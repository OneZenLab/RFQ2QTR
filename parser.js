'use strict';

/**
 * parser.js — 忠实移植自 UNIXLS.py
 * 读取 Excel 第1张表，按 Lexer.yaml 规则解析管道配件 RFQ 数据。
 */

const fs   = require('fs');
const path = require('path');
const yaml = require('js-yaml');
const XLSX = require('xlsx');

const LEXER_PATH = path.join(__dirname, 'Lexer.yaml');

/* ═══════════════════════════════════════════════════════════════
   Lexer.yaml 加载与编译
   ═══════════════════════════════════════════════════════════════ */

function loadLexer() {
  const data = yaml.load(fs.readFileSync(LEXER_PATH, 'utf8'));
  return data?.lexer ?? {};
}

/** YAML flags 列表 → RegExp flags 字符串（目前只处理 IGNORECASE） */
function flagsToStr(flagsList) {
  return (flagsList || []).some(f => f.toUpperCase() === 'IGNORECASE') ? 'i' : '';
}

/**
 * 编译所有枚举正则（对应 Python compile_enum_patterns）
 * 返回 { enumName: [[valueName, RegExp], ...] }
 * 顺序严格按 YAML 文件声明顺序，UNI_MATERIAL 额外做显式排序保证。
 */
function compileEnumPatterns(enumConf) {
  const compiled = {};
  for (const [enumName, conf] of Object.entries(enumConf)) {
    const mappings = conf.mappings ?? {};
    // 按 YAML key 声明顺序构建（js-yaml 保留插入顺序）
    const yamlOrder = Object.keys(mappings);
    const bucket = [];
    for (const valueName of yamlOrder) {
      const patConf = mappings[valueName];
      const pattern = patConf.pattern ?? '';
      if (!pattern) continue;
      try {
        bucket.push([valueName, new RegExp(pattern, flagsToStr(patConf.flags))]);
      } catch (e) { /* 跳过无效正则 */ }
    }
    // Python 对 UNI_MATERIAL 显式按 YAML 顺序排序（防御性措施）
    if (enumName === 'UNI_MATERIAL') {
      const orderMap = Object.fromEntries(yamlOrder.map((k, i) => [k, i]));
      bucket.sort((a, b) => (orderMap[a[0]] ?? Infinity) - (orderMap[b[0]] ?? Infinity));
    }
    compiled[enumName] = bucket;
    compiled[`${enumName}_KEYS`] = new Set(bucket.map(([v]) => v));
  }
  return compiled;
}

/** 编译列名映射规则（对应 Python compile_column_mapping） */
function compileColumnMapping(rulesConf) {
  return (rulesConf.column_map?.mappings ?? []).flatMap(m => {
    try {
      return [[new RegExp(m.source, flagsToStr(m.flags)), m.target]];
    } catch (e) { return []; }
  });
}

/** 编译忽略列正则（对应 Python compile_ignore_cols_pattern） */
function compileIgnoreCols(rulesConf) {
  const c = rulesConf.ignore_cols ?? {};
  if (!c.pattern) return /(?!)/;     // 永不匹配
  try { return new RegExp(c.pattern, flagsToStr(c.flags)); }
  catch (e) { return /(?!)/; }
}

/* ═══════════════════════════════════════════════════════════════
   表头检测 / 列语义推断
   ═══════════════════════════════════════════════════════════════ */

/**
 * 对前 20 行打分，关键词命中最多的行即为表头（对应 Python detect_header_row）
 */
function detectHeaderRow(rawRows, rulesConf) {
  const keywords = (rulesConf.attr_table_keywords?.keywords ?? []).map(k => String(k).toUpperCase());

  function score(values) {
    let sc = 0, nonEmpty = 0;
    for (const v of values) {
      const t = String(v ?? '').trim();
      if (!t) continue;
      nonEmpty++;
      if (keywords.some(k => t.toUpperCase().includes(k))) sc += 2;
    }
    if (nonEmpty >= 3) sc += 1;
    return sc;
  }

  let bestIdx = 0, bestScore = -1;
  for (let i = 0; i < Math.min(rawRows.length, 20); i++) {
    const sc = score(rawRows[i]);
    if (sc > bestScore) { bestScore = sc; bestIdx = i; }
  }
  return bestIdx;
}

/**
 * 将原始列名映射到内部语义字段（对应 Python infer_column_semantics）
 * 使用 column_map 规则，首个匹配胜出。
 */
function inferColumnSemantics(columns, compiledColMaps) {
  const mapping = {};
  for (const col of columns) {
    const raw    = col == null ? '' : String(col);
    // 多行表头（含 \r\n）归一化为单行再匹配
    const colStr = raw.replace(/[\r\n]+/g, ' ').replace(/\s{2,}/g, ' ').trim();
    for (const [pat, target] of compiledColMaps) {
      // Python 用 re.match 从头匹配，pattern 自带 ^ 和 $ 锚点
      if (new RegExp(pat.source, pat.flags).test(colStr)) {
        mapping[raw] = target;   // key 保留原始列名（与 rowObj 的 key 一致）
        break;
      }
    }
  }
  return mapping;
}

/* ═══════════════════════════════════════════════════════════════
   行分类：汇总行 / 父行 / 子行
   ═══════════════════════════════════════════════════════════════ */

/** 对应 Python is_summary_row */
function isSummaryRow(rowObj) {
  const text = Object.values(rowObj).join(' ');
  if (!text.trim()) return true;
  const up = text.toUpperCase();
  return ['TOTAL', 'SUBTOTAL', '合计', '小计', 'REMARK', '备注'].some(k => up.includes(k));
}

/**
 * 对应 Python is_parent_row
 * 父行特征：有描述、无数量、无价格、行号层级≤1、描述命中≥2种UNI枚举类型
 */
function isParentRow(rowObj, colSemantics, enumPatterns) {
  const descCols = [], qtyCols = [], priceCols = [], lineCols = [];
  for (const [col, sem] of Object.entries(colSemantics)) {
    if      (sem === 'QTR_LINE')  lineCols.push(col);
    else if (sem === 'QTR_QTY')   qtyCols.push(col);
    else if (sem === 'QTR_PRICE') priceCols.push(col);
    else                          descCols.push(col);
  }

  // 没有数量列也没有价格列 → 无法区分父/子行结构，全部视为子行
  if (!qtyCols.length && !priceCols.length) return false;

  // 取描述文本（优先 Item 列，与 Python 一致）
  let descText = '';
  if ('Item' in rowObj) {
    descText = String(rowObj['Item'] ?? '').trim();
  } else {
    for (const col of descCols) {
      const v = String(rowObj[col] ?? '').trim();
      if (v) { descText = v; break; }
    }
  }
  if (!descText) return false;

  // 有数量 → 子行
  if (qtyCols.some(col => String(rowObj[col] ?? '').trim())) return false;
  // 有价格 → 子行
  if (priceCols.some(col => String(rowObj[col] ?? '').trim())) return false;

  // 行号层级分析：取第一个非空行号（与 Python 一致）
  let lineNum = '';
  for (const col of lineCols) {
    const v = String(rowObj[col] ?? '').trim();
    if (v) { lineNum = v; break; }
  }
  if (lineNum && (lineNum.match(/\./g) || []).length > 1) return false;

  // 描述命中 ≥2 种 UNI 枚举类型才视为父行
  const upperDesc = descText.toUpperCase();
  const uniTypes  = [
    'UNI_TYPE', 'UNI_ANGLE', 'UNI_RADIUS', 'UNI_MISC',
    'UNI_DIM_SPEC', 'UNI_MATERIAL', 'UNI_CONSTRUCTION', 'UNI_END_PREPARATION',
  ];
  let hitCount = 0;
  for (const t of uniTypes) {
    if (!enumPatterns[t]) continue;
    for (const [, pat] of enumPatterns[t]) {
      if (new RegExp(pat.source, pat.flags).test(upperDesc)) { hitCount++; break; }
    }
  }
  return hitCount >= 2;
}

/* ═══════════════════════════════════════════════════════════════
   描述文本构建（对应 Python build_desc_text）
   ═══════════════════════════════════════════════════════════════ */

/**
 * 将行中各描述列的值拼接成完整描述文本。
 * 忽略：ignore_cols 列、行号/数量/价格列。
 * 特殊：Item 列即使映射到 QTR_LINE 也纳入描述；DN 列名前缀值。
 */
function buildDescText(rowObj, headers, colSemantics, ignoreCols) {
  const parts = [];
  for (const colName of headers) {
    const nameStr   = colName == null ? '' : String(colName);
    // 跳过 ignore_cols
    if (ignoreCols.test(nameStr)) continue;

    const internal = colSemantics[nameStr] ?? '';

    // Item 列即使映射到 QTR_LINE 也包含（Python 特例）
    if (internal === 'QTR_LINE' && nameStr.toLowerCase() === 'item') {
      const v = String(rowObj[colName] ?? '').trim();
      if (v) parts.push(v);
      continue;
    }
    // 跳过行号/数量/价格
    if (['QTR_LINE', 'QTR_QTY', 'QTR_PRICE'].includes(internal)) continue;

    const v = rowObj[colName];
    if (v == null || v === '') continue;
    let txt = String(v).trim();
    if (!txt) continue;
    // 列名含 DN → 值前加 DN
    if (nameStr.toUpperCase().includes('DN')) txt = `DN${txt}`;
    parts.push(txt);
  }
  return parts.join(' ');
}

/* ═══════════════════════════════════════════════════════════════
   枚举匹配（对应 Python match_first_enum）
   ═══════════════════════════════════════════════════════════════ */

/** 在 enumPatternList 中找第一个命中的值名（search 语义，非 match） */
function matchFirstEnum(text, enumPatternList) {
  for (const [valueName, pat] of enumPatternList) {
    if (new RegExp(pat.source, pat.flags).test(text)) return valueName;
  }
  return '';
}

/* ═══════════════════════════════════════════════════════════════
   OD 尺寸工具（对应 Python parse_od_range）
   ═══════════════════════════════════════════════════════════════ */

// "N*M" 中第二段的分数标准化
const OD_FRAC_NORM = {
  '1 1/2': '1-1/2', '1 1/4': '1-1/4', '2 1/2': '2-1/2', '3 1/2': '3-1/2',
  '1-1/2': '1-1/2', '1-1/4': '1-1/4', '2-1/2': '2-1/2', '3-1/2': '3-1/2',
};

/**
 * 解析 "N*M" / "NxM" 形式的两段尺寸（对应 Python parse_od_range）
 * 支持: 4X1、4*1、4 * 1 1/2" 等
 */
function parseOdRange(sizeText) {
  if (!sizeText) return null;
  const text = String(sizeText).trim();
  // 与 Python 相同的正则：支持分数形式的第二段
  // 加一个否定顺序环视 (?<![A-Za-z\/\-])，避免把壁厚写法（如 "S/20"、"S-40"）
  // 里紧跟分隔符的数字尾巴误当成口径——这类数字前面是字母/斜杠/短横线，
  // 不应被当作独立的口径数字
  const m = text.match(/(?<![A-Za-z\/\-])\b(\d+)\s*[*Xx×]\s*([\d\s\-\/]+?)(?:\s*["']|NB|DN|\s*$|\s*\))/i)
         ?? text.match(/(?<![A-Za-z\/\-])\b(\d+)\s*[*Xx×]\s*([\d\s\-\/]+)/i);
  if (!m) return null;
  const first      = m[1].trim();
  let   secondRaw  = m[2].trim();
  if (/^\d+\s+\d+\/\d+$/.test(secondRaw)) {
    secondRaw = secondRaw.replace(' ', '-');
  } else {
    secondRaw = OD_FRAC_NORM[secondRaw] ?? secondRaw;
  }
  return [first, secondRaw];
}

/* ═══════════════════════════════════════════════════════════════
   核心行解析（对应 Python parse_row_to_uni）
   ═══════════════════════════════════════════════════════════════ */

function parseRowToUni(rowObj, headers, rowIndex, colSemantics, enumPatterns, ignoreCols, sizeCol) {
  const uni = {
    UNI_TYPE: '', UNI_ANGLE: '', UNI_RADIUS: '', UNI_MISC: '',
    UNI_OD1: '', UNI_OD2: '', UNI_OD3: '',
    UNI_WT1: '', UNI_WT2: '', UNI_WT3: '',
    UNI_DIM_SPEC: '', UNI_MATERIAL: '', UNI_CONSTRUCTION: '', UNI_END_PREPARATION: '',
    UNI_CATEGORY: '', UNI_PRESSURE_CLASS: '', UNI_FACING: '',
    UNI_BOLT_GRADE: '', UNI_BOLT_LENGTH: '', UNI_THREAD_TYPE: '',
    UNI_GASKET_TYPE: '', UNI_GASKET_THK: '',
    QTR_LINE: '', QTR_QTY: '', QTR_PRICE: '',
  };

  // ── QTR 字段 ──────────────────────────────────────────────────
  const lineCol = Object.entries(colSemantics).find(([, t]) => t === 'QTR_LINE')?.[0];
  uni.QTR_LINE  = lineCol != null ? String(rowObj[lineCol] ?? '').trim()
                                  : String((rowIndex + 1) * 10);

  const qtyCol = Object.entries(colSemantics).find(([, t]) => t === 'QTR_QTY')?.[0];
  if (qtyCol != null) uni.QTR_QTY = String(rowObj[qtyCol] ?? '').trim();

  // 单价：找第一个有值的 QTR_PRICE 列
  for (const [col, sem] of Object.entries(colSemantics)) {
    if (sem === 'QTR_PRICE') {
      const v = rowObj[col];
      if (v != null && v !== '') { uni.QTR_PRICE = v; break; }
    }
  }

  // ── 描述文本 ──────────────────────────────────────────────────
  const descText  = buildDescText(rowObj, headers, colSemantics, ignoreCols);
  const upperDesc = descText.toUpperCase();

  // Size 列文本（优先用于 OD 解析）
  let sizeText = '';
  if (sizeCol && rowObj[sizeCol] != null) sizeText = String(rowObj[sizeCol]);
  const upperSize = sizeText.toUpperCase();

  // ── 枚举提取（TYPE / ANGLE / RADIUS / MISC / MATERIAL / ...） ─
  const me = (name) => enumPatterns[name] ? matchFirstEnum(upperDesc, enumPatterns[name]) : '';

  uni.UNI_TYPE            = me('UNI_TYPE');
  uni.UNI_ANGLE           = me('UNI_ANGLE');
  uni.UNI_RADIUS          = me('UNI_RADIUS');
  uni.UNI_MISC            = me('UNI_MISC');
  uni.UNI_MATERIAL        = me('UNI_MATERIAL');
  uni.UNI_CONSTRUCTION    = me('UNI_CONSTRUCTION');
  uni.UNI_END_PREPARATION = me('UNI_END_PREPARATION');
  const dimSpec           = me('UNI_DIM_SPEC');
  if (dimSpec) uni.UNI_DIM_SPEC = dimSpec;

  // 新增品类相关枚举提取
  uni.UNI_CATEGORY       = me('UNI_CATEGORY');
  uni.UNI_PRESSURE_CLASS = me('UNI_PRESSURE_CLASS');
  uni.UNI_FACING         = me('UNI_FACING');
  uni.UNI_BOLT_GRADE     = me('UNI_BOLT_GRADE');
  uni.UNI_THREAD_TYPE    = me('UNI_THREAD_TYPE');
  uni.UNI_GASKET_TYPE    = me('UNI_GASKET_TYPE');

  // 特例：描述同时命中 TEE 和 RED（顺序不限，如 "TEE RED" 或 "RED TEE"）
  // → TYPE=TEE，RED 归入 MISC（原逻辑只认字面 "RED TEE" 子串，此处放宽为顺序无关）
  if (uni.UNI_TYPE === 'RED') {
    const teePat = enumPatterns.UNI_TYPE?.find(([v]) => v === 'TEE')?.[1];
    if (teePat && new RegExp(teePat.source, teePat.flags).test(upperDesc)) {
      uni.UNI_TYPE = 'TEE';
      if (!uni.UNI_MISC) uni.UNI_MISC = 'RED';
    }
  }

  // ── WT 提取 ───────────────────────────────────────────────────
  let wts = [];
  const wtMatchRanges = [];  // 用于防止 OD 提取与 WT 位置重叠

  if (enumPatterns.UNI_WT) {
    // 1. 从描述文本中逐一匹配 WT 枚举（遍历所有，非首个），按文本中出现的
    //    位置排序（而非枚举字典声明顺序），保证与 OD 的先后顺序一一对应
    //    —— 例如 "6" Sch40 * 4" Sch10S" 应得到 wts=[S40, S10S]
    const wtCandidates = [];
    for (const [valueName, pat] of enumPatterns.UNI_WT) {
      const tp = new RegExp(pat.source, pat.flags);
      const m  = tp.exec(upperDesc);
      if (m) {
        wtCandidates.push([m.index, valueName]);
        wtMatchRanges.push([m.index, m.index + m[0].length]);
      }
    }
    wtCandidates.sort((a, b) => a[0] - b[0]);
    wts = wtCandidates.map(([, v]) => v);

    // 2. 直接映射到 UNI_WT1/UNI_WT2 的列（如 SCH1、SCH2）
    for (const [colName, internal] of Object.entries(colSemantics)) {
      if (internal !== 'UNI_WT1' && internal !== 'UNI_WT2') continue;
      const wtVal = String(rowObj[colName] ?? '').trim().toUpperCase();
      if (!wtVal) continue;
      for (const [valueName, pat] of enumPatterns.UNI_WT) {
        if (new RegExp(pat.source, pat.flags).test(wtVal)) { wts.push(valueName); break; }
      }
    }
  }

  // 去重（保留首次出现顺序，与 Python dict.fromkeys 一致）
  wts = [...new Set(wts)];
  // 清理重叠规格：更具体的优先（Python 原有逻辑）
  if (wts.includes('XXS') && wts.includes('XS'))   wts = wts.filter(w => w !== 'XS');
  if (wts.includes('S40S') && wts.includes('S40'))  wts = wts.filter(w => w !== 'S40');
  if (wts.includes('S80S') && wts.includes('S80'))  wts = wts.filter(w => w !== 'S80');
  if (wts.includes('S10S') && wts.includes('S10'))  wts = wts.filter(w => w !== 'S10');

  // ── OD 提取 ───────────────────────────────────────────────────
  let ods = [];
  // searchText = Size 列优先，其次用描述文本
  const searchText  = upperSize || upperDesc;
  const rangeOds    = parseOdRange(sizeText || descText);
  let   isRangePat  = false;

  if (rangeOds) {
    // "N*M" 模式直接给出两段 OD
    ods = rangeOds;
    isRangePat = true;
  } else if (enumPatterns.UNI_OD) {
    // 从文本中提取 OD 枚举，去除与 WT 位置重叠及小数误判
    //
    // 注意：候选匹配的"谁包含谁"判定不能依赖字典 key 的遍历顺序——JS 对象对
    // 形如 '1'、'10'、'20' 这类整数形字符串 key 会自动按数值升序排到最前面
    // （无视 YAML 里的原始声明顺序），导致像 '1 1/2' 这种本该优先命中的分数
    // 口径反而排到所有整数 key 之后处理，使重叠的短匹配（如 '1'、'2'）先一步
    // 被计入结果。因此这里改为：先收集全部候选（顺序无关），再统一按"位置升
    // 序 + 同位置更长匹配优先"贪心选取，重叠的候选一律跳过（不仅仅是被完全
    // 包含的情况）。
    const extractOds = (text) => {
      const candidates = [];

      for (const [valueName, pat] of enumPatterns.UNI_OD) {
        const gp = new RegExp(pat.source, pat.flags.includes('i') ? 'gi' : 'g');
        let m;
        while ((m = gp.exec(text)) !== null) {
          const s = m.index, e = m.index + m[0].length;
          // 与 WT 位置重叠（wtMatchRanges 基于 upperDesc）
          if (wtMatchRanges.some(([a, b]) => s < b && e > a)) continue;
          // 是小数的一部分（前一字符是数字或小数点）
          if (s > 0 && (text[s - 1] === '.' || /\d/.test(text[s - 1]))) continue;
          // 后一字符是小数点
          if (e < text.length && text[e] === '.') continue;
          candidates.push([s, e, valueName]);
        }
      }

      // 位置升序；同位置时更长的匹配（更具体）优先
      candidates.sort((a, b) => a[0] - b[0] || (b[1] - b[0]) - (a[1] - a[0]));

      const result = [], seen = new Set();
      let lastEnd = -1;
      for (const [s, e, valueName] of candidates) {
        if (s < lastEnd) continue; // 与已选中的更靠前/更长的匹配重叠，跳过
        lastEnd = e;
        if (!seen.has(valueName)) { seen.add(valueName); result.push(valueName); }
      }
      return result;
    };

    ods = extractOds(searchText);
    // searchText 结果不足 2 个时，也尝试从完整描述提取，取更完整的一份
    // （避免像 SIZE1 这类只含单个数值的列，遮蔽了描述文本里 "OD1 * OD2" 的完整信息）
    if (ods.length < 2 && searchText !== upperDesc) {
      const fromDesc = extractOds(upperDesc);
      if (fromDesc.length > ods.length) ods = fromDesc;
    }
  }

  // 直接映射到 UNI_OD1/UNI_OD2 的列（如 SIZE1、SIZE2）：这类结构化列数据是
  // 明确无歧义的，始终优先于从自由文本正则提取的结果（后者容易被描述里的
  // 其他数字，如未映射的序号列泄漏进描述文本，干扰匹配顺序）
  if (enumPatterns.UNI_OD) {
    const colOds = [];
    for (const [colName, internal] of Object.entries(colSemantics)) {
      if (internal !== 'UNI_OD1' && internal !== 'UNI_OD2') continue;
      const odVal = String(rowObj[colName] ?? '').trim().toUpperCase();
      if (!odVal) continue;
      const matched = matchFirstEnum(odVal, enumPatterns.UNI_OD);
      if (matched) colOds.push(matched);
    }
    if (colOds.length) ods = colOds;
  }

  // ── 无 WT 时从 "BW 数字" 推断壁厚（Python 原有逻辑） ──────────
  if (!wts.length) {
    const bwM = upperDesc.match(/\bBW\s+(\d+)\b/);
    if (bwM && ['10', '20', '40', '60', '80', '160'].includes(bwM[1])) {
      wts.push(`S${bwM[1]}`);
    }
  }

  // ── OD 分配规则 ───────────────────────────────────────────────
  const uniType = uni.UNI_TYPE.toUpperCase();
  // 异径三通（TEE RED）：主管缩径端与支管共用第二段口径/壁厚
  const isReducingTee = uniType === 'TEE' && uni.UNI_MISC === 'RED';

  if (ods.length) {
    // 减径管件（范围模式）：较大口径作为 OD1（Python 同逻辑）
    // Python: float(str(x).replace("-", ".")) → 对整数有效，分数会异常 → 保持原顺序
    if (uniType === 'RED' && isRangePat) {
      const vals = ods.map(od => parseFloat(od));  // 范围模式 OD 都是纯整数
      if (vals.every(v => !isNaN(v))) {
        ods = [...ods].sort((a, b) => parseFloat(b) - parseFloat(a));
      }
    }

    if (uniType === 'TEE') {
      // 三通：异径三通（TEE RED）填 OD1（主径）+ OD3（支管缩径），OD2 留空；
      // 等径三通（TEE EQL）及其他情况只填 OD1，OD2/OD3 均留空
      uni.UNI_OD1 = ods[0];
      if (isReducingTee && ods.length >= 2) uni.UNI_OD3 = ods[1];
    } else {
      if (ods.length >= 1) uni.UNI_OD1 = ods[0];
      if (ods.length >= 2) uni.UNI_OD2 = ods[1];
      if (ods.length >= 3) uni.UNI_OD3 = ods[2];
    }
  }

  // ── WT 分配规则 ───────────────────────────────────────────────
  if (wts.length) {
    if (uniType === 'TEE') {
      // 异径三通：WT1（主管）+ WT3（支管缩径），WT2 留空；
      // 等径三通及其他情况只填 WT1
      uni.UNI_WT1 = wts[0];
      if (isReducingTee && ods.length >= 2 && wts.length >= 2) uni.UNI_WT3 = wts[1];
    } else {
      const odCount = [uni.UNI_OD1, uni.UNI_OD2, uni.UNI_OD3].filter(Boolean).length;
      if (odCount <= 1) {
        // 只有 1 个 OD → 只显示 WT1
        uni.UNI_WT1 = wts[0];
      } else {
        if (wts.length >= 1) uni.UNI_WT1 = wts[0];
        // WT1 和 WT2 相同时只显示 WT1（Python 原有逻辑）
        if (wts.length >= 2 && wts[0] !== wts[1]) uni.UNI_WT2 = wts[1];
        if (wts.length >= 3) uni.UNI_WT3 = wts[2];
      }
    }
  }

  // ── 从专用列读取新增字段（列值覆盖枚举提取，适用于有独立列的结构化文件） ──
  const EXTRA_COL_FIELDS = new Set([
    'UNI_PRESSURE_CLASS','UNI_FACING','UNI_BOLT_GRADE',
    'UNI_BOLT_LENGTH','UNI_THREAD_TYPE','UNI_GASKET_TYPE','UNI_GASKET_THK',
  ]);
  for (const [colName, internal] of Object.entries(colSemantics)) {
    if (!EXTRA_COL_FIELDS.has(internal)) continue;
    const v = String(rowObj[colName] ?? '').trim();
    if (v) uni[internal] = v;
  }

  return uni;
}

/* ═══════════════════════════════════════════════════════════════
   输出行构建（对应 Python build_output_dataframe）
   ═══════════════════════════════════════════════════════════════ */

/** 将 UNI_* 字段按 uni_to_code_map 映射为中文列名输出行 */
function buildOutputRow(uni, uniToCodeMap) {
  const row = {};
  for (const [uniField, cnName] of Object.entries(uniToCodeMap)) {
    const v = uni[uniField] ?? '';
    row[cnName] = v !== '' ? String(v) : '';
  }
  return row;
}

/* ═══════════════════════════════════════════════════════════════
   分类检测 + 置信度评分
   ═══════════════════════════════════════════════════════════════ */

// UNI_TYPE → 品类路由表
const FITTING_TYPE_SET = new Set(['CAP','ELB','RED','TEE','STE','XOS','COUP']);
const FLANGE_TYPE_SET  = new Set(['WN','SO','BL','SW_FL','THD_FL','LJ','OR']);
const BOLT_TYPE_SET    = new Set(['STUD','HEX']);
const GASKET_TYPE_SET  = new Set(['SW_GSK','RF_GSK','FF_GSK','NRS']);

// UNI_TYPE 枚举和 UNI_CATEGORY 均未命中时的关键字兜底（OLE/PLUG 等特殊件）
const FALLBACK_KEYWORDS = [
  [/WELDOLET|SOCKOLET|THREDOLET|\bOLET\b/i, 'FITTING', 'OLE'],
  [/\bPLUG\b|BOUCHON/i,                      'FITTING', 'PLUG'],
];

/**
 * 品类路由：UNI_TYPE → UNI_CATEGORY 枚举 → 关键字兜底
 * @param {string} descText - 完整描述文本
 * @param {object} uniRow   - 已提取的 UNI 字段对象
 */
function detectCategory(descText, uniRow) {
  const t = uniRow.UNI_TYPE;
  if (FITTING_TYPE_SET.has(t)) return { category: 'FITTING', catType: t };
  if (FLANGE_TYPE_SET.has(t))  return { category: 'FLANGE',  catType: t };
  if (BOLT_TYPE_SET.has(t))    return { category: 'BOLT',    catType: t };
  if (GASKET_TYPE_SET.has(t))  return { category: 'GASKET',  catType: t };
  // UNI_CATEGORY 来自描述文本的枚举匹配（FLANGE/GASKET/BOLT 等关键字）
  const cat = uniRow.UNI_CATEGORY;
  if (cat) return { category: cat, catType: t || null };
  // 兜底：OLE / PLUG 等未进入 UNI_TYPE 枚举的特殊管件
  const upper = (descText || '').toUpperCase();
  for (const [rx, c, typ] of FALLBACK_KEYWORDS) {
    if (rx.test(upper)) return { category: c, catType: typ };
  }
  return { category: null, catType: null };
}

/**
 * 按品类评分：FITTING / FLANGE / BOLT / GASKET 各用不同字段判断置信度
 */
function scoreRow(uni, category, catType, enumPatterns) {
  if (!category) return { conf: 'low', reason: '无法确定品类(描述未命中关键字)' };

  const reasons = [];
  const od1      = uni.UNI_OD1;
  const material = uni.UNI_MATERIAL;
  const od_ok    = !!od1      && (enumPatterns['UNI_OD_KEYS']?.has(od1) ?? false);
  const mat_ok   = !!material && (enumPatterns['UNI_MATERIAL_KEYS']?.has(material) ?? false);

  if (!od_ok && od1) reasons.push(`OD "${od1}" 非标准值`);
  if (!mat_ok)       reasons.push(`材质 "${material || '(空)'}" 未命中词典`);

  switch (category) {
    case 'FITTING': {
      const typeKeys = enumPatterns['UNI_TYPE_KEYS'];
      const type_ok  = typeKeys?.has(catType) || catType === 'OLE' || catType === 'PLUG';
      if (!type_ok && catType) reasons.push(`类型 "${catType}" 未在管件枚举中`);
      if (od_ok && mat_ok && type_ok)           return { conf: 'high', reason: '' };
      if (od_ok && (mat_ok || type_ok))         return { conf: 'mid',  reason: reasons.join('；') };
      return { conf: 'low', reason: reasons.join('；') || '多个字段未命中' };
    }
    case 'FLANGE': {
      const pressure  = uni.UNI_PRESSURE_CLASS;
      const facing    = uni.UNI_FACING;
      const press_ok  = !!pressure && (enumPatterns['UNI_PRESSURE_CLASS_KEYS']?.has(pressure) ?? false);
      const facing_ok = !!facing;
      if (!pressure)      reasons.push('缺压力等级');
      else if (!press_ok) reasons.push(`压力等级 "${pressure}" 待确认`);
      if (!facing)        reasons.push('缺密封面形式');
      if (od_ok && mat_ok && press_ok && facing_ok) return { conf: 'high', reason: '' };
      if (od_ok && mat_ok)                          return { conf: 'mid',  reason: reasons.join('；') };
      return { conf: 'low', reason: reasons.join('；') || '多个字段未命中' };
    }
    case 'BOLT': {
      const grade_ok = !!uni.UNI_BOLT_GRADE;
      if (!grade_ok) reasons.push('缺螺栓等级');
      if (grade_ok && mat_ok) return { conf: 'high', reason: '' };
      if (grade_ok || mat_ok) return { conf: 'mid',  reason: reasons.join('；') };
      return { conf: 'low', reason: reasons.join('；') || '螺栓等级/材质未命中' };
    }
    case 'GASKET':
    case 'NUT':
    case 'INSULATION_KIT': {
      const pressure = uni.UNI_PRESSURE_CLASS;
      if (!uni.UNI_GASKET_TYPE && category === 'GASKET') reasons.push('缺垫片类型');
      if (!pressure) reasons.push('缺压力等级');
      reasons.unshift(`${category} 品类词典待扩展`);
      if (od_ok && mat_ok) return { conf: 'mid', reason: reasons.join('；') };
      return { conf: 'low', reason: reasons.join('；') };
    }
    default:
      return { conf: 'low', reason: `未知品类 "${category}"` };
  }
}

/* ═══════════════════════════════════════════════════════════════
   主解析入口（读第1张表）
   ═══════════════════════════════════════════════════════════════ */

/**
 * 读取 Excel 文件第1张表（索引0），按 Lexer.yaml 规则解析。
 * 返回前端所需的完整 JSON 结构。
 */
function parseExcel(filePath) {
  const lexerConf = loadLexer();
  const rulesConf = lexerConf.rules  ?? {};
  const enumsConf = lexerConf.enums  ?? {};

  const enumPatterns    = compileEnumPatterns(enumsConf);
  const compiledColMaps = compileColumnMapping(rulesConf);
  const uniToCodeMap    = rulesConf.uni_to_code_map ?? {};
  const ignoreCols      = compileIgnoreCols(rulesConf);

  // 读取第1张表，全部按字符串处理
  const wb        = XLSX.readFile(filePath);
  const sheetName = wb.SheetNames[0];
  const ws        = wb.Sheets[sheetName];
  const rawRows   = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });

  // 检测表头行
  const headerRowIdx = detectHeaderRow(rawRows, rulesConf);

  // 构建列名数组：空表头用列索引命名，重复非空表头加后缀，防止覆盖
  const headerRowRaw = rawRows[headerRowIdx] ?? [];
  const headerSeen   = {};
  const headers      = headerRowRaw.map((h, j) => {
    let name = h == null ? '' : String(h).trim();
    if (!name) {
      name = `__col${j}`;   // 空表头：用列索引作唯一 key
    } else if (headerSeen[name]) {
      name = `${name}__${j}`; // 重复表头：加列索引后缀
    }
    headerSeen[name] = true;
    return name;
  });

  // 构建带表头的数据行对象数组
  const dataRows = rawRows.slice(headerRowIdx + 1).map(arr => {
    const obj = {};
    headers.forEach((h, j) => { obj[h] = arr[j] ?? ''; });
    return obj;
  });

  // 列语义映射
  const colSemantics = inferColumnSemantics(headers, compiledColMaps);

  // 检测 Size 列
  const sizeCol = headers.find(h => h.toUpperCase().includes('SIZE')) ?? null;

  // ── 逐行解析 ─────────────────────────────────────────────────
  const parsedRows     = [];
  const preprocessRows = [];
  let   parentAttrs    = {};   // 当前父行携带的属性，供子行继承

  dataRows.forEach((rowObj, idx) => {
    const absIdx = headerRowIdx + 1 + idx;  // 在原始表中的行号

    if (isSummaryRow(rowObj)) {
      preprocessRows.push({ index: absIdx, type: 'summary', descText: '', uni: null });
      return;
    }

    if (isParentRow(rowObj, colSemantics, enumPatterns)) {
      // 父行：解析属性（不忽略任何列）并保存供子行继承
      const parentUni = parseRowToUni(rowObj, headers, idx, colSemantics, enumPatterns, /(?!)/, sizeCol);
      parentAttrs = {};
      for (const [k, v] of Object.entries(parentUni)) {
        if (v && !['QTR_LINE', 'QTR_QTY', 'QTR_PRICE'].includes(k)) parentAttrs[k] = v;
      }
      const dt = buildDescText(rowObj, headers, colSemantics, ignoreCols);
      preprocessRows.push({ index: absIdx, type: 'parent', descText: dt, uni: parentUni });
      return;
    }

    // 粗过滤：描述和数量都为空则跳过
    const descText = buildDescText(rowObj, headers, colSemantics, ignoreCols);
    const qtyCol   = Object.entries(colSemantics).find(([, t]) => t === 'QTR_QTY')?.[0];
    const hasQty   = qtyCol != null && !!String(rowObj[qtyCol] ?? '').trim();
    if (!descText.trim() && !hasQty) {
      preprocessRows.push({ index: absIdx, type: 'skip', descText: '', uni: null });
      return;
    }

    // 解析子行
    const uniRow = parseRowToUni(rowObj, headers, idx, colSemantics, enumPatterns, ignoreCols, sizeCol);

    // 序号 = Excel 实际行号（与 Tab2"行"列一致），1-based
    uniRow.QTR_LINE = String(absIdx + 1);

    // 从父行继承未解析到的属性（Python 原有逻辑）
    for (const [k, v] of Object.entries(parentAttrs)) {
      if (!uniRow[k]) uniRow[k] = v;
    }

    const { category, catType } = detectCategory(descText, uniRow);
    const { conf, reason }      = scoreRow(uniRow, category, catType, enumPatterns);

    uniRow._category = category;
    uniRow._catType  = catType;
    uniRow._conf     = conf;
    uniRow._reason   = reason;
    uniRow._srcRow   = absIdx + 1;

    parsedRows.push(uniRow);
    preprocessRows.push({
      index: absIdx, type: 'child', descText, uni: uniRow,
      conf, reason, category,
    });
  });

  // 构建输出行
  const outputRows = parsedRows.map(uni => buildOutputRow(uni, uniToCodeMap));

  const enrichedRows = parsedRows.map(uni => ({
    srcRow:       uni._srcRow,
    item:         uni.QTR_LINE,
    category:     uni._category,
    catType:      uni._catType,
    qty:          uni.QTR_QTY,
    price:        uni.QTR_PRICE,
    type:         uni.UNI_TYPE,
    angle:        uni.UNI_ANGLE,
    radius:       uni.UNI_RADIUS,
    misc:         uni.UNI_MISC,
    od1:          uni.UNI_OD1,
    od2:          uni.UNI_OD2,
    od3:          uni.UNI_OD3,
    wt:           uni.UNI_WT1,
    wt2:          uni.UNI_WT2,
    wt3:          uni.UNI_WT3,
    dimSpec:      uni.UNI_DIM_SPEC,
    material:     uni.UNI_MATERIAL,
    construction: uni.UNI_CONSTRUCTION,
    ends:         uni.UNI_END_PREPARATION,
    pressure:     uni.UNI_PRESSURE_CLASS,
    facing:       uni.UNI_FACING,
    boltGrade:    uni.UNI_BOLT_GRADE,
    boltLength:   uni.UNI_BOLT_LENGTH,
    threadType:   uni.UNI_THREAD_TYPE,
    gasketType:   uni.UNI_GASKET_TYPE,
    gasketThk:    uni.UNI_GASKET_THK,
    conf:         uni._conf,
    reason:       uni._reason,
  }));

  return { rawRows, headerRow: headerRowIdx, headers, colSemantics, preprocessRows, outputRows, enrichedRows };
}

module.exports = { parseExcel };
