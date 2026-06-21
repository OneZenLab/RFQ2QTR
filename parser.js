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
    const colStr = col == null ? '' : String(col);
    for (const [pat, target] of compiledColMaps) {
      // Python 用 re.match 从头匹配，pattern 自带 ^ 和 $ 锚点
      if (new RegExp(pat.source, pat.flags).test(colStr)) {
        mapping[colStr] = target;
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
  const m = text.match(/\b(\d+)\s*[*Xx×]\s*([\d\s\-\/]+?)(?:\s*["']|NB|DN|\s*$|\s*\))/i)
         ?? text.match(/\b(\d+)\s*[*Xx×]\s*([\d\s\-\/]+)/i);
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

  // 特例：描述含 "RED TEE" → TYPE=TEE，RED 归入 MISC（Python 原有逻辑）
  if (upperDesc.includes('RED TEE')) {
    if (uni.UNI_TYPE === 'RED') uni.UNI_TYPE = 'TEE';
    if (!uni.UNI_MISC) uni.UNI_MISC = 'RED';
  }

  // ── WT 提取 ───────────────────────────────────────────────────
  let wts = [];
  const wtMatchRanges = [];  // 用于防止 OD 提取与 WT 位置重叠

  if (enumPatterns.UNI_WT) {
    // 1. 从描述文本中逐一匹配 WT 枚举（遍历所有，非首个）
    for (const [valueName, pat] of enumPatterns.UNI_WT) {
      const tp = new RegExp(pat.source, pat.flags);
      const m  = tp.exec(upperDesc);
      if (m) {
        wts.push(valueName);
        wtMatchRanges.push([m.index, m.index + m[0].length]);
      }
    }
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
    const extractOds = (text) => {
      const allMatches  = [];
      const usedRanges  = [];

      const addIfValid = (m, valueName) => {
        const s = m.index, e = m.index + m[0].length;
        // 已被更长匹配包含
        if (usedRanges.some(([a, b]) => s >= a && e <= b)) return;
        // 与 WT 位置重叠（wtMatchRanges 基于 upperDesc）
        if (wtMatchRanges.some(([a, b]) => s < b && e > a)) return;
        // 是小数的一部分（前一字符是数字或小数点）
        if (s > 0 && (text[s - 1] === '.' || /\d/.test(text[s - 1]))) return;
        // 后一字符是小数点
        if (e < text.length && text[e] === '.') return;
        usedRanges.push([s, e]);
        allMatches.push([s, -m[0].length, valueName]);
      };

      for (const [valueName, pat] of enumPatterns.UNI_OD) {
        const gp = new RegExp(pat.source, pat.flags.includes('i') ? 'gi' : 'g');
        let m;
        while ((m = gp.exec(text)) !== null) addIfValid(m, valueName);
      }

      // 按位置升序、同位置按匹配长度降序排列（与 Python all_matches.sort() 一致）
      allMatches.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
      const seen = new Set(), result = [];
      for (const [,, v] of allMatches) {
        if (!seen.has(v)) { seen.add(v); result.push(v); }
      }
      return result;
    };

    ods = extractOds(searchText);
    // searchText 无结果时再从完整描述提取（Python 同逻辑）
    if (!ods.length) ods = extractOds(upperDesc);
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
      // 三通：OD1（主径）和 OD3（支管径），与 Python 一致
      uni.UNI_OD1 = ods[0];
      if (ods.length >= 2) uni.UNI_OD3 = ods[1];
    } else {
      if (ods.length >= 1) uni.UNI_OD1 = ods[0];
      if (ods.length >= 2) uni.UNI_OD2 = ods[1];
      if (ods.length >= 3) uni.UNI_OD3 = ods[2];
    }
  }

  // ── WT 分配规则 ───────────────────────────────────────────────
  const odCount = [uni.UNI_OD1, uni.UNI_OD2, uni.UNI_OD3].filter(Boolean).length;

  if (wts.length) {
    if (uniType === 'TEE') {
      // TEE 有 2 个 OD 且 WT≥2 时：WT1 和 WT3（Python 原有逻辑）
      if (odCount === 2 && wts.length >= 2) {
        uni.UNI_WT1 = wts[0]; uni.UNI_WT3 = wts[1];
      } else {
        uni.UNI_WT1 = wts[0];
      }
    } else {
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

    // 从父行继承未解析到的属性（Python 原有逻辑）
    for (const [k, v] of Object.entries(parentAttrs)) {
      if (!uniRow[k]) uniRow[k] = v;
    }

    parsedRows.push(uniRow);
    preprocessRows.push({ index: absIdx, type: 'child', descText, uni: uniRow });
  });

  // 构建输出行
  const outputRows = parsedRows.map(uni => buildOutputRow(uni, uniToCodeMap));

  return { rawRows, headerRow: headerRowIdx, headers, colSemantics, preprocessRows, outputRows };
}

module.exports = { parseExcel };
