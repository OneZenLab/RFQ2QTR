'use strict';

/* ── 工具 ─────────────────────────────────────────────── */
function esc(v) {
  return String(v ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function $(id) { return document.getElementById(id); }

const CONF_STYLE = {
  high:      { bg: '#eef8f1', badge: '高',     cls: 'badge-high' },
  mid:       { bg: '#fcfcfa', badge: '中',     cls: 'badge-mid'  },
  low:       { bg: '#fdf3ea', badge: '低·待审', cls: 'badge-low'  },
  confirmed: { bg: '#eef8f1', badge: '已确认',  cls: 'badge-high' },
};

const UNI_FIELDS_COMPACT = [
  ['UNI_TYPE','类型'], ['UNI_OD1','OD1'], ['UNI_OD2','OD2'],
  ['UNI_WT1','WT1'],  ['UNI_MATERIAL','材质'], ['UNI_END_PREPARATION','端口'],
];

const EDITABLE_COLS = ['qty','price','type','angle','radius','misc','od1','od2','od3','wt','wt2','wt3','dimSpec','material','construction','ends','pressure','facing','boltGrade','boltLength','threadType','gasketType','gasketThk'];

/* ── Tab 切换 ─────────────────────────────────────────── */
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    const t = tab.dataset.tab;
    document.querySelectorAll('.tab').forEach(b => b.classList.toggle('active', b === tab));
    document.querySelectorAll('.pane').forEach(p => p.classList.toggle('active', p.id === `pane-${t}`));
    // 切换 tab 时以右侧第一可见产品行为基准，将目标面板行对齐
    const R = $('result-wrap');
    const anchor = firstVisibleTr(R, 'tr[data-seq]');
    if (anchor) {
      if (t === 'pre' && _preWrap) scrollToTr(_preWrap, _preWrap.querySelector(`tr[data-seq="${anchor.dataset.seq}"]`));
      if (t === 'raw' && _rawWrap && anchor.dataset.srcRow) scrollToTr(_rawWrap, _rawWrap.querySelector(`tr[data-row="${anchor.dataset.srcRow}"]`));
    }
  });
});

/* ── 文件输入 / 拖拽 ──────────────────────────────────── */
$('file-input').addEventListener('change', e => {
  if (e.target.files[0]) handleFile(e.target.files[0]);
  e.target.value = '';
});
const dropZone = $('drop-zone');
dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', ()  => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
});

/* ── 本地文件直读 ────────────────────────────────────── */
const localBtn      = $('local-btn');
const localDropdown = $('local-dropdown');

localBtn.addEventListener('click', async e => {
  e.stopPropagation();
  if (!localDropdown.hidden) { localDropdown.hidden = true; return; }
  localDropdown.innerHTML = '<div class="file-item" style="color:#aaa">加载中…</div>';
  localDropdown.hidden = false;
  try {
    const files = await (await fetch('/api/local-files')).json();
    if (!files.length) {
      localDropdown.innerHTML = '<div class="file-item" style="color:#aaa">QTR/ 目录无 xlsx 文件</div>';
      return;
    }
    localDropdown.innerHTML = '';
    files.forEach(f => {
      const div = document.createElement('div');
      div.className   = 'file-item';
      div.textContent = f.name;
      div.title       = `${(f.size / 1024).toFixed(1)} KB`;
      div.addEventListener('click', () => {
        localDropdown.hidden = true;
        loadLocalFile(f.name);
      });
      localDropdown.appendChild(div);
    });
  } catch (err) {
    localDropdown.innerHTML = `<div class="file-item" style="color:red">${err.message}</div>`;
  }
});

document.addEventListener('click', () => { localDropdown.hidden = true; });

async function loadLocalFile(filename) {
  $('file-name').textContent    = filename;
  $('export-btn').disabled      = true;
  currentResult                 = null;
  setLoading($('pane-raw'));
  setLoading($('pane-pre'));
  setLoading($('result-wrap'));
  $('row-count').textContent    = '';
  $('conf-summary').textContent = '';

  try {
    const res  = await fetch('/api/parse-local', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filename }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || '解析失败');
    currentResult           = data;
    currentResult._filename = filename;
    renderRawTable($('pane-raw'), data.rawRows);
    renderTabPre(data);
    renderResult(data);
    attachSyncScroll();
    $('export-btn').disabled = false;
  } catch (err) {
    console.error('[RFQ2QTR] 本地文件:', err);
    setError($('pane-pre'),    err.message);
    setError($('result-wrap'), err.message);
  }
}

/* ── 读文件为 ArrayBuffer ─────────────────────────────── */
function readAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload  = e => resolve(e.target.result);
    reader.onerror = () => reject(new Error('文件读取失败'));
    reader.readAsArrayBuffer(file);
  });
}

/* ── 主逻辑：Tab1 客户端 + Tab2/右侧 服务端并行 ──────── */
let currentResult = null;
let _preTbody = null, _resultTbody = null;  // 用于跨栏同步高亮
let _rawWrap  = null;    // Tab1 滚动容器
let _preWrap  = null;    // Tab2 滚动容器
let _rawTbody = null;    // Tab1 tbody（用于 hover 同步）
let _scrollCtrl    = null;  // AbortController，防多次加载累积监听器
let _hoverScrolling = 0;    // hover 引发的滚动深度计数，阻止滚动同步反馈

async function handleFile(file) {
  $('file-name').textContent = file.name;
  $('export-btn').disabled = true;
  currentResult = null;

  // 两侧同时显示加载中
  setLoading($('pane-raw'));
  setLoading($('pane-pre'));
  setLoading($('result-wrap'));
  $('row-count').textContent    = '';
  $('conf-summary').textContent = '';

  // Tab1：浏览器端直接解析 Excel
  renderTabRaw(file);

  // Tab2 + 右侧：发往服务端解析
  const fd = new FormData();
  fd.append('file', file);
  try {
    const res  = await fetch('/api/parse', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || '解析失败');
    currentResult = data;
    currentResult._filename = file.name;
    renderTabPre(data);
    renderResult(data);
    attachSyncScroll();
    $('export-btn').disabled = false;
  } catch (err) {
    console.error('[RFQ2QTR] 服务端解析:', err);
    setError($('pane-pre'),   err.message);
    setError($('result-wrap'), err.message);
  }
}

/* ── 跨栏同步高亮（持久容器监听，事件触发时读模块变量）*/
function clearHi(tbody) {
  tbody?.querySelectorAll('tr.sync-hi').forEach(r => r.classList.remove('sync-hi'));
}

/* 把目标面板滚动到 dstTr 与 srcTr 在视口内的 Y 坐标一致 */
function alignRowVisual(srcTr, srcWrap, dstTr, dstWrap) {
  if (!srcTr || !srcWrap || !dstTr || !dstWrap) return;
  const srcTop     = srcTr.getBoundingClientRect().top  - srcWrap.getBoundingClientRect().top;
  const dstCurTop  = dstTr.getBoundingClientRect().top  - dstWrap.getBoundingClientRect().top;
  const newScrollTop = Math.max(0, dstWrap.scrollTop + dstCurTop - srcTop);
  if (Math.abs(dstWrap.scrollTop - newScrollTop) < 1) return;
  _hoverScrolling++;
  dstWrap.scrollTop = newScrollTop;
  setTimeout(() => { _hoverScrolling = Math.max(0, _hoverScrolling - 1); }, 60);
}

// 一次性挂载在持久容器上，事件触发时读模块变量
(function initHoverSync() {
  const R   = $('result-wrap');
  const Rpre = $('pane-pre');
  const Rraw = $('pane-raw');

  // ── 右侧 → Tab2 + Tab1 ───────────────────────────────
  R.addEventListener('mouseover', e => {
    const tr = e.target.closest('tbody tr');
    clearHi(_preTbody); clearHi(_rawTbody);
    if (!tr) return;
    const preMatch = tr.dataset.seq    ? _preTbody?.querySelector(`tr[data-seq="${tr.dataset.seq}"]`)    : null;
    const rawMatch = tr.dataset.srcRow ? _rawTbody?.querySelector(`tr[data-row="${tr.dataset.srcRow}"]`) : null;
    preMatch?.classList.add('sync-hi');
    rawMatch?.classList.add('sync-hi');
    if (preMatch && _preWrap) alignRowVisual(tr, R, preMatch, _preWrap);
    if (rawMatch && _rawWrap) alignRowVisual(tr, R, rawMatch, _rawWrap);
  });
  R.addEventListener('mouseleave', () => { clearHi(_preTbody); clearHi(_rawTbody); });

  // ── Tab2 → 右侧 + Tab1 ───────────────────────────────
  Rpre.addEventListener('mouseover', e => {
    const tr = e.target.closest('tbody tr');
    clearHi(_resultTbody); clearHi(_rawTbody);
    if (!tr?.dataset.seq) return;
    const rMatch  = _resultTbody?.querySelector(`tr[data-seq="${tr.dataset.seq}"]`);
    const rawMatch = rMatch?.dataset.srcRow ? _rawTbody?.querySelector(`tr[data-row="${rMatch.dataset.srcRow}"]`) : null;
    rMatch?.classList.add('sync-hi');
    rawMatch?.classList.add('sync-hi');
    if (rMatch)              alignRowVisual(tr, _preWrap, rMatch,  R);
    if (rawMatch && _rawWrap) alignRowVisual(tr, _preWrap, rawMatch, _rawWrap);
  });
  Rpre.addEventListener('mouseleave', () => { clearHi(_resultTbody); clearHi(_rawTbody); });

  // ── Tab1 → 右侧 + Tab2 ───────────────────────────────
  Rraw.addEventListener('mouseover', e => {
    const tr = e.target.closest('tbody tr');
    clearHi(_resultTbody); clearHi(_preTbody);
    if (!tr?.dataset.row) return;
    const rMatch  = _resultTbody?.querySelector(`tr[data-src-row="${tr.dataset.row}"]`);
    const preMatch = rMatch?.dataset.seq ? _preTbody?.querySelector(`tr[data-seq="${rMatch.dataset.seq}"]`) : null;
    rMatch?.classList.add('sync-hi');
    preMatch?.classList.add('sync-hi');
    if (rMatch)               alignRowVisual(tr, _rawWrap, rMatch,  R);
    if (preMatch && _preWrap) alignRowVisual(tr, _rawWrap, preMatch, _preWrap);
  });
  Rraw.addEventListener('mouseleave', () => { clearHi(_resultTbody); clearHi(_preTbody); });
})();

/* ── 滚动辅助：找第一个（部分）可见行 ───────────────── */
function firstVisibleTr(wrap, sel) {
  const top = wrap.scrollTop;
  for (const tr of wrap.querySelectorAll(sel)) {
    if (tr.offsetTop + tr.offsetHeight > top) return tr;
  }
  return null;
}
function scrollToTr(wrap, tr) {
  if (wrap && tr && wrap.scrollTop !== tr.offsetTop) wrap.scrollTop = tr.offsetTop;
}

/* ── 跨栏滚动同步（行对齐，Tab1 / Tab2 / 右侧三向联动）*/
function attachSyncScroll() {
  if (_scrollCtrl) _scrollCtrl.abort();
  _scrollCtrl = new AbortController();
  const { signal } = _scrollCtrl;
  const R = $('result-wrap');
  let lock = false;

  // Tab1 滚动 → 右侧 + Tab2 按行对齐
  if (_rawWrap) {
    _rawWrap.addEventListener('scroll', () => {
      if (_hoverScrolling || lock) return; lock = true;
      const src = firstVisibleTr(_rawWrap, 'tr[data-row]');
      if (src) {
        const rTr = R.querySelector(`tr[data-src-row="${src.dataset.row}"]`);
        scrollToTr(R, rTr);
        if (rTr && _preWrap) scrollToTr(_preWrap, _preWrap.querySelector(`tr[data-seq="${rTr.dataset.seq}"]`));
      }
      lock = false;
    }, { signal });
  }

  // Tab2 滚动 → 右侧 + Tab1 按行对齐
  if (_preWrap) {
    _preWrap.addEventListener('scroll', () => {
      if (_hoverScrolling || lock) return; lock = true;
      const src = firstVisibleTr(_preWrap, 'tr[data-seq]');
      if (src) {
        const rTr = R.querySelector(`tr[data-seq="${src.dataset.seq}"]`);
        scrollToTr(R, rTr);
        if (rTr && _rawWrap) scrollToTr(_rawWrap, _rawWrap.querySelector(`tr[data-row="${rTr.dataset.srcRow}"]`));
      }
      lock = false;
    }, { signal });
  }

  // 右侧滚动 → Tab1 + Tab2 按行对齐
  R.addEventListener('scroll', () => {
    if (_hoverScrolling || lock) return; lock = true;
    const src = firstVisibleTr(R, 'tr[data-seq]');
    if (src) {
      if (_preWrap) scrollToTr(_preWrap, _preWrap.querySelector(`tr[data-seq="${src.dataset.seq}"]`));
      if (_rawWrap && src.dataset.srcRow) scrollToTr(_rawWrap, _rawWrap.querySelector(`tr[data-row="${src.dataset.srcRow}"]`));
    }
    lock = false;
  }, { signal });
}

/* ── 辅助：loading / error 占位 ─────────────────────────*/
function setLoading(el) {
  el.innerHTML = '<div class="loading"><div class="spinner"></div>处理中…</div>';
}
function setError(el, msg) {
  el.innerHTML = `<div class="error-hint">❌ ${esc(msg)}</div>`;
}

/* ═══════════════════════════════════════════════════════
   Tab 1：原始 Excel（浏览器端解析）
   ═══════════════════════════════════════════════════════ */
async function renderTabRaw(file) {
  const pane = $('pane-raw');
  try {
    const buf  = await readAsArrayBuffer(file);
    const wb   = XLSX.read(new Uint8Array(buf), { type: 'array' });
    if (!wb.SheetNames.length) throw new Error('未找到工作表');
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
    renderRawTable(pane, rows);
    attachSyncScroll();  // Tab1 完成后重新绑定滚动同步，确保 _rawWrap 被纳入
  } catch (err) {
    setError(pane, err.message);
  }
}

function renderRawTable(pane, rows) {
  if (!rows.length) { pane.innerHTML = '<div class="drop-hint"><div>工作表为空</div></div>'; return; }

  const colCount = rows.reduce((m, r) => Math.max(m, r.length), 0);
  if (!colCount)   { pane.innerHTML = '<div class="drop-hint"><div>未检测到列数据</div></div>'; return; }

  const table = document.createElement('table');

  // 表头
  const hdrRow = table.createTHead().insertRow();
  addTh(hdrRow, '', 'row-num-col');
  for (let c = 0; c < colCount; c++) addTh(hdrRow, colLabel(c));

  // 数据行
  const tbody = table.createTBody();
  rows.forEach((row, ri) => {
    const tr = tbody.insertRow();
    tr.dataset.row = String(ri + 1);  // Excel 行号，供 hover 同步使用
    addTd(tr, ri + 1, 'row-num-col');
    for (let c = 0; c < colCount; c++) addTd(tr, row[c] ?? '');
  });

  _rawTbody = tbody;
  _rawWrap  = mountTable(pane, table);
}

/* ═══════════════════════════════════════════════════════
   Tab 2：预处理结果（服务端返回）
   14 个 UNI 属性 + 原始拼接描述
   ═══════════════════════════════════════════════════════ */

// 14 个产品描述属性（顺序与 Lexer.yaml uni_to_code_map 对应）
const UNI_FIELDS = [
  ['UNI_TYPE',            '类型'],
  ['UNI_ANGLE',           '角度'],
  ['UNI_RADIUS',          '曲率'],
  ['UNI_MISC',            '其他'],
  ['UNI_OD1',             'OD1'],
  ['UNI_OD2',             'OD2'],
  ['UNI_OD3',             'OD3'],
  ['UNI_WT1',             'WT1'],
  ['UNI_WT2',             'WT2'],
  ['UNI_WT3',             'WT3'],
  ['UNI_DIM_SPEC',        '尺寸标准'],
  ['UNI_MATERIAL',        '材质'],
  ['UNI_CONSTRUCTION',    '结构'],
  ['UNI_END_PREPARATION', '端口'],
];

function renderTabPre(data) {
  const pane = $('pane-pre');
  const rows = (data.preprocessRows || []).filter(r => r.type !== 'skip');
  if (!rows.length) { pane.innerHTML = '<div class="drop-hint"><div>无预处理结果</div></div>'; return; }

  const table  = document.createElement('table');
  const hdrRow = table.createTHead().insertRow();
  addTh(hdrRow, '#',    'row-num-col');
  addTh(hdrRow, '行');
  addTh(hdrRow, '类型');
  addTh(hdrRow, '完整描述');
  UNI_FIELDS_COMPACT.forEach(([, label]) => addTh(hdrRow, label));
  addTh(hdrRow, '置信度');

  const tbody = table.createTBody();
  let seq = 0;
  rows.forEach(r => {
    const tr = tbody.insertRow();
    if (r.type === 'parent')  tr.classList.add('row-parent');
    if (r.type === 'summary') tr.classList.add('row-summary');
    if (r.type === 'child') {
      seq++;
      tr.dataset.seq = String(seq);
      if (r.conf) tr.style.background = CONF_STYLE[r.conf]?.bg || '';
    }

    addTd(tr, r.type === 'child' ? seq : '', 'row-num-col');
    addTd(tr, r.index + 1);

    const tdType = tr.insertCell();
    tdType.innerHTML = typeBadge(r.type);

    const tdDesc = tr.insertCell();
    tdDesc.className   = 'desc-cell';
    tdDesc.textContent = r.descText || '';
    tdDesc.title       = r.descText || '';

    UNI_FIELDS_COMPACT.forEach(([field]) => {
      const td = tr.insertCell();
      const v  = r.uni?.[field] || '';
      td.textContent = v;
      td.className   = v ? 'uni-hit' : 'uni-miss';
    });

    const tdConf = tr.insertCell();
    if (r.type === 'child' && r.conf) {
      const s = CONF_STYLE[r.conf] || CONF_STYLE.mid;
      tdConf.innerHTML = `<span class="badge ${s.cls}">${s.badge}</span>`;
      if (r.reason) tdConf.title = r.reason;
    }
  });

  _preTbody = tbody;
  _preWrap = mountTable(pane, table);
}

function typeBadge(type) {
  const map = {
    parent:  ['badge-parent',  '父行'],
    child:   ['badge-child',   '产品行'],
    summary: ['badge-summary', '汇总'],
  };
  const [cls, label] = map[type] ?? ['badge-skip', type];
  return `<span class="badge ${cls}">${label}</span>`;
}

/* ═══════════════════════════════════════════════════════
   右侧：处理结果
   ═══════════════════════════════════════════════════════ */
function renderResult(data) {
  const wrap = $('result-wrap');
  const rows = data.enrichedRows || [];
  if (!rows.length) {
    wrap.innerHTML = '<div class="drop-hint"><div class="drop-icon">📋</div><div>无处理结果</div></div>';
    $('row-count').textContent = '';
    return;
  }

  $('row-count').textContent = `共 ${rows.length} 行`;
  const counts = { high: 0, mid: 0, low: 0 };
  rows.forEach(r => { counts[r.conf] = (counts[r.conf] || 0) + 1; });
  $('conf-summary').textContent =
    `高 ${counts.high}  中 ${counts.mid}  待审 ${counts.low}`;

  const cols   = ['item','category','qty','price','type','angle','radius','misc','od1','od2','od3','wt','wt2','wt3','dimSpec','material','construction','ends','pressure','facing','boltGrade','boltLength','threadType','gasketType','gasketThk','conf'];
  const labels = ['序号','品类','数量','订单价格','产品类型','角度','曲率','其他属性','OD1','OD2','OD3','WT1','WT2','WT3','尺寸标准','产品材质','产品结构','端口形式','压力等级','密封面','螺栓等级','螺栓长度','螺纹类型','垫片类型','垫片厚度','置信度'];

  const table  = document.createElement('table');
  const hdrRow = table.createTHead().insertRow();
  labels.forEach(l => addTh(hdrRow, l));

  const tbody = table.createTBody();
  rows.forEach((row, i) => {
    const tr = tbody.insertRow();
    tr.dataset.seq    = String(i + 1);
    tr.dataset.srcRow = String(row.srcRow);  // Excel 行号，与 Tab1 data-row 对应

    const style = CONF_STYLE[row.conf] || CONF_STYLE.mid;
    tr.style.background = style.bg;
    if (row.reason) tr.title = row.reason;

    cols.forEach(col => {
      const td = tr.insertCell();
      if (col === 'conf') {
        td.innerHTML = `<span class="badge ${style.cls}">${style.badge}</span>`;
        return;
      }
      if (col === 'category') {
        td.textContent = row.category || 'FITTING';
        td.title       = row.category || 'FITTING';
        return;
      }
      td.textContent = row[col] ?? '';
      td.title       = String(row[col] ?? '');
      if ((col === 'material' || col === 'od1') && row.reason?.includes(`"${row[col]}"`)) {
        td.style.color = '#a4571f';
      }
      if (EDITABLE_COLS.includes(col)) {
        td.contentEditable = 'true';
        td.addEventListener('blur', () => {
          row[col] = td.textContent.trim();
          if (row.conf !== 'high') {
            row.conf   = 'confirmed';
            row.reason = '已人工确认';
            tr.style.background = CONF_STYLE.confirmed.bg;
            const badge = tr.querySelector('.badge');
            badge.className   = `badge ${CONF_STYLE.confirmed.cls}`;
            badge.textContent = CONF_STYLE.confirmed.badge;
          }
          refreshConfSummary(rows);
        });
      }
    });
  });

  _resultTbody = tbody;
  wrap.innerHTML = '';
  wrap.appendChild(table);
}

function refreshConfSummary(rows) {
  const c = { high: 0, mid: 0, low: 0, confirmed: 0 };
  rows.forEach(r => { c[r.conf] = (c[r.conf] || 0) + 1; });
  $('conf-summary').textContent =
    `高 ${c.high + c.confirmed}  中 ${c.mid}  待审 ${c.low}`;
}

/* ── 导出 ─────────────────────────────────────────────── */
$('export-btn').addEventListener('click', async () => {
  const rows = currentResult?.enrichedRows;
  if (!rows?.length) return;

  const confLabel = { high:'高', mid:'中', low:'低·待审', confirmed:'已确认' };
  const exportData = rows.map(r => ({
    '序号':     r.item,
    '品类':     r.category     || 'FITTING',
    '数量':     r.qty          || '',
    '订单价格': r.price        || '',
    '产品类型': r.type         || '',
    '角度':     r.angle        || '',
    '曲率':     r.radius       || '',
    '其他属性': r.misc         || '',
    'OD1':      r.od1          || '',
    'OD2':      r.od2          || '',
    'OD3':      r.od3          || '',
    'WT1':      r.wt           || '',
    'WT2':      r.wt2          || '',
    'WT3':      r.wt3          || '',
    '尺寸标准': r.dimSpec      || '',
    '产品材质': r.material     || '',
    '产品结构': r.construction || '',
    '端口形式': r.ends         || '',
    '压力等级': r.pressure     || '',
    '密封面':   r.facing       || '',
    '螺栓等级': r.boltGrade    || '',
    '螺栓长度': r.boltLength   || '',
    '螺纹类型': r.threadType   || '',
    '垫片类型': r.gasketType   || '',
    '垫片厚度': r.gasketThk    || '',
    '置信度':   confLabel[r.conf] || r.conf,
    '待审原因': r.reason       || '',
  }));

  const filename = (currentResult._filename || 'result').replace(/\.[^.]+$/, '');
  try {
    const res = await fetch('/api/export', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ rows: exportData, filename }),
    });
    if (!res.ok) { setError($('result-wrap'), '导出失败：' + (await res.json()).error); return; }
    const url = URL.createObjectURL(await res.blob());
    Object.assign(document.createElement('a'), {
      href: url, download: filename + '_QTR.xlsx',
    }).click();
    URL.revokeObjectURL(url);
  } catch (e) { setError($('result-wrap'), '导出失败：' + e.message); }
});

/* ── DOM 辅助 ─────────────────────────────────────────── */
function addTh(tr, text, cls) {
  const th = document.createElement('th');
  th.textContent = String(text);
  if (cls) th.className = cls;
  tr.appendChild(th);
}
function addTd(tr, text, cls) {
  const td = tr.insertCell();
  const v  = String(text ?? '');
  td.textContent = v;
  td.title       = v;
  if (cls) td.className = cls;
  return td;
}
function mountTable(pane, table) {
  const wrap = document.createElement('div');
  wrap.className = 'table-wrap';
  wrap.appendChild(table);
  pane.innerHTML = '';
  pane.appendChild(wrap);
  return wrap;
}

/* 生成 Excel 风格列标签：0→A, 25→Z, 26→AA … */
function colLabel(n) {
  let s = '';
  for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}
