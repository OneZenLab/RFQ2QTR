'use strict';

/* ── 工具 ─────────────────────────────────────────────── */
function esc(v) {
  return String(v ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function $(id) { return document.getElementById(id); }

/* ── Tab 切换 ─────────────────────────────────────────── */
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    const t = tab.dataset.tab;
    document.querySelectorAll('.tab').forEach(b => b.classList.toggle('active', b === tab));
    document.querySelectorAll('.pane').forEach(p => p.classList.toggle('active', p.id === `pane-${t}`));
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

async function handleFile(file) {
  $('file-name').textContent = file.name;
  $('export-btn').disabled = true;
  currentResult = null;

  // 两侧同时显示加载中
  setLoading($('pane-raw'));
  setLoading($('pane-pre'));
  setLoading($('result-wrap'));
  $('row-count').textContent = '';

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
    renderTabPre(data);
    renderResult(data);
    $('export-btn').disabled = false;
  } catch (err) {
    console.error('[RFQ2QTR] 服务端解析:', err);
    setError($('pane-pre'),   err.message);
    setError($('result-wrap'), err.message);
  }
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
    addTd(tr, ri + 1, 'row-num-col');
    for (let c = 0; c < colCount; c++) addTd(tr, row[c] ?? '');
  });

  mountTable(pane, table);
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

  const table = document.createElement('table');

  // 表头：仅序号、行号、类型、完整描述
  const hdrRow = table.createTHead().insertRow();
  addTh(hdrRow, '#',    'row-num-col');
  addTh(hdrRow, '行');
  addTh(hdrRow, '类型');
  addTh(hdrRow, '完整描述');

  // 数据行
  const tbody = table.createTBody();
  let seq = 0;
  rows.forEach(r => {
    const tr = tbody.insertRow();
    if (r.type === 'parent')  tr.classList.add('row-parent');
    if (r.type === 'summary') tr.classList.add('row-summary');

    if (r.type === 'child') seq++;
    addTd(tr, r.type === 'child' ? seq : '', 'row-num-col');
    addTd(tr, r.index + 1);

    // 类型徽章
    const tdType = tr.insertCell();
    tdType.innerHTML = typeBadge(r.type);

    // 完整描述
    const tdDesc = tr.insertCell();
    tdDesc.className   = 'desc-cell';
    tdDesc.textContent = r.descText || '';
    tdDesc.title       = r.descText || '';
  });

  mountTable(pane, table);
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
  const rows = data.outputRows || [];
  if (!rows.length) { wrap.innerHTML = '<div class="drop-hint"><div class="drop-icon">📋</div><div>无处理结果</div></div>'; $('row-count').textContent = ''; return; }

  $('row-count').textContent = `共 ${rows.length} 行`;
  const cols  = Object.keys(rows[0]);
  const table = document.createElement('table');

  const hdrRow = table.createTHead().insertRow();
  cols.forEach(c => addTh(hdrRow, c));

  const tbody = table.createTBody();
  rows.forEach(row => {
    const tr = tbody.insertRow();
    cols.forEach(c => addTd(tr, row[c] ?? ''));
  });

  wrap.innerHTML = '';
  const div = document.createElement('div');
  div.className = 'table-wrap';
  div.appendChild(table);
  wrap.appendChild(div);
}

/* ── 导出 ─────────────────────────────────────────────── */
$('export-btn').addEventListener('click', async () => {
  if (!currentResult?.outputRows?.length) return;
  try {
    const res = await fetch('/api/export', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ rows: currentResult.outputRows }),
    });
    if (!res.ok) { alert('导出失败：' + (await res.json()).error); return; }
    const url = URL.createObjectURL(await res.blob());
    Object.assign(document.createElement('a'), { href: url, download: 'result.xlsx' }).click();
    URL.revokeObjectURL(url);
  } catch (e) { alert('导出失败：' + e.message); }
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
}

/* 生成 Excel 风格列标签：0→A, 25→Z, 26→AA … */
function colLabel(n) {
  let s = '';
  for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}
