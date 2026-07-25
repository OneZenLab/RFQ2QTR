'use strict';

const express = require('express');
const multer  = require('multer');
const path    = require('path');
const fs      = require('fs');
const XLSX    = require('xlsx');

const { parseExcel } = require('./parser');

const app    = express();
const upload = multer({ dest: path.join(__dirname, 'uploads') });

// 静态文件服务（前端页面），禁用缓存方便迭代调试
app.use(express.static(path.join(__dirname, 'public'), { etag: false, maxAge: 0 }));
app.use(express.json({ limit: '50mb' }));

// 向浏览器暴露 xlsx 库
app.get('/xlsx.js', (_req, res) => {
  res.sendFile(path.join(__dirname, 'node_modules/xlsx/dist/xlsx.full.min.js'));
});

// GET /api/local-files — 列出 QTR/ 目录中的 xlsx 文件
app.get('/api/local-files', (_req, res) => {
  const dir = path.join(__dirname, 'QTR');
  try {
    const files = fs.readdirSync(dir)
      .filter(f => /\.xlsx$/i.test(f) && !f.startsWith('~$'))
      .map(f => {
        const st = fs.statSync(path.join(dir, f));
        return { name: f, size: st.size, mtime: st.mtime };
      });
    res.json(files);
  } catch (e) {
    res.status(500).json({ error: `读取目录失败：${e.message}` });
  }
});

// POST /api/parse-local — 直接解析 QTR/ 目录中的本地文件
app.post('/api/parse-local', (req, res) => {
  const { filename } = req.body;
  if (!filename) return res.status(400).json({ error: '缺少 filename 参数' });

  // 路径越界检查
  const fullPath = path.resolve(path.join(__dirname, 'QTR', filename));
  if (!fullPath.startsWith(path.resolve(path.join(__dirname, 'QTR')))) {
    return res.status(400).json({ error: '非法文件路径' });
  }
  if (!fs.existsSync(fullPath)) {
    return res.status(404).json({ error: `文件不存在：${filename}` });
  }

  try {
    const result = parseExcel(fullPath);
    console.log(`\n=== 本地文件解析：${filename} ===`);
    console.log('输出行数:', result.outputRows.length);
    res.json(result);
  } catch (e) {
    console.error('解析异常:', e);
    res.status(500).json({ error: `解析失败：${e.message}` });
  }
});

// POST /api/parse — 上传 Excel 并解析第1张表
app.post('/api/parse', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: '未上传文件' });
  }
  const tmpPath = req.file.path;
  try {
    const result = parseExcel(tmpPath);
    // 调试：打印关键信息到服务端控制台
    console.log('\n=== 解析结果摘要 ===');
    console.log('表头行号:', result.headerRow);
    console.log('列名:', result.headers);
    console.log('列语义映射:', result.colSemantics);
    const sampleDesc = result.preprocessRows
      .filter(r => r.type === 'child').slice(0, 3)
      .map(r => `  行${r.index}: descText="${r.descText}" | TYPE=${r.uni?.UNI_TYPE} OD1=${r.uni?.UNI_OD1}`);
    console.log('前3个产品行:\n' + sampleDesc.join('\n'));
    console.log('输出行数:', result.outputRows.length);
    res.json(result);
  } catch (e) {
    console.error('解析异常:', e);
    res.status(500).json({ error: `解析失败：${e.message}` });
  } finally {
    // 解析完成后删除临时文件
    fs.unlink(tmpPath, () => {});
  }
});

// POST /api/export — 将结果行导出为 Excel 文件
app.post('/api/export', (req, res) => {
  const { rows } = req.body;
  if (!rows?.length) {
    return res.status(400).json({ error: '无可导出数据' });
  }
  try {
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '解析结果');
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.set('Content-Disposition', 'attachment; filename="result.xlsx"');
    res.send(buf);
  } catch (e) {
    res.status(500).json({ error: `导出失败：${e.message}` });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`RFQ2QTR 服务已启动：http://localhost:${PORT}`);
});
