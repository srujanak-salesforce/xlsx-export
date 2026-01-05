const ExcelJS = require('exceljs');
const express = require('express');
const app = express();

app.use(express.json({ limit: '50mb' }));

let workbook;
let worksheet;
let bufferChunks = [];

app.post('/append', async (req, res) => {
  if (!workbook) {
    workbook = new ExcelJS.Workbook();
    worksheet = workbook.addWorksheet('Accounts');
  }

  const rows = req.body;

  if (!worksheet.columns && rows.length) {
    worksheet.columns = Object.keys(rows[0]).map(k => ({
      header: k,
      key: k,
      width: 25
    }));
  }

  rows.forEach(r => worksheet.addRow(r));
  res.sendStatus(200);
});

app.post('/finalize', async (req, res) => {
  const buffer = await workbook.xlsx.writeBuffer();

  // 🔴 VERY IMPORTANT: reset memory
  workbook = null;
  worksheet = null;

  res.json({
    base64: Buffer.from(buffer).toString('base64')
  });
});

app.listen(3000, () =>
  console.log('✅ XLSX export service running on port 3000')
);
