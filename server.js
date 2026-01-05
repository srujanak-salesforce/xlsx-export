const ExcelJS = require('exceljs');
const express = require('express');
const app = express();

app.use(express.json({ limit: '25mb' }));

let workbook;
let worksheet;

app.post('/append', (req, res) => {

  if (!workbook) {
    workbook = new ExcelJS.Workbook();   // ✅ NOT streaming
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

  const buffer = await workbook.xlsx.writeBuffer(); // ✅ VALID NOW

  // cleanup
  workbook = null;
  worksheet = null;

  res.json({
    base64: Buffer.from(buffer).toString('base64')
  });
});

app.listen(process.env.PORT || 3000, () =>
  console.log('✅ XLSX service ready')
);
