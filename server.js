const ExcelJS = require('exceljs');
const express = require('express');
const app = express();

app.use(express.json({ limit: '25mb' }));

let workbook;
let worksheet;

app.post('/append', (req, res) => {

  if (!workbook) {
    workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
      filename: 'Accounts_Ancillary_Export.xlsx'
    });
    worksheet = workbook.addWorksheet('Accounts');
  }

  const rows = req.body;
  if (!worksheet.columns && rows.length) {
    worksheet.columns = Object.keys(rows[0]).map(k => ({
      header: k, key: k, width: 25
    }));
  }

  rows.forEach(r => worksheet.addRow(r).commit());
  res.sendStatus(200);
});

app.post('/finalize', async (req, res) => {
  await workbook.commit();
  workbook = null;
  worksheet = null;
  res.sendStatus(200);
});

app.listen(3000, () => console.log('✅ XLSX streaming service ready'));
