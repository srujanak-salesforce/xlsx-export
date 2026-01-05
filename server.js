const ExcelJS = require('exceljs');
const express = require('express');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.json({ limit: '50mb' }));

let workbook;
let worksheet;
let exportFilePath;

app.post('/append', async (req, res) => {
  try {
    if (!workbook) {
      workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
        filename: null,
        useStyles: true,
        useSharedStrings: true
      });

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

    rows.forEach(r => worksheet.addRow(r).commit());

    res.sendStatus(200);
  } catch (e) {
    console.error('Append error', e);
    res.sendStatus(500);
  }
});

app.post('/finalize', async (req, res) => {
  try {
    const fileName = `Accounts_Export_${Date.now()}.xlsx`;
    exportFilePath = path.join(__dirname, fileName);

    workbook.stream = fs.createWriteStream(exportFilePath);
    worksheet.commit();
    await workbook.commit();

    // 🔥 reset memory
    workbook = null;
    worksheet = null;

    res.json({
      downloadUrl: `https://xlsx-export.onrender.com/download/${fileName}`
    });

  } catch (e) {
    console.error('Finalize error', e);
    res.sendStatus(500);
  }
});

app.get('/download/:file', (req, res) => {
  const filePath = path.join(__dirname, req.params.file);
  res.download(filePath);
});

app.listen(3000, () =>
  console.log('✅ XLSX Streaming Export Service running on port 3000')
);
