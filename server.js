const ExcelJS = require('exceljs');
const express = require('express');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.json({ limit: '100mb' })); // safe for 40k+ rows

let workbook;
let worksheet;
let exportFilePath;

app.post('/append', async (req, res) => {
  try {
    if (!workbook) {
      const fileName = `Accounts_Export_${Date.now()}.xlsx`;
      exportFilePath = path.join('/tmp', fileName); // ✅ REQUIRED for Render

      workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
        filename: exportFilePath,
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
    res.status(500).send(e.message);
  }
});

app.post('/finalize', async (req, res) => {
  try {
    worksheet.commit();
    await workbook.commit();

    // reset memory
    workbook = null;
    worksheet = null;

    res.json({
      downloadUrl: `https://xlsx-export.onrender.com/download`
    });
  } catch (e) {
    console.error('Finalize error', e);
    res.status(500).send(e.message);
  }
});

app.get('/download', (req, res) => {
  res.download(exportFilePath);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () =>
  console.log(`✅ XLSX Streaming Export Service running on port ${PORT}`)
);
