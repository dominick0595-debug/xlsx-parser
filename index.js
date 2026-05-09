const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const iconv = require('iconv-lite');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.post('/parse', upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const csv = XLSX.utils.sheet_to_csv(sheet);
    
    // Декодируем через iconv
    const decoded = iconv.decode(Buffer.from(csv, 'binary'), 'cp1251');
    
    res.json({ csv: decoded });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
