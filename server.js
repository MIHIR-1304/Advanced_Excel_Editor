
const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const app = express();
app.use(cors());
app.use(express.json());



const DATA_DIR = path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}


app.get('/api/test', (req, res) => {
  res.json({ message: "Backend is working!" });
});


app.post('/api/save', (req, res) => {
  try {
    console.log("Received save request:", req.body);
    
    if (!req.body.filename || !req.body.data) {
      throw new Error("Filename and data are required");
    }
    
    const filePath = path.join(DATA_DIR, req.body.filename);
    fs.writeFileSync(filePath, JSON.stringify(req.body.data));
    
    console.log("File saved successfully at:", filePath);
    res.json({ 
      success: true,
      filename: req.body.filename
    });
  } catch (error) {
    console.error("Save error:", error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});
const XLSX = require('xlsx');


app.post('/api/save-excel', (req, res) => {
  try {
    console.log("Received Excel save request:", req.body);
    
    if (!req.body.filename || !req.body.data) {
      throw new Error("Filename and data are required");
    }
    
    
    const wb = XLSX.utils.book_new();
    
    
    const wsData = [
      req.body.data.headers, // Header row
      ...req.body.data.rows  // Data rows
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    
    
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    
    
    const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    
    
    const filePath = path.join(DATA_DIR, `${req.body.filename}.xlsx`);
    fs.writeFileSync(filePath, excelBuffer);
    
    console.log("Excel file saved at:", filePath);
    res.json({ 
      success: true,
      filename: `${req.body.filename}.xlsx`,
      path: filePath
    });
  } catch (error) {
    console.error("Excel save error:", error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});
// File download endpoint
app.get('/downloads/:filename', (req, res) => {
  const filePath = path.join(DATA_DIR, req.params.filename);
  
  if (fs.existsSync(filePath)) {
    res.download(filePath, req.params.filename, (err) => {
      if (err) {
        console.error('Download error:', err);
        res.status(500).send('Download failed');
      }
    });
  } else {
    res.status(404).send('File not found');
  }
});
const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});