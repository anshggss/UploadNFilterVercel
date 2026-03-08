import express from 'express';
import multer from 'multer';
import { filterExcel } from './filterExcel.js';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import cors from 'cors';

const app = express();

// allow cross‑origin requests from the client (deploy origin or any)
app.use(cors({
  origin: process.env.CLIENT_URL || '*' // set CLIENT_URL when deploying if you want to lock it down
}));

// Optimized multer configuration with memory storage for better performance
const upload = multer({ 
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024, // 50MB limit
    files: 2
  }
});

const PORT = process.env.PORT || 5901;
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);


if (process.env.SERVE_STATIC === 'true') {
  app.use(express.static(path.join(__dirname, "../client/dist")));
  app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, "../client/index.html"));
  });
}

// Add error handling middleware
app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    if (err.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({ error: 'File too large. Maximum size is 50MB.' });
    }
    if (err.code === 'LIMIT_FILE_COUNT') {
      return res.status(400).json({ error: 'Too many files. Maximum is 2 files.' });
    }
  }
  next(err);
});



app.post('/api/filter', upload.fields([
  { name: 'file', maxCount: 1 },
  { name: 'custData', maxCount: 1 }
]), async (req, res) => {
  let tempFiles = [];
  
  try {
  if (!req.files || !req.files['file'] || !req.files['custData']) {
    return res.status(400).send('Both files must be uploaded');
  }
    
    // Create temporary files from memory buffers
    const mainFile = req.files['file'][0];
    const custDataFile = req.files['custData'][0];
    
    const mainFilePath = path.join(__dirname, 'temp', `main_${Date.now()}_${Math.random().toString(36).substr(2, 9)}.xlsx`);
    const custDataFilePath = path.join(__dirname, 'temp', `cust_${Date.now()}_${Math.random().toString(36).substr(2, 9)}.xlsx`);
    
    // Ensure temp directory exists
    const tempDir = path.join(__dirname, 'temp');
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }
    
    // Write buffers to temporary files
    fs.writeFileSync(mainFilePath, mainFile.buffer);
    fs.writeFileSync(custDataFilePath, custDataFile.buffer);
    
    tempFiles = [mainFilePath, custDataFilePath];
    
    const filteredBuffer = await filterExcel(mainFilePath, custDataFilePath);
    
    res.setHeader('Content-Disposition', 'attachment; filename="filtered.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(filteredBuffer);
    
  } catch (err) {
    console.error('Error filtering file:', err);
    res.status(500).send('Error filtering file: ' + err.message);
  } finally {
    // Clean up uploaded files
    tempFiles.forEach(filePath => {
      if (fs.existsSync(filePath)) {
        fs.unlink(filePath, (err) => {
          if (err) console.error('Error deleting temp file:', err);
        });
      }
    });
  }
});


app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
}); 
