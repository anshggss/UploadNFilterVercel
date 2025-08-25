import { filterExcel } from '../server/utils/filterExcel.js';
import formidable from 'formidable';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Configure formidable for file uploads
const form = formidable({
  maxFileSize: 50 * 1024 * 1024, // 50MB
  maxFiles: 2,
  keepExtensions: true,
  uploadDir: '/tmp'
});

export default async function handler(req, res) {
  // Enable CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  let tempFiles = [];

  try {
    // Parse the multipart form data
    const [fields, files] = await form.parse(req);
    
    const mainFile = files.file?.[0];
    const custDataFile = files.custData?.[0];

    if (!mainFile || !custDataFile) {
      res.status(400).json({ error: 'Both files must be uploaded' });
      return;
    }

    // Create unique temporary file paths
    const timestamp = Date.now();
    const randomId = Math.random().toString(36).substr(2, 9);
    const mainFilePath = path.join('/tmp', `main_${timestamp}_${randomId}.xlsx`);
    const custDataFilePath = path.join('/tmp', `cust_${timestamp}_${randomId}.xlsx`);

    // Copy uploaded files to our temp paths
    fs.copyFileSync(mainFile.filepath, mainFilePath);
    fs.copyFileSync(custDataFile.filepath, custDataFilePath);
    
    tempFiles = [mainFile.filepath, custDataFile.filepath, mainFilePath, custDataFilePath];

    // Process the Excel files
    const filteredBuffer = await filterExcel(mainFilePath, custDataFilePath);

    // Set response headers for file download
    res.setHeader('Content-Disposition', 'attachment; filename="filtered.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Length', filteredBuffer.length);
    
    res.status(200).send(filteredBuffer);

  } catch (error) {
    console.error('Error processing files:', error);
    res.status(500).json({ 
      error: 'Error processing files: ' + error.message 
    });
  } finally {
    // Clean up temporary files
    tempFiles.forEach(filePath => {
      if (filePath && fs.existsSync(filePath)) {
        try {
          fs.unlinkSync(filePath);
        } catch (err) {
          console.error('Error deleting temp file:', err);
        }
      }
    });
  }
}