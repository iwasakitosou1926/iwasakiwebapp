import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import multer from 'multer';
import fs from 'fs';
import * as notion from './src/server/notion.js';
import * as workflow from './src/server/workflow.js';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const port = 3001;

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, __dirname); // Save to the project root
  },
  filename: (req, file, cb) => {
    cb(null, file.originalname); // Keep original name
  }
});
const upload = multer({ storage });

app.use(express.json());

// Initialize Notion SDK on startup if possible
try {
  notion.initNotion();
} catch (e) {
  console.warn('Could not initialize Notion on startup:', e);
}

// Upload endpoint
app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }
  res.json({ filename: req.file.originalname });
});

// API Endpoints
app.get('/api/calendar', async (req, res) => {
  try {
    const data = await notion.getCalendarPagesNextN();
    res.json(data);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/load-products', async (req, res) => {
  const { filePath } = req.body;
  if (!filePath) {
    return res.status(400).json({ error: 'filePath is required' });
  }
  try {
    const fullPath = path.isAbsolute(filePath) ? filePath : path.join(__dirname, filePath);
    const data = await workflow.loadProductsFromExcel(fullPath);
    res.json(data);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/sync', async (req, res) => {
  try {
    const { file_path, page_id, products } = req.body;
    if (!file_path || !page_id) throw new Error('file_path and page_id are required');
    
    const fullPath = path.isAbsolute(file_path) ? file_path : path.join(__dirname, file_path);
    const data = await workflow.highlightAndSync(fullPath, page_id, products || []);
    res.json(data);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
});

// Serve frontend in production (optional for now)
// app.use(express.static(path.join(__dirname, 'dist')));

app.listen(port, () => {
  console.log(`Backend server running at http://localhost:${port}`);
});
