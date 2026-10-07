import express from 'express';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

const DATA_DIR = path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

const DASHBOARD_FILE = path.join(DATA_DIR, 'dashboard.json');
const REMARKS_FILE = path.join(DATA_DIR, 'remarks.json');

// In-memory caches for maximum speed across all connected PCs
let cachedDashboard: { data: any[] | null; lastUpdated: number | null; uploadId: string | null } = {
  data: null,
  lastUpdated: null,
  uploadId: null,
};

let cachedRemarks: Record<string, { text: string; updatedAt: number | null }> = {};

// Load on startup
try {
  if (fs.existsSync(DASHBOARD_FILE)) {
    const raw = fs.readFileSync(DASHBOARD_FILE, 'utf-8');
    cachedDashboard = JSON.parse(raw);
  }
} catch (e) {
  console.error('Error reading dashboard file:', e);
}

try {
  if (fs.existsSync(REMARKS_FILE)) {
    const raw = fs.readFileSync(REMARKS_FILE, 'utf-8');
    cachedRemarks = JSON.parse(raw);
  }
} catch (e) {
  console.error('Error reading remarks file:', e);
}

// API Routes
app.get('/api/dashboard', (req, res) => {
  res.json(cachedDashboard);
});

app.post('/api/dashboard', (req, res) => {
  try {
    const { data, lastUpdated, uploadId } = req.body;
    cachedDashboard = {
      data: Array.isArray(data) ? data : null,
      lastUpdated: lastUpdated || Date.now(),
      uploadId: uploadId || Date.now().toString(),
    };
    fs.writeFileSync(DASHBOARD_FILE, JSON.stringify(cachedDashboard));
    res.json({ success: true, count: cachedDashboard.data?.length || 0 });
  } catch (err: any) {
    console.error('Error saving dashboard:', err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/remarks', (req, res) => {
  res.json(cachedRemarks);
});

app.post('/api/remarks', (req, res) => {
  try {
    const { id, text, updatedAt } = req.body;
    if (id) {
      cachedRemarks[id] = {
        text: text || '',
        updatedAt: updatedAt || Date.now(),
      };
      fs.writeFileSync(REMARKS_FILE, JSON.stringify(cachedRemarks));
    }
    res.json({ success: true, id });
  } catch (err: any) {
    console.error('Error saving remark:', err);
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/remarks/batch', (req, res) => {
  try {
    const { remarks } = req.body;
    if (remarks && typeof remarks === 'object') {
      Object.assign(cachedRemarks, remarks);
      fs.writeFileSync(REMARKS_FILE, JSON.stringify(cachedRemarks));
    }
    res.json({ success: true });
  } catch (err: any) {
    console.error('Error saving batch remarks:', err);
    res.status(500).json({ error: err.message });
  }
});

// Mount Vite or static server
async function startServer() {
  if (process.env.NODE_ENV === 'production' && fs.existsSync(path.join(__dirname, 'dist'))) {
    app.use(express.static(path.join(__dirname, 'dist')));
    app.get('*', (req, res) => {
      res.sendFile(path.join(__dirname, 'dist', 'index.html'));
    });
  } else {
    const { createServer } = await import('vite');
    const vite = await createServer({
      server: { middlewareMode: true, port: Number(PORT) },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  }

  app.listen(Number(PORT), '0.0.0.0', () => {
    console.log(`Production Dashboard Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
