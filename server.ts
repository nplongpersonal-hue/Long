import express from 'express';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function resolveWritableStoreDir() {
  const localDir = path.join(__dirname, 'templates_store');
  try {
    if (!fs.existsSync(localDir)) {
      fs.mkdirSync(localDir, { recursive: true });
    }
    fs.accessSync(localDir, fs.constants.W_OK);
    return localDir;
  } catch {
    const tmpDir = path.join(os.tmpdir(), 'templates_store');
    try {
      if (!fs.existsSync(tmpDir)) {
        fs.mkdirSync(tmpDir, { recursive: true });
      }
      // Seed from bundled templates_store if available
      if (fs.existsSync(localDir)) {
        for (const file of ['config.json', 'history.json', 'drive_config.json']) {
          const src = path.join(localDir, file);
          const dest = path.join(tmpDir, file);
          if (fs.existsSync(src) && !fs.existsSync(dest)) {
            fs.copyFileSync(src, dest);
          }
        }
      }
    } catch {
      // Ignore if tmpDir also fails; in-memory fallback is used below
    }
    return tmpDir;
  }
}

const STORE_DIR = resolveWritableStoreDir();
const CONFIG_FILE = path.join(STORE_DIR, 'config.json');
const HISTORY_FILE = path.join(STORE_DIR, 'history.json');
const DRIVE_CONFIG_FILE = path.join(STORE_DIR, 'drive_config.json');

let memoryTemplatesCache: Record<string, any> = {};
let memoryDriveConfigCache: any = null;
let memoryHistoryCache: any[] = [];

function readServerTemplates(): Record<string, any> {
  try {
    if (fs.existsSync(CONFIG_FILE)) {
      const raw = fs.readFileSync(CONFIG_FILE, 'utf-8');
      const parsed = JSON.parse(raw);
      memoryTemplatesCache = parsed;
      return parsed;
    }
  } catch (err) {
    console.error('Error reading templates_store/config.json:', err);
  }
  return memoryTemplatesCache;
}

function writeServerTemplates(configData: Record<string, any>): void {
  memoryTemplatesCache = configData || {};
  try {
    if (!fs.existsSync(STORE_DIR)) {
      fs.mkdirSync(STORE_DIR, { recursive: true });
    }
    fs.writeFileSync(CONFIG_FILE, JSON.stringify(configData, null, 2), 'utf-8');
  } catch (err) {
    console.error('Error writing templates_store/config.json:', err);
  }
}

function readServerDriveConfig(): any {
  try {
    if (fs.existsSync(DRIVE_CONFIG_FILE)) {
      const raw = fs.readFileSync(DRIVE_CONFIG_FILE, 'utf-8');
      const parsed = JSON.parse(raw);
      memoryDriveConfigCache = parsed;
      return parsed;
    }
  } catch (err) {
    console.error('Error reading templates_store/drive_config.json:', err);
  }
  return memoryDriveConfigCache;
}

function writeServerDriveConfig(driveConfig: any): void {
  memoryDriveConfigCache = driveConfig;
  try {
    if (!fs.existsSync(STORE_DIR)) {
      fs.mkdirSync(STORE_DIR, { recursive: true });
    }
    fs.writeFileSync(DRIVE_CONFIG_FILE, JSON.stringify(driveConfig, null, 2), 'utf-8');
  } catch (err) {
    console.error('Error writing templates_store/drive_config.json:', err);
  }
}

function readServerHistory(): any[] {
  try {
    if (fs.existsSync(HISTORY_FILE)) {
      const raw = fs.readFileSync(HISTORY_FILE, 'utf-8');
      const parsed = JSON.parse(raw);
      memoryHistoryCache = parsed;
      return parsed;
    }
  } catch (err) {
    console.error('Error reading templates_store/history.json:', err);
  }
  return memoryHistoryCache;
}

function writeServerHistory(historyData: any[]): void {
  const sliced = Array.isArray(historyData) ? historyData.slice(0, 50) : [];
  memoryHistoryCache = sliced;
  try {
    if (!fs.existsSync(STORE_DIR)) {
      fs.mkdirSync(STORE_DIR, { recursive: true });
    }
    fs.writeFileSync(HISTORY_FILE, JSON.stringify(sliced, null, 2), 'utf-8');
  } catch (err) {
    console.error('Error writing templates_store/history.json:', err);
  }
}

/**
 * Downloads a publicly shared JSON file ("Anyone with the link can view") directly from Google Drive
 * so that any visitor accessing the app gets the templates from Google Drive without needing OAuth login.
 */
async function fetchPublicGoogleDriveJson(fileId: string): Promise<Record<string, any> | null> {
  if (!fileId || !/^[a-zA-Z0-9_-]+$/.test(fileId)) {
    return null;
  }

  const urlsToTry = [
    `https://drive.google.com/uc?export=download&id=${encodeURIComponent(fileId)}`,
    `https://drive.usercontent.google.com/download?id=${encodeURIComponent(fileId)}&export=download&confirm=t`,
  ];

  for (const url of urlsToTry) {
    try {
      const response = await fetch(url, {
        headers: {
          Accept: 'application/json, text/plain, */*',
          'User-Agent': 'Mozilla/5.0 (compatible; SmartWordTemplatesApp/1.0)',
        },
        redirect: 'follow',
      });

      if (!response.ok) continue;
      const text = await response.text();
      if (!text || text.trim().startsWith('<!DOCTYPE') || text.trim().startsWith('<html')) {
        continue;
      }

      const parsed = JSON.parse(text);
      if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
        return parsed;
      }
    } catch (err) {
      console.warn(`Could not fetch public Drive file from ${url}:`, err);
    }
  }

  return null;
}

async function startServer() {
  const app = express();
  const PORT = Number(process.env.PORT) || 3000;

  // Increase payload limit so base64 .docx files inside config.json sync smoothly
  app.use(express.json({ limit: '50mb' }));

  // GET shared templates for all users visiting the app link
  app.get('/api/templates', async (req, res) => {
    const refreshFromDrive = req.query.refreshDrive === 'true';
    const driveConfig = readServerDriveConfig();

    if (refreshFromDrive && driveConfig?.backupFileId) {
      const driveTemplates = await fetchPublicGoogleDriveJson(driveConfig.backupFileId);
      if (driveTemplates && Object.keys(driveTemplates).length > 0) {
        writeServerTemplates(driveTemplates);
        res.json({ templates: driveTemplates, driveConfig, source: 'google_drive' });
        return;
      }
    }

    const templates = readServerTemplates();
    res.json({ templates, driveConfig, source: driveConfig ? 'google_drive_cached' : 'server' });
  });

  // POST (replace or merge) shared templates on the server
  app.post('/api/templates', (req, res) => {
    const incoming = req.body?.templates;
    const mode = req.body?.mode || 'replace';
    if (!incoming || typeof incoming !== 'object') {
      res.status(400).json({ error: 'Invalid templates payload' });
      return;
    }

    if (mode === 'merge') {
      const existing = readServerTemplates();
      const merged = { ...existing, ...incoming };
      writeServerTemplates(merged);
      res.json({ templates: merged, driveConfig: readServerDriveConfig() });
      return;
    }

    writeServerTemplates(incoming);
    res.json({ templates: incoming, driveConfig: readServerDriveConfig() });
  });

  // GET Google Drive shared storage configuration & optionally pull latest from Drive
  app.get('/api/drive-config', async (req, res) => {
    const driveConfig = readServerDriveConfig();
    const pullLatest = req.query.pull === 'true';

    if (driveConfig?.backupFileId && pullLatest) {
      const driveTemplates = await fetchPublicGoogleDriveJson(driveConfig.backupFileId);
      if (driveTemplates && Object.keys(driveTemplates).length > 0) {
        writeServerTemplates(driveTemplates);
        const updatedDriveConfig = {
          ...driveConfig,
          templateCount: Object.keys(driveTemplates).length,
          updatedAt: new Date().toISOString(),
        };
        writeServerDriveConfig(updatedDriveConfig);
        res.json({
          driveConfig: updatedDriveConfig,
          templates: driveTemplates,
          syncedFromDrive: true,
        });
        return;
      }
    }

    res.json({
      driveConfig,
      templates: readServerTemplates(),
      syncedFromDrive: false,
    });
  });

  // POST Google Drive shared storage configuration after publishing/sharing on Google Drive
  app.post('/api/drive-config', (req, res) => {
    const driveConfig = req.body?.driveConfig;
    const templates = req.body?.templates;

    if (!driveConfig || !driveConfig.backupFileId) {
      res.status(400).json({ error: 'Missing driveConfig.backupFileId' });
      return;
    }

    writeServerDriveConfig(driveConfig);
    if (templates && typeof templates === 'object') {
      writeServerTemplates(templates);
    }

    res.json({
      driveConfig,
      templates: templates || readServerTemplates(),
    });
  });

  // GET public Google Drive file by fileId (allows linking any shared Drive file or ?driveId=...)
  app.get('/api/drive-public/:fileId', async (req, res) => {
    const fileId = req.params.fileId;
    const driveTemplates = await fetchPublicGoogleDriveJson(fileId);

    if (!driveTemplates || Object.keys(driveTemplates).length === 0) {
      res.status(404).json({
        error:
          'Không thể đọc tệp từ Google Drive. Hãy đảm bảo tệp trên Google Drive đã được bật chia sẻ "Bất kỳ ai có đường liên kết đều có thể xem" (Anyone with the link).',
      });
      return;
    }

    const existingDriveConfig = readServerDriveConfig();
    const nextDriveConfig = {
      backupFileId: fileId,
      backupFileName: existingDriveConfig?.backupFileName || 'templates_store_backup.json',
      backupWebViewLink: `https://drive.google.com/file/d/${fileId}/view?usp=sharing`,
      folderId: existingDriveConfig?.folderId,
      folderName: existingDriveConfig?.folderName,
      folderWebViewLink: existingDriveConfig?.folderWebViewLink,
      ownerEmail: existingDriveConfig?.ownerEmail,
      updatedAt: new Date().toISOString(),
      templateCount: Object.keys(driveTemplates).length,
    };

    writeServerDriveConfig(nextDriveConfig);
    writeServerTemplates(driveTemplates);

    res.json({
      driveConfig: nextDriveConfig,
      templates: driveTemplates,
    });
  });

  // DELETE a specific template from the shared server store
  app.delete('/api/templates/:name', (req, res) => {
    const name = req.params.name;
    const existing = readServerTemplates();
    if (existing[name]) {
      delete existing[name];
      writeServerTemplates(existing);
    }
    res.json({ templates: existing });
  });

  // GET shared export history
  app.get('/api/history', (_req, res) => {
    const history = readServerHistory();
    res.json({ history });
  });

  // POST new export history record
  app.post('/api/history', (req, res) => {
    const record = req.body?.record;
    if (!record) {
      res.status(400).json({ error: 'Invalid history record' });
      return;
    }
    const current = readServerHistory();
    const updated = [record, ...current.filter((r: any) => r.id !== record.id)].slice(0, 50);
    writeServerHistory(updated);
    res.json({ history: updated });
  });

  // DELETE all export history
  app.delete('/api/history', (_req, res) => {
    writeServerHistory([]);
    res.json({ history: [] });
  });

  const distPath = path.join(__dirname, 'dist');
  const distIndexHtml = path.join(distPath, 'index.html');
  const hasBuiltDist = fs.existsSync(distIndexHtml);

  if (!hasBuiltDist) {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true, hmr: false },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(distPath));
    app.get('*', (_req, res) => {
      res.sendFile(distIndexHtml);
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Failed to start server:', err);
  process.exit(1);
});

