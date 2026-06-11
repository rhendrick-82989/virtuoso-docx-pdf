const express  = require('express');
const multer   = require('multer');
const { exec } = require('child_process');
const fs       = require('fs');
const path     = require('path');
const crypto   = require('crypto');
const os       = require('os');

const app    = express();
const upload = multer({ limits: { fileSize: 50 * 1024 * 1024 } }); // 50 MB cap

// ── CORS: allow requests from your Vercel domain ─────────────────────────────
const ALLOWED_ORIGINS = (process.env.ALLOWED_ORIGINS || '*').split(',').map(s => s.trim());

app.use((req, res, next) => {
  const origin = req.headers.origin || '';
  if (ALLOWED_ORIGINS.includes('*') || ALLOWED_ORIGINS.includes(origin)) {
    res.setHeader('Access-Control-Allow-Origin', origin || '*');
  }
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

// ── Health check ──────────────────────────────────────────────────────────────
app.get('/health', (_req, res) => res.json({ ok: true }));

// ── POST /convert  (multipart: field name "file", .docx) ─────────────────────
app.post('/convert', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });

  const id      = crypto.randomBytes(8).toString('hex');
  const tmpDir  = path.join(os.tmpdir(), 'docx-pdf-' + id);
  const inPath  = path.join(tmpDir, 'input.docx');
  const outPath = path.join(tmpDir, 'input.pdf');

  try {
    fs.mkdirSync(tmpDir, { recursive: true });
    fs.writeFileSync(inPath, req.file.buffer);

    await new Promise((resolve, reject) => {
      const cmd = `soffice --headless --convert-to pdf --outdir "${tmpDir}" "${inPath}"`;
      exec(cmd, { timeout: 60_000, env: { ...process.env, HOME: tmpDir } }, (err, stdout, stderr) => {
        if (err) return reject(new Error(stderr || err.message));
        resolve();
      });
    });

    if (!fs.existsSync(outPath)) throw new Error('LibreOffice produced no output.');

    const pdf = fs.readFileSync(outPath);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="response.pdf"');
    res.send(pdf);

  } catch (err) {
    console.error('[convert]', err.message);
    res.status(500).json({ error: err.message });
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`docx-pdf-service listening on :${PORT}`));
