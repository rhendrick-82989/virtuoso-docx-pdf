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

    // Use a LibreOffice Basic macro to open the doc, update all fields (incl. TOC),
    // then export to PDF — this ensures TOC page numbers are recalculated.
    const macroScript = `
import uno
from com.sun.star.beans import PropertyValue

def update_fields_and_export():
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    url = uno.systemPathToFileUrl("${inPath}")
    props = []
    doc = desktop.loadComponentFromURL(url, "_blank", 0, props)

    # Update all fields including TOC
    doc.getTextFields().refresh()
    try:
        indexes = doc.getDocumentIndexes()
        for i in range(indexes.Count):
            indexes.getByIndex(i).update()
    except:
        pass

    # Export to PDF
    out_url = uno.systemPathToFileUrl("${outPath}")
    pdf_props = [PropertyValue()]
    pdf_props[0].Name = "FilterName"
    pdf_props[0].Value = "writer_pdf_Export"
    doc.storeToURL(out_url, tuple(pdf_props))
    doc.close(True)

update_fields_and_export()
`.trim();

    const macroPath = path.join(tmpDir, 'convert.py');
    fs.writeFileSync(macroPath, macroScript);

    await new Promise((resolve, reject) => {
      // Try python-based UNO macro first; fall back to direct convert if it fails
      const unoCmd = `python3 "${macroPath}"`;
      const env = { ...process.env, HOME: tmpDir, PYTHONPATH: '/usr/lib/libreoffice/program' };
      exec(unoCmd, { timeout: 60_000, env }, (err) => {
        if (!err && fs.existsSync(outPath)) return resolve();
        // Fallback: direct headless convert (TOC page numbers may all be 1)
        const cmd = `soffice --headless --convert-to pdf --outdir "${tmpDir}" "${inPath}"`;
        exec(cmd, { timeout: 60_000, env: { ...process.env, HOME: tmpDir } }, (err2, stdout, stderr) => {
          if (err2) return reject(new Error(stderr || err2.message));
          resolve();
        });
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
