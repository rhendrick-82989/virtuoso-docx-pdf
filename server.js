const express   = require('express');
const multer    = require('multer');
const { exec }  = require('child_process');
const fs        = require('fs');
const path      = require('path');
const crypto    = require('crypto');
const os        = require('os');
const JSZip     = require('jszip');

const app    = express();
const upload = multer({ limits: { fileSize: 50 * 1024 * 1024 } });

// ── CORS ──────────────────────────────────────────────────────────────────────
const ALLOWED_ORIGINS = (process.env.ALLOWED_ORIGINS || '*').split(',').map(s => s.trim());
app.use((req, res, next) => {
  const origin = req.headers.origin || '';
  if (ALLOWED_ORIGINS.includes('*') || ALLOWED_ORIGINS.includes(origin))
    res.setHeader('Access-Control-Allow-Origin', origin || '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

app.get('/health', (_req, res) => res.json({ ok: true }));

// ── Headless soffice convert ──────────────────────────────────────────────────
function sofficeConvert(inPath, outDir) {
  return new Promise((resolve, reject) => {
    const cmd = `soffice --headless --convert-to pdf --outdir "${outDir}" "${inPath}"`;
    exec(cmd, { timeout: 90_000, env: { ...process.env, HOME: outDir } }, (err, _, stderr) => {
      if (err) return reject(new Error(stderr || err.message));
      resolve();
    });
  });
}

// ── Extract heading page numbers via invisible @@HDR_ID@@ anchors ────────────
// The docx embeds white 1pt text anchors in each heading paragraph.
// pdftotext reliably extracts these regardless of where the heading falls on the page.
function extractHeadingPages(pdfPath) {
  return new Promise((resolve) => {
    exec(`pdftotext -layout "${pdfPath}" -`, { timeout: 30_000, maxBuffer: 20 * 1024 * 1024 },
      (err, stdout) => {
        if (err) { console.warn('[pages] pdftotext failed:', err.message); return resolve({}); }
        const pageMap = {};
        const pages = stdout.split('\x0c');
        pages.forEach((pageText, idx) => {
          const pg = idx + 1;
          const matches = pageText.matchAll(/@@HDR_(\d{1,2}[a-z]?|[A-Z]{1,2}-\d{1,2})@@/g);
          for (const m of matches) {
            if (!pageMap[m[1]]) {
              pageMap[m[1]] = pg;
              console.log(`[pages] ${m[1]} -> p${pg}`);
            }
          }
        });
        console.log(`[pages] found ${Object.keys(pageMap).length} headings`);
        resolve(pageMap);
      }
    );
  });
}

// ── Patch TOC entries + strip anchors from docx XML ──────────────────────────
async function patchDocx(docxBuf, pageMap) {
  const zip = await new JSZip().loadAsync(docxBuf);
  let xml = await zip.file('word/document.xml').async('string');

  // 1. Patch TOC page numbers: find last <w:t>\d+</w:t> in each TOC1 paragraph
  //    The ID is embedded earlier in the same paragraph as "<w:t>ID.</w:t>"
  let patchCount = 0;
  xml = xml.replace(
    /(<w:pStyle w:val="TOC1"\/>(?:(?!<\/w:p>)[\s\S])*?)(<w:t[^>]*>)(\d+)(<\/w:t><\/w:r><\/w:p>)/g,
    (match, prefix, tOpen, currentNum, tail) => {
      const idMatch = prefix.match(/<w:t[^>]*>(\d{1,2}[a-z]?|[A-Z]{1,2}-\d{1,2})\.<\/w:t>/);
      if (!idMatch) { console.warn('[patch] no ID found in a TOC paragraph'); return match; }
      const realPage = pageMap[idMatch[1]];
      if (!realPage) { console.warn('[patch] no page for', idMatch[1]); return match; }
      patchCount++;
      console.log(`[patch] TOC ${idMatch[1]}: ${currentNum} -> ${realPage}`);
      return prefix + tOpen + realPage + tail;
    }
  );
  console.log(`[patch] patched ${patchCount} TOC entries`);

  // 2. Strip the invisible anchor runs so they don't appear in final PDF
  xml = xml.replace(/@@HDR_(?:\d{1,2}[a-z]?|[A-Z]{1,2}-\d{1,2})@@/g, '');

  zip.file('word/document.xml', xml);
  return zip.generateAsync({ type: 'nodebuffer' });
}

// ── POST /convert ─────────────────────────────────────────────────────────────
app.post('/convert', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });

  const id      = crypto.randomBytes(8).toString('hex');
  const tmpDir  = path.join(os.tmpdir(), 'docx-pdf-' + id);
  const pass1In = path.join(tmpDir, 'pass1.docx');
  const pass2In = path.join(tmpDir, 'pass2.docx');

  try {
    fs.mkdirSync(tmpDir, { recursive: true });
    const originalBuf = req.file.buffer;
    fs.writeFileSync(pass1In, originalBuf);

    // Pass 1: convert to establish page layout
    console.log('[convert] pass 1...');
    await sofficeConvert(pass1In, tmpDir);
    const p1pdf = path.join(tmpDir, 'pass1.pdf');
    if (!fs.existsSync(p1pdf)) throw new Error('Pass 1 produced no PDF.');

    // Extract heading page numbers from anchors
    const pageMap = await extractHeadingPages(p1pdf);

    // Pass 2: patch TOC + strip anchors, reconvert
    console.log('[convert] pass 2...');
    const patched = await patchDocx(originalBuf, pageMap);
    fs.writeFileSync(pass2In, patched);
    await sofficeConvert(pass2In, tmpDir);
    const p2pdf = path.join(tmpDir, 'pass2.pdf');
    if (!fs.existsSync(p2pdf)) throw new Error('Pass 2 produced no PDF.');

    const pdf = fs.readFileSync(p2pdf);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="DEN250054_Complete_Response.pdf"');
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
