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

// ── Convert: headless soffice ─────────────────────────────────────────────────
function sofficeConvert(inPath, outDir) {
  return new Promise((resolve, reject) => {
    const cmd = `soffice --headless --convert-to pdf --outdir "${outDir}" "${inPath}"`;
    exec(cmd, { timeout: 90_000, env: { ...process.env, HOME: outDir } }, (err, _, stderr) => {
      if (err) return reject(new Error(stderr || err.message));
      resolve();
    });
  });
}

// ── Extract page numbers for headings from a PDF using pdftotext ─────────────
// Returns { "heading text": pageNumber, ... }
async function extractHeadingPages(pdfPath) {
  const pageMap = {};
  return new Promise((resolve) => {
    exec(`pdftotext -layout "${pdfPath}" -`, { timeout: 30_000, maxBuffer: 10 * 1024 * 1024 },
      (err, stdout) => {
        if (err) { console.warn('pdftotext failed:', err.message); return resolve(pageMap); }
        // pdftotext inserts \f (form feed, \x0c) between pages
        const pages = stdout.split('\x0c');
        pages.forEach((pageText, idx) => {
          const pageNum = idx + 1;
          const lines = pageText.split('\n');
          lines.forEach(line => {
            // Match heading lines: "ID.  Title" pattern
            const m = line.trim().match(/^([A-Z0-9]{1,4}[a-z]?-?\d*)\.\s+(.+)/);
            if (m) {
              const key = m[1] + '. ' + m[2].trim();
              if (!pageMap[key]) pageMap[key] = pageNum;
            }
          });
        });
        resolve(pageMap);
      }
    );
  });
}

// ── Patch docx TOC entries with real page numbers ─────────────────────────────
async function patchTocPageNumbers(docxBuf, pageMap) {
  const zip = await new JSZip().loadAsync(docxBuf);
  let xml = await zip.file('word/document.xml').async('string');

  // Replace each TOC entry placeholder: find w:pStyle TOC1 paragraphs
  // and update the page number run at the end of each entry
  xml = xml.replace(
    /(<w:p>(?:(?!<w:p>).)*?<w:pStyle w:val="TOC1"\/>(?:(?!<\/w:p>).)*?<\/w:p>)/gs,
    (match) => {
      // Extract the title text from the entry
      const titleMatch = match.match(/<w:t[^>]*>([^<]+)<\/w:t>/g);
      if (!titleMatch) return match;
      const titleText = titleMatch.map(t => t.replace(/<[^>]+>/g, '')).join('').trim();
      // Find the page number for this title
      let pageNum = null;
      for (const [heading, pg] of Object.entries(pageMap)) {
        // Match on ID prefix (e.g. "1a.") since titles may be truncated
        const idMatch = titleText.match(/^([A-Z0-9]{1,4}[a-z]?-?\d*)\./);
        const headingId = heading.match(/^([A-Z0-9]{1,4}[a-z]?-?\d*)\./);
        if (idMatch && headingId && idMatch[1] === headingId[1]) {
          pageNum = pg;
          break;
        }
      }
      if (!pageNum) return match;
      // Replace the last <w:t> run (page number placeholder) with actual number
      return match.replace(/<w:r><w:t>(\d+)<\/w:t><\/w:r>(<\/w:p>)$/, 
        `<w:r><w:t>${pageNum}</w:t></w:r>$2`);
    }
  );

  zip.file('word/document.xml', xml);
  return zip.generateAsync({ type: 'nodebuffer' });
}

// ── POST /convert ─────────────────────────────────────────────────────────────
app.post('/convert', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });

  const id      = crypto.randomBytes(8).toString('hex');
  const tmpDir  = path.join(os.tmpdir(), 'docx-pdf-' + id);
  const pass1In = path.join(tmpDir, 'pass1.docx');
  const pass1Pdf = path.join(tmpDir, 'pass1.pdf');
  const pass2In = path.join(tmpDir, 'pass2.docx');
  const pass2Pdf = path.join(tmpDir, 'pass2.pdf');

  try {
    fs.mkdirSync(tmpDir, { recursive: true });
    fs.writeFileSync(pass1In, req.file.buffer);

    // ── Pass 1: convert to PDF to get page layout ─────────────────────────
    await sofficeConvert(pass1In, tmpDir);
    const p1pdf = path.join(tmpDir, 'pass1.pdf');
    if (!fs.existsSync(p1pdf)) throw new Error('Pass 1 conversion failed.');

    // ── Extract heading → page number mapping ────────────────────────────
    const pageMap = await extractHeadingPages(p1pdf);
    console.log('[convert] page map:', JSON.stringify(pageMap));

    // ── Pass 2: patch TOC entries with real page numbers, reconvert ──────
    const patched = await patchTocPageNumbers(req.file.buffer, pageMap);
    fs.writeFileSync(pass2In, patched);
    await sofficeConvert(pass2In, tmpDir);
    const p2pdf = path.join(tmpDir, 'pass2.pdf');
    if (!fs.existsSync(p2pdf)) throw new Error('Pass 2 conversion failed.');

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
