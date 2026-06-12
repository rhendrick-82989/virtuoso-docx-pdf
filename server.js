const express   = require('express');
const multer    = require('multer');
const { exec, execSync }  = require('child_process');
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

// ── Extract heading → page map using pdftotext page-by-page ──────────────────
// Returns { "1a": 3, "2a": 5, "M-1": 28, ... }
async function extractHeadingPages(pdfPath) {
  const pageMap = {};
  return new Promise((resolve) => {
    // Get total page count first
    exec(`pdfinfo "${pdfPath}"`, { timeout: 10_000 }, (err, stdout) => {
      if (err) { console.warn('pdfinfo failed:', err.message); return resolve(pageMap); }
      const m = stdout.match(/Pages:\s+(\d+)/);
      const totalPages = m ? parseInt(m[1]) : 0;
      if (!totalPages) return resolve(pageMap);

      let completed = 0;
      for (let pg = 1; pg <= totalPages; pg++) {
        const pgNum = pg;
        exec(`pdftotext -f ${pgNum} -l ${pgNum} -layout "${pdfPath}" -`, 
          { timeout: 10_000 },
          (err2, text) => {
            if (!err2 && text) {
              text.split('\n').forEach(line => {
                // Match "1a.   Title" or "M-1.  Title" or "AC-1.  Title"
                const hm = line.trim().match(/^([A-Z]{0,2}\d{1,2}[a-z]?(?:-\d+)?)\.\s{2,}(.+)/);
                if (hm && !pageMap[hm[1]]) {
                  pageMap[hm[1]] = pgNum;
                  console.log(`[pages] ${hm[1]} -> p${pgNum}`);
                }
              });
            }
            completed++;
            if (completed === totalPages) resolve(pageMap);
          }
        );
      }
    });
  });
}

// ── Patch TOC entries in docx XML with real page numbers ─────────────────────
async function patchTocPageNumbers(docxBuf, pageMap) {
  if (!Object.keys(pageMap).length) {
    console.warn('[patch] empty page map — skipping');
    return docxBuf;
  }

  const zip = await new JSZip().loadAsync(docxBuf);
  let xml = await zip.file('word/document.xml').async('string');

  // Each TOC entry paragraph has structure:
  //   <w:t>ID.</w:t> ... <w:t>\tTitle\t</w:t> ... <w:t>1</w:t></w:p>
  // Strategy: find the last <w:t> in each TOC1-styled paragraph and replace with real page num
  xml = xml.replace(
    /(<w:pStyle w:val="TOC1"\/>(?:(?!<\/w:p>)[\s\S])*?)(<w:r[^>]*><w:t[^>]*>)(\d+)(<\/w:t><\/w:r>)(<\/w:p>)/g,
    (match, prefix, runOpen, currentNum, runClose, paraClose) => {
      // Extract the ID from earlier in the paragraph — find first <w:t> content
      const idMatch = prefix.match(/<w:t[^>]*>([A-Z]{0,2}\d{1,2}[a-z]?(?:-\d+)?)\.<\/w:t>/);
      if (!idMatch) return match;
      const id = idMatch[1];
      const realPage = pageMap[id];
      if (!realPage) { console.warn('[patch] no page for', id); return match; }
      console.log(`[patch] ${id}: ${currentNum} -> ${realPage}`);
      return prefix + runOpen + realPage + runClose + paraClose;
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
  const pass2In = path.join(tmpDir, 'pass2.docx');

  try {
    fs.mkdirSync(tmpDir, { recursive: true });
    const originalBuf = req.file.buffer;
    fs.writeFileSync(pass1In, originalBuf);

    // ── Pass 1: convert to PDF to establish real page layout ─────────────
    console.log('[convert] pass 1...');
    await sofficeConvert(pass1In, tmpDir);
    const p1pdf = path.join(tmpDir, 'pass1.pdf');
    if (!fs.existsSync(p1pdf)) throw new Error('Pass 1 conversion produced no PDF.');

    // ── Extract heading page numbers ──────────────────────────────────────
    console.log('[convert] extracting page numbers...');
    const pageMap = await extractHeadingPages(p1pdf);
    console.log('[convert] found', Object.keys(pageMap).length, 'headings');

    // ── Pass 2: patch TOC + reconvert ─────────────────────────────────────
    console.log('[convert] pass 2...');
    const patched = await patchTocPageNumbers(originalBuf, pageMap);
    fs.writeFileSync(pass2In, patched);
    await sofficeConvert(pass2In, tmpDir);
    const p2pdf = path.join(tmpDir, 'pass2.pdf');
    if (!fs.existsSync(p2pdf)) throw new Error('Pass 2 conversion produced no PDF.');

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
