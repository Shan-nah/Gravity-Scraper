require('dotenv').config();
const express  = require('express');
const axios    = require('axios');
const cheerio  = require('cheerio');
const ExcelJS  = require('exceljs');
const cors     = require('cors');
const path     = require('path');
const fs       = require('fs');
const crypto   = require('crypto');
const { Readable } = require('stream');

const app  = express();
const PORT = process.env.PORT || 8080;

process.on('uncaughtException',  err => console.error('Uncaught Exception:', err));
process.on('unhandledRejection', err => console.error('Unhandled Rejection:', err));

// ── Gemini API setup (optional — enables EMD Exemption extraction)
//    Set GEMINI_API_KEY in .env or add keys to keys.txt (one per line)
let geminiApiKeys = [];
(function initGemini() {
  const envKey = process.env.GEMINI_API_KEY;
  if (envKey) {
    geminiApiKeys = envKey.split(',').map(k => k.trim()).filter(Boolean);
    console.log(`✓  Gemini API enabled (${geminiApiKeys.length} key(s)) — EMD Exemption will be extracted\n`);
    return;
  }
  const keyFile = path.join(__dirname, 'keys.txt');
  if (fs.existsSync(keyFile)) {
    const lines = fs.readFileSync(keyFile, 'utf8').split('\n').map(l => l.trim()).filter(Boolean);
    if (lines.length > 0) {
      geminiApiKeys = lines;
      console.log(`✓  Gemini API enabled via keys.txt (${geminiApiKeys.length} key(s)) — EMD Exemption will be extracted\n`);
    }
  }
})();


const GEMINI_MODELS = [
  'gemini-2.5-flash',
  'gemini-2.5-flash-lite',
  'gemini-3.1-flash-lite',
  'gemma-4-26b-a4b-it',
  'gemma-4-31b-it',
];


// Strip chain-of-thought reasoning that verbose models (Gemma, etc.) prepend to their answers.
// The actual answer is always the last clean non-markdown line.
function cleanGeminiOutput(text) {
  if (!text) return text;

  // Detect reasoning markers used by verbose models
  const hasReasoning = /\*\s+(?:Task|Constraint|Source|Searching|Found|Check|Draft|Require):/i.test(text) ||
                       /^\s*\|\s/m.test(text) && text.includes('Task:');
  if (hasReasoning) {
    const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
    for (let i = lines.length - 1; i >= 0; i--) {
      const line = lines[i];
      if (/^[\*\-\|#]/.test(line) || line.length < 15) continue;
      // Deduplicate if model repeated the answer twice on one line
      const mid = Math.ceil(line.length / 2);
      const a = line.slice(0, mid).trim().toLowerCase();
      const b = line.slice(line.length - mid).trim().toLowerCase();
      return (a === b ? line.slice(0, mid) : line).trim();
    }
  }

  // Deduplicate answer-doubled-without-newline (also common in reasoning models)
  const t = text.trim();
  const mid = Math.ceil(t.length / 2);
  const a = t.slice(0, mid).trim().toLowerCase().replace(/\s+/g, ' ');
  const b = t.slice(t.length - mid).trim().toLowerCase().replace(/\s+/g, ' ');
  if (a.length > 20 && a === b) return t.slice(0, mid).trim();

  return t;
}

// In-memory cache: identical snippets (e.g. repeated GeM clause) → instant result, no API call
const _geminiCache    = new Map();
// Per-key quota tracker: skip calls when all keys were recently 429'd
const _keyLastQuota   = new Map();

async function callGemini(text, prompt, extraKey) {
  const keys = [...new Set([extraKey, ...geminiApiKeys].filter(Boolean))];
  if (!keys.length) return null;

  // Cache check — first 500 chars uniquely identify repeated clauses (GeM standard text, etc.)
  const cacheKey = text.slice(0, 500);
  if (_geminiCache.has(cacheKey)) {
    console.log('  Gemini cache hit');
    return _geminiCache.get(cacheKey);
  }

  // Quota guard — if ALL keys were 429'd in the last 60s, skip immediately
  const allExhausted = keys.every(k => {
    const t = _keyLastQuota.get(k);
    return t && Date.now() - t < 60000;
  });
  if (allExhausted) {
    console.warn('  All Gemini keys recently quota-limited, skipping');
    return null;
  }

  const { GoogleGenerativeAI } = require('@google/generative-ai');

  for (const modelName of GEMINI_MODELS) {
    for (const key of keys) {
      try {
        const genAI = new GoogleGenerativeAI(key);
        const model = genAI.getGenerativeModel({ model: modelName });

        // Disable thinking mode (thinkingBudget:0 cuts response time from 30–120s → 1–3s)
        const callPromise = model.generateContent({
          contents: [{ role: 'user', parts: [{ text: prompt + '\n\n' + text }] }],
          generationConfig: { thinkingConfig: { thinkingBudget: 0 } },
        });
        // 10s safety timeout — if model hangs, try next key/model immediately
        const result = await Promise.race([
          callPromise,
          new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), 10000)),
        ]);

        const out = cleanGeminiOutput(result.response.text());
        if (out) {
          _geminiCache.set(cacheKey, out); // cache for subsequent identical snippets
          return out;
        }
      } catch (e) {
        const msg = e.message || '';
        if (/429|quota|rate.?limit|resource.?exhaust/i.test(msg)) {
          _keyLastQuota.set(key, Date.now());
          console.warn(`  Key …${key.slice(-4)} rate-limited on ${modelName}, trying next key`);
        } else {
          console.warn(`  Gemini ${modelName} failed: ${msg.slice(0, 80)}`);
          break; // non-quota error: skip to next model
        }
      }
    }
  }
  return null;
}

async function fetchDocText(url) {
  try {
    const resp = await axios.get(url, {
      headers: { ...HEADERS, Accept: 'application/pdf,*/*' },
      responseType: 'arraybuffer',
      timeout: 12000,
      maxContentLength: 10 * 1024 * 1024,
    });
    const ct = (resp.headers['content-type'] || '').toLowerCase();
    const buf = Buffer.from(resp.data);

    if (ct.includes('pdf') || /\.pdf/i.test(url)) {
      const tmpPath = path.join(TMP_DIR, `pdf_${Date.now()}_${crypto.randomBytes(4).toString('hex')}.pdf`);
      await fs.promises.writeFile(tmpPath, buf);
      try {
        if (PDFTOTEXT) {
          // Fast path: native binary, 5–40× faster than JS parsing
          const { execFile } = require('child_process');
          const { promisify } = require('util');
          const { stdout } = await promisify(execFile)(
            PDFTOTEXT,
            ['-q', '-enc', 'UTF-8', '-f', '1', '-l', '15', tmpPath, '-'],
            { maxBuffer: 50 * 1024 * 1024, timeout: 15000 }
          );
          return stdout.replace(/\s+/g, ' ').trim();
        }
        // Fallback: PDFParse capped at 30 pages
        const { PDFParse } = require('pdf-parse');
        const parser = new PDFParse({ url: `file://${tmpPath}` });
        const result = await parser.getText({ first: 30 });
        return (result.text || '').replace(/\s+/g, ' ').trim();
      } finally {
        fs.unlink(tmpPath, () => {});
      }
    }
    return buf.toString('utf8').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  } catch (e) {
    return '';
  }
}

// Scan extracted PDF text for EMD exemption-related passages.
// Returns a short snippet (≤3000 chars) of the relevant text, or null if none found.
function extractEmdSnippet(text) {
  if (!text) return null;
  const HALF_WIN = 500;
  const lower = text.toLowerCase();
  const positions = [];

  // Strong patterns: directly mention exemption + EMD only
  const strongPats = [
    /emd\s*exemption/g,
    /exemption\s+(?:from|of)\s+(?:emd|earnest)/g,
    /exempt(?:ed|ion)?\s+from\s+(?:paying\s+)?(?:emd|earnest)/g,
    /emd\s+(?:is\s+)?(?:not\s+required|waived|nil|exempted|exempt)/g,
    /no\s+emd\s+(?:is\s+)?(?:required|applicable)/g,
    /waiver\s+of\s+(?:emd|earnest)/g,
    /emd\s*[:\-]\s*nil/g,
    /nil\s+emd\b/g,
    /earnest\s+money\s+deposit\s+(?:is\s+)?(?:not\s+applicable|waived|nil|exempt)/g,
    // "EMD required: No" / "EMD Detail → Required: No" — EMD not required at all
    // The text after pdftotext looks like "EMD Detail ... Required No" with Hindi chars between
    /emd\s+required\s*[:\-]?\s*no\b/g,
    /required\s*[:\-]\s*no\b[^.]*emd/g,
    /emd[^.]{0,120}required[^.]{0,30}no\b/g,   // "EMD Detail ... Required No" broader match
    /required[^.]{0,30}no\b[^.]{0,120}emd/g,   // "Required No ... EMD" reverse order
    // "registered under MSME/NSIC → EMD exemption certificate"
    /emd\s+exemption\s+certificate/g,
    /registered\s+under\s+(?:msme|nsic|msmed)[^.]*emd/g,
    /(?:msme|nsic|msmed)\s+registered[^.]*emd/g,
  ];
  for (const pat of strongPats) {
    let m;
    while ((m = pat.exec(lower)) !== null) positions.push(m.index);
  }

  // Contextual: category keywords near an EMD mention (within 300 chars)
  const emdPos = [];
  const emdPat = /(?:emd|earnest\s+money)/g;
  let em;
  while ((em = emdPat.exec(lower)) !== null) emdPos.push(em.index);

  const catPat = /(?:msme|micro\s+(?:and\s+small|enterprise)|mse\b|sc\/st|scheduled\s+(?:caste|tribe)|startup|dpiit|nsic|udyam)/g;
  let cm;
  while ((cm = catPat.exec(lower)) !== null) {
    for (const ep of emdPos) {
      if (Math.abs(cm.index - ep) <= 300) { positions.push(Math.min(cm.index, ep)); break; }
    }
  }

  if (positions.length === 0) return null;

  // Build ±HALF_WIN windows, sort, merge overlapping ones
  const ranges = positions
    .map(p => [Math.max(0, p - HALF_WIN), Math.min(text.length, p + HALF_WIN)])
    .sort((a, b) => a[0] - b[0]);

  const merged = [ranges[0]];
  for (let i = 1; i < ranges.length; i++) {
    const last = merged[merged.length - 1];
    if (ranges[i][0] <= last[1] + 100) { last[1] = Math.max(last[1], ranges[i][1]); }
    else merged.push(ranges[i]);
  }

  return merged.map(([s, e]) => text.slice(s, e).trim()).join('\n---\n').slice(0, 3000);
}

// Returns only the NIT document link (UUID match from tender URL) — or falls back to all links.
function pickTargetDocLinks(viewLink, docLinks) {
  const uuidMatch = viewLink.match(
    /TenderNotice\/\d+\/([A-F0-9]{8}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{12})/i
  );
  if (uuidMatch) {
    const uuid = uuidMatch[1].toLowerCase();
    const matched = docLinks.filter(l => l.toLowerCase().includes(uuid));
    if (matched.length > 0) return matched; // UUID found → only this doc
  }
  return docLinks.slice(0, 3); // fallback: up to 3 docs
}

// True when the scraped value is a placeholder meaning "look in the document".
function needsDocLookup(val) {
  if (!val || val === 'N/A' || typeof val === 'number') return false;
  return /refer\s*doc|as\s*per\s*doc|as\s*per\s*boq|as\s*per\s*nit|per\s*nit|see\s*doc|tbd|as\s*per\s*schedule|as\s*per\s*estimate|as\s*per\s*tender|as\s*per\s*drawing/i.test(String(val));
}

// Pure-regex extraction of EMD amount and Tender Value from pre-fetched document texts.
function extractAmountsFromTexts(texts) {
  const result = { emd: null, tv: null };
  const emdPats = [
    /emd\s+(?:amount|value)\s*[:\|=]\s*(?:rs\.?|inr|₹)?\s*([0-9,\.]+(?:\s*(?:crore[s]?|cr|lac[s]?|lakh[s]?|thousand|k))?)/i,
    /earnest\s+money\s+(?:deposit\s+)?(?:amount\s*)?[:\|=]\s*(?:rs\.?|inr|₹)?\s*([0-9,\.]+(?:\s*(?:crore[s]?|cr|lac[s]?|lakh[s]?|thousand|k))?)/i,
  ];
  const tvPats = [
    /estimated\s*(?:bid\s*)?(?:value|cost|amount)\s*[:\|=]\s*(?:rs\.?|inr|₹)?\s*([0-9,\.]+(?:\s*(?:crore[s]?|cr|lac[s]?|lakh[s]?|thousand|k))?)/i,
    /tender\s*value\s*[:\|=]\s*(?:rs\.?|inr|₹)?\s*([0-9,\.]+(?:\s*(?:crore[s]?|cr|lac[s]?|lakh[s]?|thousand|k))?)/i,
    /(?:amount\s*put\s*to\s*tender|contract\s*value|work\s*(?:value|amount|estimate)|nit\s*(?:value|amount))\s*[:\|=]\s*(?:rs\.?|inr|₹)?\s*([0-9,\.]+(?:\s*(?:crore[s]?|cr|lac[s]?|lakh[s]?|thousand|k))?)/i,
    /bid\s*amount\s*[:\|=]\s*(?:rs\.?|inr|₹)?\s*([0-9,\.]+(?:\s*(?:crore[s]?|cr|lac[s]?|lakh[s]?|thousand|k))?)/i,
  ];
  for (const text of texts) {
    if (!text) continue;
    if (!result.emd) {
      for (const p of emdPats) {
        const m = text.match(p);
        if (m?.[1]) { const n = normalizeAmount(m[1].trim()); if (typeof n === 'number') { result.emd = n; break; } }
      }
    }
    if (!result.tv) {
      for (const p of tvPats) {
        const m = text.match(p);
        if (m?.[1]) { const n = normalizeAmount(m[1].trim()); if (typeof n === 'number') { result.tv = n; break; } }
      }
    }
    if (result.emd && result.tv) break;
  }
  return result;
}

// Takes pre-fetched document texts (not URLs) — call after fetching once in scrapeTenderDetail.
async function extractEmdExemptionFromTexts(docTexts, apiKeyOverride) {
  let snippet = null;
  for (const text of docTexts) {
    if (!text) continue;
    snippet = extractEmdSnippet(text);
    if (snippet) { console.log(`  EMD snippet found (${snippet.length} chars)`); break; }
  }

  if (!snippet) { console.log('  No EMD snippet in docs'); return 'No exemption mentioned'; }

  if (!geminiApiKeys.length && !apiKeyOverride) return snippet.trim();

  const prompt = `From this tender document excerpt, state exactly who is exempt from paying EMD (Earnest Money Deposit). Use the exact wording from the document. 1-2 sentences only. If no EMD exemption is mentioned, respond: No exemption mentioned

Excerpt:`;

  const result = await callGemini(snippet, prompt, apiKeyOverride);
  if (result) return result;

  // Direct "EMD not required" signal — no Gemini interpretation needed
  if (/emd[^.]{0,120}required[^.]{0,30}no\b/i.test(snippet) ||
      /required[^.]{0,30}no\b[^.]{0,120}emd/i.test(snippet)) {
    return 'EMD not required';
  }

  const fallback =
    snippet.match(/Under\s+(?:the\s+)?MSE\s+category[^.]+\.[^.]*Traders[^.]+\./i)?.[0] ||
    snippet.match(/Under\s+(?:the\s+)?MSE\s+category[^.]+\./i)?.[0] ||
    snippet.match(/EMD\s+EXEMPTION\s*:\s*(?:The\s+bidder[^.]+\.\s*)?([^/\n]{20,300})/i)?.[1];
  return fallback ? fallback.trim() : 'N/A';
}

// ── Google Drive / Sheets upload (optional)
//    Set GOOGLE_SERVICE_ACCOUNT_KEY in .env (JSON string or file path)
//    Optionally set GOOGLE_SHEETS_SHARED_WITH=your@email.com for editor access
let driveClient = null;
(function initGoogleDrive() {
  const raw = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
  if (!raw) return;
  try {
    const { google } = require('googleapis');
    const creds = raw.trim().startsWith('{') ? JSON.parse(raw) : require(path.resolve(raw));
    const auth  = new google.auth.GoogleAuth({
      credentials: creds,
      scopes: ['https://www.googleapis.com/auth/drive'],
    });
    driveClient = google.drive({ version: 'v3', auth });
    console.log('✓  Google Drive enabled — scrapes will also create a Google Sheet\n');
  } catch (e) {
    console.warn('Google Drive init failed:', e.message);
  }
})();

async function uploadToGoogleSheets(xlsxBuffer, title) {
  const scriptUrl = process.env.GOOGLE_SCRIPT_URL;
  if (!scriptUrl) {
    if (!driveClient) return null;
    // Fallback to direct drive upload (will likely fail on personal accounts due to 0 quota)
    try {
      const res = await driveClient.files.create({
        requestBody: { name: title || 'Tender Data', mimeType: 'application/vnd.google-apps.spreadsheet' },
        media: { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', body: Readable.from(xlsxBuffer) },
        fields: 'id',
      });
      const fileId = res.data.id;
      await driveClient.permissions.create({ fileId, requestBody: { role: 'writer', type: 'anyone' } });
      return `https://docs.google.com/spreadsheets/d/${fileId}/edit`;
    } catch (e) {
      console.warn('Direct Google Sheets upload failed:', e.message);
      return null;
    }
  }

  try {
    const response = await axios.post(scriptUrl, {
      base64: xlsxBuffer.toString('base64'),
      filename: (title || 'Tender Data') + '.xlsx'
    }, { timeout: 30000 });

    if (response.data && response.data.url) {
      console.log('✓  Google Sheet created via Bridge:', response.data.url);
      return response.data.url;
    }
    if (response.data && response.data.error) {
      console.warn('Bridge reported error:', response.data.error);
    }
    return null;
  } catch (e) {
    console.warn('Google Script Bridge failed:', e.message);
    return null;
  }
}

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (_req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

const TMP_DIR = path.join(__dirname, 'tmp');
if (!fs.existsSync(TMP_DIR)) fs.mkdirSync(TMP_DIR);
cleanOldFiles();

// Detect pdftotext (poppler) — 5–40× faster than JS-based PDF parsing
let PDFTOTEXT = null;
(function detectPdftotext() {
  try {
    const { execSync } = require('child_process');
    const p = execSync('which pdftotext 2>/dev/null', { timeout: 2000 }).toString().trim();
    if (p) { PDFTOTEXT = p; console.log(`✓  pdftotext found at ${p} — fast PDF extraction enabled\n`); }
  } catch (_) {}
})();

const HEADERS = {
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
  'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
  'Accept-Language': 'en-US,en;q=0.5',
};
const CONCURRENCY = Math.max(6, geminiApiKeys.length * 3);

function clean(t) { return (t || '').replace(/\s+/g, ' ').trim(); }


// Convert EMD / Tender Value strings to plain rupee integers (JS number type).
// Returning a number — not a string — is what makes Excel sort numerically.
// Handles: "2.5 Crores", "50 Lacs", "₹ 1,23,456", "50000/-", "Nil", etc.
function normalizeAmount(val) {
  if (!val) return 'N/A';
  const raw = String(val).trim();
  if (/^(nil|n\/a|not\s*applicable|na|-|exempt)$/i.test(raw)) return 'N/A';
  const lo = raw.toLowerCase();
  // "X crore Y lac" combo
  const combo = lo.match(/([0-9.]+)\s*crore[s]?\s*(?:and\s*)?([0-9.]+)\s*(?:lac|lakh)/);
  if (combo) return Math.round(+combo[1] * 1e7 + +combo[2] * 1e5);
  // crore / cr
  const cr = lo.match(/([0-9.]+)\s*(?:crore[s]?|cr\.?)\b/);
  if (cr) return Math.round(+cr[1] * 1e7);
  // lac / lakh
  const lac = lo.match(/([0-9.]+)\s*(?:lac[s]?|lakh[s]?)\b/);
  if (lac) return Math.round(+lac[1] * 1e5);
  // thousand / k
  const k = lo.match(/([0-9.]+)\s*(?:thousand|k)\b/);
  if (k) return Math.round(+k[1] * 1e3);
  // plain number (strip currency symbols, commas, slashes)
  const num = parseFloat(raw.replace(/[₹$,\s\/-]/g, '').replace(/[^0-9.]/g, ''));
  if (!isNaN(num) && num > 0) return Math.round(num);
  return val;
}


// Convert 1-based column number → Excel letter (1→A, 26→Z, 27→AA …)
function colNumToLetter(n) {
  let r = '';
  while (n > 0) { n--; r = String.fromCharCode(65 + n % 26) + r; n = Math.floor(n / 26); }
  return r;
}

// Block private/loopback addresses to prevent SSRF
function validateUrl(rawUrl) {
  try {
    const u = new URL(rawUrl);
    if (!['http:', 'https:'].includes(u.protocol)) return false;
    const h = u.hostname.toLowerCase();
    if (/^(localhost|127\.|10\.|172\.(1[6-9]|2\d|3[01])\.|192\.168\.|::1|0\.0\.0\.0)/.test(h)) return false;
    return true;
  } catch { return false; }
}

// ══════════════════════════════════════════════════════════════
//  EXCEL SCHEMA
//  First two columns are user-input dropdowns (Company / Status).
//  18 fixed scraped fields follow, then two catch-all text cols.
// ══════════════════════════════════════════════════════════════
const COLS = [
  // ── User-input columns (blank on scrape; filled by user in Excel)
  { key: 'Company', label: 'Company', width: 16 },
  { key: 'Important', label: 'Important', width: 12 },
  { key: 'Filled Date', label: 'Filled Date', width: 16 },
  { key: 'Filled By', label: 'Filled By', width: 18 },
  { key: 'Bid Status', label: 'Bid Status', width: 14 },
  // ── Fixed scraped fields
  { key: 'TDR', label: 'TDR', width: 14 },
  { key: 'Tender No', label: 'Tender No', width: 26 },
  { key: 'Tendering Authority', label: 'Tendering Authority', width: 36 },
  { key: 'Tender Brief', label: 'Tender Brief', width: 62 },
  { key: 'City', label: 'City', width: 16 },
  { key: 'State', label: 'State', width: 18 },
  { key: 'Document Fees', label: 'Document Fees', width: 16 },
  { key: 'EMD', label: 'EMD (₹)', width: 18 },
  { key: 'EMD Exempt?', label: 'EMD Exempt?', width: 13 },
  { key: 'EMD Exemption', label: 'EMD Exemption', width: 44 },
  { key: 'Tender Value', label: 'Tender Value (₹)', width: 18 },
  { key: 'Tender Type', label: 'Tender Type', width: 16 },
  { key: 'Bidding Type', label: 'Bidding Type', width: 16 },
  { key: 'Competition Type', label: 'Competition Type', width: 18 },
  { key: 'Publish Date', label: 'Publish Date', width: 14 },
  { key: 'Last Date of Bid Submission', label: 'Last Date of Bid Submission', width: 26 },
  { key: 'Tender Opening Date', label: 'Tender Opening Date', width: 22 },
  { key: 'Address', label: 'Address', width: 34 },
  { key: 'Information Source', label: 'Information Source', width: 30 },
  { key: 'View Link', label: 'View Link', width: 60 },
  // ── Variable / catch-all
  { key: 'Additional Details', label: 'Additional Details', width: 48 },
  { key: 'Bid Document Details', label: 'Bid Document Details', width: 70 },
];

// Columns whose values may contain \n — used for row-height calculation
const TEXT_WRAP_COLS = new Set(['Additional Details', 'Bid Document Details', 'EMD Exemption']);


// ══════════════════════════════════════════════════════════════
//  STYLE HELPERS
// ══════════════════════════════════════════════════════════════
function headerStyle(row, colorHex, colCount) {
  row.height = 26;
  for (let c = 1; c <= colCount; c++) {
    const cell = row.getCell(c);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + colorHex } };
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: false };
    cell.border = {
      top: { style: 'medium', color: { argb: 'FF' + colorHex } },
      bottom: { style: 'medium', color: { argb: 'FF' + colorHex } },
      left: { style: 'thin', color: { argb: 'FFB0C4DE' } },
      right: { style: 'thin', color: { argb: 'FFB0C4DE' } },
    };
  }
}

function dataRowStyle(row, rowIdx, colCount) {
  const bg = rowIdx % 2 === 0 ? 'FFEAF4FB' : 'FFFFFFFF';
  for (let c = 1; c <= colCount; c++) {
    const cell = row.getCell(c);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    cell.font = { name: 'Calibri', size: 10 };
    cell.alignment = { vertical: 'top', wrapText: false };
    cell.border = {
      top: { style: 'hair', color: { argb: 'FFCCE0F5' } },
      bottom: { style: 'hair', color: { argb: 'FFCCE0F5' } },
      left: { style: 'thin', color: { argb: 'FFCCE0F5' } },
      right: { style: 'thin', color: { argb: 'FFCCE0F5' } },
    };
  }
}



// ══════════════════════════════════════════════════════════════
//  DATA SHEET — coloured headers, zebra rows, frozen row, filter
// ══════════════════════════════════════════════════════════════
function fillDataSheet(ws, cols, rows, tabArgb, headerColor) {
  ws.properties = { tabColor: { argb: tabArgb } };
  ws.views = [{ state: 'frozen', ySplit: 1, showGridLines: true }];
  ws.columns = cols.map(c => ({ header: c.label, key: c.key, width: c.width }));

  const hdr = ws.getRow(1);
  hdr.values = cols.map(c => c.label);
  headerStyle(hdr, headerColor, cols.length);

  // Make the two text-heavy column headers italic to signal they're free-text
  cols.forEach((c, i) => {
    if (TEXT_WRAP_COLS.has(c.key)) {
      hdr.getCell(i + 1).font = { name: 'Calibri', size: 11, bold: true, italic: true, color: { argb: 'FFFFFFFF' } };
    }
  });

  // Pre-compute column indices used in per-row overrides
  const linkIdx = cols.findIndex(c => c.key === 'View Link');
  const addlIdx = cols.findIndex(c => c.key === 'Additional Details');
  const bidIdx = cols.findIndex(c => c.key === 'Bid Document Details');
  const companyIdx   = cols.findIndex(c => c.key === 'Company');
  const importantIdx = cols.findIndex(c => c.key === 'Important');
  const emdIdx      = cols.findIndex(c => c.key === 'EMD');
  const emdExemptIdx = cols.findIndex(c => c.key === 'EMD Exempt?');
  const tvIdx       = cols.findIndex(c => c.key === 'Tender Value');
  const bidStatusIdx = cols.findIndex(c => c.key === 'Bid Status');

  rows.forEach((r, i) => {
    const values = cols.map(c => {
      const v = r[c.key] ?? '';
      return typeof v === 'string' ? v.replace(/\n/g, ' | ') : v;
    });
    const row = ws.addRow(values);
    dataRowStyle(row, i + 1, cols.length);

    row.height = 18;

    // Company dropdown — arrow list; for multi-company type comma-separated after selecting
    if (companyIdx >= 0) {
      row.getCell(companyIdx + 1).dataValidation = {
        type: 'list', allowBlank: true,
        formulae: ['"Gravity,Total Tech,Quickman"'],
        showErrorMessage: false,
        showInputMessage: true,
        promptTitle: 'Company',
        prompt: 'Pick one or type comma-separated e.g. Gravity,Total Tech',
      };
    }

    // Important — checkbox (TRUE/FALSE boolean)
    if (importantIdx >= 0) {
      const cell = row.getCell(importantIdx + 1);
      if (cell.value === '' || cell.value == null) cell.value = false;
      cell.dataValidation = { type: 'list', allowBlank: true, formulae: ['"TRUE,FALSE"'], showErrorMessage: false };
    }

    // Bid Status dropdown — Accepted / Rejected
    if (bidStatusIdx >= 0) {
      row.getCell(bidStatusIdx + 1).dataValidation = {
        type: 'list', allowBlank: true,
        formulae: ['"Accepted,Rejected"'],
        showErrorMessage: false,
      };
    }

    // Hyperlink on View Link column — use HYPERLINK formula so it works in both Excel and Google Sheets
    if (linkIdx >= 0) {
      const cell = row.getCell(linkIdx + 1);
      const href = r['View Link'];
      if (href && href.startsWith('http')) {
        cell.value = { formula: `HYPERLINK("${href}","View Tender")`, result: 'View Tender' };
        cell.font = { name: 'Calibri', size: 10, color: { argb: 'FF1565C0' }, underline: true };
      }
    }

    // Additional Details — compact Calibri, left-aligned, subtle colour
    if (addlIdx >= 0) {
      const cell = row.getCell(addlIdx + 1);
      cell.font = { name: 'Calibri', size: 9, color: { argb: 'FF1E3A5F' } };
      cell.alignment = { vertical: 'middle', wrapText: false, horizontal: 'left' };
    }

    // Bid Document Details — monospace so pipe-delimited columns line up
    if (bidIdx >= 0) {
      const cell = row.getCell(bidIdx + 1);
      cell.font = { name: 'Consolas', size: 9 };
      cell.alignment = { vertical: 'middle', wrapText: false, horizontal: 'left' };
    }

    // EMD and Tender Value — comma-separated number format so Excel sorts numerically
    if (emdIdx >= 0 && typeof r['EMD'] === 'number') {
      row.getCell(emdIdx + 1).numFmt = '#,##0';
    }
    if (tvIdx >= 0 && typeof r['Tender Value'] === 'number') {
      row.getCell(tvIdx + 1).numFmt = '#,##0';
    }

    // EMD Exempt? — centre-align
    if (emdExemptIdx >= 0) {
      const cell = row.getCell(emdExemptIdx + 1);
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.font = { name: 'Calibri', size: 10, bold: true };
    }
  });

  ws.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1 + rows.length, column: cols.length },
  };

  const endRow = rows.length + 10;
  const endCol = colNumToLetter(cols.length);

  // Bid Status column — Accepted=dark green, Rejected=dark red
  if (bidStatusIdx >= 0) {
    const bsCol = colNumToLetter(bidStatusIdx + 1);
    ws.addConditionalFormatting({
      ref: `${bsCol}2:${bsCol}${endRow}`,
      rules: [
        {
          type: 'cellIs', operator: 'equal', formulae: ['"Accepted"'], priority: 1,
          style: {
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D32' } },
            font: { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } }
          }
        },
        {
          type: 'cellIs', operator: 'equal', formulae: ['"Rejected"'], priority: 2,
          style: {
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC62828' } },
            font: { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } }
          }
        },
      ],
    });
  }

  // EMD Exempt? column — Yes=green, No=light grey
  if (emdExemptIdx >= 0) {
    const exCol = colNumToLetter(emdExemptIdx + 1);
    ws.addConditionalFormatting({
      ref: `${exCol}2:${exCol}${endRow}`,
      rules: [
        {
          type: 'cellIs', operator: 'equal', formulae: ['"Yes"'], priority: 3,
          style: {
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D32' } },
            font: { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } }
          }
        },
        {
          type: 'cellIs', operator: 'equal', formulae: ['"No"'], priority: 4,
          style: {
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB0BEC5' } },
            font: { name: 'Calibri', size: 10, bold: true, color: { argb: 'FF37474F' } }
          }
        },
      ],
    });
  }

  // Row highlighting — yellow when Important is checked
  if (importantIdx >= 0) {
    const impCol = colNumToLetter(importantIdx + 1);
    ws.addConditionalFormatting({
      ref: `A2:${endCol}${endRow}`,
      rules: [
        { type: 'expression', formulae: [`$${impCol}2=TRUE`], priority: 101,
          style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF59D' } } } },
      ],
    });
  }
}

// ══════════════════════════════════════════════════════════════
//  BUILD WORKBOOK
// ══════════════════════════════════════════════════════════════
async function buildExcel(sections) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Gravity Scraper';
  wb.company = 'Gravity';
  wb.created = new Date();
  wb.modified = new Date();

  const usedNames = new Set();
  function sectionBaseName(raw) {
    return raw.replace(/\(\d+\)\s*$/, '').replace(/[\\/:?*[\]]/g, '').trim() || 'Sheet';
  }
  function uniqueName(raw) {
    let base = raw.replace(/[\\/:?*[\]]/g, '_').trim().slice(0, 31) || 'Sheet';
    let attempt = base;
    let n = 2;
    while (usedNames.has(attempt.toLowerCase())) attempt = base.slice(0, 28) + '_' + (n++);
    usedNames.add(attempt.toLowerCase());
    return attempt;
  }

  const allRows = sections.flatMap(s => s.tenders.map(t => ({ Section: s.section, ...t })));

  // Sort GeM tenders first across all sections (stable — preserves original order within each group)
  allRows.sort((a, b) => {
    const aG = /gem/i.test(String(a['Information Source'] || '')) ? 0 : 1;
    const bG = /gem/i.test(String(b['Information Source'] || '')) ? 0 : 1;
    return aG - bG;
  });

  // Use fixed columns, dropping the unused Bid Document Details
  const finalCols = COLS.filter(c => c.key !== 'Bid Document Details');

  // Columns in All Sections — Filled / Filled Date / Filled By / Bid Status live only on company sheets
  const MASTER_INPUT_KEYS = new Set(['Company', 'Important']);
  const ALL_INPUT_KEYS    = new Set(['Company', 'Important', 'Filled Date', 'Filled By', 'Bid Status']);
  const masterCols = [
    ...finalCols.filter(c => MASTER_INPUT_KEYS.has(c.key)),
    { key: 'Section', label: 'Section', width: 28 },
    ...finalCols.filter(c => !ALL_INPUT_KEYS.has(c.key)),
  ];

  const lastCol = colNumToLetter(masterCols.length);
  const maxRow  = Math.max(allRows.length + 10, 1000);
  // Dynamic column letter for Section — used in Corrigendum and section-tab filter formulas
  const sectionColLetter = colNumToLetter(masterCols.findIndex(c => c.key === 'Section') + 1);

  // ── "All Sections" — only the 5 user-input columns (A–E) are editable.
  //    Everything else (scraped data) is locked.
  const allWs = wb.addWorksheet(uniqueName('All Sections'));
  fillDataSheet(allWs, masterCols, allRows, 'FF2C3E50', '2C3E50');

  // Hide Company column in All Sections — company is now assigned in the Important tab
  allWs.getColumn(1).hidden = true;

  // Important (col B) stays visible as a checkbox
  allWs.getRow(1).getCell(2).note = 'Check this box to mark the tender as Important, then assign a company in the Important tab';

  //  Helper: builds a formula-driven filter sheet.
  //
  //  isCompanySheet = true:
  //    • Columns A = "Filled Date", B = "Filled By" (editable, placed first)
  //    • Columns C onwards = FILTER formula result (Company col skipped via CHOOSECOLS)
  //    • No sheet protection (Apps Script locks cols C+ via range protection)
  //  Other sheets:
  //    • Columns A onwards = FILTER formula result (all masterCols)
  //    • Apps Script applies data-row range protection
  function makeFilterSheet(name, tabArgb, hdrHex, filterCol, filterValue, options = {}) {
    const ws = wb.addWorksheet(uniqueName(name));
    ws.properties = { tabColor: { argb: tabArgb } };
    ws.views      = [{ state: 'frozen', ySplit: 1, showGridLines: true }];

    // Company sheets: Filled/FilledDate/FilledBy prefix
    // Important tab: Company prefix (user assigns company after marking as important)
    const inputColDefs = options.isCompanySheet
      ? [
          { label: 'Filled',      key: 'FilledCheck', width: 10 },
          { label: 'Filled Date', key: 'FilledDate',  width: 16 },
          { label: 'Filled By',   key: 'FilledBy',    width: 18 },
        ]
      : name === 'Important'
        ? [{ label: 'Filled?', key: 'FilledStatus', width: 10 }, { label: 'Company', key: 'Company', width: 16 }]
        : [];

    // Important tab: strip Company+Important from formula output (both are handled separately)
    // Company sheets: strip Company+Important (masterCols.slice(2))
    // Filled: widen Company col; others: plain masterCols
    const displayCols = options.isCompanySheet
      ? masterCols.slice(2)
      : name === 'Important'
        ? masterCols.filter(function(c) { return c.key !== 'Company' && c.key !== 'Important'; })
        : name === 'Filled'
          ? masterCols.map(function(c) { return c.key === 'Company' ? Object.assign({}, c, { width: 32 }) : c; })
          : masterCols;
    const displayColCount = displayCols.length;

    const allSheetCols  = [...inputColDefs, ...displayCols];
    const totalColCount = allSheetCols.length;
    const totalLastCol  = colNumToLetter(totalColCount);

    const inputOffset = inputColDefs.length;  // 0 or 3

    ws.properties.defaultRowHeight = 18;
    ws.columns = allSheetCols.map(c => ({
      header: c.label, key: c.key, width: c.width,
      style: { font: { name: 'Calibri', size: 10 }, alignment: { vertical: 'top', wrapText: false } },
    }));
    const hdr  = ws.getRow(1);
    hdr.values = allSheetCols.map(c => c.label);
    headerStyle(hdr, hdrHex, totalColCount);

    // Column letters inside All Sections (for formula references)
    const sColLetter = colNumToLetter(masterCols.findIndex(c => c.key === 'Section') + 1);
    const bColLetter = colNumToLetter(masterCols.findIndex(c => c.key === 'Tender Brief') + 1);

    // Company sheets pull from Important tab (cols A-B are Filled?+Company; formula output starts at C)
    // All other sheets pull from All Sections as usual
    const impOutputLastCol = colNumToLetter(masterCols.length); // Important has 2 input cols → formula output ends at masterCols.length
    const src       = options.isCompanySheet
      ? `'Important'!C2:${impOutputLastCol}${maxRow}`
      : `'All Sections'!A2:${lastCol}${maxRow}`;
    const filterSheetName = options.isCompanySheet ? 'Important' : 'All Sections';
    const critRange = `'${filterSheetName}'!$${filterCol}$2:$${filterCol}$${maxRow}`;

    // CHOOSECOLS seq for Important tab: strip col 1 (Company) and col 2 (Important) from All Sections output
    const impColSeq = '{' + Array.from({ length: masterCols.length - 2 }, (_, i) => i + 3).join(',') + '}';
    // Fallback empty row — width must match the formula output column count
    const emptyRow = '{' + Array(displayColCount).fill('""').join(',') + '}';

    // ── Build formula
    let filterFormula;
    if (name === 'Corrigendum') {
      const sectionRange = `'All Sections'!$${sColLetter}$2:$${sColLetter}$${maxRow}`;
      const briefRange   = `'All Sections'!$${bColLetter}$2:$${bColLetter}$${maxRow}`;
      filterFormula = `IFERROR(FILTER(${src},(ISNUMBER(SEARCH("corrigendum",${sectionRange})))+(ISNUMBER(SEARCH("corrigendum",${briefRange})))>0),IF(ROWS(A2:A2)=1,"No Corrigendum tenders found",""))`;
    } else if (name === 'Important') {
      // Filter Important=TRUE rows, strip Company (col1) + Important (col2) from output via CHOOSECOLS
      // Formula placed at C2 (col A = Filled?, col B = Company)
      filterFormula = `IFERROR(CHOOSECOLS(FILTER(${src},${critRange}=TRUE),${impColSeq}),IF(ROWS(C2:C2)=1,"No important tenders",""))`;
    } else if (options.isCompanySheet) {
      // Filter from Important tab by Company prefix (col A); source already has Company+Important stripped
      const cond = `ISNUMBER(SEARCH("${filterValue}",${critRange}))`;
      filterFormula = `IFERROR(FILTER(${src},${cond}),${emptyRow})`;
    } else {
      filterFormula = `IFERROR(FILTER(${src},ISNUMBER(SEARCH("${filterValue}",${critRange}))),IF(ROWS(A2:A2)=1,"No tenders assigned to ${name}",""))`;
    }

    // Place formula: C2 for company sheets (cols A-B are the editable input cols), A2 otherwise
    const formulaCell = inputOffset > 0 ? colNumToLetter(inputOffset + 1) + '2' : 'A2';
    ws.getCell(formulaCell).value = { formula: filterFormula };

    // Do NOT set fixed row heights — Google Sheets auto-sizes rows based on FILTER formula content

    // Zebra stripe
    ws.addConditionalFormatting({
      ref: `A2:${totalLastCol}${maxRow}`,
      rules: [
        { type: 'expression', formulae: ['MOD(ROW(),2)=0'], priority: 200,
          style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF4FB' } } } },
      ],
    });

    // Row highlighting:
    // • Company sheets: green when Filled checkbox (col A) = TRUE
    // • All other sheets: yellow when Important checkbox = TRUE (col B in masterCols display)
    if (options.isCompanySheet) {
      ws.addConditionalFormatting({
        ref: `A2:${totalLastCol}${maxRow}`,
        rules: [
          { type: 'expression', formulae: ['$A2=TRUE'], priority: 100,
            style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFA5D6A7' } } } },
        ],
      });
    } else {
      const impIdx = displayCols.findIndex(function(c) { return c.key === 'Important'; });
      if (impIdx >= 0) {
        const impCF = colNumToLetter(impIdx + 1 + inputOffset);
        ws.addConditionalFormatting({
          ref: `A2:${totalLastCol}${maxRow}`,
          rules: [
            { type: 'expression', formulae: [`$${impCF}2=TRUE`], priority: 101,
              style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF59D' } } } },
          ],
        });
      }
    }

    // EMD Exempt? column — Yes=green, No=light grey (matches All Sections styling)
    const emdExemptPos = allSheetCols.findIndex(c => c.key === 'EMD Exempt?');
    if (emdExemptPos >= 0) {
      const exCol = colNumToLetter(emdExemptPos + 1);
      ws.addConditionalFormatting({
        ref: `${exCol}2:${exCol}${maxRow}`,
        rules: [
          { type: 'cellIs', operator: 'equal', formulae: ['"Yes"'], priority: 3,
            style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D32' } },
                     font: { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } } } },
          { type: 'cellIs', operator: 'equal', formulae: ['"No"'], priority: 4,
            style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB0BEC5' } },
                     font: { name: 'Calibri', size: 10, bold: true, color: { argb: 'FF37474F' } } } },
        ],
      });
    }

    // AutoFilter across all display columns (input cols A-B + formula cols C+)
    ws.autoFilter = {
      from: { row: 1, column: 1 },
      to:   { row: 1, column: totalColCount },
    };

    // Important tab: Filled? in col A (BYROW formula), Company dropdown in col B
    if (name === 'Important') {
      // Company dropdown — col B (col 2), since col A is Filled?
      for (let r = 2; r <= Math.min(maxRow, 1000); r++) {
        ws.getCell(r, 2).dataValidation = {
          type: 'list', allowBlank: true,
          formulae: ['"Gravity,Total Tech,Quickman"'],
          showErrorMessage: false,
          showInputMessage: true,
          promptTitle: 'Company',
          prompt: 'Pick one or type comma-separated e.g. Gravity,Total Tech',
        };
      }

      // BYROW formula in col A — checks TDR against all company sheets' Filled checkbox
      const tdrPosInDisplay = displayCols.findIndex(c => c.key === 'TDR');
      if (tdrPosInDisplay >= 0) {
        // inputOffset = 2 (Filled?, Company), so TDR is at col 2 + tdrPosInDisplay + 1
        const tdrColInImp  = colNumToLetter(inputOffset + tdrPosInDisplay + 1);  // D
        const tdrColInComp = colNumToLetter(3 + tdrPosInDisplay + 1);            // E (company sheets have 3 input cols)

        ws.getCell(2, 1).value = { formula:
          `BYROW(${tdrColInImp}2:${tdrColInImp}${maxRow},LAMBDA(r,IF(r="","",` +
          `IF(COUNTIFS(Gravity!$${tdrColInComp}$2:$${tdrColInComp}$${maxRow},r,Gravity!$A$2:$A$${maxRow},TRUE)+` +
          `COUNTIFS('Total Tech'!$${tdrColInComp}$2:$${tdrColInComp}$${maxRow},r,'Total Tech'!$A$2:$A$${maxRow},TRUE)+` +
          `COUNTIFS(Quickman!$${tdrColInComp}$2:$${tdrColInComp}$${maxRow},r,Quickman!$A$2:$A$${maxRow},TRUE)>0,` +
          `"Yes","No"))))` };

        // Green for Yes, red for No in col A
        ws.addConditionalFormatting({
          ref: `A2:A${maxRow}`,
          rules: [
            { type: 'containsText', operator: 'containsText', text: 'Yes', priority: 48,
              style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFA5D6A7' } }, font: { bold: true } } },
            { type: 'containsText', operator: 'containsText', text: 'No',  priority: 49,
              style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEF9A9A' } }, font: { bold: true } } },
          ],
        });
      }
    }

    // Hide Company and Important in formula-driven sheets — they show redundant TRUE/FALSE values
    // (Company is hidden in All Sections; Important is only meaningful as a checkbox there)
    ['Company', 'Important'].forEach(function(key) {
      const idx = displayCols.findIndex(function(c) { return c.key === key; });
      if (idx >= 0) ws.getColumn(idx + 1 + inputOffset).hidden = true;
    });

    return ws;
  }

  // ── Sheet order: All Sections → Gravity → TT → Quickman → Important → Filled → Corrigendum → sections
  // Company sheets come right after All Sections so they're the first thing users see
  const masterColor = '2C3E50';
  const masterArgb  = 'FF2C3E50';

  // Company sheets — fully unlocked; Filled Date / Filled By entered here
  makeFilterSheet('Gravity',    masterArgb, masterColor, 'B', 'Gravity',    { isCompanySheet: true });
  makeFilterSheet('Total Tech', masterArgb, masterColor, 'B', 'Total Tech', { isCompanySheet: true });
  makeFilterSheet('Quickman',   masterArgb, masterColor, 'B', 'Quickman',   { isCompanySheet: true });

  // Important: filterCol = Important column (B in All Sections), formula uses =TRUE (checkbox)
  makeFilterSheet('Important', 'FFB71C1C', masterColor, colNumToLetter(masterCols.findIndex(function(c) { return c.key === 'Important'; }) + 1), 'Yes');

  // Filled sheet — references each company sheet's Filled checkbox (col A) directly
  (function() {
    const filledWs = wb.addWorksheet(uniqueName('Filled'));
    filledWs.properties = { tabColor: { argb: 'FF00695C' } };
    filledWs.views = [{ state: 'frozen', ySplit: 1, showGridLines: true }];

    // Company sheet layout: A=Filled, B=FilledDate, C=FilledBy, D..=FILTER output (masterCols.slice(2))
    const coDisplayCols  = masterCols.slice(2);          // Section, TDR, ...
    const coLastCol      = colNumToLetter(3 + coDisplayCols.length); // D + displayColCount - 1

    // Filled sheet columns: Company | FilledDate | FilledBy | Section | TDR | ...
    const filledSheetCols = [
      { key: 'Company',    label: 'Company',     width: 32 },
      { key: 'FilledDate', label: 'Filled Date', width: 16 },
      { key: 'FilledBy',   label: 'Filled By',   width: 18 },
    ].concat(coDisplayCols);
    const filledTotal   = filledSheetCols.length;
    const filledLastCol = colNumToLetter(filledTotal);

    filledWs.columns = filledSheetCols.map(function(c) { return { header: c.label, key: c.key, width: c.width }; });
    const filledHdr = filledWs.getRow(1);
    filledHdr.values = filledSheetCols.map(function(c) { return c.label; });
    headerStyle(filledHdr, masterColor, filledTotal);

    // VSTACK: for each company, filter rows where col A = TRUE, prefix with company name
    const companies = ['Gravity', 'Total Tech', 'Quickman'];
    const filledEmpty = '{' + Array(filledTotal).fill('""').join(',') + '}';
    const fParts = companies.map(function(co) {
      const sq = co.indexOf(' ') >= 0 ? ("'" + co + "'") : co;
      return 'IFERROR(FILTER(HSTACK({"' + co + '"},' + sq + '!B2:' + coLastCol + maxRow + '),' + sq + '!A2:A' + maxRow + '=TRUE),' + filledEmpty + ')';
    });
    filledWs.getCell('A2').value = { formula: 'IFERROR(VSTACK(' + fParts.join(',') + '),IF(ROWS(A2:A2)=1,"No filled tenders",""))' };

    filledWs.addConditionalFormatting({
      ref: 'A2:' + filledLastCol + maxRow,
      rules: [
        { type: 'expression', formulae: ['MOD(ROW(),2)=0'], priority: 200,
          style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF4FB' } } } },
        { type: 'expression', formulae: ['$A2<>""'], priority: 100,
          style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFA5D6A7' } } } },
      ],
    });
    filledWs.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: filledTotal } };
  }());
  makeFilterSheet('Corrigendum', 'FFFF6F00', masterColor, sectionColLetter, 'corrigendum');

  const SECTION_COLORS = [
    'FF1565C0', 'FF00695C', 'FFE65100', 'FF4A148C',
    'FF006064', 'FF1B5E20', 'FF4E342E', 'FF37474F',
    'FF880E4F', 'FF0D47A1',
  ];
  sections.forEach((sec, idx) => {
    const name      = sectionBaseName(sec.section);
    const tabArgb   = SECTION_COLORS[idx % SECTION_COLORS.length];
    const hdrColor  = tabArgb.slice(2); // strip FF prefix for headerStyle
    makeFilterSheet(name, tabArgb, hdrColor, sectionColLetter, name);
  });

  return wb.xlsx.writeBuffer();
}

// ══════════════════════════════════════════════════════════════
//  PHASE 1 — Parse listing page → sections + view links
// ══════════════════════════════════════════════════════════════
function parseDailyDigest($) {
  const sections = [];
  let curSection = null;
  let curTenders = [];

  $('p.m-r-government-tenders, div.m-mainTR').each((_, el) => {
    const $el = $(el);
    const cls = $el.attr('class') || '';

    if (cls.includes('m-r-government-tenders')) {
      if (curTenders.length > 0) sections.push({ section: curSection || 'All Tenders', tenders: curTenders });
      curSection = clean($el.text());
      curTenders = [];
      return;
    }

    const $row2 = $el.children('div.row').eq(1);
    const $desc = $el.children('div.row').eq(0).find('.col-md-12').first();
    let tenderId = 'N/A';
    $desc.find('p.m-td-brief').each((_, p) => {
      const t = clean($(p).text());
      if (/TDR:\d+/i.test(t) && tenderId === 'N/A') {
        const m = t.match(/TDR:(\d+)/i);
        if (m) tenderId = m[1];
      }
    });

    let viewLink = $row2.find('a').first().attr('href') || null;
    if (viewLink && !viewLink.startsWith('http')) viewLink = `https://www.tenderdetail.com${viewLink}`;
    if (viewLink) curTenders.push({ tenderId, viewLink });
  });

  if (curTenders.length > 0) sections.push({ section: curSection || 'All Tenders', tenders: curTenders });
  return sections;
}

// ══════════════════════════════════════════════════════════════
//  PHASE 2 — Scrape one detail page
// ══════════════════════════════════════════════════════════════

// Keys that are mapped to dedicated columns — anything else goes
// into the "Additional Details" cell.
const MAPPED_KEYS = new Set([
  'TDR', 'Tender No', 'Tender ID', 'Tendering Authority', 'Company Name',
  'Tender Brief', 'City', 'State', 'Document Fees', 'EMD', 'Tender Value',
  'Tender Type', 'Bidding Type', 'Competition Type', 'Publish Date',
  'Last Date of Bid Submission', 'Tender Opening Date', 'Address', 'Information Source',
]);

async function scrapeTenderDetail(viewLink, geminiKey) {
  try {
    const { data } = await axios.get(viewLink, { headers: HEADERS, timeout: 15000, maxRedirects: 4 });
    const $ = cheerio.load(data);
    const record = {};
    const docLinks = [];

    $('table tr').each((_, row) => {
      const $tds = $(row).find('td');
      if ($tds.length >= 2) {
        const label = clean($tds.eq(0).text());
        const value = clean($tds.eq(1).text());
        if (label.startsWith('Download')) {
          // Check both columns for hrefs (link is in first td on tenderdetail.com)
          $(row).find('a[href]').each((__, a) => {
            let href = ($(a).attr('href') || '').replace(/\\/g, '/');
            if (href && !href.startsWith('#') && !href.startsWith('javascript')) {
              if (!href.startsWith('http')) href = `https://www.tenderdetail.com${href}`;
              if (!docLinks.includes(href)) docLinks.push(href);
            }
          });
        } else if (label && value && label.length < 80) {
          record[label] = value;
        }
      }
    });

    // Broader fallback: any link on the page that looks like a document
    if (docLinks.length === 0) {
      $('a[href]').each((_, a) => {
        const raw  = $(a).attr('href') || '';
        const href = raw.replace(/\\/g, '/');
        const text = clean($(a).text()).toLowerCase();
        const lhref = href.toLowerCase();
        if (
          /\.(pdf|doc|docx)(\?|$)/i.test(href) ||
          /download|document|corrigendum|nit|boq|tender.doc/i.test(text) ||
          /\/download\//i.test(lhref) ||
          /\/document\//i.test(lhref) ||
          /tenderfiles\.com/i.test(lhref)
        ) {
          if (!href.startsWith('#') && !href.startsWith('javascript')) {
            const full = href.startsWith('http') ? href : `https://www.tenderdetail.com${href}`;
            if (!docLinks.includes(full)) docLinks.push(full);
          }
        }
      });
    }

    console.log(`[${record['TDR'] || viewLink}] docLinks found: ${docLinks.length}`, docLinks.slice(0, 3));

    // Collect every detail-page field that is not already in a dedicated column
    const additionalLines = [];
    for (const [k, v] of Object.entries(record)) {
      if (!MAPPED_KEYS.has(k)) additionalLines.push(`${k}: ${v}`);
    }

    // Pick only the UUID-matching NIT document (or ≤3 fallback), fetch once, reuse for everything
    const targetLinks = pickTargetDocLinks(viewLink, docLinks);
    const docTexts    = targetLinks.length ? await Promise.all(targetLinks.map(link => fetchDocText(link))) : [];

    const emdExemption = await extractEmdExemptionFromTexts(docTexts, geminiKey);

    // Use HTML-scraped values; fill from doc text if HTML only says "refer document"
    let emdAmount   = normalizeAmount(record['EMD']);
    let tenderValue = normalizeAmount(record['Tender Value']);
    if (docTexts.length && (needsDocLookup(emdAmount) || needsDocLookup(tenderValue))) {
      const amounts = extractAmountsFromTexts(docTexts);
      if (needsDocLookup(emdAmount)   && amounts.emd !== null) emdAmount   = amounts.emd;
      if (needsDocLookup(tenderValue) && amounts.tv  !== null) tenderValue = amounts.tv;
    }

    return {
      'Company': '',
      'Important': false,
      'Filled Date': '',
      'Filled By': '',
      'Bid Status': '',
      'TDR': record['TDR'] || 'N/A',
      'Tender No': record['Tender No'] || record['Tender ID'] || 'N/A',
      'Tendering Authority': record['Tendering Authority'] || record['Company Name'] || 'N/A',
      'Tender Brief': record['Tender Brief'] || 'N/A',
      'City': record['City'] || 'N/A',
      'State': record['State'] || 'N/A',
      'Document Fees': record['Document Fees'] || 'N/A',
      'EMD': emdAmount,
      'EMD Exempt?': /^(no exemption mentioned|n\/a)$/i.test(emdExemption.trim()) ? 'No' : 'Yes',
      'EMD Exemption': emdExemption,
      'Tender Value': tenderValue,
      'Tender Type': record['Tender Type'] || 'N/A',
      'Bidding Type': record['Bidding Type'] || 'N/A',
      'Competition Type': record['Competition Type'] || 'N/A',
      'Publish Date': record['Publish Date'] || 'N/A',
      'Last Date of Bid Submission': record['Last Date of Bid Submission'] || 'N/A',
      'Tender Opening Date': record['Tender Opening Date'] || 'N/A',
      'Address': record['Address'] || 'N/A',
      'Information Source': record['Information Source'] || 'N/A',
      'View Link': viewLink,
      'Additional Details': additionalLines.join('\n') || 'N/A'
    };
  } catch (e) {
    return {
      'Company': '', 'Important': false, 'Filled Date': '', 'Filled By': '', 'Bid Status': '',
      'TDR': 'N/A', 'Tender No': 'N/A', 'Tendering Authority': 'N/A',
      'Tender Brief': `Error: ${e.message}`, 'City': 'N/A', 'State': 'N/A',
      'Document Fees': 'N/A', 'EMD': 'N/A', 'EMD Exempt?': 'N/A', 'EMD Exemption': 'N/A', 'Tender Value': 'N/A',
      'Tender Type': 'N/A', 'Bidding Type': 'N/A', 'Competition Type': 'N/A',
      'Publish Date': 'N/A', 'Last Date of Bid Submission': 'N/A',
      'Tender Opening Date': 'N/A', 'Address': 'N/A', 'Information Source': 'N/A',
      'View Link': viewLink, 'Additional Details': 'N/A', 'Bid Document Details': 'N/A',
    };
  }
}

// ══════════════════════════════════════════════════════════════
//  Concurrency pool
// ══════════════════════════════════════════════════════════════
async function pooledMap(items, fn, limit, onProgress) {
  const results = new Array(items.length);
  let index = 0;
  async function worker() {
    while (index < items.length) {
      const i = index++;
      results[i] = await fn(items[i]);
      if (onProgress) onProgress(i + 1, items.length);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return results;
}

function sseWrite(res, event, data) {
  res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
  if (typeof res.flush === 'function') res.flush();
}

// ══════════════════════════════════════════════════════════════
//  GET /scrape-deep — SSE: fetch all detail pages → build Excel
// ══════════════════════════════════════════════════════════════
app.get('/scrape-deep', async (req, res) => {
  const { url, geminiKey: reqGeminiKey } = req.query;
  if (!url || !validateUrl(url)) {
    res.status(400).end();
    return;
  }

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders();

  try {
    sseWrite(res, 'status', { phase: 1, message: 'Fetching listing page…' });

    const { data } = await axios.get(url.split('#')[0], { headers: HEADERS, timeout: 20000, maxRedirects: 5 });
    const $ = cheerio.load(data);
    const pageTitle = clean($('title').text()) || 'Tenders';
    
    // Try to extract date from title (e.g., "15-05-2026")
    const dateMatch = pageTitle.match(/\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}/);
    const tenderDate = dateMatch ? dateMatch[0].replace(/\//g, '-') : new Date().toLocaleDateString('en-GB').replace(/\//g, '-');

    $('script, style').remove();

    const sections = parseDailyDigest($);
    const total = sections.reduce((n, s) => n + s.tenders.length, 0);

    sseWrite(res, 'status', {
      phase: 2,
      message: `Found ${total} tenders across ${sections.length} sections. Visiting detail pages…`,
      total,
      sections: sections.map(s => ({ section: s.section, count: s.tenders.length })),
    });

    let globalDone = 0;
    const enriched = [];

    for (const sec of sections) {
      sseWrite(res, 'section_start', { section: sec.section, count: sec.tenders.length });

      const records = await pooledMap(
        sec.tenders,
        (t) => scrapeTenderDetail(t.viewLink, reqGeminiKey || undefined),
        CONCURRENCY,
        (sectionDone) => {
          globalDone++;
          sseWrite(res, 'progress', {
            globalDone, total,
            sectionDone, sectionTotal: sec.tenders.length,
            section: sec.section,
          });
        }
      );

      enriched.push({ section: sec.section, tenders: records });
      sseWrite(res, 'section_done', { section: sec.section, count: records.length });
    }

    sseWrite(res, 'status', { phase: 3, message: 'Building Excel file…' });
    const excelBuf = await buildExcel(enriched);

    const token    = crypto.randomBytes(16).toString('hex');
    const safeName = tenderDate; // Use the date as the filename
    const filePath = path.join(TMP_DIR, `${token}.xlsx`);
    fs.writeFileSync(filePath, excelBuf);
    setTimeout(() => { try { fs.unlinkSync(filePath); } catch (_) {} }, 30 * 60 * 1000);

    // Phase 4 — Google Sheets upload with fake progress ticks
    sseWrite(res, 'status', { phase: 4, message: 'Uploading to Google Sheets…', total });
    let sheetsPct = 0;
    const sheetsTicker = setInterval(() => {
      sheetsPct = Math.min(90, sheetsPct + Math.floor(Math.random() * 8) + 4);
      sseWrite(res, 'sheets_progress', { pct: sheetsPct });
    }, 1800);
    let sheetsUrl;
    try {
      // Retry up to 3 times — transient DNS failures (ENOTFOUND script.google.com) are common
      for (let attempt = 1; attempt <= 3; attempt++) {
        sheetsUrl = await uploadToGoogleSheets(excelBuf, safeName);
        if (sheetsUrl) break;
        if (attempt < 3) {
          console.warn(`Sheets upload attempt ${attempt} failed, retrying in ${attempt * 2}s…`);
          await new Promise(r => setTimeout(r, attempt * 2000));
        }
      }
    } finally {
      clearInterval(sheetsTicker);
    }
    sseWrite(res, 'sheets_progress', { pct: 100 });

    sseWrite(res, 'done', { token, filename: safeName, totalTenders: total, sections: sections.length, pageTitle, sheetsUrl });
  } catch (err) {
    sseWrite(res, 'error', { message: err.message });
  } finally {
    res.end();
  }
});

// ══════════════════════════════════════════════════════════════
//  GET /download/:token
// ══════════════════════════════════════════════════════════════
app.get('/download/:token', (req, res) => {
  const token = req.params.token.replace(/[^a-f0-9]/gi, '');
  const filePath = path.join(TMP_DIR, `${token}.xlsx`);
  const filename = (req.query.name || 'tender_data').replace(/[^a-z0-9_\-]/gi, '_') + '.xlsx';

  if (!fs.existsSync(filePath)) return res.status(404).send('File expired. Please scrape again.');

  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.sendFile(filePath);
});

// ══════════════════════════════════════════════════════════════
//  GET /retry-sheets/:token — re-attempt Sheets upload without re-scraping
// ══════════════════════════════════════════════════════════════
app.get('/retry-sheets/:token', async (req, res) => {
  const token = req.params.token.replace(/[^a-f0-9]/gi, '');
  const filePath = path.join(TMP_DIR, `${token}.xlsx`);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'File expired. Please scrape again.' });

  try {
    const buf = fs.readFileSync(filePath);
    const name = (req.query.name || 'tender_data').replace(/[^a-z0-9_\-]/gi, '_');
    let sheetsUrl;
    for (let attempt = 1; attempt <= 3; attempt++) {
      sheetsUrl = await uploadToGoogleSheets(buf, name);
      if (sheetsUrl) break;
      if (attempt < 3) await new Promise(r => setTimeout(r, attempt * 2000));
    }
    if (sheetsUrl) return res.json({ sheetsUrl });
    return res.status(502).json({ error: 'Upload failed after 3 attempts. Check server connectivity.' });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
//  Helpers
// ══════════════════════════════════════════════════════════════
function cleanOldFiles() {
  try {
    if (!fs.existsSync(TMP_DIR)) return;
    const cutoff = Date.now() - 30 * 60 * 1000;
    fs.readdirSync(TMP_DIR).forEach(f => {
      const fp = path.join(TMP_DIR, f);
      try { if (fs.statSync(fp).mtimeMs < cutoff) fs.unlinkSync(fp); } catch (_) { }
    });
  } catch (_) { }
}

app.listen(PORT, () => console.log(`\n🚀 Gravity Scraper → http://localhost:${PORT}\n`));
