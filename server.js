const express  = require('express');
const axios    = require('axios');
const cheerio  = require('cheerio');
const ExcelJS  = require('exceljs');
const cors     = require('cors');
const path     = require('path');
const fs       = require('fs');
const crypto   = require('crypto');
const { PDFParse } = require('pdf-parse');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

const TMP_DIR = path.join(__dirname, 'tmp');
if (!fs.existsSync(TMP_DIR)) fs.mkdirSync(TMP_DIR);
cleanOldFiles();

const HEADERS = {
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
  'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
  'Accept-Language': 'en-US,en;q=0.5',
};
const CONCURRENCY = 8;


function clean(t) { return (t || '').replace(/\s+/g, ' ').trim(); }

// Normalise multi-line strings that go into text-heavy cells:
// trim trailing spaces per line, collapse 3+ blank lines → 1, trim ends.
function cleanBidText(text) {
  return text
    .split('\n')
    .map(l => l.trimEnd())
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

// Convert EMD / Tender Value strings to plain rupee integers.
// Handles: "2.5 Crores", "50 Lacs", "₹ 1,23,456", "50000/-", "Nil", etc.
function normalizeAmount(val) {
  if (!val) return 'N/A';
  const raw = String(val).trim();
  if (/^(nil|n\/a|not\s*applicable|na|-|exempt)$/i.test(raw)) return 'N/A';
  const lo = raw.toLowerCase();
  // "X crore Y lac" combo
  const combo = lo.match(/([0-9.]+)\s*crore[s]?\s*(?:and\s*)?([0-9.]+)\s*(?:lac|lakh)/);
  if (combo) return String(Math.round(+combo[1] * 1e7 + +combo[2] * 1e5));
  // crore / cr
  const cr = lo.match(/([0-9.]+)\s*(?:crore[s]?|cr\.?)\b/);
  if (cr) return String(Math.round(+cr[1] * 1e7));
  // lac / lakh
  const lac = lo.match(/([0-9.]+)\s*(?:lac[s]?|lakh[s]?)\b/);
  if (lac) return String(Math.round(+lac[1] * 1e5));
  // thousand / k
  const k = lo.match(/([0-9.]+)\s*(?:thousand|k)\b/);
  if (k) return String(Math.round(+k[1] * 1e3));
  // plain number (strip currency symbols, commas, slashes)
  const num = parseFloat(raw.replace(/[₹$,\s\/-]/g, '').replace(/[^0-9.]/g, ''));
  if (!isNaN(num) && num > 0) return String(Math.round(num));
  return val;
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
  // ── User-input dropdowns (blank on scrape; filled by user in Excel)
  { key: 'Company', label: 'Company', width: 16 },
  { key: 'Status',  label: 'Status',  width: 14 },
  // ── Fixed scraped fields
  { key: 'TDR',                         label: 'TDR',                          width: 14 },
  { key: 'Tender No',                   label: 'Tender No',                    width: 26 },
  { key: 'Tendering Authority',         label: 'Tendering Authority',          width: 36 },
  { key: 'Tender Brief',                label: 'Tender Brief',                 width: 62 },
  { key: 'City',                        label: 'City',                         width: 16 },
  { key: 'State',                       label: 'State',                        width: 18 },
  { key: 'Document Fees',               label: 'Document Fees',                width: 16 },
  { key: 'EMD',                         label: 'EMD (₹)',                      width: 18 },
  { key: 'Tender Value',                label: 'Tender Value (₹)',             width: 18 },
  { key: 'Tender Type',                 label: 'Tender Type',                  width: 16 },
  { key: 'Bidding Type',                label: 'Bidding Type',                 width: 16 },
  { key: 'Competition Type',            label: 'Competition Type',             width: 18 },
  { key: 'Publish Date',                label: 'Publish Date',                 width: 14 },
  { key: 'Last Date of Bid Submission', label: 'Last Date of Bid Submission',  width: 26 },
  { key: 'Tender Opening Date',         label: 'Tender Opening Date',          width: 22 },
  { key: 'Address',                     label: 'Address',                      width: 34 },
  { key: 'Information Source',          label: 'Information Source',           width: 30 },
  { key: 'View Link',                   label: 'View Link',                    width: 60 },
  // ── Variable / catch-all
  { key: 'Additional Details',   label: 'Additional Details',   width: 48 },
  { key: 'Bid Document Details', label: 'Bid Document Details', width: 70 },
];

// Columns whose values may contain \n — used for row-height calculation
const TEXT_WRAP_COLS = new Set(['Additional Details', 'Bid Document Details']);

// ── Per-section tab colours (ARGB, fully opaque)
const TAB_COLORS = [
  'FF1565C0', 'FF00695C', 'FFB71C1C', 'FFE65100',
  'FF1B5E20', 'FF4E342E', 'FF4A148C', 'FF006064',
  'FF1A237E', 'FF37474F',
];

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
    cell.alignment = { vertical: 'top', wrapText: true };
    cell.border = {
      top: { style: 'hair', color: { argb: 'FFCCE0F5' } },
      bottom: { style: 'hair', color: { argb: 'FFCCE0F5' } },
      left: { style: 'thin', color: { argb: 'FFCCE0F5' } },
      right: { style: 'thin', color: { argb: 'FFCCE0F5' } },
    };
  }
}

// Calculate row height based on the number of newlines in text-heavy cells
function calcRowHeight(r, cols) {
  let maxLines = 1;
  cols.forEach(c => {
    if (TEXT_WRAP_COLS.has(c.key)) {
      const val = r[c.key] || '';
      maxLines = Math.max(maxLines, (val.match(/\n/g) || []).length + 1);
    }
  });
  return maxLines > 2 ? Math.min(maxLines * 14, 409) : undefined; // 409pt = Excel max
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
  const linkIdx    = cols.findIndex(c => c.key === 'View Link');
  const addlIdx    = cols.findIndex(c => c.key === 'Additional Details');
  const bidIdx     = cols.findIndex(c => c.key === 'Bid Document Details');
  const companyIdx = cols.findIndex(c => c.key === 'Company');
  const statusIdx  = cols.findIndex(c => c.key === 'Status');

  rows.forEach((r, i) => {
    const values = cols.map(c => r[c.key] ?? '');
    const row    = ws.addRow(values);
    dataRowStyle(row, i + 1, cols.length);

    const h = calcRowHeight(r, cols);
    if (h) row.height = h;

    // Company dropdown — Gravity / Total Tech / Quickman
    if (companyIdx >= 0) {
      row.getCell(companyIdx + 1).dataValidation = {
        type: 'list', allowBlank: true,
        formulae: ['"Gravity,Total Tech,Quickman"'],
        showErrorMessage: false,
      };
    }

    // Status dropdown — Important / Filled
    if (statusIdx >= 0) {
      row.getCell(statusIdx + 1).dataValidation = {
        type: 'list', allowBlank: true,
        formulae: ['"Important,Filled"'],
        showErrorMessage: false,
      };
    }

    // Hyperlink on View Link column
    if (linkIdx >= 0) {
      const cell = row.getCell(linkIdx + 1);
      const href = r['View Link'];
      if (href && href.startsWith('http')) {
        cell.value = { text: 'View Tender', hyperlink: href };
        cell.font = { name: 'Calibri', size: 10, color: { argb: 'FF1565C0' }, underline: true };
      }
    }

    // Additional Details — compact Calibri, left-aligned, subtle colour
    if (addlIdx >= 0) {
      const cell = row.getCell(addlIdx + 1);
      cell.font      = { name: 'Calibri', size: 9, color: { argb: 'FF1E3A5F' } };
      cell.alignment = { vertical: 'top', wrapText: true, horizontal: 'left' };
    }

    // Bid Document Details — monospace so pipe-delimited columns line up
    if (bidIdx >= 0) {
      const cell = row.getCell(bidIdx + 1);
      cell.font      = { name: 'Consolas', size: 9 };
      cell.alignment = { vertical: 'top', wrapText: true, horizontal: 'left' };
    }
  });

  ws.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1 + rows.length, column: cols.length },
  };
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
    return raw.replace(/\(\d+\)\s*$/, '').replace(/[\\/:?*[\]]/g, '').trim();
  }
  function uniqueName(raw) {
    let base = raw.replace(/[\\/:?*[\]]/g, '_').trim().slice(0, 31);
    let attempt = base;
    let n = 2;
    while (usedNames.has(attempt.toLowerCase())) attempt = base.slice(0, 28) + '_' + (n++);
    usedNames.add(attempt.toLowerCase());
    return attempt;
  }

  // Helper: Excel column letter (1→A, 26→Z, 27→AA, …)
  function colLetter(n) {
    let r = '';
    while (n > 0) { n--; r = String.fromCharCode(65 + n % 26) + r; n = Math.floor(n / 26); }
    return r;
  }

  // ── Master "All Sections" sheet — Company, Status, Section + fixed cols, no Bid Doc Details
  const masterCols = [
    ...COLS.filter(c => c.key === 'Company' || c.key === 'Status'),
    { key: 'Section', label: 'Section', width: 28 },
    ...COLS.filter(c => !['Company', 'Status', 'Bid Document Details'].includes(c.key)),
  ];
  const allRows  = sections.flatMap(s => s.tenders.map(t => ({ Section: s.section, ...t })));
  const lastCol  = colLetter(masterCols.length);
  const maxRow   = Math.max(allRows.length + 10, 1000);
  const allWs    = wb.addWorksheet(uniqueName('All Sections'));
  fillDataSheet(allWs, masterCols, allRows, 'FF2C3E50', '2C3E50');

  // ── Helper: build a standard linked filter sheet (FILTER formula)
  function makeFilterSheet(name, tabArgb, hdrHex, filterCol, filterValue) {
    const ws = wb.addWorksheet(uniqueName(name));
    ws.properties = { tabColor: { argb: tabArgb } };
    ws.views      = [{ state: 'frozen', ySplit: 1, showGridLines: true }];
    ws.columns    = masterCols.map(c => ({ header: c.label, key: c.key, width: c.width }));
    const hdr     = ws.getRow(1);
    hdr.values    = masterCols.map(c => c.label);
    headerStyle(hdr, hdrHex, masterCols.length);
    const src  = `'All Sections'!A2:${lastCol}${maxRow}`;
    const crit = `'All Sections'!${filterCol}2:${filterCol}${maxRow}`;
    ws.getCell('A2').value = {
      formula: `IFERROR(FILTER(${src},${crit}="${filterValue}"),IF(ROWS(A2:A2)=1,"No tenders assigned to ${name}",""))`,
    };
    return ws;
  }

  // ── Sheet order: All Sections → Important → Filled → Gravity/TT/QM → per-section → Corrigendum

  // 1. Important
  makeFilterSheet('Important', 'FFB71C1C', 'B71C1C', 'B', 'Important');

  // 2. Filled — same filter base, plus three extra input columns
  const filledWs = makeFilterSheet('Filled', 'FF00695C', '00695C', 'B', 'Filled');
  const filledExtraCols = [
    { letter: colLetter(masterCols.length + 1), label: 'Filled Date', width: 16 },
    { letter: colLetter(masterCols.length + 2), label: 'Filled By',   width: 18 },
    { letter: colLetter(masterCols.length + 3), label: 'Bid Status',  width: 14 },
  ];
  filledExtraCols.forEach(({ letter, label, width }) => {
    filledWs.getColumn(letter).width = width;
    const cell      = filledWs.getCell(`${letter}1`);
    cell.value      = label;
    cell.fill       = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00695C' } };
    cell.font       = { name: 'Calibri', size: 11, bold: true, italic: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment  = { vertical: 'middle', horizontal: 'center' };
    cell.border     = {
      top: { style: 'medium', color: { argb: 'FF00695C' } },
      bottom: { style: 'medium', color: { argb: 'FF00695C' } },
      left: { style: 'thin', color: { argb: 'FFB0C4DE' } },
      right: { style: 'thin', color: { argb: 'FFB0C4DE' } },
    };
  });
  // Bid Status column: dropdown + conditional formatting (green = Accepted, red = Rejected)
  const bidStatusCol = colLetter(masterCols.length + 3);
  for (let r = 2; r <= 1002; r++) {
    filledWs.getCell(`${bidStatusCol}${r}`).dataValidation = {
      type: 'list', allowBlank: true,
      formulae: ['"Accepted,Rejected"'],
      showErrorMessage: false,
    };
  }
  filledWs.addConditionalFormatting({
    ref: `${bidStatusCol}2:${bidStatusCol}1002`,
    rules: [
      {
        type: 'containsText', operator: 'containsText', text: 'Accepted', priority: 1,
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D32' } },
          font: { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } },
        },
      },
      {
        type: 'containsText', operator: 'containsText', text: 'Rejected', priority: 2,
        style: {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC62828' } },
          font: { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } },
        },
      },
    ],
  });

  // 3. Company filter sheets
  makeFilterSheet('Gravity',    'FF1B5E20', '1B5E20', 'A', 'Gravity');
  makeFilterSheet('Total Tech', 'FF1565C0', '1565C0', 'A', 'Total Tech');
  makeFilterSheet('Quickman',   'FF4A148C', '4A148C', 'A', 'Quickman');

  // 4. Per-section sheets — full COLS including Bid Document Details
  sections.forEach((sec, idx) => {
    const tabArgb   = TAB_COLORS[idx % TAB_COLORS.length];
    const headerHex = tabArgb.slice(2);
    const ws        = wb.addWorksheet(uniqueName(sectionBaseName(sec.section)));
    fillDataSheet(ws, COLS, sec.tenders, tabArgb, headerHex);
  });

  // 5. Corrigendum — only if any tender contains "corrigendum" in any field
  const corrigendumRows = allRows.filter(r =>
    Object.values(r).some(v => typeof v === 'string' && v.toLowerCase().includes('corrigendum'))
  );
  if (corrigendumRows.length) {
    const ws = wb.addWorksheet('Corrigendum');
    fillDataSheet(ws, masterCols, corrigendumRows, 'FFFF6F00', 'E65100');
  }

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
//  PHASE 2 — Bid document parsers
//  Both return a plain string (newline-separated rows / pairs).
//  All tabular and numerical data (manpower, costs, quantities,
//  BOQ rows, etc.) is captured as-is and stored in one cell.
// ══════════════════════════════════════════════════════════════

async function parsePdfBidDocument(url) {
  let parser;
  try {
    const { data } = await axios.get(url, { responseType: 'arraybuffer', timeout: 20000, maxRedirects: 4 });
    parser = new PDFParse({ data: new Uint8Array(data) });

    const [tableResult, textResult] = await Promise.all([
      parser.getTable().catch(() => ({ pages: [] })),
      parser.getText(),
    ]);

    const tableCellTexts = new Set();
    const tableBlocks    = [];

    for (const page of tableResult.pages) {
      for (const table of (page.tables || [])) {
        if (!table?.length) continue;
        const rows = [];
        table.forEach((row, rowIdx) => {
          const cells = row.map(c => (c || '').trim());
          if (!cells.some(Boolean)) return;
          cells.forEach(c => { if (c) tableCellTexts.add(c); });
          const line = cells.join('  |  ');
          rows.push(line);
          if (rowIdx === 0) rows.push('─'.repeat(Math.min(line.length, 60)));
        });
        if (rows.length) tableBlocks.push(rows.join('\n'));
      }
    }

    const seen = new Set(tableCellTexts);
    const textLines = [];
    let blankRun = 0;

    for (const rawLine of textResult.text.split('\n')) {
      const t = rawLine.trim();
      if (!t || /^\d{1,3}$/.test(t) || /^Page\s+\d+/i.test(t) || seen.has(t)) {
        if (textLines.length && blankRun < 1) { textLines.push(''); blankRun++; }
        continue;
      }
      blankRun = 0;
      seen.add(t);
      textLines.push(rawLine.trimEnd());
    }

    const sections = [];
    if (tableBlocks.length) sections.push(tableBlocks.join('\n\n'));
    const plainText = textLines.join('\n').trim();
    if (plainText) sections.push(plainText);

    return cleanBidText(sections.join('\n\n'));
  } catch { return ''; } finally {
    if (parser) await parser.destroy().catch(() => {});
  }
}

async function parseHtmlBidDocument(url) {
  try {
    const { data } = await axios.get(url, { headers: HEADERS, timeout: 15000, maxRedirects: 4 });
    const $        = cheerio.load(data);
    $('script, style, nav, footer, header').remove();

    const bodyText = $('body').text().replace(/\s{3,}/g, '\n').trim();
    const sections = [];
    const seen     = new Set();

    const addUniq = (arr, line) => {
      const l = line.trimEnd();
      if (!l || l.length < 2 || seen.has(l)) return false;
      seen.add(l); arr.push(l); return true;
    };

    $('table').each((_, table) => {
      const $tbl  = $(table);
      const cap   = clean($tbl.find('caption').first().text());
      const prev  = clean($tbl.prevAll('h1,h2,h3,h4,h5,h6,p.heading,.section-title').first().text());
      const title = cap || prev;
      const tableLines = [];
      $tbl.find('tr').each((rowIdx, tr) => {
        const cells = [];
        $(tr).find('th, td').each((_, td) => cells.push(clean($(td).text())));
        if (cells.filter(Boolean).length >= 2) {
          const line = cells.join('  |  ');
          addUniq(tableLines, line);
          if (rowIdx === 0 && tableLines.length) tableLines.push('─'.repeat(Math.min(line.length, 60)));
        }
      });
      if (tableLines.length) {
        const block = [];
        if (title) block.push(`▌ ${title}`);
        block.push(...tableLines);
        sections.push(block.join('\n'));
      }
    });

    $('dl').each((_, dl) => {
      const dlLines = [];
      $(dl).find('dt').each((_, dt) => {
        const key = clean($(dt).text());
        const val = clean($(dt).next('dd').text());
        if (key && val) addUniq(dlLines, `${key}: ${val}`);
      });
      if (dlLines.length) sections.push(dlLines.join('\n'));
    });

    if (!sections.length) {
      const lines    = bodyText.split('\n').map(l => l.trim()).filter(Boolean);
      const fallback = [];
      for (let i = 0; i < lines.length - 1; i++) {
        const k = lines[i], v = lines[i + 1];
        if (k.length > 3 && k.length < 100 && v && !/^[\d\.]+$/.test(k)) {
          addUniq(fallback, `${k}: ${v}`); i++;
        }
      }
      if (fallback.length) sections.push(fallback.join('\n'));
    }

    return cleanBidText(sections.join('\n\n'));
  } catch { return ''; }
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

async function scrapeTenderDetail(viewLink) {
  try {
    const { data } = await axios.get(viewLink, { headers: HEADERS, timeout: 15000, maxRedirects: 4 });
    const $ = cheerio.load(data);
    const record = {};

    $('table tr').each((_, row) => {
      const $tds = $(row).find('td');
      if ($tds.length >= 2) {
        const label = clean($tds.eq(0).text());
        const value = clean($tds.eq(1).text());
        if (label && value && label.length < 80 && !label.startsWith('Download')) record[label] = value;
      }
    });

    // Collect every detail-page field that is not already in a dedicated column
    const additionalLines = [];
    for (const [k, v] of Object.entries(record)) {
      if (!MAPPED_KEYS.has(k)) additionalLines.push(`${k}: ${v}`);
    }

    // Locate bid document link (HTML preferred over PDF)
    let htmlLink = null, pdfLink = null;
    $('a').each((_, a) => {
      const href = $(a).attr('href') || '';
      if (href.includes('tenderfiles.com')) {
        if (!htmlLink && href.endsWith('.html')) htmlLink = href;
        if (!pdfLink && href.endsWith('.pdf')) pdfLink = href;
      }
    });

    let bidDocDetails = '';
    if (htmlLink) bidDocDetails = await parseHtmlBidDocument(htmlLink);
    else if (pdfLink) bidDocDetails = await parsePdfBidDocument(pdfLink);

    return {
      'Company': '',
      'Status':  '',
      'TDR':                         record['TDR']                                   || 'N/A',
      'Tender No':                   record['Tender No'] || record['Tender ID']      || 'N/A',
      'Tendering Authority':         record['Tendering Authority'] || record['Company Name'] || 'N/A',
      'Tender Brief':                record['Tender Brief']                          || 'N/A',
      'City':                        record['City']                                  || 'N/A',
      'State':                       record['State']                                 || 'N/A',
      'Document Fees':               record['Document Fees']                         || 'N/A',
      'EMD':                         normalizeAmount(record['EMD']),
      'Tender Value':                normalizeAmount(record['Tender Value']),
      'Tender Type':                 record['Tender Type']                           || 'N/A',
      'Bidding Type':                record['Bidding Type']                          || 'N/A',
      'Competition Type':            record['Competition Type']                      || 'N/A',
      'Publish Date':                record['Publish Date']                          || 'N/A',
      'Last Date of Bid Submission': record['Last Date of Bid Submission']           || 'N/A',
      'Tender Opening Date':         record['Tender Opening Date']                   || 'N/A',
      'Address':                     record['Address']                               || 'N/A',
      'Information Source':          record['Information Source']                    || 'N/A',
      'View Link':                   viewLink,
      'Additional Details':          additionalLines.join('\n')                      || 'N/A',
      'Bid Document Details':        bidDocDetails                                   || 'N/A',
    };
  } catch (e) {
    return {
      'Company': '', 'Status': '',
      'TDR': 'N/A', 'Tender No': 'N/A', 'Tendering Authority': 'N/A',
      'Tender Brief': `Error: ${e.message}`, 'City': 'N/A', 'State': 'N/A',
      'Document Fees': 'N/A', 'EMD': 'N/A', 'Tender Value': 'N/A',
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
}

// ══════════════════════════════════════════════════════════════
//  GET /scrape-deep — SSE: fetch all detail pages → build Excel
// ══════════════════════════════════════════════════════════════
app.get('/scrape-deep', async (req, res) => {
  const { url } = req.query;
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
        (t) => scrapeTenderDetail(t.viewLink),
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

    sseWrite(res, 'status', { phase: 3, message: 'Building formatted Excel file…' });
    const excelBuf = await buildExcel(enriched);

    const token = crypto.randomBytes(16).toString('hex');
    const safeName = pageTitle.replace(/[^a-z0-9]/gi, '_').slice(0, 40) || 'tender_data';
    const filePath = path.join(TMP_DIR, `${token}.xlsx`);
    fs.writeFileSync(filePath, excelBuf);
    setTimeout(() => { try { fs.unlinkSync(filePath); } catch (_) { } }, 30 * 60 * 1000);

    sseWrite(res, 'done', { token, filename: safeName, totalTenders: total, sections: sections.length, pageTitle });
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
