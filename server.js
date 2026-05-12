const express  = require('express');
const axios    = require('axios');
const cheerio  = require('cheerio');
const ExcelJS  = require('exceljs');
const cors     = require('cors');
const path     = require('path');
const fs       = require('fs');
const crypto   = require('crypto');

const app  = express();
const PORT = 3000;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const TMP_DIR = path.join(__dirname, 'tmp');
if (!fs.existsSync(TMP_DIR)) fs.mkdirSync(TMP_DIR);
cleanOldFiles();

const HEADERS = {
  'User-Agent':      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
  'Accept':          'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
  'Accept-Language': 'en-US,en;q=0.5',
};
const CONCURRENCY = 8;

function clean(t) { return (t || '').replace(/\s+/g, ' ').trim(); }

// ══════════════════════════════════════════════════════════════
//  EXCEL SCHEMA — columns in order
// ══════════════════════════════════════════════════════════════
const COLS = [
  { key: 'TDR',                        label: 'TDR',                         width: 14 },
  { key: 'Tender No',                  label: 'Tender No',                   width: 26 },
  { key: 'Tendering Authority',        label: 'Tendering Authority',         width: 36 },
  { key: 'Tender Brief',              label: 'Tender Brief',                width: 62 },
  { key: 'City',                       label: 'City',                        width: 16 },
  { key: 'State',                      label: 'State',                       width: 18 },
  { key: 'Document Fees',             label: 'Document Fees',               width: 16 },
  { key: 'EMD',                        label: 'EMD',                         width: 22 },
  { key: 'Tender Value',              label: 'Tender Value',                width: 20 },
  { key: 'Tender Type',              label: 'Tender Type',                  width: 16 },
  { key: 'Bidding Type',             label: 'Bidding Type',                 width: 16 },
  { key: 'Competition Type',         label: 'Competition Type',             width: 18 },
  { key: 'Publish Date',             label: 'Publish Date',                 width: 14 },
  { key: 'Last Date of Bid Submission', label: 'Last Date of Bid Submission', width: 26 },
  { key: 'Tender Opening Date',      label: 'Tender Opening Date',          width: 22 },
  { key: 'Address',                   label: 'Address',                     width: 34 },
  { key: 'Information Source',       label: 'Information Source',           width: 30 },
  { key: 'View Link',                 label: 'View Link',                   width: 60 },
];

// ── Per-section tab colours (ARGB, fully opaque)
const TAB_COLORS = [
  'FF1565C0', // Blue
  'FF00695C', // Teal
  'FFB71C1C', // Red
  'FFE65100', // Deep Orange
  'FF1B5E20', // Dark Green
  'FF4E342E', // Brown
  'FF4A148C', // Deep Purple
  'FF006064', // Cyan
  'FF1A237E', // Indigo
  'FF37474F', // Blue Grey
];

// ══════════════════════════════════════════════════════════════
//  STYLE HELPERS
// ══════════════════════════════════════════════════════════════
function headerStyle(row, colorHex, colCount) {
  row.height = 26;
  for (let c = 1; c <= colCount; c++) {
    const cell = row.getCell(c);
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + colorHex } };
    cell.font      = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: false };
    cell.border    = {
      top:    { style: 'medium', color: { argb: 'FF' + colorHex } },
      bottom: { style: 'medium', color: { argb: 'FF' + colorHex } },
      left:   { style: 'thin',   color: { argb: 'FFB0C4DE' } },
      right:  { style: 'thin',   color: { argb: 'FFB0C4DE' } },
    };
  }
}

function dataRowStyle(row, rowIdx, colCount) {
  const bg = rowIdx % 2 === 0 ? 'FFEAF4FB' : 'FFFFFFFF';
  for (let c = 1; c <= colCount; c++) {
    const cell = row.getCell(c);
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    cell.font      = { name: 'Calibri', size: 10 };
    cell.alignment = { vertical: 'top', wrapText: true };
    cell.border    = {
      top:    { style: 'hair',   color: { argb: 'FFCCE0F5' } },
      bottom: { style: 'hair',   color: { argb: 'FFCCE0F5' } },
      left:   { style: 'thin',   color: { argb: 'FFCCE0F5' } },
      right:  { style: 'thin',   color: { argb: 'FFCCE0F5' } },
    };
  }
}

// ══════════════════════════════════════════════════════════════
//  DATA SHEET — coloured headers, zebra rows, frozen row, filter
// ══════════════════════════════════════════════════════════════
function fillDataSheet(ws, cols, rows, tabArgb, headerColor) {
  ws.properties = { tabColor: { argb: tabArgb } };
  ws.views       = [{ state: 'frozen', ySplit: 1, showGridLines: true }];

  // Columns
  ws.columns = cols.map(c => ({ header: c.label, key: c.key, width: c.width }));

  // Header row
  const hdr = ws.getRow(1);
  hdr.values = cols.map(c => c.label);
  headerStyle(hdr, headerColor, cols.length);

  // Data rows
  rows.forEach((r, i) => {
    const values = cols.map(c => r[c.key] || 'N/A');
    const row    = ws.addRow(values);
    dataRowStyle(row, i + 1, cols.length);

    // Make View Link a clickable hyperlink
    const linkIdx = cols.findIndex(c => c.key === 'View Link');
    if (linkIdx >= 0) {
      const cell = row.getCell(linkIdx + 1);
      const href = r['View Link'];
      if (href && href.startsWith('http')) {
        cell.value = { text: 'View Tender', hyperlink: href };
        cell.font  = { name: 'Calibri', size: 10, color: { argb: 'FF1565C0' }, underline: true };
      }
    }
  });

  // AutoFilter
  ws.autoFilter = {
    from: { row: 1, column: 1 },
    to:   { row: 1 + rows.length, column: cols.length },
  };
}

// ══════════════════════════════════════════════════════════════
//  SORT SHEET — instruction banner + data with AutoFilter
//  The user clicks any column header ▼ to sort the whole sheet.
// ══════════════════════════════════════════════════════════════
function fillSortSheet(ws, cols, rows, tabArgb, headerColor) {
  ws.properties = { tabColor: { argb: tabArgb } };

  // Row 1 — instruction banner (merged across all columns + serial col)
  const totalCols = cols.length + 1; // +1 for #
  ws.mergeCells(1, 1, 1, totalCols);
  const banner       = ws.getCell('A1');
  banner.value       = '💡  Click the ▼ dropdown arrow on any column header to sort all data by that column  •  Use the filter dropdown to show only matching rows';
  banner.font        = { name: 'Calibri', size: 11, italic: true, color: { argb: 'FF1A3A5F' } };
  banner.fill        = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD6E8F7' } };
  banner.alignment   = { horizontal: 'center', vertical: 'middle' };
  banner.border      = { bottom: { style: 'medium', color: { argb: 'FFB0C4DE' } } };
  ws.getRow(1).height = 24;

  // Row 2 — headers (with serial # column prepended)
  ws.getColumn(1).width = 6;
  cols.forEach((c, i) => { ws.getColumn(i + 2).width = c.width; });

  const hdr    = ws.getRow(2);
  hdr.values   = ['#', ...cols.map(c => c.label)];
  headerStyle(hdr, headerColor, totalCols);

  // Freeze rows 1+2 so instructions and headers stay visible
  ws.views = [{ state: 'frozen', ySplit: 2, showGridLines: true }];

  // Data rows
  rows.forEach((r, i) => {
    const values = [i + 1, ...cols.map(c => r[c.key] || 'N/A')];
    const row    = ws.addRow(values);
    dataRowStyle(row, i + 1, totalCols);

    // Serial number: grey, centred
    const numCell      = row.getCell(1);
    numCell.font       = { name: 'Calibri', size: 10, color: { argb: 'FF8A9BBF' } };
    numCell.alignment  = { horizontal: 'center', vertical: 'top' };

    // Hyperlink on View Link column
    const linkIdx = cols.findIndex(c => c.key === 'View Link');
    if (linkIdx >= 0) {
      const cell = row.getCell(linkIdx + 2);
      const href = r['View Link'];
      if (href && href.startsWith('http')) {
        cell.value = { text: 'View Tender', hyperlink: href };
        cell.font  = { name: 'Calibri', size: 10, color: { argb: 'FF1565C0' }, underline: true };
      }
    }
  });

  // AutoFilter on row 2
  ws.autoFilter = {
    from: { row: 2, column: 1 },
    to:   { row: 2 + rows.length, column: totalCols },
  };
}

// ══════════════════════════════════════════════════════════════
//  BUILD WORKBOOK
// ══════════════════════════════════════════════════════════════
async function buildExcel(sections, pageTitle) {
  const wb      = new ExcelJS.Workbook();
  wb.creator    = 'Gravity Scraper';
  wb.company    = 'Gravity';
  wb.created    = new Date();
  wb.modified   = new Date();

  const usedNames = new Set();
  // Strip section count "(149)" from name, sanitize, enforce 31-char limit
  function sectionBaseName(raw) {
    return raw.replace(/\(\d+\)\s*$/, '').replace(/[\\/:?*[\]]/g, '').trim();
  }
  function uniqueName(raw) {
    let base    = raw.replace(/[\\/:?*[\]]/g, '_').trim().slice(0, 31);
    let attempt = base;
    let n       = 2;
    while (usedNames.has(attempt.toLowerCase())) attempt = base.slice(0, 28) + '_' + (n++);
    usedNames.add(attempt.toLowerCase());
    return attempt;
  }

  // ── 1. Combined "All Sections" sheet ────────────────────────
  const allWs   = wb.addWorksheet(uniqueName('All Sections'));
  const allCols = [{ key: 'Section', label: 'Section', width: 28 }, ...COLS];
  const allRows = sections.flatMap(s => s.tenders.map(t => ({ Section: s.section, ...t })));
  fillDataSheet(allWs, allCols, allRows, 'FF2C3E50', '2C3E50');

  // ── 2. Per-section: data sheet + sort sheet ──────────────────
  sections.forEach((sec, idx) => {
    const tabArgb   = TAB_COLORS[idx % TAB_COLORS.length];
    const headerHex = tabArgb.slice(2);
    const shortName = sectionBaseName(sec.section); // e.g. "Facility Management"

    // Data sheet  — named "Facility Management"
    const dataWs = wb.addWorksheet(uniqueName(shortName));
    fillDataSheet(dataWs, COLS, sec.tenders, tabArgb, headerHex);
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

    const $row2  = $el.children('div.row').eq(1);
    const $desc  = $el.children('div.row').eq(0).find('.col-md-12').first();
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

    return {
      'TDR':                         record['TDR']                             || 'N/A',
      'Tender No':                   record['Tender No'] || record['Tender ID'] || 'N/A',
      'Tendering Authority':         record['Tendering Authority'] || record['Company Name'] || 'N/A',
      'Tender Brief':                record['Tender Brief']                    || 'N/A',
      'City':                        record['City']                            || 'N/A',
      'State':                       record['State']                           || 'N/A',
      'Document Fees':               record['Document Fees']                   || 'N/A',
      'EMD':                         record['EMD']                             || 'N/A',
      'Tender Value':                record['Tender Value']                    || 'N/A',
      'Tender Type':                 record['Tender Type']                     || 'N/A',
      'Bidding Type':                record['Bidding Type']                    || 'N/A',
      'Competition Type':            record['Competition Type']                || 'N/A',
      'Publish Date':                record['Publish Date']                    || 'N/A',
      'Last Date of Bid Submission': record['Last Date of Bid Submission']     || 'N/A',
      'Tender Opening Date':         record['Tender Opening Date']             || 'N/A',
      'Address':                     record['Address']                         || 'N/A',
      'Information Source':          record['Information Source']              || 'N/A',
      'View Link':                   viewLink,
    };
  } catch (e) {
    return {
      'TDR': 'N/A', 'Tender No': 'N/A', 'Tendering Authority': 'N/A',
      'Tender Brief': `Error: ${e.message}`, 'City': 'N/A', 'State': 'N/A',
      'Document Fees': 'N/A', 'EMD': 'N/A', 'Tender Value': 'N/A',
      'Tender Type': 'N/A', 'Bidding Type': 'N/A', 'Competition Type': 'N/A',
      'Publish Date': 'N/A', 'Last Date of Bid Submission': 'N/A',
      'Tender Opening Date': 'N/A', 'Address': 'N/A', 'Information Source': 'N/A',
      'View Link': viewLink,
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
  if (!url) { res.status(400).end(); return; }

  res.setHeader('Content-Type',  'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection',    'keep-alive');
  res.flushHeaders();

  try {
    sseWrite(res, 'status', { phase: 1, message: 'Fetching listing page…' });

    const { data } = await axios.get(url.split('#')[0], { headers: HEADERS, timeout: 20000, maxRedirects: 5 });
    const $         = cheerio.load(data);
    const pageTitle = clean($('title').text()) || 'Tenders';
    $('script, style').remove();

    const sections = parseDailyDigest($);
    const total    = sections.reduce((n, s) => n + s.tenders.length, 0);

    sseWrite(res, 'status', {
      phase: 2,
      message: `Found ${total} tenders across ${sections.length} sections. Visiting detail pages…`,
      total,
      sections: sections.map(s => ({ section: s.section, count: s.tenders.length })),
    });

    let globalDone    = 0;
    const enriched    = [];

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

    // Build styled Excel
    sseWrite(res, 'status', { phase: 3, message: 'Building formatted Excel file…' });
    const excelBuf = await buildExcel(enriched, pageTitle);

    const token    = crypto.randomBytes(16).toString('hex');
    const safeName = pageTitle.replace(/[^a-z0-9]/gi, '_').slice(0, 40) || 'tender_data';
    const filePath = path.join(TMP_DIR, `${token}.xlsx`);
    fs.writeFileSync(filePath, excelBuf);
    setTimeout(() => { try { fs.unlinkSync(filePath); } catch (_) {} }, 30 * 60 * 1000);

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
  const token    = req.params.token.replace(/[^a-f0-9]/gi, '');
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
      try { if (fs.statSync(fp).mtimeMs < cutoff) fs.unlinkSync(fp); } catch (_) {}
    });
  } catch (_) {}
}

app.listen(PORT, () => console.log(`\n🚀 Gravity Scraper → http://localhost:${PORT}\n`));
