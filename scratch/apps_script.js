/**
 * Gravity Scraper — Google Apps Script Bridge
 * ─────────────────────────────────────────────
 * PROTECTION MODEL
 *   All Sections        → cols A (Company) and B (Important) editable; cols C+ locked
 *   Gravity/TT/Quickman → cols A (Filled), B (Filled Date), C (Filled By) editable; cols D+ locked
 *   All other sheets    → header row (row 1) free so AutoFilter works for everyone;
 *                         data rows (row 2+) fully locked
 *
 * CHIP DROPDOWNS & CHECKBOXES (Google Sheets native)
 *   All Sections → Company (col A) chip + multi-select panel
 *                  Important (col B) BOOLEAN checkbox
 *   Company sheets → Filled (col A) BOOLEAN checkbox
 *
 * WHY RANGE PROTECTION INSTEAD OF SHEET PROTECTION
 *   Sheet-level protection blocks the AutoFilter arrows for non-owners.
 *   Range-level protection (startRowIndex 1+) leaves the header row free, so
 *   everyone with the link can click filter arrows and filter rows without
 *   being able to edit the protected cell values.
 */

var COMPANY_SHEETS = ['Gravity', 'Total Tech', 'Quickman'];

// ── Chip dropdown definitions for All Sections ────────────────────────────────
// showCustomUi:true → chip pill UI; clicking opens a multi-select checkbox panel
// Important (col B) is a BOOLEAN checkbox — handled separately below, not here.
var ALL_SECTIONS_CHIPS = [
  {
    // Company — col A (index 0) — multi-select chip
    col: 0,
    values: ['Gravity', 'Total Tech', 'Quickman']
  }
];

// ─────────────────────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var blob = Utilities.newBlob(
      Utilities.base64Decode(data.base64),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      data.filename
    );

    // 1 ── Upload XLSX and convert to a native Google Sheet
    var file = Drive.Files.create(
      { name: data.filename, mimeType: 'application/vnd.google-apps.spreadsheet' },
      blob
    );
    var fileId = file.id;

    // 2 ── Grant Writer access to anyone with the link.
    //      Writer (not just Viewer) is required so users can click AutoFilter arrows.
    //      The range protections below prevent them from editing the locked cells.
    Drive.Permissions.create({ role: 'writer', type: 'anyone' }, fileId);

    // 3 ── Fetch all sheet metadata
    var ss = Sheets.Spreadsheets.get(fileId);
    var sheets = ss.sheets;
    var owner = Session.getEffectiveUser().getEmail();

    var requests = [];

    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var title = sheet.properties.title;
      var sheetId = sheet.properties.sheetId;
      var grid = sheet.properties.gridProperties || {};
      var rows = grid.rowCount || 1000;
      var cols = grid.columnCount || 30;

      // ── STEP A  Remove any protection that was imported from the XLSX ────────
      // Google Sheets may import XLSX sheet-level protection. We remove it so our
      // finer-grained range protection takes effect correctly.
      var existingProtections = sheet.protectedRanges || [];
      for (var p = 0; p < existingProtections.length; p++) {
        requests.push({
          deleteProtectedRange: {
            protectedRangeId: existingProtections[p].protectedRangeId
          }
        });
      }

      // ── STEP B  Apply new range-based protection ─────────────────────────────

      if (title === 'All Sections') {
        // Lock cols C+ (index 2+); leave Company (A=0) and Status (B=1) free
        requests.push({
          addProtectedRange: {
            protectedRange: {
              range: {
                sheetId: sheetId,
                startRowIndex: 1,      // skip header row
                endRowIndex: rows,
                startColumnIndex: 2,      // col C onwards
                endColumnIndex: cols
              },
              description: 'Scraped data — read-only. Only Company (A) and Important (B) are editable.',
              editors: { users: [owner] }
            }
          }
        });

      } else if (COMPANY_SHEETS.indexOf(title) !== -1) {
        // Company sheets: lock cols D+ (index 3+); leave Filled (A=0), Filled Date (B=1), Filled By (C=2) free
        requests.push({
          addProtectedRange: {
            protectedRange: {
              range: {
                sheetId: sheetId,
                startRowIndex: 1,
                endRowIndex: rows,
                startColumnIndex: 3,      // col D onwards (formula data starts here)
                endColumnIndex: cols
              },
              description: 'Formula-driven data — read-only. Filled (A), Filled Date (B) and Filled By (C) are editable.',
              editors: { users: [owner] }
            }
          }
        });

        // Filled (col A) — BOOLEAN checkbox
        requests.push({
          setDataValidation: {
            range: {
              sheetId: sheetId,
              startRowIndex: 1,
              endRowIndex: rows,
              startColumnIndex: 0,    // col A
              endColumnIndex: 1
            },
            rule: {
              condition: { type: 'BOOLEAN' },
              showCustomUi: true
            }
          }
        });

      } else {
        // All other sheets (Important, Filled, Corrigendum, section tabs):
        // Lock data rows 2+ entirely. Header row (index 0) stays free so AutoFilter arrows work.
        requests.push({
          addProtectedRange: {
            protectedRange: {
              range: {
                sheetId: sheetId,
                startRowIndex: 1,      // row 2 onwards
                endRowIndex: rows,
                startColumnIndex: 0,
                endColumnIndex: cols
              },
              description: 'Scraped data — read-only. Use the filter arrows in the header to filter.',
              editors: { users: [owner] }
            }
          }
        });
      }

      // ── STEP C  Chip dropdowns + checkboxes on All Sections ─────────────────
      if (title === 'All Sections') {
        // Company (col A) — chip pill with multi-select checkbox panel
        for (var c = 0; c < ALL_SECTIONS_CHIPS.length; c++) {
          var chip = ALL_SECTIONS_CHIPS[c];
          requests.push({
            setDataValidation: {
              range: {
                sheetId: sheetId,
                startRowIndex: 1,      // skip header
                endRowIndex: rows,
                startColumnIndex: chip.col,
                endColumnIndex: chip.col + 1
              },
              rule: {
                condition: {
                  type: 'ONE_OF_LIST',
                  values: chip.values.map(function (v) { return { userEnteredValue: v }; })
                },
                showCustomUi: true,   // ← chip pill + multi-select checkbox panel
                strict: false         // ← allow typed comma-separated values too
              }
            }
          });
        }

        // Important (col B) — BOOLEAN checkbox
        requests.push({
          setDataValidation: {
            range: {
              sheetId: sheetId,
              startRowIndex: 1,
              endRowIndex: rows,
              startColumnIndex: 1,    // col B
              endColumnIndex: 2
            },
            rule: {
              condition: { type: 'BOOLEAN' },
              showCustomUi: true
            }
          }
        });
      }
    } // end for sheets

    // ── STEP D  Execute all requests in a single batch ───────────────────────
    if (requests.length > 0) {
      Sheets.Spreadsheets.batchUpdate({ requests: requests }, fileId);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        url: 'https://docs.google.com/spreadsheets/d/' + fileId + '/edit'
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        error: err.toString(),
        details: 'Ensure Google Sheets API and Drive API are enabled under Services → Add a service.'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Run once in the Apps Script IDE to grant OAuth permissions before deploying
function authorizeMe() {
  DriveApp.getRootFolder();
  SpreadsheetApp.getActiveSpreadsheet();
}
