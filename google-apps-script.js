// ============================================================
// Google Apps Script — Monthly Report API
// ============================================================
// INSTRUCTIONS:
// 1. Create a new Google Sheet
// 2. Go to Extensions → Apps Script
// 3. Delete any existing code and paste this entire file
// 4. Click Deploy → New deployment
//    - Type: Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copy the Web App URL and paste it into app.js & admin.js
// ============================================================

const SHEET_NAME = 'Reports';

const EXPECTED_HEADERS = [
    'ID',
    'Submitted Date',
    'Name',
    'Role',
    'Project',
    'Key Highlights',
    'Upcoming Focus',
    'Issues',
    'Concerns',
    'Risks',
    'Need Support'
];

const COLUMN_WIDTHS = [180, 160, 150, 220, 250, 400, 400, 400, 400, 400, 400];

/**
 * Get or create the sheet. Ensures schema matches EXPECTED_HEADERS.
 */
function getOrCreateSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME);
    }

    ensureSchema(sheet);
    return sheet;
}

/**
 * Write expected headers + formatting. Idempotent.
 */
function writeHeaders(sheet) {
    sheet.getRange(1, 1, 1, EXPECTED_HEADERS.length).setValues([EXPECTED_HEADERS]);
    const headerRange = sheet.getRange(1, 1, 1, EXPECTED_HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4f46e5');
    headerRange.setFontColor('#ffffff');
    COLUMN_WIDTHS.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    sheet.setFrozenRows(1);
}

/**
 * Ensure sheet schema matches EXPECTED_HEADERS.
 * - Empty sheet: write headers.
 * - Legacy 10-col sheet (no Role, Project at col 4): insert Role at col 4.
 * - Otherwise: overwrite header row (data is preserved).
 */
function ensureSchema(sheet) {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow === 0 || lastCol === 0) {
        writeHeaders(sheet);
        return;
    }

    const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // Already matches expected schema — nothing to do.
    const matches = currentHeaders.length >= EXPECTED_HEADERS.length &&
        EXPECTED_HEADERS.every((h, i) => currentHeaders[i] === h);
    if (matches) return;

    // Legacy 10-col schema (pre-Role): insert a blank column at position 4
    // so existing data rows shift correctly and Role lands in the right spot.
    const hasRole = currentHeaders.indexOf('Role') !== -1;
    if (!hasRole && currentHeaders[3] === 'Project') {
        sheet.insertColumnBefore(4);
    }

    writeHeaders(sheet);
}

/**
 * Manual utility — run this from the Apps Script editor once to fix a
 * corrupted header row. Rewrites headers and trims any columns beyond 11.
 * Data rows are NOT modified, so verify alignment afterward.
 */
function resetHeaders() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
        Logger.log('Sheet "' + SHEET_NAME + '" not found.');
        return;
    }
    const lastCol = sheet.getLastColumn();
    if (lastCol > EXPECTED_HEADERS.length) {
        sheet.deleteColumns(EXPECTED_HEADERS.length + 1, lastCol - EXPECTED_HEADERS.length);
    }
    writeHeaders(sheet);
    Logger.log('Headers reset. Verify data alignment in row 2+.');
}

/**
 * Manual utility — prints current sheet state to the execution log.
 * Run from the Apps Script editor to diagnose column alignment issues.
 */
function diagnoseSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
        Logger.log('Sheet "' + SHEET_NAME + '" not found.');
        return;
    }
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    Logger.log('Rows: ' + lastRow + ', Cols: ' + lastCol);
    if (lastCol > 0) {
        Logger.log('Headers: ' + JSON.stringify(sheet.getRange(1, 1, 1, lastCol).getValues()[0]));
    }
    if (lastRow >= 2) {
        Logger.log('Row 2: ' + JSON.stringify(sheet.getRange(2, 1, 1, lastCol).getValues()[0]));
    }
    if (lastRow >= 3) {
        Logger.log('Last row: ' + JSON.stringify(sheet.getRange(lastRow, 1, 1, lastCol).getValues()[0]));
    }
}

/**
 * Manual utility — DESTRUCTIVE. Wipes all data rows but keeps the header.
 * Use this to start fresh after schema issues. Run from the Apps Script editor.
 */
function clearAllData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
        Logger.log('Sheet "' + SHEET_NAME + '" not found.');
        return;
    }
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
    }
    ensureSchema(sheet);
    Logger.log('Cleared all data rows. Schema enforced.');
}

/**
 * Format array items as numbered list string
 */
function formatItems(items) {
    if (!items || items.length === 0) return '';
    if (items.length === 1) return items[0];
    return items.map((item, i) => `${i + 1}. ${item}`).join('\n');
}

/**
 * Parse numbered list string back to array
 */
function parseItems(str) {
    if (!str || str.trim() === '') return [];
    // Check if it's a numbered list
    const lines = str.split('\n');
    if (lines.length === 1 && !lines[0].match(/^\d+\.\s/)) {
        return [lines[0]];
    }
    return lines.map(line => line.replace(/^\d+\.\s*/, '').trim()).filter(l => l);
}

/**
 * Handle POST requests — save reports
 */
function doPost(e) {
    try {
        const data = JSON.parse(e.postData.contents);
        const action = data.action || 'submit';

        if (action === 'submit') {
            return handleSubmit(data);
        } else if (action === 'delete') {
            return handleDelete(data);
        } else if (action === 'clearMonth') {
            return handleClearMonth(data);
        }

        return jsonResponse({ success: false, error: 'Unknown action' });
    } catch (err) {
        return jsonResponse({ success: false, error: err.toString() });
    }
}

/**
 * Handle GET requests — fetch reports
 */
function doGet(e) {
    try {
        const sheet = getOrCreateSheet();
        const lastRow = sheet.getLastRow();

        if (lastRow <= 1) {
            return jsonResponse({ success: true, reports: [] });
        }

        const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
        const reports = data.map(row => ({
            id: row[0],
            submittedAt: row[1],
            name: row[2],
            role: row[3],
            project: row[4],
            keyHighlights: parseItems(row[5]),
            upcomingFocus: parseItems(row[6]),
            issues: parseItems(row[7]),
            concerns: parseItems(row[8]),
            risks: parseItems(row[9]),
            needSupport: parseItems(row[10])
        }));

        return jsonResponse({ success: true, reports: reports });
    } catch (err) {
        return jsonResponse({ success: false, error: err.toString() });
    }
}

/**
 * Submit a new report
 */
function handleSubmit(data) {
    const sheet = getOrCreateSheet();
    const id = data.id || new Date().getTime().toString();
    const submittedAt = data.submittedAt || new Date().toISOString();
    const name = data.name;
    const role = data.role || '';
    const projects = data.projects || [];

    const rows = [];
    projects.forEach(proj => {
        rows.push([
            id,
            submittedAt,
            name,
            role,
            proj.projectName || '',
            formatItems(proj.keyHighlights || []),
            formatItems(proj.upcomingFocus || []),
            formatItems(proj.issues || []),
            formatItems(proj.concerns || []),
            formatItems(proj.risks || []),
            formatItems(proj.needSupport || [])
        ]);
    });

    if (rows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 11).setValues(rows);
        // Enable text wrap for data columns
        const lastRow = sheet.getLastRow();
        sheet.getRange(lastRow - rows.length + 1, 6, rows.length, 6).setWrap(true);
    }

    return jsonResponse({
        success: true,
        message: `Saved ${rows.length} project(s)`,
        id: id
    });
}

/**
 * Delete a report by ID
 */
function handleDelete(data) {
    const sheet = getOrCreateSheet();
    const targetId = data.id;
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
        return jsonResponse({ success: true, message: 'No reports to delete' });
    }

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let deletedCount = 0;

    // Delete from bottom to top to preserve row indices
    for (let i = ids.length - 1; i >= 0; i--) {
        if (ids[i][0].toString() === targetId.toString()) {
            sheet.deleteRow(i + 2); // +2 because 1-indexed and header row
            deletedCount++;
        }
    }

    return jsonResponse({
        success: true,
        message: `Deleted ${deletedCount} row(s)`
    });
}

/**
 * Clear all reports for a specific month
 */
function handleClearMonth(data) {
    const sheet = getOrCreateSheet();
    const targetYear = data.year;
    const targetMonth = data.month; // 0-indexed
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
        return jsonResponse({ success: true, message: 'No reports to clear' });
    }

    const dates = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    let deletedCount = 0;

    for (let i = dates.length - 1; i >= 0; i--) {
        const d = new Date(dates[i][0]);
        if (d.getFullYear() === targetYear && d.getMonth() === targetMonth) {
            sheet.deleteRow(i + 2);
            deletedCount++;
        }
    }

    return jsonResponse({
        success: true,
        message: `Cleared ${deletedCount} row(s)`
    });
}

/**
 * Create a JSON response
 */
function jsonResponse(obj) {
    return ContentService
        .createTextOutput(JSON.stringify(obj))
        .setMimeType(ContentService.MimeType.JSON);
}
