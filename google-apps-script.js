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

/**
 * Initialize the sheet with headers if empty
 */
function getOrCreateSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME);
    }

    if (sheet.getLastRow() === 0) {
        sheet.appendRow([
            'ID',
            'Submitted Date',
            'Name',
            'Project',
            'Key Highlights',
            'Issues',
            'Concerns',
            'Risks',
            'Need Support'
        ]);

        // Format header row
        const headerRange = sheet.getRange(1, 1, 1, 9);
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#4f46e5');
        headerRange.setFontColor('#ffffff');

        // Set column widths
        sheet.setColumnWidth(1, 180);  // ID
        sheet.setColumnWidth(2, 160);  // Date
        sheet.setColumnWidth(3, 150);  // Name
        sheet.setColumnWidth(4, 250);  // Project
        sheet.setColumnWidth(5, 400);  // Key Highlights
        sheet.setColumnWidth(6, 400);  // Issues
        sheet.setColumnWidth(7, 400);  // Concerns
        sheet.setColumnWidth(8, 400);  // Risks
        sheet.setColumnWidth(9, 400);  // Need Support

        sheet.setFrozenRows(1);
    }

    return sheet;
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

        const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
        const reports = data.map(row => ({
            id: row[0],
            submittedAt: row[1],
            name: row[2],
            project: row[3],
            keyHighlights: parseItems(row[4]),
            issues: parseItems(row[5]),
            concerns: parseItems(row[6]),
            risks: parseItems(row[7]),
            needSupport: parseItems(row[8])
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
    const projects = data.projects || [];

    const rows = [];
    projects.forEach(proj => {
        rows.push([
            id,
            submittedAt,
            name,
            proj.projectName || '',
            formatItems(proj.keyHighlights || []),
            formatItems(proj.issues || []),
            formatItems(proj.concerns || []),
            formatItems(proj.risks || []),
            formatItems(proj.needSupport || [])
        ]);
    });

    if (rows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
        // Enable text wrap for data columns
        const lastRow = sheet.getLastRow();
        sheet.getRange(lastRow - rows.length + 1, 5, rows.length, 5).setWrap(true);
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
