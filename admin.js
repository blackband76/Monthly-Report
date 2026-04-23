// ===========================
// Configuration
// ===========================
// IMPORTANT: Replace this URL with your deployed Google Apps Script Web App URL
// Must be the same URL used in app.js
const GOOGLE_SHEETS_API_URL = 'https://script.google.com/macros/s/AKfycbxHyASbWgYOzIHc166QfMiZoETO976P88FSjkAgaDnFekeqDMge15J6hKkEWZ3zgsOJ/exec';
const STORAGE_KEY = 'monthlyReports';

// ===========================
// Constants
// ===========================
const MONTH_NAMES = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
];
const MONTH_SHORT = [
    'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
];

// ===========================
// State
// ===========================
const now = new Date();
let selectedYear = now.getFullYear();
let selectedMonth = now.getMonth();
let cachedReports = null;
let isLoading = false;

// ===========================
// API Helpers
// ===========================
function isApiConfigured() {
    return GOOGLE_SHEETS_API_URL && GOOGLE_SHEETS_API_URL.trim() !== '';
}

// ===========================
// Database (Google Sheets + localStorage fallback)
// ===========================
function getReportsLocal() {
    try {
        const data = localStorage.getItem(STORAGE_KEY);
        return data ? JSON.parse(data) : [];
    } catch (e) {
        console.error('Error reading reports:', e);
        return [];
    }
}

function saveReportsLocal(reports) {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(reports));
    } catch (e) {
        console.error('Error saving reports:', e);
    }
}

async function fetchReportsFromSheets() {
    const response = await fetch(GOOGLE_SHEETS_API_URL);
    const result = await response.json();
    if (!result.success) {
        throw new Error(result.error || 'Failed to fetch reports');
    }
    return result.reports || [];
}

async function getAllReports(forceRefresh = false) {
    if (cachedReports && !forceRefresh) {
        return cachedReports;
    }

    if (isApiConfigured()) {
        try {
            const sheetsReports = await fetchReportsFromSheets();
            // Convert flat rows back to report format for filtering
            cachedReports = sheetsReports;
            return cachedReports;
        } catch (err) {
            console.error('Failed to fetch from Google Sheets, falling back to localStorage:', err);
            showToast('Failed to load from Google Sheets. Showing local data.', 'warning');
            return flattenLocalReports();
        }
    } else {
        return flattenLocalReports();
    }
}

function flattenLocalReports() {
    const localReports = getReportsLocal();
    const rows = [];
    localReports.forEach(report => {
        if (report.projects && Array.isArray(report.projects)) {
            report.projects.forEach(proj => {
                rows.push({
                    id: report.id,
                    submittedAt: report.submittedAt,
                    name: report.name,
                    role: report.role || '',
                    project: proj.projectName,
                    keyHighlights: proj.keyHighlights || [],
                    upcomingFocus: proj.upcomingFocus || [],
                    issues: proj.issues || [],
                    concerns: proj.concerns || [],
                    risks: proj.risks || [],
                    needSupport: proj.needSupport || []
                });
            });
        } else {
            rows.push({
                id: report.id,
                submittedAt: report.submittedAt,
                name: report.name,
                role: report.role || '',
                project: report.project || '',
                keyHighlights: report.keyHighlights || [],
                upcomingFocus: report.upcomingFocus || [],
                issues: report.issues || [],
                concerns: report.concerns || [],
                risks: report.risks || [],
                needSupport: report.needSupport || []
            });
        }
    });
    return rows;
}

function filterByMonth(rows) {
    return rows.filter(row => {
        const date = new Date(row.submittedAt);
        return date.getFullYear() === selectedYear && date.getMonth() === selectedMonth;
    });
}

function getUniqueReportCount(rows) {
    const ids = new Set(rows.map(r => r.id));
    return ids.size;
}

async function deleteReport(id) {
    if (isApiConfigured()) {
        try {
            showTableLoading(true);
            const response = await fetch(GOOGLE_SHEETS_API_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({ action: 'delete', id: id })
            });
            const result = await response.json();
            if (!result.success) throw new Error(result.error);
            cachedReports = null;
            showToast('Report deleted', 'success');
        } catch (err) {
            console.error('Delete error:', err);
            showToast('Failed to delete report', 'error');
        }
    }

    // Also remove from localStorage
    let reports = getReportsLocal();
    reports = reports.filter(r => r.id !== id);
    saveReportsLocal(reports);

    await loadAndRender();
}

async function clearAllReports() {
    const allRows = await getAllReports();
    const filtered = filterByMonth(allRows);
    if (filtered.length === 0) return;

    const monthLabel = `${MONTH_NAMES[selectedMonth]} ${selectedYear}`;
    if (!confirm(`Are you sure you want to delete all ${filtered.length} report(s) for ${monthLabel}? This cannot be undone.`)) {
        return;
    }

    if (isApiConfigured()) {
        try {
            showTableLoading(true);
            const response = await fetch(GOOGLE_SHEETS_API_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({
                    action: 'clearMonth',
                    year: selectedYear,
                    month: selectedMonth
                })
            });
            const result = await response.json();
            if (!result.success) throw new Error(result.error);
            cachedReports = null;
            showToast(`Cleared all reports for ${monthLabel}`, 'success');
        } catch (err) {
            console.error('Clear error:', err);
            showToast('Failed to clear reports', 'error');
        }
    }

    // Also clear from localStorage
    let reports = getReportsLocal();
    reports = reports.filter(report => {
        const date = new Date(report.submittedAt);
        return !(date.getFullYear() === selectedYear && date.getMonth() === selectedMonth);
    });
    saveReportsLocal(reports);

    await loadAndRender();
}

// ===========================
// Toast Notifications
// ===========================
function showToast(message, type = 'success') {
    let container = document.getElementById('toastContainer');
    if (!container) {
        container = document.createElement('div');
        container.id = 'toastContainer';
        container.className = 'toast-container';
        document.body.appendChild(container);
    }

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    const iconMap = {
        success: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>',
        error: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>',
        warning: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>'
    };

    toast.innerHTML = `
        <span class="toast-icon">${iconMap[type] || iconMap.success}</span>
        <span class="toast-message">${message}</span>
    `;

    container.appendChild(toast);
    requestAnimationFrame(() => toast.classList.add('show'));

    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

// ===========================
// Month/Year Picker
// ===========================
async function changeYear(delta) {
    selectedYear += delta;
    document.getElementById('pickerYear').textContent = selectedYear;
    await loadAndRender();
}

async function selectMonth(month) {
    selectedMonth = month;
    await loadAndRender();
}

async function renderMonthGrid(allRows) {
    const grid = document.getElementById('monthGrid');

    const monthCounts = new Array(12).fill(0);
    allRows.forEach(row => {
        const date = new Date(row.submittedAt);
        if (date.getFullYear() === selectedYear) {
            monthCounts[date.getMonth()]++;
        }
    });

    grid.innerHTML = MONTH_SHORT
        .map((name, index) => {
            const isSelected = index === selectedMonth;
            const count = monthCounts[index];
            const hasReports = count > 0;
            return `
                <button type="button"
                    class="month-btn ${isSelected ? 'selected' : ''} ${hasReports ? 'has-reports' : ''}"
                    onclick="selectMonth(${index})">
                    <span class="month-name">${name}</span>
                    ${hasReports ? `<span class="month-badge">${count}</span>` : ''}
                </button>
            `;
        })
        .join('');
}

// ===========================
// Table Loading State
// ===========================
function showTableLoading(show) {
    isLoading = show;
    const wrapper = document.getElementById('reportsTableWrapper');
    const table = document.getElementById('reportsTable');
    const noMsg = document.getElementById('noReportsMsg');
    const loader = document.getElementById('tableLoader');

    if (show) {
        table.style.display = 'none';
        noMsg.style.display = 'none';
        if (loader) loader.style.display = '';
    } else {
        if (loader) loader.style.display = 'none';
    }
}

// ===========================
// Render Reports Table
// ===========================
function renderReportsTable(allRows) {
    const rows = filterByMonth(allRows);
    const table = document.getElementById('reportsTable');
    const tbody = document.getElementById('reportsTableBody');
    const noMsg = document.getElementById('noReportsMsg');
    const countEl = document.getElementById('reportCount');
    const titleEl = document.getElementById('reportsSectionTitle');
    const loader = document.getElementById('tableLoader');

    if (loader) loader.style.display = 'none';

    const monthLabel = `${MONTH_NAMES[selectedMonth]} ${selectedYear}`;
    titleEl.textContent = `Reports — ${monthLabel}`;

    const reportCount = getUniqueReportCount(rows);
    countEl.textContent = `${reportCount} report${reportCount !== 1 ? 's' : ''} (${rows.length} project${rows.length !== 1 ? 's' : ''})`;

    if (rows.length === 0) {
        table.style.display = 'none';
        noMsg.style.display = '';
        return;
    }

    table.style.display = '';
    noMsg.style.display = 'none';

    tbody.innerHTML = rows
        .map(
            (row, index) => `
            <tr>
                <td class="cell-number">${index + 1}</td>
                <td class="cell-date">${formatDate(row.submittedAt)}</td>
                <td>${escapeHtml(row.name)}</td>
                <td>${escapeHtml(row.role || '')}</td>
                <td>${escapeHtml(row.project)}</td>
                <td>${formatListCell(row.keyHighlights)}</td>
                <td>${formatListCell(row.upcomingFocus)}</td>
                <td>${formatListCell(row.issues)}</td>
                <td>${formatListCell(row.concerns)}</td>
                <td>${formatListCell(row.risks)}</td>
                <td>${formatListCell(row.needSupport)}</td>
                <td>
                    <button type="button" class="btn-remove-report" onclick="deleteReport('${row.id}')" title="Delete report">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <polyline points="3 6 5 6 21 6"/>
                            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
                        </svg>
                    </button>
                </td>
            </tr>
        `
        )
        .join('');
}

// ===========================
// Main Load & Render
// ===========================
async function loadAndRender() {
    showTableLoading(true);
    try {
        const allRows = await getAllReports(true);
        renderMonthGrid(allRows);
        renderReportsTable(allRows);
    } catch (err) {
        console.error('Error loading reports:', err);
        showToast('Error loading reports', 'error');
    } finally {
        showTableLoading(false);
    }
}

// ===========================
// Helpers
// ===========================
function formatListCell(items) {
    if (!items || items.length === 0) return '<span class="cell-empty">—</span>';
    if (items.length === 1) return escapeHtml(items[0]).replace(/\n/g, '<br>');
    return items.map((item, i) => `${i + 1}. ${escapeHtml(item).replace(/\n/g, '<br>')}`).join('<br>');
}

function formatDate(isoString) {
    const d = new Date(isoString);
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    const hours = String(d.getHours()).padStart(2, '0');
    const minutes = String(d.getMinutes()).padStart(2, '0');
    return `${day}/${month}/${year} ${hours}:${minutes}`;
}

// ===========================
// Export to Excel
// ===========================
async function exportToExcel() {
    const allRows = await getAllReports();
    const rows = filterByMonth(allRows);

    if (rows.length === 0) {
        alert('No reports to export for the selected period.');
        return;
    }

    const monthLabel = `${MONTH_NAMES[selectedMonth]} ${selectedYear}`;

    const excelData = rows.map((row, index) => ({
        'No.': index + 1,
        'Submitted Date': formatDate(row.submittedAt),
        'Name': row.name,
        'Role': row.role || '',
        'Project': row.project,
        'Key Highlights': (row.keyHighlights || []).map((h, i) => `${i + 1}. ${h}`).join('\n'),
        'Upcoming Focus': (row.upcomingFocus || []).map((h, i) => `${i + 1}. ${h}`).join('\n'),
        'Issues': (row.issues || []).map((h, i) => `${i + 1}. ${h}`).join('\n'),
        'Concerns': (row.concerns || []).map((h, i) => `${i + 1}. ${h}`).join('\n'),
        'Risks': (row.risks || []).map((h, i) => `${i + 1}. ${h}`).join('\n'),
        'Need Support': (row.needSupport || []).map((h, i) => `${i + 1}. ${h}`).join('\n')
    }));

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(excelData);

    ws['!cols'] = [
        { wch: 5 },   // No.
        { wch: 18 },  // Date
        { wch: 20 },  // Name
        { wch: 28 },  // Role
        { wch: 25 },  // Project
        { wch: 50 },  // Key Highlights
        { wch: 50 },  // Upcoming Focus
        { wch: 50 },  // Issues
        { wch: 50 },  // Concerns
        { wch: 50 },  // Risks
        { wch: 50 }   // Need Support
    ];

    XLSX.utils.book_append_sheet(wb, ws, monthLabel);

    const monthNum = String(selectedMonth + 1).padStart(2, '0');
    const filename = `Monthly_Reports_${selectedYear}_${monthNum}.xlsx`;
    XLSX.writeFile(wb, filename);
}

// ===========================
// Utility
// ===========================
function escapeHtml(str) {
    const div = document.createElement('div');
    div.appendChild(document.createTextNode(str));
    return div.innerHTML;
}

// ===========================
// Initialize
// ===========================
document.getElementById('pickerYear').textContent = selectedYear;
loadAndRender();
