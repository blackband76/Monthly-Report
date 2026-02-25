// ===========================
// Configuration
// ===========================
// IMPORTANT: Replace this URL with your deployed Google Apps Script Web App URL
const GOOGLE_SHEETS_API_URL = 'https://script.google.com/macros/s/AKfycbxHyASbWgYOzIHc166QfMiZoETO976P88FSjkAgaDnFekeqDMge15J6hKkEWZ3zgsOJ/exec';
const STORAGE_KEY = 'monthlyReports';

// ===========================
// Constants
// ===========================
const CATEGORIES = [
    { key: 'keyHighlight', label: 'Key Highlights', placeholder: 'Enter a key highlight...', iconClass: 'highlight-icon', icon: '<polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>' },
    { key: 'issue', label: 'Issues', placeholder: 'Enter an issue...', iconClass: 'issue-icon', icon: '<circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>' },
    { key: 'concern', label: 'Concerns', placeholder: 'Enter a concern...', iconClass: 'concern-icon', icon: '<path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>' },
    { key: 'risk', label: 'Risks', placeholder: 'Enter a risk...', iconClass: 'risk-icon', icon: '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>' },
    { key: 'needSupport', label: 'Need Support', placeholder: 'Enter support needed...', iconClass: 'support-icon', icon: '<circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/><line x1="12" y1="17" x2="12.01" y2="17"/>' }
];

const PROJECT_LIST = [
    "APEX WCiD",
    "API Hub",
    "Apigee",
    "Asset End-to-End Workflow Integration Study",
    "Auto Well and Final Well Report",
    "BCMS",
    "BRM Tasks",
    "Career & Talent Management - 2026 Enhancements",
    "Career & Talent Management - Module Hi-Potential Mgmt. (IDP)",
    "Cash Forecasting Ph. 2",
    "Cognite Data Fusion (CDF) for Production Operation",
    "Contract Intelligence",
    "Cost Excellence Ph.2",
    "Crew Boat Route Optimization (GOT)",
    "Crew Boat Routing Optimization - 2024",
    "Data Analyst & Analytic Engineer Pools",
    "Data Foundation Drilling",
    "Data Foundation Engineering",
    "Data Foundation Finance",
    "Data Foundation GCP support SAP SHANA",
    "Data Foundation HR",
    "Data Foundation Maintenance",
    "Data Foundation Production",
    "Data Foundation Supplychain",
    "Data Governance Program Lead Consultancy",
    "DCOE Meeting / Activities",
    "Delorean",
    "DevOps",
    "Digital PMO Ph.2",
    "DigitalX Hub",
    "Emergency Evacuation Ph.2",
    "EOT Timesheet",
    "EP - Common",
    "EP - DWE",
    "EP - DWE (Centralized Cost Sheet Application (Auto AFE))",
    "EP - EDM",
    "EP - EDM (Well-FAST)",
    "EP - EMI",
    "EP - EPI (APM)",
    "EP - HR (HR Hyperspace)",
    "EP - IOCC",
    "EP - IOCC (Main Dashboard)",
    "EP - IOCC (Workspace)",
    "EP - SCM",
    "EP - SCM (AI for Bid Evaluation (Vessel))",
    "EP - SCM (AI for Bid Evaluation)",
    "EP - SCM (AI for Contract Management)",
    "EP - SCM (AI for Inventory Management)",
    "EP - SCM (AI for Material Master)",
    "EP - SCM (AI for Pre-Qualification)",
    "EP - SCM (AI for Quality Assurance)",
    "EP - SCM (AI for Vessel Route Optimization)",
    "EP - SCM (Aviation & Marine Reallocation)",
    "EP - SCM (Contract Dashboard)",
    "EP - SCM (Inventory Dashboard)",
    "EP - SCM (Marine Dashboard)",
    "EP - SCM (Marketplace)",
    "EP - SCM (Material Tracking)",
    "EP - SCM (MSB Dashboard)",
    "EP - SCM (Probot)",
    "EP - SCM (PR-PO Performance Dashboard)",
    "EP - SCM (PSB Budget Management)",
    "EP - SCM (PSB Performance Dashboard)",
    "EP - SCM (Vessel Entry Permit)",
    "EP - Subsurface (KM)",
    "EP - Support",
    "EP - Support (Ucontrol Ph.3)",
    "EP - Sustainability (Decarbonization Intelligence Platform (DIP))",
    "EP Data Platform",
    "EP Platform",
    "EP Platform X.Brain",
    "ESP Curve Digitizer",
    "ETR Co-Pilot",
    "Foresight",
    "GRC One - UControl Ph.2",
    "GRC One - UControl Ph.3",
    "IDCA Offshore Zawtika",
    "iGFDP Ph.2",
    "INSURE Ph.3",
    "ITPM Governance",
    "iTQ",
    "LCM Insight",
    "LoC Enhancement",
    "Machine Learning Projects",
    "Medical Web Service",
    "NC Easy",
    "ODC/A Tasks",
    "ODC/E Tasks",
    "ODC/S Tasks",
    "ODT Tasks",
    "Offshore Timesheet",
    "OpEx Digital Assessment Platform",
    "OpEx GPT",
    "OWO Tasks",
    "PDA for Outsource",
    "Personal Leave",
    "Production and Load Management System",
    "PSL Rate Charge",
    "PTTEP Holidays",
    "Risk Register",
    "Running Club",
    "Salary Adjustment Letter",
    "SAP Data Platform Migration",
    "SFs Security Improvement for All Legacies (API)",
    "SHANA Migration Support, GCP data",
    "Sick Leave",
    "Sight Solution",
    "Single SOS",
    "Smart Doc AI",
    "Top of the Line Corrosion",
    "Turnaround Management",
    "WISDOM",
    "WPB Portal",
    "X-Tra Claim",
    "YourPA"
];

// ===========================
// State
// ===========================
let projects = [];
let projectIdCounter = 0;
let activeDropdown = null;
let isSubmitting = false;

// ===========================
// Database (Google Sheets + localStorage fallback)
// ===========================
function isApiConfigured() {
    return GOOGLE_SHEETS_API_URL && GOOGLE_SHEETS_API_URL.trim() !== '';
}

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

function saveReportLocal(report) {
    const reports = getReportsLocal();
    reports.push(report);
    saveReportsLocal(reports);
}

async function saveReportToSheets(report) {
    const response = await fetch(GOOGLE_SHEETS_API_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
            action: 'submit',
            id: report.id,
            submittedAt: report.submittedAt,
            name: report.name,
            projects: report.projects
        })
    });

    const result = await response.json();
    if (!result.success) {
        throw new Error(result.error || 'Failed to save report');
    }
    return result;
}

async function saveReport(report) {
    report.id = Date.now().toString();
    report.submittedAt = new Date().toISOString();

    if (isApiConfigured()) {
        try {
            await saveReportToSheets(report);
            // Also save locally as backup
            saveReportLocal(report);
            return { success: true };
        } catch (err) {
            console.error('Google Sheets API error, saving locally:', err);
            saveReportLocal(report);
            return { success: true, warning: 'Saved locally (Google Sheets unavailable)' };
        }
    } else {
        saveReportLocal(report);
        return { success: true, warning: 'No API configured — saved to local storage only' };
    }
}

// ===========================
// Toast Notifications
// ===========================
function showToast(message, type = 'success') {
    const container = document.getElementById('toastContainer');
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

    // Trigger animation
    requestAnimationFrame(() => {
        toast.classList.add('show');
    });

    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

// ===========================
// Project Management
// ===========================
function createProjectData() {
    const id = ++projectIdCounter;
    const project = {
        id: id,
        name: '',
        collapsed: false,
        categories: {}
    };
    CATEGORIES.forEach(cat => {
        project.categories[cat.key] = [];
    });
    return project;
}

function getProject(projectId) {
    return projects.find(p => p.id === projectId);
}

function addProject() {
    const project = createProjectData();
    projects.push(project);
    renderAllProjects();
    setTimeout(() => {
        const input = document.getElementById(`projectSearch_${project.id}`);
        if (input) input.focus();
    }, 100);
}

function removeProject(projectId) {
    if (projects.length <= 1) return;
    projects = projects.filter(p => p.id !== projectId);
    renderAllProjects();
}

function toggleProject(projectId) {
    const project = getProject(projectId);
    if (project) {
        project.collapsed = !project.collapsed;
        const body = document.getElementById(`projectBody_${projectId}`);
        const chevron = document.getElementById(`projectChevron_${projectId}`);
        if (body) body.classList.toggle('collapsed', project.collapsed);
        if (chevron) chevron.classList.toggle('rotated', project.collapsed);
    }
}

// ===========================
// Searchable Dropdown
// ===========================
function openDropdown(projectId) {
    closeAllDropdowns();
    activeDropdown = projectId;
    const dropdown = document.getElementById(`projectDropdown_${projectId}`);
    const searchInput = document.getElementById(`projectSearch_${projectId}`);
    if (dropdown) {
        dropdown.classList.add('open');
        filterDropdown(projectId, searchInput.value);
    }
}

function closeAllDropdowns() {
    if (activeDropdown !== null) {
        const dd = document.getElementById(`projectDropdown_${activeDropdown}`);
        if (dd) dd.classList.remove('open');
        activeDropdown = null;
    }
}

function filterDropdown(projectId, query) {
    const listEl = document.getElementById(`projectList_${projectId}`);
    if (!listEl) return;

    const q = query.toLowerCase().trim();
    const filtered = q === ''
        ? PROJECT_LIST
        : PROJECT_LIST.filter(p => p.toLowerCase().includes(q));

    if (filtered.length === 0) {
        listEl.innerHTML = '<div class="dropdown-empty">No matching projects</div>';
        return;
    }

    listEl.innerHTML = filtered.map(p => {
        const highlighted = q ? highlightMatch(p, q) : escapeHtml(p);
        return `<div class="dropdown-option" onmousedown="selectProject(${projectId}, '${escapeAttr(p)}')">${highlighted}</div>`;
    }).join('');
}

function highlightMatch(text, query) {
    const escaped = escapeHtml(text);
    const idx = text.toLowerCase().indexOf(query);
    if (idx === -1) return escaped;
    const before = escapeHtml(text.substring(0, idx));
    const match = escapeHtml(text.substring(idx, idx + query.length));
    const after = escapeHtml(text.substring(idx + query.length));
    return `${before}<mark>${match}</mark>${after}`;
}

function selectProject(projectId, projectName) {
    const project = getProject(projectId);
    if (project) {
        project.name = projectName;
        const searchInput = document.getElementById(`projectSearch_${projectId}`);
        if (searchInput) {
            searchInput.value = projectName;
        }
    }
    closeAllDropdowns();
}

function onProjectSearchInput(projectId) {
    const searchInput = document.getElementById(`projectSearch_${projectId}`);
    const project = getProject(projectId);
    if (project) project.name = searchInput.value;
    filterDropdown(projectId, searchInput.value);

    const dropdown = document.getElementById(`projectDropdown_${projectId}`);
    if (dropdown && !dropdown.classList.contains('open')) {
        dropdown.classList.add('open');
        activeDropdown = projectId;
    }
}

function onProjectSearchBlur(projectId) {
    setTimeout(() => {
        closeAllDropdowns();
    }, 200);
}

document.addEventListener('click', function (e) {
    if (!e.target.closest('.project-search-wrapper')) {
        closeAllDropdowns();
    }
});

// ===========================
// Show Input / Confirm Item
// ===========================
function showInput(projectId, catKey) {
    const inputRow = document.getElementById(`inputRow_${projectId}_${catKey}`);
    const textarea = document.getElementById(`input_${projectId}_${catKey}`);

    // Auto-confirm existing text if any
    if (textarea && textarea.value.trim()) {
        silentConfirm(projectId, catKey);
    }

    if (inputRow) {
        inputRow.style.display = '';
        textarea.value = '';
        textarea.focus();
    }
}

function silentConfirm(projectId, catKey) {
    const textarea = document.getElementById(`input_${projectId}_${catKey}`);
    const inputRow = document.getElementById(`inputRow_${projectId}_${catKey}`);
    const value = textarea.value.trim();

    if (!value) {
        // Blank — just hide the input
        if (inputRow) inputRow.style.display = 'none';
        return;
    }

    const project = getProject(projectId);
    if (project) {
        project.categories[catKey].push(value);
        textarea.value = '';
        inputRow.style.display = 'none';
        renderCategoryList(projectId, catKey);
    }
}

function onItemBlur(projectId, catKey) {
    // Small delay to allow "Add" button click to fire first
    setTimeout(() => {
        const textarea = document.getElementById(`input_${projectId}_${catKey}`);
        // Only act if textarea is not focused (user didn't click back into it)
        if (textarea && document.activeElement !== textarea) {
            silentConfirm(projectId, catKey);
        }
    }, 200);
}

function removeItem(projectId, catKey, index) {
    const project = getProject(projectId);
    if (!project) return;

    const list = document.getElementById(`list_${projectId}_${catKey}`);
    const items = list.querySelectorAll('.list-item');
    const item = items[index];

    if (item) {
        item.style.animation = 'slideOut 0.3s ease forwards';
        setTimeout(() => {
            project.categories[catKey].splice(index, 1);
            renderCategoryList(projectId, catKey);
        }, 280);
    }
}

// ===========================
// Render
// ===========================
function renderAllProjects() {
    projects.forEach(project => {
        const nameInput = document.getElementById(`projectSearch_${project.id}`);
        if (nameInput) project.name = nameInput.value;
    });

    const container = document.getElementById('projectsContainer');
    container.innerHTML = projects.map((project, index) => renderProjectCard(project, index)).join('');
}

function renderProjectCard(project, index) {
    const sectionNumber = index + 1;
    const canDelete = projects.length > 1;

    let sectionsHtml = CATEGORIES.map(cat => {
        const items = project.categories[cat.key];
        const listHtml = renderCategoryListHtml(project.id, cat.key, items);

        return `
            <div class="project-category" id="catSection_${project.id}_${cat.key}">
                <div class="category-header">
                    <div class="section-icon ${cat.iconClass}">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${cat.icon}</svg>
                    </div>
                    <h3>${cat.label}</h3>
                </div>
                <div class="dynamic-list" id="list_${project.id}_${cat.key}">
                    ${listHtml}
                </div>
                <div class="add-item-row" id="inputRow_${project.id}_${cat.key}" style="display:none;">
                    <textarea id="input_${project.id}_${cat.key}" class="add-item-input" placeholder="${cat.placeholder}" rows="3" onblur="onItemBlur(${project.id}, '${cat.key}')"></textarea>
                </div>
                <div class="add-btn-center">
                    <button type="button" class="btn-add-center" onclick="showInput(${project.id}, '${cat.key}')">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                            <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
                        </svg>
                        Add
                    </button>
                </div>
            </div>
        `;
    }).join('');

    return `
        <div class="project-card" id="project_${project.id}">
            <div class="project-card-header" onclick="toggleProject(${project.id})">
                <div class="project-card-title">
                    <span class="project-number">${sectionNumber}</span>
                    <div class="project-search-wrapper" onclick="event.stopPropagation()">
                        <input type="text" class="project-search-input" id="projectSearch_${project.id}"
                            placeholder="Search project..."
                            value="${escapeAttr(project.name)}"
                            autocomplete="off"
                            onfocus="openDropdown(${project.id})"
                            oninput="onProjectSearchInput(${project.id})"
                            onblur="onProjectSearchBlur(${project.id})"
                            required>
                        <svg class="project-search-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
                        </svg>
                        <div class="project-dropdown" id="projectDropdown_${project.id}">
                            <div class="project-dropdown-list" id="projectList_${project.id}"></div>
                        </div>
                    </div>
                </div>
                <div class="project-card-actions">
                    ${canDelete ? `
                        <button type="button" class="btn-remove-project" onclick="event.stopPropagation(); removeProject(${project.id})" title="Remove project">
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
                            </svg>
                        </button>
                    ` : ''}
                    <svg class="project-chevron ${project.collapsed ? 'rotated' : ''}" id="projectChevron_${project.id}" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <polyline points="6 9 12 15 18 9"/>
                    </svg>
                </div>
            </div>
            <div class="project-card-body ${project.collapsed ? 'collapsed' : ''}" id="projectBody_${project.id}">
                ${sectionsHtml}
            </div>
        </div>
    `;
}

function renderCategoryListHtml(projectId, catKey, items) {
    if (!items || items.length === 0) {
        return `<div class="empty-state">No items added yet.</div>`;
    }
    return items.map((item, index) => `
        <div class="list-item">
            <span class="item-number cat-${catKey}">${index + 1}</span>
            <span class="item-text">${escapeHtml(item).replace(/\n/g, '<br>')}</span>
            <button type="button" class="btn-remove" onclick="removeItem(${projectId}, '${catKey}', ${index})" title="Remove item">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
            </button>
        </div>
    `).join('');
}

function renderCategoryList(projectId, catKey) {
    const project = getProject(projectId);
    if (!project) return;
    const list = document.getElementById(`list_${projectId}_${catKey}`);
    list.innerHTML = renderCategoryListHtml(projectId, catKey, project.categories[catKey]);
}

// ===========================
// Clear Form
// ===========================
function clearForm() {
    document.getElementById('nameInput').value = '';
    projects = [];
    projectIdCounter = 0;
    const first = createProjectData();
    projects.push(first);
    renderAllProjects();
}

// ===========================
// Submit Form
// ===========================
function setSubmitLoading(loading) {
    isSubmitting = loading;
    const btn = document.getElementById('submitReportBtn');
    if (loading) {
        btn.disabled = true;
        btn.innerHTML = `
            <span class="btn-spinner"></span>
            Submitting...
        `;
    } else {
        btn.disabled = false;
        btn.innerHTML = `
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M22 2L11 13"/>
                <path d="M22 2L15 22L11 13L2 9L22 2Z"/>
            </svg>
            Submit Report
        `;
    }
}

document.getElementById('reportForm').addEventListener('submit', async function (e) {
    e.preventDefault();
    if (isSubmitting) return;

    const name = document.getElementById('nameInput').value.trim();
    if (!name) {
        document.getElementById('nameInput').focus();
        return;
    }

    const projectsData = [];
    let hasError = false;

    projects.forEach(project => {
        const searchInput = document.getElementById(`projectSearch_${project.id}`);
        const projectName = searchInput ? searchInput.value.trim() : '';

        if (!projectName) {
            searchInput.focus();
            searchInput.classList.add('shake');
            setTimeout(() => searchInput.classList.remove('shake'), 500);
            hasError = true;
            return;
        }

        projectsData.push({
            projectName: projectName,
            keyHighlights: [...project.categories.keyHighlight],
            issues: [...project.categories.issue],
            concerns: [...project.categories.concern],
            risks: [...project.categories.risk],
            needSupport: [...project.categories.needSupport]
        });
    });

    if (hasError) return;

    const report = {
        name: name,
        projects: projectsData
    };

    setSubmitLoading(true);

    try {
        const result = await saveReport(report);

        document.getElementById('modalSummary').innerHTML = buildSummaryHtml(name, projectsData);
        document.getElementById('successModal').classList.add('active');

        if (result.warning) {
            showToast(result.warning, 'warning');
        } else {
            showToast('Report submitted successfully!', 'success');
        }

        clearForm();
    } catch (err) {
        console.error('Submit error:', err);
        showToast('Failed to submit report. Please try again.', 'error');
    } finally {
        setSubmitLoading(false);
    }
});

// ===========================
// Build Summary HTML
// ===========================
function buildSummaryHtml(name, projectsData) {
    const catMeta = [
        { key: 'keyHighlights', label: 'Key Highlights', colorClass: 'highlight' },
        { key: 'issues', label: 'Issues', colorClass: 'issue' },
        { key: 'concerns', label: 'Concerns', colorClass: 'concern' },
        { key: 'risks', label: 'Risks', colorClass: 'risk' },
        { key: 'needSupport', label: 'Need Support', colorClass: 'support' }
    ];

    let html = `
        <div class="summary-row">
            <span class="summary-label">Name:</span>
            <span class="summary-value">${escapeHtml(name)}</span>
        </div>
    `;

    projectsData.forEach((proj, idx) => {
        html += `<div class="summary-project">
            <div class="summary-project-title">Project ${idx + 1}: ${escapeHtml(proj.projectName)}</div>`;

        catMeta.forEach(cat => {
            const items = proj[cat.key] || [];
            html += `<div class="summary-section">
                <div class="summary-section-title ${cat.colorClass}">${cat.label}</div>`;
            if (items.length > 0) {
                html += '<ul class="summary-list">';
                items.forEach(item => {
                    html += `<li>${escapeHtml(item).replace(/\n/g, '<br>')}</li>`;
                });
                html += '</ul>';
            } else {
                html += `<div class="summary-empty">No items reported</div>`;
            }
            html += '</div>';
        });

        html += '</div>';
    });

    return html;
}

// ===========================
// Close Modal
// ===========================
function closeModal() {
    document.getElementById('successModal').classList.remove('active');
}

document.getElementById('successModal').addEventListener('click', function (e) {
    if (e.target === this) closeModal();
});

document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') closeModal();
});

// ===========================
// Utilities
// ===========================
function escapeHtml(str) {
    const div = document.createElement('div');
    div.appendChild(document.createTextNode(str));
    return div.innerHTML;
}

function escapeAttr(str) {
    return str.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ===========================
// Initialize
// ===========================
const firstProject = createProjectData();
projects.push(firstProject);
renderAllProjects();

// Show config warning if API not set
if (!isApiConfigured()) {
    console.warn('Google Sheets API URL not configured. Reports will only be saved to localStorage.');
}

// Shake animation
const shakeStyle = document.createElement('style');
shakeStyle.textContent = `
    @keyframes shake {
        0%, 100% { transform: translateX(0); }
        20% { transform: translateX(-6px); }
        40% { transform: translateX(6px); }
        60% { transform: translateX(-4px); }
        80% { transform: translateX(4px); }
    }
    .shake {
        animation: shake 0.4s ease;
        border-color: var(--issue-color) !important;
    }
`;
document.head.appendChild(shakeStyle);
