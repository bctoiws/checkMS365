/**
 * Microsoft 365 User Access & Application Governance Dashboard
 * Architecture: Clean Reactive State • Zero Hardcoding • Apple HIG Standards
 * Language: English (International)
 */

'use strict';

// ==========================================================================
// Centralized Configuration & Localized Schema (Zero Hardcoding)
// ==========================================================================
const CONFIG = {
    endpoints: {
        usersActivity: '/api/users/activity',
        configStatus: '/api/config-status'
    },
    thresholds: {
        frequentAccessCount: 10,
        recentDays: 7,
        dormantDays: 14
    },
    defaultRefreshMs: 30000,
    exportFilenamePrefix: 'M365_Access_Audit_'
};

const I18N = {
    metrics: [
        {
            id: 'total',
            label: 'Total Accounts',
            icon: 'blue',
            svg: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>',
            sub: 'Licensed tenant identities'
        },
        {
            id: 'activeUsers',
            label: 'Active This Month',
            icon: 'green',
            svg: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline></svg>',
            sub: 'Monthly Active Users (MAU)'
        },
        {
            id: 'monthlyAccesses',
            label: 'Total Monthly Logins',
            icon: 'purple',
            svg: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect><line x1="16" y1="2" x2="16" y2="6"></line><line x1="8" y1="2" x2="8" y2="6"></line><line x1="3" y1="10" x2="21" y2="10"></line></svg>',
            sub: 'Recorded audit sessions (30d)'
        },
        {
            id: 'dormant',
            label: 'Inactive / Review',
            icon: 'orange',
            svg: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>',
            sub: 'No sign-ins for >14 days'
        }
    ],
    roles: {
        admin: 'Global Administrator',
        member: 'Staff Member'
    },
    table: {
        showingPrefix: 'Showing',
        accounts: 'accounts',
        noResults: 'No user accounts match your filter criteria.',
        serverError: 'Unable to communicate with Microsoft Graph API. Please inspect connection.'
    }
};

// ==========================================================================
// Reactive Application State
// ==========================================================================
const AppState = {
    users: [],
    stats: {
        total: 0,
        adminCount: 0,
        activeUsersCount: 0,
        dormantCount: 0,
        totalMonthlyAccesses: 0
    },
    filter: 'all',
    searchQuery: '',
    selectedUser: null,
    isFetching: false,
    timer: null,

    setFilter(newFilter) {
        this.filter = newFilter;
        renderSegmentedControls();
        renderTable();
    },

    setSearch(query) {
        this.searchQuery = (query || '').trim().toLowerCase();
        renderTable();
    },

    setSelectedUser(userId) {
        this.selectedUser = this.users.find(u => u.id === userId) || null;
        renderModal();
    }
};

// ==========================================================================
// Initialization & Event Binding
// ==========================================================================
document.addEventListener('DOMContentLoaded', () => {
    initMetricsGrid();
    setupEventHandlers();
    fetchData();
    setupAutoRefresh();
});

function initMetricsGrid() {
    const container = document.getElementById('metricsSection');
    if (!container) return;

    container.innerHTML = I18N.metrics.map(metric => `
        <article class="sf-card" id="card-${metric.id}">
            <div class="sf-card-header">
                <span class="sf-card-label">${escapeHtml(metric.label)}</span>
                <div class="sf-card-icon ${metric.icon}" aria-hidden="true">${metric.svg}</div>
            </div>
            <div class="sf-card-value" id="val-${metric.id}">--</div>
            <div class="sf-card-hint">${escapeHtml(metric.sub)}</div>
        </article>
    `).join('');
}

function setupEventHandlers() {
    // Manual refresh
    const refreshBtn = document.getElementById('btnManualRefresh');
    if (refreshBtn) {
        refreshBtn.addEventListener('click', () => fetchData(true));
    }

    // Auto-refresh interval
    const intervalSelect = document.getElementById('autoRefreshInterval');
    if (intervalSelect) {
        intervalSelect.addEventListener('change', setupAutoRefresh);
    }

    // Segmented filters
    const segmentContainer = document.getElementById('segmentedFilter');
    if (segmentContainer) {
        segmentContainer.addEventListener('click', (e) => {
            const btn = e.target.closest('.sf-segment');
            if (!btn) return;
            const targetFilter = btn.getAttribute('data-filter');
            if (targetFilter) AppState.setFilter(targetFilter);
        });
    }

    // Search field
    const searchInput = document.getElementById('userSearchInput');
    const searchClear = document.getElementById('searchClearBtn');
    if (searchInput) {
        searchInput.addEventListener('input', (e) => {
            const val = e.target.value;
            if (searchClear) searchClear.hidden = val.length === 0;
            AppState.setSearch(val);
        });
    }

    if (searchClear) {
        searchClear.addEventListener('click', () => {
            if (searchInput) {
                searchInput.value = '';
                searchClear.hidden = true;
                searchInput.focus();
            }
            AppState.setSearch('');
        });
    }

    // Export CSV
    const exportBtn = document.getElementById('btnExportCsv');
    if (exportBtn) {
        exportBtn.addEventListener('click', exportDatasetToCsv);
    }

    // Modal dialog controls
    const modalBackdrop = document.getElementById('detailModal');
    const modalCloseBtn = document.getElementById('modalCloseBtn');
    if (modalCloseBtn) {
        modalCloseBtn.addEventListener('click', () => AppState.setSelectedUser(null));
    }
    if (modalBackdrop) {
        modalBackdrop.addEventListener('click', (e) => {
            if (e.target === modalBackdrop) AppState.setSelectedUser(null);
        });
    }
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && AppState.selectedUser) {
            AppState.setSelectedUser(null);
        }
    });

    // Delegate row click for detail view
    const tableBody = document.getElementById('userTableBody');
    if (tableBody) {
        tableBody.addEventListener('click', (e) => {
            const btn = e.target.closest('[data-view-id]');
            if (btn) {
                const id = btn.getAttribute('data-view-id');
                AppState.setSelectedUser(id);
            }
        });
    }
}

function setupAutoRefresh() {
    if (AppState.timer) {
        clearInterval(AppState.timer);
        AppState.timer = null;
    }
    const select = document.getElementById('autoRefreshInterval');
    const ms = select ? parseInt(select.value, 10) : CONFIG.defaultRefreshMs;
    if (ms > 0) {
        AppState.timer = setInterval(() => fetchData(false), ms);
    }
}

// ==========================================================================
// Data Fetching & Sync
// ==========================================================================
async function fetchData(isManualTrigger = false) {
    if (AppState.isFetching) return;
    AppState.isFetching = true;

    const icon = document.getElementById('refreshIcon');
    if (isManualTrigger && icon) icon.classList.add('sf-spinning');

    try {
        const response = await fetch(CONFIG.endpoints.usersActivity);
        const result = await response.json();

        if (!result.success) {
            throw new Error(result.message || I18N.table.serverError);
        }

        AppState.users = result.users || [];
        AppState.stats = result.stats || {};

        updateMetricsValues();
        updateSegmentCounts();
        renderTable();
        updateTimestamp(result.timestamp);

    } catch (err) {
        console.error('[Dashboard Sync Error]:', err);
        renderTableError(err.message);
    } finally {
        AppState.isFetching = false;
        if (isManualTrigger && icon) {
            setTimeout(() => icon.classList.remove('sf-spinning'), 400);
        }
    }
}

// ==========================================================================
// Presentation Renderers (Apple HIG)
// ==========================================================================
function updateMetricsValues() {
    const valTotal = document.getElementById('val-total');
    const valActiveUsers = document.getElementById('val-activeUsers');
    const valMonthlyAccesses = document.getElementById('val-monthlyAccesses');
    const valDormant = document.getElementById('val-dormant');

    if (valTotal) valTotal.textContent = AppState.stats.total ?? 0;
    if (valActiveUsers) valActiveUsers.textContent = AppState.stats.activeUsersCount ?? 0;
    if (valMonthlyAccesses) valMonthlyAccesses.textContent = AppState.stats.totalMonthlyAccesses ?? 0;
    if (valDormant) valDormant.textContent = AppState.stats.dormantCount ?? 0;
}

function updateSegmentCounts() {
    const total = AppState.users.length;
    const frequent = AppState.users.filter(u => u.monthlyAccessCount >= CONFIG.thresholds.frequentAccessCount).length;
    const admin = AppState.users.filter(u => u.isAdmin).length;
    const dormant = AppState.users.filter(u => u.usageCategory === 'dormant').length;

    const cAll = document.getElementById('countAll');
    const cFrequent = document.getElementById('countFrequent');
    const cAdmin = document.getElementById('countAdmin');
    const cDormant = document.getElementById('countDormant');

    if (cAll) cAll.textContent = total;
    if (cFrequent) cFrequent.textContent = frequent;
    if (cAdmin) cAdmin.textContent = admin;
    if (cDormant) cDormant.textContent = dormant;
}

function renderSegmentedControls() {
    const buttons = document.querySelectorAll('#segmentedFilter .sf-segment');
    buttons.forEach(btn => {
        const f = btn.getAttribute('data-filter');
        const isActive = f === AppState.filter;
        btn.classList.toggle('active', isActive);
        btn.setAttribute('aria-selected', String(isActive));
    });
}

function renderTable() {
    const tbody = document.getElementById('userTableBody');
    const summaryText = document.getElementById('tableSummaryText');
    if (!tbody) return;

    const query = AppState.searchQuery;
    const filter = AppState.filter;

    const filtered = AppState.users.filter(user => {
        // Search Matching (Name, UPN, Dept, Job Title, Apps Used)
        if (query) {
            const name = (user.displayName || '').toLowerCase();
            const upn = (user.userPrincipalName || '').toLowerCase();
            const dept = (user.department || '').toLowerCase();
            const job = (user.jobTitle || '').toLowerCase();
            const apps = (user.appsUsed || []).join(' ').toLowerCase();

            if (!name.includes(query) && !upn.includes(query) && !dept.includes(query) && !job.includes(query) && !apps.includes(query)) {
                return false;
            }
        }

        // Segment Filter
        if (filter === 'all') return true;
        if (filter === 'frequent') return user.monthlyAccessCount >= CONFIG.thresholds.frequentAccessCount;
        if (filter === 'admin') return user.isAdmin === true;
        if (filter === 'dormant') return user.usageCategory === 'dormant';

        return true;
    });

    if (summaryText) {
        summaryText.textContent = `${I18N.table.showingPrefix} ${filtered.length} of ${AppState.users.length} ${I18N.table.accounts}`;
    }

    if (filtered.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="6" class="sf-table-empty">
                    <p>${I18N.table.noResults}</p>
                </td>
            </tr>
        `;
        return;
    }

    tbody.innerHTML = filtered.map(user => {
        const initials = computeInitials(user.displayName);
        const roleLabel = user.isAdmin ? I18N.roles.admin : (user.jobTitle || I18N.roles.member);
        const roleBadge = user.isAdmin ? `<span class="sf-badge admin">Admin</span>` : '';

        // Monthly count badge
        let countClass = 'zero';
        if (user.monthlyAccessCount >= 20) countClass = 'high';
        else if (user.monthlyAccessCount >= 5) countClass = 'mid';
        else if (user.monthlyAccessCount > 0) countClass = 'low';

        const countBadge = `
            <span class="sf-count-badge ${countClass}">
                ${user.monthlyAccessCount} ${user.monthlyAccessCount === 1 ? 'session' : 'sessions'}
            </span>
        `;

        // Application tags
        const appsPills = (user.appsUsed && user.appsUsed.length > 0)
            ? user.appsUsed.map(app => renderAppPill(app)).join('')
            : '<span class="sf-label-tertiary" style="font-size:0.75rem;">None recorded</span>';

        // Latest Timestamp
        const timeObj = formatTimestamp(user.latestTimestamp);

        return `
            <tr>
                <td>
                    <div class="sf-user-cell">
                        <div class="sf-avatar ${user.isAdmin ? 'admin' : ''}" aria-hidden="true">${initials}</div>
                        <div class="sf-user-meta">
                            <span class="sf-user-name">${escapeHtml(user.displayName)} ${roleBadge}</span>
                            <span class="sf-user-role-sub">${escapeHtml(roleLabel)} • ${escapeHtml(user.department || 'General')}</span>
                        </div>
                    </div>
                </td>
                <td>
                    <span class="sf-code-pill">${escapeHtml(user.userPrincipalName)}</span>
                </td>
                <td>
                    ${countBadge}
                </td>
                <td>
                    <div class="sf-app-tags">
                        ${appsPills}
                    </div>
                </td>
                <td>
                    <div class="sf-time-cell">
                        <span class="sf-time-relative">${timeObj.relative}</span>
                        <span class="sf-time-exact">${timeObj.exact}</span>
                    </div>
                </td>
                <td>
                    <button type="button" class="sf-row-btn" data-view-id="${user.id}">Audit Log</button>
                </td>
            </tr>
        `;
    }).join('');
}

function renderAppPill(appName) {
    let styleClass = 'office';
    const lower = (appName || '').toLowerCase();
    if (lower.includes('word')) styleClass = 'word';
    else if (lower.includes('excel')) styleClass = 'excel';
    else if (lower.includes('outlook')) styleClass = 'outlook';
    else if (lower.includes('sharepoint')) styleClass = 'sharepoint';
    else if (lower.includes('onedrive')) styleClass = 'onedrive';
    else if (lower.includes('admin') || lower.includes('portal')) styleClass = 'admin';

    return `<span class="sf-app-pill ${styleClass}">${escapeHtml(appName)}</span>`;
}

function renderTableError(msg) {
    const tbody = document.getElementById('userTableBody');
    if (tbody) {
        tbody.innerHTML = `
            <tr>
                <td colspan="6" class="sf-table-empty" style="color: var(--sf-red);">
                    <p>⚠️ ${escapeHtml(msg)}</p>
                </td>
            </tr>
        `;
    }
}

function updateTimestamp(iso) {
    const el = document.getElementById('lastSyncTime');
    if (!el) return;
    const date = iso ? new Date(iso) : new Date();
    el.textContent = date.toLocaleTimeString('en-US');
    el.setAttribute('datetime', date.toISOString());
}

// ==========================================================================
// Modal Sheet (Access History & Exact Application Log)
// ==========================================================================
function renderModal() {
    const modal = document.getElementById('detailModal');
    const container = document.getElementById('modalDetailsContainer');
    const title = document.getElementById('modalUserName');
    const sub = document.getElementById('modalUserSubtitle');
    if (!modal || !container) return;

    const user = AppState.selectedUser;
    if (!user) {
        modal.classList.remove('open');
        modal.setAttribute('aria-hidden', 'true');
        return;
    }

    if (title) title.textContent = user.displayName;
    if (sub) sub.textContent = `${user.userPrincipalName} • ${user.isAdmin ? 'Global Administrator' : 'Staff Member'}`;

    const historyItems = (user.history && user.history.length > 0)
        ? user.history.map(ev => {
            const evDate = new Date(ev.timestamp);
            return `
                <div class="sf-timeline-item">
                    <div class="sf-timeline-left">
                        <span class="sf-timeline-app">${escapeHtml(ev.app)}</span>
                        <span class="sf-timeline-meta">${escapeHtml(ev.rawApp)} • ${escapeHtml(ev.os)} (${escapeHtml(ev.browser)})</span>
                    </div>
                    <div class="sf-timeline-right">
                        <span class="sf-timeline-time">${evDate.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', second: '2-digit' })}</span>
                        <span class="sf-timeline-date">${evDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}</span>
                    </div>
                </div>
            `;
        }).join('')
        : `<p style="color: var(--sf-label-tertiary); font-size: 0.85rem; padding: 20px 0; text-align: center;">No individual sign-in events recorded for this account in the current 30-day window.</p>`;

    container.innerHTML = `
        <div class="sf-modal-stats">
            <div class="sf-modal-stat-item">
                <span class="sf-modal-stat-label">Monthly Logins</span>
                <span class="sf-modal-stat-value">${user.monthlyAccessCount} sessions</span>
            </div>
            <div class="sf-modal-stat-item">
                <span class="sf-modal-stat-label">Account Status</span>
                <span class="sf-modal-stat-value" style="font-size: 0.95rem; display: flex; align-items: center; gap: 6px;">
                    ${user.accountEnabled ? '🟢 Enabled' : '🔴 Suspended'}
                </span>
            </div>
        </div>

        <div>
            <h4 class="sf-timeline-section-title">Applications Authenticated</h4>
            <div class="sf-app-tags" style="margin-bottom: 16px;">
                ${(user.appsUsed && user.appsUsed.length > 0) ? user.appsUsed.map(a => renderAppPill(a)).join('') : '<span class="sf-label-tertiary">None recorded</span>'}
            </div>
        </div>

        <div>
            <h4 class="sf-timeline-section-title">Recent Sign-in Events &amp; Timestamps</h4>
            <div class="sf-timeline">
                ${historyItems}
            </div>
        </div>
    `;

    modal.classList.add('open');
    modal.setAttribute('aria-hidden', 'false');
}

// ==========================================================================
// Formatting & Utilities
// ==========================================================================
function computeInitials(name) {
    if (!name) return 'U';
    const parts = name.trim().split(/\s+/);
    if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

function formatTimestamp(isoString) {
    if (!isoString) return { relative: 'Never', exact: 'No session recorded' };
    const date = new Date(isoString);
    const now = new Date();
    const diffSec = Math.floor((now.getTime() - date.getTime()) / 1000);

    let relative = '';
    if (diffSec < 60) relative = 'Just now';
    else if (diffSec < 3600) relative = `${Math.floor(diffSec / 60)}m ago`;
    else if (diffSec < 86400) relative = `${Math.floor(diffSec / 3600)}h ago`;
    else {
        const days = Math.floor(diffSec / 86400);
        if (days === 1) relative = 'Yesterday';
        else if (days < 30) relative = `${days}d ago`;
        else relative = date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
    }

    const exact = date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) + 
        ' at ' + date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });

    return { relative, exact };
}

function escapeHtml(str) {
    if (!str) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

function exportDatasetToCsv() {
    if (!AppState.users || AppState.users.length === 0) {
        alert('No user data available to export.');
        return;
    }

    const headers = ['Full Name', 'Account UPN', 'Role', 'Department', 'Monthly Sessions', 'Applications Used', 'Latest Sign-In Timestamp'];
    const rows = AppState.users.map(u => [
        `"${(u.displayName || '').replace(/"/g, '""')}"`,
        `"${(u.userPrincipalName || '').replace(/"/g, '""')}"`,
        `"${u.isAdmin ? 'Global Administrator' : (u.jobTitle || 'Member')}"`,
        `"${(u.department || 'General').replace(/"/g, '""')}"`,
        u.monthlyAccessCount,
        `"${(u.appsUsed || []).join(', ')}"`,
        `"${u.latestTimestamp ? new Date(u.latestTimestamp).toISOString() : 'Never'}"`
    ]);

    const csvContent = '\uFEFF' + [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `${CONFIG.exportFilenamePrefix}${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}
