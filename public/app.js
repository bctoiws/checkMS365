/**
 * Microsoft 365 Identity & Activity Governance Dashboard
 * Architecture: Clean Reactive State • Zero Hardcoding • Apple Human Interface Guidelines
 */

'use strict';

// ==========================================================================
// Centralized Configuration & Localized Schema (No Hardcoded Strings)
// ==========================================================================
const CONFIG = {
    endpoints: {
        usersActivity: '/api/users/activity',
        configStatus: '/api/config-status'
    },
    thresholds: {
        activeDays: 3,
        moderateDays: 14,
        idleDays: 30
    },
    defaultRefreshMs: 30000,
    exportFilenamePrefix: 'M365_Activity_Report_'
};

const I18N = {
    metrics: [
        {
            id: 'total',
            label: 'Tổng Tài Khoản Cấp Phát',
            icon: 'blue',
            svg: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>',
            sub: 'Toàn bộ giấy phép nội bộ'
        },
        {
            id: 'admin',
            label: 'Quản Trị Viên Toàn Cầu',
            icon: 'purple',
            svg: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"></path></svg>',
            sub: 'Quyền kiểm soát cao nhất (Global Admin)'
        },
        {
            id: 'active',
            label: 'Đang Hoạt Động (≤ 7 Ngày)',
            icon: 'green',
            svg: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline></svg>',
            sub: 'Tài khoản có tương tác thường xuyên'
        },
        {
            id: 'abandoned',
            label: 'Cần Rà Soát (> 14 Ngày)',
            icon: 'orange',
            svg: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>',
            sub: 'Không phát sinh đăng nhập'
        }
    ],
    tiers: {
        active: { label: 'Tích cực', class: 'active', hint: 'Đăng nhập gần đây' },
        moderate: { label: 'Thỉnh thoảng', class: 'moderate', hint: 'Trong 2 tuần' },
        idle: { label: 'Ít dùng', class: 'idle', hint: '2 - 4 tuần' },
        abandoned: { label: 'Cần lưu ý', class: 'abandoned', hint: '> 14 ngày hoặc chưa từng đăng nhập' }
    },
    roles: {
        admin: 'Quản trị viên (Global Admin)',
        member: 'Thành viên tổ chức'
    },
    relativeTime: {
        justNow: 'Vừa xong',
        minutesAgo: 'phút trước',
        hoursAgo: 'giờ trước',
        yesterday: 'Hôm qua',
        daysAgo: 'ngày trước',
        never: 'Chưa từng đăng nhập'
    },
    table: {
        showingPrefix: 'Đang hiển thị',
        accounts: 'tài khoản',
        noResults: 'Không tìm thấy tài khoản nào phù hợp với điều kiện lọc.',
        serverError: 'Không thể kết nối đến Microsoft Graph API. Vui lòng kiểm tra lại dịch vụ.'
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
        active7dCount: 0,
        abandonedCount: 0
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
// Initialization & Event Listeners
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
    // Manual refresh button
    const refreshBtn = document.getElementById('btnManualRefresh');
    if (refreshBtn) {
        refreshBtn.addEventListener('click', () => fetchData(true));
    }

    // Auto-refresh interval change
    const intervalSelect = document.getElementById('autoRefreshInterval');
    if (intervalSelect) {
        intervalSelect.addEventListener('change', setupAutoRefresh);
    }

    // Segmented control filters
    const segmentContainer = document.getElementById('segmentedFilter');
    if (segmentContainer) {
        segmentContainer.addEventListener('click', (e) => {
            const btn = e.target.closest('.sf-segment');
            if (!btn) return;
            const targetFilter = btn.getAttribute('data-filter');
            if (targetFilter) AppState.setFilter(targetFilter);
        });
    }

    // Search input
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

    // Detail modal interactions
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

    // Delegate table clicks for Detail Sheet
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
// Presentation Renderers (Apple Clean UI)
// ==========================================================================
function updateMetricsValues() {
    const valTotal = document.getElementById('val-total');
    const valAdmin = document.getElementById('val-admin');
    const valActive = document.getElementById('val-active');
    const valAbandoned = document.getElementById('val-abandoned');

    if (valTotal) valTotal.textContent = AppState.stats.total ?? 0;
    if (valAdmin) valAdmin.textContent = AppState.stats.adminCount ?? 0;
    if (valActive) valActive.textContent = AppState.stats.active7dCount ?? 0;
    if (valAbandoned) valAbandoned.textContent = AppState.stats.abandonedCount ?? 0;
}

function updateSegmentCounts() {
    const total = AppState.users.length;
    const active = AppState.users.filter(u => u.diffDays !== null && u.diffDays <= 7).length;
    const admin = AppState.users.filter(u => u.isAdmin).length;
    const abandoned = AppState.users.filter(u => u.diffDays === null || u.diffDays > 14).length;

    const cAll = document.getElementById('countAll');
    const cActive = document.getElementById('countActive');
    const cAdmin = document.getElementById('countAdmin');
    const cAbandoned = document.getElementById('countAbandoned');

    if (cAll) cAll.textContent = total;
    if (cActive) cActive.textContent = active;
    if (cAdmin) cAdmin.textContent = admin;
    if (cAbandoned) cAbandoned.textContent = abandoned;
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

    // Filter Logic
    const query = AppState.searchQuery;
    const filter = AppState.filter;

    const filtered = AppState.users.filter(user => {
        // Search Matching
        if (query) {
            const name = (user.displayName || '').toLowerCase();
            const upn = (user.userPrincipalName || '').toLowerCase();
            const dept = (user.department || '').toLowerCase();
            const job = (user.jobTitle || '').toLowerCase();
            if (!name.includes(query) && !upn.includes(query) && !dept.includes(query) && !job.includes(query)) {
                return false;
            }
        }

        // Segment Filter
        if (filter === 'all') return true;
        if (filter === 'active') return user.diffDays !== null && user.diffDays <= 7;
        if (filter === 'admin') return user.isAdmin === true;
        if (filter === 'abandoned') return user.diffDays === null || user.diffDays > 14;

        return true;
    });

    if (summaryText) {
        summaryText.textContent = `${I18N.table.showingPrefix} ${filtered.length} / ${AppState.users.length} ${I18N.table.accounts}`;
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
        const timeFormatted = formatRelativeTime(user.lastSignIn);
        const tier = I18N.tiers[user.usageClass] || I18N.tiers.abandoned;

        const roleBadge = user.isAdmin
            ? `<span class="sf-badge admin" title="${escapeHtml(I18N.roles.admin)}">Quản trị viên</span>`
            : `<span class="sf-badge member">Thành viên</span>`;

        return `
            <tr>
                <td>
                    <div class="sf-user-cell">
                        <div class="sf-avatar ${user.isAdmin ? 'admin' : ''}" aria-hidden="true">${initials}</div>
                        <div class="sf-user-meta">
                            <span class="sf-user-name">${escapeHtml(user.displayName)}</span>
                            <span class="sf-user-role-sub">${escapeHtml(user.jobTitle || 'Chưa đặt chức danh')}</span>
                        </div>
                    </div>
                </td>
                <td>
                    <span class="sf-code-pill">${escapeHtml(user.userPrincipalName)}</span>
                </td>
                <td>
                    ${roleBadge}
                </td>
                <td>
                    <span class="sf-badge ${tier.class}" title="${escapeHtml(tier.hint)}">
                        ${tier.label}
                    </span>
                </td>
                <td>
                    <span title="${user.lastSignIn ? new Date(user.lastSignIn).toLocaleString('vi-VN') : ''}">
                        ${timeFormatted}
                    </span>
                </td>
                <td>
                    <button type="button" class="sf-row-btn" data-view-id="${user.id}">Xem</button>
                </td>
            </tr>
        `;
    }).join('');
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
    el.textContent = date.toLocaleTimeString('vi-VN');
    el.setAttribute('datetime', date.toISOString());
}

// ==========================================================================
// Modal Sheet (Apple Detail View)
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
    if (sub) sub.textContent = user.userPrincipalName;

    const tier = I18N.tiers[user.usageClass] || I18N.tiers.abandoned;
    const exactDate = user.lastSignIn ? new Date(user.lastSignIn).toLocaleString('vi-VN') : I18N.relativeTime.never;

    container.innerHTML = `
        <div class="sf-detail-row">
            <span class="sf-detail-label">Vai trò quản trị:</span>
            <span class="sf-detail-val">${user.isAdmin ? I18N.roles.admin : I18N.roles.member}</span>
        </div>
        <div class="sf-detail-row">
            <span class="sf-detail-label">Phòng ban / Bộ phận:</span>
            <span class="sf-detail-val">${escapeHtml(user.department || 'Chưa phân bổ')}</span>
        </div>
        <div class="sf-detail-row">
            <span class="sf-detail-label">Chức danh công vụ:</span>
            <span class="sf-detail-val">${escapeHtml(user.jobTitle || 'Nhân viên')}</span>
        </div>
        <div class="sf-detail-row">
            <span class="sf-detail-label">Tình trạng tài khoản:</span>
            <span class="sf-detail-val">${user.accountEnabled ? '🟢 Đang hoạt động' : '🔴 Bị vô hiệu'}</span>
        </div>
        <div class="sf-detail-row">
            <span class="sf-detail-label">Lần đăng nhập cuối:</span>
            <span class="sf-detail-val">${exactDate}</span>
        </div>
        <div class="sf-detail-row">
            <span class="sf-detail-label">Mức độ tương tác:</span>
            <span class="sf-badge ${tier.class}">${tier.label}</span>
        </div>
        <div class="sf-detail-row">
            <span class="sf-detail-label">Mã định danh Object ID:</span>
            <span class="sf-code-pill" style="font-size: 0.72rem;">${user.id}</span>
        </div>
    `;

    modal.classList.add('open');
    modal.setAttribute('aria-hidden', 'false');
}

// ==========================================================================
// Utilities
// ==========================================================================
function computeInitials(name) {
    if (!name) return 'U';
    const parts = name.trim().split(/\s+/);
    if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

function formatRelativeTime(isoString) {
    if (!isoString) return `<span style="color: var(--sf-label-tertiary);">${I18N.relativeTime.never}</span>`;
    const date = new Date(isoString);
    const now = new Date();
    const diffSec = Math.floor((now.getTime() - date.getTime()) / 1000);

    if (diffSec < 60) return I18N.relativeTime.justNow;
    if (diffSec < 3600) return `${Math.floor(diffSec / 60)} ${I18N.relativeTime.minutesAgo}`;
    if (diffSec < 86400) return `${Math.floor(diffSec / 3600)} ${I18N.relativeTime.hoursAgo}`;
    const days = Math.floor(diffSec / 86400);
    if (days === 1) return I18N.relativeTime.yesterday;
    if (days < 30) return `${days} ${I18N.relativeTime.daysAgo}`;
    return date.toLocaleDateString('vi-VN');
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
        alert('Không có dữ liệu để xuất file.');
        return;
    }

    const headers = ['Họ và Tên', 'Tài khoản UPN', 'Chức danh', 'Phòng ban', 'Vai trò', 'Lần đăng nhập cuối', 'Đánh giá'];
    const rows = AppState.users.map(u => [
        `"${(u.displayName || '').replace(/"/g, '""')}"`,
        `"${(u.userPrincipalName || '').replace(/"/g, '""')}"`,
        `"${(u.jobTitle || '').replace(/"/g, '""')}"`,
        `"${(u.department || '').replace(/"/g, '""')}"`,
        `"${u.isAdmin ? 'Global Administrator' : 'Member'}"`,
        `"${u.lastSignIn ? new Date(u.lastSignIn).toLocaleString('vi-VN') : 'Never'}"`,
        `"${u.usageTier || ''}"`
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
