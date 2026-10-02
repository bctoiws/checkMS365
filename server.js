const express = require('express');
const cors = require('cors');
const axios = require('axios');
const path = require('path');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// In-memory token & sign-in cache
let cachedToken = null;
let tokenExpiry = 0;
let cachedSignIns = null;
let signInsExpiry = 0;

/**
 * Acquire OAuth2 Token via Client Credentials
 */
async function getAccessToken() {
    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;

    if (!tenantId || !clientId || !clientSecret) {
        throw new Error('Missing TENANT_ID, CLIENT_ID or CLIENT_SECRET in environment variables.');
    }

    const now = Math.floor(Date.now() / 1000);
    if (cachedToken && tokenExpiry > now + 120) {
        return cachedToken;
    }

    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams();
    params.append('client_id', clientId);
    params.append('scope', 'https://graph.microsoft.com/.default');
    params.append('client_secret', clientSecret);
    params.append('grant_type', 'client_credentials');

    const res = await axios.post(tokenUrl, params.toString(), {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        timeout: 10000
    });

    cachedToken = res.data.access_token;
    tokenExpiry = now + (res.data.expires_in || 3600);
    return cachedToken;
}

/**
 * Return genuine Microsoft application name without artificial modification
 */
function cleanApplicationName(rawName, resourceName) {
    if (rawName && rawName.trim()) {
        return rawName.trim();
    }
    if (resourceName && resourceName.trim()) {
        return resourceName.trim();
    }
    return 'Microsoft 365 Service';
}

/**
 * =========================================================================
 * Microsoft 365 E5 Developer Auto-Renewal Keep-Alive Engine
 * Periodically executes authentic Graph API requests across multiple services
 * simulating genuine developer development activity to ensure 90-day renewals.
 * =========================================================================
 */
const E5RenewalEngine = {
    enabled: true,
    intervalMinutes: parseInt(process.env.E5_INTERVAL_MINUTES || '60', 10), // Every 60 minutes default
    lastRunTime: null,
    nextRunTime: null,
    totalCycles: 0,
    callsToday: 0,
    lastDayTracked: new Date().getUTCDate(),
    recentCycles: [],
    timer: null,

    // Target Graph API endpoints simulating active developer activity
    targetEndpoints: [
        { name: 'Tenant Organization', path: '/v1.0/organization', weight: 'Core Tenant' },
        { name: 'Subscription Licenses & SKUs', path: '/v1.0/subscribedSkus', weight: 'License Allocation' },
        { name: 'Directory User Roster', path: '/v1.0/users?$top=25&$select=id,displayName,userPrincipalName', weight: 'Identity' },
        { name: 'Directory Groups', path: '/v1.0/groups?$top=15&$select=id,displayName', weight: 'Collaboration' },
        { name: 'Directory Administrative Roles', path: '/v1.0/directoryRoles', weight: 'Governance' },
        { name: 'Directory Audit Logs', path: '/v1.0/auditLogs/directoryAudits?$top=10', weight: 'Security Compliance' },
        { name: 'Enterprise Service Principals', path: '/v1.0/servicePrincipals?$top=10&$select=id,appDisplayName', weight: 'App Registration' }
    ],

    async executeCycle(triggeredBy = 'automated_schedule') {
        const currentDay = new Date().getUTCDate();
        if (currentDay !== this.lastDayTracked) {
            this.callsToday = 0;
            this.lastDayTracked = currentDay;
        }

        const cycleStart = Date.now();
        const cycleId = 'cycle_' + cycleStart;
        const results = [];

        console.log(`[E5 Engine] Starting keep-alive renewal cycle (${triggeredBy})...`);

        try {
            const token = await getAccessToken();
            const headers = { Authorization: `Bearer ${token}` };

            for (const ep of this.targetEndpoints) {
                const epStart = Date.now();
                try {
                    const res = await axios.get(`https://graph.microsoft.com${ep.path}`, {
                        headers,
                        timeout: 15000
                    });
                    const duration = Date.now() - epStart;
                    this.callsToday++;
                    results.push({
                        name: ep.name,
                        path: ep.path,
                        weight: ep.weight,
                        status: res.status,
                        durationMs: duration,
                        success: true,
                        itemCount: Array.isArray(res.data?.value) ? res.data.value.length : 1
                    });
                } catch (err) {
                    results.push({
                        name: ep.name,
                        path: ep.path,
                        weight: ep.weight,
                        status: err.response?.status || 500,
                        durationMs: Date.now() - epStart,
                        success: false,
                        error: err.response?.data?.error?.message || err.message
                    });
                }
            }

            this.totalCycles++;
            this.lastRunTime = new Date().toISOString();
            this.nextRunTime = new Date(Date.now() + this.intervalMinutes * 60 * 1000).toISOString();

            const cycleRecord = {
                id: cycleId,
                timestamp: this.lastRunTime,
                triggeredBy,
                totalEndpoints: this.targetEndpoints.length,
                successCount: results.filter(r => r.success).length,
                totalDurationMs: Date.now() - cycleStart,
                results
            };

            this.recentCycles.unshift(cycleRecord);
            if (this.recentCycles.length > 25) {
                this.recentCycles.pop();
            }

            console.log(`[E5 Engine] Cycle complete: ${cycleRecord.successCount}/${cycleRecord.totalEndpoints} successful (${cycleRecord.totalDurationMs}ms). Next run: ${this.nextRunTime}`);
            return cycleRecord;

        } catch (authErr) {
            console.error('[E5 Engine] Auth failure during cycle:', authErr.message);
            throw authErr;
        }
    },

    start() {
        if (!this.enabled) return;
        console.log(`[E5 Engine] Initialized: Recurring cycle every ${this.intervalMinutes} minutes.`);

        // Schedule first run 5s after boot
        setTimeout(() => {
            this.executeCycle('startup_init').catch(e => console.error('[E5 Engine] Startup run error:', e.message));
        }, 5000);

        // Continuous recurring timer
        this.nextRunTime = new Date(Date.now() + this.intervalMinutes * 60 * 1000).toISOString();
        this.timer = setInterval(() => {
            this.executeCycle('automated_schedule').catch(e => console.error('[E5 Engine] Recurring run error:', e.message));
        }, this.intervalMinutes * 60 * 1000);
    }
};

/**
 * E5 Engine APIs
 */
app.get('/api/e5-engine/status', (req, res) => {
    res.json({
        success: true,
        enabled: E5RenewalEngine.enabled,
        intervalMinutes: E5RenewalEngine.intervalMinutes,
        lastRunTime: E5RenewalEngine.lastRunTime,
        nextRunTime: E5RenewalEngine.nextRunTime,
        totalCycles: E5RenewalEngine.totalCycles,
        callsToday: E5RenewalEngine.callsToday,
        targetEndpoints: E5RenewalEngine.targetEndpoints.map(e => ({ name: e.name, path: e.path, weight: e.weight })),
        recentCycles: E5RenewalEngine.recentCycles
    });
});

app.post('/api/e5-engine/trigger', async (req, res) => {
    try {
        const cycle = await E5RenewalEngine.executeCycle('manual_dashboard_trigger');
        res.json({ success: true, cycle });
    } catch (err) {
        res.status(500).json({ success: false, message: err.message });
    }
});

// External Cron Ping Endpoint (For UptimeRobot, cron-job.org, Render Keep-Alive, GitHub Actions)
app.get('/api/keep-alive', async (req, res) => {
    try {
        const cycle = await E5RenewalEngine.executeCycle('external_cron_webhook');
        res.json({
            success: true,
            message: 'E5 Keep-Alive renewal cycle executed successfully',
            timestamp: new Date().toISOString(),
            successEndpoints: `${cycle.successCount}/${cycle.totalEndpoints}`,
            durationMs: cycle.totalDurationMs
        });
    } catch (err) {
        res.status(500).json({ success: false, message: err.message });
    }
});

/**
 * Health & configuration probe
 */
app.get('/api/config-status', (req, res) => {
    res.json({
        configured: Boolean(process.env.TENANT_ID && process.env.CLIENT_ID && process.env.CLIENT_SECRET),
        tenantId: process.env.TENANT_ID ? `${process.env.TENANT_ID.slice(0, 8)}...` : null
    });
});

/**
 * Primary API: Users, Directory Roles, Monthly Access Count & Exact App Usage
 */
app.get('/api/users/activity', async (req, res) => {
    try {
        const token = await getAccessToken();
        const headers = { Authorization: `Bearer ${token}` };

        // 1. Fetch Users
        const usersUrl = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,jobTitle,department,accountEnabled,createdDateTime,signInActivity&$top=100';
        const usersRes = await axios.get(usersUrl, { headers, timeout: 15000 });
        const rawUsers = usersRes.data.value || [];

        // 2. Fetch Directory Roles to detect Global Administrators
        const adminUserIds = new Set();
        try {
            const rolesRes = await axios.get('https://graph.microsoft.com/v1.0/directoryRoles', { headers, timeout: 10000 });
            const globalAdminRole = (rolesRes.data.value || []).find(r => 
                r.displayName === 'Global Administrator' || r.displayName === 'Company Administrator'
            );
            if (globalAdminRole) {
                const membersRes = await axios.get(`https://graph.microsoft.com/v1.0/directoryRoles/${globalAdminRole.id}/members`, { headers, timeout: 10000 });
                (membersRes.data.value || []).forEach(m => adminUserIds.add(m.id));
            }
        } catch (rErr) {
            console.warn('Role inspection warning:', rErr.message);
        }

        // 3. Fetch Sign-in Audit Logs (Cached for 60s to ensure fast response)
        const userAccessMap = {};
        let totalMonthlyAccesses = 0;
        const now = Math.floor(Date.now() / 1000);

        try {
            let signIns = [];
            if (cachedSignIns && signInsExpiry > now) {
                signIns = cachedSignIns;
            } else {
                let currentUrl = 'https://graph.microsoft.com/v1.0/auditLogs/signIns?$top=500';
                const fetched = [];
                while (currentUrl && fetched.length < 2500) {
                    const signInsRes = await axios.get(currentUrl, { headers, timeout: 35000 });
                    const pageData = signInsRes.data.value || [];
                    fetched.push(...pageData);
                    currentUrl = signInsRes.data['@odata.nextLink'] || null;
                }
                signIns = fetched;
                cachedSignIns = signIns;
                signInsExpiry = now + 60; // 60 seconds TTL
            }

            totalMonthlyAccesses = signIns.length;

            signIns.forEach(entry => {
                const upn = (entry.userPrincipalName || '').toLowerCase();
                if (!upn) return;

                if (!userAccessMap[upn]) {
                    userAccessMap[upn] = {
                        monthlyAccessCount: 0,
                        appsSet: new Set(),
                        events: []
                    };
                }

                userAccessMap[upn].monthlyAccessCount++;

                const cleanApp = cleanApplicationName(entry.appDisplayName, entry.resourceDisplayName);
                userAccessMap[upn].appsSet.add(cleanApp);

                if (userAccessMap[upn].events.length < 30) {
                    userAccessMap[upn].events.push({
                        id: entry.id,
                        app: cleanApp,
                        rawApp: entry.appDisplayName || entry.resourceDisplayName || 'Microsoft 365 Service',
                        resource: entry.resourceDisplayName || '',
                        timestamp: entry.createdDateTime,
                        clientApp: entry.clientAppUsed || 'Desktop / Browser',
                        os: entry.deviceDetail?.operatingSystem || 'Unknown OS',
                        browser: entry.deviceDetail?.browser || 'Browser',
                        location: entry.location?.city ? `${entry.location.city}, ${entry.location.countryOrRegion}` : (entry.location?.countryOrRegion || 'VN'),
                        ipAddress: entry.ipAddress || '',
                        status: entry.status?.errorCode === 0 ? 'Success' : 'Failed'
                    });
                }
            });
        } catch (sErr) {
            console.warn('AuditLogs sign-in inspection warning:', sErr.message);
        }

        // 4. Assemble Data for Each User
        const nowMs = Date.now();
        const processedUsers = rawUsers.map(user => {
            const upn = (user.userPrincipalName || '').toLowerCase();
            const isAdmin = adminUserIds.has(user.id);
            const accessData = userAccessMap[upn] || { monthlyAccessCount: 0, appsSet: new Set(), events: [] };

            // Determine latest sign-in timestamp (from sign-ins or signInActivity)
            let latestTimestamp = null;
            if (accessData.events.length > 0) {
                latestTimestamp = accessData.events[0].timestamp;
            } else if (user.signInActivity?.lastSignInDateTime) {
                latestTimestamp = user.signInActivity.lastSignInDateTime;
            }

            const latestMs = latestTimestamp ? new Date(latestTimestamp).getTime() : null;
            let diffDays = null;
            if (latestMs) {
                diffDays = Math.floor((nowMs - latestMs) / (1000 * 60 * 60 * 24));
            }

            // Usage categorization
            let usageCategory = 'inactive';
            if (accessData.monthlyAccessCount >= 10 || (diffDays !== null && diffDays <= 3)) {
                usageCategory = 'frequent';
            } else if (accessData.monthlyAccessCount > 0 || (diffDays !== null && diffDays <= 14)) {
                usageCategory = 'moderate';
            } else {
                usageCategory = 'dormant';
            }

            const appsList = Array.from(accessData.appsSet);
            const latestApp = accessData.events.length > 0 ? accessData.events[0].app : (appsList[0] || 'None recorded');

            return {
                id: user.id,
                displayName: user.displayName || 'Unnamed User',
                userPrincipalName: user.userPrincipalName,
                department: user.department || 'General',
                jobTitle: user.jobTitle || (isAdmin ? 'Global Administrator' : 'Staff Member'),
                isAdmin,
                accountEnabled: user.accountEnabled !== false,
                monthlyAccessCount: accessData.monthlyAccessCount,
                appsUsed: appsList,
                latestApp,
                latestTimestamp,
                diffDays,
                usageCategory,
                history: accessData.events
            };
        });

        // 5. Aggregate Summary Statistics
        const total = processedUsers.length;
        const adminCount = processedUsers.filter(u => u.isAdmin).length;
        const activeUsersCount = processedUsers.filter(u => u.monthlyAccessCount > 0 || (u.diffDays !== null && u.diffDays <= 7)).length;
        const dormantCount = processedUsers.filter(u => u.monthlyAccessCount === 0 && (u.diffDays === null || u.diffDays > 14)).length;

        res.json({
            success: true,
            timestamp: new Date().toISOString(),
            stats: {
                total,
                adminCount,
                activeUsersCount,
                dormantCount,
                totalMonthlyAccesses
            },
            users: processedUsers
        });

    } catch (err) {
        console.error('Microsoft Graph processing error:', err.response?.data || err.message);
        res.status(500).json({
            success: false,
            message: err.response?.data?.error?.message || err.message
        });
    }
});

// Default route
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
    console.log(`[M365 Monitor] Server running on port ${PORT}`);
    // Boot the automated E5 renewal keep-alive engine
    E5RenewalEngine.start();
});

