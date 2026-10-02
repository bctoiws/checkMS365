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

// Cache token in memory
let cachedToken = null;
let tokenExpiry = 0;

/**
 * Get OAuth2 Access Token using Client Credentials Flow
 */
async function getAccessToken() {
    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;

    if (!tenantId || !clientId || !clientSecret) {
        throw new Error('Chưa cấu hình TENANT_ID, CLIENT_ID hoặc CLIENT_SECRET trong Environment Variables');
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
 * API Health & Configuration check
 */
app.get('/api/config-status', (req, res) => {
    const hasTenant = Boolean(process.env.TENANT_ID);
    const hasClient = Boolean(process.env.CLIENT_ID);
    const hasSecret = Boolean(process.env.CLIENT_SECRET);

    res.json({
        configured: hasTenant && hasClient && hasSecret,
        details: {
            tenantIdConfigured: hasTenant,
            clientIdConfigured: hasClient,
            clientSecretConfigured: hasSecret
        }
    });
});

/**
 * API: Fetch User List with Activity, Admin Roles & Realtime Status
 */
app.get('/api/users/activity', async (req, res) => {
    try {
        const token = await getAccessToken();
        const headers = { Authorization: `Bearer ${token}` };

        // 1. Fetch Users + SignInActivity
        const usersUrl = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,jobTitle,department,userType,accountEnabled,createdDateTime,signInActivity&$top=100';
        const usersRes = await axios.get(usersUrl, { headers, timeout: 15000 });
        const rawUsers = usersRes.data.value || [];

        // 2. Fetch Directory Roles to detect Admins
        let adminUserIds = new Set();
        try {
            const rolesRes = await axios.get('https://graph.microsoft.com/v1.0/directoryRoles', { headers, timeout: 10000 });
            const globalAdminRole = (rolesRes.data.value || []).find(r => r.displayName === 'Global Administrator' || r.displayName === 'Company Administrator');
            if (globalAdminRole) {
                const membersRes = await axios.get(`https://graph.microsoft.com/v1.0/directoryRoles/${globalAdminRole.id}/members`, { headers, timeout: 10000 });
                (membersRes.data.value || []).forEach(m => adminUserIds.add(m.id));
            }
        } catch (rErr) {
            console.warn('Role fetch warning:', rErr.message);
        }

        // 3. Process & Format Activity
        const nowMs = Date.now();
        const processedUsers = rawUsers.map(user => {
            const isAdmin = adminUserIds.has(user.id);

            // Extract sign-in times
            const lastInteractive = user.signInActivity?.lastSignInDateTime ? new Date(user.signInActivity.lastSignInDateTime).getTime() : null;
            const lastNonInteractive = user.signInActivity?.lastNonInteractiveSignInDateTime ? new Date(user.signInActivity.lastNonInteractiveSignInDateTime).getTime() : null;

            // Pick latest activity time
            let latestActivityTime = null;
            if (lastInteractive && lastNonInteractive) {
                latestActivityTime = Math.max(lastInteractive, lastNonInteractive);
            } else {
                latestActivityTime = lastInteractive || lastNonInteractive;
            }

            // Calculate usage tier
            let usageTier = 'Chưa hoạt động';
            let usageClass = 'abandoned';
            let diffDays = null;

            if (latestActivityTime) {
                diffDays = Math.floor((nowMs - latestActivityTime) / (1000 * 60 * 60 * 24));
                if (diffDays <= 3) {
                    usageTier = 'Rất tích cực';
                    usageClass = 'active';
                } else if (diffDays <= 14) {
                    usageTier = 'Thỉnh thoảng';
                    usageClass = 'moderate';
                } else if (diffDays <= 30) {
                    usageTier = 'Ít sử dụng';
                    usageClass = 'idle';
                } else {
                    usageTier = 'Có nguy cơ bỏ hoang';
                    usageClass = 'abandoned';
                }
            }

            return {
                id: user.id,
                displayName: user.displayName || 'Chưa đặt tên',
                userPrincipalName: user.userPrincipalName,
                department: user.department || 'Chưa phân ban',
                jobTitle: user.jobTitle || (isAdmin ? 'Quản trị viên hệ thống' : 'Nhân viên'),
                isAdmin,
                accountEnabled: user.accountEnabled,
                lastSignIn: latestActivityTime ? new Date(latestActivityTime).toISOString() : null,
                diffDays,
                usageTier,
                usageClass
            };
        });

        // 4. Calculate Summary Statistics
        const total = processedUsers.length;
        const adminCount = processedUsers.filter(u => u.isAdmin).length;
        const active7dCount = processedUsers.filter(u => u.diffDays !== null && u.diffDays <= 7).length;
        const abandonedCount = processedUsers.filter(u => u.diffDays === null || u.diffDays > 14).length;

        res.json({
            success: true,
            timestamp: new Date().toISOString(),
            stats: {
                total,
                adminCount,
                active7dCount,
                abandonedCount
            },
            users: processedUsers
        });

    } catch (err) {
        console.error('Lỗi truy vấn Microsoft Graph:', err.response?.data || err.message);
        res.status(500).json({
            success: false,
            message: err.response?.data?.error?.message || err.message
        });
    }
});

// Fallback to index.html
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
    console.log(`[M365 Monitor] Server running on port ${PORT}`);
});
