# Microsoft 365 Tenant Activity Monitor & E5 Keep-Alive Engine

An enterprise-grade, privacy-centric monitoring dashboard and automated keep-alive engine for Microsoft 365 E5 Developer tenants. Designed following Apple Human Interface Guidelines (HIG) with zero external CDN dependencies, minimal overhead, and 100% authentic Microsoft Graph telemetry.

---

## Features

- **Authentic Application Auditing**: Tracks exact Microsoft applications used by tenant identities (Word, Excel, PowerPoint, Outlook, OneDrive, SharePoint, Teams, Azure Portal, Copilot).
- **Monthly Access Frequencies**: Aggregates 30-day session telemetry across all licensed directory members.
- **E5 Developer Auto-Renewal Engine**: Background keep-alive engine periodically querying 7 core Graph API endpoints (`/organization`, `/subscribedSkus`, `/users`, `/groups`, `/directoryRoles`, `/directoryAudits`, `/servicePrincipals`) to guarantee 90-day subscription renewals without spamming.
- **Apple HIG Light Mode Interface**: Minimalist layout inspired by macOS Activity Monitor and App Store Connect, featuring interactive two-way table sorting (Name A-Z, Accesses High/Low, Recent Time) and zero marketing fluff.
- **Zero Token Leakage**: Credentials are kept exclusively in environment variables and protected by `.gitignore`.

---

## Prerequisites & Microsoft Entra ID Permissions

Configure an App Registration in [Microsoft Entra admin center](https://entra.microsoft.com) with the following Application permissions:
- `Directory.Read.All`
- `User.Read.All`
- `AuditLog.Read.All`

*Make sure to click **Grant admin consent**.*

---

## Local Setup

1. Clone the repository and install dependencies:
   ```bash
   git clone https://github.com/bctoiws/checkMS365.git
   cd checkMS365
   npm install
   ```

2. Copy `.env.example` to `.env` and fill in your tenant credentials:
   ```bash
   cp .env.example .env
   ```

3. Start the application:
   ```bash
   npm start
   ```
   Open `http://localhost:3000` in your browser.

---

## Cloud Deployment (Render.com)

1. Create a new **Web Service** on [Render.com](https://render.com) connected to this GitHub repository.
2. Set configuration:
   - **Environment**: `Node`
   - **Build Command**: `npm install`
   - **Start Command**: `node server.js`
   - **Instance Type**: `Free`
3. Add the following **Environment Variables**:
   - `TENANT_ID`: Your Azure Tenant ID
   - `CLIENT_ID`: Your App Registration Client ID
   - `CLIENT_SECRET`: Your App Registration Client Secret Value
   - `E5_INTERVAL_MINUTES`: `60` (or `120`)
4. Deploy the service.

---

## Automated Cloud Keep-Alive (GitHub Actions)

If deployed on Render's free tier (which sleeps after 15 minutes of inactivity), the built-in GitHub Actions workflow will automatically ping your service every 4 hours to keep the E5 renewal engine pulsing:

1. In your GitHub repository, go to **Settings** > **Secrets and variables** > **Actions**.
2. Click **New repository secret**.
3. Name: `DEPLOYED_APP_URL`
4. Value: `https://your-service-name.onrender.com`
5. The workflow in `.github/workflows/e5-keepalive.yml` will now automatically keep your tenant active 24/7!
