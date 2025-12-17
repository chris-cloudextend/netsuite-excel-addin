# XAVI for NetSuite - Engineering Handoff Documentation

## Overview

XAVI is an Excel Add-in that provides financial reporting formulas for NetSuite. Users can type formulas like `=XAVI.BALANCE("4010", "Jan 2025", "Dec 2025")` directly in Excel cells to pull live data from their NetSuite account.

---

## Current Architecture

### Why Public GitHub?

The project currently uses a **public GitHub repository** for one primary reason:

**GitHub Pages requires public repos on the free tier.** The Excel Add-in manifest points to GitHub Pages URLs for hosting the static frontend files (HTML, JS, CSS). This was the fastest path to a working prototype.

```
Excel Add-in → GitHub Pages (static files) → Cloudflare Worker (proxy) → Cloudflare Tunnel → Local Backend → NetSuite REST API
```

### Component Breakdown

| Component | Current Location | Purpose |
|-----------|-----------------|---------|
| **Excel Manifest** | `excel-addin/manifest-claude.xml` | Tells Excel where to load the add-in from |
| **Frontend (Taskpane)** | `docs/taskpane.html` | The sidebar UI users interact with |
| **Custom Functions** | `docs/functions.js` + `docs/functions.json` | Excel formulas (XAVI.BALANCE, etc.) |
| **Backend Server** | `backend/server.py` | Flask server that queries NetSuite |
| **Cloudflare Worker** | `CLOUDFLARE-WORKER-CODE.js` | Proxy that routes requests to the tunnel |

### Current Request Flow

```
1. User types =XAVI.BALANCE(...) in Excel
2. Excel calls functions.js
3. functions.js calls Cloudflare Worker (netsuite-proxy.chris-corcoran.workers.dev)
4. Worker proxies to Cloudflare Tunnel (*.trycloudflare.com)
5. Tunnel connects to local Flask backend (localhost:5002)
6. Backend authenticates with NetSuite REST API using OAuth 1.0
7. Response flows back through the chain to Excel
```

### Why Cloudflare Tunnel?

The backend needs to connect to NetSuite using OAuth 1.0 credentials stored in `netsuite_config.json`. During development:
- The backend runs locally on the developer's machine
- Cloudflare Tunnel exposes it to the internet with a temporary URL
- The Cloudflare Worker provides a stable URL that forwards to whatever tunnel is active

This allowed rapid iteration without deploying to a server for every change.

---

## Migration to Private Git + Cloud Hosting

### What Needs to Change

1. **Static File Hosting** - Move from GitHub Pages to:
   - AWS S3 + CloudFront, OR
   - Azure Blob Storage + CDN, OR
   - Any static file hosting service

2. **Backend Hosting** - Deploy Flask backend to:
   - AWS: EC2, ECS, Lambda + API Gateway, or Elastic Beanstalk
   - Azure: App Service, Container Instances, or Functions

3. **Remove Cloudflare Tunnel** - Replace with direct cloud hosting

4. **Update Manifest URLs** - Point to new hosting locations

### Recommended AWS Architecture

```
                                    ┌─────────────────┐
                                    │   CloudFront    │
                                    │   (CDN + HTTPS) │
                                    └────────┬────────┘
                                             │
              ┌──────────────────────────────┼──────────────────────────────┐
              │                              │                              │
              ▼                              ▼                              ▼
    ┌─────────────────┐          ┌─────────────────┐          ┌─────────────────┐
    │   S3 Bucket     │          │  API Gateway    │          │   Cognito       │
    │  (Static Files) │          │  (REST API)     │          │  (CEFI Auth)    │
    │  taskpane.html  │          │                 │          │                 │
    │  functions.js   │          └────────┬────────┘          └─────────────────┘
    └─────────────────┘                   │
                                          ▼
                               ┌─────────────────┐
                               │  Lambda / ECS   │
                               │  (Python Flask) │
                               └────────┬────────┘
                                        │
                                        ▼
                               ┌─────────────────┐
                               │  Secrets Manager│
                               │  (NetSuite creds)│
                               └────────┬────────┘
                                        │
                                        ▼
                               ┌─────────────────┐
                               │  NetSuite REST  │
                               │      API        │
                               └─────────────────┘
```

### Recommended Azure Architecture

```
                                    ┌─────────────────┐
                                    │   Azure CDN     │
                                    │   (HTTPS)       │
                                    └────────┬────────┘
                                             │
              ┌──────────────────────────────┼──────────────────────────────┐
              │                              │                              │
              ▼                              ▼                              ▼
    ┌─────────────────┐          ┌─────────────────┐          ┌─────────────────┐
    │  Blob Storage   │          │  API Management │          │   Azure AD B2C  │
    │  (Static Files) │          │  (REST API)     │          │  (CEFI Auth)    │
    └─────────────────┘          └────────┬────────┘          └─────────────────┘
                                          │
                                          ▼
                               ┌─────────────────┐
                               │  App Service    │
                               │  (Python Flask) │
                               └────────┬────────┘
                                        │
                                        ▼
                               ┌─────────────────┐
                               │  Key Vault      │
                               │  (NetSuite creds)│
                               └─────────────────┘
```

---

## Code Changes Required for Cloud Deployment

### 1. Update Manifest URLs

In `excel-addin/manifest-claude.xml`, replace all GitHub Pages URLs:

```xml
<!-- FROM -->
<SourceLocation DefaultValue="https://chris-cloudextend.github.io/netsuite-excel-addin/taskpane.html"/>

<!-- TO (example for AWS) -->
<SourceLocation DefaultValue="https://d1234567890.cloudfront.net/taskpane.html"/>
```

### 2. Update SERVER_URL in Frontend

In `docs/taskpane.html` and `docs/functions.js`, update the server URL:

```javascript
// FROM
const SERVER_URL = 'https://netsuite-proxy.chris-corcoran.workers.dev';

// TO (example)
const SERVER_URL = 'https://api.xavi.cloudextend.io';
```

### 3. Backend Configuration

The backend (`backend/server.py`) currently reads credentials from `netsuite_config.json`. For cloud deployment:

**Option A: Environment Variables**
```python
ACCOUNT_ID = os.environ.get('NETSUITE_ACCOUNT_ID')
CONSUMER_KEY = os.environ.get('NETSUITE_CONSUMER_KEY')
# etc.
```

**Option B: Secrets Manager (AWS) / Key Vault (Azure)**
```python
import boto3
client = boto3.client('secretsmanager')
secret = client.get_secret_value(SecretId='netsuite-credentials')
```

### 4. Remove Cloudflare Dependencies

- Delete `CLOUDFLARE-WORKER-CODE.js` (no longer needed)
- Remove any references to trycloudflare.com tunnel URLs

---

## Multi-Tenant Architecture (CEFI Login)

### Current State
The backend currently uses a single set of NetSuite credentials configured in `netsuite_config.json`. All users share these credentials.

### Target State
Each user authenticates via CEFI (Celigo's identity platform), and the backend retrieves their NetSuite credentials from a secure store.

### Required Changes

1. **Frontend Authentication Flow**
   ```javascript
   // On add-in load, check if user is authenticated
   async function checkAuth() {
       const token = localStorage.getItem('cefi_token');
       if (!token) {
           // Redirect to CEFI login
           window.location.href = 'https://auth.celigo.com/login?redirect=...';
       }
       // Validate token with backend
       const response = await fetch(`${SERVER_URL}/auth/validate`, {
           headers: { 'Authorization': `Bearer ${token}` }
       });
   }
   ```

2. **Backend Token Validation**
   ```python
   @app.before_request
   def validate_token():
       token = request.headers.get('Authorization')
       # Validate with CEFI
       # Retrieve user's NetSuite credentials from secure store
       # Set credentials for this request
   ```

3. **Credential Storage**
   - Store per-user NetSuite credentials in encrypted database
   - Or use CEFI's credential vault if available
   - Credentials should include: Account ID, Consumer Key/Secret, Token Key/Secret

---

## Files to Review

| File | Description |
|------|-------------|
| `backend/server.py` | Main Flask backend - all NetSuite API calls |
| `backend/netsuite_config.json` | Current credentials (DO NOT COMMIT to public repo) |
| `docs/taskpane.html` | Main UI + JavaScript logic |
| `docs/functions.js` | Excel custom functions implementation |
| `docs/functions.json` | Excel function definitions/metadata |
| `excel-addin/manifest-claude.xml` | Excel add-in manifest with all URLs |
| `DEVELOPER_CHECKLIST.md` | Integration points for adding new formulas |

---

## Security Considerations

1. **Never commit credentials** - `netsuite_config.json` is in `.gitignore`
2. **HTTPS required** - Excel add-ins require HTTPS for all resources
3. **CORS configuration** - Backend must allow requests from Excel's origin
4. **Token expiration** - Implement proper token refresh for CEFI auth
5. **Rate limiting** - Consider adding rate limits to prevent abuse

---

## Testing the Migration

1. Deploy static files to new hosting
2. Deploy backend to cloud
3. Update manifest with new URLs
4. Sideload updated manifest in Excel
5. Test all formulas: BALANCE, BUDGET, NAME, TYPEBALANCE, etc.
6. Test tutorial flow
7. Test drill-down functionality
8. Verify multi-subsidiary support

---

## Questions for Engineering

1. Which cloud provider (AWS or Azure)?
2. How will CEFI credentials be passed to the add-in?
3. Will there be a credential storage service, or should we build one?
4. What's the domain for the production API? (e.g., api.xavi.cloudextend.io)
5. Do we need to support on-premise NetSuite deployments?

---

## Contact

This prototype was developed by the CloudExtend team. For questions about the codebase, refer to:
- `DEVELOPER_CHECKLIST.md` - How to add new formulas
- `DOCUMENTATION.md` - API and feature documentation
- `USER_GUIDE.md` - End-user documentation

---

*Last updated: December 2024*

