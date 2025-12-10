# Azure Deployment Guide for NetSuite Excel Add-in

This guide covers deploying the backend to **Azure Functions** and frontend to **Azure Storage Static Website**.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Backend: Azure Functions](#backend-azure-functions)
- [Frontend: Azure Storage](#frontend-azure-storage)
- [Post-Deployment Configuration](#post-deployment-configuration)
- [Monitoring & Troubleshooting](#monitoring--troubleshooting)
- [Cost Overview](#cost-overview)

---

## Prerequisites

### 1. Azure Account
- Active Azure subscription
- Sufficient permissions to create resources

### 2. NetSuite OAuth Credentials

| Credential | Where to Find |
|------------|---------------|
| **Account ID** | Setup → Company → Company Information |
| **Consumer Key** | Setup → Integration → Manage Integrations |
| **Consumer Secret** | (Shown once when creating Integration) |
| **Token ID** | Setup → Users/Roles → Access Tokens |
| **Token Secret** | (Shown once when creating Token) |

### 3. Tools

```bash
# Install Azure CLI
brew install azure-cli

# Install Azure Functions Core Tools
brew install azure-functions-core-tools@4

# Login to Azure
az login
```

---

## Backend: Azure Functions

Azure Functions with **Consumption Plan** provides a cost-effective serverless solution (~$0/month for low usage).

### Deploy Backend

```bash
cd backend/azure-functions
./deploy-functions.sh
```

### What Gets Created

| Resource | Name | Purpose |
|----------|------|---------|
| Resource Group | `netsuite-excel-func-rg` | Container for resources |
| Storage Account | `netsuiteexcelstor` | Required for Functions |
| Function App | `netsuite-excel-func` | API backend |

### API Endpoints

After deployment, your API will be available at:

| Endpoint | URL |
|----------|-----|
| Health | `https://netsuite-excel-func.azurewebsites.net/health` |
| Balance | `https://netsuite-excel-func.azurewebsites.net/api/balance` |
| Budget | `https://netsuite-excel-func.azurewebsites.net/api/budget` |
| Account Name | `https://netsuite-excel-func.azurewebsites.net/api/account/name` |
| Account Type | `https://netsuite-excel-func.azurewebsites.net/api/account/type` |
| List Accounts | `https://netsuite-excel-func.azurewebsites.net/api/accounts` |
| List Periods | `https://netsuite-excel-func.azurewebsites.net/api/periods` |
| SuiteQL | `https://netsuite-excel-func.azurewebsites.net/api/suiteql` |

---

## Frontend: Azure Storage

Azure Storage Static Website provides low-cost hosting for the Excel Add-in frontend.

### Deploy Frontend

```bash
cd docs
./deploy-frontend.sh
```

### What Gets Created

| Resource | Name | Purpose |
|----------|------|---------|
| Storage Account | `netsuiteexcelweb` | Static website hosting |

### Frontend URLs

| Page | URL |
|------|-----|
| Home | `https://netsuiteexcelweb.z13.web.core.windows.net/` |
| Functions | `https://netsuiteexcelweb.z13.web.core.windows.net/functions.html` |
| Taskpane | `https://netsuiteexcelweb.z13.web.core.windows.net/taskpane.html` |

---

## Post-Deployment Configuration

### Configure NetSuite Credentials

After deploying the backend, configure your NetSuite credentials:

1. Open [Azure Portal](https://portal.azure.com)
2. Navigate to: **Function App** → `netsuite-excel-func` → **Settings** → **Environment variables**
3. Update these variables:

| Variable | Value |
|----------|-------|
| `NETSUITE_ACCOUNT_ID` | Your NetSuite account ID |
| `NETSUITE_CONSUMER_KEY` | Your consumer key |
| `NETSUITE_CONSUMER_SECRET` | Your consumer secret |
| `NETSUITE_TOKEN_ID` | Your token ID |
| `NETSUITE_TOKEN_SECRET` | Your token secret |

4. Click **Apply** → **Confirm** to save and restart

### Verify Deployment

```bash
# Test health endpoint
curl https://netsuite-excel-func.azurewebsites.net/health

# Test frontend
curl https://netsuiteexcelweb.z13.web.core.windows.net/functions.json
```

---

## Monitoring & Troubleshooting

### View Logs

```bash
# Stream live logs
func azure functionapp logstream netsuite-excel-func
```

Or in Azure Portal: **Function App** → **Monitoring** → **Log stream**

### Common Issues

| Issue | Solution |
|-------|----------|
| 500 errors | Check NetSuite credentials in Environment variables |
| CORS errors | Verify CORS is configured in Function App settings |
| Function not starting | Check Application Insights for detailed errors |

---

## Cost Overview

| Service | Monthly Cost |
|---------|-------------|
| **Azure Functions (Consumption)** | ~$0 (1M free executions/month) |
| **Azure Storage (Frontend)** | ~$0.02/GB |
| **Azure Storage (Functions)** | ~$0.02/GB |
| **Total** | **~$0-1/month** |

---

## Useful Commands

```bash
# Redeploy backend
cd backend/azure-functions && ./deploy-functions.sh

# Redeploy frontend  
cd docs && ./deploy-frontend.sh

# View backend logs
func azure functionapp logstream netsuite-excel-func

# Delete all resources
az group delete --name netsuite-excel-func-rg --yes --no-wait
```

---

## Support

For issues:
1. Check the [Troubleshooting](#monitoring--troubleshooting) section
2. Review Azure Function logs
3. Verify NetSuite credentials and permissions
