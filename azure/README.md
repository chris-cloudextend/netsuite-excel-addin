# Azure Deployment for XAVI NetSuite Excel Add-in

This folder contains everything needed to deploy XAVI to Azure, **without modifying any existing code**.

## Architecture

```
┌─────────────────────────────────┐
│     Excel Add-in (Browser)      │
│   - Custom Functions (JS)       │
│   - Taskpane UI (HTML/JS)       │
└─────────────┬───────────────────┘
              │
              ▼
┌─────────────────────────────────┐
│  Azure Storage Static Website   │
│  netsuiteexcelweb.z13.web...    │
└─────────────┬───────────────────┘
              │
              ▼
┌─────────────────────────────────┐
│   Azure Container Apps          │
│   netsuite-flask (scale-to-0)   │
│   - All 29 API endpoints        │
└─────────────┬───────────────────┘
              │
              ▼
┌─────────────────────────────────┐
│         NetSuite API            │
└─────────────────────────────────┘
```

## Folder Structure

```
azure/
├── README.md                           # This file
├── manifests/
│   └── manifest-azure.xml              # Excel Add-in manifest
└── scripts/
    ├── deploy-flask-container.sh       # Deploy Flask to Container Apps
    └── deploy-frontend-container.sh    # Deploy frontend to Azure Storage
```

## Quick Start

### Prerequisites

```bash
# Install Azure CLI
brew install azure-cli

# Login to Azure
az login
```

### Step 1: Deploy Flask Backend

```bash
cd azure/scripts
chmod +x deploy-flask-container.sh
./deploy-flask-container.sh
```

### Step 2: Deploy Frontend

```bash
chmod +x deploy-frontend-container.sh
./deploy-frontend-container.sh
```

### Step 3: Configure NetSuite Credentials

```bash
az containerapp update \
    --name netsuite-flask \
    --resource-group netsuite-excel-func-rg \
    --set-env-vars \
        NETSUITE_ACCOUNT_ID=your_account_id \
        NETSUITE_CONSUMER_KEY=your_consumer_key \
        NETSUITE_CONSUMER_SECRET=your_consumer_secret \
        NETSUITE_TOKEN_ID=your_token_id \
        NETSUITE_TOKEN_SECRET=your_token_secret
```

Or configure via Azure Portal → Container Apps → netsuite-flask → Environment variables

### Step 4: Install Excel Add-in

**Mac:**
```bash
cp azure/manifests/manifest-azure.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```

**Windows:**
```
Copy manifest-azure.xml to %USERPROFILE%\AppData\Local\Microsoft\Office\16.0\Wef\
```

Then in Excel: **Insert → My Add-ins → Shared Folder** → Select the manifest

Or: **Insert → My Add-ins → Upload My Add-in** → Browse to manifest-azure.xml

### Step 5: Verify

```bash
curl https://netsuite-flask.ashymeadow-6282430c.eastus.azurecontainerapps.io/health
```

## Azure Resources

| Resource | Name | URL |
|----------|------|-----|
| Resource Group | `netsuite-excel-func-rg` | - |
| Storage (Frontend) | `netsuiteexcelweb` | https://netsuiteexcelweb.z13.web.core.windows.net |
| Container Apps (Backend) | `netsuite-flask` | https://netsuite-flask.ashymeadow-6282430c.eastus.azurecontainerapps.io |

## Cost

| Service | Monthly Cost |
|---------|-------------|
| Container Apps (scale-to-zero) | ~$0-5 (usage-based) |
| Storage (Frontend) | ~$0.02/GB |
| **Total** | **~$0-5/month** for testing |

### Stop to Save Costs

```bash
# Stop (no costs)
az containerapp update --name netsuite-flask --resource-group netsuite-excel-func-rg --min-replicas 0 --max-replicas 0

# Start
az containerapp update --name netsuite-flask --resource-group netsuite-excel-func-rg --min-replicas 0 --max-replicas 3
```

## Troubleshooting

```bash
# Health check
curl https://netsuite-flask.ashymeadow-6282430c.eastus.azurecontainerapps.io/health

# View logs
az containerapp logs show --name netsuite-flask --resource-group netsuite-excel-func-rg --tail 50

# Check status
az containerapp revision list --name netsuite-flask --resource-group netsuite-excel-func-rg -o table
```

## Delete All Resources

```bash
az group delete --name netsuite-excel-func-rg --yes --no-wait
```

---

*Last Updated: December 2025*
