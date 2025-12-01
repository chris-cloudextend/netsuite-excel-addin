# NetSuite Excel Formulas

Excel add-in that provides custom formulas for querying NetSuite data directly in Excel.

## Features

- **NS.GLATITLE(account)** - Get GL account name
- **NS.GLABAL(account, fromPeriod, toPeriod, [subsidiary], [department], [location], [class])** - Get GL account balance
- **NS.GLABUD(account, fromPeriod, toPeriod, [subsidiary], [department], [location], [class])** - Get GL account budget

## Architecture

- **Frontend**: Excel add-in (Office.js) hosted on GitHub Pages
- **Backend**: Flask server for NetSuite OAuth & SuiteQL queries
- **Proxy**: Cloudflare Worker for stable URL (no manifest updates needed)
- **Deployment**: 
  - Add-in files: GitHub Pages (`docs/` folder)
  - Backend: Local Flask server exposed via Cloudflare Tunnel

## Quick Start

### 1. Backend Setup

```bash
cd backend
pip3 install -r requirements.txt

# Edit netsuite_config.json with your credentials
python3 server.py
```

### 2. Cloudflare Tunnel

```bash
cloudflared tunnel --url http://localhost:5002
# Update TUNNEL_URL in Cloudflare Worker with the new tunnel URL
```

### 3. Excel Installation

1. Copy `excel-addin/manifest-claude.xml` to:
   - Mac: `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
   - Windows: `%USERPROFILE%\AppData\Local\Microsoft\Office\16.0\Wef\`

2. In Excel: Insert → My Add-ins → Shared Folder
3. Select the manifest file

## Configuration

### NetSuite Account Setup

Create `backend/netsuite_config.json`:

```json
{
  "account_id": "YOUR_ACCOUNT_ID",
  "consumer_key": "YOUR_CONSUMER_KEY",
  "consumer_secret": "YOUR_CONSUMER_SECRET",
  "token_id": "YOUR_TOKEN_ID",
  "token_secret": "YOUR_TOKEN_SECRET",
  "realm": "YOUR_ACCOUNT_ID"
}
```

### Cloudflare Worker

The Cloudflare Worker at `https://netsuite-proxy.chris-corcoran.workers.dev/` proxies requests to your Cloudflare Tunnel, providing a stable URL so the manifest doesn't need to be updated when the tunnel restarts.

Update the `TUNNEL_URL` variable in the worker code when you restart your tunnel.

## Development

### File Structure

```
├── backend/
│   ├── server.py              # Flask backend
│   ├── requirements.txt       # Python dependencies
│   └── netsuite_config.json   # NetSuite credentials (gitignored)
├── docs/                      # GitHub Pages (add-in files)
│   ├── functions.js           # Custom function implementations
│   ├── functions.json         # Function metadata
│   ├── functions.html         # Functions runtime page
│   ├── taskpane.html          # Task pane UI
│   └── commands.html          # Ribbon commands
├── excel-addin/
│   └── manifest-claude.xml    # Add-in manifest
└── clear-excel-cache.sh       # Utility to clear Excel cache
```

### Making Changes

1. Edit files in `docs/` folder
2. Commit and push to GitHub
3. Wait 2-3 minutes for GitHub Pages to deploy
4. Bump manifest version to force Excel to reload:
   - Update `<Version>` in `manifest-claude.xml`
   - Update `?v=XXXX` in all URLs
5. Reload add-in in Excel

## Troubleshooting

### Clear Excel Cache

```bash
./clear-excel-cache.sh
```

Or manually:
```bash
rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Application\ Support/Microsoft/Office/16.0/Wef/
```

### Check Backend Logs

Backend logs to console. Check for errors related to NetSuite authentication or SuiteQL queries.

### Verify Tunnel

Visit your Cloudflare Worker URL to ensure the tunnel is running and responding.

## License

Private project - All rights reserved
