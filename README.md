# XAVI for NetSuite

Excel Add-in providing custom formulas to query NetSuite financial data directly in Excel.

## Quick Start

### 1. Start the Backend Server

```bash
cd backend
pip3 install -r requirements.txt
cp netsuite_config.template.json netsuite_config.json
# Edit netsuite_config.json with your NetSuite credentials
python3 server.py
```

### 2. Start Cloudflare Tunnel

```bash
cloudflared tunnel --url http://localhost:5002
# Copy the tunnel URL and update CLOUDFLARE-WORKER-CODE.js
```

### 3. Install the Excel Add-in

**Mac:**
```bash
cp excel-addin/manifest-claude.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```

**Windows:**
```
Copy manifest-claude.xml to %USERPROFILE%\AppData\Local\Microsoft\Office\16.0\Wef\
```

Then in Excel: **Insert → My Add-ins → Shared Folder** → Select the manifest

## Available Formulas

| Formula | Description |
|---------|-------------|
| `=XAVI.BALANCE("4010", "Jan 2025", "Jan 2025")` | Get GL account balance |
| `=XAVI.BUDGET("4010", "Jan 2025", "Dec 2025")` | Get budget amount |
| `=XAVI.NAME("4010")` | Get account name |
| `=XAVI.TYPE("4010")` | Get account type |
| `=XAVI.RETAINEDEARNINGS("Dec 2024")` | Calculate Retained Earnings |
| `=XAVI.NETINCOME("Mar 2025")` | Calculate Net Income YTD |
| `=XAVI.CTA("Dec 2024")` | Calculate CTA (multi-currency) |

## Documentation

| Document | Audience | Content |
|----------|----------|---------|
| [DOCUMENTATION.md](DOCUMENTATION.md) | All | Complete guide (CPA + Engineering) |
| [SPECIAL_FORMULAS_REFERENCE.md](SPECIAL_FORMULAS_REFERENCE.md) | Engineers | RE, NI, CTA calculation details |
| [SPEED-REFERENCE.md](SPEED-REFERENCE.md) | Engineers | Performance optimization techniques |
| [FUTURE_DATA_SOURCES.md](FUTURE_DATA_SOURCES.md) | Product | Multi-ERP expansion roadmap |
| [PROJECT_SUMMARY.md](PROJECT_SUMMARY.md) | Engineers | Technical architecture reference |

## Project Structure

```
├── backend/           # Python Flask server
│   ├── server.py      # API endpoints + SuiteQL queries
│   └── constants.py   # Account type constants
├── docs/              # Excel Add-in files (GitHub Pages)
│   ├── functions.js   # Custom functions
│   └── taskpane.html  # Taskpane UI
├── excel-addin/       # Manifest file
└── DOCUMENTATION.md   # Main documentation
```

## Support

For issues with:
- **Formulas showing #N/A:** Check the connection status in the taskpane
- **Values not matching NetSuite:** Verify subsidiary, period, and accounting book
- **Performance:** Use "Refresh Accounts" instead of individual cell refreshes

See [DOCUMENTATION.md](DOCUMENTATION.md) for detailed troubleshooting.

---

*Current Version: 1.5.36.0*
