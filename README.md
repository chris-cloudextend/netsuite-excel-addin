# NetSuite Excel Formulas Add-in

Real-time NetSuite SuiteQL data directly in Excel via custom formulas.

## 🎯 Features

- **NS.GLATITLE** - Get account names
- **NS.GLABAL** - Get GL balances with optional filters (subsidiary, department, location, class)
- **NS.GLABUD** - Get budget amounts with optional filters
- **Intelligent batching** - Multiple formulas consolidated into single efficient queries
- **Manual refresh** - Task pane button to update all formulas with latest NetSuite data
- **Transaction drill-down** - Click any balance cell to see underlying transactions with NetSuite hyperlinks

---

## 📁 Project Structure

```
├── backend/                    # Flask backend server
│   ├── server.py              # Main API (OAuth + SuiteQL)
│   ├── netsuite_config.json   # NetSuite credentials (not in git)
│   └── requirements.txt       # Python dependencies
│
├── docs/                       # GitHub Pages (public hosting)
│   ├── taskpane.html          # Task pane UI with refresh button
│   ├── functions.js           # Custom functions implementation
│   ├── functions.json         # Function metadata
│   └── icon-*.png             # Add-in icons
│
└── excel-addin/
    └── manifest-claude.xml    # Excel add-in manifest (PRODUCTION)
```

---

## 🚀 Quick Start

### 1. Start Backend Server
```bash
cd backend
python3 server.py
```

### 2. Start Cloudflare Tunnel
```bash
cloudflared tunnel --url http://localhost:5002
```
Copy the tunnel URL and update `docs/functions.js`

### 3. Deploy to Excel
- Upload `excel-addin/manifest-claude.xml` to Microsoft 365 Admin Center
- Use Centralized Deployment

---

## 💼 Usage in Excel

### Insert Formulas:
```excel
=NS.GLATITLE(4010)
=NS.GLABAL("4010", "Jan 2025", "Dec 2025")
=NS.GLABAL("4010", "Jan 2025", "Dec 2025", "", "13", "", "")
=NS.GLABUD("5000", "Jan 2025", "Dec 2025")
```

### Refresh Data:
1. Data tab → CloudExtend → "NetSuite Formulas"
2. Click "Refresh All Data" button

### Drill Down to Transactions:
1. Select any cell with an **NS.GLABAL** formula
2. Data tab → CloudExtend → "NetSuite Formulas"
3. Click "View Transactions" button
4. New sheet created with transaction details and **clickable NetSuite links**!

---

## 🔧 Configuration

### NetSuite Credentials
Edit `backend/netsuite_config.json`:
```json
{
  "account_id": "YOUR_ACCOUNT_ID",
  "consumer_key": "YOUR_CONSUMER_KEY",
  "consumer_secret": "YOUR_CONSUMER_SECRET",
  "token_id": "YOUR_TOKEN_ID",
  "token_secret": "YOUR_TOKEN_SECRET"
}
```

### Tunnel URL
When you restart the Cloudflare tunnel, update `docs/functions.js`:
```javascript
const SERVER_URL = 'https://your-new-tunnel-url.trycloudflare.com';
```

---

## 📊 Architecture

```
Excel Cell (=NS.GLABAL(...))
    ↓
GitHub Pages (functions.js)
    ↓
Cloudflare Tunnel (HTTPS)
    ↓
Flask Backend (localhost:5002)
    ↓
NetSuite SuiteQL API (OAuth 1.0a)
```

---

## 🎨 Current Configuration

- **Manifest Version:** 1.0.0.9
- **Cache-Busting:** ?v=1009
- **Backend:** localhost:5002
- **Tunnel:** https://load-scanner-nathan-targeted.trycloudflare.com
- **GitHub Pages:** https://chris-cloudextend.github.io/netsuite-excel-addin/

---

## 📚 Documentation

See `PROJECT-STRUCTURE.md` for detailed project organization and deployment instructions.

---

## 🔒 Security Note

Never commit `backend/netsuite_config.json` to git - it contains sensitive credentials.
Use `netsuite_config.template.json` as a template.

