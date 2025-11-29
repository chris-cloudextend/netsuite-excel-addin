# NetSuite Excel Formulas - Project Structure

## 📁 Project Organization

```
NetSuite Formulas Revised/
├── backend/                          # Flask backend server
│   ├── server.py                    # Main backend API
│   ├── netsuite_config.json         # NetSuite credentials
│   └── requirements.txt             # Python dependencies
│
├── docs/                             # GitHub Pages files (deployed)
│   ├── taskpane.html                # Task pane UI with refresh button
│   ├── functions.js                 # Custom functions (NS.GLATITLE, NS.GLABAL, NS.GLABUD)
│   ├── functions.json               # Function metadata
│   ├── functions.html               # Function page
│   ├── icon-16.png                  # Add-in icons (4 sizes)
│   ├── icon-32.png
│   ├── icon-64.png
│   └── icon-80.png
│
├── excel-addin/                      # Manifest file
│   └── manifest-claude.xml          # Excel add-in manifest (PRODUCTION)
│
└── REFRESH-GUIDE.md                 # User documentation for refresh
```

---

## 🎯 Key Files

### **Manifest (Deployment)**
- `excel-addin/manifest-claude.xml` - **Upload this to Microsoft Admin Center**
  - Current version: 1.0.0.9
  - Cache-busting: ?v=1009
  - Data tab button configuration

### **Backend Server**
- `backend/server.py` - Flask API server
  - Endpoints: /account/<>/name, /balance, /budget, /batch/balance
  - Port: localhost:5002
  - Must be running for formulas to work

- `backend/netsuite_config.json` - NetSuite credentials
  - OAuth 1.0a (TBA) credentials
  - Account ID, tokens, secrets

### **GitHub Pages (Public)**
- `docs/functions.js` - Custom functions implementation
  - Contains: NS.GLATITLE, NS.GLABAL, NS.GLABUD
  - Intelligent batching logic
  - Non-volatile (manual refresh only)

- `docs/taskpane.html` - Task pane UI
  - "Refresh All Data" button
  - Formula documentation
  - Help and examples

---

## 🚀 Deployment

### **1. Manifest Upload**
- Go to: Microsoft 365 Admin Center
- Upload: `excel-addin/manifest-claude.xml`
- Method: Centralized Deployment

### **2. GitHub Pages**
- Repo: chris-cloudextend/netsuite-excel-addin
- URL: https://chris-cloudextend.github.io/netsuite-excel-addin/
- Auto-deploys from `docs/` folder

### **3. Backend Server**
- Run: `cd backend && python3 server.py`
- Port: localhost:5002
- Cloudflare tunnel exposes to internet

---

## 🔧 Services Required

Three services must be running:

1. **Flask Backend**
   ```bash
   cd backend
   python3 server.py &
   ```

2. **Cloudflare Tunnel**
   ```bash
   cloudflared tunnel --url http://localhost:5002 &
   ```
   - Get tunnel URL from output
   - Update `docs/functions.js` with new URL
   - Push to GitHub Pages

3. **GitHub Pages**
   - Automatically serves `docs/` folder
   - No manual deployment needed

---

## 📝 Making Changes

### **Update Manifest:**
1. Edit `excel-addin/manifest-claude.xml`
2. Increment version (e.g., 1.0.0.9 → 1.0.0.10)
3. Update cache-busting (e.g., ?v=1009 → ?v=1010)
4. Upload to Microsoft Admin Center

### **Update Custom Functions:**
1. Edit `docs/functions.js`
2. Commit and push to GitHub
3. Wait 2-3 minutes for GitHub Pages
4. Users quit/reopen Excel (cache-busting ensures fresh load)

### **Update Task Pane:**
1. Edit `docs/taskpane.html`
2. Commit and push to GitHub
3. Wait 2-3 minutes for GitHub Pages

---

## 🧹 Cleanup Done

**Removed:**
- ❌ All old manifest files (manifest.xml, manifest-*.xml except claude)
- ❌ Duplicate files in excel-addin/ folder
- ❌ Old documentation files
- ❌ Backup files (.bak)

**Kept:**
- ✅ manifest-claude.xml (production manifest)
- ✅ backend/ (Flask server)
- ✅ docs/ (GitHub Pages files)
- ✅ REFRESH-GUIDE.md (user documentation)

---

## 📊 Current Configuration

- **Manifest Version:** 1.0.0.9
- **Cache-Busting:** ?v=1009
- **Tunnel URL:** https://load-scanner-nathan-targeted.trycloudflare.com
- **Backend Port:** localhost:5002
- **GitHub Pages:** https://chris-cloudextend.github.io/netsuite-excel-addin/

---

## ✅ Everything Ready for Production!

