# Platform Compatibility

## Overview

This add-in is built with Office.js, which is designed to be cross-platform. However, there are important compatibility considerations for each platform.

## ✅ Excel for Mac (Desktop) - **FULLY SUPPORTED**

**Status:** ✅ Tested and Working

### Installation:
1. Copy `manifest-claude.xml` to: `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
2. In Excel: Insert → My Add-ins → Shared Folder
3. Select the manifest

### Known Issues:
- Excel on Mac sometimes passes "preview invocations" (without `.close()` method)
- **Already handled** in our code via `safeFinishInvocation()`

---

## ✅ Excel for Windows (Desktop) - **SHOULD WORK**

**Status:** ⚠️ Not Tested, But Should Work

### Installation:
1. Copy `manifest-claude.xml` to: `%USERPROFILE%\AppData\Local\Microsoft\Office\16.0\Wef\`
2. In Excel: Insert → My Add-ins → Shared Folder
3. Select the manifest

### Differences from Mac:
- Windows Excel typically passes full streaming invocations (with `.close()` method)
- Our code handles both preview and full invocations, so it should work fine
- File paths are different (already documented above)
- Manifest format is identical

### Testing Needed:
- Verify formulas work correctly
- Check if invocation handling differs from Mac
- Test cache clearing on Windows

---

## ⚠️ Excel Online (Office on the web) - **PARTIALLY SUPPORTED**

**Status:** ⚠️ Works with Limitations

### What Works:
- ✅ The add-in CAN be loaded in Excel Online
- ✅ Office.js custom functions are supported
- ✅ HTTPS endpoints work (GitHub Pages, Cloudflare Worker)

### What Doesn't Work:
- ❌ **Cannot access localhost directly** from browser (CORS + security)
- ❌ **Must use Cloudflare Tunnel** (which we already do)
- ⚠️ **Performance may be slower** (extra network hops)

### Installation in Excel Online:

#### Option 1: Upload Manifest to SharePoint/OneDrive
1. Upload `manifest-claude.xml` to SharePoint or OneDrive
2. In Excel Online: Insert → Office Add-ins → MY ADD-INS → Manage My Add-ins
3. Click "Upload My Add-in"
4. Select the manifest file from SharePoint/OneDrive

#### Option 2: Use App Catalog (Enterprise)
1. Admin uploads manifest to Office 365 App Catalog
2. Users install from "MY ORGANIZATION" tab
3. Best for enterprise deployments

### Important Notes for Excel Online:

1. **Cloudflare Tunnel is REQUIRED**
   - Excel Online runs in a browser
   - Browsers cannot access `localhost`
   - The tunnel exposes your local backend via HTTPS
   - Our Cloudflare Worker already provides a stable URL

2. **HTTPS is REQUIRED**
   - Cloudflare tunnels provide HTTPS by default ✅
   - Our GitHub Pages uses HTTPS ✅
   - Mixed content (HTTP + HTTPS) is blocked by browsers ❌

3. **Performance Considerations**
   - Excel Online → GitHub Pages → Cloudflare Worker → Tunnel → localhost:5002
   - More network hops = slightly slower
   - Still acceptable for most use cases

4. **Testing Limitations**
   - Sideloading (shared folder) is NOT available in Excel Online
   - Must use SharePoint/OneDrive upload or App Catalog
   - Changes require re-uploading manifest (version bumps still help with caching)

---

## 🔧 Making It Work on All Platforms

### Current Setup (Already Done):
1. ✅ **Frontend on GitHub Pages** (HTTPS, accessible from all platforms)
2. ✅ **Cloudflare Worker proxy** (stable URL, HTTPS)
3. ✅ **Cloudflare Tunnel** (exposes localhost via HTTPS)
4. ✅ **Cross-platform code** (handles both preview and full invocations)

### What You Need to Do for Excel Online:

1. **Keep the tunnel running**:
   ```bash
   cloudflared tunnel --url http://localhost:5002
   ```

2. **Ensure Worker is updated** with current tunnel URL

3. **Upload manifest** to SharePoint/OneDrive or use App Catalog

### Architecture Diagram:

```
Excel Desktop (Mac/Windows):
  Excel → GitHub Pages (functions.js)
       → Cloudflare Worker
       → Cloudflare Tunnel
       → localhost:5002 (Flask)
       → NetSuite API

Excel Online:
  Browser → GitHub Pages (functions.js)
         → Cloudflare Worker
         → Cloudflare Tunnel
         → localhost:5002 (Flask)
         → NetSuite API
```

---

## 📊 Compatibility Matrix

| Feature | Mac Desktop | Windows Desktop | Excel Online |
|---------|-------------|-----------------|--------------|
| Custom Functions | ✅ Tested | ✅ Should Work | ✅ Should Work |
| Task Pane | ✅ Tested | ✅ Should Work | ✅ Should Work |
| Ribbon Button | ✅ Tested | ✅ Should Work | ⚠️ Limited* |
| Sideloading | ✅ Yes | ✅ Yes | ❌ No |
| Performance | ✅ Fast | ✅ Fast | ⚠️ Slower |
| Backend Access | ✅ Via Tunnel | ✅ Via Tunnel | ✅ Via Tunnel |
| Caching | ✅ Yes | ✅ Yes | ✅ Yes |

\* Some ribbon features are limited in Excel Online

---

## 🚨 Known Limitations

### All Platforms:
- Backend must be running on your machine
- Cloudflare tunnel must be active
- Internet connection required (for backend API calls)

### Excel Online Specific:
- Cannot use sideloading (Shared Folder)
- Must upload manifest to SharePoint/OneDrive
- Additional network latency
- Some advanced Excel features may not be available

### Network/Firewall:
- Corporate firewalls may block Cloudflare tunnel
- Some networks block WebSocket connections
- VPN may interfere with localhost access

---

## 🔮 Future Improvements for Excel Online

### Option 1: Deploy Backend to Cloud
- Host Flask server on Heroku, AWS, Azure, etc.
- Eliminates need for local backend + tunnel
- Better performance for Excel Online
- Requires cloud hosting costs

### Option 2: Use Office.js Server-Side
- Use Azure Functions or AWS Lambda
- Serverless backend
- Only pay for what you use
- More complex setup

### Option 3: NetSuite Direct (No Backend)
- Use NetSuite RESTlet with CORS headers
- Direct browser → NetSuite communication
- Eliminates Flask backend entirely
- Requires NetSuite configuration changes

---

## 📋 Testing Checklist

### Before Deploying to Windows:
- [ ] Test on Windows Excel (2016, 2019, or Microsoft 365)
- [ ] Verify sideloading works
- [ ] Check if invocation handling differs
- [ ] Test all three formulas (GLATITLE, GLABAL, GLABUD)
- [ ] Verify period ranges work (Jan-Mar, Jan-Dec)
- [ ] Test optional filters (subsidiary, department, etc.)

### Before Deploying to Excel Online:
- [ ] Upload manifest to SharePoint/OneDrive
- [ ] Verify Cloudflare tunnel is running
- [ ] Check HTTPS for all endpoints
- [ ] Test formulas in Excel Online
- [ ] Verify performance is acceptable
- [ ] Test with different browsers (Chrome, Edge, Safari)

---

## 💡 Recommendations

1. **Primary Platform: Excel Desktop (Mac/Windows)**
   - Best performance
   - Easiest to deploy (sideloading)
   - Full feature support

2. **Secondary Platform: Excel Online**
   - Works, but requires more setup
   - Good for remote access
   - Consider cloud backend for better performance

3. **For Enterprise Deployment:**
   - Use Office 365 App Catalog
   - Deploy backend to cloud (Azure/AWS)
   - Eliminate Cloudflare tunnel dependency
   - Centralized management

