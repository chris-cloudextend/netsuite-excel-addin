# Excel Custom Functions Installation Troubleshooting Guide

## Status Check
✅ **All files are deployed and accessible on GitHub Pages:**
- ✅ `functions.js` - Custom functions code
- ✅ `functions.json` - Function metadata
- ✅ `functions.html` - Runtime page

## Common Installation Issues & Solutions

### Issue 1: Excel Cache (Most Common)

Excel aggressively caches add-in files. Even when you update files on GitHub Pages, Excel might still use old cached versions.

**Solution:**

1. **Run the cache clearing script:**
   ```bash
   ./clear-excel-cache-mac.sh
   ```

2. **Manual cache clearing (if script doesn't work):**
   - Close Excel completely
   - Delete these folders:
     ```
     ~/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/16.0/Wef/
     ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/
     ```
   - Restart Excel

3. **Reinstall the add-in:**
   - Open Excel
   - Go to `Insert` > `Add-ins` > `My Add-ins`
   - Click the three dots (...) next to "NetSuite Formulas"
   - Click "Remove"
   - Click "Upload My Add-in"
   - Select `excel-addin/manifest-claude.xml`

### Issue 2: Certificate/CORS Errors

If you see CORS or certificate errors in the console:

**Check the browser console:**
1. In Excel, open Developer Tools (`Insert` > `Add-ins` > `My Add-ins` > Right-click add-in > `Inspect`)
2. Look for errors like:
   - `CORS policy` errors
   - `net::ERR_CERT_` errors
   - `Mixed Content` errors

**Solution:**
- The manifest uses HTTPS URLs, which should work
- GitHub Pages has valid SSL certificates
- If you see CORS errors, the backend URL might be blocked

### Issue 3: Functions Not Registering

If the taskpane loads but formulas show `#NAME?` error:

**Check console for:**
```
❌ CustomFunctions not available!
```

**Solution:**
1. Verify `functions.html` loads correctly (no 404 errors)
2. Check that `functions.js` loads without errors
3. Look for JavaScript errors in the console

### Issue 4: Manifest Version Mismatch

The manifest has cache-busting query parameters:
- `functions.js?v=1094`
- `functions.json?v=1094`
- `taskpane.html?v=1094`

**If you recently updated files, bump these version numbers:**

1. Edit `excel-addin/manifest-claude.xml`
2. Change `v=1094` to `v=1095` (or higher)
3. Update the main `<Version>` tag (currently `1.0.0.94`)
4. Re-upload the manifest to Excel

### Issue 5: Office.js API Not Loading

If you see errors about `Office` or `Excel` not being defined:

**Solution:**
- Check internet connection (Office.js loads from CDN)
- Try using Excel Online instead of Desktop
- Update Excel to latest version

## Diagnostic Checklist

Run through this checklist to identify the issue:

### ✅ Files Deployed
- [ ] Visit https://chris-cloudextend.github.io/netsuite-excel-addin/functions.js - Should load
- [ ] Visit https://chris-cloudextend.github.io/netsuite-excel-addin/functions.json - Should load
- [ ] Visit https://chris-cloudextend.github.io/netsuite-excel-addin/functions.html - Should load

### ✅ Excel Setup
- [ ] Excel version is up to date
- [ ] Add-in is installed (visible in `Insert` > `Add-ins`)
- [ ] Taskpane opens without errors
- [ ] No errors in Developer Tools console

### ✅ Functions Working
- [ ] Type `=NS.GLATITLE("4000")` in a cell
- [ ] Should return account name (not `#NAME?` or `#VALUE!`)
- [ ] Check console for logs: `✅ Custom functions registered with Excel`

## Still Having Issues?

### Get Detailed Error Information

1. **Open Excel Developer Tools:**
   - Mac: `Insert` > `Add-ins` > `My Add-ins` > Right-click "NetSuite Formulas" > `Inspect`
   
2. **Check Console Tab:**
   - Look for red error messages
   - Copy the full error text
   
3. **Check Network Tab:**
   - Look for failed requests (red)
   - Check if `functions.js`, `functions.json`, `functions.html` loaded successfully
   
4. **Try a simple formula:**
   ```
   =NS.GLATITLE("4000")
   ```
   - What error do you see?
   - `#NAME?` = Function not registered
   - `#VALUE!` = Function registered but error in execution
   - `#N/A` = Function executed but returned N/A (check server)

### Common Error Messages & Fixes

| Error Message | Cause | Fix |
|--------------|-------|-----|
| `#NAME?` | Excel doesn't recognize the function | Clear cache, reinstall add-in |
| `#VALUE!` | Function error | Check console for JavaScript errors |
| `#N/A` | API returned N/A | Check backend server is running |
| `CustomFunctions not available` | Functions.js not loaded | Check Network tab, clear cache |
| `CORS policy` | Backend blocking requests | Check backend CORS settings |
| `net::ERR_CERT_` | SSL certificate issue | Use HTTPS URLs in manifest |

## Quick Fix (Nuclear Option)

If nothing else works:

1. **Completely uninstall Excel:**
   ```bash
   # Remove Excel
   rm -rf ~/Library/Containers/com.microsoft.Excel
   rm -rf ~/Library/Group\ Containers/UBF8T346G9.Office
   rm -rf ~/Library/Preferences/com.microsoft.Excel.plist
   ```

2. **Reinstall Excel from Office 365**

3. **Install add-in fresh**

---

## Need More Help?

**Provide this information:**
1. What specific error message do you see?
2. Screenshot of Developer Tools console
3. What happens when you type `=NS.GLATITLE("4000")`?
4. Excel version (Excel > About)
5. Mac OS version

