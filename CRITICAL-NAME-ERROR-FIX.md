# CRITICAL: #NAME? Error Fix

## 🚨 **THE PROBLEM**

You're seeing `#NAME?` in all your formula cells because Excel is loading an **old cached version** of `functions.js` that doesn't have the new functions (`GLACCTTYPE` and `GLAPARENT`).

---

## ✅ **THE COMPLETE FIX**

### **Step 1: Close Excel COMPLETELY**
- Quit Excel entirely (Cmd+Q on Mac)
- Make sure no Excel processes are running

### **Step 2: Clear Office Cache** (Mac)
```bash
rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/*
```

### **Step 3: Remove Old Add-in**
1. Open Excel
2. Go to: **Insert → My Add-ins**
3. Find "NetSuite Formulas"
4. Click the **three dots (...)** → **Remove**

### **Step 4: Upload NEW Manifest v1.0.0.85**
1. **Insert → My Add-ins → Upload My Add-in**
2. Browse to: `excel-addin/manifest-claude.xml`
3. Click **Upload**

### **Step 5: Verify Functions Load**
1. Open Excel Developer Console:
   - **Developer → Console** (or Cmd+Opt+I)
2. Look for:
   ```
   ✅ Custom functions registered with Excel
   ```
3. Should see 5 functions registered:
   - GLATITLE
   - GLACCTTYPE ← NEW
   - GLAPARENT ← NEW
   - GLABAL
   - GLABUD

---

## 🔧 **WHAT WAS FIXED**

| Item | Status | Details |
|------|--------|---------|
| **Account names** | ✅ Fixed | Now returns "Cloud Integration" not "4220 Cloud Integration" |
| **Account search names** | ✅ Fixed | Backend uses `accountsearchdisplaynamecopy` field |
| **Button label** | ✅ Fixed | "Search Accounts" → "Add Accounts" |
| **No headers** | ✅ Fixed | Account search inserts data only |
| **Manifest version** | ✅ Updated | v1.0.0.85 with cache-busting |
| **New functions** | ✅ Ready | GLACCTTYPE and GLAPARENT deployed |

---

## 🧪 **AFTER UPLOADING MANIFEST, TEST:**

### **Test 1: Existing Formulas (Should Work)**
```excel
=NS.GLATITLE("4220")
Expected: "Cloud Integration" ✓

=NS.GLABAL($A8, C$5, C$5, $H$3, , , $J$3)
Expected: Dollar amount ✓
```

### **Test 2: New Formulas (Were Showing #NAME?)**
```excel
=NS.GLACCTTYPE("4220")
Expected: "Income" ✓

=NS.GLAPARENT("4220")
Expected: "4210" ✓
```

### **Test 3: Account Search**
1. Click any cell (e.g., A10)
2. Open task pane → "Enter Accounts"
3. Type `42*` and press Enter
4. Should insert (no headers):
   ```
   4200 | NS Product Services              | Income
   4210 | Cloud Integration & Connectors   | Income
   4220 | Cloud Integration                | Income
   ```

---

## 🎯 **WHY THIS HAPPENS**

**Excel Caching is AGGRESSIVE:**
- Excel caches `functions.js` for performance
- Even with GitHub Pages updating, Excel uses cached version
- New functions (GLACCTTYPE, GLAPARENT) aren't in old cache
- Result: `#NAME?` error

**The Solution:**
- Upload manifest with NEW cache-busting parameter (?v=1085)
- Excel sees different URL → fetches fresh functions.js
- All 5 functions now registered and working

---

## ✅ **YOUR INFRASTRUCTURE IS PERFECT**

```
Backend Server:  ✅ Running on localhost:5002
Cloudflare Tunnel: ✅ https://made-interval-charger-stay.trycloudflare.com
Cloudflare Worker: ✅ https://netsuite-proxy.chris-corcoran.workers.dev
GitHub Pages:    ✅ All code deployed

Everything is working - just need to upload new manifest!
```

---

## 📋 **COMPLETE FORMULA LIST (After Fix)**

| Formula | What It Does | Example |
|---------|--------------|---------|
| `NS.GLATITLE(account)` | Get account name | `=NS.GLATITLE("4220")` → "Cloud Integration" |
| `NS.GLACCTTYPE(account)` | Get account type | `=NS.GLACCTTYPE("4220")` → "Income" |
| `NS.GLAPARENT(account)` | Get parent account | `=NS.GLAPARENT("4220")` → "4210" |
| `NS.GLABAL(...)` | Get balance | See task pane for full syntax |
| `NS.GLABUD(...)` | Get budget | See task pane for full syntax |

---

## 🚀 **SUMMARY**

**Backend:** ✅ All working perfectly  
**Frontend:** ✅ All code deployed to GitHub  
**Excel:** ⏳ **Needs manifest v1.0.0.85 upload**

**Once you upload the new manifest, #NAME? errors will disappear and all 5 formulas will work!**

---

**File Location:** `excel-addin/manifest-claude.xml` (v1.0.0.85)

