# Cache Bug Fix & Default Subsidiary

## Issues Resolved

### Issue 1: Values Turn to $0 on Second Recalc ❌ → ✅

**Symptom:**
- User changes subsidiary filter
- First recalc shows correct values
- Second recalc shows $0 for all cells

**Root Cause:**
In `docs/functions.js` line 560-562:
```javascript
} catch (error) {
    console.error('Error distributing result:', error, key);
    cache.balance.set(key, 0);  // ❌ BUG: Caching 0 on error!
    safeFinishInvocation(req.invocation, 0);
}
```

When an error occurred (e.g., changing filters mid-batch), we cached `0`. The second recalc then returned `0` from cache instead of retrying.

**Fix:**
```javascript
} catch (error) {
    console.error('Error distributing result:', error, key);
    // ✅ Don't cache errors - just return 0 without polluting cache
    safeFinishInvocation(req.invocation, 0);
}
```

**Result:** Cache only stores successful results. Errors return `0` but don't prevent retry on next recalc.

---

### Issue 2: No Subsidiary Should Default to Consolidated Parent ❌ → ✅

**Requirement:**
When user doesn't select a subsidiary, formulas should default to **Celigo Inc. (Consolidated)** - the top-level parent with all children included.

**Old Behavior:**
```python
if subsidiary and subsidiary != '':
    target_sub = subsidiary
else:
    target_sub = None  # ❌ No consolidation, returns 0
```

**New Behavior:**
```python
if subsidiary and subsidiary != '':
    target_sub = subsidiary
else:
    target_sub = '1'  # ✅ Default to Celigo Inc. (Consolidated)
```

**Result:** 
- No subsidiary selected = Celigo Inc. (Consolidated)
- Subsidiary selected = That specific subsidiary (consolidated if parent)

---

## Test Results

```
Test: Account 59999, Jan 2024

NO subsidiary selected:
  Result: $1,317,187.91 ✅
  Expected: $1,317,188 (NetSuite Consolidated)
  Match: YES!

WITH subsidiary=1:
  Result: $1,317,187.91 ✅
  Expected: $1,317,188 (NetSuite Consolidated)
  Match: YES!

Both scenarios return the same consolidated value!
```

---

## Files Modified

1. **`docs/functions.js`**
   - Removed caching of errors (line 561)
   - Deployed to GitHub Pages ✅

2. **`backend/server.py`**
   - Changed default subsidiary from `None` to `'1'` (Celigo Inc.)
   - Running on localhost:5002 ✅
   - Proxied via Cloudflare Tunnel ✅

3. **`excel-addin/manifest-claude.xml`**
   - Bumped version to `1.0.0.72`
   - Updated cache-busting params `?v=1072`
   - Deployed to GitHub Pages ✅

---

## How to Update Excel

1. **Remove old add-in:**
   - Excel → Insert → My Add-ins
   - Click "..." on NetSuite Formulas
   - Select "Remove"

2. **Upload new manifest:**
   - Download latest: `excel-addin/manifest-claude.xml`
   - Excel → Insert → My Add-ins → Upload My Add-in
   - Browse to `manifest-claude.xml` v1.0.0.72
   - Click "Upload"

3. **Test:**
   - Clear subsidiary filter (leave blank)
   - Formulas should return Celigo Inc. (Consolidated) values
   - Changing filters should work without $0 bug

---

## What Changed in User Experience

### Before:
- ❌ Changing filters caused $0 on second recalc
- ❌ No subsidiary = $0 (useless)
- ⚠️ Had to manually select "Celigo Inc. (Consolidated)" every time

### After:
- ✅ Changing filters works smoothly
- ✅ No subsidiary = Celigo Inc. (Consolidated) automatically
- ✅ Cache only stores successful results

---

## Technical Details

### Why Default to Subsidiary 1?

**Celigo Inc.** (ID=1) is the top-level parent subsidiary. When we pass `target_sub=1` to `BUILTIN.CONSOLIDATE`:

1. NetSuite consolidates:
   - Celigo Inc. (parent) transactions
   - Celigo Australia Pty Ltd (child)
   - Celigo Europe B.V. (child)
   - Celigo India Pvt Ltd (child)
   - Elimination entries

2. Currency conversion is handled
3. Eliminations are applied
4. Result matches NetSuite's consolidated reports exactly

This is the most useful default for financial reporting!

---

**Date:** December 1, 2025  
**Status:** ✅ DEPLOYED  
**Version:** v1.0.0.72

