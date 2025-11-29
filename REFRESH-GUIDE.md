# How to Refresh NetSuite Data in Excel

## The ONLY Way to Refresh

Excel custom functions (Office.js) do not respond to F9 or standard Excel recalculation commands. This is a platform limitation.

**To refresh NetSuite data, you MUST use the manual refresh button:**

1. **Open the task pane:** Insert → My Add-ins → NetSuite Formulas
2. **Click "Refresh All Data"** button (yellow section at top)
3. All NetSuite formulas will recalculate with fresh data

---

## Why F9 Doesn't Work

Excel custom functions created with Office.js are **not the same as built-in Excel functions**. They do not respond to:
- ❌ F9 (Recalculate)
- ❌ Ctrl+Alt+F9
- ❌ Formulas → Calculate Now
- ❌ Any Excel recalculation command

This is a **platform limitation** of Office.js custom functions, not a bug.

### Two Options Were Considered:

1. **Volatile functions** (like `NOW()` or `RAND()`)
   - ❌ Would recalculate on EVERY cell edit
   - ❌ Would make hundreds of API calls during normal editing
   - ❌ Would freeze Excel for 30-60 seconds per edit
   - ❌ Terrible user experience

2. **Manual refresh button** (chosen approach)
   - ✅ User controls when to refresh
   - ✅ Excel stays fast during editing
   - ✅ No unexpected API calls
   - ✅ Professional user experience

---

## How It Works

Because there is **no caching**:
- ✅ Every recalculation = fresh API call to NetSuite
- ✅ Always get the most recent data
- ✅ No stale cached values
- ✅ No need to "clear cache"

**Note:** Intelligent batching still works! When you recalculate many formulas at once, they're bundled into efficient batch requests for better performance.

---

## Recommended Workflow

### Daily Reporting
1. Open your workbook
2. Press **Cmd+Option+F9** (Mac) or **Ctrl+Alt+F9** (Windows)
3. Wait for all formulas to refresh (~5-10 seconds for typical reports)
4. Work with current NetSuite data

### Before Important Meetings
1. Press **Cmd+Option+F9** / **Ctrl+Alt+F9**
2. Ensure all numbers are current
3. Present with confidence!

---

## Performance

With intelligent batching enabled:
- 100 cells (10 accounts × 12 months): **~5-10 seconds**
- Fresh data from NetSuite every time
- No waiting for individual cells

---

## Automatic Recalculation

By default, Excel recalculates automatically when you:
- Change any cell value
- Open the workbook
- Press Enter in a cell

To verify automatic calculation is enabled:
- **Formulas** tab → **Calculation Options** → **Automatic**

---

## Troubleshooting

**Formulas not updating?**
1. Try **Cmd+Option+F9** (Mac) or **Ctrl+Alt+F9** (Windows)
2. Check that backend server is running
3. Check console (F12) for any errors

**Taking too long?**
- Normal for first calculation (fetching from NetSuite)
- Subsequent recalculations use batching for efficiency
- Large reports (100+ cells) may take 10-20 seconds

---

Remember: **No caching = Always fresh!** Just press F9 when you need the latest data. 🎉

