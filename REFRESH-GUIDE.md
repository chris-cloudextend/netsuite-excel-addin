# How to Refresh NetSuite Data in Excel

Since the add-in does not use caching, every recalculation fetches fresh data from NetSuite.

## Quick Reference

### Mac
**Cmd + Option + F9** - Force refresh all formulas

### Windows
**Ctrl + Alt + F9** - Force refresh all formulas

---

## All Methods

### Keyboard Shortcuts

#### Mac
- `Cmd + =` - Recalculate current worksheet
- `F9` (or `Fn+F9`) - Recalculate entire workbook
- `Cmd + Option + F9` - **Force full recalculation (recommended)**

#### Windows
- `F9` - Recalculate entire workbook
- `Ctrl + Alt + F9` - **Force full recalculation (recommended)**
- `Ctrl + Shift + Alt + F9` - Recalculate all open workbooks

### Menu Method

1. Go to **Formulas** tab
2. Click **Calculate Now** (for entire workbook)
   OR
   Click **Calculate Sheet** (for current sheet only)

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

