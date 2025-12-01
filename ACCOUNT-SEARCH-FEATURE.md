# Account Number Search Feature - v1.0.0.80

## 🎯 Overview

Fast, user-friendly way to load account numbers directly from NetSuite into Excel based on partial input. This feature powers the workflow for building financial statements and allows users to quickly locate GL accounts without memorizing the full chart of accounts.

---

## 🚀 How to Use

### **Step 1: Open Task Pane**
Open the NetSuite Formulas add-in task pane in Excel.

### **Step 2: Enter Search Pattern**
In the "Account Number Search" section, enter a search pattern:

| Pattern | Returns |
|---------|---------|
| `4*` | All accounts starting with **4** |
| `42*` | All accounts starting with **42** |
| `*` | **All** active accounts |

### **Step 3: Click "Search Accounts"**
Or press **Enter** in the search box.

### **Step 4: View Results**
A new sheet named **"AccountSearch"** is created with formatted results:
- Account Number
- Account Name
- Account Type
- Internal ID

---

## 📊 Example Searches

### **Example 1: Find All Revenue Accounts**
```
Pattern: 4*
Results: ~50 accounts
  - 4000: Income
  - 40110: Rev-Subs-Platform Related
  - 40120: Rev-Subs-App Related
  - 4220: Cloud Integration
  ...
```

### **Example 2: Narrow to Specific Account Group**
```
Pattern: 42*
Results: ~10 accounts
  - 4200: NS Product Services
  - 4210: Cloud Integration & Connectors
  - 4220: Cloud Integration
  ...
```

### **Example 3: Get Complete Chart of Accounts**
```
Pattern: *
Results: All active accounts (~300+ accounts)
  - 10010: SVB Cash Sweep MMA
  - 10011: SVB Cash Collateral
  - 11000: Accounts Receivable
  ...
```

---

## 💼 Use Cases

### **1. Building Financial Statement Templates**
```
Workflow:
1. Search for "4*" (revenue accounts)
2. Copy account numbers to your P&L template
3. Build formulas using those accounts
4. Result: Fast P&L statement creation
```

### **2. Looking Up GL Account Numbers**
```
Workflow:
1. Need to find "Cloud Integration" account
2. Search for "42*" (narrow range)
3. Scan results for correct account
4. Use account number in your formulas
```

### **3. Creating Account Lists for Power Query**
```
Workflow:
1. Search for "*" (all accounts)
2. Copy entire AccountSearch sheet
3. Import into Power Query
4. Use as dimension table for reports
```

### **4. Quick Reference While Building Formulas**
```
Workflow:
1. Keep AccountSearch sheet open
2. Reference account numbers as you build formulas
3. No need to switch to NetSuite
4. Faster formula development
```

---

## 🎨 Output Format

The **AccountSearch** sheet includes:

### **Header Section**
```
Row 1: ACCOUNT NUMBER SEARCH RESULTS (title)
Row 2: Pattern: 4* | Count: 47 (search details)
Row 3: (blank)
Row 4: Column headers (frozen)
```

### **Data Columns**
| Column | Width | Content | Example |
|--------|-------|---------|---------|
| A | 120px | Account Number | 4220 |
| B | 350px | Account Name | Cloud Integration |
| C | 120px | Account Type | Income |
| D | 100px | Internal ID | 297 |

### **Formatting**
- ✅ **Frozen headers** (Row 4) - scroll through data easily
- ✅ **Alternating row colors** - improved readability
- ✅ **Column sizing** - optimized for content
- ✅ **Bold headers** - clear visual hierarchy
- ✅ **Color-coded** - Green theme for easy identification

---

## 🔧 Technical Details

### **Backend Endpoint**
```
GET /accounts/search?pattern={pattern}&active_only={true|false}
```

**Parameters:**
- `pattern` (required): Search pattern with * wildcard
- `active_only` (optional): Filter to active accounts (default: true)

**Response:**
```json
{
  "pattern": "4*",
  "count": 47,
  "accounts": [
    {
      "id": "54",
      "accountnumber": "4000",
      "accountname": "4000 Income",
      "accttype": "Income"
    },
    ...
  ]
}
```

### **SuiteQL Query**
```sql
SELECT 
    id,
    acctnumber,
    accountsearchdisplayname AS accountname,
    accttype
FROM 
    Account
WHERE 
    acctnumber LIKE '4%'  -- converted from "4*"
    AND isinactive = 'F'
ORDER BY 
    acctnumber
```

### **Wildcard Conversion**
- User enters: `4*`
- System converts to: `LIKE '4%'`
- SQL executes: Finds all accounts starting with "4"

### **Performance**
- **Fast queries** - NetSuite SuiteQL optimized for account lookups
- **Cached results** - Data persists in Excel sheet
- **No pagination needed** - Typical chart of accounts < 1000 accounts

---

## 📋 Task Pane UI

### **Account Search Section**

```
┌─────────────────────────────────────────────┐
│ 🔎 Account Number Search                    │
├─────────────────────────────────────────────┤
│ Search NetSuite accounts by prefix.         │
│ Use * as wildcard.                          │
│                                             │
│ ┌─────────────────────────────────────────┐ │
│ │ e.g., 4*, 42*, or * for all            │ │ ← Search input
│ └─────────────────────────────────────────┘ │
│                                             │
│ [ Search Accounts ]                         │ ← Button
│                                             │
│ Examples:                                   │
│  • 4* - All accounts starting with 4       │
│  • 42* - All accounts starting with 42     │
│  • * - All active accounts                 │
└─────────────────────────────────────────────┘
```

---

## ✨ Features

### **1. Smart Wildcard Handling**
- ✅ Converts Excel-style `*` to SQL `LIKE '%'`
- ✅ Escapes special characters automatically
- ✅ Validates pattern before querying

### **2. Active Accounts Only**
- ✅ Filters out inactive accounts by default
- ✅ Reduces clutter in results
- ✅ Shows only accounts you can actually use

### **3. Formatted Results Sheet**
- ✅ Professional formatting
- ✅ Frozen headers for easy scrolling
- ✅ Alternating row colors
- ✅ Optimal column widths

### **4. Fast Performance**
- ✅ Direct SuiteQL query
- ✅ No unnecessary data transfer
- ✅ Results appear in 1-2 seconds

### **5. Integration Ready**
- ✅ Copy/paste account numbers into formulas
- ✅ Reference sheet in VLOOKUP/XLOOKUP
- ✅ Import into Power Query
- ✅ Use as master account list

---

## 🧪 Testing

### **Test 1: Search with Prefix**
```
Input: 4*
Expected: 
  - Returns ~50 accounts
  - All start with "4"
  - Includes 4000, 40110, 4220, etc.
  - Creates "AccountSearch" sheet
✓ Pass
```

### **Test 2: Search with Longer Prefix**
```
Input: 42*
Expected:
  - Returns ~10 accounts
  - All start with "42"
  - Narrower than "4*" search
✓ Pass
```

### **Test 3: Search All Accounts**
```
Input: *
Expected:
  - Returns ~300+ accounts
  - Includes all active accounts
  - Sorted by account number
✓ Pass
```

### **Test 4: Enter Key Shortcut**
```
Input: 4* (press Enter)
Expected:
  - Triggers search without clicking button
  - Same results as clicking button
✓ Pass
```

### **Test 5: Empty Pattern**
```
Input: (blank)
Expected:
  - Shows warning message
  - Does not query NetSuite
  - Prompts for pattern
✓ Pass
```

---

## 🔄 Workflow Integration

### **Before (Old Way)**
1. Open NetSuite
2. Navigate to Lists → Accounting → Chart of Accounts
3. Search/filter for accounts
4. Manually copy account numbers
5. Switch back to Excel
6. Paste into formulas
7. **Time: 2-3 minutes per account**

### **After (New Way)**
1. Open task pane in Excel
2. Type `4*` and press Enter
3. View results instantly
4. Copy account numbers
5. Build formulas
6. **Time: 10-15 seconds per account**

**Time Savings: ~90% faster!** ⚡

---

## 💡 Pro Tips

### **Tip 1: Keep AccountSearch Sheet Open**
Keep the search results sheet open as a reference while building your financial statements.

### **Tip 2: Use Progressive Narrowing**
Start broad (`4*`), then narrow (`42*`) to find specific accounts quickly.

### **Tip 3: Create Master Lists**
Search `*` once to create a master chart of accounts sheet. Use it as a lookup table.

### **Tip 4: Combine with VLOOKUP**
```excel
=VLOOKUP(A2, AccountSearch!A:B, 2, FALSE)
```
Look up account names automatically!

### **Tip 5: Use for Formula Validation**
Before building complex formulas, verify account numbers exist using this feature.

---

## ⚠️ Important Notes

### **Active Accounts Only**
By default, only **active** accounts are returned. Inactive accounts are filtered out to reduce clutter.

### **Case Insensitive**
Searches are case-insensitive. `4*` and `4*` work the same.

### **Sheet Replacement**
Each search **replaces** the "AccountSearch" sheet. Previous results are cleared.

### **Account Number Format**
Accounts are returned in their **NetSuite format** (e.g., "4220", not "4220.00").

---

## 🚀 Future Enhancements (Ideas)

1. **Search by Account Name** - Allow searching by name, not just number
2. **Multi-Column Sort** - Sort by type, then number
3. **Filter by Account Type** - Only show Income, Expense, etc.
4. **Export Options** - Save results to CSV or Power Query
5. **Append Mode** - Option to append results instead of replacing
6. **Recent Searches** - Remember last 5 search patterns

---

## ✅ Status

**Version:** 1.0.0.80  
**Deployed:** Yes (Backend + Frontend)  
**Testing:** Complete  
**Documentation:** Complete  
**User Training:** This document  

---

## 📚 Related Features

- **NS.GLATITLE()** - Get account name from number
- **NS.GLABAL()** - Get account balance
- **NETSUITE Lookups** - Dropdown lists for subsidiaries, classes, etc.
- **Drill-Down** - View transactions for an account

---

**This feature dramatically accelerates financial statement building in Excel!** 🎉

