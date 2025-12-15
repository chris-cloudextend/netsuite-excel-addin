# XAVI.TYPEBALANCE User Guide

## Understanding Account Type vs. Special Account Type

A quick guide for choosing the right option in **XAVI.TYPEBALANCE**

When pulling financial data into Excel with XAVI, NetSuite provides two ways to classify accounts. Choosing the right one determines how precise or broad your results will be.

---

## 1. Account Type (Standard Financial Category)

Account Type is the traditional classification used for financial statements:

| Account Type | Description |
|--------------|-------------|
| `Bank` | Bank and cash accounts |
| `AcctRec` | Accounts Receivable |
| `OthCurrAsset` | Other Current Assets |
| `FixedAsset` | Fixed Assets |
| `OthAsset` | Other Assets |
| `AcctPay` | Accounts Payable |
| `CredCard` | Credit Card |
| `OthCurrLiab` | Other Current Liabilities |
| `LongTermLiab` | Long-Term Liabilities |
| `Equity` | Equity |
| `Income` | Income |
| `COGS` | Cost of Goods Sold |
| `Expense` | Expenses |
| `OthIncome` | Other Income |
| `OthExpense` | Other Expense |

Every account has exactly one Account Type. NetSuite uses these types for the Balance Sheet, Income Statement, and Cash Flow Statement.

### When to Use Account Type

Use this whenever you want broad groupings, such as:

* All assets of a certain type
* All liabilities
* All expenses
* All income accounts

**Example - Total of all Other Current Liabilities:**

```excel
=XAVI.TYPEBALANCE("OthCurrLiab",,"Dec 2025")
```

**Example - Total Expenses for a period:**

```excel
=XAVI.TYPEBALANCE("Expense","Jan 2025","Dec 2025")
```

---

## 2. Special Account Type (System-Defined Control Accounts)

Special Account Type identifies NetSuite's *system accounts*, created automatically for functions like AR, AP, Inventory, Taxes, and Multi-Currency.

### Common Special Account Types

| Code | Description | Category |
|------|-------------|----------|
| `AcctRec` | Accounts Receivable (control) | Balance Sheet |
| `AcctPay` | Accounts Payable (control) | Balance Sheet |
| `InvtAsset` | Inventory Asset | Balance Sheet |
| `UndepFunds` | Undeposited Funds | Balance Sheet |
| `DeferRevenue` | Deferred Revenue | Balance Sheet |
| `DeferExpense` | Deferred Expense / Prepaid | Balance Sheet |
| `RetEarnings` | Retained Earnings | Balance Sheet |
| `SalesTaxPay` | Sales Tax Payable | Balance Sheet |
| `CumulTransAdj` | Cumulative Translation Adjustment | Balance Sheet |
| `COGS` | Cost of Goods Sold (system) | P&L |
| `RealizedERV` | Realized FX Gain/Loss | P&L |
| `UnrERV` | Unrealized FX Gain/Loss | P&L |

Only true control accounts have these values. User-created accounts typically have a blank special account type.

### When to Use Special Account Type

Choose this when you need:

* The *real* AR or AP control account (not user-created AR sub-accounts)
* Consistent results across subsidiaries
* A precise cash flow model using working-capital accounts
* To avoid including custom user-created accounts

**Example - True Accounts Receivable balance:**

```excel
=XAVI.TYPEBALANCE("AcctRec",,"Dec 2025",,,,,,1)
```

**Example - Inventory Asset for cash flow:**

```excel
=XAVI.TYPEBALANCE("InvtAsset",,"Dec 2025",,,,,,1)
```

The `1` at the end tells the formula to use Special Account Type instead of regular Account Type.

---

## Formula Syntax

```
=XAVI.TYPEBALANCE(accountType, fromPeriod, toPeriod, subsidiary, department, location, classId, accountingBook, useSpecialAccount)
```

| Parameter | Position | Description |
|-----------|----------|-------------|
| `accountType` | 1 | Required. The account type or special account type code |
| `fromPeriod` | 2 | Start period (required for P&L, ignored for BS) |
| `toPeriod` | 3 | End period (required) |
| `subsidiary` | 4 | Optional. Subsidiary name or ID |
| `department` | 5 | Optional. Department filter |
| `location` | 6 | Optional. Location filter |
| `classId` | 7 | Optional. Class filter |
| `accountingBook` | 8 | Optional. Accounting Book ID |
| `useSpecialAccount` | 9 | Optional. Set to `1` to use Special Account Type |

### Balance Sheet vs P&L Behavior

- **Balance Sheet types**: `fromPeriod` is ignored; calculates cumulative from inception through `toPeriod`
- **P&L types**: Uses the range from `fromPeriod` to `toPeriod`

---

## Examples

### Using Regular Account Type (Position 9 = 0 or blank)

```excel
// All Other Current Assets as of Dec 2025
=XAVI.TYPEBALANCE("OthCurrAsset",,"Dec 2025")

// All Expenses for full year 2025
=XAVI.TYPEBALANCE("Expense","Jan 2025","Dec 2025")

// All Income for Q1 2025 for specific subsidiary
=XAVI.TYPEBALANCE("Income","Jan 2025","Mar 2025","Celigo Inc.")
```

### Using Special Account Type (Position 9 = 1)

```excel
// True A/R control account balance
=XAVI.TYPEBALANCE("AcctRec",,"Dec 2025",,,,,,1)

// True A/P control account balance
=XAVI.TYPEBALANCE("AcctPay",,"Dec 2025",,,,,,1)

// Inventory Asset for cash flow analysis
=XAVI.TYPEBALANCE("InvtAsset",,"Dec 2025",,,,,,1)

// Deferred Revenue balance
=XAVI.TYPEBALANCE("DeferRevenue",,"Dec 2025",,,,,,1)

// Retained Earnings
=XAVI.TYPEBALANCE("RetEarnings",,"Dec 2025",,,,,,1)
```

---

## Which Should I Choose?

| Goal | Use Account Type | Use Special Account Type |
|------|------------------|--------------------------|
| Build Balance Sheet | ✔ | |
| Build Income Statement | ✔ | |
| Broad financial categories | ✔ | |
| True AR/AP/Inventory accounts | | ✔ |
| Cash Flow (working-capital deltas) | | ✔ |
| Identify Undeposited Funds or Tax accounts | | ✔ |
| Avoid including user-created accounts | | ✔ |
| Multi-currency FX gain/loss accounts | | ✔ |

---

## Rule of Thumb

> **Use Account Type for broad financial reporting.**
> 
> **Use Special Account Type when you need precision control accounts.**

---

## Cash Flow Statement Example

For a proper indirect method cash flow statement, use Special Account Types to get precise working capital changes:

```excel
// Change in Accounts Receivable
=XAVI.TYPEBALANCE("AcctRec",,"Dec 2024",,,,,,1) - XAVI.TYPEBALANCE("AcctRec",,"Dec 2025",,,,,,1)

// Change in Inventory
=XAVI.TYPEBALANCE("InvtAsset",,"Dec 2024",,,,,,1) - XAVI.TYPEBALANCE("InvtAsset",,"Dec 2025",,,,,,1)

// Change in Accounts Payable
=XAVI.TYPEBALANCE("AcctPay",,"Dec 2025",,,,,,1) - XAVI.TYPEBALANCE("AcctPay",,"Dec 2024",,,,,,1)

// Change in Deferred Revenue
=XAVI.TYPEBALANCE("DeferRevenue",,"Dec 2025",,,,,,1) - XAVI.TYPEBALANCE("DeferRevenue",,"Dec 2024",,,,,,1)
```

---

## See Also

- [SPECIAL_ACCOUNT_TYPES.md](SPECIAL_ACCOUNT_TYPES.md) - Complete list of all special account type codes
- [DEVELOPER_CHECKLIST.md](../DEVELOPER_CHECKLIST.md) - Developer integration guide

