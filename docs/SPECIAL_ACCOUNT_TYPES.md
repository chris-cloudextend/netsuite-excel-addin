# NetSuite Special Account Types (sspecacct) Reference

This document provides a comprehensive reference for the `sspecacct` field in NetSuite, which is used by `XAVI.TYPEBALANCE` when the `useSpecialAccount` parameter is set to 1.

## Purpose

The Special Account Type field (`sspecacct`) is an internal classification used on account records to flag certain system-defined or feature-specific accounts. These are accounts that NetSuite automatically creates or expects for various sub-ledgers and processes.

- Regular accounts have this field **blank**
- Any non-blank value indicates the type of system account

## Special Account Types by Category

### Accounts Receivable & Customer-Related
| Code | Description | Account Type |
|------|-------------|--------------|
| `AcctRec` | Accounts Receivable (main A/R control) | Balance Sheet |
| `UnbilledRec` | Unbilled Receivable (earned but not billed) | Balance Sheet |
| `CustDep` | Customer Deposit (prepayments) | Balance Sheet (Liability) |
| `CustAuth` | Customer Payment Authorizations | Balance Sheet (Asset) |
| `RefundPay` | Refunds Payable | Balance Sheet (Liability) |
| `UnappvPymt` | Unapproved Customer Payments | Non-Posting |

### Accounts Payable & Vendor-Related
| Code | Description | Account Type |
|------|-------------|--------------|
| `AcctPay` | Accounts Payable (main A/P control) | Balance Sheet |
| `AdvPaid` | Advances Paid to vendors | Balance Sheet (Asset) |
| `RecvNotBill` | Inventory Received Not Billed | Balance Sheet (Liability) |

### Inventory and COGS
| Code | Description | Account Type |
|------|-------------|--------------|
| `InvtAsset` | Inventory Asset | Balance Sheet |
| `COGS` | Cost of Goods Sold | P&L |
| `InvInTransit` | Inventory In Transit | Balance Sheet |
| `InvInTransitExt` | External Inventory In Transit | Balance Sheet |
| `RtnNotCredit` | Inventory Returned Not Credited | Balance Sheet |

### Deferred Revenue/Expense
| Code | Description | Account Type |
|------|-------------|--------------|
| `DeferRevenue` | Deferred Revenue | Balance Sheet (Liability) |
| `DeferExpense` | Deferred Expense/Prepaid | Balance Sheet (Asset) |
| `DeferRevClearing` | Deferred Revenue Clearing | Balance Sheet |

### Equity and Retained Earnings
| Code | Description | Account Type |
|------|-------------|--------------|
| `OpeningBalEquity` | Opening Balance Equity | Balance Sheet |
| `RetEarnings` | Retained Earnings | Balance Sheet |
| `CumulTransAdj` | Cumulative Translation Adjustment (CTA) | Balance Sheet |
| `CTA-E` | CTA - Elimination | Balance Sheet |

### Multi-Currency Gain/Loss
| Code | Description | Account Type |
|------|-------------|--------------|
| `FxRateVariance` | Foreign Currency Rate Variance | P&L |
| `RealizedERV` | Realized Gain/Loss (Exchange Rate) | P&L |
| `UnrERV` | Unrealized Gain/Loss (Exchange Rate) | P&L |
| `MatchingUnrERV` | Unrealized Matching Gain/Loss | P&L |
| `RndERV` | Rounding Gain/Loss | P&L |

### Tax Accounts
| Code | Description | Account Type |
|------|-------------|--------------|
| `SalesTaxPay` | Sales Taxes Payable | Balance Sheet (Liability) |
| `Tax` | Tax (various) | Varies |
| `TaxLiability` | Tax Liability | Balance Sheet (Liability) |
| `PSTExp` | PST Expense (Canada) | P&L |
| `PSTPay` | PST Payable (Canada) | Balance Sheet |

### Payroll and Compensation
| Code | Description | Account Type |
|------|-------------|--------------|
| `CommPay` | Commissions Payable | Balance Sheet (Liability) |
| `PayrollExp` | Payroll Expense | P&L |
| `PayrollLiab` | Payroll Liability | Balance Sheet (Liability) |
| `PayrollFloat` | Payroll Float | Balance Sheet (Asset) |
| `PayWage` | Payroll Wage | P&L |
| `PayAdjst` | Payroll Adjustment | Balance Sheet (Liability) |
| `UnappvExpRept` | Unapproved Expense Reports | Non-Posting |

### Cash and Banking
| Code | Description | Account Type |
|------|-------------|--------------|
| `UndepFunds` | Undeposited Funds | Balance Sheet (Asset) |

### Non-Posting Transaction Accounts
| Code | Description |
|------|-------------|
| `Opprtnty` | Opportunity |
| `Estimate` | Estimate/Quote |
| `SalesOrd` | Sales Order |
| `PurchOrd` | Purchase Order |
| `PurchReq` | Requisitions |
| `WorkOrd` | Work Orders |
| `RtnAuth` | Return Authorization |
| `VendAuth` | Vendor Return Authorizations |
| `RevArrng` | Revenue Arrangement |
| `TrnfrOrd` | Transfer Order |

## Usage in XAVI.TYPEBALANCE

```excel
=XAVI.TYPEBALANCE("AcctRec", , "Dec 2025", , , , , , 1)
```

The last parameter (`useSpecialAccount`) when set to `1` tells the function to filter by `sspecacct` instead of `accttype`.

## Cash Flow Statement Mapping

For building cash flow statements, these special accounts map to:

### Operating Activities (Changes in Working Capital)
- `AcctRec` - Changes in Accounts Receivable
- `AcctPay` - Changes in Accounts Payable  
- `InvtAsset` - Changes in Inventory
- `DeferRevenue` - Changes in Deferred Revenue
- `DeferExpense` - Changes in Prepaid Expenses
- `UnbilledRec` - Changes in Unbilled Receivables

### Investing Activities
- `InvInTransit` - Inventory transfers
- `FixedAsset` (use accttype) - Capital expenditures

### Financing Activities
- `RetEarnings` - Retained earnings changes
- `CumulTransAdj` - Currency translation adjustments

## Sources
- NetSuite Schema Browser
- Marty Zigman, NetSuite Account Special Types Reference (Prolecto)
- NetSuite Help Documentation

