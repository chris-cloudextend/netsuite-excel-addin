# CTA Investigation - Complete Account-Level Analysis

## Summary

| Component | My Query | Claude's Expected | Difference |
|-----------|----------|-------------------|------------|
| **RE** | -$88,022,956.91 | -$88,022,956.91 | **✅ EXACT** |
| **NI** | -$8,781,243.65 | -$8,781,243.65 | **✅ EXACT** |
| **Total Assets** | $53,681,594.33 | $53,322,353.28 | **+$359,241** |
| **Total Liabilities** | $60,099,076.91 | $59,987,254.09 | **+$111,823** |
| **Total Equity** | -$6,417,482.58 | -$6,664,900.80 | **+$247,418** |
| **CTA (plug)** | -$29,747.85 | -$239,639.06 | **~$210K off** |

## Key Finding: Intercompany Accounts

### IC Assets (eliminate='T')
| Account | Name | Balance |
|---------|------|---------|
| 15401 | InterCompany Receivable - IntX UK | +$41,747.36 |
| 15200 | InterCompany Receivable - IntX EMEA | +$40,914.55 |
| 15400-1 | InterCompany Receivable | +$34,300.48 |
| 15200-1 | InterCompany Receivable | +$22,041.86 |
| 15600 | InterCompany Receivable - IntX AU | +$2,340.65 |
| 15400 | InterCompany Receivable - IntX IN | -$3,175.96 |
| 15000-1 | InterCompany Receivable | -$9,564.22 |
| 15210-1 | InterCompany Receivable | -$14,520.10 |
| 15401-1 | InterCompany Receivable | -$17,499.63 |
| 15900 | Due From | -$17,762.92 |
| 15500 | InterCompany Receivable - IntX | -$50,286.03 |
| **TOTAL IC ASSETS** | | **+$28,536.04** |

### Tested Approaches

| Approach | Assets Result | Diff from NetSuite |
|----------|---------------|-------------------|
| WITH IC accounts | $53,681,594 | +$359K |
| WITHOUT IC accounts (`eliminate != 'T'`) | $53,653,058 | +$331K |

**Conclusion:** Excluding IC accounts only reduces discrepancy by ~$28K. The remaining ~$331K is from non-IC accounts.

## Attempted Consolidation Parameters

### Tried 'ELIMINATE' parameter:
```sql
BUILTIN.CONSOLIDATE(tal.amount, 'LEDGER', 'DEFAULT', 'ELIMINATE', 1, t.postingperiod, 'DEFAULT')
```
**Result:** Error - `"Subsidiary rate type const doesn't exist: ELIMINATE"`

### Current (working) parameters:
```sql
BUILTIN.CONSOLIDATE(tal.amount, 'LEDGER', 'DEFAULT', 'DEFAULT', 1, t.postingperiod, 'DEFAULT')
```

## Account Type Breakdown

### Assets by Type (My Query)
| Type | Balance |
|------|---------|
| Bank | $20,850,298.77 |
| AcctRec | $10,597,247.92 |
| DeferExpense | $15,640,814.07 |
| FixedAsset | $411,084.02 |
| OthAsset | $1,889,891.51 |
| OthCurrAsset | $4,292,258.05 |
| **TOTAL** | **$53,681,594.33** |

Note: `UnbilledRec` has 12,241 transactions but consolidated balance = $0.00

### Liabilities by Type (My Query)
| Type | Balance |
|------|---------|
| AcctPay | $464,932.81 |
| CredCard | $66,251.44 |
| DeferRevenue | $46,566,219.61 |
| LongTermLiab | $1,530,179.49 |
| OthCurrLiab | $11,471,493.56 |
| **TOTAL** | **$60,099,076.91** |

## Questions for Claude

1. **How did you get your expected Total Assets = $53,322,353.28?**
   - Did you use a different date or subsidiary?
   - Did you run NetSuite's native Balance Sheet report?
   - What specific account types did you include/exclude?

2. **My +$359K excess is NOT explained by IC accounts alone**
   - IC accounts only contribute +$28K
   - Excluding them leaves +$331K unexplained
   - Which non-IC accounts might I be including that NetSuite excludes?

3. **What are the correct BUILTIN.CONSOLIDATE parameters?**
   - 'ELIMINATE' doesn't work as a parameter
   - What's the proper way to force intercompany elimination?

4. **Could this be a timing/date issue?**
   - Are you certain the expected values are for Dec 31, 2024?
   - Were they from a specific saved report configuration?

## My CTA Plug Query (Complete Code)

```python
# Asset types
asset_types = "'Bank', 'AcctRec', 'OthCurrAsset', 'FixedAsset', 'OthAsset', 'DeferExpense'"

# Liability types
liability_types = "'AcctPay', 'CredCard', 'OthCurrLiab', 'LongTermLiab', 'DeferRevenue'"

# Consolidation SQL
cons_amount = f"""TO_NUMBER(BUILTIN.CONSOLIDATE(tal.amount, 'LEDGER', 'DEFAULT', 'DEFAULT', {target_sub}, t.postingperiod, 'DEFAULT'))"""

# Asset query
SELECT SUM({cons_amount}) AS total_assets
FROM transactionaccountingline tal
JOIN transaction t ON t.id = tal.transaction
JOIN account a ON a.id = tal.account
JOIN accountingperiod ap ON ap.id = t.postingperiod
WHERE t.posting = 'T'
  AND tal.posting = 'T'
  AND a.accttype IN ({asset_types})
  AND ap.enddate <= TO_DATE('2024-12-31', 'YYYY-MM-DD')
  AND tal.accountingbook = 1
```

## Context
- Period: Dec 2024
- Subsidiary: Celigo Inc. (ID: 1) - Consolidated
- Accounting Book: Primary (ID: 1)
- NetSuite Account: 589861
