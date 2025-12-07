# CTA Investigation - Plug Method Still Has Discrepancy

## Current Status (After Implementing Plug Method)

| Component | My Query | Claude's Expected | Difference |
|-----------|----------|-------------------|------------|
| **Total Equity (A-L)** | -$6,417,482.58 | -$6,664,900.80 | +$247,418 |
| **Posted Equity** | $90,416,465.83 | $90,378,938.81 | +$37,527 |
| **Retained Earnings** | -$88,022,956.91 | -$88,022,956.91 | **✅ EXACT** |
| **Net Income** | -$8,781,243.65 | -$8,781,243.65 | **✅ EXACT** |
| **CTA (plug result)** | -$29,747.85 | -$239,639.06 | ~$210K off |

## The Issue

The plug formula is:
```
CTA = Total Equity - Posted Equity - RE - NI
```

My RE and NI are EXACT matches. So the discrepancy comes from:
1. **Total Equity being $247K too high** (less negative)
2. **Posted Equity being $37K too high**

## Key Finding: Elimination Accounts

I found 42 accounts with `eliminate='T'` (intercompany accounts). Their consolidated balances:

| Account Type | Consolidated Balance |
|--------------|---------------------|
| AcctRec (Asset) | +$31,540.57 |
| AcctPay (Liability) | +$3,396.63 |
| DeferRevenue | +$3,000.00 |
| Equity | -$251,256.97 |
| OthCurrLiab | -$61,988.24 |
| OthCurrAsset | -$3,004.54 |
| COGS | $0.00 |
| Income | -$16.20 |
| OthAsset | $0.00 |
| **TOTAL** | **-$278,328.75** |

**Attempted fix:** Exclude accounts where `eliminate='T'` from Assets/Liabilities/Equity queries.

**Result:** Made things WORSE! CTA went to +$248K instead of -$240K.

## My Asset/Liability Values vs NetSuite

| Component | My Value | NetSuite | Difference |
|-----------|----------|----------|------------|
| Total Assets | $53,681,594.33 | $53,322,353.28 | +$359,241 |
| Total Liabilities | $60,099,076.91 | $59,987,254.09 | +$111,823 |
| **Total Equity** | **-$6,417,482.58** | **-$6,664,900.80** | **+$247,418** |

## My Current CTA Queries

### Building the Consolidation Amount
```python
cons_amount = f"""TO_NUMBER(BUILTIN.CONSOLIDATE(tal.amount, 'LEDGER', 'DEFAULT', 'DEFAULT', {target_sub}, t.postingperiod, 'DEFAULT'))"""
```

### Total Assets Query
```sql
SELECT SUM({cons_amount}) AS total_assets
FROM transactionaccountingline tal
JOIN transaction t ON t.id = tal.transaction
JOIN account a ON a.id = tal.account
JOIN accountingperiod ap ON ap.id = t.postingperiod
WHERE t.posting = 'T'
  AND tal.posting = 'T'
  AND a.accttype IN ('Bank', 'AcctRec', 'OthCurrAsset', 'FixedAsset', 'OthAsset', 'DeferExpense', 'UnbilledRec')
  AND ap.enddate <= TO_DATE('2024-12-31', 'YYYY-MM-DD')
  AND tal.accountingbook = 1
```

### Total Liabilities Query
```sql
SELECT SUM({cons_amount} * -1) AS total_liabilities
FROM transactionaccountingline tal
JOIN transaction t ON t.id = tal.transaction
JOIN account a ON a.id = tal.account
JOIN accountingperiod ap ON ap.id = t.postingperiod
WHERE t.posting = 'T'
  AND tal.posting = 'T'
  AND a.accttype IN ('AcctPay', 'CredCard', 'OthCurrLiab', 'LongTermLiab', 'DeferRevenue')
  AND ap.enddate <= TO_DATE('2024-12-31', 'YYYY-MM-DD')
  AND tal.accountingbook = 1
```

### Posted Equity Query (excludes RE/NI/CTA by name)
```sql
SELECT SUM({cons_amount} * -1) AS posted_equity
FROM transactionaccountingline tal
JOIN transaction t ON t.id = tal.transaction
JOIN account a ON a.id = tal.account
JOIN accountingperiod ap ON ap.id = t.postingperiod
WHERE t.posting = 'T'
  AND tal.posting = 'T'
  AND a.accttype IN ('Equity', 'RetainedEarnings')
  AND LOWER(a.fullname) NOT LIKE '%retained earnings%'
  AND LOWER(a.fullname) NOT LIKE '%translation%'
  AND LOWER(a.fullname) NOT LIKE '%cta%'
  AND LOWER(a.fullname) NOT LIKE '%net income%'
  AND LOWER(a.fullname) NOT LIKE '%cumulative translation%'
  AND ap.enddate <= TO_DATE('2024-12-31', 'YYYY-MM-DD')
  AND tal.accountingbook = 1
```

## Questions for Claude

1. **How did you get Total Assets = $53,322,353.28?**
   - My query returns $53,681,594.33 (+$359K more)
   - Are you using different account types?
   - Are you filtering out some accounts I'm not?

2. **How did you get Total Liabilities = $59,987,254.09?**
   - My query returns $60,099,076.91 (+$112K more)

3. **How did you handle elimination accounts?**
   - I tried excluding `eliminate='T'` accounts but it made things worse
   - Should these accounts be INCLUDED in the consolidation query?

4. **What BUILTIN.CONSOLIDATE parameters did you use?**
   - I'm using: `'LEDGER', 'DEFAULT', 'DEFAULT', 1, t.postingperiod, 'DEFAULT'`
   - Is there a different parameter set that handles eliminations properly?

5. **Did you use a different approach entirely?**
   - Maybe query the Balance Summary saved search?
   - Maybe use PERIOD_END rate instead of t.postingperiod?

## Context
- Period: Dec 2024
- Subsidiary: Celigo Inc. (ID: 1) - Consolidated
- Accounting Book: Primary (ID: 1)
- NetSuite Account: 589861
