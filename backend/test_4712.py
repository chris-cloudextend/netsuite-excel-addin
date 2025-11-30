#!/usr/bin/env python3
"""Test account 4712 with different date ranges"""

import requests
from requests_oauthlib import OAuth1
import json

# Production account
ACCOUNT_ID = "589861"
CONSUMER_KEY = "a432d93c007d27337151ee010d734bb9858556addc3d5961923fbf15ef2d8246"
CONSUMER_SECRET = "953405056dc879569d03e074422ca1792bf5c34f4d724fdafd001dbf6a8e5df3"
TOKEN_ID = "fd15642ac1360a727dee4076a137e1825a987e4b07d6216a1f7674311dfd7db0"
TOKEN_SECRET = "8138f3291f9a1fa97dd8aacda7bca1a0d3bff876dc3e732a35abbf8366608d2d"

auth = OAuth1(
    client_key=CONSUMER_KEY,
    client_secret=CONSUMER_SECRET,
    resource_owner_key=TOKEN_ID,
    resource_owner_secret=TOKEN_SECRET,
    realm=ACCOUNT_ID,
    signature_method='HMAC-SHA256'
)

suiteql_url = f"https://{ACCOUNT_ID}.suitetalk.api.netsuite.com/services/rest/query/v1/suiteql"

def query_netsuite(query):
    response = requests.post(
        suiteql_url,
        auth=auth,
        headers={'Content-Type': 'application/json', 'Prefer': 'transient'},
        json={'q': query},
        timeout=30
    )
    
    if response.status_code == 200:
        return response.json().get('items', [])
    else:
        return {'error': response.status_code, 'details': response.text}

print("=" * 80)
print("Testing Account 4712 - Date Range Issue")
print("=" * 80)
print()

# Test 1: What periods have data?
print("Test 1: What periods have transactions for account 4712?")
print("-" * 40)
query1 = """
    SELECT DISTINCT ap.periodname, COUNT(*) as transactions
    FROM Transaction t
    INNER JOIN TransactionAccountingLine tal ON t.id = tal.transaction
    INNER JOIN Account a ON tal.account = a.id
    INNER JOIN AccountingPeriod ap ON t.postingperiod = ap.id
    WHERE t.posting = 'T' 
      AND tal.posting = 'T'
      AND a.acctnumber = '4712'
    GROUP BY ap.periodname
    ORDER BY ap.periodname
"""

result = query_netsuite(query1)
if isinstance(result, list):
    print(f"Found data in {len(result)} periods:")
    for row in result:
        print(f"  - {row['periodname']}: {row['transactions']} transactions")
else:
    print(f"Error: {result}")
print()

# Test 2: Mar 2025 balance (should match Excel)
print("Test 2: Mar 2025 balance")
print("-" * 40)
query2 = """
    SELECT 
        CASE 
            WHEN a.accttype IN ('Income', 'OthIncome') 
            THEN SUM(COALESCE(tal.credit, 0)) - SUM(COALESCE(tal.debit, 0))
            ELSE SUM(COALESCE(tal.debit, 0)) - SUM(COALESCE(tal.credit, 0))
        END AS balance
    FROM Transaction t
    INNER JOIN TransactionAccountingLine tal ON t.id = tal.transaction
    INNER JOIN Account a ON tal.account = a.id
    INNER JOIN AccountingPeriod ap ON t.postingperiod = ap.id
    WHERE t.posting = 'T' 
      AND tal.posting = 'T'
      AND a.acctnumber = '4712'
      AND ap.periodname >= 'Mar 2025'
      AND ap.periodname <= 'Mar 2025'
    GROUP BY a.accttype
"""

result = query_netsuite(query2)
if isinstance(result, list) and len(result) > 0:
    print(f"Balance: {result[0].get('balance', 0)}")
else:
    print(f"Error or no data: {result}")
print()

# Test 3: Jan-Dec 2025 balance
print("Test 3: Jan 2025 to Dec 2025 balance")
print("-" * 40)
query3 = """
    SELECT 
        CASE 
            WHEN a.accttype IN ('Income', 'OthIncome') 
            THEN SUM(COALESCE(tal.credit, 0)) - SUM(COALESCE(tal.debit, 0))
            ELSE SUM(COALESCE(tal.debit, 0)) - SUM(COALESCE(tal.credit, 0))
        END AS balance
    FROM Transaction t
    INNER JOIN TransactionAccountingLine tal ON t.id = tal.transaction
    INNER JOIN Account a ON tal.account = a.id
    INNER JOIN AccountingPeriod ap ON t.postingperiod = ap.id
    WHERE t.posting = 'T' 
      AND tal.posting = 'T'
      AND a.acctnumber = '4712'
      AND ap.periodname >= 'Jan 2025'
      AND ap.periodname <= 'Dec 2025'
    GROUP BY a.accttype
"""

result = query_netsuite(query3)
if isinstance(result, list) and len(result) > 0:
    print(f"Balance: {result[0].get('balance', 0)}")
    print(f"Expected: 3,101,698.76")
else:
    print(f"Error or no data: {result}")
print()

# Test 4: Check period name format
print("Test 4: Checking period names in NetSuite")
print("-" * 40)
query4 = """
    SELECT DISTINCT periodname 
    FROM AccountingPeriod 
    WHERE periodname LIKE '%2025%'
    ORDER BY periodname
"""

result = query_netsuite(query4)
if isinstance(result, list):
    print(f"2025 periods found ({len(result)} total):")
    for row in result[:15]:
        print(f"  - {row['periodname']}")
else:
    print(f"Error: {result}")

print()
print("=" * 80)

