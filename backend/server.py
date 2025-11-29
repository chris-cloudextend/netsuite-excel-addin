"""
NetSuite Excel Formulas - Backend Server
Flask server that provides REST API for NetSuite SuiteQL queries
"""

from flask import Flask, jsonify, request
from flask_cors import CORS
import json
import requests
from requests_oauthlib import OAuth1
import sys
from datetime import datetime

app = Flask(__name__)
CORS(app)  # Enable CORS for Excel add-in

# Load NetSuite configuration
try:
    with open('netsuite_config.json', 'r') as f:
        config = json.load(f)
except FileNotFoundError:
    print("ERROR: netsuite_config.json not found!")
    print("Please create netsuite_config.json with your NetSuite credentials.")
    sys.exit(1)

account_id = config['account_id']
suiteql_url = f"https://{account_id}.suitetalk.api.netsuite.com/services/rest/query/v1/suiteql"

# Create OAuth1 authentication
auth = OAuth1(
    client_key=config['consumer_key'],
    client_secret=config['consumer_secret'],
    resource_owner_key=config['token_id'],
    resource_owner_secret=config['token_secret'],
    realm=account_id,
    signature_method='HMAC-SHA256'
)


def query_netsuite(sql_query):
    """Execute a SuiteQL query against NetSuite"""
    try:
        response = requests.post(
            suiteql_url,
            auth=auth,
            headers={'Content-Type': 'application/json', 'Prefer': 'transient'},
            json={'q': sql_query},
            timeout=30
        )
        
        if response.status_code == 200:
            return response.json().get('items', [])
        else:
            error_msg = f"NetSuite error: {response.status_code}"
            print(f"=== NetSuite Error ===", file=sys.stderr)
            print(f"Query: {sql_query[:200]}...", file=sys.stderr)
            print(f"Status: {response.status_code}", file=sys.stderr)
            print(f"Response: {response.text}", file=sys.stderr)
            print(f"=====================", file=sys.stderr)
            return {'error': error_msg, 'details': response.text}
            
    except Exception as e:
        print(f"Exception querying NetSuite: {str(e)}", file=sys.stderr)
        return {'error': str(e)}


def escape_sql(text):
    """Escape single quotes in SQL strings"""
    if text is None:
        return ""
    return str(text).replace("'", "''")


@app.route('/')
def home():
    """Health check endpoint"""
    return jsonify({
        'status': 'running',
        'service': 'NetSuite Excel Formulas API',
        'account': account_id,
        'version': '1.0',
        'endpoints': {
            '/account/{account_number}/name': 'Get account name (NSGLATITLE)',
            '/balance': 'Get GL balance (NSGLABAL)',
            '/budget': 'Get budget amount (NSGLABUD)',
            '/health': 'Health check'
        }
    })


@app.route('/health')
def health():
    """Health check"""
    return jsonify({'status': 'healthy', 'account': account_id})


@app.route('/batch/balance', methods=['POST'])
def batch_balance():
    """
    BATCH ENDPOINT - Get balances for MULTIPLE accounts and periods in ONE call
    This is much faster than individual requests!
    
    POST JSON:
    {
        "accounts": ["4010", "5000", "6000"],
        "periods": ["Jan 2025", "Feb 2025", "Mar 2025"],
        "subsidiary": "",
        "class": "",
        "department": "13",
        "location": ""
    }
    
    Returns:
    {
        "balances": {
            "4010": {
                "Jan 2025": 12400000,
                "Feb 2025": 13200000
            },
            "5000": {
                "Jan 2025": 5000000,
                "Feb 2025": 5200000
            }
        }
    }
    """
    data = request.get_json()
    
    if not data or 'accounts' not in data or 'periods' not in data:
        return jsonify({'error': 'accounts and periods required'}), 400
    
    accounts = data.get('accounts', [])
    periods = data.get('periods', [])
    subsidiary = data.get('subsidiary', '')
    class_id = data.get('class', '')
    department = data.get('department', '')
    location = data.get('location', '')
    
    if not accounts or not periods:
        return jsonify({'error': 'accounts and periods must be non-empty'}), 400
    
    try:
        # Build WHERE clause with optional filters
        where_clauses = [
            "t.posting = 'T'",
            "tal.posting = 'T'"
        ]
        
        # Add accounts IN clause
        accounts_in = ','.join([f"'{escape_sql(acc)}'" for acc in accounts])
        where_clauses.append(f"a.acctnumber IN ({accounts_in})")
        
        # Add optional filters
        if subsidiary and subsidiary != '':
            where_clauses.append(f"t.subsidiary = {subsidiary}")
        
        # Need TransactionLine join if filtering by department, class, or location
        needs_line_join = (department and department != '') or (class_id and class_id != '') or (location and location != '')
        
        if class_id and class_id != '':
            where_clauses.append(f"tl.class = {class_id}")
        
        if department and department != '':
            where_clauses.append(f"tl.department = {department}")
        
        if location and location != '':
            where_clauses.append(f"tl.location = {location}")
        
        # Build periods IN clause
        periods_in = ','.join([f"'{escape_sql(p)}'" for p in periods])
        where_clauses.append(f"ap.periodname IN ({periods_in})")
        
        where_clause = " AND ".join(where_clauses)
        
        # Build query with TransactionLine join if needed
        if needs_line_join:
            query = f"""
                SELECT 
                    a.acctnumber,
                    ap.periodname,
                    a.accttype,
                    CASE 
                        WHEN a.accttype IN ('Income', 'Other Income', 'OthIncome', 'Liability', 
                                            'LongTermLiab', 'OthCurrLiab', 'Equity') 
                        THEN SUM(COALESCE(tal.credit, 0)) - SUM(COALESCE(tal.debit, 0))
                        ELSE SUM(COALESCE(tal.debit, 0)) - SUM(COALESCE(tal.credit, 0))
                    END AS balance
                FROM Transaction t
                INNER JOIN TransactionLine tl ON t.id = tl.transaction
                INNER JOIN TransactionAccountingLine tal ON t.id = tal.transaction AND tl.id = tal.transactionline
                INNER JOIN Account a ON tal.account = a.id
                INNER JOIN AccountingPeriod ap ON t.postingperiod = ap.id
                WHERE {where_clause}
                GROUP BY a.acctnumber, ap.periodname, a.accttype
                ORDER BY a.acctnumber, ap.periodname
            """
        else:
            query = f"""
                SELECT 
                    a.acctnumber,
                    ap.periodname,
                    a.accttype,
                    CASE 
                        WHEN a.accttype IN ('Income', 'Other Income', 'OthIncome', 'Liability', 
                                            'LongTermLiab', 'OthCurrLiab', 'Equity') 
                        THEN SUM(COALESCE(tal.credit, 0)) - SUM(COALESCE(tal.debit, 0))
                        ELSE SUM(COALESCE(tal.debit, 0)) - SUM(COALESCE(tal.credit, 0))
                    END AS balance
                FROM Transaction t
                INNER JOIN TransactionAccountingLine tal ON t.id = tal.transaction
                INNER JOIN Account a ON tal.account = a.id
                INNER JOIN AccountingPeriod ap ON t.postingperiod = ap.id
                WHERE {where_clause}
                GROUP BY a.acctnumber, ap.periodname, a.accttype
                ORDER BY a.acctnumber, ap.periodname
            """
        
        print(f"DEBUG - Batch query:\n{query}", file=sys.stderr)
        result = query_netsuite(query)
        
        # Check for errors
        if isinstance(result, dict) and 'error' in result:
            return jsonify(result), 500
        
        # Organize results by account, then period
        balances = {}
        for row in result:
            account_num = row['acctnumber']
            period = row['periodname']
            balance = float(row['balance']) if row['balance'] else 0
            
            if account_num not in balances:
                balances[account_num] = {}
            
            if period in balances[account_num]:
                balances[account_num][period] += balance
            else:
                balances[account_num][period] = balance
        
        # Fill in zeros for missing account/period combinations
        for account_num in accounts:
            if account_num not in balances:
                balances[account_num] = {}
            for period in periods:
                if period not in balances[account_num]:
                    balances[account_num][period] = 0
        
        return jsonify({'balances': balances})
        
    except Exception as e:
        print(f"Error in batch_balance: {str(e)}", file=sys.stderr)
        return jsonify({'error': str(e)}), 500


@app.route('/department/<department_name>')
def get_department_id(department_name):
    """
    Get department ID from department name
    Used to lookup department IDs for filtering
    
    Returns: Department ID (number) or error
    """
    try:
        query = f"""
            SELECT id, name
            FROM Department
            WHERE LOWER(name) LIKE LOWER('%{escape_sql(department_name)}%')
        """
        
        result = query_netsuite(query)
        
        if isinstance(result, dict) and 'error' in result:
            return jsonify({'error': result['error']}), 500
        
        if result and len(result) > 0:
            return jsonify(result)
        else:
            return jsonify({'error': 'Department not found'}), 404
            
    except Exception as e:
        print(f"Error in get_department_id: {str(e)}", file=sys.stderr)
        return jsonify({'error': str(e)}), 500


@app.route('/account/<account_number>/name')
def get_account_name(account_number):
    """
    Get account name from account number
    Used by: NS.GLATITLE(accountNumber)
    
    Returns: Account display name (string)
    """
    try:
        # Build SuiteQL query
        query = f"""
            SELECT a.acctname AS account_name
            FROM Account a
            WHERE a.acctnumber = '{escape_sql(account_number)}'
        """
        
        result = query_netsuite(query)
        
        # Check for errors
        if isinstance(result, dict) and 'error' in result:
            return jsonify({'error': result['error']}), 500
        
        # Return account name or "Not Found"
        if result and len(result) > 0:
            return result[0].get('account_name', 'Not Found')
        else:
            return 'Not Found'
            
    except Exception as e:
        print(f"Error in get_account_name: {str(e)}", file=sys.stderr)
        return jsonify({'error': str(e)}), 500


@app.route('/balance')
def get_balance():
    """
    Get GL account balance with filters
    Used by: NS.GLABAL(subsidiary, account, fromPeriod, toPeriod, class, dept, location)
    
    Query params:
        - account: Account number (required)
        - subsidiary: Subsidiary ID (optional)
        - from_period: Starting period name (optional)
        - to_period: Ending period name (optional)
        - class: Class ID (optional)
        - department: Department ID (optional)
        - location: Location ID (optional)
    
    Returns: Balance amount (number)
    """
    try:
        # Get parameters
        account = request.args.get('account', '')
        subsidiary = request.args.get('subsidiary', '')
        from_period = request.args.get('from_period', '')
        to_period = request.args.get('to_period', '')
        class_id = request.args.get('class', '')
        department = request.args.get('department', '')
        location = request.args.get('location', '')
        
        if not account:
            return jsonify({'error': 'Account number required'}), 400
        
        # Build WHERE clause
        where_clauses = [
            "t.posting = 'T'",
            "tal.posting = 'T'",
            f"a.acctnumber = '{escape_sql(account)}'"
        ]
        
        # Add optional filters
        if subsidiary and subsidiary != '':
            where_clauses.append(f"t.subsidiary = {subsidiary}")
        
        # Handle period filters - support both period IDs and names
        if from_period and to_period:
            # Check if it's a number (period ID) or text (period name)
            if from_period.isdigit() and to_period.isdigit():
                where_clauses.append(f"t.postingperiod >= {from_period}")
                where_clauses.append(f"t.postingperiod <= {to_period}")
            else:
                # Use period names (requires join with AccountingPeriod)
                where_clauses.append(f"ap.periodname >= '{escape_sql(from_period)}'")
                where_clauses.append(f"ap.periodname <= '{escape_sql(to_period)}'")
        elif from_period:
            if from_period.isdigit():
                where_clauses.append(f"t.postingperiod = {from_period}")
            else:
                where_clauses.append(f"ap.periodname = '{escape_sql(from_period)}'")
        
        if class_id and class_id != '':
            where_clauses.append(f"tl.class = {class_id}")
        
        if department and department != '':
            # Department is on TransactionLine table for journal entries
            where_clauses.append(f"tl.department = {department}")
        
        if location and location != '':
            where_clauses.append(f"tl.location = {location}")
        
        where_clause = " AND ".join(where_clauses)
        
        # Build SuiteQL query - use CASE for correct balance by account type
        # Only join AccountingPeriod if we're using period names
        # Note: Department filtering requires TransactionLine join for journal entries
        print(f"DEBUG - WHERE clause: {where_clause}", file=sys.stderr)
        print(f"DEBUG - Department param: {department}", file=sys.stderr)
        
        # Need TransactionLine join if filtering by department, class, or location
        needs_line_join = (department and department != '') or (class_id and class_id != '') or (location and location != '')
        
        if (from_period and not from_period.isdigit()) or (to_period and not to_period.isdigit()):
            if needs_line_join:
                query = f"""
                    SELECT 
                        a.accttype,
                        CASE 
                            WHEN a.accttype IN ('Income', 'Other Income', 'OthIncome', 'Liability', 
                                                'LongTermLiab', 'OthCurrLiab', 'Equity') 
                            THEN SUM(COALESCE(tal.credit, 0)) - SUM(COALESCE(tal.debit, 0))
                            ELSE SUM(COALESCE(tal.debit, 0)) - SUM(COALESCE(tal.credit, 0))
                        END AS balance
                    FROM Transaction t
                    INNER JOIN TransactionLine tl ON t.id = tl.transaction  
                    INNER JOIN TransactionAccountingLine tal ON t.id = tal.transaction AND tl.id = tal.transactionline
                    INNER JOIN Account a ON tal.account = a.id
                    INNER JOIN AccountingPeriod ap ON t.postingperiod = ap.id
                    WHERE {where_clause}
                    GROUP BY a.accttype
                """
            else:
                query = f"""
                    SELECT 
                        a.accttype,
                        CASE 
                            WHEN a.accttype IN ('Income', 'Other Income', 'OthIncome', 'Liability', 
                                                'LongTermLiab', 'OthCurrLiab', 'Equity') 
                            THEN SUM(COALESCE(tal.credit, 0)) - SUM(COALESCE(tal.debit, 0))
                            ELSE SUM(COALESCE(tal.debit, 0)) - SUM(COALESCE(tal.credit, 0))
                        END AS balance
                    FROM Transaction t
                    INNER JOIN TransactionAccountingLine tal ON t.id = tal.transaction
                    INNER JOIN Account a ON tal.account = a.id
                    INNER JOIN AccountingPeriod ap ON t.postingperiod = ap.id
                    WHERE {where_clause}
                    GROUP BY a.accttype
                """
        else:
            if needs_line_join:
                query = f"""
                    SELECT 
                        a.accttype,
                        CASE 
                            WHEN a.accttype IN ('Income', 'Other Income', 'OthIncome', 'Liability', 
                                                'LongTermLiab', 'OthCurrLiab', 'Equity') 
                            THEN SUM(COALESCE(tal.credit, 0)) - SUM(COALESCE(tal.debit, 0))
                            ELSE SUM(COALESCE(tal.debit, 0)) - SUM(COALESCE(tal.credit, 0))
                        END AS balance
                    FROM Transaction t
                    INNER JOIN TransactionLine tl ON t.id = tl.transaction
                    INNER JOIN TransactionAccountingLine tal ON t.id = tal.transaction AND tl.id = tal.transactionline
                    INNER JOIN Account a ON tal.account = a.id
                    WHERE {where_clause}
                    GROUP BY a.accttype
                """
            else:
                query = f"""
                    SELECT 
                        a.accttype,
                        CASE 
                            WHEN a.accttype IN ('Income', 'Other Income', 'OthIncome', 'Liability', 
                                                'LongTermLiab', 'OthCurrLiab', 'Equity') 
                            THEN SUM(COALESCE(tal.credit, 0)) - SUM(COALESCE(tal.debit, 0))
                            ELSE SUM(COALESCE(tal.debit, 0)) - SUM(COALESCE(tal.credit, 0))
                        END AS balance
                    FROM Transaction t
                    INNER JOIN TransactionAccountingLine tal ON t.id = tal.transaction
                    INNER JOIN Account a ON tal.account = a.id
                    WHERE {where_clause}
                    GROUP BY a.accttype
                """
        
        print(f"DEBUG - Full query:\n{query}", file=sys.stderr)
        result = query_netsuite(query)
        
        # Check for errors
        if isinstance(result, dict) and 'error' in result:
            return jsonify({'error': result['error']}), 500
        
        # Return balance (default to 0 if no data)
        if result and len(result) > 0:
            balance = result[0].get('balance')
            if balance is None:
                return '0'
            return str(float(balance))
        else:
            return '0'
            
    except Exception as e:
        print(f"Error in get_balance: {str(e)}", file=sys.stderr)
        return jsonify({'error': str(e)}), 500


@app.route('/budget')
def get_budget():
    """
    Get budget amount with filters
    Used by: NS.GLABUD(subsidiary, budgetCategory, account, fromPeriod, toPeriod, class, dept, location)
    
    Query params:
        - account: Account number (required)
        - subsidiary: Subsidiary ID (optional)
        - budget_category: Budget category name (optional - not currently filtered)
        - from_period: Starting period name (optional)
        - to_period: Ending period name (optional)
        - class: Class ID (optional)
        - department: Department ID (optional)
        - location: Location ID (optional)
    
    Returns: Budget amount (number)
    """
    try:
        # Get parameters
        account = request.args.get('account', '')
        subsidiary = request.args.get('subsidiary', '')
        budget_category = request.args.get('budget_category', '')  # Future enhancement
        from_period = request.args.get('from_period', '')
        to_period = request.args.get('to_period', '')
        class_id = request.args.get('class', '')
        department = request.args.get('department', '')
        location = request.args.get('location', '')
        
        if not account:
            return jsonify({'error': 'Account number required'}), 400
        
        # Build WHERE clause
        where_clauses = [
            f"a.acctnumber = '{escape_sql(account)}'"
        ]
        
        # Add optional filters
        if subsidiary and subsidiary != '':
            where_clauses.append(f"b.subsidiary = {subsidiary}")
        
        # Handle period filters - support both period IDs and names
        if from_period and to_period:
            if from_period.isdigit() and to_period.isdigit():
                where_clauses.append(f"b.accountingperiod >= {from_period}")
                where_clauses.append(f"b.accountingperiod <= {to_period}")
            else:
                where_clauses.append(f"ap.periodname >= '{escape_sql(from_period)}'")
                where_clauses.append(f"ap.periodname <= '{escape_sql(to_period)}'")
        elif from_period:
            if from_period.isdigit():
                where_clauses.append(f"b.accountingperiod = {from_period}")
            else:
                where_clauses.append(f"ap.periodname = '{escape_sql(from_period)}'")
        
        if class_id and class_id != '':
            where_clauses.append(f"b.class = {class_id}")
        
        if department and department != '':
            where_clauses.append(f"b.department = {department}")
        
        if location and location != '':
            where_clauses.append(f"b.location = {location}")
        
        where_clause = " AND ".join(where_clauses)
        
        # Build SuiteQL query - only join AccountingPeriod if using period names
        if (from_period and not from_period.isdigit()) or (to_period and not to_period.isdigit()):
            query = f"""
                SELECT SUM(b.amount) AS budget_amount
                FROM Budget b
                INNER JOIN Account a ON b.account = a.id
                INNER JOIN AccountingPeriod ap ON b.accountingperiod = ap.id
                WHERE {where_clause}
            """
        else:
            query = f"""
                SELECT SUM(b.amount) AS budget_amount
                FROM Budget b
                INNER JOIN Account a ON b.account = a.id
                WHERE {where_clause}
            """
        
        result = query_netsuite(query)
        
        # Check for errors - Budget table may not exist or have no data in test accounts
        if isinstance(result, dict) and 'error' in result:
            # Return 0 if budget table doesn't exist (common in test accounts)
            print(f"Budget query failed (this is normal for test accounts): {result.get('error')}", file=sys.stderr)
            return '0'
        
        # Return budget amount (default to 0 if no data)
        if result and len(result) > 0:
            budget = result[0].get('budget_amount')
            if budget is None:
                return '0'
            return str(float(budget))
        else:
            return '0'
            
    except Exception as e:
        print(f"Error in get_budget: {str(e)}", file=sys.stderr)
        return jsonify({'error': str(e)}), 500


@app.route('/transactions', methods=['GET'])
def get_transactions():
    """
    Get transaction-level details for drill-down
    Used for: Drill-down from balance cells to see underlying transactions
    
    Query params:
        - account: Account number (required)
        - period: Period name (required)
        - subsidiary: Subsidiary ID (optional)
        - class: Class ID (optional)
        - department: Department ID (optional)
        - location: Location ID (optional)
    
    Returns: JSON with transaction details including NetSuite URLs
    """
    try:
        account = request.args.get('account')
        period = request.args.get('period')
        subsidiary = request.args.get('subsidiary', '')
        class_id = request.args.get('class', '')
        department = request.args.get('department', '')
        location = request.args.get('location', '')
        
        if not account or not period:
            return jsonify({'error': 'Missing required parameters: account and period'}), 400
        
        # Build WHERE clause with filters
        where_conditions = [
            "t.posting = 'T'",
            "tal.posting = 'T'",
            f"a.acctnumber = '{escape_sql(account)}'",
            f"ap.periodname = '{escape_sql(period)}'"
        ]
        
        if subsidiary:
            where_conditions.append(f"t.subsidiary = {subsidiary}")
        
        # Need TransactionLine join if filtering by department, class, or location
        needs_line_join = (department and department != '') or (class_id and class_id != '') or (location and location != '')
        
        if class_id:
            where_conditions.append(f"tl.class = {class_id}")
        if department:
            where_conditions.append(f"tl.department = {department}")
        if location:
            where_conditions.append(f"tl.location = {location}")
        
        where_clause = " AND ".join(where_conditions)
        
        # SuiteQL query for transaction details
        # GROUP BY transaction to avoid duplicates when multiple lines hit same account
        if needs_line_join:
            query = f"""
                SELECT 
                    t.id AS transaction_id,
                    t.tranid AS transaction_number,
                    t.trandisplayname AS transaction_type,
                    t.recordtype AS record_type,
                    TO_CHAR(t.trandate, 'YYYY-MM-DD') AS transaction_date,
                    e.entityid AS entity_name,
                    e.id AS entity_id,
                    t.memo,
                    SUM(COALESCE(tal.debit, 0)) AS debit,
                    SUM(COALESCE(tal.credit, 0)) AS credit,
                    a.acctnumber AS account_number,
                    a.accountsearchdisplayname AS account_name
                FROM 
                    Transaction t
                INNER JOIN 
                    TransactionLine tl ON t.id = tl.transaction
                INNER JOIN 
                    TransactionAccountingLine tal ON t.id = tal.transaction AND tl.id = tal.transactionline
                INNER JOIN 
                    Account a ON tal.account = a.id
                INNER JOIN
                    AccountingPeriod ap ON t.postingperiod = ap.id
                LEFT JOIN
                    Entity e ON t.entity = e.id
                WHERE 
                    {where_clause}
                GROUP BY
                    t.id, t.tranid, t.trandisplayname, t.recordtype, t.trandate,
                    e.entityid, e.id, t.memo, a.acctnumber, a.accountsearchdisplayname
                ORDER BY
                    t.trandate, t.tranid
            """
        else:
            query = f"""
                SELECT 
                    t.id AS transaction_id,
                    t.tranid AS transaction_number,
                    t.trandisplayname AS transaction_type,
                    t.recordtype AS record_type,
                    TO_CHAR(t.trandate, 'YYYY-MM-DD') AS transaction_date,
                    e.entityid AS entity_name,
                    e.id AS entity_id,
                    t.memo,
                    SUM(COALESCE(tal.debit, 0)) AS debit,
                    SUM(COALESCE(tal.credit, 0)) AS credit,
                    a.acctnumber AS account_number,
                    a.accountsearchdisplayname AS account_name
                FROM 
                    Transaction t
                INNER JOIN 
                    TransactionAccountingLine tal ON t.id = tal.transaction
                INNER JOIN 
                    Account a ON tal.account = a.id
                INNER JOIN
                    AccountingPeriod ap ON t.postingperiod = ap.id
                LEFT JOIN
                    Entity e ON t.entity = e.id
                WHERE 
                    {where_clause}
                GROUP BY
                    t.id, t.tranid, t.trandisplayname, t.recordtype, t.trandate,
                    e.entityid, e.id, t.memo, a.acctnumber, a.accountsearchdisplayname
                ORDER BY
                    t.trandate, t.tranid
            """
        
        print(f"DEBUG - Transaction drill-down query:\n{query}", file=sys.stderr)
        result = query_netsuite(query)
        
        if isinstance(result, dict) and 'error' in result:
            return jsonify(result), 500
        
        # Add NetSuite URL to each transaction
        for row in result:
            transaction_id = row.get('transaction_id')
            record_type = row.get('record_type', '').lower()
            
            # Map record types to NetSuite URL paths
            type_map = {
                'invoice': 'custinvc',
                'bill': 'vendorbill',
                'journalentry': 'journal',
                'journal': 'journal',
                'payment': 'custpymt',
                'vendorpayment': 'vendpymt',
                'creditmemo': 'custcred',
                'vendorcredit': 'vendcred',
                'check': 'check',
                'deposit': 'deposit',
                'cashsale': 'cashsale',
                'cashrefund': 'cashrfnd',
                'expensereport': 'exprept'
            }
            
            url_type = type_map.get(record_type, record_type)
            row['netsuite_url'] = f"https://{account_id}.app.netsuite.com/app/accounting/transactions/{url_type}.nl?id={transaction_id}"
            
            # Calculate net amount for this account
            debit = float(row.get('debit', 0)) if row.get('debit') else 0
            credit = float(row.get('credit', 0)) if row.get('credit') else 0
            row['net_amount'] = debit - credit
        
        return jsonify({
            'transactions': result,
            'count': len(result),
            'filters': {
                'account': account,
                'period': period,
                'subsidiary': subsidiary,
                'class': class_id,
                'department': department,
                'location': location
            }
        })
        
    except Exception as e:
        print(f"Error in get_transactions: {str(e)}", file=sys.stderr)
        return jsonify({'error': str(e)}), 500


@app.route('/test')
def test_connection():
    """Test NetSuite connection"""
    try:
        # Simple query to test connection
        query = "SELECT COUNT(*) as count FROM Account WHERE isinactive = 'F'"
        result = query_netsuite(query)
        
        if isinstance(result, dict) and 'error' in result:
            return jsonify({
                'status': 'error',
                'error': result['error'],
                'details': result.get('details', '')
            }), 500
        
        return jsonify({
            'status': 'success',
            'message': 'NetSuite connection successful',
            'account': account_id,
            'active_accounts': result[0].get('count', 0) if result else 0
        })
        
    except Exception as e:
        return jsonify({
            'status': 'error',
            'error': str(e)
        }), 500


# ============================================================================
# LOOKUP ENDPOINTS - For Excel dropdowns/data validation
# ============================================================================

@app.route('/lookups/all')
def get_all_lookups():
    """
    Get all lookups at once - Subsidiary, Department, Location, Class
    Returns actual IDs from transaction data
    """
    try:
        # Start with known values (you can add more manually)
        lookups = {
            'subsidiaries': [
                {'id': '1', 'name': 'Parent Company'},
                # Add more as you discover them
            ],
            'departments': [
                {'id': '13', 'name': 'Demo'},
                # Will populate from NetSuite below
            ],
            'classes': [
                # Will populate from NetSuite below
            ],
            'locations': [
                # Will populate from NetSuite below
            ]
        }
        
        # Try to fetch actual values from NetSuite (might fail on some accounts)
        try:
            # Get unique department IDs from recent transactions
            dept_query = """
                SELECT DISTINCT tl.department as id
                FROM TransactionLine tl
                WHERE tl.department IS NOT NULL
                AND tl.department != 0
                AND ROWNUM <= 100
            """
            dept_result = query_netsuite(dept_query)
            if isinstance(dept_result, list) and len(dept_result) > 0:
                # Clear hardcoded and use real data
                lookups['departments'] = []
                for row in dept_result:
                    dept_id = str(row['id'])
                    lookups['departments'].append({
                        'id': dept_id,
                        'name': f'Department {dept_id}'
                    })
        except Exception as e:
            print(f"Department query error: {e}")
        
        try:
            # Get unique subsidiary IDs from transactions (actual usage)
            sub_query = """
                SELECT DISTINCT t.subsidiary as id
                FROM Transaction t
                WHERE t.subsidiary IS NOT NULL
                AND t.subsidiary != 0
                AND ROWNUM <= 100
            """
            sub_result = query_netsuite(sub_query)
            
            if isinstance(sub_result, list) and len(sub_result) > 0:
                # Clear the hardcoded list and use real data
                lookups['subsidiaries'] = []
                for row in sub_result:
                    sub_id = str(row['id'])
                    lookups['subsidiaries'].append({
                        'id': sub_id,
                        'name': f'Subsidiary {sub_id}'
                    })
                print(f"Found {len(lookups['subsidiaries'])} subsidiaries")
        except Exception as e:
            print(f"Subsidiary query error: {e}")
        
        try:
            # Get unique class IDs
            class_query = """
                SELECT DISTINCT tl.class as id
                FROM TransactionLine tl
                WHERE tl.class IS NOT NULL
                AND tl.class != 0
                AND ROWNUM <= 100
            """
            class_result = query_netsuite(class_query)
            if isinstance(class_result, list) and len(class_result) > 0:
                for row in class_result:
                    class_id = str(row['id'])
                    lookups['classes'].append({
                        'id': class_id,
                        'name': f'Class {class_id}'
                    })
        except Exception as e:
            print(f"Class query error: {e}")
        
        try:
            # Get unique location IDs
            loc_query = """
                SELECT DISTINCT tl.location as id
                FROM TransactionLine tl
                WHERE tl.location IS NOT NULL
                AND tl.location != 0
                AND ROWNUM <= 100
            """
            loc_result = query_netsuite(loc_query)
            if isinstance(loc_result, list) and len(loc_result) > 0:
                for row in loc_result:
                    loc_id = str(row['id'])
                    lookups['locations'].append({
                        'id': loc_id,
                        'name': f'Location {loc_id}'
                    })
        except Exception as e:
            print(f"Location query error: {e}")
        
        return jsonify(lookups)
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    print("=" * 80)
    print("NetSuite Excel Formulas - Backend Server")
    print("=" * 80)
    print()
    print(f"NetSuite Account: {account_id}")
    print(f"Server starting on: http://localhost:5002")
    print()
    print("Endpoints:")
    print("  GET  /                              - Service info")
    print("  GET  /health                        - Health check")
    print("  GET  /test                          - Test NetSuite connection")
    print("  GET  /account/<number>/name         - Get account name")
    print("  GET  /balance?account=...           - Get GL balance")
    print("  GET  /budget?account=...            - Get budget amount")
    print("  POST /batch/balance                 - Batch balance queries")
    print("  GET  /transactions?account=...      - Transaction drill-down")
    print("  GET  /lookups/subsidiaries          - Get subsidiaries list")
    print("  GET  /lookups/departments           - Get departments list")
    print("  GET  /lookups/classes               - Get classes list")
    print("  GET  /lookups/locations             - Get locations list")
    print()
    print("Press Ctrl+C to stop")
    print("=" * 80)
    print()
    
    # Run server
    app.run(host='127.0.0.1', port=5002, debug=False)

