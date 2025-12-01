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
from dateutil import parser as date_parser
from dateutil.relativedelta import relativedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

app = Flask(__name__)
CORS(app)  # Enable CORS for Excel add-in

# In-memory cache for name-to-ID lookups (refreshes on server restart)
lookup_cache = {
    'subsidiaries': {},  # name → id
    'departments': {},   # name → id
    'classes': {},       # name → id
    'locations': {},     # name → id
    'periods': {}        # period name → id (for date range performance)
}
cache_loaded = False

# Default subsidiary ID (top-level parent) - loaded at startup
# This is used when no subsidiary is specified by the user
default_subsidiary_id = None

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


def query_netsuite(sql_query, timeout=30):
    """Execute a SuiteQL query against NetSuite
    
    Args:
        sql_query: The SuiteQL query to execute
        timeout: Request timeout in seconds (default 30, increase for complex BS queries)
    """
    try:
        response = requests.post(
            suiteql_url,
            auth=auth,
            headers={'Content-Type': 'application/json', 'Prefer': 'transient'},
            json={'q': sql_query},
            timeout=timeout
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


def is_balance_sheet_account(accttype):
    """
    Determine if an account type is a Balance Sheet account.
    
    Balance Sheet accounts are cumulative (beginning of time through period end).
    P&L accounts are for a specific period only.
    
    Args:
        accttype: Account type from NetSuite (e.g., 'Bank', 'Income', 'Expense')
        
    Returns:
        True if Balance Sheet account, False if P&L account
    """
    # P&L account types (Income Statement)
    pl_types = {
        'Income', 'OthIncome', 'Other Income',
        'COGS', 'Cost of Goods Sold',
        'Expense', 'OthExpense', 'Other Expense'
    }
    
    # If it's a P&L type, return False (not balance sheet)
    return accttype not in pl_types


def get_period_dates_from_name(period_name):
    """Convert period name (e.g., 'Mar 2025') to start/end dates for proper date range queries
    Returns tuple: (startdate, enddate) or (None, None) if not found
    Uses cache for performance (avoids repeated NetSuite queries)"""
    
    # Check cache first
    cache_key = f"{period_name}_dates"
    if cache_key in lookup_cache['periods']:
        return lookup_cache['periods'][cache_key]
    
    try:
        query = f"""
            SELECT startdate, enddate, id
            FROM AccountingPeriod
            WHERE periodname = '{escape_sql(period_name)}'
            AND isquarter = 'F'
            AND isyear = 'F'
            AND ROWNUM = 1
        """
        result = query_netsuite(query)
        if isinstance(result, list) and len(result) > 0:
            dates = (result[0].get('startdate'), result[0].get('enddate'), result[0].get('id'))
            # Cache it
            lookup_cache['periods'][cache_key] = dates
            return dates
        return (None, None, None)
    except Exception as e:
        print(f"Error getting period dates for '{period_name}': {e}", file=sys.stderr)
        return (None, None, None)


def get_months_between_periods(from_period, to_period):
    """Calculate the number of months between two periods
    Returns number of months, or 0 if calculation fails"""
    try:
        from_start, _ = get_period_dates_from_name(from_period)
        _, to_end = get_period_dates_from_name(to_period)
        
        if not from_start or not to_end:
            return 0
        
        # Parse dates (NetSuite returns dates like "1/1/2025")
        start = date_parser.parse(from_start)
        end = date_parser.parse(to_end)
        
        # Calculate months difference
        months = (end.year - start.year) * 12 + (end.month - start.month) + 1
        return months
    except Exception as e:
        print(f"Error calculating months between periods: {e}", file=sys.stderr)
        return 0


def generate_quarterly_chunks(from_period, to_period):
    """Break a large date range into quarterly chunks
    Returns list of (from_period, to_period) tuples"""
    
    # Map month names to numbers
    month_map = {
        'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
        'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
    }
    
    try:
        # Parse "Jan 2025" format
        from_parts = from_period.split()
        to_parts = to_period.split()
        
        if len(from_parts) != 2 or len(to_parts) != 2:
            return [(from_period, to_period)]  # Return original if parsing fails
        
        from_month = month_map.get(from_parts[0].lower()[:3])
        from_year = int(from_parts[1])
        to_month = month_map.get(to_parts[0].lower()[:3])
        to_year = int(to_parts[1])
        
        if not from_month or not to_month:
            return [(from_period, to_period)]
        
        # Generate quarters (3-month chunks)
        chunks = []
        current_month = from_month
        current_year = from_year
        
        while (current_year < to_year) or (current_year == to_year and current_month <= to_month):
            # Calculate chunk end (current + 2 months, or to_month if sooner)
            chunk_end_month = min(current_month + 2, 12)
            chunk_end_year = current_year
            
            # Don't exceed the target end date
            if chunk_end_year > to_year or (chunk_end_year == to_year and chunk_end_month > to_month):
                chunk_end_month = to_month
                chunk_end_year = to_year
            
            # Convert back to period names
            month_names = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                          'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            chunk_from = f"{month_names[current_month]} {current_year}"
            chunk_to = f"{month_names[chunk_end_month]} {chunk_end_year}"
            
            chunks.append((chunk_from, chunk_to))
            
            # Move to next quarter
            current_month = chunk_end_month + 1
            if current_month > 12:
                current_month = 1
                current_year += 1
            
            # Safety: prevent infinite loops
            if len(chunks) > 20:
                break
        
        return chunks if chunks else [(from_period, to_period)]
        
    except Exception as e:
        print(f"Error generating chunks: {e}", file=sys.stderr)
        return [(from_period, to_period)]  # Return original on error


def load_lookup_cache():
    """Load all name-to-ID mappings into memory cache"""
    global cache_loaded
    
    if cache_loaded:
        return
    
    print("Loading name-to-ID lookup cache...")
    
    # Load Classes - use known names + query NetSuite
    class_known = {
        'hardware': '1',
        'furniture': '2',
        'racks': '7',
        'accessories': '10',
        'consumer goods': '13',
        'interior': '20',
        'electronics': '22',
        'electrical': '29'
    }
    lookup_cache['classes'] = class_known.copy()
    
    # Try to load more from NetSuite
    try:
        class_query = """
            SELECT DISTINCT c.id, c.name
            FROM Classification c
            WHERE c.id IN (
                SELECT DISTINCT tl.class
                FROM TransactionLine tl
                WHERE tl.class IS NOT NULL AND tl.class != 0
                AND ROWNUM <= 100
            )
        """
        class_result = query_netsuite(class_query)
        if isinstance(class_result, list):
            for row in class_result:
                class_name_lower = row['name'].lower()
                class_id = str(row['id'])
                # Add or update
                lookup_cache['classes'][class_name_lower] = class_id
            print(f"✓ Loaded {len(lookup_cache['classes'])} classes")
    except Exception as e:
        print(f"✗ Class lookup error: {e}, using {len(lookup_cache['classes'])} known classes")
    
    # Load Locations (has real names from NetSuite)
    try:
        loc_query = """
            SELECT DISTINCT l.id, l.name
            FROM Location l
            WHERE l.id IN (
                SELECT DISTINCT tl.location
                FROM TransactionLine tl
                WHERE tl.location IS NOT NULL AND tl.location != 0
                AND ROWNUM <= 100
            )
        """
        loc_result = query_netsuite(loc_query)
        if isinstance(loc_result, list):
            for row in loc_result:
                lookup_cache['locations'][row['name'].lower()] = str(row['id'])
            print(f"✓ Loaded {len(lookup_cache['locations'])} locations")
    except Exception as e:
        print(f"✗ Location lookup error: {e}")
    
    # Departments - use known names + IDs
    dept_known = {
        'demo': '13',
        'corporate': '1',
        'sales': '2',
        'operations': '7',
        'marketing': '11'
    }
    lookup_cache['departments'] = dept_known.copy()
    
    # Try to load more from NetSuite
    try:
        dept_query = """
            SELECT DISTINCT tl.department as id
            FROM TransactionLine tl
            WHERE tl.department IS NOT NULL AND tl.department != 0
            AND ROWNUM <= 100
        """
        dept_result = query_netsuite(dept_query)
        if isinstance(dept_result, list):
            for row in dept_result:
                dept_id = str(row['id'])
                # Add as "Department {id}" if not in known list
                if dept_id not in dept_known.values():
                    lookup_cache['departments'][f'department {dept_id}'.lower()] = dept_id
            print(f"✓ Loaded {len(lookup_cache['departments'])} departments")
    except Exception as e:
        print(f"✗ Department lookup error: {e}")
    
    # Subsidiaries - now we have access to the Subsidiary table!
    try:
        sub_query = """
            SELECT 
                s.id,
                s.name,
                s.fullName AS hierarchy
            FROM 
                Subsidiary s
            ORDER BY 
                s.fullName
        """
        sub_result = query_netsuite(sub_query)
        if isinstance(sub_result, list):
            for row in sub_result:
                sub_id = str(row['id'])
                # Use hierarchy (full path) if available, otherwise just name
                sub_name = row.get('hierarchy', row['name']).lower()
                lookup_cache['subsidiaries'][sub_name] = sub_id
            print(f"✓ Loaded {len(lookup_cache['subsidiaries'])} subsidiaries with hierarchy")
    except Exception as e:
        print(f"✗ Subsidiary lookup error: {e}")
        # Fallback to known values
        lookup_cache['subsidiaries'] = {'parent company': '1'}
    
    # Find top-level parent subsidiary (where parent IS NULL)
    # This is used as default when no subsidiary is specified
    load_default_subsidiary()
    
    cache_loaded = True
    print("✓ Lookup cache loaded!")


def load_default_subsidiary():
    """
    Find the top-level parent subsidiary (where parent IS NULL)
    This subsidiary will be used as the default when no subsidiary is specified.
    For consolidated reporting, this gives the full company view.
    """
    global default_subsidiary_id
    
    try:
        # Query for top-level parent (where parent IS NULL and active)
        # Note: SuiteQL doesn't support LIMIT, use ROWNUM instead
        parent_query = """
            SELECT id, name
            FROM Subsidiary
            WHERE parent IS NULL
              AND isinactive = 'F'
              AND ROWNUM = 1
            ORDER BY id
        """
        result = query_netsuite(parent_query)
        
        if isinstance(result, list) and len(result) > 0:
            default_subsidiary_id = str(result[0]['id'])
            parent_name = result[0]['name']
            print(f"✓ Default subsidiary: {parent_name} (ID: {default_subsidiary_id})")
        else:
            # Fallback: use '1' if query fails
            default_subsidiary_id = '1'
            print(f"⚠ Could not determine parent subsidiary, defaulting to ID=1")
            
    except Exception as e:
        # Fallback: use '1' if query fails
        default_subsidiary_id = '1'
        print(f"⚠ Error finding parent subsidiary: {e}, defaulting to ID=1")


def convert_name_to_id(dimension_type, value):
    """
    Convert a dimension name to its ID
    Args:
        dimension_type: 'subsidiary', 'department', 'class', 'location'
        value: Name (string) or ID (string/number)
    Returns:
        ID as string, or EMPTY STRING if name not found (to prevent SQL errors)
    """
    if not value or value == '':
        return ''
    
    # Load cache if not loaded
    if not cache_loaded:
        load_lookup_cache()
    
    # If it's already a number (ID), return it
    if str(value).isdigit():
        return str(value)
    
    # Look up name in cache (case-insensitive)
    value_lower = str(value).lower().strip()
    
    # Map dimension type to cache key (handle 'class' → 'classes')
    cache_key_map = {
        'subsidiary': 'subsidiaries',
        'department': 'departments',
        'class': 'classes',  # NOT 'classs'!
        'location': 'locations'
    }
    cache_key = cache_key_map.get(dimension_type, dimension_type + 's')
    
    if cache_key in lookup_cache:
        if value_lower in lookup_cache[cache_key]:
            found_id = lookup_cache[cache_key][value_lower]
            print(f"✓ Converted {dimension_type} '{value}' → ID {found_id}")
            return found_id
    
    # Not found - return EMPTY to prevent SQL errors
    # (better to ignore the filter than break the query)
    print(f"⚠ {dimension_type} '{value}' not found in cache, ignoring filter")
    return ''


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


@app.route('/accounts/search', methods=['GET'])
def search_accounts():
    """
    Search for accounts by partial account number
    
    Query params:
        - pattern: Search pattern (e.g., "4*", "42*", "*")
        - active_only: Filter to active accounts only (default: true)
    
    Returns: List of matching accounts with number, name, ID, and type
    
    Example:
        GET /accounts/search?pattern=4*
        Returns all accounts starting with "4"
    """
    try:
        pattern = request.args.get('pattern', '')
        active_only = request.args.get('active_only', 'true').lower() == 'true'
        
        if not pattern:
            return jsonify({'error': 'Pattern parameter is required'}), 400
        
        # Convert Excel wildcard (*) to SQL wildcard (%)
        sql_pattern = pattern.replace('*', '%')
        
        # Escape any existing special characters
        sql_pattern = escape_sql(sql_pattern)
        
        # Build WHERE clause
        where_conditions = []
        
        # Filter by pattern
        where_conditions.append(f"acctnumber LIKE '{sql_pattern}'")
        
        # Filter by active status
        if active_only:
            where_conditions.append("isinactive = 'F'")
        
        where_clause = " AND ".join(where_conditions)
        
        # Build SuiteQL query
        query = f"""
            SELECT 
                id,
                acctnumber,
                accountsearchdisplayname AS accountname,
                accttype
            FROM 
                Account
            WHERE 
                {where_clause}
            ORDER BY 
                acctnumber
        """
        
        print(f"DEBUG - Account search query: {query}", file=sys.stderr)
        
        result = query_netsuite(query)
        
        if isinstance(result, dict) and 'error' in result:
            return jsonify(result), 500
        
        # Format response
        accounts = []
        for row in result:
            accounts.append({
                'id': row.get('id'),
                'accountnumber': row.get('acctnumber'),
                'accountname': row.get('accountname'),
                'accttype': row.get('accttype')
            })
        
        return jsonify({
            'pattern': pattern,
            'count': len(accounts),
            'accounts': accounts
        })
        
    except Exception as e:
        print(f"Error in search_accounts: {str(e)}", file=sys.stderr)
        return jsonify({'error': str(e)}), 500


def build_pl_query(accounts, periods, base_where, target_sub, needs_line_join):
    """
    Build query for P&L accounts (Income Statement)
    P&L accounts show activity within the specific period only
    """
    accounts_in = ','.join([f"'{escape_sql(acc)}'" for acc in accounts])
    periods_in = ','.join([f"'{escape_sql(p)}'" for p in periods])
    
    # Add account and period filters
    where_clause = f"{base_where} AND a.acctnumber IN ({accounts_in}) AND apf.periodname IN ({periods_in})"
    
    # Only include P&L account types
    where_clause += " AND a.accttype IN ('Income', 'OthIncome', 'COGS', 'Expense', 'OthExpense')"
    
    amount_calc = f"""CASE
                        WHEN subs_count > 1 THEN
                            TO_NUMBER(
                                BUILTIN.CONSOLIDATE(
                                    tal.amount,
                                    'LEDGER',
                                    'DEFAULT',
                                    'DEFAULT',
                                    {target_sub},
                                    t.postingperiod,
                                    'DEFAULT'
                                )
                            )
                        ELSE tal.amount
                    END""" if target_sub else "tal.amount"
    
    if needs_line_join:
        return f"""
            SELECT 
                a.acctnumber,
                ap.periodname,
                SUM(cons_amt) AS balance
            FROM (
                SELECT
                    tal.account,
                    t.postingperiod,
                    {amount_calc}
                    * CASE WHEN a.accttype IN ('Income','OthIncome') THEN -1 ELSE 1 END AS cons_amt
                FROM TransactionAccountingLine tal
                    JOIN Transaction t ON t.id = tal.transaction
                    JOIN TransactionLine tl ON t.id = tl.transaction AND tal.transactionline = tl.id
                    JOIN Account a ON a.id = tal.account
                    JOIN AccountingPeriod apf ON apf.id = t.postingperiod
                    CROSS JOIN (
                        SELECT COUNT(*) AS subs_count
                        FROM Subsidiary
                        WHERE isinactive = 'F'
                    ) subs_cte
                WHERE {where_clause}
                    AND COALESCE(a.eliminate, 'F') = 'F'
            ) x
            JOIN Account a ON a.id = x.account
            JOIN AccountingPeriod ap ON ap.id = x.postingperiod
            GROUP BY a.acctnumber, ap.periodname
            ORDER BY a.acctnumber, ap.periodname
        """
    else:
        return f"""
            SELECT 
                a.acctnumber,
                ap.periodname,
                SUM(cons_amt) AS balance
            FROM (
                SELECT
                    tal.account,
                    t.postingperiod,
                    {amount_calc}
                    * CASE WHEN a.accttype IN ('Income','OthIncome') THEN -1 ELSE 1 END AS cons_amt
                FROM TransactionAccountingLine tal
                    JOIN Transaction t ON t.id = tal.transaction
                    JOIN Account a ON a.id = tal.account
                    JOIN AccountingPeriod apf ON apf.id = t.postingperiod
                    CROSS JOIN (
                        SELECT COUNT(*) AS subs_count
                        FROM Subsidiary
                        WHERE isinactive = 'F'
                    ) subs_cte
                WHERE {where_clause}
                    AND COALESCE(a.eliminate, 'F') = 'F'
            ) x
            JOIN Account a ON a.id = x.account
            JOIN AccountingPeriod ap ON ap.id = x.postingperiod
            GROUP BY a.acctnumber, ap.periodname
            ORDER BY a.acctnumber, ap.periodname
        """


def build_bs_query_single_period(accounts, period_name, period_info, base_where, target_sub, needs_line_join):
    """
    Build query for Balance Sheet accounts for a SINGLE period
    Balance Sheet = CUMULATIVE balance from inception through period end
    
    Returns one row per account with the cumulative balance as of period end
    """
    from datetime import datetime
    
    accounts_in = ','.join([f"'{escape_sql(acc)}'" for acc in accounts])
    
    enddate = period_info['enddate']
    period_id = period_info['id']
    
    # Parse enddate
    try:
        end_date_obj = datetime.strptime(enddate, '%m/%d/%Y')
        end_date_str = end_date_obj.strftime('%Y-%m-%d')
    except:
        end_date_str = enddate
    
    # Build WHERE clause
    where_clause = f"{base_where} AND a.acctnumber IN ({accounts_in})"
    where_clause += " AND a.accttype NOT IN ('Income', 'OthIncome', 'COGS', 'Expense', 'OthExpense')"
    # CUMULATIVE: All transactions through period end (no lower bound)
    where_clause += f" AND t.trandate <= TO_DATE('{end_date_str}', 'YYYY-MM-DD')"
    where_clause += " AND tal.accountingbook = 1"
    
    amount_calc = f"""CASE
                        WHEN subs_count > 1 THEN
                            TO_NUMBER(
                                BUILTIN.CONSOLIDATE(
                                    tal.amount,
                                    'LEDGER',
                                    'DEFAULT',
                                    'DEFAULT',
                                    {target_sub},
                                    {period_id},
                                    'DEFAULT'
                                )
                            )
                        ELSE tal.amount
                    END""" if target_sub else "tal.amount"
    
    if needs_line_join:
        return f"""
            SELECT 
                a.acctnumber,
                SUM({amount_calc}) AS balance
            FROM TransactionAccountingLine tal
                JOIN Transaction t ON t.id = tal.transaction
                JOIN TransactionLine tl ON t.id = tl.transaction AND tal.transactionline = tl.id
                JOIN Account a ON a.id = tal.account
                CROSS JOIN (
                    SELECT COUNT(*) AS subs_count
                    FROM Subsidiary
                    WHERE isinactive = 'F'
                ) subs_cte
            WHERE {where_clause}
                AND COALESCE(a.eliminate, 'F') = 'F'
            GROUP BY a.acctnumber
        """
    else:
        return f"""
            SELECT 
                a.acctnumber,
                SUM({amount_calc}) AS balance
            FROM TransactionAccountingLine tal
                JOIN Transaction t ON t.id = tal.transaction
                JOIN Account a ON a.id = tal.account
                CROSS JOIN (
                    SELECT COUNT(*) AS subs_count
                    FROM Subsidiary
                    WHERE isinactive = 'F'
                ) subs_cte
            WHERE {where_clause}
                AND COALESCE(a.eliminate, 'F') = 'F'
            GROUP BY a.acctnumber
        """


def build_bs_query(accounts, period_info, base_where, target_sub, needs_line_join):
    """
    Build query for Balance Sheet accounts (Assets/Liabilities/Equity)
    Balance Sheet accounts show CUMULATIVE balance from inception through period end
    
    Key difference: For each period, use t.trandate <= period.enddate
    Returns row-based output (like P&L) - one row per account per period
    
    Performance optimization: 
    1. Query ONE period at a time (UNION ALL)
    2. Limit to fiscal year scope (not ALL history) to avoid timeouts
    """
    from datetime import datetime
    from dateutil.relativedelta import relativedelta
    
    accounts_in = ','.join([f"'{escape_sql(acc)}'" for acc in accounts])
    
    # Find the earliest period to determine fiscal year start
    earliest_enddate = min([info['enddate'] for info in period_info.values()])
    try:
        earliest_date = datetime.strptime(earliest_enddate, '%m/%d/%Y')
        # Get fiscal year start (January 1 of that year)
        fiscal_year_start = datetime(earliest_date.year, 1, 1)
        min_date_str = fiscal_year_start.strftime('%Y-%m-%d')
        print(f"DEBUG - BS query using fiscal year start: {min_date_str}", file=sys.stderr)
    except Exception as e:
        # Fallback: 1 year back
        min_date_str = '2024-01-01'
        print(f"DEBUG - BS query using fallback date: {min_date_str} (error: {e})", file=sys.stderr)
    
    union_queries = []
    
    for period, info in period_info.items():
        enddate = info['enddate']
        period_id = info['id']
        
        # Parse enddate
        try:
            end_date_obj = datetime.strptime(enddate, '%m/%d/%Y')
            end_date_str = end_date_obj.strftime('%Y-%m-%d')
        except:
            end_date_str = enddate
        
        # Build WHERE clause for this period
        period_where = f"{base_where} AND a.acctnumber IN ({accounts_in})"
        period_where += " AND a.accttype NOT IN ('Income', 'OthIncome', 'COGS', 'Expense', 'OthExpense')"
        # CRITICAL: Balance Sheet is CUMULATIVE - ALL transactions through period end (like user's reference)
        # No lower bound to get true cumulative balance
        period_where += f" AND t.trandate <= TO_DATE('{end_date_str}', 'YYYY-MM-DD')"
        # Add accountingbook filter (like user's reference query)
        period_where += " AND tal.accountingbook = 1"
        
        amount_calc = f"""CASE
                            WHEN subs_count > 1 THEN
                                TO_NUMBER(
                                    BUILTIN.CONSOLIDATE(
                                        tal.amount,
                                        'LEDGER',
                                        'DEFAULT',
                                        'DEFAULT',
                                        {target_sub},
                                        {period_id},
                                        'DEFAULT'
                                    )
                                )
                            ELSE tal.amount
                        END""" if target_sub else "tal.amount"
        
        # Query for THIS period only
        if needs_line_join:
            period_query = f"""
                SELECT 
                    a.acctnumber,
                    '{escape_sql(period)}' AS periodname,
                    SUM({amount_calc}) AS balance
                FROM TransactionAccountingLine tal
                    JOIN Transaction t ON t.id = tal.transaction
                    JOIN TransactionLine tl ON t.id = tl.transaction AND tal.transactionline = tl.id
                    JOIN Account a ON a.id = tal.account
                    CROSS JOIN (
                        SELECT COUNT(*) AS subs_count
                        FROM Subsidiary
                        WHERE isinactive = 'F'
                    ) subs_cte
                WHERE {period_where}
                    AND COALESCE(a.eliminate, 'F') = 'F'
                GROUP BY a.acctnumber
            """
        else:
            period_query = f"""
                SELECT 
                    a.acctnumber,
                    '{escape_sql(period)}' AS periodname,
                    SUM({amount_calc}) AS balance
                FROM TransactionAccountingLine tal
                    JOIN Transaction t ON t.id = tal.transaction
                    JOIN Account a ON a.id = tal.account
                    CROSS JOIN (
                        SELECT COUNT(*) AS subs_count
                        FROM Subsidiary
                        WHERE isinactive = 'F'
                    ) subs_cte
                WHERE {period_where}
                    AND COALESCE(a.eliminate, 'F') = 'F'
                GROUP BY a.acctnumber
            """
        
        union_queries.append(period_query)
    
    # UNION all period queries
    full_query = " UNION ALL ".join(union_queries)
    full_query += " ORDER BY acctnumber, periodname"
    
    return full_query


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
    
    # Convert names to IDs (accepts names OR IDs)
    subsidiary = convert_name_to_id('subsidiary', subsidiary)
    class_id = convert_name_to_id('class', class_id)
    department = convert_name_to_id('department', department)
    location = convert_name_to_id('location', location)
    
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
        
        # IMPORTANT: Do NOT filter by t.subsidiary here!
        # For consolidation, we let BUILTIN.CONSOLIDATE handle subsidiary filtering
        # The target_sub parameter tells CONSOLIDATE which subsidiary hierarchy to use
        
        # Need TransactionLine join if filtering by department, class, or location
        needs_line_join = (department and department != '') or (class_id and class_id != '') or (location and location != '')
        
        if class_id and class_id != '':
            where_clauses.append(f"tl.class = {class_id}")
        
        if department and department != '':
            where_clauses.append(f"tl.department = {department}")
        
        if location and location != '':
            where_clauses.append(f"tl.location = {location}")
        
        # Get period enddates for Balance Sheet calculation
        # Balance Sheet accounts need cumulative balance (inception through period end)
        period_info = {}
        for period in periods:
            start, end, period_id = get_period_dates_from_name(period)
            if end and period_id:
                period_info[period] = {'enddate': end, 'id': period_id}
        
        # Determine target subsidiary for consolidation
        # If subsidiary filter is applied, consolidate to that subsidiary (for Consolidated view)
        # If no subsidiary, default to top-level parent (dynamically determined at startup)
        if subsidiary and subsidiary != '':
            target_sub = subsidiary
        else:
            target_sub = default_subsidiary_id or '1'
        
        # Build base WHERE clause (without period filter yet)
        base_where = " AND ".join(where_clauses)
        
        # ============================================================================
        # STRATEGY: Run SEPARATE queries for P&L vs Balance Sheet accounts
        # Then merge the results - this is cleaner and more maintainable
        # ============================================================================
        
        all_balances = {}
        
        # QUERY 1: P&L Accounts (Income Statement)
        # Use current period-specific logic (ap.periodname IN periods)
        pl_query = build_pl_query(accounts, periods, base_where, target_sub, needs_line_join)
        
        print(f"DEBUG - P&L Query:\n{pl_query[:500]}...", file=sys.stderr)
        pl_result = query_netsuite(pl_query)
        
        if isinstance(pl_result, list):
            print(f"DEBUG - P&L returned {len(pl_result)} rows", file=sys.stderr)
            for row in pl_result:
                account_num = row['acctnumber']
                period_name = row['periodname']
                balance = float(row['balance']) if row['balance'] else 0
                
                if account_num not in all_balances:
                    all_balances[account_num] = {}
                all_balances[account_num][period_name] = balance
        
        # QUERY 2: Balance Sheet Accounts (Assets/Liabilities/Equity)
        # Use cumulative logic (t.trandate <= period.enddate)
        # Query each period SEPARATELY (UNION ALL causes 400 errors with complex queries)
        if period_info:
            print(f"DEBUG - Querying {len(period_info)} periods for Balance Sheet accounts...", file=sys.stderr)
            
            for period, info in period_info.items():
                try:
                    # Build query for THIS period only
                    period_query = build_bs_query_single_period(
                        accounts, period, info, base_where, target_sub, needs_line_join
                    )
                    
                    print(f"DEBUG - BS Query for {period}:\n{period_query[:300]}...", file=sys.stderr)
                    
                    # Balance Sheet queries can be slower - use 90 second timeout
                    bs_result = query_netsuite(period_query, timeout=90)
                    
                    if isinstance(bs_result, list):
                        print(f"DEBUG - BS returned {len(bs_result)} rows for {period}", file=sys.stderr)
                        # Process results for this period
                        for row in bs_result:
                            account_num = row['acctnumber']
                            balance = float(row['balance']) if row['balance'] else 0
                            
                            if account_num not in all_balances:
                                all_balances[account_num] = {}
                            all_balances[account_num][period] = balance
                    elif isinstance(bs_result, dict) and 'error' in bs_result:
                        print(f"ERROR - BS query failed for {period}: {bs_result['error']}", file=sys.stderr)
                    else:
                        print(f"ERROR - BS query unexpected result type for {period}: {type(bs_result)}", file=sys.stderr)
                except Exception as e:
                    print(f"ERROR - BS query exception for {period}: {str(e)}", file=sys.stderr)
        
        print(f"DEBUG - Final merged balances: {list(all_balances.keys())}", file=sys.stderr)
        
        # Fill in zeros for missing account/period combinations
        for account_num in accounts:
            if account_num not in all_balances:
                all_balances[account_num] = {}
            for period in periods:
                if period not in all_balances[account_num]:
                    all_balances[account_num][period] = 0
        
        # Return merged results
        return jsonify({'balances': all_balances})
        
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
        
        # Convert names to IDs (accepts names OR IDs)
        subsidiary = convert_name_to_id('subsidiary', subsidiary)
        class_id = convert_name_to_id('class', class_id)
        department = convert_name_to_id('department', department)
        location = convert_name_to_id('location', location)
        
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
                # Convert period names to DATE ranges
                # Period IDs don't work because they include quarterly/fiscal periods
                from_start, from_end = get_period_dates_from_name(from_period)
                to_start, to_end = get_period_dates_from_name(to_period)
                if from_start and to_end:
                    # Use date strings directly (NetSuite returns dates as strings)
                    where_clauses.append(f"ap.startdate >= '{from_start}'")
                    where_clauses.append(f"ap.enddate <= '{to_end}'")
                else:
                    # Fallback to period name if conversion fails
                    where_clauses.append(f"ap.periodname = '{escape_sql(from_period)}'")
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
        
        # Determine target subsidiary for consolidation
        target_sub = subsidiary if subsidiary and subsidiary != '' else 'NULL'
        
        # Need TransactionLine join if filtering by department, class, or location
        needs_line_join = (department and department != '') or (class_id and class_id != '') or (location and location != '')
        
        if (from_period and not from_period.isdigit()) or (to_period and not to_period.isdigit()):
            if needs_line_join:
                query = f"""
                    SELECT SUM(x.cons_amt) AS balance
                    FROM (
                        SELECT
                            CASE
                                WHEN subs_count > 1 THEN
                                    TO_NUMBER(
                                        BUILTIN.CONSOLIDATE(
                                            tal.amount,
                                            'LEDGER',
                                            'DEFAULT',
                                            'DEFAULT',
                                            {target_sub},
                                            t.postingperiod,
                                            'DEFAULT'
                                        )
                                    )
                                ELSE tal.amount
                            END
                            * CASE WHEN a.accttype IN ('Income','OthIncome') THEN -1 ELSE 1 END AS cons_amt
                        FROM TransactionAccountingLine tal
                            JOIN Transaction t ON t.id = tal.transaction
                            JOIN TransactionLine tl ON t.id = tl.transaction AND tal.transactionline = tl.id
                            JOIN Account a ON a.id = tal.account
                            JOIN AccountingPeriod ap ON ap.id = t.postingperiod
                            CROSS JOIN (
                                SELECT COUNT(*) AS subs_count
                                FROM Subsidiary
                                WHERE isinactive = 'F'
                            ) subs_cte
                        WHERE {where_clause}
                            AND COALESCE(a.eliminate, 'F') = 'F'
                    ) x
                """
            else:
                query = f"""
                    SELECT SUM(x.cons_amt) AS balance
                    FROM (
                        SELECT
                            CASE
                                WHEN subs_count > 1 THEN
                                    TO_NUMBER(
                                        BUILTIN.CONSOLIDATE(
                                            tal.amount,
                                            'LEDGER',
                                            'DEFAULT',
                                            'DEFAULT',
                                            {target_sub},
                                            t.postingperiod,
                                            'DEFAULT'
                                        )
                                    )
                                ELSE tal.amount
                            END
                            * CASE WHEN a.accttype IN ('Income','OthIncome') THEN -1 ELSE 1 END AS cons_amt
                        FROM TransactionAccountingLine tal
                            JOIN Transaction t ON t.id = tal.transaction
                            JOIN Account a ON a.id = tal.account
                            JOIN AccountingPeriod ap ON ap.id = t.postingperiod
                            CROSS JOIN (
                                SELECT COUNT(*) AS subs_count
                                FROM Subsidiary
                                WHERE isinactive = 'F'
                            ) subs_cte
                        WHERE {where_clause}
                            AND COALESCE(a.eliminate, 'F') = 'F'
                    ) x
                """
        else:
            if needs_line_join:
                query = f"""
                    SELECT SUM(x.cons_amt) AS balance
                    FROM (
                        SELECT
                            CASE
                                WHEN subs_count > 1 THEN
                                    TO_NUMBER(
                                        BUILTIN.CONSOLIDATE(
                                            tal.amount,
                                            'LEDGER',
                                            'DEFAULT',
                                            'DEFAULT',
                                            {target_sub},
                                            t.postingperiod,
                                            'DEFAULT'
                                        )
                                    )
                                ELSE tal.amount
                            END
                            * CASE WHEN a.accttype IN ('Income','OthIncome') THEN -1 ELSE 1 END AS cons_amt
                        FROM TransactionAccountingLine tal
                            JOIN Transaction t ON t.id = tal.transaction
                            JOIN TransactionLine tl ON t.id = tl.transaction AND tal.transactionline = tl.id
                            JOIN Account a ON a.id = tal.account
                            CROSS JOIN (
                                SELECT COUNT(*) AS subs_count
                                FROM Subsidiary
                                WHERE isinactive = 'F'
                            ) subs_cte
                        WHERE {where_clause}
                            AND COALESCE(a.eliminate, 'F') = 'F'
                    ) x
                """
            else:
                query = f"""
                    SELECT SUM(x.cons_amt) AS balance
                    FROM (
                        SELECT
                            CASE
                                WHEN subs_count > 1 THEN
                                    TO_NUMBER(
                                        BUILTIN.CONSOLIDATE(
                                            tal.amount,
                                            'LEDGER',
                                            'DEFAULT',
                                            'DEFAULT',
                                            {target_sub},
                                            t.postingperiod,
                                            'DEFAULT'
                                        )
                                    )
                                ELSE tal.amount
                            END
                            * CASE WHEN a.accttype IN ('Income','OthIncome') THEN -1 ELSE 1 END AS cons_amt
                        FROM TransactionAccountingLine tal
                            JOIN Transaction t ON t.id = tal.transaction
                            JOIN Account a ON a.id = tal.account
                            CROSS JOIN (
                                SELECT COUNT(*) AS subs_count
                                FROM Subsidiary
                                WHERE isinactive = 'F'
                            ) subs_cte
                        WHERE {where_clause}
                            AND COALESCE(a.eliminate, 'F') = 'F'
                    ) x
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
        
        # Convert names to IDs (accepts names OR IDs)
        subsidiary = convert_name_to_id('subsidiary', subsidiary)
        class_id = convert_name_to_id('class', class_id)
        department = convert_name_to_id('department', department)
        location = convert_name_to_id('location', location)
        
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
                # Convert period names to DATE ranges (same fix as balance query)
                from_start, from_end = get_period_dates_from_name(from_period)
                to_start, to_end = get_period_dates_from_name(to_period)
                if from_start and to_end:
                    # Use date strings directly (NetSuite returns dates as strings)
                    where_clauses.append(f"ap.startdate >= '{from_start}'")
                    where_clauses.append(f"ap.enddate <= '{to_end}'")
                else:
                    # Fallback to period name if conversion fails
                    where_clauses.append(f"ap.periodname = '{escape_sql(from_period)}'")
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
        
        # Determine target subsidiary for consolidation
        target_sub = subsidiary if subsidiary and subsidiary != '' else 'NULL'
        
        # Build SuiteQL query - only join AccountingPeriod if using period names
        # Note: Budget amounts also need BUILTIN.CONSOLIDATE for multi-currency
        # Budgets are typically 'LEDGER' type (not INCOME)
        if (from_period and not from_period.isdigit()) or (to_period and not to_period.isdigit()):
            query = f"""
                SELECT SUM(
                    BUILTIN.CONSOLIDATE(
                        b.amount, 'LEDGER', 'DEFAULT', 'DEFAULT',
                        {target_sub}, b.accountingperiod, 'DEFAULT'
                    )
                ) AS budget_amount
                FROM Budget b
                INNER JOIN Account a ON b.account = a.id
                INNER JOIN AccountingPeriod ap ON b.accountingperiod = ap.id
                WHERE {where_clause}
            """
        else:
            query = f"""
                SELECT SUM(
                    BUILTIN.CONSOLIDATE(
                        b.amount, 'LEDGER', 'DEFAULT', 'DEFAULT',
                        {target_sub}, b.accountingperiod, 'DEFAULT'
                    )
                ) AS budget_amount
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
        
        # Convert names to IDs (accepts names OR IDs)
        subsidiary = convert_name_to_id('subsidiary', subsidiary)
        class_id = convert_name_to_id('class', class_id)
        department = convert_name_to_id('department', department)
        location = convert_name_to_id('location', location)
        
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
        # For drill-down, we show RAW transaction amounts (no consolidation)
        # This gives users the actual transaction detail, not consolidated view
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
                    {debit_expr} AS debit,
                    {credit_expr} AS credit,
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
    Returns data from the in-memory cache (already loaded at startup)
    
    For subsidiaries that are parents (have children), we also add a "(Consolidated)" option
    which uses BUILTIN.CONSOLIDATE to include parent + all children transactions
    """
    try:
        # Load cache if not already loaded
        if not cache_loaded:
            load_lookup_cache()
        
        # Convert cache format (name→id) to list format (id, name) for frontend
        lookups = {
            'subsidiaries': [],
            'departments': [],
            'classes': [],
            'locations': []
        }
        
        # Get subsidiary hierarchy to identify parents
        try:
            hierarchy_query = """
                SELECT id, name, parent
                FROM Subsidiary
                WHERE isinactive = 'F'
                ORDER BY name
            """
            hierarchy_result = query_netsuite(hierarchy_query)
            
            # Identify parent subsidiaries (those with children)
            parent_ids = set()
            all_subs = {}
            
            if isinstance(hierarchy_result, list):
                for row in hierarchy_result:
                    sub_id = str(row['id'])
                    all_subs[sub_id] = row['name']
                    if row.get('parent'):
                        parent_ids.add(str(row['parent']))
                
                # Add all subsidiaries
                for sub_id, sub_name in all_subs.items():
                    lookups['subsidiaries'].append({
                        'id': sub_id,
                        'name': sub_name
                    })
                    
                    # If this is a parent, also add "(Consolidated)" version
                    if sub_id in parent_ids:
                        lookups['subsidiaries'].append({
                            'id': sub_id,  # Same ID, BUILTIN.CONSOLIDATE handles consolidation
                            'name': f"{sub_name} (Consolidated)"
                        })
        except Exception as e:
            print(f"Error loading subsidiary hierarchy: {e}", file=sys.stderr)
            # Fallback to cache
            for name, id_val in lookup_cache['subsidiaries'].items():
                lookups['subsidiaries'].append({
                    'id': id_val,
                    'name': name.title()
                })
        
        # Convert cache data for other lookups
        for name, id_val in lookup_cache['departments'].items():
            lookups['departments'].append({
                'id': id_val,
                'name': name.title()  # Capitalize first letter
            })
        
        for name, id_val in lookup_cache['classes'].items():
            lookups['classes'].append({
                'id': id_val,
                'name': name.title()  # Capitalize first letter
            })
        
        for name, id_val in lookup_cache['locations'].items():
            lookups['locations'].append({
                'id': id_val,
                'name': name  # Keep location names as-is
            })
        
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
    print("Loading name-to-ID lookup cache...")
    load_lookup_cache()
    print()
    print("Press Ctrl+C to stop")
    print("=" * 80)
    print()
    
    # Run server
    app.run(host='127.0.0.1', port=5002, debug=False)

