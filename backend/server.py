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

# In-memory cache for balance data (from full year refresh)
# Structure: { 'account:period:filters_hash': balance_value }
# Expires after 5 minutes
balance_cache = {}
balance_cache_timestamp = None
BALANCE_CACHE_TTL = 300  # 5 minutes in seconds

# In-memory cache for BS ACTIVITY data (used to compute cumulative balances)
# Structure: { 'account:period:filters_hash': activity_value }
# Backend computes cumulative by summing activity from Jan through requested period
bs_activity_cache = {}
bs_activity_cache_timestamp = None

# Track which accounts are Balance Sheet (for cumulative calculation)
# Structure: { 'account_number': True }
bs_account_set = set()

# In-memory cache for account titles (permanent, rarely changes)
# Structure: { 'account_number': 'account_name' }
account_title_cache = {}

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


def calculate_period_end_date(period_name):
    """Calculate the end date of a period from its name (e.g., 'Jan 2025' -> '01/31/2025')
    Used as a fallback when the period doesn't exist in NetSuite's AccountingPeriod table
    """
    import calendar
    
    month_map = {
        'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
        'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
    }
    
    try:
        parts = period_name.strip().split()
        if len(parts) != 2:
            return None
        
        month_str = parts[0].lower()[:3]
        year = int(parts[1])
        month = month_map.get(month_str)
        
        if not month:
            return None
        
        # Get last day of the month
        last_day = calendar.monthrange(year, month)[1]
        
        # Return in MM/DD/YYYY format (same as NetSuite returns)
        return f"{month:02d}/{last_day:02d}/{year}"
    except Exception as e:
        print(f"Error calculating period end date for '{period_name}': {e}", file=sys.stderr)
        return None


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
            print(f"DEBUG: Found period '{period_name}' -> {dates}", file=sys.stderr)
            return dates
        print(f"DEBUG: Period '{period_name}' NOT found in NetSuite AccountingPeriod table", file=sys.stderr)
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


@app.route('/admin/restart', methods=['POST'])
def admin_restart():
    """
    Restart the server (called from add-in settings)
    Uses os.execv to replace the current process with a fresh one
    """
    print("=" * 60, file=sys.stderr)
    print("🔄 SERVER RESTART REQUESTED FROM ADD-IN", file=sys.stderr)
    print("=" * 60, file=sys.stderr)
    
    def do_restart():
        import time
        time.sleep(1)  # Give time for response to be sent
        os.execv(sys.executable, ['python3', '-u'] + sys.argv)
    
    import threading
    threading.Thread(target=do_restart).start()
    
    return jsonify({
        'status': 'restarting',
        'message': 'Server will restart in 1 second. Refresh the taskpane in a few seconds.'
    })


@app.route('/accounts/search', methods=['GET'])
def search_accounts():
    """
    Search for accounts by account number OR account type
    
    Query params:
        - pattern: Search pattern
          Examples:
            - "4*"        → Accounts starting with "4"
            - "*"         → All accounts
            - "*income"   → All accounts with type containing "income" (Income, Other Income)
            - "income*"   → All accounts with type starting with "income"
            - "expense"   → All accounts with type containing "expense"
            - "bank"      → All accounts with type containing "bank"
        - active_only: Filter to active accounts only (default: true)
    
    Returns: List of matching accounts with number, name, ID, and type
    """
    try:
        pattern = request.args.get('pattern', '')
        active_only = request.args.get('active_only', 'true').lower() == 'true'
        
        if not pattern:
            return jsonify({'error': 'Pattern parameter is required'}), 400
        
        # Determine if this is a TYPE search or ACCOUNT NUMBER search
        # Type search: contains letters (other than wildcards)
        # Account number search: only numbers and wildcards
        pattern_without_wildcards = pattern.replace('*', '').strip()
        is_type_search = bool(pattern_without_wildcards) and any(c.isalpha() for c in pattern_without_wildcards)
        
        # Build WHERE clause
        where_conditions = []
        
        if is_type_search:
            # ACCOUNT TYPE search
            # Convert pattern to SQL LIKE pattern
            sql_pattern = pattern.replace('*', '%').upper()
            sql_pattern = escape_sql(sql_pattern)
            
            # NetSuite account type mapping for better matching
            # Map common user inputs to actual NetSuite type values
            type_mappings = {
                'INCOME': ['Income', 'OthIncome'],
                'EXPENSE': ['Expense', 'OthExpense'],
                'COGS': ['COGS', 'Cost of Goods Sold'],
                'ASSET': ['Bank', 'AcctRec', 'OthCurrAsset', 'FixedAsset', 'OthAsset', 'DeferExpense', 'Unbilled'],
                'LIABILITY': ['AcctPay', 'CreditCard', 'OthCurrLiab', 'LongTermLiab', 'DeferRevenue'],
                'EQUITY': ['Equity']
            }
            
            # Check if pattern matches a category
            pattern_upper = pattern_without_wildcards.upper()
            matched_types = []
            
            for category, types in type_mappings.items():
                if category.startswith(pattern_upper) or pattern_upper in category:
                    matched_types.extend(types)
            
            if matched_types:
                # Use exact match for mapped types
                type_list = "','".join(matched_types)
                where_conditions.append(f"accttype IN ('{type_list}')")
            else:
                # Use LIKE for direct type matching
                where_conditions.append(f"UPPER(accttype) LIKE '{sql_pattern}'")
            
            print(f"DEBUG - Type search: pattern='{pattern}', sql_pattern='{sql_pattern}', mapped_types={matched_types}", file=sys.stderr)
            
        else:
            # ACCOUNT NUMBER search
            # Convert Excel wildcard (*) to SQL wildcard (%)
            sql_pattern = pattern.replace('*', '%')
            sql_pattern = escape_sql(sql_pattern)
            where_conditions.append(f"acctnumber LIKE '{sql_pattern}'")
            
            print(f"DEBUG - Account number search: pattern='{pattern}', sql_pattern='{sql_pattern}'", file=sys.stderr)
        
        # Filter by active status
        if active_only:
            where_conditions.append("isinactive = 'F'")
        
        where_clause = " AND ".join(where_conditions)
        
        # Build SuiteQL query
        # Use accountsearchdisplaynamecopy for clean name (without number prefix)
        query = f"""
            SELECT 
                id,
                acctnumber,
                accountsearchdisplaynamecopy AS accountname,
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
            'search_type': 'account_type' if is_type_search else 'account_number',
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
    
    # If period_id is None (period doesn't exist in NetSuite), skip consolidation
    # This happens for future periods that haven't been created yet
    if period_id and target_sub:
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
                        END"""
    else:
        # No consolidation - use raw amount
        # This is a fallback for periods not in NetSuite's AccountingPeriod table
        print(f"WARNING: Using non-consolidated amounts for BS query (period_id={period_id})", file=sys.stderr)
        amount_calc = "tal.amount"
    
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


def build_full_year_bs_query(fiscal_year, target_sub, filters):
    """
    OPTIMIZED: Build full-year BALANCE SHEET query returning ACTIVITY per month.
    
    Instead of expensive 12 CROSS JOINs for cumulative balances, this query:
    1. Returns monthly ACTIVITY (like P&L query pattern)
    2. Backend/Excel computes cumulative from activity
    
    This reduces query time from 60-90 seconds to ~15-25 seconds.
    
    Expected performance: ~15-25 seconds (same as P&L)
    """
    # Build optional filter clauses
    filter_clauses = []
    if filters.get('subsidiary'):
        filter_clauses.append(f"t.subsidiary = {filters['subsidiary']}")
    if filters.get('department'):
        filter_clauses.append(f"tal.department = {filters['department']}")
    if filters.get('location'):
        filter_clauses.append(f"tal.location = {filters['location']}")
    if filters.get('class'):
        filter_clauses.append(f"tal.class = {filters['class']}")
    
    filter_sql = (" AND " + " AND ".join(filter_clauses)) if filter_clauses else ""
    
    # Query returns ACTIVITY per month (like P&L) - NOT cumulative
    # Backend will compute cumulative from activity when needed
    query = f"""
    WITH sub_cte AS (
      SELECT COUNT(*) AS subs_count
      FROM Subsidiary
      WHERE isinactive = 'F'
    ),
    base AS (
      SELECT
        tal.account AS account_id,
        t.postingperiod AS period_id,
        CASE
          WHEN (SELECT subs_count FROM sub_cte) > 1 THEN
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
        END AS cons_amt
      FROM TransactionAccountingLine tal
      JOIN Transaction t ON t.id = tal.transaction
      JOIN Account a ON a.id = tal.account
      JOIN AccountingPeriod ap ON ap.id = t.postingperiod
      CROSS JOIN sub_cte
      WHERE t.posting = 'T'
        AND tal.posting = 'T'
        AND tal.accountingbook = 1
        AND ap.isyear = 'F'
        AND ap.isquarter = 'F'
        AND EXTRACT(YEAR FROM ap.startdate) = {fiscal_year}
        AND COALESCE(a.eliminate, 'F') = 'F'
        AND a.accttype NOT IN ('Income','COGS','Cost of Goods Sold','Expense','OthIncome','OthExpense')
        {filter_sql}
    )
    SELECT
      a.acctnumber AS account_number,
      a.accttype AS account_type,
      TO_CHAR(ap.startdate,'YYYY-MM') AS month,
      SUM(b.cons_amt) AS amount
    FROM base b
    JOIN AccountingPeriod ap ON ap.id = b.period_id
    JOIN Account a ON a.id = b.account_id
    GROUP BY a.acctnumber, a.accttype, ap.startdate
    ORDER BY a.acctnumber, ap.startdate
    """
    
    return query


def build_full_year_pl_query(fiscal_year, target_sub, filters):
    """
    Build optimized full-year P&L query using CTE pattern.
    This query consolidates FIRST (in the CTE), then groups - MUCH faster than grouping then consolidating.
    
    IMPORTANT: Query ends with ORDER BY for stable pagination.
    OFFSET/LIMIT will be added by the paginator.
    
    Expected performance: ~15 seconds per 1000 rows
    """
    # Build optional filter clauses
    filter_clauses = []
    if filters.get('subsidiary'):
        filter_clauses.append(f"t.subsidiary = {filters['subsidiary']}")
    if filters.get('department'):
        filter_clauses.append(f"tal.department = {filters['department']}")
    if filters.get('location'):
        filter_clauses.append(f"tal.location = {filters['location']}")
    if filters.get('class'):
        filter_clauses.append(f"tal.class = {filters['class']}")
    
    filter_sql = (" AND " + " AND ".join(filter_clauses)) if filter_clauses else ""
    
    # Base query WITHOUT OFFSET/LIMIT - those are added by paginator
    query = f"""
    WITH sub_cte AS (
      SELECT COUNT(*) AS subs_count
      FROM Subsidiary
      WHERE isinactive = 'F'
    ),
    base AS (
      SELECT
        tal.account AS account_id,
        t.postingperiod AS period_id,
        CASE
          WHEN (SELECT subs_count FROM sub_cte) > 1 THEN
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
        * CASE WHEN a.accttype IN ('Income','OthIncome') THEN -1 ELSE 1 END
        AS cons_amt
      FROM TransactionAccountingLine tal
      JOIN Transaction t ON t.id = tal.transaction
      JOIN Account a ON a.id = tal.account
      JOIN AccountingPeriod ap ON ap.id = t.postingperiod
      CROSS JOIN sub_cte
      WHERE t.posting = 'T'
        AND tal.posting = 'T'
        AND tal.accountingbook = 1
        AND ap.isyear = 'F'
        AND ap.isquarter = 'F'
        AND EXTRACT(YEAR FROM ap.startdate) = {fiscal_year}
        AND COALESCE(a.eliminate, 'F') = 'F'
        AND a.accttype IN ('Income','COGS','Cost of Goods Sold','Expense','OthIncome','OthExpense')
        {filter_sql}
    )
    SELECT
      a.acctnumber AS account_number,
      a.accttype AS account_type,
      TO_CHAR(ap.startdate,'YYYY-MM') AS month,
      SUM(b.cons_amt) AS amount
    FROM base b
    JOIN AccountingPeriod ap ON ap.id = b.period_id
    JOIN Account a ON a.id = b.account_id
    GROUP BY a.acctnumber, a.accttype, ap.startdate
    ORDER BY a.acctnumber, ap.startdate
    """
    
    return query


def run_paginated_suiteql(base_query, page_size=1000, max_pages=20, timeout=120):
    """
    Execute a SuiteQL query with pagination to overcome NetSuite's 1000-row limit.
    
    NetSuite SuiteQL uses API-level pagination via URL parameters, NOT SQL OFFSET/LIMIT.
    The 'offset' parameter is added to the API URL.
    
    Args:
        base_query: SQL query (the API handles pagination)
        page_size: Rows per page (max 1000 for NetSuite)
        max_pages: Safety limit to prevent infinite loops
        timeout: Request timeout in seconds
    
    Returns:
        List of all rows from all pages
    """
    all_rows = []
    offset = 0
    page_num = 0
    
    while page_num < max_pages:
        page_num += 1
        
        # NetSuite pagination is done via URL parameters, not SQL syntax!
        # Add offset to the URL: /query/v1/suiteql?offset=X&limit=Y
        paginated_url = f"{suiteql_url}?limit={page_size}&offset={offset}"
        
        response = requests.post(
            paginated_url,
            auth=auth,
            headers={'Content-Type': 'application/json', 'Prefer': 'transient'},
            json={'q': base_query},
            timeout=timeout
        )
        
        if response.status_code != 200:
            print(f"❌ NetSuite error on page {page_num}: {response.status_code}", flush=True)
            print(f"   Response: {response.text[:500]}", flush=True)
            raise Exception(f"NetSuite API error: {response.status_code}")
        
        result = response.json()
        rows = result.get('items', [])
        
        print(f"   Page {page_num}: {len(rows)} rows (total: {len(all_rows) + len(rows)})", flush=True)
        
        all_rows.extend(rows)
        
        # If we got fewer rows than page_size, we've reached the end
        if len(rows) < page_size:
            break
        
        offset += page_size
    
    if page_num >= max_pages:
        print(f"⚠️ Reached max page limit ({max_pages})", flush=True)
    
    return all_rows


def convert_month_to_period_name(month_str):
    """Convert 'YYYY-MM' to 'Mon YYYY' format"""
    try:
        dt = datetime.strptime(month_str, '%Y-%m')
        return dt.strftime('%b %Y')
    except:
        return month_str


def extract_year_from_period(period_name):
    """Extract year from 'Jan 2024' format"""
    try:
        parts = period_name.split()
        if len(parts) == 2:
            return int(parts[1])
    except:
        pass
    return datetime.now().year


@app.route('/batch/full_year_refresh', methods=['POST'])
def batch_full_year_refresh():
    """
    OPTIMIZED FULL-YEAR REFRESH - Get ALL P&L accounts for an entire fiscal year in ONE query.
    Uses optimized CTE pattern that consolidates FIRST, then groups (much faster).
    
    Expected performance: < 30 seconds for ALL accounts × 12 months
    
    POST JSON:
    {
        "year": 2025,  // Optional - defaults to current year
        "subsidiary": "",
        "class": "",
        "department": "",
        "location": ""
    }
    
    Returns:
    {
        "balances": {
            "4010": {
                "Jan 2025": 12400000,
                "Feb 2025": 13200000,
                ...
            },
            "5000": {
                "Jan 2025": 5000000,
                ...
            }
        }
    }
    """
    data = request.get_json() or {}
    
    # Extract year - default to current year if not provided
    fiscal_year = data.get('year')
    if not fiscal_year:
        # Try to extract from first period if provided
        periods = data.get('periods', [])
        if periods:
            fiscal_year = extract_year_from_period(periods[0])
        else:
            fiscal_year = datetime.now().year
    
    # Get filters
    subsidiary = data.get('subsidiary', '')
    class_id = data.get('class', '')
    department = data.get('department', '')
    location = data.get('location', '')
    
    # Convert names to IDs
    subsidiary = convert_name_to_id('subsidiary', subsidiary)
    class_id = convert_name_to_id('class', class_id)
    department = convert_name_to_id('department', department)
    location = convert_name_to_id('location', location)
    
    # Determine target subsidiary for consolidation
    if subsidiary and subsidiary != '':
        target_sub = subsidiary
    else:
        target_sub = default_subsidiary_id or '1'
    
    # Build filters dict
    filters = {}
    if subsidiary and subsidiary != '':
        filters['subsidiary'] = subsidiary
    if class_id and class_id != '':
        filters['class'] = class_id
    if department and department != '':
        filters['department'] = department
    if location and location != '':
        filters['location'] = location
    
    try:
        print(f"\n{'='*80}", flush=True)
        print(f"🚀 FULL YEAR REFRESH: {fiscal_year}", flush=True)
        print(f"   Target subsidiary: {target_sub}", flush=True)
        print(f"   Filters: {filters}", flush=True)
        print(f"{'='*80}\n", flush=True)
        
        # Build the base query (WITHOUT OFFSET/LIMIT)
        base_query = build_full_year_pl_query(fiscal_year, target_sub, filters)
        
        # Execute with pagination to overcome NetSuite's 1000-row limit
        # Using OFFSET X LIMIT Y syntax for stable pagination
        start_time = datetime.now()
        
        try:
            items = run_paginated_suiteql(base_query, page_size=1000, max_pages=20)
        except Exception as e:
            print(f"❌ Pagination error: {e}", flush=True)
            return jsonify({'error': f'NetSuite query failed: {str(e)}'}), 500
        
        elapsed = (datetime.now() - start_time).total_seconds()
        print(f"⏱️  Total query time: {elapsed:.2f} seconds", flush=True)
        
        print(f"✅ Received {len(items)} rows")
        
        # Transform results to nested dict: { account: { period: value } }
        # Also track account types for frontend filtering
        balances = {}
        account_types = {}  # { account_number: "Income" | "Expense" | etc. }
        
        # Debug: Track period counts per account
        period_counts = {}
        
        for row in items:
            account = row.get('account_number')
            acct_type = row.get('account_type', '')
            month_str = row.get('month')  # 'YYYY-MM' format
            amount = float(row.get('amount', 0))
            
            # Convert 'YYYY-MM' to 'Mon YYYY' format
            period_name = convert_month_to_period_name(month_str)
            
            if account not in balances:
                balances[account] = {}
                account_types[account] = acct_type
                period_counts[account] = 0
            balances[account][period_name] = amount
            period_counts[account] += 1
        
        print(f"📊 Returning {len(balances)} accounts (P&L)")
        
        # Debug: Show accounts with less than 12 periods
        incomplete_accounts = {k: v for k, v in period_counts.items() if v < 12}
        if incomplete_accounts:
            print(f"⚠️  Accounts with < 12 periods: {len(incomplete_accounts)}")
            for acct, count in list(incomplete_accounts.items())[:10]:
                print(f"   Account {acct}: {count} periods, data: {list(balances[acct].keys())}")
                
            # DEEP DEBUG: Check if account 4270 is in the raw query results
            # This helps identify if it's a query issue vs processing issue
            print(f"\n🔍 DEBUG: Checking raw query for account 4270...")
            debug_query = f"""
                SELECT 
                    a.acctnumber, 
                    TO_CHAR(ap.startdate,'YYYY-MM') AS month,
                    COUNT(*) as transaction_count
                FROM TransactionAccountingLine tal
                JOIN Transaction t ON t.id = tal.transaction
                JOIN Account a ON a.id = tal.account
                JOIN AccountingPeriod ap ON ap.id = t.postingperiod
                WHERE t.posting = 'T'
                    AND tal.posting = 'T'
                    AND tal.accountingbook = 1
                    AND EXTRACT(YEAR FROM ap.startdate) = {fiscal_year}
                    AND a.acctnumber = '4270'
                    AND ap.isyear = 'F' AND ap.isquarter = 'F'
                GROUP BY a.acctnumber, ap.startdate
                ORDER BY ap.startdate
            """
            try:
                debug_result = query_netsuite(debug_query)
                if isinstance(debug_result, list):
                    print(f"   Raw data for 4270: {len(debug_result)} periods")
                    for row in debug_result:
                        print(f"      {row.get('month')}: {row.get('transaction_count')} transactions")
                else:
                    print(f"   Debug query error: {debug_result}")
            except Exception as e:
                print(f"   Debug query exception: {e}")
        
        # CRITICAL: Cache all results in backend for fast lookups
        # This allows individual formula requests to be instant after full refresh
        global balance_cache, balance_cache_timestamp
        balance_cache = {}
        balance_cache_timestamp = datetime.now()
        
        filters_hash = f"{subsidiary}:{department}:{location}:{class_id}"
        cached_count = 0
        
        print(f"🔑 Cache key format:")
        print(f"   subsidiary='{subsidiary}', department='{department}', location='{location}', class='{class_id}'")
        print(f"   filters_hash='{filters_hash}' (length: {len(filters_hash)}, colons: {filters_hash.count(':')})")
        
        for account, periods_data in balances.items():
            for period, amount in periods_data.items():
                cache_key = f"{account}:{period}:{filters_hash}"
                balance_cache[cache_key] = amount
                cached_count += 1
                
                # Show first 3 keys as examples
                if cached_count <= 3:
                    print(f"   Example key #{cached_count}: '{cache_key}' (length: {len(cache_key)}, colons: {cache_key.count(':')})")
        
        print(f"💾 Cached {cached_count} values on backend for instant formula lookups")
        print(f"{'='*80}\n")
        
        # ALSO fetch Balance Sheet accounts for the same year
        # OPTIMIZED: Query returns ACTIVITY per month, backend computes cumulative
        print(f"\n📊 Now fetching Balance Sheet accounts (OPTIMIZED - activity query)...", flush=True)
        
        # Clear BS activity cache
        global bs_activity_cache, bs_activity_cache_timestamp, bs_account_set
        bs_activity_cache = {}
        bs_activity_cache_timestamp = datetime.now()
        bs_account_set = set()
        
        try:
            bs_query = build_full_year_bs_query(fiscal_year, target_sub, filters)
            print(f"   BS Query (first 500 chars):\n{bs_query[:500]}...", flush=True)
            bs_start = datetime.now()
            # OPTIMIZED: Activity query is much faster than old cumulative query
            # Timeout reduced from 240s to 120s
            bs_items = run_paginated_suiteql(bs_query, page_size=1000, max_pages=20, timeout=120)
            bs_elapsed = (datetime.now() - bs_start).total_seconds()
            print(f"⏱️  BS query time: {bs_elapsed:.2f} seconds", flush=True)
            print(f"✅ BS returned {len(bs_items)} rows (account × month)", flush=True)
            
            # Process BS results - same format as P&L now (account, month, amount)
            # Store ACTIVITY in bs_activity_cache
            bs_account_count = 0
            bs_activity_data = {}  # { account: { period: activity } }
            
            for row in bs_items:
                account = row.get('account_number')
                acct_type = row.get('account_type', '')
                month_str = row.get('month')  # 'YYYY-MM' format
                amount = float(row.get('amount', 0))
                
                if not account or not month_str:
                    continue
                
                # Convert 'YYYY-MM' to 'Mon YYYY' format
                period_name = convert_month_to_period_name(month_str)
                
                # Track this as a BS account
                bs_account_set.add(account)
                
                if account not in bs_activity_data:
                    bs_activity_data[account] = {}
                    account_types[account] = acct_type
                    bs_account_count += 1
                
                # Store ACTIVITY (not cumulative)
                bs_activity_data[account][period_name] = amount
                
                # Cache activity for later cumulative calculations
                activity_cache_key = f"activity:{account}:{period_name}:{filters_hash}"
                bs_activity_cache[activity_cache_key] = amount
            
            print(f"📊 Loaded activity for {bs_account_count} Balance Sheet accounts", flush=True)
            
            # Now compute CUMULATIVE balances from activity
            # This is instant in Python (vs 60-90s in NetSuite)
            month_order = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                          'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            cumulative_count = 0
            
            for account, activity_by_period in bs_activity_data.items():
                if account not in balances:
                    balances[account] = {}
                
                cumulative = 0
                for month_abbrev in month_order:
                    period_name = f"{month_abbrev} {fiscal_year}"
                    
                    # Add this month's activity to running cumulative
                    activity = activity_by_period.get(period_name, 0)
                    cumulative += activity
                    
                    # Store CUMULATIVE balance (what formulas expect)
                    balances[account][period_name] = cumulative
                    
                    # Cache cumulative for formula lookups
                    cache_key = f"{account}:{period_name}:{filters_hash}"
                    balance_cache[cache_key] = cumulative
                    cached_count += 1
                    cumulative_count += 1
            
            print(f"📊 Computed {cumulative_count} cumulative BS balances from activity", flush=True)
            print(f"⚡ Cumulative calculation: instant (vs 60-90s with old query)", flush=True)
            
        except Exception as bs_error:
            print(f"⚠️  BS query error (P&L still succeeded): {bs_error}", flush=True)
            import traceback
            traceback.print_exc()
            # Don't fail the whole request - P&L data is still valid
        
        total_elapsed = elapsed + bs_elapsed if 'bs_elapsed' in dir() else elapsed
        print(f"💾 Total cached: {cached_count} values (P&L + BS)")
        print(f"📊 Total accounts: {len(balances)} (P&L + BS)")
        print(f"⏱️  Total time: {total_elapsed:.2f} seconds")
        
        # Fetch account names in ONE query to avoid 429 concurrency errors
        # This prevents 35+ parallel requests when Guide Me writes formulas
        account_names = {}
        try:
            account_numbers = list(balances.keys())
            if account_numbers:
                # Batch query for account names - IN clause with all account numbers
                account_list = "', '".join([escape_sql(str(a)) for a in account_numbers])
                names_query = f"""
                    SELECT acctnumber AS number, accountsearchdisplaynamecopy AS name
                    FROM Account
                    WHERE acctnumber IN ('{account_list}')
                """
                names_result = query_netsuite(names_query)
                if isinstance(names_result, list):
                    for row in names_result:
                        account_names[str(row.get('number', ''))] = row.get('name', '')
                print(f"📛 Fetched {len(account_names)} account names in ONE query")
        except Exception as names_error:
            print(f"⚠️  Account names fetch error (non-fatal): {names_error}")
        
        print(f"{'='*80}\n")
        
        return jsonify({
            'balances': balances,
            'account_types': account_types,  # { account_number: "Income" | "Expense" | etc. }
            'account_names': account_names,  # { account_number: "Account Name" }
            'query_time': total_elapsed, 
            'cached_count': cached_count,
            'pl_time': elapsed,
            'bs_time': bs_elapsed if 'bs_elapsed' in dir() else 0
        })
    
    except requests.exceptions.Timeout:
        print("❌ Query timeout (> 5 minutes)")
        return jsonify({'error': 'Query timeout - this should not happen with optimized query!'}), 504
    
    except Exception as e:
        print(f"❌ Error in full_year_refresh: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/batch/full_year_refresh_bs', methods=['POST'])
def batch_full_year_refresh_bs():
    """
    BALANCE SHEET ONLY full-year refresh.
    Use this when you specifically need BS accounts without P&L.
    
    POST JSON:
    {
        "year": 2025,
        "subsidiary": "",
        ...
    }
    
    Returns same structure as full_year_refresh but only BS accounts.
    """
    data = request.get_json() or {}
    
    fiscal_year = data.get('year')
    if not fiscal_year:
        periods = data.get('periods', [])
        if periods:
            fiscal_year = extract_year_from_period(periods[0])
        else:
            fiscal_year = datetime.now().year
    
    subsidiary = convert_name_to_id('subsidiary', data.get('subsidiary', ''))
    class_id = convert_name_to_id('class', data.get('class', ''))
    department = convert_name_to_id('department', data.get('department', ''))
    location = convert_name_to_id('location', data.get('location', ''))
    
    target_sub = subsidiary if subsidiary else (default_subsidiary_id or '1')
    
    filters = {}
    if subsidiary: filters['subsidiary'] = subsidiary
    if class_id: filters['class'] = class_id
    if department: filters['department'] = department
    if location: filters['location'] = location
    
    try:
        print(f"\n{'='*80}", flush=True)
        print(f"📊 BALANCE SHEET FULL YEAR REFRESH: {fiscal_year}", flush=True)
        print(f"   Target subsidiary: {target_sub}", flush=True)
        print(f"   Filters: {filters}", flush=True)
        print(f"{'='*80}\n", flush=True)
        
        bs_query = build_full_year_bs_query(fiscal_year, target_sub, filters)
        print(f"   BS Query (first 500 chars):\n{bs_query[:500]}...", flush=True)
        
        start_time = datetime.now()
        # BS query with WIDE format returns 1 row per account, unlikely to exceed 1000
        # Use 4-minute timeout since this is a complex query with 12 CONSOLIDATE calls
        items = run_paginated_suiteql(bs_query, page_size=1000, max_pages=5, timeout=240)
        elapsed = (datetime.now() - start_time).total_seconds()
        
        print(f"⏱️  Query time: {elapsed:.2f} seconds", flush=True)
        print(f"✅ Received {len(items)} rows")
        
        # Process BS results - WIDE format with columns like Jan_2024, Feb_2024, etc.
        balances = {}
        month_map = {
            'jan': 'Jan', 'feb': 'Feb', 'mar': 'Mar', 'apr': 'Apr',
            'may': 'May', 'jun': 'Jun', 'jul': 'Jul', 'aug': 'Aug',
            'sep': 'Sep', 'oct': 'Oct', 'nov': 'Nov', 'dec': 'Dec'
        }
        
        global balance_cache, balance_cache_timestamp
        filters_hash = f"{subsidiary}:{department}:{location}:{class_id}"
        cached_count = 0
        
        for row in items:
            account = row.get('account_number')
            if not account:
                continue
                
            if account not in balances:
                balances[account] = {}
            
            # Extract month columns (format: Jan_2024, Feb_2024, etc.)
            for key, value in row.items():
                if key in ('account_number', 'account_type'):
                    continue
                
                # Parse column name like "Jan_2024" -> "Jan 2024"
                if '_' in str(key):
                    parts = key.split('_')
                    if len(parts) == 2:
                        month_abbrev = parts[0].lower()
                        year = parts[1]
                        if month_abbrev in month_map:
                            period_name = f"{month_map[month_abbrev]} {year}"
                            amount = float(value or 0)
                            balances[account][period_name] = amount
                            
                            # Cache this result
                            cache_key = f"{account}:{period_name}:{filters_hash}"
                            balance_cache[cache_key] = amount
                            cached_count += 1
        
        balance_cache_timestamp = datetime.now()
        
        print(f"📊 Returning {len(balances)} BS accounts")
        print(f"💾 Cached {cached_count} BS values")
        print(f"{'='*80}\n")
        
        return jsonify({'balances': balances, 'query_time': elapsed, 'cached_count': cached_count})
        
    except Exception as e:
        print(f"❌ Error in full_year_refresh_bs: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


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
    
    # Check if we can serve this request from the backend balance cache
    # (populated by full year refresh)
    global balance_cache, balance_cache_timestamp
    
    if balance_cache and balance_cache_timestamp:
        from datetime import timedelta
        cache_age = (datetime.now() - balance_cache_timestamp).total_seconds()
        
        if cache_age < BALANCE_CACHE_TTL:
            # Cache is fresh! Try to serve from cache
            filters_hash = f"{subsidiary}:{department}:{location}:{class_id}"
            
            print(f"🔍 Cache lookup:")
            print(f"   subsidiary='{subsidiary}', department='{department}', location='{location}', class='{class_id}'")
            print(f"   Filters hash: '{filters_hash}' (length: {len(filters_hash)}, colons: {filters_hash.count(':')})")
            print(f"   Sample accounts: {accounts[:3]}")
            print(f"   Sample periods: {periods[:3]}")
            print(f"   Total cached keys: {len(balance_cache)}")
            print(f"   Sample cached keys: {list(balance_cache.keys())[:3]}")
            
            # Try building a sample key to compare
            if accounts and periods:
                sample_key = f"{accounts[0]}:{periods[0]}:{filters_hash}"
                print(f"   Sample lookup key: '{sample_key}' (length: {len(sample_key)}, colons: {sample_key.count(':')})")
                print(f"   Key exists in cache: {sample_key in balance_cache}")
            
            # Check if ALL requested data is in cache
            all_in_cache = True
            missing_keys = []
            for account in accounts:
                for period in periods:
                    cache_key = f"{account}:{period}:{filters_hash}"
                    if cache_key not in balance_cache:
                        all_in_cache = False
                        if len(missing_keys) < 5:  # Only collect first 5 for debugging
                            missing_keys.append(cache_key)
            
            if all_in_cache:
                # Serve entirely from cache!
                print(f"⚡ BACKEND CACHE HIT: {len(accounts)} accounts × {len(periods)} periods (age: {cache_age:.1f}s)")
                
                result_balances = {}
                for account in accounts:
                    result_balances[account] = {}
                    for period in periods:
                        cache_key = f"{account}:{period}:{filters_hash}"
                        result_balances[account][period] = balance_cache.get(cache_key, 0)
                
                return jsonify({'balances': result_balances, 'from_cache': True})
            else:
                print(f"⚠️  Partial cache miss - missing keys (showing first 5):")
                for key in missing_keys:
                    print(f"     Missing: '{key}'")
        else:
            print(f"⚠️  Backend cache expired ({cache_age:.1f}s old) - falling back to full query")
    
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
            else:
                # FALLBACK: Period not in NetSuite's AccountingPeriod table
                # Calculate the end date from the period name (e.g., "Jan 2025" -> 1/31/2025)
                print(f"WARNING: Period '{period}' not found in NetSuite, calculating date...", file=sys.stderr)
                calc_end = calculate_period_end_date(period)
                if calc_end:
                    # Use period_id=None - BS query will need to handle this
                    period_info[period] = {'enddate': calc_end, 'id': None}
                    print(f"   Calculated end date: {calc_end}", file=sys.stderr)
        
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


@app.route('/account/<account_number>/type')
def get_account_type(account_number):
    """
    Get account type from account number
    Used by: NS.GLACCTTYPE(accountNumber)
    
    Returns: Account type (Income, Expense, Bank, etc.)
    """
    try:
        query = f"""
            SELECT accttype AS account_type
            FROM Account
            WHERE acctnumber = '{escape_sql(account_number)}'
        """
        
        result = query_netsuite(query)
        
        if isinstance(result, dict) and 'error' in result:
            return jsonify({'error': result['error']}), 500
            
        if not result or len(result) == 0:
            return 'Not Found', 404
            
        account_type = result[0].get('account_type', '')
        return account_type, 200, {'Content-Type': 'text/plain'}
        
    except Exception as e:
        print(f"Error in get_account_type: {str(e)}", file=sys.stderr)
        return jsonify({'error': str(e)}), 500


@app.route('/account/<account_number>/parent')
def get_account_parent(account_number):
    """
    Get parent account number from account number
    Used by: NS.GLAPARENT(accountNumber)
    
    Returns: Parent account number (or empty string if no parent)
    """
    try:
        query = f"""
            SELECT 
                a.acctnumber,
                p.acctnumber AS parent_number
            FROM Account a
            LEFT JOIN Account p ON a.parent = p.id
            WHERE a.acctnumber = '{escape_sql(account_number)}'
        """
        
        result = query_netsuite(query)
        
        if isinstance(result, dict) and 'error' in result:
            return jsonify({'error': result['error']}), 500
            
        if not result or len(result) == 0:
            return 'Not Found', 404
            
        parent_number = result[0].get('parent_number', '')
        return parent_number or '', 200, {'Content-Type': 'text/plain'}
        
    except Exception as e:
        print(f"Error in get_account_parent: {str(e)}", file=sys.stderr)
        return jsonify({'error': str(e)}), 500


@app.route('/account/preload_titles')
def preload_account_titles():
    """
    Preload ALL account titles into cache with a single query
    This prevents 429 rate limit errors from concurrent individual requests
    
    Returns: Count of titles loaded
    """
    global account_title_cache
    
    try:
        print("🔄 Preloading ALL account titles...")
        
        # Query ALL active accounts in one go
        query = """
            SELECT acctnumber, accountsearchdisplaynamecopy AS account_name
            FROM Account
            WHERE isinactive = 'F'
            ORDER BY acctnumber
        """
        
        result = query_netsuite(query)
        
        if isinstance(result, dict) and 'error' in result:
            return jsonify({'error': result['error']}), 500
        
        # Populate cache
        loaded_count = 0
        if isinstance(result, list):
            for row in result:
                account_num = str(row.get('acctnumber', ''))
                account_name = row.get('account_name', 'Unknown')
                if account_num:
                    account_title_cache[account_num] = account_name
                    loaded_count += 1
        
        print(f"✅ Preloaded {loaded_count} account titles into cache")
        return jsonify({'loaded': loaded_count, 'status': 'success'})
            
    except Exception as e:
        print(f"Error preloading account titles: {str(e)}", file=sys.stderr)
        return jsonify({'error': str(e)}), 500


@app.route('/account/<account_number>/name')
def get_account_name(account_number):
    """
    Get account name from account number
    Used by: NS.GLATITLE(accountNumber)
    
    Returns: Account display name (string)
    """
    global account_title_cache
    
    try:
        # Check cache first
        if account_number in account_title_cache:
            # print(f"⚡ Title cache HIT: {account_number}")  # Uncomment for debugging
            return account_title_cache[account_number]
        
        # Cache miss - query NetSuite (ONLY if not preloaded)
        # This should rarely happen if preload_titles was called
        print(f"⚠️  Title cache MISS for account {account_number} - querying NetSuite")
        
        # Build SuiteQL query
        # Use accountsearchdisplaynamecopy to get name WITHOUT account number prefix
        query = f"""
            SELECT accountsearchdisplaynamecopy AS account_name
            FROM Account
            WHERE acctnumber = '{escape_sql(account_number)}'
        """
        
        result = query_netsuite(query)
        
        # Check for errors
        if isinstance(result, dict) and 'error' in result:
            return jsonify({'error': result['error']}), 500
        
        # Return account name or "Not Found"
        if result and len(result) > 0:
            account_name = result[0].get('account_name', 'Not Found')
        else:
            account_name = 'Not Found'
        
        # Cache the result (even if Not Found, to avoid repeated queries)
        account_title_cache[account_number] = account_name
        print(f"📝 Cached title for account {account_number}: {account_name}")
        
        return account_name
            
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
                from_start, from_end, _ = get_period_dates_from_name(from_period)
                to_start, to_end, _ = get_period_dates_from_name(to_period)
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
        elif to_period:
            # BALANCE SHEET CASE: Empty from_period, only to_period provided
            # This means "cumulative from beginning of time through to_period"
            if to_period.isdigit():
                where_clauses.append(f"t.postingperiod <= {to_period}")
            else:
                _, to_end, _ = get_period_dates_from_name(to_period)
                if to_end:
                    where_clauses.append(f"ap.enddate <= '{to_end}'")
                else:
                    where_clauses.append(f"ap.periodname = '{escape_sql(to_period)}'")
        
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
        
        print(f"DEBUG - Transaction drill-down request:", file=sys.stderr)
        print(f"  Account: {account}", file=sys.stderr)
        print(f"  Period: {period}", file=sys.stderr)
        print(f"  Subsidiary: {subsidiary}", file=sys.stderr)
        print(f"  Department: {department}", file=sys.stderr)
        print(f"  Class: {class_id}", file=sys.stderr)
        print(f"  Location: {location}", file=sys.stderr)
        
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
        
        print(f"DEBUG - Transaction drill-down query:\n{query[:500]}...", file=sys.stderr)
        result = query_netsuite(query)
        
        print(f"DEBUG - Query result type: {type(result)}", file=sys.stderr)
        if isinstance(result, list):
            print(f"DEBUG - Found {len(result)} transactions", file=sys.stderr)
        
        if isinstance(result, dict) and 'error' in result:
            print(f"DEBUG - Query error: {result}", file=sys.stderr)
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

@app.route('/lookups/accounts')
def get_all_accounts():
    """
    Get Income accounts for the Guide Me wizard.
    Returns account type, number, and name (clean display name).
    Filters to Income accounts only for a cleaner starter report.
    """
    try:
        # Filter to Income accounts for a focused starter report
        # Use accountsearchdisplaynamecopy for clean name (without hierarchy prefix)
        query = """
            SELECT 
                acctnumber AS number,
                accountsearchdisplaynamecopy AS name,
                accttype AS type
            FROM Account
            WHERE isinactive = 'F'
              AND accttype = 'Income'
            ORDER BY acctnumber
        """
        
        result = query_netsuite(query)
        
        if isinstance(result, dict) and 'error' in result:
            return jsonify({'error': result['error']}), 500
            
        accounts = []
        if isinstance(result, list):
            for row in result:
                accounts.append({
                    'type': row.get('type', ''),
                    'number': str(row.get('number', '')),
                    'name': row.get('name', '')
                })
        
        print(f"✓ Returning {len(accounts)} Income accounts")
        return jsonify(accounts)
        
    except Exception as e:
        print(f"❌ Account lookup error: {e}")
        return jsonify({'error': str(e)}), 500


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

