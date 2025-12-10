"""
Azure Functions wrapper for NetSuite Excel Add-in Backend
Uses HTTP triggers to expose the Flask-like API endpoints
"""

import azure.functions as func
import json
import os
import logging
from datetime import datetime
from dateutil import parser as date_parser
from dateutil.relativedelta import relativedelta
import requests
from requests_oauthlib import OAuth1
from concurrent.futures import ThreadPoolExecutor, as_completed

# Import constants (same directory in deployment)
from constants import (
    AccountType, PL_TYPES_SQL, SIGN_FLIP_TYPES_SQL, INCOME_TYPES_SQL,
    BS_ASSET_TYPES_SQL, BS_LIABILITY_TYPES_SQL, BS_EQUITY_TYPES_SQL
)

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

# Configuration from environment variables
def get_config():
    """Load NetSuite configuration from environment variables"""
    config = {
        'account_id': os.environ.get('NETSUITE_ACCOUNT_ID'),
        'consumer_key': os.environ.get('NETSUITE_CONSUMER_KEY'),
        'consumer_secret': os.environ.get('NETSUITE_CONSUMER_SECRET'),
        'token_id': os.environ.get('NETSUITE_TOKEN_ID'),
        'token_secret': os.environ.get('NETSUITE_TOKEN_SECRET'),
    }
    
    if not all(config.values()):
        missing = [k for k, v in config.items() if not v]
        raise ValueError(f"Missing environment variables: {missing}")
    
    return config

def get_auth():
    """Create OAuth1 authentication object"""
    config = get_config()
    return OAuth1(
        client_key=config['consumer_key'],
        client_secret=config['consumer_secret'],
        resource_owner_key=config['token_id'],
        resource_owner_secret=config['token_secret'],
        realm=config['account_id'],
        signature_method='HMAC-SHA256'
    )

def get_suiteql_url():
    """Get the SuiteQL API URL"""
    config = get_config()
    return f"https://{config['account_id']}.suitetalk.api.netsuite.com/services/rest/query/v1/suiteql"

def query_netsuite(sql_query: str, timeout: int = 30) -> dict:
    """Execute a SuiteQL query against NetSuite"""
    try:
        headers = {
            'Content-Type': 'application/json',
            'Prefer': 'transient'
        }
        
        response = requests.post(
            get_suiteql_url(),
            auth=get_auth(),
            headers=headers,
            json={'q': sql_query},
            timeout=timeout
        )
        
        if response.status_code == 200:
            return response.json()
        else:
            logging.error(f"NetSuite API error: {response.status_code} - {response.text}")
            return {'error': f"NetSuite API error: {response.status_code}"}
            
    except requests.exceptions.Timeout:
        return {'error': 'Request timeout'}
    except Exception as e:
        logging.error(f"Query error: {str(e)}")
        return {'error': str(e)}

def json_response(data: dict, status_code: int = 200) -> func.HttpResponse:
    """Create a JSON HTTP response with CORS headers for web access"""
    return func.HttpResponse(
        json.dumps(data),
        status_code=status_code,
        mimetype="application/json",
        headers={
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Requested-With",
            "Access-Control-Max-Age": "86400",
            "Cache-Control": "no-cache"
        }
    )


# ============================================================================
# Root Endpoint (for easy web access testing)
# ============================================================================

@app.route(route="", methods=["GET", "OPTIONS"])
def root(req: func.HttpRequest) -> func.HttpResponse:
    """Root endpoint - API information"""
    if req.method == "OPTIONS":
        return json_response({})
    
    return json_response({
        "name": "NetSuite Excel Add-in API",
        "version": "2.0.0",
        "status": "running",
        "endpoints": {
            "/": "API information (this page)",
            "/health": "Health check",
            "/api/balance": "Get GL account balance",
            "/api/budget": "Get budget amount",
            "/api/account/name": "Get account name",
            "/api/account/type": "Get account type",
            "/api/accounts": "List all accounts",
            "/api/periods": "List accounting periods",
            "/api/suiteql": "Execute custom SuiteQL query"
        },
        "docs": "https://github.com/chris-cloudextend/netsuite-excel-addin"
    })

# ============================================================================
# HTTP Trigger Functions
# ============================================================================

@app.route(route="health", methods=["GET", "OPTIONS"])
def health(req: func.HttpRequest) -> func.HttpResponse:
    """Health check endpoint"""
    if req.method == "OPTIONS":
        return json_response({})
    
    return json_response({
        "status": "ok",
        "message": "NetSuite Excel Add-in Backend (Azure Functions) is running",
        "version": "2.0.0",
        "timestamp": datetime.utcnow().isoformat()
    })

@app.route(route="api/balance", methods=["GET", "OPTIONS"])
def get_balance(req: func.HttpRequest) -> func.HttpResponse:
    """Get GL account balance for a period range"""
    if req.method == "OPTIONS":
        return json_response({})
    
    try:
        account = req.params.get('account')
        start_period = req.params.get('startPeriod')
        end_period = req.params.get('endPeriod')
        subsidiary = req.params.get('subsidiary', '')
        
        if not account or not start_period:
            return json_response({"error": "Missing required parameters: account, startPeriod"}, 400)
        
        # Build the SuiteQL query
        sql = f"""
            SELECT 
                SUM(tal.amount) as balance
            FROM 
                transactionaccountingline tal
                JOIN transaction t ON tal.transaction = t.id
                JOIN accountingperiod ap ON t.postingperiod = ap.id
                JOIN account a ON tal.account = a.id
            WHERE 
                a.acctnumber = '{account}'
                AND ap.periodname = '{start_period}'
                AND t.posting = 'T'
        """
        
        if subsidiary:
            sql += f" AND t.subsidiary = '{subsidiary}'"
        
        result = query_netsuite(sql)
        
        if 'error' in result:
            return json_response(result, 500)
        
        items = result.get('items', [])
        balance = items[0].get('balance', 0) if items else 0
        
        return json_response({
            "account": account,
            "startPeriod": start_period,
            "endPeriod": end_period,
            "balance": float(balance) if balance else 0
        })
        
    except Exception as e:
        logging.error(f"Balance error: {str(e)}")
        return json_response({"error": str(e)}, 500)

@app.route(route="api/account/name", methods=["GET", "OPTIONS"])
def get_account_name(req: func.HttpRequest) -> func.HttpResponse:
    """Get account name by account number"""
    if req.method == "OPTIONS":
        return json_response({})
    
    try:
        account = req.params.get('account')
        
        if not account:
            return json_response({"error": "Missing required parameter: account"}, 400)
        
        sql = f"SELECT acctnumber, acctname FROM account WHERE acctnumber = '{account}'"
        result = query_netsuite(sql)
        
        if 'error' in result:
            return json_response(result, 500)
        
        items = result.get('items', [])
        if items:
            return json_response({
                "account": account,
                "name": items[0].get('acctname', '')
            })
        else:
            return json_response({"error": f"Account {account} not found"}, 404)
            
    except Exception as e:
        logging.error(f"Account name error: {str(e)}")
        return json_response({"error": str(e)}, 500)

@app.route(route="api/account/type", methods=["GET", "OPTIONS"])
def get_account_type(req: func.HttpRequest) -> func.HttpResponse:
    """Get account type by account number"""
    if req.method == "OPTIONS":
        return json_response({})
    
    try:
        account = req.params.get('account')
        
        if not account:
            return json_response({"error": "Missing required parameter: account"}, 400)
        
        sql = f"SELECT acctnumber, accttype FROM account WHERE acctnumber = '{account}'"
        result = query_netsuite(sql)
        
        if 'error' in result:
            return json_response(result, 500)
        
        items = result.get('items', [])
        if items:
            return json_response({
                "account": account,
                "type": items[0].get('accttype', '')
            })
        else:
            return json_response({"error": f"Account {account} not found"}, 404)
            
    except Exception as e:
        logging.error(f"Account type error: {str(e)}")
        return json_response({"error": str(e)}, 500)

@app.route(route="api/budget", methods=["GET", "OPTIONS"])
def get_budget(req: func.HttpRequest) -> func.HttpResponse:
    """Get budget amount for an account and period"""
    if req.method == "OPTIONS":
        return json_response({})
    
    try:
        account = req.params.get('account')
        start_period = req.params.get('startPeriod')
        end_period = req.params.get('endPeriod')
        
        if not account or not start_period:
            return json_response({"error": "Missing required parameters: account, startPeriod"}, 400)
        
        sql = f"""
            SELECT 
                SUM(b.amount) as budget
            FROM 
                budgets b
                JOIN account a ON b.account = a.id
                JOIN accountingperiod ap ON b.accountingperiod = ap.id
            WHERE 
                a.acctnumber = '{account}'
                AND ap.periodname = '{start_period}'
        """
        
        result = query_netsuite(sql)
        
        if 'error' in result:
            return json_response(result, 500)
        
        items = result.get('items', [])
        budget = items[0].get('budget', 0) if items else 0
        
        return json_response({
            "account": account,
            "startPeriod": start_period,
            "endPeriod": end_period,
            "budget": float(budget) if budget else 0
        })
        
    except Exception as e:
        logging.error(f"Budget error: {str(e)}")
        return json_response({"error": str(e)}, 500)

@app.route(route="api/suiteql", methods=["POST", "OPTIONS"])
def execute_suiteql(req: func.HttpRequest) -> func.HttpResponse:
    """Execute a custom SuiteQL query"""
    if req.method == "OPTIONS":
        return json_response({})
    
    try:
        body = req.get_json()
        sql = body.get('query') or body.get('q')
        
        if not sql:
            return json_response({"error": "Missing query in request body"}, 400)
        
        result = query_netsuite(sql)
        return json_response(result)
        
    except ValueError:
        return json_response({"error": "Invalid JSON in request body"}, 400)
    except Exception as e:
        logging.error(f"SuiteQL error: {str(e)}")
        return json_response({"error": str(e)}, 500)

@app.route(route="api/accounts", methods=["GET", "OPTIONS"])
def list_accounts(req: func.HttpRequest) -> func.HttpResponse:
    """List all accounts"""
    if req.method == "OPTIONS":
        return json_response({})
    
    try:
        account_type = req.params.get('type', '')
        
        sql = "SELECT acctnumber, acctname, accttype FROM account WHERE isinactive = 'F'"
        
        if account_type:
            sql += f" AND accttype = '{account_type}'"
        
        sql += " ORDER BY acctnumber"
        
        result = query_netsuite(sql)
        
        if 'error' in result:
            return json_response(result, 500)
        
        return json_response({
            "accounts": result.get('items', []),
            "count": len(result.get('items', []))
        })
        
    except Exception as e:
        logging.error(f"List accounts error: {str(e)}")
        return json_response({"error": str(e)}, 500)

@app.route(route="api/periods", methods=["GET", "OPTIONS"])
def list_periods(req: func.HttpRequest) -> func.HttpResponse:
    """List accounting periods"""
    if req.method == "OPTIONS":
        return json_response({})
    
    try:
        sql = """
            SELECT id, periodname, startdate, enddate, isquarter, isyear
            FROM accountingperiod 
            WHERE isinactive = 'F' AND isyear = 'F' AND isquarter = 'F'
            ORDER BY startdate DESC
        """
        
        result = query_netsuite(sql)
        
        if 'error' in result:
            return json_response(result, 500)
        
        return json_response({
            "periods": result.get('items', []),
            "count": len(result.get('items', []))
        })
        
    except Exception as e:
        logging.error(f"List periods error: {str(e)}")
        return json_response({"error": str(e)}, 500)

