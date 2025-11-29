#!/usr/bin/env python3
"""
Test script for NetSuite Excel Formulas backend
Run this to verify the backend server is working correctly
"""

import requests
import sys

SERVER_URL = "http://localhost:5002"

def test_health():
    """Test the health endpoint"""
    print("Testing /health endpoint...")
    try:
        response = requests.get(f"{SERVER_URL}/health", timeout=5)
        if response.status_code == 200:
            data = response.json()
            print(f"  ✓ Health check passed")
            print(f"    Account: {data.get('account')}")
            return True
        else:
            print(f"  ✗ Health check failed: {response.status_code}")
            return False
    except Exception as e:
        print(f"  ✗ Could not connect: {e}")
        return False


def test_netsuite_connection():
    """Test NetSuite connection"""
    print("\nTesting NetSuite connection...")
    try:
        response = requests.get(f"{SERVER_URL}/test", timeout=30)
        if response.status_code == 200:
            data = response.json()
            print(f"  ✓ NetSuite connection successful")
            print(f"    Account: {data.get('account')}")
            print(f"    Active accounts: {data.get('active_accounts')}")
            return True
        else:
            data = response.json()
            print(f"  ✗ NetSuite connection failed")
            print(f"    Error: {data.get('error')}")
            return False
    except Exception as e:
        print(f"  ✗ Error: {e}")
        return False


def test_account_name():
    """Test getting account name"""
    print("\nTesting account name lookup...")
    try:
        response = requests.get(f"{SERVER_URL}/account/1000/name", timeout=10)
        if response.status_code == 200:
            name = response.text
            print(f"  ✓ Account name retrieved")
            print(f"    Account 1000: {name}")
            return True
        else:
            print(f"  ✗ Failed to get account name: {response.status_code}")
            return False
    except Exception as e:
        print(f"  ✗ Error: {e}")
        return False


def test_balance():
    """Test getting account balance"""
    print("\nTesting account balance...")
    try:
        params = {
            'account': '1000',
            'from_period': 'Jan 2025',
            'to_period': 'Mar 2025'
        }
        response = requests.get(f"{SERVER_URL}/balance", params=params, timeout=10)
        if response.status_code == 200:
            balance = response.text
            print(f"  ✓ Balance retrieved")
            print(f"    Account 1000 (Jan-Mar 2025): ${balance}")
            return True
        else:
            print(f"  ✗ Failed to get balance: {response.status_code}")
            return False
    except Exception as e:
        print(f"  ✗ Error: {e}")
        return False


def test_budget():
    """Test getting budget amount"""
    print("\nTesting budget amount...")
    try:
        params = {
            'account': '5000',
            'from_period': 'Jan 2025',
            'to_period': 'Dec 2025'
        }
        response = requests.get(f"{SERVER_URL}/budget", params=params, timeout=10)
        if response.status_code == 200:
            budget = response.text
            print(f"  ✓ Budget retrieved")
            print(f"    Account 5000 (Jan-Dec 2025): ${budget}")
            return True
        else:
            print(f"  ✗ Failed to get budget: {response.status_code}")
            return False
    except Exception as e:
        print(f"  ✗ Error: {e}")
        return False


def main():
    print("=" * 80)
    print("NetSuite Excel Formulas - Backend Test")
    print("=" * 80)
    print()
    print(f"Testing server at: {SERVER_URL}")
    print()
    
    results = []
    
    # Run tests
    results.append(("Health Check", test_health()))
    results.append(("NetSuite Connection", test_netsuite_connection()))
    results.append(("Account Name", test_account_name()))
    results.append(("Account Balance", test_balance()))
    results.append(("Budget Amount", test_budget()))
    
    # Summary
    print()
    print("=" * 80)
    print("Test Summary")
    print("=" * 80)
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for name, result in results:
        status = "✓ PASS" if result else "✗ FAIL"
        print(f"{status:8} {name}")
    
    print()
    print(f"Results: {passed}/{total} tests passed")
    print("=" * 80)
    
    if passed == total:
        print()
        print("🎉 All tests passed! Backend is working correctly.")
        print()
        print("Next steps:")
        print("  1. Load the Excel add-in (see QUICKSTART.md)")
        print("  2. Try the formulas in Excel")
        return 0
    else:
        print()
        print("⚠️  Some tests failed. Check the errors above.")
        print()
        print("Common issues:")
        print("  - Make sure the server is running: python backend/server.py")
        print("  - Check your NetSuite credentials in backend/netsuite_config.json")
        print("  - Verify your NetSuite account has SuiteQL enabled")
        return 1


if __name__ == '__main__':
    sys.exit(main())

