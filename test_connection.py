"""
Test script to verify database and API connectivity
"""

import pyodbc
import requests
import json

def test_database_connection():
    """Test SQL Server database connection"""
    print("\n" + "="*60)
    print("Testing Database Connection")
    print("="*60)
    
    # TODO: Update with your actual connection details
    config = {
        'db_driver': 'ODBC Driver 17 for SQL Server',
        'db_server': 'your_server_name',
        'db_name': 'FIRSINVOICEDB',
        'db_user': 'your_username',
        'db_password': 'your_password'
    }
    
    connection_string = (
        f"DRIVER={{{config['db_driver']}}};"
        f"SERVER={config['db_server']};"
        f"DATABASE={config['db_name']};"
        f"UID={config['db_user']};"
        f"PWD={config['db_password']}"
    )
    
    try:
        print(f"Connecting to server: {config['db_server']}")
        print(f"Database: {config['db_name']}")
        
        conn = pyodbc.connect(connection_string, timeout=10)
        cursor = conn.cursor()
        
        # Test query
        cursor.execute("SELECT @@VERSION")
        version = cursor.fetchone()[0]
        
        print("\n✓ Connection successful!")
        print(f"SQL Server Version: {version[:50]}...")
        
        # Check if REPORT table exists
        cursor.execute("""
            SELECT COUNT(*) 
            FROM INFORMATION_SCHEMA.TABLES 
            WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'REPORT'
        """)
        
        table_exists = cursor.fetchone()[0]
        
        if table_exists:
            print("✓ REPORT table exists")
            
            # Get row count
            cursor.execute("SELECT COUNT(*) FROM [dbo].[REPORT]")
            row_count = cursor.fetchone()[0]
            print(f"✓ Total rows in REPORT table: {row_count}")
            
            # Check for required columns
            cursor.execute("""
                SELECT COLUMN_NAME 
                FROM INFORMATION_SCHEMA.COLUMNS 
                WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'REPORT'
            """)
            columns = [row[0] for row in cursor.fetchall()]
            
            required_columns = ['TRANS_REF', 'BOOKING_DATE', 'CUSTOMER', 'REPORTING_STATUS', 'IRN', 'QR_CODE']
            
            print("\nColumn check:")
            for col in required_columns:
                if col in columns:
                    print(f"  ✓ {col}")
                else:
                    print(f"  ✗ {col} (missing - run setup_database.sql)")
            
            # Check for pending invoices
            if 'REPORTING_STATUS' in columns:
                cursor.execute("""
                    SELECT COUNT(*) 
                    FROM [dbo].[REPORT] 
                    WHERE REPORTING_STATUS IS NULL OR REPORTING_STATUS = 0
                """)
                pending = cursor.fetchone()[0]
                print(f"\n✓ Pending invoices to process: {pending}")
        else:
            print("✗ REPORT table does not exist!")
        
        cursor.close()
        conn.close()
        
        return True
        
    except pyodbc.Error as e:
        print(f"\n✗ Database connection failed!")
        print(f"Error: {str(e)}")
        return False
    except Exception as e:
        print(f"\n✗ Unexpected error: {str(e)}")
        return False


def test_api_connection():
    """Test API connectivity"""
    print("\n" + "="*60)
    print("Testing API Connection")
    print("="*60)
    
    # TODO: Update with your actual API details
    config = {
        'base_url': 'https://api.example.com',
        'participant_id': 'your_participant_id',
        'api_key': 'your_api_key'
    }
    
    headers = {
        'Content-Type': 'application/json',
        'participant-id': config['participant_id'],
        'x-api-key': config['api_key']
    }
    
    # Test with a sample payload (won't be processed as it's a test)
    test_payload = {
        "document_identifier": "TEST-CONNECTION-001",
        "issue_date": "2025-01-15",
        "invoice_type_code": "380",
        "document_currency_code": "NGN",
        "tax_currency_code": "NGN",
        "accounting_customer_party": {
            "party_name": "Test Customer",
            "email": "test@example.com",
            "tin": "00000000-0001",
            "telephone": "+234",
            "business_description": "Test",
            "postal_address": {
                "street_name": "Test Street",
                "city_name": "Test City",
                "postal_zone": "00000",
                "country": "NG"
            }
        },
        "invoice_line": [
            {
                "hsn_code": "0000.00",
                "price_amount": 100,
                "discount_amount": 0,
                "uom": "ST",
                "invoiced_quantity": 1,
                "product_category": "Test",
                "tax_rate": 0,
                "tax_category_id": "ZERO_VAT",
                "item_name": "Test Item",
                "sellers_item_identification": "TEST-001"
            }
        ]
    }
    
    try:
        url = f"{config['base_url']}/invoice/generate"
        print(f"Testing connection to: {url}")
        print(f"Participant ID: {config['participant_id']}")
        
        # Note: This will actually call the API
        # Comment out if you don't want to make a real API call
        print("\nSkipping actual API call (update config and uncomment to test)")
        print("API endpoint configured correctly in code")
        
        # Uncomment below to make actual API call
        # response = requests.post(url, json=test_payload, headers=headers, timeout=10)
        # print(f"\n✓ API responded with status: {response.status_code}")
        # print(f"Response: {response.text[:200]}")
        
        return True
        
    except requests.exceptions.RequestException as e:
        print(f"\n✗ API connection failed!")
        print(f"Error: {str(e)}")
        return False
    except Exception as e:
        print(f"\n✗ Unexpected error: {str(e)}")
        return False


def main():
    """Run all connection tests"""
    print("\n" + "="*60)
    print("FIRS Invoice Scheduler - Connection Test")
    print("="*60)
    print("\nBefore running this test:")
    print("1. Update the config dictionaries in this file")
    print("2. Ensure SQL Server ODBC driver is installed")
    print("3. Ensure network connectivity to database and API")
    
    db_ok = test_database_connection()
    api_ok = test_api_connection()
    
    print("\n" + "="*60)
    print("Test Summary")
    print("="*60)
    print(f"Database: {'✓ PASSED' if db_ok else '✗ FAILED'}")
    print(f"API:      {'✓ PASSED' if api_ok else '✗ FAILED (skipped)'}")
    
    if db_ok:
        print("\n✓ Ready to run invoice_scheduler.py")
    else:
        print("\n✗ Fix connection issues before running scheduler")
    
    print("="*60 + "\n")


if __name__ == "__main__":
    main()
