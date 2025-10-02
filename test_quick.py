"""
Quick Test - Verify MCP Server Basic Functionality

A minimal test script that quickly verifies the MCP server can connect
and read basic data from Univer.

Usage:
    python test_quick.py
"""

import asyncio
from mcp_server import UniverSheetsController


async def quick_test():
    """Quick smoke test"""
    print("🚀 Starting quick test...\n")
    
    controller = UniverSheetsController()
    
    try:
        # Connect
        print("1. Connecting to Univer...")
        await controller.start("http://localhost:3002/sheets/", headless=False)
        print("   ✅ Connected!\n")
        
        # Get status
        print("2. Getting workbook status...")
        status = await controller.get_activity_status()
        print(f"   ✅ Active sheet: {status['activeSheetName']}")
        print(f"   ✅ Total sheets: {status['sheetCount']}\n")
        
        # Get some data
        print("3. Reading range A1:C3...")
        data = await controller.get_range_data("A1:C3")
        print(f"   ✅ Retrieved {len(data['values'])} rows\n")
        
        # List sheets
        print("4. Getting all sheets...")
        sheets = await controller.get_sheets()
        print(f"   ✅ Found {len(sheets)} sheet(s):")
        for sheet in sheets:
            print(f"      - {sheet['name']}")
        
        print("\n✨ All tests passed! MCP server is working.\n")
        
    except Exception as e:
        print(f"\n❌ Test failed: {e}\n")
        import traceback
        traceback.print_exc()
    finally:
        await controller.cleanup()


if __name__ == "__main__":
    asyncio.run(quick_test())

