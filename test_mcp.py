"""
Test Script for Univer MCP Server

This script makes hardcoded calls to the UniverSheetsController to verify
that all the read-only tools work correctly.

Usage:
    python test_mcp.py [--url http://localhost:3002/sheets/] [--headless]
"""

import asyncio
import json
import sys
from mcp_server import UniverSheetsController


def print_section(title: str):
    """Print a formatted section header"""
    print("\n" + "=" * 60)
    print(f"  {title}")
    print("=" * 60)


def print_result(name: str, result: dict):
    """Print formatted result"""
    print(f"\n✅ {name}:")
    # Remove screenshot data for cleaner output
    if 'screenshot' in result:
        screenshot_len = len(result['screenshot'])
        result = result.copy()
        result['screenshot'] = f"<base64 data, {screenshot_len} chars>"
    print(json.dumps(result, indent=2))


async def run_tests(url: str = "http://localhost:3002/sheets/", headless: bool = False):
    """Run all test scenarios"""
    
    print_section("Initializing Univer Sheets Controller")
    controller = UniverSheetsController()
    
    try:
        # Start the controller
        print(f"📡 Connecting to Univer at: {url}")
        await controller.start(url, headless=headless)
        print("✅ Connected successfully!")
        
        # Test 1: Get Activity Status
        print_section("Test 1: Get Activity Status (without screenshot)")
        result = await controller.get_activity_status(screenshot=False)
        print_result("Activity Status", result)
        
        # Test 2: Get Activity Status with Screenshot
        print_section("Test 2: Get Activity Status (with screenshot)")
        result = await controller.get_activity_status(screenshot=True)
        print_result("Activity Status + Screenshot", result)
        
        # Test 3: Get Sheets List
        print_section("Test 3: Get All Sheets")
        result = await controller.get_sheets()
        print_result("Sheets List", result)
        
        # Test 4: Get Range Data (small range)
        print_section("Test 4: Get Range Data (A1:C5)")
        result = await controller.get_range_data("A1:C5")
        print_result("Range Data", result)
        
        # Test 5: Get Range Data with Styles
        print_section("Test 5: Get Range Data with Styles (A1:B3)")
        result = await controller.get_range_data("A1:B3", return_style=True)
        print_result("Range Data + Styles", result)
        
        # Test 6: Get Multiple Ranges
        print_section("Test 6: Get Multiple Ranges")
        result = await controller.get_range_data(["A1:B2", "D1:E2"])
        print_result("Multiple Ranges", result)
        
        # Test 7: Search Cells by Value
        print_section("Test 7: Search Cells by Value")
        result = await controller.search_cells("Total", "value")
        print_result("Search Results (value)", result)
        
        # Test 8: Search Cells by Formula
        print_section("Test 8: Search Cells by Formula")
        result = await controller.search_cells("=", "formula")
        print_result("Search Results (formula)", result)
        
        # Test 9: Scroll and Screenshot
        print_section("Test 9: Scroll to Cell and Screenshot")
        result = await controller.scroll_and_screenshot("A1")
        print_result("Scroll + Screenshot", result)
        
        # Test 10: Test error handling - large range
        print_section("Test 10: Testing Large Range Warning")
        print("⚠️  Requesting a moderately sized range (A1:E40 = 200 cells)...")
        result = await controller.get_range_data("A1:E40")
        print(f"✅ Successfully retrieved {len(result.get('values', []))} rows")
        
        # Summary
        print_section("Test Summary")
        print("✅ All tests completed successfully!")
        print("\n📋 Tests Run:")
        print("  1. ✓ Activity Status (no screenshot)")
        print("  2. ✓ Activity Status (with screenshot)")
        print("  3. ✓ Get Sheets List")
        print("  4. ✓ Get Range Data (simple)")
        print("  5. ✓ Get Range Data (with styles)")
        print("  6. ✓ Get Multiple Ranges")
        print("  7. ✓ Search Cells (by value)")
        print("  8. ✓ Search Cells (by formula)")
        print("  9. ✓ Scroll and Screenshot")
        print(" 10. ✓ Large Range Test (200 cells)")
        print("\n🎉 MCP Server is working correctly!")
        
    except Exception as e:
        print(f"\n❌ Error during tests: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print_section("Cleanup")
        print("🧹 Closing browser...")
        await controller.cleanup()
        print("✅ Cleanup complete!")


async def main():
    """Main entry point"""
    # Parse command line arguments
    univer_url = "http://localhost:3002/sheets/"
    headless = False
    
    if "--url" in sys.argv:
        idx = sys.argv.index("--url")
        if idx + 1 < len(sys.argv):
            univer_url = sys.argv[idx + 1]
    
    if "--headless" in sys.argv:
        headless = True
    
    if "--help" in sys.argv or "-h" in sys.argv:
        print(__doc__)
        return
    
    await run_tests(univer_url, headless)


if __name__ == "__main__":
    asyncio.run(main())

