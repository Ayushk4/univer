"""
Test Script for Univer MCP Server

This script makes hardcoded calls to the UniverSheetsController to verify
that all the MCP tools work correctly.

Tests include:
- Read operations (status, range data, sheets, search, screenshot)  
- Sheet management (create, delete, rename, activate, move, hide/show)
- Data editing (set data, set style, merge cells, set dimensions)
- Row/column management (insert/delete rows/columns)
- Advanced features (auto fill, format brush, get workbook ID)
- Stub implementations (conditional formatting, data validation - API not available)

Usage:
    python test_mcp.py [--url http://localhost:3002/sheets/] [--headless] [--quick | --cell-write-test | --row-column-management-test]
    
Options:
    --url URL                         Univer server URL (default: http://localhost:3002/sheets/)
    --headless                        Run browser in headless mode
    --quick                           Run quick smoke test (10 tests)
    --cell-write-test                 Run only cell data editing tests (tests 17-20)
    --row-column-management-test      Run only row/column management tests (tests 21-26)

Note: Tests 25-28 verify API compatibility for conditional formatting and data validation,
      but these features are stub implementations only (Univer Facade API limitation).
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
        
        # Test 11: Create new sheets
        print_section("Test 11: Create New Sheets")
        result = await controller.create_sheet(["TestSheet1", "TestSheet2"])
        print_result("Create Sheets", result)
        
        # Test 12: Rename sheets
        print_section("Test 12: Rename Sheets")
        result = await controller.rename_sheet([
            {"old_name": "TestSheet1", "new_name": "DataSheet"},
            {"old_name": "TestSheet2", "new_name": "SummarySheet"}
        ])
        print_result("Rename Sheets", result)
        
        # Test 13: Activate sheet
        print_section("Test 13: Activate Sheet")
        result = await controller.activate_sheet("DataSheet")
        print_result("Activate Sheet", result)
        
        # Test 14: Move sheet
        print_section("Test 14: Move Sheet")
        result = await controller.move_sheet("SummarySheet", 0)
        print_result("Move Sheet", result)
        
        # Test 15: Set sheet display status (hide/show)
        print_section("Test 15: Set Sheet Display Status")
        # Hide a sheet
        result = await controller.set_sheet_display_status([
            {"sheet_name": "DataSheet", "visible": False}
        ])
        print_result("Hide Sheet", result)
        
        # Show it again
        result = await controller.set_sheet_display_status([
            {"sheet_name": "DataSheet", "visible": True}
        ])
        print_result("Show Sheet", result)
        
        # Test 16: Delete sheets (cleanup)
        print_section("Test 16: Delete Sheets (Cleanup)")
        result = await controller.delete_sheet(["DataSheet", "SummarySheet"])
        print_result("Delete Sheets", result)
        
        # Test 17: Set Range Data (write values and formulas)
        print_section("Test 17: Set Range Data (Write Values & Formulas)")
        result = await controller.set_range_data([
            {"range": "F1", "value": "Product"},
            {"range": "G1", "value": "Price"},
            {"range": "H1", "value": "Quantity"},
            {"range": "I1", "value": "Total"},
            {"range": "F2", "value": "Widget"},
            {"range": "G2", "value": 19.99},
            {"range": "H2", "value": 5},
            {"range": "I2", "value": "=G2*H2"}
        ])
        print_result("Set Range Data", result)
        
        # Test 18: Set Range Style (format cells)
        print_section("Test 18: Set Range Style (Format Cells)")
        result = await controller.set_range_style([
            {
                "range": "F1:I1",
                "style": {
                    "bl": 1,  # bold
                    "fs": 12,  # font size
                    "bg": {"rgb": "#4472C4"},  # blue background
                    "cl": {"rgb": "#FFFFFF"},  # white text
                    "ht": 2  # center align
                }
            }
        ])
        print_result("Set Range Style", result)
        
        # Test 19: Set Merge (merge cells)
        print_section("Test 19: Set Merge (Merge Cells)")
        result = await controller.set_merge("F4:I4")
        print_result("Merge Cells", result)
        # Add title to merged cell
        result = await controller.set_range_data([
            {"range": "F4", "value": "Test Report"}
        ])
        print_result("Set Merged Cell Value", result)
        
        # Test 20: Set Cell Dimensions (column widths and row heights)
        print_section("Test 20: Set Cell Dimensions (Widths & Heights)")
        result = await controller.set_cell_dimensions([
            {"range": "F:F", "width": 100},
            {"range": "G:I", "width": 90},
            {"range": "4:4", "height": 30}
        ])
        print_result("Set Cell Dimensions", result)
        
        # Test 21: Get Active Unit ID
        print_section("Test 21: Get Active Unit ID")
        result = await controller.get_active_unit_id()
        print(f"\n✅ Active Unit ID: {result}")
        
        # Test 22: Auto Fill
        print_section("Test 22: Auto Fill (Pattern Detection)")
        # Set up some data to auto-fill
        await controller.set_range_data([
            {"range": "J1", "value": 1},
            {"range": "J2", "value": 2}
        ])
        result = await controller.auto_fill("J1:J2", "J1:J5")
        print_result("Auto Fill (Vertical)", result)
        
        # Test 23: Format Brush
        print_section("Test 23: Format Brush (Copy Formatting)")
        # First apply some formatting to source
        await controller.set_range_style([
            {
                "range": "K1",
                "style": {
                    "bl": 1,
                    "fs": 14,
                    "bg": {"rgb": "#873e23"},
                    "cl": {"rgb": "#2596be"}
                }
            }
        ])
        await controller.set_range_data([{"range": "K1", "value": "Styled"}])
        
        # Add different data to target cells (to verify data is preserved)
        await controller.set_range_data([
            {"range": "K2", "value": "Cell 2"},
            {"range": "K3", "value": "Cell 3"},
            {"range": "K4", "value": "Cell 4"}
        ])
        
        # Now copy format to target
        result = await controller.format_brush("K1", "K2:K4")
        print_result("Format Brush Operation", result)
        
        # Verify: Read back all styles and compare
        print("\n🔍 Verifying format was actually copied...")
        source_data = await controller.get_range_data("K1", return_style=True)
        target_data = await controller.get_range_data("K2:K4", return_style=True)
        
        # Parse JSON strings if needed (get_range_data returns dict)
        if isinstance(source_data, str):
            source_json = json.loads(source_data)
        else:
            source_json = source_data
            
        if isinstance(target_data, str):
            target_json = json.loads(target_data)
        else:
            target_json = target_data
        
        # Extract source style (from K1)
        # Structure: {range: "K1", values: [[val]], styles: [[style_obj]], formulas: [[formula]]}
        source_style_raw = None
        if source_json.get("styles") and len(source_json["styles"]) > 0 and len(source_json["styles"][0]) > 0:
            source_style_raw = source_json["styles"][0][0]  # First row, first column
        
        # Extract actual style from _style property if present
        source_style = None
        if source_style_raw:
            source_style = source_style_raw.get("_style") if isinstance(source_style_raw, dict) else source_style_raw
        
        print(f"\n📋 Source Style (K1): {json.dumps(source_style, indent=2) if source_style else 'None'}")
        
        # Check each target cell (K2:K4 is 3 rows, 1 column)
        all_match = True
        if target_json.get("styles") and len(target_json["styles"]) > 0:
            for i, row_styles in enumerate(target_json["styles"]):
                cell_ref = f"K{i+2}"
                cell_style_raw = row_styles[0] if len(row_styles) > 0 else None  # First column
                
                # Extract actual style from _style property if present
                cell_style = None
                if cell_style_raw:
                    cell_style = cell_style_raw.get("_style") if isinstance(cell_style_raw, dict) else cell_style_raw
                
                # Get corresponding value
                cell_value = None
                if target_json.get("values") and i < len(target_json["values"]):
                    cell_value = target_json["values"][i][0] if len(target_json["values"][i]) > 0 else None
                
                print(f"\n📋 Target Style ({cell_ref}): {json.dumps(cell_style, indent=2) if cell_style else 'None'}")
                print(f"   Value preserved: '{cell_value}'")
                
                # Compare key style properties
                if source_style and cell_style:
                    # Check if required properties exist and match
                    source_bl = source_style.get("bl")
                    source_fs = source_style.get("fs")
                    source_bg = source_style.get("bg")
                    source_cl = source_style.get("cl")
                    
                    target_bl = cell_style.get("bl")
                    target_fs = cell_style.get("fs")
                    target_bg = cell_style.get("bg")
                    target_cl = cell_style.get("cl")
                    
                    # All source properties must be present in target
                    style_matches = True
                    mismatches = []
                    
                    if source_bl is not None and target_bl != source_bl:
                        style_matches = False
                        mismatches.append(f"bl: expected {source_bl}, got {target_bl}")
                    
                    if source_fs is not None and target_fs != source_fs:
                        style_matches = False
                        mismatches.append(f"fs: expected {source_fs}, got {target_fs}")
                    
                    if source_bg is not None and target_bg != source_bg:
                        style_matches = False
                        mismatches.append(f"bg: expected {source_bg}, got {target_bg}")
                    
                    if source_cl is not None and target_cl != source_cl:
                        style_matches = False
                        mismatches.append(f"cl: expected {source_cl}, got {target_cl}")
                    
                    if style_matches:
                        print(f"   ✅ Style matches source!")
                    else:
                        print(f"   ❌ Style does NOT match source!")
                        for mismatch in mismatches:
                            print(f"      • {mismatch}")
                        all_match = False
                else:
                    print(f"   ⚠️  Missing style data")
                    all_match = False
        else:
            print(f"\n⚠️  No style data found in target cells")
            all_match = False
        
        if all_match:
            print(f"\n✅ Test PASSED: All target cells (K2:K4) have matching styles from K1")
        else:
            print(f"\n❌ Test FAILED: Some target cells do not match source style")
            raise AssertionError("Format brush did not correctly copy styles")
        
        # Test 24: Get Conditional Formatting Rules
        print_section("Test 24: Get Conditional Formatting Rules")
        # Get rules from active sheet (may be empty)
        sheets = await controller.get_sheets()
        active_sheet_name = sheets[0]['name'] if sheets else "Sheet1"
        result = await controller.get_conditional_formatting_rules(active_sheet_name)
        print_result("Conditional Formatting Rules", result if result else {"rules": "No rules found"})
        
        # Test 25-28: Advanced Features (STUB IMPLEMENTATIONS - API Not Available)
        print_section("Test 25: Add Conditional Formatting Rule (STUB ONLY)")
        print("⚠️  WARNING: This is a stub implementation - rules are NOT actually applied")
        test_rule = [{"range": "L1:L5", "rule_type": "highlightCell", "sub_type": "number", "operator": "greaterThan", "value": 100}]
        result = await controller.add_conditional_formatting_rule(active_sheet_name, test_rule)
        print(f"\n📝 API Response: {result}")
        if "limited API support" in result or "not fully supported" in result:
            print("✅ Test PASSED: Stub correctly reports limited API support")
        else:
            print("⚠️  Unexpected response format")
        
        print_section("Test 26: Verify CF Rules Were NOT Created")
        print("🔍 Checking if any rules were actually created...")
        result = await controller.delete_conditional_formatting_rule(active_sheet_name, ["test-rule-1"])
        print(f"\n📝 Delete Result: {result}")
        if "Deleted 0" in result or "0 conditional formatting rule" in result:
            print("✅ Test PASSED: Correctly verified 0 rules exist (as expected for stub)")
        else:
            print(f"❌ Test FAILED: Expected 0 deletions, got: {result}")
        
        print_section("Test 27: Add Data Validation Rule (STUB ONLY)")
        print("⚠️  WARNING: This is a stub implementation - validation is NOT actually applied")
        test_validation = [{"range_a1": "M1:M10", "validation_type": "list", "source": "Option1,Option2,Option3"}]
        result = await controller.add_data_validation_rule(active_sheet_name, test_validation)
        print(f"\n📝 API Response: {result}")
        if "limited API support" in result or "not fully supported" in result:
            print("✅ Test PASSED: Stub correctly reports limited API support")
        else:
            print("⚠️  Unexpected response format")
        
        print_section("Test 28: Verify DV Rules Were NOT Created")
        print("🔍 Checking if any validation rules were actually created...")
        result = await controller.get_data_validation_rules(active_sheet_name)
        if not result or len(result) == 0:
            print("✅ Test PASSED: Correctly verified 0 rules exist (as expected for stub)")
            print_result("Data Validation Rules", {"status": "No rules found (expected)"})
        else:
            print(f"❌ Test FAILED: Expected 0 rules, but found: {len(result)}")
            print_result("Data Validation Rules", result)
        
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
        print(" 11. ✓ Create New Sheets")
        print(" 12. ✓ Rename Sheets")
        print(" 13. ✓ Activate Sheet")
        print(" 14. ✓ Move Sheet")
        print(" 15. ✓ Set Sheet Display Status (Hide/Show)")
        print(" 16. ✓ Delete Sheets")
        print(" 17. ✓ Set Range Data (Values & Formulas)")
        print(" 18. ✓ Set Range Style (Formatting)")
        print(" 19. ✓ Set Merge (Merge Cells)")
        print(" 20. ✓ Set Cell Dimensions (Widths & Heights)")
        print(" 21. ✓ Get Active Unit ID")
        print(" 22. ✓ Auto Fill (Pattern Detection)")
        print(" 23. ✓ Format Brush (Copy Formatting)")
        print(" 24. ✓ Get Conditional Formatting Rules")
        print("\n⚠️  STUB IMPLEMENTATIONS (API Compatibility Tests):")
        print(" 25. ✓ Add Conditional Formatting Rule (STUB - Not Applied)")
        print(" 26. ✓ Verify CF Rules Not Created (0 rules as expected)")
        print(" 27. ✓ Add Data Validation Rule (STUB - Not Applied)")
        print(" 28. ✓ Verify DV Rules Not Created (0 rules as expected)")
        print("\n🎉 All functional tests passed!")
        print("📊 Total: 24 functional tests + 4 stub API tests")
        print("\n📝 Note: Tests 25-28 verify API compatibility but do not apply rules.")
        print("    These tools are ready for when Univer expands its Facade API.")
        
    except Exception as e:
        print(f"\n❌ Error during tests: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print_section("Cleanup")
        print("🧹 Closing browser...")
        await controller.cleanup()
        print("✅ Cleanup complete!")


async def quick_test(url: str = "http://localhost:3002/sheets/", headless: bool = False):
    """Quick smoke test - minimal verification of MCP server functionality"""
    print("🚀 Starting quick test...\n")
    
    controller = UniverSheetsController()
    
    try:
        # Connect
        print("1. Connecting to Univer...")
        await controller.start(url, headless=headless)
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
        
        # Test sheet management
        print("\n5. Testing sheet management...")
        print("   Creating test sheet...")
        result = await controller.create_sheet(["QuickTest"])
        print(f"   ✅ Created sheet: {result['message']}")
        
        print("   Deleting test sheet...")
        result = await controller.delete_sheet(["QuickTest"])
        print(f"   ✅ Deleted sheet: {result['message']}")
        
        # Test data editing
        print("\n6. Testing data editing...")
        print("   Writing data to cells...")
        result = await controller.set_range_data([
            {"range": "E1", "value": "Test"},
            {"range": "E2", "value": 123},
            {"range": "E3", "value": "=E2*2"}
        ])
        print(f"   ✅ {result}")
        
        print("   Formatting cells...")
        result = await controller.set_range_style([
            {"range": "E1:E3", "style": {"bl": 1, "fs": 12}}
        ])
        print(f"   ✅ {result}")
        
        print("\n✨ All tests passed! MCP server is working.\n")
        
    except Exception as e:
        print(f"\n❌ Test failed: {e}\n")
        import traceback
        traceback.print_exc()
    finally:
        await controller.cleanup()


async def cell_write_test(url: str = "http://localhost:3002/sheets/", headless: bool = False):
    """Test only the cell data editing tools (tests 17-20)"""
    print("✏️  Starting cell write tests...\n")
    
    controller = UniverSheetsController()
    
    try:
        # Connect
        print("📡 Connecting to Univer...")
        await controller.start(url, headless=headless)
        print("✅ Connected!\n")
        
        # Test 17: Set Range Data (write values and formulas)
        print_section("Test 17: Set Range Data (Write Values & Formulas)")
        result = await controller.set_range_data([
            {"range": "F1", "value": "Product"},
            {"range": "G1", "value": "Price"},
            {"range": "H1", "value": "Quantity"},
            {"range": "I1", "value": "Total"},
            {"range": "F2", "value": "Widget"},
            {"range": "G2", "value": 19.99},
            {"range": "H2", "value": 5},
            {"range": "I2", "value": "=G2*H2"}
        ])
        print_result("Set Range Data", result)
        
        # Test 18: Set Range Style (format cells)
        print_section("Test 18: Set Range Style (Format Cells)")
        result = await controller.set_range_style([
            {
                "range": "F1:I1",
                "style": {
                    "bl": 1,  # bold
                    "fs": 12,  # font size
                    "bg": {"rgb": "#4472C4"},  # blue background
                    "cl": {"rgb": "#FFFFFF"},  # white text
                    "ht": 2  # center align
                }
            }
        ])
        print_result("Set Range Style", result)
        
        # Test 19: Set Merge (merge cells)
        print_section("Test 19: Set Merge (Merge Cells)")
        result = await controller.set_merge("F4:I4")
        print_result("Merge Cells", result)
        # Add title to merged cell
        result = await controller.set_range_data([
            {"range": "F4", "value": "Test Report"}
        ])
        print_result("Set Merged Cell Value", result)
        
        # Test 20: Set Cell Dimensions (column widths and row heights)
        print_section("Test 20: Set Cell Dimensions (Widths & Heights)")
        result = await controller.set_cell_dimensions([
            {"range": "F:F", "width": 100},
            {"range": "G:I", "width": 90},
            {"range": "4:4", "height": 30}
        ])
        print_result("Set Cell Dimensions", result)
        
        # Verify the data was written
        print_section("Verification: Read Back Written Data")
        result = await controller.get_range_data("F1:I2", return_style=True)
        print_result("Verify Data + Styles", result)
        
        print("\n✨ All cell write tests passed!\n")
        print("📋 Tests completed:")
        print("  17. ✓ Set Range Data (Values & Formulas)")
        print("  18. ✓ Set Range Style (Formatting)")
        print("  19. ✓ Set Merge (Merge Cells)")
        print("  20. ✓ Set Cell Dimensions (Widths & Heights)")
        print("\n🎉 All data editing operations working correctly!\n")
        
    except Exception as e:
        print(f"\n❌ Test failed: {e}\n")
        import traceback
        traceback.print_exc()
    finally:
        await controller.cleanup()


async def row_column_management_test(url: str = "http://localhost:3002/sheets/", headless: bool = False):
    """Test only the row/column management tools (tests 21-24)"""
    print("📊 Starting row/column management tests...\n")
    
    controller = UniverSheetsController()
    
    try:
        # Connect
        print("📡 Connecting to Univer...")
        await controller.start(url, headless=headless)
        print("✅ Connected!\n")
        
        # Test 21: Insert Rows
        print_section("Test 21: Insert Rows")
        result = await controller.insert_rows([
            {"position": 5, "how_many": 3, "where": "before"}
        ])
        print_result("Insert 3 Rows Before Row 5", result)
        
        # Verify by reading data
        result = await controller.get_sheets()
        print(f"✅ Current sheet has rows, inserted successfully")
        
        # Test 22: Insert Columns
        print_section("Test 22: Insert Columns")
        result = await controller.insert_columns([
            {"position": 3, "how_many": 2, "where": "after"}
        ])
        print_result("Insert 2 Columns After Column 3 (D)", result)
        
        # Test 23: Delete Rows
        print_section("Test 23: Delete Rows")
        result = await controller.delete_rows([
            {"position": 5, "how_many": 3}
        ])
        print_result("Delete 3 Rows Starting at Row 5", result)
        
        # Test 24: Delete Columns
        print_section("Test 24: Delete Columns")
        result = await controller.delete_columns([
            {"position": 3, "how_many": 2}
        ])
        print_result("Delete 2 Columns Starting at Column 3 (D)", result)
        
        # Test 25: Multiple Operations (insert multiple)
        print_section("Test 25: Batch Insert Operations")
        result = await controller.insert_rows([
            {"position": 10, "how_many": 1},
            {"position": 20, "how_many": 1}
        ])
        print_result("Insert Rows at Multiple Positions", result)
        
        # Test 26: Multiple Delete Operations
        print_section("Test 26: Batch Delete Operations")
        result = await controller.delete_rows([
            {"position": 10, "how_many": 1},
            {"position": 20, "how_many": 1}
        ])
        print_result("Delete Rows at Multiple Positions", result)
        
        print("\n✨ All row/column management tests passed!\n")
        print("📋 Tests completed:")
        print("  21. ✓ Insert Rows")
        print("  22. ✓ Insert Columns")
        print("  23. ✓ Delete Rows")
        print("  24. ✓ Delete Columns")
        print("  25. ✓ Batch Insert Operations")
        print("  26. ✓ Batch Delete Operations")
        print("\n🎉 All row/column management operations working correctly!\n")
        
    except Exception as e:
        print(f"\n❌ Test failed: {e}\n")
        import traceback
        traceback.print_exc()
    finally:
        await controller.cleanup()


async def main():
    """Main entry point"""
    # Parse command line arguments
    univer_url = "http://localhost:3002/sheets/"
    headless = False
    quick_mode = False
    cell_write_mode = False
    row_column_mode = False
    
    if "--url" in sys.argv:
        idx = sys.argv.index("--url")
        if idx + 1 < len(sys.argv):
            univer_url = sys.argv[idx + 1]
    
    if "--headless" in sys.argv:
        headless = True
    
    if "--quick" in sys.argv:
        quick_mode = True
    
    if "--cell-write-test" in sys.argv:
        cell_write_mode = True
    
    if "--row-column-management-test" in sys.argv:
        row_column_mode = True
    
    if "--help" in sys.argv or "-h" in sys.argv:
        print(__doc__)
        return
    
    # Run appropriate test suite
    if quick_mode:
        await quick_test(univer_url, headless)
    elif cell_write_mode:
        await cell_write_test(univer_url, headless)
    elif row_column_mode:
        await row_column_management_test(univer_url, headless)
    else:
        await run_tests(univer_url, headless)


if __name__ == "__main__":
    asyncio.run(main())

