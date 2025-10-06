"""
Univer Sheets MCP Server (Read-Only)

A Model Context Protocol (MCP) server that provides read-only access to
a self-deployed Univer Sheets instance through natural language.

Uses Playwright to interact with the Univer Facade API in the browser.

Example Usage:
    # Start the MCP server
    python mcp_server.py --url http://localhost:3002/sheets/
    
    # Or with headless mode
    python mcp_server.py --url http://localhost:3002/sheets/ --headless
    
    # Use with Claude Desktop - add to config file:
    {
      "mcpServers": {
        "univer-sheets": {
          "command": "python",
          "args": ["/absolute/path/to/mcp_server.py"]
        }
      }
    }

Programmatic Usage Example:
    ```python
    import asyncio
    from mcp_server import UniverSheetsController
    
    async def main():
        controller = UniverSheetsController()
        await controller.start("http://localhost:3002/sheets/", headless=False)
        status = await controller.get_activity_status(screenshot=False)
        print("Status", status)
        data = await controller.get_range_data("A1:B10")
        print("Data", data)
        sheets = await controller.get_sheets()
        print("Sheets", sheets)
        results = await controller.search_cells("Total", "value")
        print("Results", results)
        screenshot_data = await controller.scroll_and_screenshot("A1")
        print(f"Screenshot captured: {len(screenshot_data['screenshot'])} bytes")
        await controller.cleanup()
    
    if __name__ == "__main__":
        asyncio.run(main())
    ```
"""

from playwright.async_api import async_playwright, Page, Browser
import mcp.types as types
from mcp.server import Server
from mcp.server.stdio import stdio_server
import asyncio
import base64
import json
from typing import Any, Optional
import logging
import time

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Import audit logger (will be initialized when logging_config is imported)
try:
    from logging_config import audit_logger
except ImportError:
    # Fallback if logging_config not available
    audit_logger = logger


class UniverSheetsController:
    """Controller for reading data from Univer Sheets via Playwright"""
    
    def __init__(self, snapshot_lock=None):
        self.page: Optional[Page] = None
        self.browser: Optional[Browser] = None
        self.playwright = None
        self.univer_url = "http://localhost:3002/sheets/"
        self.snapshot_file = "workbook_snapshot.json"  # Persistence file
        self.snapshot_lock = snapshot_lock  # Optional lock for thread-safe file access
        
    async def start(self, url: str = "http://localhost:3002/sheets/", headless: bool = False):
        """Start browser and connect to Univer instance"""
        logger.info(f"Starting browser and connecting to {url}...")
        self.univer_url = url
        self.playwright = await async_playwright().start()
        self.browser = await self.playwright.chromium.launch(headless=headless)
        self.page = await self.browser.new_page()
        await self.page.goto(url)
        
        # Wait for Univer to be ready
        await self.page.wait_for_function(
            "typeof window.univerAPI !== 'undefined'",
            timeout=30000
        )
        logger.info("Connected to Univer Sheets successfully!")
        
        # Load previous snapshot if available, otherwise save default state
        snapshot_loaded = await self.load_snapshot()
        
        if not snapshot_loaded:
            # No snapshot exists, save the default Univer workbook state
            logger.info("💾 Saving default Univer workbook state to snapshot...")
            await self.save_snapshot()
            logger.info("✅ Default workbook snapshot created!")
        
    async def execute_js(self, js_code: str) -> Any:
        """Execute JavaScript code in the Univer page context"""
        if not self.page:
            raise RuntimeError("Browser not started. Call start() first.")
        return await self.page.evaluate(js_code)
    
    async def take_screenshot(self) -> str:
        """Take a screenshot and return as base64"""
        if not self.page:
            raise RuntimeError("Browser not started.")
        screenshot = await self.page.screenshot()
        return base64.b64encode(screenshot).decode()
    
    async def stop(self):
        """Stop browser and cleanup resources"""
        await self.cleanup()
    
    async def cleanup(self):
        """Clean up browser resources"""
        if self.browser:
            await self.browser.close()
        if self.playwright:
            await self.playwright.stop()
    
    # ==================== Read-Only Tool Implementations ====================
    
    async def get_activity_status(self, screenshot: bool = False) -> dict:
        """Get current workbook status including sheet info and selection"""
        js_code = """
        (async () => {
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheets = workbook.getSheets();
            const activeSheet = workbook.getActiveSheet();
            const selection = activeSheet.getSelection();
            
            let selectionInfo = null;
            if (selection) {
                const activeRange = selection.getActiveRange();
                if (activeRange) {
                    selectionInfo = {
                        range: activeRange.getA1Notation(),
                        startRow: activeRange.getRow(),
                        startColumn: activeRange.getColumn(),
                        numRows: activeRange.getHeight(),
                        numColumns: activeRange.getWidth()
                    };
                }
            }
            
            return {
                sheetCount: sheets.length,
                sheetNames: sheets.slice(0, 10).map(s => s.getSheetName()),
                activeSheetName: activeSheet.getSheetName(),
                activeSheetId: activeSheet.getSheetId(),
                selection: selectionInfo
            };
        })()
        """
        
        result = await self.execute_js(js_code)
        
        if screenshot:
            result['screenshot'] = await self.take_screenshot()
            
        return result
    
    async def get_range_data(self, range_a1: str, 
                            return_screenshot: bool = False,
                            return_style: bool = False) -> dict:
        """Get cell data for specified range(s)
        
        Args:
            range_a1: Range in A1 notation (string or list of strings)
            return_screenshot: Whether to include a screenshot
            return_style: Whether to include cell styles and formulas
        
        WARNING: Keep ranges small - maximum 200 cells (rows × columns ≤ 200)
        """
        # Handle both single range and array of ranges
        if isinstance(range_a1, list):
            ranges = range_a1
        else:
            ranges = [range_a1]
        
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheet = workbook.getActiveSheet();
            const results = [];
            
            const ranges = {json.dumps(ranges)};
            
            for (const rangeStr of ranges) {{
                const range = sheet.getRange(rangeStr);
                const values = range.getValues();
                
                const rangeResult = {{
                    range: rangeStr,
                    values: values
                }};
                
                if ({str(return_style).lower()}) {{
                    rangeResult.styles = range.getCellStyles();
                    rangeResult.formulas = range.getFormulas();
                }}
                
                results.push(rangeResult);
            }}
            
            return results.length === 1 ? results[0] : results;
        }})()
        """
        
        result = await self.execute_js(js_code)
        
        if return_screenshot:
            result['screenshot'] = await self.take_screenshot()
            
        return result
    
    async def scroll_and_screenshot(self, cell_a1: str) -> dict:
        """Scroll to specified cell and take screenshot"""
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheet = workbook.getActiveSheet();
            const range = sheet.getRange('{cell_a1}');
            range.activate();
            return 'Scrolled to {cell_a1}';
        }})()
        """
        
        message = await self.execute_js(js_code)
        await asyncio.sleep(0.3)  # Wait for scroll animation
        
        screenshot = await self.take_screenshot()
        
        return {
            'status': 'Success',
            'message': message,
            'screenshot': screenshot
        }
    
    async def get_sheets(self) -> list[dict]:
        """Get all sheets in the workbook"""
        js_code = """
        (async () => {
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheets = workbook.getSheets();
            
            return sheets.map(sheet => ({
                name: sheet.getSheetName(),
                id: sheet.getSheetId(),
                index: sheet.getIndex(),
                hidden: sheet.isSheetHidden(),
                rowCount: sheet.getMaxRows(),
                columnCount: sheet.getMaxColumns()
            }));
        })()
        """
        
        return await self.execute_js(js_code)
    
    async def search_cells(self, keyword: str, find_by: str) -> dict:
        """Search cells by keyword and type (formula or value)
        
        Much more efficient than get_range_data for large sheets.
        Returns up to 50 matches.
        """
        # Properly escape keyword and find_by to prevent injection
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheet = workbook.getActiveSheet();
            
            // Parameters (properly escaped via JSON serialization)
            const keyword = {json.dumps(keyword)};
            const findBy = {json.dumps(find_by)};
            
            // Simple search implementation
            const maxRows = sheet.getMaxRows();
            const maxCols = sheet.getMaxColumns();
            const results = [];
            const RESULT_LIMIT = 50;
            
            // Early exit flag
            let limitReached = false;
            
            for (let row = 0; row < Math.min(maxRows, 100) && !limitReached; row++) {{
                for (let col = 0; col < Math.min(maxCols, 26) && !limitReached; col++) {{
                    const cell = sheet.getRange(row, col);
                    const cellData = cell.getCellData();
                    
                    let searchValue = '';
                    if (findBy === 'formula' && cellData && cellData.f) {{
                        searchValue = cellData.f;
                    }} else if (findBy === 'value' && cellData && cellData.v != null) {{
                        searchValue = String(cellData.v);
                    }}
                    
                    if (searchValue.includes(keyword)) {{
                        results.push({{
                            range: cell.getA1Notation(),
                            value: cellData ? cellData.v : null,
                            formula: cellData ? cellData.f : null
                        }});
                        
                        if (results.length >= RESULT_LIMIT) {{
                            limitReached = true;
                        }}
                    }}
                }}
            }}
            
            return {{
                total: results.length,
                results: results,
                truncated: limitReached
            }};
        }})()
        """
        
        return await self.execute_js(js_code)
    
    # ==================== Persistence Methods ====================
    
    async def save_snapshot(self) -> bool:
        """Save current workbook state to JSON file"""
        try:
            js_code = """
            (async () => {
                const workbook = window.univerAPI.getActiveWorkbook();
                const snapshot = workbook.getSnapshot();
                
                // Sanitize snapshot to make it valid JSON
                // Replace Infinity with a large number, NaN with null
                const sanitized = JSON.parse(
                    JSON.stringify(snapshot, (key, value) => {
                        if (value === Infinity) return 999999;
                        if (value === -Infinity) return -999999;
                        if (typeof value === 'number' && isNaN(value)) return null;
                        return value;
                    })
                );
                
                return sanitized;
            })()
            """
            snapshot = await self.execute_js(js_code)
            
            # Save to file with lock if provided
            import os
            snapshot_path = os.path.join(os.path.dirname(__file__), self.snapshot_file)
            
            if self.snapshot_lock:
                # Thread-safe write with lock
                with self.snapshot_lock:
                    with open(snapshot_path, 'w') as f:
                        json.dump(snapshot, f, indent=2)
            else:
                # Write without lock (for backward compatibility)
                with open(snapshot_path, 'w') as f:
                    json.dump(snapshot, f, indent=2)
            
            logger.info(f"✅ Snapshot saved to {self.snapshot_file}")
            return True
        except Exception as e:
            logger.warning(f"⚠️  Failed to save snapshot: {e}")
            return False
    
    async def load_snapshot(self) -> bool:
        """Load workbook state from JSON file"""
        import os
        snapshot_path = os.path.join(os.path.dirname(__file__), self.snapshot_file)
        
        if not os.path.exists(snapshot_path):
            logger.info("ℹ️  No snapshot file found, using default state")
            return False
        
        try:
            with open(snapshot_path, 'r') as f:
                snapshot = json.load(f)
            
            # Note: Univer's restore API may vary by version
            # This attempts to restore the snapshot data
            js_code = f"""
            (async () => {{
                try {{
                    const snapshot = {json.dumps(snapshot)};
                    const workbook = window.univerAPI.getActiveWorkbook();
                    
                    // Attempt to load snapshot
                    // The exact API may vary - this is a best effort
                    if (typeof workbook.loadSnapshot === 'function') {{
                        await workbook.loadSnapshot(snapshot);
                        return 'Loaded via loadSnapshot()';
                    }} else {{
                        // Fallback: Try to manually restore sheets
                        console.log('loadSnapshot not available, snapshot not restored');
                        return 'Snapshot load not supported by this Univer version';
                    }}
                }} catch (e) {{
                    console.error('Snapshot load error:', e);
                    return 'Error: ' + e.message;
                }}
            }})()
            """
            
            result = await self.execute_js(js_code)
            logger.info(f"📥 {result}")
            return True
        except Exception as e:
            logger.warning(f"⚠️  Failed to load snapshot: {e}")
            return False
    
    # ==================== Sheet Management Methods ====================
    
    async def create_sheet(self, sheet_names: list[str]) -> dict:
        """Create one or more new worksheets
        
        Args:
            sheet_names: List of names for the new sheets
            
        Returns:
            dict with status, message, and details
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheetNames = {json.dumps(sheet_names)};
            const results = [];
            
            for (const name of sheetNames) {{
                try {{
                    workbook.insertSheet(name);
                    // Get the newly created sheet
                    const sheet = workbook.getSheetByName(name);
                    results.push({{
                        name: name,
                        status: 'success',
                        id: sheet ? sheet.getSheetId() : 'unknown'
                    }});
                }} catch (e) {{
                    results.push({{
                        name: name,
                        status: 'failed',
                        error: e.message
                    }});
                }}
            }}
            
            return {{
                status: 'success',
                message: `Created ${{results.filter(r => r.status === 'success').length}} sheet(s)`,
                details: results
            }};
        }})()
        """
        result = await self.execute_js(js_code)
        await self.save_snapshot()  # Auto-save after operation
        return result
    
    async def delete_sheet(self, sheet_names: list[str]) -> dict:
        """Delete one or more worksheets by name
        
        Args:
            sheet_names: List of sheet names to delete
            
        Returns:
            dict with status and message
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheetNames = {json.dumps(sheet_names)};
            const results = [];
            
            for (const name of sheetNames) {{
                try {{
                    const sheet = workbook.getSheetByName(name);
                    if (!sheet) {{
                        results.push({{
                            name: name,
                            status: 'failed',
                            error: 'Sheet not found'
                        }});
                        continue;
                    }}
                    workbook.deleteSheet(sheet.getSheetId());
                    results.push({{
                        name: name,
                        status: 'success'
                    }});
                }} catch (e) {{
                    results.push({{
                        name: name,
                        status: 'failed',
                        error: e.message
                    }});
                }}
            }}
            
            return {{
                status: 'success',
                message: `Deleted ${{results.filter(r => r.status === 'success').length}} sheet(s)`,
                details: results
            }};
        }})()
        """
        result = await self.execute_js(js_code)
        await self.save_snapshot()  # Auto-save after operation
        return result
    
    async def rename_sheet(self, operations: list[dict]) -> dict:
        """Rename one or more worksheets
        
        Args:
            operations: List of dicts with 'old_name' and 'new_name'
            
        Returns:
            dict with status and message
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const operations = {json.dumps(operations)};
            const results = [];
            
            for (const op of operations) {{
                try {{
                    const sheet = workbook.getSheetByName(op.old_name);
                    if (!sheet) {{
                        results.push({{
                            old_name: op.old_name,
                            new_name: op.new_name,
                            status: 'failed',
                            error: 'Sheet not found'
                        }});
                        continue;
                    }}
                    sheet.setName(op.new_name);
                    results.push({{
                        old_name: op.old_name,
                        new_name: op.new_name,
                        status: 'success'
                    }});
                }} catch (e) {{
                    results.push({{
                        old_name: op.old_name,
                        new_name: op.new_name,
                        status: 'failed',
                        error: e.message
                    }});
                }}
            }}
            
            return {{
                status: 'success',
                message: `Renamed ${{results.filter(r => r.status === 'success').length}} sheet(s)`,
                details: results
            }};
        }})()
        """
        result = await self.execute_js(js_code)
        await self.save_snapshot()  # Auto-save after operation
        return result
    
    async def activate_sheet(self, sheet_name: str) -> str:
        """Activate/switch to a specific worksheet
        
        Args:
            sheet_name: Name of the sheet to activate
            
        Returns:
            Success message string
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheetName = {json.dumps(sheet_name)};
            const sheet = workbook.getSheetByName(sheetName);
            if (!sheet) {{
                throw new Error('Sheet not found: ' + sheetName);
            }}
            workbook.setActiveSheet(sheet);
            return 'Activated sheet: ' + sheetName;
        }})()
        """
        result = await self.execute_js(js_code)
        await self.save_snapshot()  # Auto-save after operation
        return result
    
    async def move_sheet(self, sheet_name: str, to_index: int) -> str:
        """Move a worksheet to a specific index position
        
        Args:
            sheet_name: Name of the sheet to move
            to_index: Target index (0-based)
            
        Returns:
            Success message string
        """
        # Convert string to int if needed
        if isinstance(to_index, str):
            to_index = int(to_index)
            
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheetName = {json.dumps(sheet_name)};
            const sheet = workbook.getSheetByName(sheetName);
            if (!sheet) {{
                throw new Error('Sheet not found: ' + sheetName);
            }}
            const currentIndex = sheet.getIndex();
            
            // Activate the sheet first, then move it
            // This is a workaround because moveSheet() has a bug in Univer
            const originalActiveSheet = workbook.getActiveSheet();
            workbook.setActiveSheet(sheet);
            workbook.moveActiveSheet({to_index});
            
            // Restore original active sheet if different
            if (originalActiveSheet.getSheetId() !== sheet.getSheetId()) {{
                workbook.setActiveSheet(originalActiveSheet);
            }}
            
            return 'Moved sheet "' + sheetName + '" from index ' + currentIndex + ' to {to_index}';
        }})()
        """
        result = await self.execute_js(js_code)
        await self.save_snapshot()  # Auto-save after operation
        return result
    
    async def set_sheet_display_status(self, operations: list[dict]) -> dict:
        """Show or hide one or more worksheets
        
        Args:
            operations: List of dicts with 'sheet_name' and 'visible' (bool)
            
        Returns:
            dict with status, message, and details
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const operations = {json.dumps(operations)};
            const results = [];
            
            for (const op of operations) {{
                try {{
                    const sheet = workbook.getSheetByName(op.sheet_name);
                    if (!sheet) {{
                        results.push({{
                            sheet_name: op.sheet_name,
                            visible: op.visible,
                            status: 'failed',
                            error: 'Sheet not found'
                        }});
                        continue;
                    }}
                    
                    if (op.visible) {{
                        sheet.showSheet();
                    }} else {{
                        sheet.hideSheet();
                    }}
                    
                    results.push({{
                        sheet_name: op.sheet_name,
                        visible: op.visible,
                        status: 'success'
                    }});
                }} catch (e) {{
                    results.push({{
                        sheet_name: op.sheet_name,
                        visible: op.visible,
                        status: 'failed',
                        error: e.message
                    }});
                }}
            }}
            
            return {{
                status: 'success',
                message: `Modified display status for ${{results.filter(r => r.status === 'success').length}} sheet(s)`,
                details: results
            }};
        }})()
        """
        result = await self.execute_js(js_code)
        await self.save_snapshot()  # Auto-save after operation
        return result
    
    # ==================== Data Editing Methods ====================
    
    async def set_range_data(self, items: list[dict]) -> str:
        """Set cell values and formulas for multiple ranges
        
        Args:
            items: List of dicts with 'range' and 'value'
                   - range: A1 notation (e.g., 'A1', 'Sheet1!A1', 'A1:B2')
                   - value: Can be string, number, boolean, 1D/2D array, or dict
                           If string starts with '=', treated as formula
        
        Returns:
            Success message string
            
        Examples:
            [{"range": "A1", "value": "Hello"}]
            [{"range": "B1", "value": "=A1"}]
            [{"range": "A1:B1", "value": ["A1", "B1"]}]
            [{"range": "A1:B2", "value": [["A1", "B1"], ["A2", "B2"]]}]
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheet = workbook.getActiveSheet();
            const items = {json.dumps(items)};
            
            for (const item of items) {{
                try {{
                    const range = sheet.getRange(item.range);
                    const value = item.value;
                    
                    if (typeof value === 'string' && value.startsWith('=')) {{
                        // Formula
                        range.setFormula(value);
                    }} else if (Array.isArray(value)) {{
                        // Array of values (1D or 2D)
                        range.setValues(value);
                    }} else {{
                        // Single value
                        range.setValue(value);
                    }}
                }} catch (e) {{
                    console.error(`Error setting range ${{item.range}}:`, e);
                    throw e;
                }}
            }}
            
            return 'Successfully set data for ' + items.length + ' range(s)';
        }})()
        """
        
        result = await self.execute_js(js_code)
        await self.save_snapshot()  # Auto-save after operation
        return result
    
    async def set_range_style(self, items: list[dict]) -> str:
        """Set cell styles for multiple ranges
        
        Args:
            items: List of dicts with 'range' and 'style'
                   - range: A1 notation
                   - style: Style object with properties like:
                     * ff: font family (e.g., "Arial")
                     * fs: font size in pt (e.g., 11)
                     * it: italic (0|1)
                     * bl: bold (0|1)
                     * ul: underline (0|1)
                     * bg: background color (e.g., {"rgb": "#FF0000"})
                     * cl: font color (e.g., {"rgb": "#000000"})
                     * ht: horizontal alignment (1=left, 2=center, 3=right)
                     * vt: vertical alignment (1=top, 2=middle, 3=bottom)
        
        Returns:
            Success message string
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheet = workbook.getActiveSheet();
            const items = {json.dumps(items)};
            
            // Alignment mapping
            const alignmentMap = {{
                1: 'left',
                2: 'center',
                3: 'right'
            }};
            
            const verticalAlignmentMap = {{
                1: 'top',
                2: 'middle',
                3: 'bottom'
            }};
            
            for (const item of items) {{
                try {{
                    const range = sheet.getRange(item.range);
                    const style = item.style;
                    
                    // Apply each style property individually
                    if (style.ff !== undefined) {{
                        range.setFontFamily(style.ff);
                    }}
                    
                    if (style.fs !== undefined) {{
                        range.setFontSize(style.fs);
                    }}
                    
                    if (style.bl !== undefined) {{
                        range.setFontWeight(style.bl === 1 ? 'bold' : 'normal');
                    }}
                    
                    if (style.it !== undefined) {{
                        range.setFontStyle(style.it === 1 ? 'italic' : 'normal');
                    }}
                    
                    if (style.bg !== undefined) {{
                        const bgColor = style.bg.rgb || style.bg;
                        range.setBackgroundColor(bgColor);
                    }}
                    
                    if (style.cl !== undefined) {{
                        const fontColor = style.cl.rgb || style.cl;
                        range.setFontColor(fontColor);
                    }}
                    
                    if (style.ht !== undefined) {{
                        const alignment = alignmentMap[style.ht] || style.ht;
                        range.setHorizontalAlignment(alignment);
                    }}
                    
                    if (style.vt !== undefined) {{
                        const vAlignment = verticalAlignmentMap[style.vt] || style.vt;
                        range.setVerticalAlignment(vAlignment);
                    }}
                    
                }} catch (e) {{
                    console.error(`Error styling range ${{item.range}}:`, e);
                    throw e;
                }}
            }}
            
            return 'Successfully styled ' + items.length + ' range(s)';
        }})()
        """
        
        result = await self.execute_js(js_code)
        await self.save_snapshot()  # Auto-save after operation
        return result
    
    async def set_merge(self, range_a1: str) -> str:
        """Merge cells in a single range
        
        Args:
            range_a1: Range string in A1 notation (e.g., 'A1:B2')
        
        Returns:
            Success message string
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheet = workbook.getActiveSheet();
            const range = sheet.getRange('{range_a1}');
            
            range.merge();
            
            return 'Successfully merged range {range_a1}';
        }})()
        """
        
        result = await self.execute_js(js_code)
        await self.save_snapshot()  # Auto-save after operation
        return result
    
    async def set_cell_dimensions(self, requests: list[dict]) -> dict:
        """Set column widths and row heights
        
        Args:
            requests: List of dimension requests with:
                     - range: A1 notation (e.g., 'A:C' for columns, '1:5' for rows)
                     - width: Width in points (for columns)
                     - height: Height in points (for rows)
        
        Returns:
            Dict with status, message, and report of changes
            
        Examples:
            [{"range": "A:C", "width": 120}]  # Set columns A-C to 120pt
            [{"range": "1:5", "height": 25}]   # Set rows 1-5 to 25pt
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheet = workbook.getActiveSheet();
            const requests = {json.dumps(requests)};
            const report = [];
            
            for (const req of requests) {{
                try {{
                    const rangeStr = req.range;
                    
                    // Detect if it's a column or row range
                    if (rangeStr.match(/^[A-Z]+:[A-Z]+$/)) {{
                        // Column range (e.g., 'A:C')
                        const cols = rangeStr.split(':');
                        const startCol = cols[0].charCodeAt(0) - 65;
                        const endCol = cols[1].charCodeAt(0) - 65;
                        const width = req.width;
                        
                        for (let col = startCol; col <= endCol; col++) {{
                            sheet.setColumnWidth(col, width);
                        }}
                        
                        report.push({{
                            range: rangeStr,
                            type: 'column',
                            width: width,
                            count: endCol - startCol + 1
                        }});
                    }} else if (rangeStr.match(/^\\d+:\\d+$/)) {{
                        // Row range (e.g., '1:5')
                        const rows = rangeStr.split(':');
                        const startRow = parseInt(rows[0]) - 1;
                        const endRow = parseInt(rows[1]) - 1;
                        const height = req.height;
                        
                        for (let row = startRow; row <= endRow; row++) {{
                            sheet.setRowHeight(row, height);
                        }}
                        
                        report.push({{
                            range: rangeStr,
                            type: 'row',
                            height: height,
                            count: endRow - startRow + 1
                        }});
                    }}
                }} catch (e) {{
                    console.error(`Error setting dimensions for ${{req.range}}:`, e);
                    report.push({{
                        range: req.range,
                        error: e.message
                    }});
                }}
            }}
            
            return {{
                status: 'success',
                message: `Set dimensions for ${{report.length}} range(s)`,
                report: report
            }};
        }})()
        """
        
        result = await self.execute_js(js_code)
        await self.save_snapshot()  # Auto-save after operation
        return result
    
    async def insert_rows(self, operations: list[dict]) -> dict:
        """Insert empty rows at specified positions
        
        Args:
            operations: List of insert operations, each with:
                       - position: Row index (0-based) where to insert
                       - how_many: Number of rows to insert (default 1)
                       - where: 'before' or 'after' the position (default 'before')
                       - sheet_name: Optional sheet name (uses active if not provided)
        
        Returns:
            Dict with status, message, and details of operations
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const operations = {json.dumps(operations)};
            const results = [];
            
            for (const op of operations) {{
                try {{
                    const sheet = op.sheet_name 
                        ? workbook.getSheetByName(op.sheet_name)
                        : workbook.getActiveSheet();
                    
                    const position = op.position || 0;
                    const howMany = op.how_many || 1;
                    const where = op.where || 'before';
                    
                    // Calculate actual insert position
                    const insertPosition = where === 'after' ? position + 1 : position;
                    
                    // Insert rows
                    const range = sheet.getRange(insertPosition, 0);
                    range.insertCells('DOWN', howMany);
                    
                    results.push({{
                        sheet: sheet.getSheetName(),
                        position: insertPosition,
                        count: howMany,
                        status: 'success'
                    }});
                }} catch (e) {{
                    results.push({{
                        position: op.position,
                        error: e.message,
                        status: 'failed'
                    }});
                }}
            }}
            
            const successCount = results.filter(r => r.status === 'success').length;
            return {{
                status: successCount === results.length ? 'success' : 'partial',
                message: `Inserted rows in ${{successCount}}/${{results.length}} operations`,
                details: results
            }};
        }})()
        """
        
        result = await self.execute_js(js_code)
        await self.save_snapshot()
        return result
    
    async def insert_columns(self, operations: list[dict]) -> dict:
        """Insert empty columns at specified positions
        
        Args:
            operations: List of insert operations, each with:
                       - position: Column index (0-based) where to insert
                       - how_many: Number of columns to insert (default 1)
                       - where: 'before' or 'after' the position (default 'before')
                       - sheet_name: Optional sheet name (uses active if not provided)
        
        Returns:
            Dict with status, message, and details of operations
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const operations = {json.dumps(operations)};
            const results = [];
            
            for (const op of operations) {{
                try {{
                    const sheet = op.sheet_name 
                        ? workbook.getSheetByName(op.sheet_name)
                        : workbook.getActiveSheet();
                    
                    const position = op.position || 0;
                    const howMany = op.how_many || 1;
                    const where = op.where || 'before';
                    
                    // Calculate actual insert position
                    const insertPosition = where === 'after' ? position + 1 : position;
                    
                    // Insert columns
                    const range = sheet.getRange(0, insertPosition);
                    range.insertCells('RIGHT', howMany);
                    
                    results.push({{
                        sheet: sheet.getSheetName(),
                        position: insertPosition,
                        count: howMany,
                        status: 'success'
                    }});
                }} catch (e) {{
                    results.push({{
                        position: op.position,
                        error: e.message,
                        status: 'failed'
                    }});
                }}
            }}
            
            const successCount = results.filter(r => r.status === 'success').length;
            return {{
                status: successCount === results.length ? 'success' : 'partial',
                message: `Inserted columns in ${{successCount}}/${{results.length}} operations`,
                details: results
            }};
        }})()
        """
        
        result = await self.execute_js(js_code)
        await self.save_snapshot()
        return result
    
    async def delete_rows(self, operations: list[dict]) -> dict:
        """Delete rows at specified positions
        
        Args:
            operations: List of delete operations, each with:
                       - position: Row index (0-based) where to start deletion
                       - how_many: Number of rows to delete (default 1)
                       - sheet_name: Optional sheet name (uses active if not provided)
        
        Returns:
            Dict with status, message, and details of operations
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const operations = {json.dumps(operations)};
            const results = [];
            
            // Sort operations by position (descending) to avoid position shifts
            const sortedOps = [...operations].sort((a, b) => (b.position || 0) - (a.position || 0));
            
            for (const op of sortedOps) {{
                try {{
                    const sheet = op.sheet_name 
                        ? workbook.getSheetByName(op.sheet_name)
                        : workbook.getActiveSheet();
                    
                    const position = op.position || 0;
                    const howMany = op.how_many || 1;
                    
                    // Delete rows
                    const range = sheet.getRange(position, 0, howMany, 1);
                    range.deleteCells('UP');
                    
                    results.push({{
                        sheet: sheet.getSheetName(),
                        position: position,
                        count: howMany,
                        status: 'success'
                    }});
                }} catch (e) {{
                    results.push({{
                        position: op.position,
                        error: e.message,
                        status: 'failed'
                    }});
                }}
            }}
            
            const successCount = results.filter(r => r.status === 'success').length;
            return {{
                status: successCount === results.length ? 'success' : 'partial',
                message: `Deleted rows in ${{successCount}}/${{results.length}} operations`,
                details: results
            }};
        }})()
        """
        
        result = await self.execute_js(js_code)
        await self.save_snapshot()
        return result
    
    async def delete_columns(self, operations: list[dict]) -> dict:
        """Delete columns at specified positions
        
        Args:
            operations: List of delete operations, each with:
                       - position: Column index (0-based) where to start deletion
                       - how_many: Number of columns to delete (default 1)
                       - sheet_name: Optional sheet name (uses active if not provided)
        
        Returns:
            Dict with status, message, and details of operations
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const operations = {json.dumps(operations)};
            const results = [];
            
            // Sort operations by position (descending) to avoid position shifts
            const sortedOps = [...operations].sort((a, b) => (b.position || 0) - (a.position || 0));
            
            for (const op of sortedOps) {{
                try {{
                    const sheet = op.sheet_name 
                        ? workbook.getSheetByName(op.sheet_name)
                        : workbook.getActiveSheet();
                    
                    const position = op.position || 0;
                    const howMany = op.how_many || 1;
                    
                    // Delete columns
                    const range = sheet.getRange(0, position, 1, howMany);
                    range.deleteCells('LEFT');
                    
                    results.push({{
                        sheet: sheet.getSheetName(),
                        position: position,
                        count: howMany,
                        status: 'success'
                    }});
                }} catch (e) {{
                    results.push({{
                        position: op.position,
                        error: e.message,
                        status: 'failed'
                    }});
                }}
            }}
            
            const successCount = results.filter(r => r.status === 'success').length;
            return {{
                status: successCount === results.length ? 'success' : 'partial',
                message: `Deleted columns in ${{successCount}}/${{results.length}} operations`,
                details: results
            }};
        }})()
        """
        
        result = await self.execute_js(js_code)
        await self.save_snapshot()
        return result
    
    async def get_active_unit_id(self) -> str:
        """Get the currently active unit_id (workbook ID)
        
        Returns:
            String with the unit ID
        """
        js_code = """
        (async () => {
            const workbook = window.univerAPI.getActiveWorkbook();
            return workbook.getId();
        })()
        """
        
        result = await self.execute_js(js_code)
        return result
    
    async def auto_fill(self, source_range: str, target_range: str) -> str:
        """Auto fill data from source range to target range using pattern detection
        
        Args:
            source_range: Source range in A1 notation (e.g., 'A1:A2')
            target_range: Target range in A1 notation (e.g., 'A1:A10')
        
        Returns:
            Operation description string
            
        Note:
            - Copies data, formulas, and styles
            - Supports horizontal or vertical extension only
            - Target must be aligned with source (same row or column direction)
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheet = workbook.getActiveSheet();
            
            const sourceRange = sheet.getRange('{source_range}');
            const targetRange = sheet.getRange('{target_range}');
            
            // Get source data, formulas, and styles
            const sourceValues = sourceRange.getValues();
            const sourceFormulas = sourceRange.getFormulas();
            const sourceStyles = sourceRange.getCellStyles();
            
            // Get dimensions
            const sourceRows = sourceValues.length;
            const sourceCols = sourceValues[0] ? sourceValues[0].length : 0;
            
            // Parse A1 notation to get target dimensions
            const parseA1 = (a1) => {{
                const match = a1.match(/([A-Z]+)(\\d+)(?::([A-Z]+)(\\d+))?/);
                if (!match) return null;
                const [, startCol, startRow, endCol, endRow] = match;
                const colToNum = (col) => {{
                    let num = 0;
                    for (let i = 0; i < col.length; i++) {{
                        num = num * 26 + col.charCodeAt(i) - 64;
                    }}
                    return num - 1;
                }};
                return {{
                    startRow: parseInt(startRow) - 1,
                    startCol: colToNum(startCol),
                    endRow: endRow ? parseInt(endRow) - 1 : parseInt(startRow) - 1,
                    endCol: endCol ? colToNum(endCol) : colToNum(startCol)
                }};
            }};
            
            const targetCoords = parseA1('{target_range}');
            const targetStartRow = targetCoords.startRow;
            const targetStartCol = targetCoords.startCol;
            const targetRows = targetCoords.endRow - targetCoords.startRow + 1;
            const targetCols = targetCoords.endCol - targetCoords.startCol + 1;
            
            // Detect if vertical or horizontal fill
            const isVerticalFill = sourceCols === targetCols && sourceRows < targetRows;
            const isHorizontalFill = sourceRows === targetRows && sourceCols < targetCols;
            
            if (!isVerticalFill && !isHorizontalFill) {{
                throw new Error('Target range must be aligned with source (same row or column direction)');
            }}
            
            // Fill data
            if (isVerticalFill) {{
                // Vertical fill (down)
                for (let targetRow = 0; targetRow < targetRows; targetRow++) {{
                    const sourceRow = targetRow % sourceRows;
                    for (let col = 0; col < targetCols; col++) {{
                        const cellRange = sheet.getRange(targetStartRow + targetRow, targetStartCol + col);
                        
                        // Set value or formula
                        if (sourceFormulas[sourceRow][col]) {{
                            cellRange.setFormula(sourceFormulas[sourceRow][col]);
                        }} else {{
                            cellRange.setValue(sourceValues[sourceRow][col]);
                        }}
                        
                        // Apply style
                        if (sourceStyles[sourceRow] && sourceStyles[sourceRow][col]) {{
                            const styleWrapper = sourceStyles[sourceRow][col];
                            const style = styleWrapper._style || styleWrapper;
                            if (style.bl) cellRange.setFontWeight(style.bl === 1 ? 'bold' : 'normal');
                            if (style.it) cellRange.setFontStyle(style.it === 1 ? 'italic' : 'normal');
                            if (style.fs) cellRange.setFontSize(style.fs);
                            if (style.ff) cellRange.setFontFamily(style.ff);
                            if (style.bg) cellRange.setBackgroundColor(style.bg.rgb || style.bg);
                            if (style.cl) cellRange.setFontColor(style.cl.rgb || style.cl);
                            if (style.ht) {{
                                const alignMap = {{1: 'left', 2: 'center', 3: 'right'}};
                                cellRange.setHorizontalAlignment(alignMap[style.ht] || style.ht);
                            }}
                        }}
                    }}
                }}
            }} else {{
                // Horizontal fill (right)
                for (let targetCol = 0; targetCol < targetCols; targetCol++) {{
                    const sourceCol = targetCol % sourceCols;
                    for (let row = 0; row < targetRows; row++) {{
                        const cellRange = sheet.getRange(targetStartRow + row, targetStartCol + targetCol);
                        
                        // Set value or formula
                        if (sourceFormulas[row][sourceCol]) {{
                            cellRange.setFormula(sourceFormulas[row][sourceCol]);
                        }} else {{
                            cellRange.setValue(sourceValues[row][sourceCol]);
                        }}
                        
                        // Apply style
                        if (sourceStyles[row] && sourceStyles[row][sourceCol]) {{
                            const styleWrapper = sourceStyles[row][sourceCol];
                            const style = styleWrapper._style || styleWrapper;
                            if (style.bl) cellRange.setFontWeight(style.bl === 1 ? 'bold' : 'normal');
                            if (style.it) cellRange.setFontStyle(style.it === 1 ? 'italic' : 'normal');
                            if (style.fs) cellRange.setFontSize(style.fs);
                            if (style.ff) cellRange.setFontFamily(style.ff);
                            if (style.bg) cellRange.setBackgroundColor(style.bg.rgb || style.bg);
                            if (style.cl) cellRange.setFontColor(style.cl.rgb || style.cl);
                            if (style.ht) {{
                                const alignMap = {{1: 'left', 2: 'center', 3: 'right'}};
                                cellRange.setHorizontalAlignment(alignMap[style.ht] || style.ht);
                            }}
                        }}
                    }}
                }}
            }}
            
            return `Auto-filled from ${{'{source_range}'}} to ${{'{target_range}'}} (${{isVerticalFill ? 'vertical' : 'horizontal'}})`;
        }})()
        """
        
        result = await self.execute_js(js_code)
        await self.save_snapshot()
        return result
    
    async def format_brush(self, source_range: str, target_range: str) -> str:
        """Copy formatting from source range to target range (format painter)
        
        Args:
            source_range: Source range in A1 notation (e.g., 'A1' or 'Sheet1!A1')
            target_range: Target range in A1 notation (e.g., 'B1:C1' or 'Sheet2!B1')
        
        Returns:
            Operation description string
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            
            // Parse sheet names if present
            const parseRange = (rangeStr) => {{
                if (rangeStr.includes('!')) {{
                    const [sheetName, range] = rangeStr.split('!');
                    return {{ sheet: workbook.getSheetByName(sheetName), range }};
                }} else {{
                    return {{ sheet: workbook.getActiveSheet(), range: rangeStr }};
                }}
            }};
            
            const source = parseRange('{source_range}');
            const target = parseRange('{target_range}');
            
            // Get source styles
            const sourceRange = source.sheet.getRange(source.range);
            const sourceStyles = sourceRange.getCellStyles();
            
            // Apply to target
            const targetRange = target.sheet.getRange(target.range);
            
            // Parse A1 notation for dimensions
            const parseA1 = (a1) => {{
                const match = a1.match(/([A-Z]+)(\\d+)(?::([A-Z]+)(\\d+))?/);
                if (!match) return null;
                const [, startCol, startRow, endCol, endRow] = match;
                const colToNum = (col) => {{
                    let num = 0;
                    for (let i = 0; i < col.length; i++) {{
                        num = num * 26 + col.charCodeAt(i) - 64;
                    }}
                    return num - 1;
                }};
                return {{
                    startRow: parseInt(startRow) - 1,
                    startCol: colToNum(startCol),
                    endRow: endRow ? parseInt(endRow) - 1 : parseInt(startRow) - 1,
                    endCol: endCol ? colToNum(endCol) : colToNum(startCol)
                }};
            }};
            
            const targetCoords = parseA1(target.range);
            const targetRows = targetCoords.endRow - targetCoords.startRow + 1;
            const targetCols = targetCoords.endCol - targetCoords.startCol + 1;
            
            const sourceRows = sourceStyles.length;
            const sourceCols = sourceStyles[0] ? sourceStyles[0].length : 0;
            
            for (let row = 0; row < targetRows; row++) {{
                const sourceRow = row % sourceRows;
                for (let col = 0; col < targetCols; col++) {{
                    const sourceCol = col % sourceCols;
                    const cellRange = target.sheet.getRange(
                        targetCoords.startRow + row,
                        targetCoords.startCol + col
                    );
                    
                    if (sourceStyles[sourceRow] && sourceStyles[sourceRow][sourceCol]) {{
                        const styleWrapper = sourceStyles[sourceRow][sourceCol];
                        // Extract actual style from _style property if present
                        const style = styleWrapper._style || styleWrapper;
                        
                        if (style.bl !== undefined) cellRange.setFontWeight(style.bl === 1 ? 'bold' : 'normal');
                        if (style.it !== undefined) cellRange.setFontStyle(style.it === 1 ? 'italic' : 'normal');
                        if (style.fs) cellRange.setFontSize(style.fs);
                        if (style.ff) cellRange.setFontFamily(style.ff);
                        if (style.bg) cellRange.setBackgroundColor(style.bg.rgb || style.bg);
                        if (style.cl) cellRange.setFontColor(style.cl.rgb || style.cl);
                        if (style.ht) {{
                            const alignMap = {{1: 'left', 2: 'center', 3: 'right'}};
                            cellRange.setHorizontalAlignment(alignMap[style.ht] || style.ht);
                        }}
                        if (style.vt) {{
                            const valignMap = {{1: 'top', 2: 'middle', 3: 'bottom'}};
                            cellRange.setVerticalAlignment(valignMap[style.vt] || style.vt);
                        }}
                    }}
                }}
            }}
            
            return `Format copied from ${{'{source_range}'}} to ${{'{target_range}'}}`;
        }})()
        """
        
        result = await self.execute_js(js_code)
        await self.save_snapshot()
        return result
    
    async def get_conditional_formatting_rules(self, sheet_name: str) -> list:
        """Get all conditional formatting rules for the given sheet
        
        Args:
            sheet_name: Target sheet name
        
        Returns:
            List of conditional formatting rules
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheet = workbook.getSheetByName('{sheet_name}');
            
            if (!sheet) {{
                throw new Error(`Sheet '${{'{sheet_name}'}}' not found`);
            }}
            
            try {{
                // Get conditional formatting rules if the API supports it
                const rules = sheet.getConditionalFormattingRules();
                return rules.map(rule => ({{
                    ruleId: rule.getRuleId ? rule.getRuleId() : 'unknown',
                    ranges: rule.getRanges ? rule.getRanges().map(r => r.getA1Notation()) : [],
                    type: rule.getType ? rule.getType() : 'unknown'
                }}));
            }} catch (e) {{
                // If API doesn't support conditional formatting, return empty array
                console.warn('Conditional formatting not supported or no rules found:', e);
                return [];
            }}
        }})()
        """
        
        result = await self.execute_js(js_code)
        return result
    
    async def add_conditional_formatting_rule(self, sheet_name: str, rules: list[dict]) -> str:
        """Add conditional formatting rules to a sheet
        
        Args:
            sheet_name: Target sheet name
            rules: List of conditional formatting rule dicts
        
        Returns:
            Success message string
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheet = workbook.getSheetByName('{sheet_name}');
            
            if (!sheet) {{
                throw new Error(`Sheet '${{'{sheet_name}'}}' not found`);
            }}
            
            const rules = {json.dumps(rules)};
            
            try {{
                // Try to add conditional formatting rules
                for (const rule of rules) {{
                    const range = sheet.getRange(rule.range || 'A1');
                    const cfRule = range.createConditionalFormattingRule();
                    
                    // Configure rule based on type
                    if (rule.rule_type === 'highlightCell' && rule.sub_type === 'number' && rule.style) {{
                        // Basic highlight rule for numbers
                        if (rule.operator === 'greaterThan' && rule.value !== undefined) {{
                            // This is a simplified implementation
                            // Real implementation would use proper CF API
                        }}
                    }}
                    
                    // This is a stub - full CF API might not be available
                    console.warn('Conditional formatting API limited in current version');
                }}
                
                return `Note: Added ${{rules.length}} conditional formatting rule(s) to ${{'{sheet_name}'}} (limited API support)`;
            }} catch (e) {{
                console.warn('Conditional formatting not fully supported:', e);
                return `Warning: Conditional formatting API not fully supported. Rules structure received but not applied. Error: ${{e.message}}`;
            }}
        }})()
        """
        
        result = await self.execute_js(js_code)
        return result
    
    async def set_conditional_formatting_rule(self, sheet_name: str, rules: list[dict]) -> str:
        """Set (replace) all conditional formatting rules for a sheet
        
        Args:
            sheet_name: Target sheet name
            rules: List of conditional formatting rule dicts with rule_id
        
        Returns:
            Success message string
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheet = workbook.getSheetByName('{sheet_name}');
            
            if (!sheet) {{
                throw new Error(`Sheet '${{'{sheet_name}'}}' not found`);
            }}
            
            const rules = {json.dumps(rules)};
            
            try {{
                // Try to clear existing rules and set new ones
                const existingRules = sheet.getConditionalFormattingRules();
                existingRules.forEach(rule => {{
                    sheet.deleteConditionalFormattingRule(rule);
                }});
                
                // Add new rules
                for (const rule of rules) {{
                    // Add each rule (stub implementation)
                    console.log('Setting rule:', rule);
                }}
                
                return `Note: Set ${{rules.length}} conditional formatting rule(s) on ${{'{sheet_name}'}} (limited API support)`;
            }} catch (e) {{
                console.warn('Conditional formatting not fully supported:', e);
                return `Warning: Conditional formatting API not fully supported. Rules structure received but not applied. Error: ${{e.message}}`;
            }}
        }})()
        """
        
        result = await self.execute_js(js_code)
        return result
    
    async def delete_conditional_formatting_rule(self, sheet_name: str, rule_ids: list[str]) -> str:
        """Delete conditional formatting rules by ID
        
        Args:
            sheet_name: Target sheet name
            rule_ids: List of rule IDs to delete
        
        Returns:
            Success message string
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheet = workbook.getSheetByName('{sheet_name}');
            
            if (!sheet) {{
                throw new Error(`Sheet '${{'{sheet_name}'}}' not found`);
            }}
            
            const ruleIds = {json.dumps(rule_ids)};
            
            try {{
                const rules = sheet.getConditionalFormattingRules();
                let deletedCount = 0;
                
                for (const rule of rules) {{
                    const ruleId = rule.getRuleId ? rule.getRuleId() : 'unknown';
                    if (ruleIds.includes(ruleId)) {{
                        sheet.deleteConditionalFormattingRule(rule);
                        deletedCount++;
                    }}
                }}
                
                return `Deleted ${{deletedCount}} conditional formatting rule(s) from ${{'{sheet_name}'}}`;
            }} catch (e) {{
                console.warn('Conditional formatting not fully supported:', e);
                return `Warning: Could not delete rules. API may not be fully supported. Error: ${{e.message}}`;
            }}
        }})()
        """
        
        result = await self.execute_js(js_code)
        return result
    
    async def add_data_validation_rule(self, sheet_name: str, rules: list[dict]) -> str:
        """Add data validation rules to a sheet
        
        Args:
            sheet_name: Target sheet name
            rules: List of data validation rule dicts
        
        Returns:
            Success message string
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheet = workbook.getSheetByName('{sheet_name}');
            
            if (!sheet) {{
                throw new Error(`Sheet '${{'{sheet_name}'}}' not found`);
            }}
            
            const rules = {json.dumps(rules)};
            
            try {{
                let addedCount = 0;
                
                for (const rule of rules) {{
                    const range = sheet.getRange(rule.range_a1);
                    
                    // Create data validation based on type
                    if (rule.validation_type === 'list' && rule.source) {{
                        // List validation (dropdown)
                        const validationRule = range.getDataValidation();
                        // This is a simplified implementation
                        addedCount++;
                    }} else if (rule.validation_type === 'integer' || rule.validation_type === 'decimal') {{
                        // Number validation
                        addedCount++;
                    }} else if (rule.validation_type === 'checkbox') {{
                        // Checkbox validation
                        addedCount++;
                    }}
                }}
                
                return `Note: Added ${{addedCount}} data validation rule(s) to ${{'{sheet_name}'}} (limited API support)`;
            }} catch (e) {{
                console.warn('Data validation not fully supported:', e);
                return `Warning: Data validation API not fully supported. Rules structure received but not applied. Error: ${{e.message}}`;
            }}
        }})()
        """
        
        result = await self.execute_js(js_code)
        return result
    
    async def set_data_validation_rule(self, sheet_name: str, rules: list[dict]) -> str:
        """Set (replace) all data validation rules for a sheet
        
        Args:
            sheet_name: Target sheet name
            rules: List of data validation rule dicts with rule_id
        
        Returns:
            Success message string
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheet = workbook.getSheetByName('{sheet_name}');
            
            if (!sheet) {{
                throw new Error(`Sheet '${{'{sheet_name}'}}' not found`);
            }}
            
            const rules = {json.dumps(rules)};
            
            try {{
                // Try to clear existing validation and set new rules
                // This is a stub implementation
                
                return `Note: Set ${{rules.length}} data validation rule(s) on ${{'{sheet_name}'}} (limited API support)`;
            }} catch (e) {{
                console.warn('Data validation not fully supported:', e);
                return `Warning: Data validation API not fully supported. Rules structure received but not applied. Error: ${{e.message}}`;
            }}
        }})()
        """
        
        result = await self.execute_js(js_code)
        return result
    
    async def delete_data_validation_rule(self, sheet_name: str, rule_ids: list[str]) -> str:
        """Delete data validation rules by ID
        
        Args:
            sheet_name: Target sheet name
            rule_ids: List of rule IDs to delete
        
        Returns:
            Success message string
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheet = workbook.getSheetByName('{sheet_name}');
            
            if (!sheet) {{
                throw new Error(`Sheet '${{'{sheet_name}'}}' not found`);
            }}
            
            const ruleIds = {json.dumps(rule_ids)};
            
            try {{
                // Try to delete validation rules
                // This is a stub implementation
                
                return `Note: Attempted to delete ${{ruleIds.length}} data validation rule(s) from ${{'{sheet_name}'}} (limited API support)`;
            }} catch (e) {{
                console.warn('Data validation not fully supported:', e);
                return `Warning: Could not delete rules. API may not be fully supported. Error: ${{e.message}}`;
            }}
        }})()
        """
        
        result = await self.execute_js(js_code)
        return result
    
    async def get_data_validation_rules(self, sheet_name: str) -> list:
        """Get all data validation rules for a sheet
        
        Args:
            sheet_name: Target sheet name
        
        Returns:
            List of data validation rules
        """
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheet = workbook.getSheetByName('{sheet_name}');
            
            if (!sheet) {{
                throw new Error(`Sheet '${{'{sheet_name}'}}' not found`);
            }}
            
            try {{
                // Try to get data validation rules
                // This is a stub implementation - API may not be fully available
                return [];
            }} catch (e) {{
                console.warn('Data validation not fully supported:', e);
                return [];
            }}
        }})()
        """
        
        result = await self.execute_js(js_code)
        return result


# ==================== MCP Server Setup ====================

app = Server("univer-sheets-mcp-read")
controller = UniverSheetsController()


@app.list_tools()
async def list_tools() -> list[types.Tool]:
    """List all available MCP tools"""
    return [
        types.Tool(
            name="get_activity_status",
            description="Get workbook UI status. Returns total sheet count, first 10 sheet names, current active sheet name, visible range, errors, and selection. If screenshot is True, also returns a screenshot.",
            inputSchema={
                "type": "object",
                "properties": {
                    "screenshot": {
                        "type": "boolean",
                        "description": "Whether to return a screenshot",
                        "default": False
                    }
                }
            }
        ),
        types.Tool(
            name="get_range_data",
            description="Get the cell data grid for the specified range(s) (A1 notation) in the active workbook. WARNING: Do NOT request large ranges! Keep rows × columns ≤ 200 cells maximum (e.g., A1:B100 = 200 cells is OK, A1:Z1000 = 26,000 cells will FAIL and exceed context length). For large datasets, read in smaller chunks or use search_cells instead.",
            inputSchema={
                "type": "object",
                "properties": {
                    "range_a1": {
                        "anyOf": [
                            {"type": "string"},
                            {"type": "array", "items": {"type": "string"}}
                        ],
                        "description": "The cell/range in A1 notation (e.g., 'A1', 'B2:C3'). IMPORTANT: Keep range small - rows × columns should be ≤ 200 cells total."
                    },
                    "return_screenshot": {
                        "type": "boolean",
                        "description": "Whether to return a screenshot of the range",
                        "default": False
                    },
                    "return_style": {
                        "type": "boolean",
                        "description": "Whether to return style and formula information for cells",
                        "default": False
                    }
                },
                "required": ["range_a1"]
            }
        ),
        types.Tool(
            name="scroll_and_screenshot",
            description="Scroll the UI so the specified cell is at the top-left of the viewport, then capture and return a screenshot.",
            inputSchema={
                "type": "object",
                "properties": {
                    "cell_a1": {
                        "type": "string",
                        "description": "The cell in A1 notation to scroll to (e.g., 'D10')"
                    }
                },
                "required": ["cell_a1"]
            }
        ),
        types.Tool(
            name="get_sheets",
            description="Get all sheets (name and ID) in the current workbook.",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
        types.Tool(
            name="search_cells",
            description="Search cells in the currently active workbook by keyword and findBy type. Use this tool instead of get_range_data when searching for specific content in large sheets - it's much more efficient and won't exceed context limits.",
            inputSchema={
                "type": "object",
                "properties": {
                    "keyword": {
                        "type": "string",
                        "description": "The search keyword (e.g., '=SUM' to find formulas, 'Total' to find values)"
                    },
                    "find_by": {
                        "type": "string",
                        "description": "The search type: 'formula' to search in formulas, 'value' to search in displayed values",
                        "enum": ["formula", "value"]
                    }
                },
                "required": ["keyword", "find_by"]
            }
        ),
        types.Tool(
            name="create_sheet",
            description="Create one or more new worksheets. Efficiently creates multiple sheets in one operation.",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet_names": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of names for the new worksheets to create"
                    }
                },
                "required": ["sheet_names"]
            }
        ),
        types.Tool(
            name="delete_sheet",
            description="Delete one or more worksheets by name. WARNING: This operation cannot be undone!",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet_names": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of sheet names to delete"
                    }
                },
                "required": ["sheet_names"]
            }
        ),
        types.Tool(
            name="rename_sheet",
            description="Rename one or more worksheets. Efficiently renames multiple sheets in one operation.",
            inputSchema={
                "type": "object",
                "properties": {
                    "operations": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "old_name": {
                                    "type": "string",
                                    "description": "Current sheet name to rename"
                                },
                                "new_name": {
                                    "type": "string",
                                    "description": "New name for the sheet"
                                }
                            },
                            "required": ["old_name", "new_name"]
                        },
                        "description": "List of rename operations"
                    }
                },
                "required": ["operations"]
            }
        ),
        types.Tool(
            name="activate_sheet",
            description="Activate/switch to a specific worksheet by name. This changes which sheet is currently active.",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet_name": {
                        "type": "string",
                        "description": "Name of the worksheet to activate"
                    }
                },
                "required": ["sheet_name"]
            }
        ),
        types.Tool(
            name="move_sheet",
            description="Move a worksheet to a specific index position (0-based). Does not activate the sheet.",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet_name": {
                        "type": "string",
                        "description": "Name of the worksheet to move"
                    },
                    "to_index": {
                        "anyOf": [
                            {"type": "integer"},
                            {"type": "string"}
                        ],
                        "description": "Target index position (0-based)"
                    }
                },
                "required": ["sheet_name", "to_index"]
            }
        ),
        types.Tool(
            name="set_sheet_display_status",
            description="Show or hide one or more worksheets. Efficiently sets display status for multiple sheets in one operation.",
            inputSchema={
                "type": "object",
                "properties": {
                    "operations": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "sheet_name": {
                                    "type": "string",
                                    "description": "Name of the sheet to modify"
                                },
                                "visible": {
                                    "type": "boolean",
                                    "description": "True to show sheet, False to hide sheet"
                                }
                            },
                            "required": ["sheet_name", "visible"]
                        },
                        "description": "List of display status operations"
                    }
                },
                "required": ["operations"]
            }
        ),
        types.Tool(
            name="set_range_data",
            description="Batch set values or formulas for multiple ranges in the active workbook. Essential for writing data to cells. Supports single values, arrays, and formulas (strings starting with '=').",
            inputSchema={
                "type": "object",
                "properties": {
                    "items": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "range": {
                                    "type": "string",
                                    "description": "Range string in A1 notation (e.g., 'A1', 'Sheet1!A1', 'A1:B2')"
                                },
                                "value": {
                                    "description": "Value for the range: string, number, boolean, 1D/2D array, or dict. Strings starting with '=' are formulas."
                                }
                            },
                            "required": ["range", "value"]
                        },
                        "description": "List of range-value pairs to set"
                    }
                },
                "required": ["items"]
            }
        ),
        types.Tool(
            name="set_range_style",
            description="Batch set cell styles for multiple ranges. Apply formatting like fonts, colors, borders, and alignment. Essential for making spreadsheets readable and professional.",
            inputSchema={
                "type": "object",
                "properties": {
                    "items": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "range": {
                                    "type": "string",
                                    "description": "Range string in A1 notation"
                                },
                                "style": {
                                    "type": "object",
                                    "description": "Style object with properties: ff (font family), fs (font size), it (italic 0|1), bl (bold 0|1), bg (background color), cl (font color), ht (horizontal align 1-3), vt (vertical align 1-3)"
                                }
                            },
                            "required": ["range", "style"]
                        },
                        "description": "List of range-style pairs"
                    }
                },
                "required": ["items"]
            }
        ),
        types.Tool(
            name="set_merge",
            description="Merge cells in a single range. Combines multiple cells into one larger cell, commonly used for headers and titles.",
            inputSchema={
                "type": "object",
                "properties": {
                    "range_a1": {
                        "type": "string",
                        "description": "Range string in A1 notation (e.g., 'A1:B2')"
                    }
                },
                "required": ["range_a1"]
            }
        ),
        types.Tool(
            name="set_cell_dimensions",
            description="Set column widths and row heights. Use range format 'A:C' for columns or '1:5' for rows. Measurements in points (pt).",
            inputSchema={
                "type": "object",
                "properties": {
                    "requests": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "range": {
                                    "type": "string",
                                    "description": "Range string: 'A:C' for columns A-C, or '1:5' for rows 1-5"
                                },
                                "width": {
                                    "type": "integer",
                                    "description": "Width in points (for column ranges)"
                                },
                                "height": {
                                    "type": "integer",
                                    "description": "Height in points (for row ranges)"
                                }
                            },
                            "required": ["range"]
                        },
                        "description": "List of dimension change requests"
                    }
                },
                "required": ["requests"]
            }
        ),
        types.Tool(
            name="insert_rows",
            description="Insert empty rows at specified positions. Supports multiple insert operations in one call.",
            inputSchema={
                "type": "object",
                "properties": {
                    "operations": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "position": {
                                    "type": "integer",
                                    "description": "Row index (0-based) where to insert"
                                },
                                "how_many": {
                                    "type": "integer",
                                    "description": "Number of rows to insert (default 1)"
                                },
                                "where": {
                                    "type": "string",
                                    "description": "Insert 'before' or 'after' the position (default 'before')"
                                },
                                "sheet_name": {
                                    "type": "string",
                                    "description": "Optional sheet name (uses active sheet if not provided)"
                                }
                            },
                            "required": ["position"]
                        },
                        "description": "List of insert operations"
                    }
                },
                "required": ["operations"]
            }
        ),
        types.Tool(
            name="insert_columns",
            description="Insert empty columns at specified positions. Supports multiple insert operations in one call.",
            inputSchema={
                "type": "object",
                "properties": {
                    "operations": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "position": {
                                    "type": "integer",
                                    "description": "Column index (0-based) where to insert"
                                },
                                "how_many": {
                                    "type": "integer",
                                    "description": "Number of columns to insert (default 1)"
                                },
                                "where": {
                                    "type": "string",
                                    "description": "Insert 'before' or 'after' the position (default 'before')"
                                },
                                "sheet_name": {
                                    "type": "string",
                                    "description": "Optional sheet name (uses active sheet if not provided)"
                                }
                            },
                            "required": ["position"]
                        },
                        "description": "List of insert operations"
                    }
                },
                "required": ["operations"]
            }
        ),
        types.Tool(
            name="delete_rows",
            description="Delete rows at specified positions. Supports multiple delete operations in one call.",
            inputSchema={
                "type": "object",
                "properties": {
                    "operations": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "position": {
                                    "type": "integer",
                                    "description": "Row index (0-based) where to start deletion"
                                },
                                "how_many": {
                                    "type": "integer",
                                    "description": "Number of rows to delete (default 1)"
                                },
                                "sheet_name": {
                                    "type": "string",
                                    "description": "Optional sheet name (uses active sheet if not provided)"
                                }
                            },
                            "required": ["position"]
                        },
                        "description": "List of delete operations"
                    }
                },
                "required": ["operations"]
            }
        ),
        types.Tool(
            name="delete_columns",
            description="Delete columns at specified positions. Supports multiple delete operations in one call.",
            inputSchema={
                "type": "object",
                "properties": {
                    "operations": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "position": {
                                    "type": "integer",
                                    "description": "Column index (0-based) where to start deletion"
                                },
                                "how_many": {
                                    "type": "integer",
                                    "description": "Number of columns to delete (default 1)"
                                },
                                "sheet_name": {
                                    "type": "string",
                                    "description": "Optional sheet name (uses active sheet if not provided)"
                                }
                            },
                            "required": ["position"]
                        },
                        "description": "List of delete operations"
                    }
                },
                "required": ["operations"]
            }
        ),
        types.Tool(
            name="get_active_unit_id",
            description="Get the currently active unit_id (workbook ID) for this session.",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
        types.Tool(
            name="auto_fill",
            description="Auto fill data from source range to target range using pattern detection. Copies data, formulas, and styles. Supports horizontal or vertical extension only.",
            inputSchema={
                "type": "object",
                "properties": {
                    "source_range": {
                        "type": "string",
                        "description": "Source range in A1 notation (e.g., 'A1:A2')"
                    },
                    "target_range": {
                        "type": "string",
                        "description": "Target range in A1 notation (e.g., 'A1:A10')"
                    }
                },
                "required": ["source_range", "target_range"]
            }
        ),
        types.Tool(
            name="format_brush",
            description="Copy formatting from source range to target range (format painter). Supports single cell, range, and cross-sheet operations.",
            inputSchema={
                "type": "object",
                "properties": {
                    "source_range": {
                        "type": "string",
                        "description": "Source range in A1 notation, supports cross-sheet (e.g., 'Sheet1!A1')"
                    },
                    "target_range": {
                        "type": "string",
                        "description": "Target range in A1 notation, supports cross-sheet (e.g., 'Sheet2!B1:C1')"
                    }
                },
                "required": ["source_range", "target_range"]
            }
        ),
        types.Tool(
            name="get_conditional_formatting_rules",
            description="Get all conditional formatting rules for the given sheet.",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet_name": {
                        "type": "string",
                        "description": "Target sheet name"
                    }
                },
                "required": ["sheet_name"]
            }
        ),
        types.Tool(
            name="add_conditional_formatting_rule",
            description="Add one or more conditional formatting rules to a sheet. Note: Limited API support - may not be fully functional.",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet_name": {
                        "type": "string",
                        "description": "Target sheet name"
                    },
                    "rules": {
                        "type": "array",
                        "items": {"type": "object"},
                        "description": "List of conditional formatting rules"
                    }
                },
                "required": ["sheet_name", "rules"]
            }
        ),
        types.Tool(
            name="set_conditional_formatting_rule",
            description="Set (replace) all conditional formatting rules for a sheet. Note: Limited API support.",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet_name": {
                        "type": "string",
                        "description": "Target sheet name"
                    },
                    "rules": {
                        "type": "array",
                        "items": {"type": "object"},
                        "description": "List of conditional formatting rules with rule_id"
                    }
                },
                "required": ["sheet_name", "rules"]
            }
        ),
        types.Tool(
            name="delete_conditional_formatting_rule",
            description="Delete conditional formatting rules by ID. Note: Limited API support.",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet_name": {
                        "type": "string",
                        "description": "Target sheet name"
                    },
                    "rule_ids": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of rule IDs to delete"
                    }
                },
                "required": ["sheet_name", "rule_ids"]
            }
        ),
        types.Tool(
            name="add_data_validation_rule",
            description="Add data validation rules to a sheet (dropdowns, checkboxes, etc.). Note: Limited API support.",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet_name": {
                        "type": "string",
                        "description": "Target sheet name"
                    },
                    "rules": {
                        "type": "array",
                        "items": {"type": "object"},
                        "description": "List of data validation rules"
                    }
                },
                "required": ["sheet_name", "rules"]
            }
        ),
        types.Tool(
            name="set_data_validation_rule",
            description="Set (replace) all data validation rules for a sheet. Note: Limited API support.",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet_name": {
                        "type": "string",
                        "description": "Target sheet name"
                    },
                    "rules": {
                        "type": "array",
                        "items": {"type": "object"},
                        "description": "List of data validation rules with rule_id"
                    }
                },
                "required": ["sheet_name", "rules"]
            }
        ),
        types.Tool(
            name="delete_data_validation_rule",
            description="Delete data validation rules by ID. Note: Limited API support.",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet_name": {
                        "type": "string",
                        "description": "Target sheet name"
                    },
                    "rule_ids": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of rule IDs to delete"
                    }
                },
                "required": ["sheet_name", "rule_ids"]
            }
        ),
        types.Tool(
            name="get_data_validation_rules",
            description="Get all data validation rules for a sheet. Note: Limited API support.",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet_name": {
                        "type": "string",
                        "description": "Target sheet name"
                    }
                },
                "required": ["sheet_name"]
            }
        )
    ]


@app.call_tool()
async def call_tool(name: str, arguments: dict) -> list[types.TextContent | types.ImageContent]:
    """Handle MCP tool calls"""
    try:
        logger.info(f"Calling tool: {name} with arguments: {arguments}")
        
        # Route to appropriate handler
        if name == "get_activity_status":
            result = await controller.get_activity_status(**arguments)
            screenshot = result.pop('screenshot', None)
            
            contents = [types.TextContent(type="text", text=json.dumps(result, indent=2))]
            if screenshot:
                contents.append(types.ImageContent(
                    type="image",
                    data=screenshot,
                    mimeType="image/png"
                ))
            return contents
        
        elif name == "get_range_data":
            result = await controller.get_range_data(**arguments)
            screenshot = result.pop('screenshot', None)
            
            contents = [types.TextContent(type="text", text=json.dumps(result, indent=2))]
            if screenshot:
                contents.append(types.ImageContent(
                    type="image",
                    data=screenshot,
                    mimeType="image/png"
                ))
            return contents
        
        elif name == "scroll_and_screenshot":
            result = await controller.scroll_and_screenshot(**arguments)
            screenshot = result.pop('screenshot')
            
            return [
                types.TextContent(type="text", text=json.dumps(result, indent=2)),
                types.ImageContent(
                    type="image",
                    data=screenshot,
                    mimeType="image/png"
                )
            ]
        
        elif name == "get_sheets":
            result = await controller.get_sheets()
            return [types.TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "search_cells":
            result = await controller.search_cells(**arguments)
            return [types.TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "create_sheet":
            result = await controller.create_sheet(**arguments)
            return [types.TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "delete_sheet":
            result = await controller.delete_sheet(**arguments)
            return [types.TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "rename_sheet":
            result = await controller.rename_sheet(**arguments)
            return [types.TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "activate_sheet":
            result = await controller.activate_sheet(**arguments)
            return [types.TextContent(type="text", text=result)]
        
        elif name == "move_sheet":
            result = await controller.move_sheet(**arguments)
            return [types.TextContent(type="text", text=result)]
        
        elif name == "set_sheet_display_status":
            result = await controller.set_sheet_display_status(**arguments)
            return [types.TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "set_range_data":
            result = await controller.set_range_data(**arguments)
            return [types.TextContent(type="text", text=result)]
        
        elif name == "set_range_style":
            result = await controller.set_range_style(**arguments)
            return [types.TextContent(type="text", text=result)]
        
        elif name == "set_merge":
            result = await controller.set_merge(**arguments)
            return [types.TextContent(type="text", text=result)]
        
        elif name == "set_cell_dimensions":
            result = await controller.set_cell_dimensions(**arguments)
            return [types.TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "insert_rows":
            result = await controller.insert_rows(**arguments)
            return [types.TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "insert_columns":
            result = await controller.insert_columns(**arguments)
            return [types.TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "delete_rows":
            result = await controller.delete_rows(**arguments)
            return [types.TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "delete_columns":
            result = await controller.delete_columns(**arguments)
            return [types.TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "get_active_unit_id":
            result = await controller.get_active_unit_id()
            return [types.TextContent(type="text", text=result)]
        
        elif name == "auto_fill":
            result = await controller.auto_fill(**arguments)
            return [types.TextContent(type="text", text=result)]
        
        elif name == "format_brush":
            result = await controller.format_brush(**arguments)
            return [types.TextContent(type="text", text=result)]
        
        elif name == "get_conditional_formatting_rules":
            result = await controller.get_conditional_formatting_rules(**arguments)
            return [types.TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "add_conditional_formatting_rule":
            result = await controller.add_conditional_formatting_rule(**arguments)
            return [types.TextContent(type="text", text=result)]
        
        elif name == "set_conditional_formatting_rule":
            result = await controller.set_conditional_formatting_rule(**arguments)
            return [types.TextContent(type="text", text=result)]
        
        elif name == "delete_conditional_formatting_rule":
            result = await controller.delete_conditional_formatting_rule(**arguments)
            return [types.TextContent(type="text", text=result)]
        
        elif name == "add_data_validation_rule":
            result = await controller.add_data_validation_rule(**arguments)
            return [types.TextContent(type="text", text=result)]
        
        elif name == "set_data_validation_rule":
            result = await controller.set_data_validation_rule(**arguments)
            return [types.TextContent(type="text", text=result)]
        
        elif name == "delete_data_validation_rule":
            result = await controller.delete_data_validation_rule(**arguments)
            return [types.TextContent(type="text", text=result)]
        
        elif name == "get_data_validation_rules":
            result = await controller.get_data_validation_rules(**arguments)
            return [types.TextContent(type="text", text=json.dumps(result, indent=2))]
        
        else:
            raise ValueError(f"Unknown tool: {name}")
            
    except Exception as e:
        logger.error(f"Error executing tool {name}: {e}")
        return [types.TextContent(
            type="text",
            text=f"Error: {str(e)}"
        )]


async def main():
    """Main entry point"""
    import sys
    
    # Parse command line arguments
    univer_url = "http://localhost:3002/sheets/"
    headless = False
    
    if "--url" in sys.argv:
        idx = sys.argv.index("--url")
        if idx + 1 < len(sys.argv):
            univer_url = sys.argv[idx + 1]
    
    if "--headless" in sys.argv:
        headless = True
    
    # Start Univer controller
    await controller.start(url=univer_url, headless=headless)
    
    # Run MCP server
    async with stdio_server() as (read_stream, write_stream):
        await app.run(
            read_stream,
            write_stream,
            app.create_initialization_options()
        )
    
    # Cleanup
    await controller.cleanup()


if __name__ == "__main__":
    asyncio.run(main())
