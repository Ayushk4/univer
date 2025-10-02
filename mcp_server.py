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

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class UniverSheetsController:
    """Controller for reading data from Univer Sheets via Playwright"""
    
    def __init__(self):
        self.page: Optional[Page] = None
        self.browser: Optional[Browser] = None
        self.playwright = None
        self.univer_url = "http://localhost:3002/sheets/"
        
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
        js_code = f"""
        (async () => {{
            const workbook = window.univerAPI.getActiveWorkbook();
            const sheet = workbook.getActiveSheet();
            
            // Simple search implementation
            const maxRows = sheet.getMaxRows();
            const maxCols = sheet.getMaxColumns();
            const results = [];
            
            for (let row = 0; row < Math.min(maxRows, 100); row++) {{
                for (let col = 0; col < Math.min(maxCols, 26); col++) {{
                    const cell = sheet.getRange(row, col);
                    const cellData = cell.getCellData();
                    
                    let searchValue = '';
                    if ('{find_by}' === 'formula' && cellData && cellData.f) {{
                        searchValue = cellData.f;
                    }} else if ('{find_by}' === 'value' && cellData && cellData.v != null) {{
                        searchValue = String(cellData.v);
                    }}
                    
                    if (searchValue.includes('{keyword}')) {{
                        results.push({{
                            range: cell.getA1Notation(),
                            value: cellData ? cellData.v : null,
                            formula: cellData ? cellData.f : null
                        }});
                    }}
                    
                    if (results.length >= 50) break;
                }}
                if (results.length >= 50) break;
            }}
            
            return {{
                total: results.length,
                results: results
            }};
        }})()
        """
        
        return await self.execute_js(js_code)


# ==================== MCP Server Setup ====================

app = Server("univer-sheets-mcp-read")
controller = UniverSheetsController()


@app.list_tools()
async def list_tools() -> list[types.Tool]:
    """List all available MCP tools (read-only)"""
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
