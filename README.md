# Univer Sheets MCP Server (Read-Only)

A [Model Context Protocol (MCP)](https://modelcontextprotocol.io) server that provides read-only access to a self-deployed [Univer Sheets](https://github.com/dream-num/univer) instance. This server exposes 5 read tools for viewing and analyzing spreadsheets through natural language using AI assistants like Claude.

## 🌟 Features

- **5 Read-Only MCP Tools** for viewing spreadsheet data
- **Visual Feedback** with screenshot support
- **Efficient Search** for large datasets
- **Real-time Access** via browser automation with Playwright
- **Python-Native** implementation with async support

## 🏗️ Architecture

```
┌─────────────────────────────────────────────────┐
│  MCP Client (Claude Desktop, etc.)              │
│  Sends natural language requests                │
└────────────────┬────────────────────────────────┘
                 │ MCP Protocol
                 ▼
┌─────────────────────────────────────────────────┐
│  MCP Server (Python)                            │
│  ┌───────────────────────────────────────────┐ │
│  │  Read-Only Tool Handlers                  │ │
│  └───────────────┬───────────────────────────┘ │
│  ┌───────────────▼───────────────────────────┐ │
│  │  UniverSheetsController                   │ │
│  │  - Playwright browser automation          │ │
│  │  - Facade API interaction                 │ │
│  │  - Screenshot capture                     │ │
│  └───────────────┬───────────────────────────┘ │
└──────────────────┼─────────────────────────────┘
                   │ Browser Control
                   ▼
┌─────────────────────────────────────────────────┐
│  Browser (Chromium)                             │
│  http://localhost:3002/sheets/                  │
│  ┌───────────────────────────────────────────┐ │
│  │  Univer Sheets Instance                   │ │
│  │  window.univerAPI (Facade API)            │ │
│  └───────────────────────────────────────────┘ │
└─────────────────────────────────────────────────┘
```

## 📋 Prerequisites

1. **Node.js** >= 18.17.0
2. **Python** >= 3.10
3. **pnpm** package manager
4. **Univer Sheets** instance running locally

## 🚀 Quick Start (5 minutes)

### Step 1: Start Univer Sheets

```bash
pnpm install

pnpm dev
```

✅ **Verify**: Open http://localhost:3002/sheets/ in your browser. You should see Univer Sheets.

### Step 2: Setup MCP Server

```bash
cd univer-mcp-server
# uv venv --python 3.12 --seed && source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

### Step 3: Run MCP Server

```bash
python mcp_server.py --url http://localhost:3002/sheets/
# Or run in headless mode (no visible browser)
python mcp_server.py --url http://localhost:3002/sheets/ --headless
```

**Important:** Make sure to use the `/sheets/` path, not just the root URL.

### Step 4: Configure Claude Desktop (Optional)

Add to your Claude Desktop configuration:
- macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
- Windows: `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "univer-sheets": {
      "command": "python",
      "args": [
        "/absolute/path/to/univer-mcp-server/mcp_server.py"
      ]
    }
  }
}
```

Restart Claude Desktop and you'll see the Univer tools available!

## ✅ Testing the MCP Server

Before integrating with Claude Desktop, verify the MCP server works correctly using the test scripts:

### Quick Test (30 seconds)

Run a fast smoke test to verify basic functionality:

```bash
cd univer-mcp-server
python test_quick.py
```

**Expected output:**
```
🚀 Starting quick test...

1. Connecting to Univer...
   ✅ Connected!

2. Getting workbook status...
   ✅ Active sheet: Sheet1
   ✅ Total sheets: 1

3. Reading range A1:C3...
   ✅ Retrieved 3 rows

4. Getting all sheets...
   ✅ Found 1 sheet(s):
      - Sheet1

✨ All tests passed! MCP server is working.
```

### Comprehensive Test (2 minutes)

Run all test scenarios to verify every tool:

```bash
python test_mcp.py
```

This tests:
- Activity status (with and without screenshots)
- Range data reading (single and multiple ranges)
- Style information retrieval
- Sheet listing
- Cell searching (by value and formula)
- Scrolling and screenshots
- Large range handling (200 cells)

**Options:**
```bash
# Test with custom URL
python test_mcp.py --url http://localhost:3002/sheets/

# Run in headless mode (no visible browser)
python test_mcp.py --headless
```

### Programmatic Usage

You can also use the controller directly in your Python code:

```python
import asyncio
from mcp_server import UniverSheetsController

async def main():
    controller = UniverSheetsController()
    await controller.start("http://localhost:3002/sheets/", headless=False)
    
    # Get workbook status
    status = await controller.get_activity_status()
    print(f"Active sheet: {status['activeSheetName']}")
    
    # Read data
    data = await controller.get_range_data("A1:B10")
    print(f"Values: {data['values']}")
    
    # Search cells
    results = await controller.search_cells("Total", "value")
    print(f"Found {results['total']} matches")
    
    await controller.cleanup()

asyncio.run(main())
```

## 🛠️ Available Tools

| Tool | Description | Example |
|------|-------------|---------|
| `get_activity_status` | Get workbook status, active sheet, selection | Get current sheet info |
| `get_range_data` | Read cell values and formulas | Read A1:B10 |
| `get_sheets` | List all sheets in workbook | Get all sheet names |
| `search_cells` | Search by value or formula | Find cells with "=SUM" |
| `scroll_and_screenshot` | Scroll to cell and capture | Scroll to Z100 and screenshot |

## 💡 Usage Examples

### Example 1: Analyze Data

**Prompt to Claude:**
```
Look at the data in Sheet1, range A1:D10, and tell me:
1. What are the column headers?
2. How many rows of data are there?
3. Show me a screenshot
```

**What happens:**
1. `get_range_data` reads A1:D10
2. Claude analyzes the data
3. `scroll_and_screenshot` captures the view

### Example 2: Find Specific Content

**Prompt to Claude:**
```
Find all cells that contain "Total" and tell me their locations
```

**What happens:**
1. `search_cells` searches for "Total"
2. Claude reports the found cells and their positions

### Example 3: Get Overview

**Prompt to Claude:**
```
Give me an overview of this workbook - how many sheets, what are they called, and what's currently selected?
```

**What happens:**
1. `get_activity_status` gets workbook info
2. `get_sheets` lists all sheets
3. Claude summarizes the information

## 📊 Data Access Limits

### The Rule: Maximum 200 Cells Per Request

When using `get_range_data`, **NEVER** request ranges larger than **200 cells** (rows × columns ≤ 200).

### ✅ Good Examples (Safe)

| Range | Calculation | Total Cells | Status |
|-------|-------------|-------------|--------|
| `A1:B100` | 2 × 100 | 200 | ✅ OK |
| `A1:J20` | 10 × 20 | 200 | ✅ OK |
| `A1:C50` | 3 × 50 | 150 | ✅ OK |

### ❌ Bad Examples (Will Fail)

| Range | Calculation | Total Cells | Status |
|-------|-------------|-------------|--------|
| `A1:Z1000` | 26 × 1000 | 26,000 | ❌ TOO MUCH! |
| `A1:AA100` | 27 × 100 | 2,700 | ❌ TOO MUCH! |

### Alternative Strategies

**1. Read in Chunks**
```python
# Instead of A1:Z1000, do:
get_range_data("A1:C200")  # First chunk
get_range_data("D1:F200")  # Second chunk
```

**2. Use Search**
```python
search_cells("Total", "value")  # Find cells containing "Total"
search_cells("=SUM", "formula")  # Find all SUM formulas
```

## 🎯 Univer Demo URLs

When you run `pnpm dev` in the Univer repository, it starts at **http://localhost:3002** with multiple demos:

### Primary Target (Recommended)
**URL:** `http://localhost:3002/sheets/`
- Full-featured spreadsheet
- All plugins enabled
- **Use this for MCP**

### Alternative (Without Workers)
**URL:** `http://localhost:3002/sheets-no-worker/`
- Same features but no web workers
- Useful for debugging

### Other Demos
- `/docs/` - Document editor
- `/slides/` - Presentation editor
- `/sheets-multi/` - Multiple instances
- `/uni/` - Unified mode

**For MCP:** Always use `/sheets/` unless you have a specific reason otherwise.

## 🔧 Advanced Configuration

### Custom Univer URL

```bash
# Different port
python mcp_server.py --url http://localhost:8080/sheets/

# Different demo
python mcp_server.py --url http://localhost:3002/sheets-no-worker/
```

### Headless Mode

For server deployments without GUI:

```bash
python mcp_server.py --headless
```

### Debugging

Enable detailed logging in `mcp_server.py`:

```python
logging.basicConfig(level=logging.DEBUG)
```

## 💻 Programmatic Usage

You can use the `UniverSheetsController` directly in your Python code:

```python
import asyncio
from mcp_server import UniverSheetsController

async def main():
    controller = UniverSheetsController()
    
    # Connect to Univer
    await controller.start("http://localhost:3002/sheets/", headless=False)
    
    # Get workbook status
    status = await controller.get_activity_status(screenshot=False)
    print(status)
    
    # Read data
    data = await controller.get_range_data("A1:B10")
    print(data)
    
    # Get all sheets
    sheets = await controller.get_sheets()
    print(sheets)
    
    # Search for cells
    results = await controller.search_cells("Total", "value")
    print(results)
    
    # Screenshot
    screenshot_data = await controller.scroll_and_screenshot("A1")
    print(f"Screenshot: {len(screenshot_data['screenshot'])} bytes")
    
    # Cleanup
    await controller.cleanup()

if __name__ == "__main__":
    asyncio.run(main())
```

## 🐛 Troubleshooting

### "Browser not started" error
```bash
playwright install chromium
```

### "univerAPI is undefined" error
- Ensure Univer is running at http://localhost:3002
- Use the `/sheets/` path (not just the root)
- Wait for Univer to fully load
- Check browser console for errors

### Screenshot not working
```bash
playwright install-deps
```

### Port already in use
```bash
# Check what's using the port
lsof -i :3002

# Start Univer on different port
pnpm dev -- --port 3003

# Update MCP server URL
python mcp_server.py --url http://localhost:3003/sheets/
```

### MCP server not showing in Claude
1. Check config file path is correct
2. Use absolute paths (not relative)
3. Restart Claude Desktop completely
4. Check Claude Desktop logs for errors

## 🎓 Demo Application

Included in `llm-demo/` is a simple web demo showing how to use the MCP tools with an LLM:

```bash
# Set your OpenRouter API key
export OPENROUTER_API_KEY='your-key-here'

# Run the demo server
cd llm-demo
python app.py

# Open in browser
open http://localhost:3003
```

This provides:
- Univer spreadsheet on the left
- Chat interface on the right
- LLM can read and analyze the spreadsheet data

## 🤖 PydanticAI Agent

For direct agentic interaction with spreadsheets, use `pydantic_agent.py`:

```bash
# Install dependencies
pip install pydantic-ai[openai]

# Set your OpenAI API key
export OPENAI_API_KEY='your-key-here'

# Run with a custom prompt
python pydantic_agent.py "What sheets are in the workbook?"

# Run with default prompt
python pydantic_agent.py

# Custom URL and headless mode
python pydantic_agent.py "Get data from A1 to C5" --url http://localhost:3002/sheets/ --headless
```

This provides:
- Natural language queries processed by GPT-4
- Automatic tool selection and execution
- Conversation trace showing all tool calls
- Single-turn agentic interaction with spreadsheets

Example queries:
- "What sheets are available?"
- "Get the data from cells A1 to C5"
- "Search for cells containing 'Total'"
- "What is the active sheet name?"

## 📚 API Reference

### UniverSheetsController Methods

```python
# Start browser and connect
await controller.start(url="http://localhost:3002/sheets/", headless=False)

# Get workbook status
status = await controller.get_activity_status(screenshot=False)

# Read range data
data = await controller.get_range_data("A1:B10", return_screenshot=False, return_style=False)

# Get all sheets
sheets = await controller.get_sheets()

# Search cells
results = await controller.search_cells("keyword", "value")  # or "formula"

# Scroll and screenshot
screenshot_data = await controller.scroll_and_screenshot("A1")

# Execute raw JavaScript
result = await controller.execute_js("window.univerAPI.getActiveWorkbook().getName()")

# Take screenshot
screenshot_b64 = await controller.take_screenshot()

# Cleanup
await controller.cleanup()
```

## 🤝 Contributing

This is a simplified read-only version. For a full-featured MCP server with write capabilities, see the [official Univer MCP](https://github.com/dream-num/univer-mcp).

## 📖 Related Documentation

- [Univer Documentation](https://univer.ai)
- [Univer Facade API](https://univer.ai/guides/sheets/getting-started/facade)
- [Model Context Protocol](https://modelcontextprotocol.io)
- [Playwright Python](https://playwright.dev/python/)

## 📄 License

This project follows the same license as Univer (Apache-2.0).

## 🙏 Acknowledgments

- [Univer Team](https://github.com/dream-num/univer) for the amazing spreadsheet framework
- [Anthropic](https://anthropic.com) for the Model Context Protocol
- [Playwright Team](https://playwright.dev) for browser automation tools

## 🔗 Links

- [Univer Repository](https://github.com/dream-num/univer)
- [Univer MCP (Official)](https://github.com/dream-num/univer-mcp)
- [MCP Specification](https://spec.modelcontextprotocol.io)

---

**Built with ❤️ for the Univer community**
