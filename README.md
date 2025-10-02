# Univer Sheets Tool Calling & MCP Server

Packages for Tool Calling & [Model Context Protocol (MCP)](https://modelcontextprotocol.io) server that provides read-only access to a self-deployed [Univer Sheets](https://github.com/dream-num/univer) instance. This server exposes 5 read tools for viewing and analyzing spreadsheets through natural language using AI assistants like Claude.

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

### Step 1: Setup MCP Server

```bash
# uv venv --python 3.12 --seed && source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

### Step 2: Start Univer Sheets

```bash
git clone https://github.com/dream-num/univer univer/
cd univer
pnpm install

pnpm dev
```

✅ **Verify**: Open http://localhost:3002/sheets/ in your browser. You should see Univer Sheets.

### Step 3: Test Tool calling & MCP Server

```bash
python test_mcp.py --quick
python test_mcp.py

python pydantic_agent.py # A simple agent call - make sure you have `OPEN_ROUTER_KEYS` set.
cd llm-demo && python test.py # Test the simple demo server.
```

To Launch MCP Server:

```bash
python mcp_server.py --url http://localhost:3002/sheets/
# Or run in headless mode (no visible browser)
python mcp_server.py --url http://localhost:3002/sheets/ --headless
```

**Important:** Make sure to use the `/sheets/` path, not just the root URL.

To launch demo server with chat on right side:
```bash
python app.py
```

## 🛠️ Available Tools

| Tool | Description | Example |
|------|-------------|---------|
| `get_activity_status` | Get workbook status, active sheet, selection | Get current sheet info |
| `get_range_data` | Read cell values and formulas | Read A1:B10 |
| `get_sheets` | List all sheets in workbook | Get all sheet names |
| `search_cells` | Search by value or formula | Find cells with "=SUM" |
| `scroll_and_screenshot` | Scroll to cell and capture | Scroll to Z100 and screenshot |
