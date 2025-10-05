"""
Pydantic AI Agent for Univer Sheets

A minimal agent that uses PydanticAI to process natural language queries
about spreadsheet data using the UniverSheetsController tools.

Usage:
    python pydantic_agent.py "What sheets are in the workbook?"
    python pydantic_agent.py "Get the data from cells A1 to C5"
    python pydantic_agent.py  # Uses default prompt
    
Can also be imported as an async generator:
    import asyncio
    from pydantic_agent import create_agent, register_tools, run_query
    
    async def run():
        agent, controller = create_agent()
        register_tools(agent, controller)
        await controller.start(url='http://localhost:3002/sheets/')
        
        # Stream results - get each step as it happens
        async for step in run_query(agent, controller, "What sheets are available?"):
            if hasattr(step, 'output'):
                # This is the final result
                print(f"Final: {step.output}\n\n")
            else:
                # This is an intermediate step
                print(f"Step: {step}\n\n")
        
        await controller.cleanup()
    
    asyncio.run(run())
"""

import asyncio
import argparse
import os
from pydantic_ai import Agent, RunContext, ModelSettings
from pydantic_ai.models.openai import OpenAIChatModel
from pydantic_ai.providers.openrouter import OpenRouterProvider
from mcp_server import UniverSheetsController


SYSTEM_PROMPT = """You are an assistant that helps users query spreadsheet data.
You have access to tools to read data from Univer Sheets.
Always use the tools to answer questions about the spreadsheet.
Keep your responses concise and focused on the data."""


def create_agent():
    """Create and return the agent with OpenAI model."""
    model = OpenAIChatModel(
        # 'openai/gpt-oss-120b',
        'anthropic/claude-sonnet-4.5',
        provider=OpenRouterProvider(api_key=os.environ["OPEN_ROUTER_KEYS"]),
    )
    agent = Agent(
        model=model, system_prompt=SYSTEM_PROMPT,
        # model_settings=ModelSettings(extra_body={"provider": {"only": ["cerebras"], "allow_fallbacks": False}}),
    )
    controller = UniverSheetsController()
    return agent, controller


def register_tools(agent, controller):
    """Register all tools with the agent."""
    
    @agent.tool
    async def get_activity_status(_ctx: RunContext) -> dict:
        """Get current workbook status including active sheet name and selection info."""
        result = await controller.get_activity_status(screenshot=False)
        return result


    @agent.tool
    async def get_sheets(_ctx: RunContext) -> list:
        """Get all sheets in the workbook with their names and metadata."""
        result = await controller.get_sheets()
        return result


    @agent.tool
    async def get_range_data(_ctx: RunContext, range_a1: str) -> dict:
        """Get cell data for a specified range in A1 notation (e.g., 'A1:C5').
        
        WARNING: Keep ranges small - maximum 200 cells (rows × columns ≤ 200).
        """
        result = await controller.get_range_data(
            range_a1=range_a1,
            return_screenshot=False,
            return_style=False
        )
        return result


    @agent.tool
    async def search_cells(_ctx: RunContext, keyword: str, find_by: str) -> dict:
        """Search for cells containing a keyword.
        
        Args:
            keyword: The text to search for
            find_by: Either 'formula' to search formulas or 'value' to search displayed values
        
        Use this for large sheets instead of get_range_data.
        """
        result = await controller.search_cells(keyword=keyword, find_by=find_by)
        return result


    @agent.tool
    async def create_sheet(_ctx: RunContext, sheet_names: list[str]) -> dict:
        """Create one or more new worksheets.
        
        Args:
            sheet_names: List of names for the new sheets to create
        
        Returns a dict with status, message, and details of each operation.
        """
        result = await controller.create_sheet(sheet_names=sheet_names)
        return result


    @agent.tool
    async def delete_sheet(_ctx: RunContext, sheet_names: list[str]) -> dict:
        """Delete one or more worksheets by name. WARNING: Cannot be undone!
        
        Args:
            sheet_names: List of sheet names to delete
        
        Returns a dict with status and message.
        """
        result = await controller.delete_sheet(sheet_names=sheet_names)
        return result


    @agent.tool
    async def rename_sheet(_ctx: RunContext, operations: list[dict]) -> dict:
        """Rename one or more worksheets.
        
        Args:
            operations: List of dicts, each with 'old_name' and 'new_name' keys
        
        Example: [{"old_name": "Sheet1", "new_name": "Sales Data"}]
        """
        result = await controller.rename_sheet(operations=operations)
        return result


    @agent.tool
    async def activate_sheet(_ctx: RunContext, sheet_name: str) -> str:
        """Activate/switch to a specific worksheet by name.
        
        Args:
            sheet_name: Name of the sheet to activate
        
        Returns a success message.
        """
        result = await controller.activate_sheet(sheet_name=sheet_name)
        return result


    @agent.tool
    async def move_sheet(_ctx: RunContext, sheet_name: str, to_index: int) -> str:
        """Move a worksheet to a specific position (0-based index).
        
        Args:
            sheet_name: Name of the sheet to move
            to_index: Target index position (0-based)
        
        Returns a success message.
        """
        result = await controller.move_sheet(sheet_name=sheet_name, to_index=to_index)
        return result


    @agent.tool
    async def set_sheet_display_status(_ctx: RunContext, operations: list[dict]) -> dict:
        """Show or hide one or more worksheets.
        
        Args:
            operations: List of dicts, each with 'sheet_name' and 'visible' (bool) keys
        
        Example: [{"sheet_name": "Sheet1", "visible": False}]
        """
        result = await controller.set_sheet_display_status(operations=operations)
        return result


    @agent.tool
    async def set_range_data(_ctx: RunContext, items: list[dict]) -> str:
        """Set cell values and formulas for multiple ranges.
        
        Args:
            items: List of dicts with 'range' and 'value'
                  - range: A1 notation (e.g., 'A1', 'Sheet1!A1', 'A1:B2')
                  - value: string, number, boolean, or array
                          Strings starting with '=' are treated as formulas
        
        Examples:
            [{"range": "A1", "value": "Hello"}]
            [{"range": "B1", "value": "=A1"}]
            [{"range": "A1:B2", "value": [["A1", "B1"], ["A2", "B2"]]}]
        """
        result = await controller.set_range_data(items=items)
        return result


    @agent.tool
    async def set_range_style(_ctx: RunContext, items: list[dict]) -> str:
        """Set cell styles for multiple ranges (fonts, colors, alignment, etc.).
        
        Args:
            items: List of dicts with 'range' and 'style'
                  - range: A1 notation
                  - style: Style object with properties like:
                    * ff: font family (e.g., "Arial")
                    * fs: font size in pt (e.g., 14)
                    * bl: bold (0 or 1)
                    * it: italic (0 or 1)
                    * bg: background color (e.g., {"rgb": "#FF0000"})
                    * cl: font color (e.g., {"rgb": "#000000"})
                    * ht: horizontal alignment (1=left, 2=center, 3=right)
                    * vt: vertical alignment (1=top, 2=middle, 3=bottom)
        
        Example: [{"range": "A1", "style": {"bl": 1, "fs": 16}}]
        """
        result = await controller.set_range_style(items=items)
        return result


    @agent.tool
    async def set_merge(_ctx: RunContext, range_a1: str) -> str:
        """Merge cells in a single range.
        
        Args:
            range_a1: Range in A1 notation (e.g., 'A1:B2')
        
        Merges multiple cells into one, commonly used for headers.
        """
        result = await controller.set_merge(range_a1=range_a1)
        return result


    @agent.tool
    async def set_cell_dimensions(_ctx: RunContext, requests: list[dict]) -> dict:
        """Set column widths and row heights.
        
        Args:
            requests: List of dicts with:
                     - range: 'A:C' for columns or '1:5' for rows
                     - width: Width in points (for columns)
                     - height: Height in points (for rows)
        
        Examples:
            [{"range": "A:C", "width": 120}]  # Set columns A-C to 120pt
            [{"range": "1:5", "height": 25}]   # Set rows 1-5 to 25pt
        """
        result = await controller.set_cell_dimensions(requests=requests)
        return result


    @agent.tool
    async def insert_rows(_ctx: RunContext, operations: list[dict]) -> dict:
        """Insert empty rows at specified positions.
        
        Args:
            operations: List of dicts with:
                       - position: Row index (0-based) where to insert
                       - how_many: Number of rows to insert (default 1)
                       - where: 'before' or 'after' the position (default 'before')
                       - sheet_name: Optional sheet name (uses active if not provided)
        
        Example: [{"position": 5, "how_many": 3, "where": "before"}]
        """
        result = await controller.insert_rows(operations=operations)
        return result


    @agent.tool
    async def insert_columns(_ctx: RunContext, operations: list[dict]) -> dict:
        """Insert empty columns at specified positions.
        
        Args:
            operations: List of dicts with:
                       - position: Column index (0-based) where to insert
                       - how_many: Number of columns to insert (default 1)
                       - where: 'before' or 'after' the position (default 'before')
                       - sheet_name: Optional sheet name (uses active if not provided)
        
        Example: [{"position": 2, "how_many": 2, "where": "after"}]
        """
        result = await controller.insert_columns(operations=operations)
        return result


    @agent.tool
    async def delete_rows(_ctx: RunContext, operations: list[dict]) -> dict:
        """Delete rows at specified positions.
        
        Args:
            operations: List of dicts with:
                       - position: Row index (0-based) where to start deletion
                       - how_many: Number of rows to delete (default 1)
                       - sheet_name: Optional sheet name (uses active if not provided)
        
        Example: [{"position": 10, "how_many": 5}]
        """
        result = await controller.delete_rows(operations=operations)
        return result


    @agent.tool
    async def delete_columns(_ctx: RunContext, operations: list[dict]) -> dict:
        """Delete columns at specified positions.
        
        Args:
            operations: List of dicts with:
                       - position: Column index (0-based) where to start deletion
                       - how_many: Number of columns to delete (default 1)
                       - sheet_name: Optional sheet name (uses active if not provided)
        
        Example: [{"position": 3, "how_many": 2}]
        """
        result = await controller.delete_columns(operations=operations)
        return result


    @agent.tool
    async def get_active_unit_id(_ctx: RunContext) -> str:
        """Get the currently active unit_id (workbook ID) for this session.
        
        Returns:
            The workbook ID string
        """
        result = await controller.get_active_unit_id()
        return result


    @agent.tool
    async def auto_fill(_ctx: RunContext, source_range: str, target_range: str) -> str:
        """Auto fill data from source range to target range using pattern detection.
        
        Copies data, formulas, and styles. Supports horizontal or vertical extension only.
        The target range must be aligned with the source (same row or column direction).
        
        Args:
            source_range: Source range in A1 notation (e.g., 'A1:A2')
            target_range: Target range in A1 notation (e.g., 'A1:A10')
        
        Example: auto_fill('A1:A2', 'A1:A10') fills down from pattern
        """
        result = await controller.auto_fill(source_range=source_range, target_range=target_range)
        return result


    @agent.tool
    async def format_brush(_ctx: RunContext, source_range: str, target_range: str) -> str:
        """Copy formatting from source range to target range (format painter).
        
        Supports single cell, range, and cross-sheet operations.
        
        Args:
            source_range: Source range in A1 notation, supports cross-sheet (e.g., 'Sheet1!A1')
            target_range: Target range in A1 notation, supports cross-sheet (e.g., 'Sheet2!B1:C1')
        
        Example: format_brush('A1', 'B1:C1') copies formatting from A1 to B1:C1
        """
        result = await controller.format_brush(source_range=source_range, target_range=target_range)
        return result


    @agent.tool
    async def get_conditional_formatting_rules(_ctx: RunContext, sheet_name: str) -> list:
        """Get all conditional formatting rules for the given sheet.
        
        Args:
            sheet_name: Target sheet name
        
        Returns:
            List of conditional formatting rules
        """
        result = await controller.get_conditional_formatting_rules(sheet_name=sheet_name)
        return result

    @agent.tool
    async def add_conditional_formatting_rule(_ctx: RunContext, sheet_name: str, rules: list[dict]) -> str:
        """Add conditional formatting rules. Note: Limited API support."""
        result = await controller.add_conditional_formatting_rule(sheet_name=sheet_name, rules=rules)
        return result

    @agent.tool
    async def set_conditional_formatting_rule(_ctx: RunContext, sheet_name: str, rules: list[dict]) -> str:
        """Set (replace) conditional formatting rules. Note: Limited API support."""
        result = await controller.set_conditional_formatting_rule(sheet_name=sheet_name, rules=rules)
        return result

    @agent.tool
    async def delete_conditional_formatting_rule(_ctx: RunContext, sheet_name: str, rule_ids: list[str]) -> str:
        """Delete conditional formatting rules by ID. Note: Limited API support."""
        result = await controller.delete_conditional_formatting_rule(sheet_name=sheet_name, rule_ids=rule_ids)
        return result

    @agent.tool
    async def add_data_validation_rule(_ctx: RunContext, sheet_name: str, rules: list[dict]) -> str:
        """Add data validation rules (dropdowns, etc.). Note: Limited API support."""
        result = await controller.add_data_validation_rule(sheet_name=sheet_name, rules=rules)
        return result

    @agent.tool
    async def set_data_validation_rule(_ctx: RunContext, sheet_name: str, rules: list[dict]) -> str:
        """Set (replace) data validation rules. Note: Limited API support."""
        result = await controller.set_data_validation_rule(sheet_name=sheet_name, rules=rules)
        return result

    @agent.tool
    async def delete_data_validation_rule(_ctx: RunContext, sheet_name: str, rule_ids: list[str]) -> str:
        """Delete data validation rules by ID. Note: Limited API support."""
        result = await controller.delete_data_validation_rule(sheet_name=sheet_name, rule_ids=rule_ids)
        return result

    @agent.tool
    async def get_data_validation_rules(_ctx: RunContext, sheet_name: str) -> list:
        """Get all data validation rules for a sheet. Note: Limited API support."""
        result = await controller.get_data_validation_rules(sheet_name=sheet_name)
        return result


async def run_query(agent, controller, prompt: str):
    """
    Run a query against the spreadsheet using the agent.
    
    This is an async generator that yields each step/node as the agent processes,
    followed by the final result.
    
    Args:
        agent: The PydanticAI agent
        controller: The UniverSheetsController instance
        prompt: The natural language query
    
    Yields:
        Each message/node during processing, then the final AgentRun result
    
    Example:
        async for item in run_query(agent, controller, "What sheets exist?"):
            if hasattr(item, 'output'):
                # This is the final result
                final_answer = item.output
            else:
                # This is an intermediate step (node)
                print(f"Processing: {item}")
    """
    # Use iter() to get streaming updates for each node
    async with agent.iter(prompt) as result:
        async for message in result:
            yield message
    
    # Yield the final result
    yield result


async def main():
    """Main entry point for CLI usage"""
    # Parse command line arguments
    parser = argparse.ArgumentParser(description="Query Univer Sheets with natural language")
    parser.add_argument(
        'prompt',
        nargs='?',
        default='What sheets are available in this workbook?',
        help='Your question about the spreadsheet'
    )
    parser.add_argument(
        '--url',
        default='http://localhost:3002/sheets/',
        help='URL of the Univer Sheets instance'
    )
    parser.add_argument(
        '--headless',
        action='store_true',
        help='Run browser in headless mode'
    )
    args = parser.parse_args()

    # Create agent and controller
    agent, controller = create_agent()
    register_tools(agent, controller)

    try:
        # Start the controller
        print(f"🚀 Connecting to Univer at {args.url}...\n")
        await controller.start(url=args.url, headless=args.headless)
        print("✅ Connected!\n")
        
        print(f"💬 User: {args.prompt}\n")
        print("🤖 Agent processing...\n")

        # Run the query and consume the generator
        final_result = None
        async for item in run_query(agent, controller, args.prompt):
            # Check if this is the final result (last item yielded)
            # The final item will be the AgentRun result object
            if hasattr(item, 'data') and hasattr(item, '__class__') and 'End' in item.__class__.__name__:
                final_result = item
                print(f"📝 Step: {item}\n\n")
                print("-" * 60)
            elif hasattr(item, 'output'):
                # This is the final result object
                final_result = item
            else:
                # This is an intermediate step
                print(f"📝 Step: {item}\n\n")
                print("-" * 60)
        
        print("\n" + "=" * 60)
        print("FINAL ANSWER")
        print(f"\n{final_result}\n")

    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        await controller.cleanup()


if __name__ == "__main__":
    asyncio.run(main())

