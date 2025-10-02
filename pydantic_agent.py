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

