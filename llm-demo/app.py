"""
Flask backend for Univer Sheets AI Agent
Uses agent.iter() for streaming without AG-UI protocol dependencies
"""

from flask import Flask, request, jsonify, send_from_directory, Response
from flask_cors import CORS
import asyncio
import json
import os
import sys
import threading

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from pydantic_agent import create_agent, register_tools

app = Flask(__name__)
CORS(app)

# Create agent and controller
agent, controller = create_agent()
register_tools(agent, controller)

# Store the event loop
event_loop = None
loop_thread = None


@app.route('/')
def index():
    """Serve the HTML frontend"""
    return send_from_directory('.', 'index.html')


@app.route('/query', methods=['POST'])
def query():
    """Handle queries with streaming SSE response"""
    try:
        data = request.get_json()
        # Support both direct prompt and AG-UI message format
        if 'prompt' in data:
            prompt = data['prompt']
        elif 'messages' in data and len(data['messages']) > 0:
            prompt = data['messages'][-1].get('content', '')
        else:
            return jsonify({'error': 'No prompt or messages provided'}), 400
        
        if not prompt:
            return jsonify({'error': 'Empty prompt'}), 400
            
    except Exception as e:
        return jsonify({'error': f'Invalid request: {str(e)}'}), 400
    
    def generate():
        """Generator that yields formatted SSE events"""
        try:
            print(f"\n💬 Query: {prompt}\n")
            
            import queue
            result_queue = queue.Queue()
            
            async def run_and_stream():
                try:
                    # Use agent.iter() to get node-by-node execution
                    async with agent.iter(prompt, deps=controller) as agent_run:
                        async for node in agent_run:
                            # Format each node as an event
                            events = format_node_to_events(node)
                            for event in events:
                                result_queue.put(('event', event))
                    
                    result_queue.put(('done', None))
                except Exception as e:
                    result_queue.put(('error', e))
            
            # Submit to event loop
            future = asyncio.run_coroutine_threadsafe(run_and_stream(), event_loop)
            
            # Stream results
            import time
            while True:
                try:
                    msg_type, data = result_queue.get(timeout=2.0)
                    if msg_type == 'event':
                        # Send as SSE
                        yield f"data: {json.dumps(data)}\n\n"
                    elif msg_type == 'done':
                        print("✅ Stream completed\n")
                        break
                    elif msg_type == 'error':
                        raise data
                except queue.Empty:
                    # Keepalive
                    yield f": keepalive {time.time()}\n\n"
                
        except Exception as e:
            print(f"❌ Error: {e}")
            import traceback
            traceback.print_exc()
            error_event = {'type': 'error', 'error': str(e)}
            yield f"data: {json.dumps(error_event)}\n\n"
    
    return Response(
        generate(),
        mimetype='text/event-stream',
        headers={
            'Cache-Control': 'no-cache',
            'X-Accel-Buffering': 'no'
        }
    )


def format_node_to_events(node):
    """Convert PydanticAI node to event format (returns list of events)"""
    from pydantic_ai import Agent
    from pydantic_ai.messages import (
        UserPromptPart, SystemPromptPart, TextPart, 
        ToolCallPart, ToolReturnPart
    )
    
    events = []
    
    if Agent.is_user_prompt_node(node):
        events.append({
            'type': 'agent_message',
            'content': f"Processing: {node.user_prompt}"
        })
    
    elif Agent.is_model_request_node(node):
        # Extract meaningful parts from the request
        for part in node.request.parts:
            if isinstance(part, ToolReturnPart):
                events.append({
                    'type': 'tool_result',
                    'tool_name': part.tool_name,
                    'result': str(part.content)
                })
    
    elif Agent.is_call_tools_node(node):
        # Process each part of the response
        for part in node.model_response.parts:
            if isinstance(part, TextPart):
                events.append({
                    'type': 'agent_message_delta',
                    'content': part.content
                })
            elif isinstance(part, ToolCallPart):
                events.append({
                    'type': 'tool_call',
                    'tool_name': part.tool_name,
                    'arguments': part.args_as_dict()
                })
    
    elif Agent.is_end_node(node):
        events.append({
            'type': 'final_result',
            'result': str(node.data.output)
        })
    
    return events if events else []


@app.route('/health', methods=['GET'])
def health():
    """Health check endpoint"""
    return jsonify({'status': 'healthy'})


def run_event_loop(loop):
    """Run event loop in a background thread"""
    asyncio.set_event_loop(loop)
    loop.run_forever()


async def init_controller():
    """Initialize the controller on startup"""
    url = os.environ.get('UNIVER_URL', 'http://localhost:3002/sheets/')
    headless = os.environ.get('HEADLESS', 'false').lower() == 'true'
    print(f"🚀 Connecting to Univer at {url}...")
    await controller.start(url=url, headless=headless)
    print("✅ Connected to Univer!")


if __name__ == '__main__':
    # Create event loop and run it in a background thread
    event_loop = asyncio.new_event_loop()
    loop_thread = threading.Thread(target=run_event_loop, args=(event_loop,), daemon=True)
    loop_thread.start()
    
    # Initialize controller
    future = asyncio.run_coroutine_threadsafe(init_controller(), event_loop)
    future.result()
    
    # Start Flask server
    port = int(os.environ.get('PORT', 5001))
    print(f"🌐 Server running on http://localhost:{port}")
    app.run(debug=False, port=port, host='0.0.0.0')