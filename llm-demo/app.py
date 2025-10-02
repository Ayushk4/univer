"""
Flask backend for Univer Sheets AI Agent
Uses pydantic_agent.py for agent functionality
"""

from flask import Flask, request, jsonify, send_from_directory, Response
from flask_cors import CORS
import asyncio
import json
import os
import sys
import threading
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from pydantic_agent import create_agent, register_tools, run_query

app = Flask(__name__)
CORS(app)

# Create agent and controller
agent, controller = create_agent()
register_tools(agent, controller)

# Store the event loop that will be used for all async operations
event_loop = None
loop_thread = None


@app.route('/')
def index():
    """Serve the HTML frontend"""
    return send_from_directory('.', 'index.html')


@app.route('/query', methods=['GET'])
def query():
    """Stream query results in real-time using SSE"""
    prompt = request.args.get('prompt', '')
    
    if not prompt:
        return jsonify({'error': 'No prompt provided'}), 400
    
    def generate():
        """Generator that yields SSE events for each step"""
        try:
            print(f"\n💬 Query: {prompt}\n")
            
            # Stream results as they come in using the same event loop as controller
            async def stream_steps():
                async for item in run_query(agent, controller, prompt):
                    step_str = str(item)
                    event_data = {'type': 'step', 'data': step_str}
                    
                    if hasattr(item, 'output'):
                        event_data['type'] = 'final'
                        event_data['answer'] = str(item.output)
                        if hasattr(item, 'usage'):
                            event_data['usage'] = str(item.usage)
                        print(f"✅ Final result received\n")
                    else:
                        print(f"📝 Step: {item}\n")
                    
                    yield event_data
            
            # Use run_coroutine_threadsafe to execute in the event loop thread
            import queue
            result_queue = queue.Queue()
            
            async def run_and_queue():
                try:
                    async for event_data in stream_steps():
                        result_queue.put(('data', event_data))
                    result_queue.put(('done', None))
                except Exception as e:
                    result_queue.put(('error', e))
            
            # Submit to the event loop running in the background thread
            future = asyncio.run_coroutine_threadsafe(run_and_queue(), event_loop)
            
            # Yield results as they come in from the queue with timeout/keepalive
            import time
            while True:
                try:
                    # Use timeout to send keepalive if no data for 2 seconds
                    msg_type, data = result_queue.get(timeout=2.0)
                    if msg_type == 'data':
                        yield f"data: {json.dumps(data)}\n\n"
                    elif msg_type == 'done':
                        yield f"data: {json.dumps({'type': 'done'})}\n\n"
                        break
                    elif msg_type == 'error':
                        raise data
                except queue.Empty:
                    # Send keepalive/heartbeat to prevent connection timeout
                    yield f": keepalive {time.time()}\n\n"
                
        except Exception as e:
            print(f"❌ Error: {e}")
            import traceback
            traceback.print_exc()
            yield f"data: {json.dumps({'type': 'error', 'message': str(e)})}\n\n"
    
    return Response(generate(), mimetype='text/event-stream')


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
    print("✅ Connected!")


if __name__ == '__main__':
    # Create event loop and run it in a background thread
    event_loop = asyncio.new_event_loop()
    loop_thread = threading.Thread(target=run_event_loop, args=(event_loop,), daemon=True)
    loop_thread.start()
    
    # Initialize controller in the event loop
    future = asyncio.run_coroutine_threadsafe(init_controller(), event_loop)
    future.result()  # Wait for initialization to complete
    
    # Start Flask server
    port = int(os.environ.get('PORT', 5001))
    app.run(debug=False, port=port, host='0.0.0.0')

