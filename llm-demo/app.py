"""
Flask backend for Univer Sheets AI Agent
Uses agent.iter() for streaming without AG-UI protocol dependencies
"""

from flask import Flask, request, jsonify, send_from_directory, Response
from flask_cors import CORS
from flask_socketio import SocketIO
import asyncio
import json
import os
import sys
import threading
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from pydantic_agent import create_agent, register_tools

# Import logging configuration
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    from logging_config import audit_logger, setup_logging
    setup_logging()
    import logging
    logger = logging.getLogger(__name__)
except ImportError:
    import logging
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger(__name__)
    audit_logger = logger

app = Flask(__name__)
CORS(app)

# Initialize SocketIO with simple-websocket backend (Python 3.12 compatible)
# Use default async_mode which will auto-detect the best option
socketio = SocketIO(app, cors_allowed_origins="*", logger=False, engineio_logger=False)

# Snapshot file path
SNAPSHOT_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'workbook_snapshot.json')

# Lock for synchronizing snapshot file access between writer and watcher
snapshot_lock = threading.Lock()

# Create agent and controller with shared snapshot lock
agent, controller = create_agent()
controller.snapshot_lock = snapshot_lock  # Share the lock with controller
register_tools(agent, controller)

# Store the event loop
event_loop = None
loop_thread = None

# File watcher for snapshot changes
class SnapshotChangeHandler(FileSystemEventHandler):
    """Watch for changes to workbook_snapshot.json and broadcast via WebSocket"""
    
    def __init__(self):
        super().__init__()
        self.last_broadcast_time = 0
        self.debounce_seconds = 0.5  # Debounce to avoid multiple rapid broadcasts
    
    def on_modified(self, event):
        if event.src_path.endswith('workbook_snapshot.json'):
            import time
            current_time = time.time()
            
            # Debounce: only broadcast if enough time has passed
            if current_time - self.last_broadcast_time < self.debounce_seconds:
                return
            
            print(f"📥 Snapshot file changed, broadcasting update...")
            
            # Use lock to ensure file is completely written
            with snapshot_lock:
                try:
                    with open(SNAPSHOT_FILE, 'r') as f:
                        snapshot = json.load(f)
                    socketio.emit('snapshot_updated', snapshot)
                    self.last_broadcast_time = current_time
                    print(f"✅ Broadcast snapshot update to all clients")
                except json.JSONDecodeError as e:
                    print(f"⚠️  JSON parse error: {e}")
                except Exception as e:
                    print(f"❌ Error broadcasting snapshot: {e}")

# Initialize file watcher
file_watcher_observer = None


@app.route('/')
def index():
    """Serve the HTML frontend"""
    return send_from_directory('.', 'index.html')


@app.route('/univer_sdk_loader.html')
def univer_sdk_loader():
    """Serve the Univer SDK loader HTML for iframe"""
    return send_from_directory('.', 'univer_sdk_loader.html')


@app.route('/query', methods=['POST'])
def query():
    """Handle queries with streaming SSE response"""
    start_time = time.time()
    prompt = None
    
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
        
        # AUDIT: Log incoming query
        audit_logger.info(
            f"Query received: {prompt[:100]}{'...' if len(prompt) > 100 else ''}",
            extra={
                'user_action': 'query',
                'operation': 'ai_query',
                'result': 'processing'
            }
        )
            
    except Exception as e:
        logger.error(f"Invalid request: {e}", exc_info=True)
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
                    async with agent.iter(prompt) as agent_run:
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
            while True:
                try:
                    msg_type, data = result_queue.get(timeout=2.0)
                    if msg_type == 'event':
                        # Send as SSE
                        yield f"data: {json.dumps(data)}\n\n"
                    elif msg_type == 'done':
                        # AUDIT LOG: Query completed successfully
                        duration_ms = (time.time() - start_time) * 1000
                        print("✅ Stream completed\n")
                        audit_logger.info(
                            "Query completed successfully",
                            extra={
                                'user_action': 'query',
                                'operation': 'ai_query',
                                'result': 'success',
                                'duration_ms': round(duration_ms, 2)
                            }
                        )
                        break
                    elif msg_type == 'error':
                        raise data
                except queue.Empty:
                    # Keepalive
                    yield f": keepalive {time.time()}\n\n"
                
        except Exception as e:
            # ERROR + AUDIT LOG
            duration_ms = (time.time() - start_time) * 1000
            logger.error(f"Query error: {e}", exc_info=True)
            audit_logger.error(
                f"Query failed: {str(e)}",
                extra={
                    'user_action': 'query',
                    'operation': 'ai_query',
                    'result': 'failed',
                    'duration_ms': round(duration_ms, 2)
                }
            )
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


@app.route('/snapshot', methods=['GET'])
def get_snapshot():
    """Serve the current workbook snapshot"""
    try:
        if os.path.exists(SNAPSHOT_FILE):
            with open(SNAPSHOT_FILE, 'r') as f:
                snapshot = json.load(f)
            return jsonify(snapshot)
        else:
            return jsonify({'error': 'No snapshot found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@socketio.on('connect')
def handle_connect():
    """Handle WebSocket connection"""
    print(f"✅ Client connected")
    audit_logger.info(
        "Client connected to WebSocket",
        extra={
            'user_action': 'websocket_connect',
            'operation': 'connection',
            'result': 'success'
        }
    )


@socketio.on('disconnect')
def handle_disconnect():
    """Handle WebSocket disconnection"""
    print(f"⚠️  Client disconnected")
    audit_logger.info(
        "Client disconnected from WebSocket",
        extra={
            'user_action': 'websocket_disconnect',
            'operation': 'connection',
            'result': 'success'
        }
    )


def run_event_loop(loop):
    """Run event loop in a background thread"""
    asyncio.set_event_loop(loop)
    loop.run_forever()


async def init_controller():
    """Initialize the controller on startup"""
    url = os.environ.get('UNIVER_URL', 'http://localhost:3002/sheets/')
    # Set headless=False to see the actual browser where changes happen
    headless = os.environ.get('HEADLESS', 'false').lower() == 'true'
    print(f"🚀 Connecting to Univer at {url}...")
    print(f"   Browser mode: {'headless' if headless else 'visible'}")
    await controller.start(url=url, headless=headless)
    print("✅ Connected to Univer!")


def start_file_watcher():
    """Start watching the snapshot file for changes"""
    global file_watcher_observer
    
    # Watch the directory containing the snapshot file
    watch_dir = os.path.dirname(SNAPSHOT_FILE)
    
    if os.path.exists(watch_dir):
        event_handler = SnapshotChangeHandler()
        file_watcher_observer = Observer()
        file_watcher_observer.schedule(event_handler, watch_dir, recursive=False)
        file_watcher_observer.start()
        print(f"👁️  Watching {SNAPSHOT_FILE} for changes...")
    else:
        print(f"⚠️  Snapshot directory not found: {watch_dir}")


if __name__ == '__main__':
    # Create event loop and run it in a background thread
    event_loop = asyncio.new_event_loop()
    loop_thread = threading.Thread(target=run_event_loop, args=(event_loop,), daemon=True)
    loop_thread.start()
    
    # Initialize controller
    future = asyncio.run_coroutine_threadsafe(init_controller(), event_loop)
    future.result()
    
    # Start file watcher for snapshot changes
    start_file_watcher()
    print("🔌 WebSocket enabled for real-time sync")
    
    # Start Flask server with SocketIO
    port = int(os.environ.get('PORT', 5001))
    print(f"🌐 Server running on http://localhost:{port}")
    socketio.run(app, debug=False, port=port, host='0.0.0.0')