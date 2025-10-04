#!/usr/bin/env python3
"""
Single command to test the backend with SSE streaming.
Launches server, runs test, and cleans up.

Usage: python test.py "What sheets are here?"
"""

import sys
import os
import time
import signal
import subprocess
import requests
import json
from pathlib import Path

# Change to script directory
os.chdir(Path(__file__).parent)

def test_sse(prompt: str = "What sheets are available?"):
    """Test SSE endpoint"""
    print(f"🔍 Testing: '{prompt}'\n")
    
    url = "http://127.0.0.1:5001/query"
    
    try:
        response = requests.get(
            url,
            params={'prompt': prompt},
            stream=True,
            timeout=(5, 120)  # 5s connect, 120s read
        )
        
        print("=" * 60)
        step_count = 0
        
        for line in response.iter_lines(decode_unicode=True):
            if line.startswith('data: '):
                data = json.loads(line[6:])
                
                if data['type'] == 'step':
                    step_count += 1
                    print(f"\n[Step {step_count}]")
                    print(data['data'][:500])  # Truncate long output
                    print("-" * 40)
                
                elif data['type'] == 'final':
                    print(f"\n{'=' * 60}")
                    print("✅ ANSWER")
                    print(f"{'=' * 60}")
                    print(data['answer'])
                    if 'usage' in data:
                        print(f"\n📊 {data['usage']}")
                
                elif data['type'] == 'done':
                    print(f"\n{'=' * 60}")
                    print("✓ Completed")
                    print(f"{'=' * 60}\n")
                    return True
                
                elif data['type'] == 'error':
                    print(f"\n❌ Error: {data['message']}")
                    return False
        
        return True
        
    except requests.exceptions.ConnectionError:
        print("❌ Server not running on port 5001")
        return False
    except requests.exceptions.Timeout:
        print("❌ Request timed out (>120s)")
        return False
    except Exception as e:
        print(f"❌ Error: {e}")
        return False


def main():
    prompt = sys.argv[1] if len(sys.argv) > 1 else "What sheets are available?"
    
    print("🚀 Starting Flask server...\n")
    
    # Start server
    server = subprocess.Popen(
        [sys.executable, "app.py"],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1
    )
    
    # Wait for server ready
    print("⏳ Waiting for server to initialize...\n")
    ready = False
    for i in range(60):
        try:
            if requests.get("http://127.0.0.1:5001/", timeout=1).status_code == 200:
                ready = True
                print("✅ Server ready!\n")
                time.sleep(2)  # Extra time for full initialization
                break
        except:
            pass
        time.sleep(0.5)
    
    if not ready:
        print("❌ Server failed to start")
        server.kill()
        return 1
    
    try:
        # Run test
        success = test_sse(prompt)
        if(success):
            print("✅ Test passed")
        else:
            print("❌ Test failed")
        return 0 if success else 1
    finally:
        print("\n🛑 Stopping server...")
        server.send_signal(signal.SIGINT)
        try:
            server.wait(timeout=5)
        except subprocess.TimeoutExpired:
            server.kill()


if __name__ == "__main__":
    sys.exit(main())

