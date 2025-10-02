"""
UNIVER_KEY=sk-lxIlTQ8zT2SFuk1lOEczBAy7d649c16c5c1674d python
"""
import os
import requests
import json

API_KEY = os.environ.get('UNIVER_KEY')
if not API_KEY:
    raise ValueError("UNIVER_KEY not set")

SESSION_ID = "test-session-1"
MCP_URL = f"https://mcp.univer.ai/mcp/?univer_session_id={SESSION_ID}"

headers = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json",
    "Accept": "application/json, text/event-stream"
}

for method in ['tools/list', 'resources/list', 'prompts/list']:
    print("\n\n\n\n")
    print("================", method, "================")
    payload = {"jsonrpc": "2.0", "id": 1, "method": method, "params": {}}
    response = requests.post(MCP_URL, headers=headers, json=payload)
    data = None
    for line in response.text.split('\n'):
        if line.startswith('data: '):
            data = json.loads(line[6:])  # Parse JSON after "data: "
            break
    for k in data['result'][method.split('/')[0]]:
        print('\n\n\n\n')
        print('---------------------', method, ":", k['name'], '---------------------')
        print(k['description'])
