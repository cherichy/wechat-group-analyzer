import urllib.request
import json
import sys

url = 'http://127.0.0.1:5200/api/v1/chatrooms/57327409534@chatroom'
try:
    req = urllib.request.Request(url)
    with urllib.request.urlopen(req, timeout=30) as resp:
        raw_bytes = resp.read()
    # Save raw bytes directly
    with open('chatroom_info_raw.json', 'wb') as f:
        f.write(raw_bytes)
    # Try parsing
    data = json.loads(raw_bytes.decode('utf-8'))
    print(f"Success: {len(data['data']['users'])} members")
    print(f"Owner: {data['data']['owner']}")
except Exception as e:
    print(f"Error: {e}")
    sys.exit(1)
