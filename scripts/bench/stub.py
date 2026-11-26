"""MCP-сервер на stdio с выдуманными ответами: сети и 1С нет. Его подключают все участники замера."""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import data

TOOLS = data.load_tools()
for line in sys.stdin:
    m = json.loads(line)
    i, method = m.get("id"), m.get("method")
    if i is None:
        continue
    if method == "initialize":
        r = {
            "protocolVersion": m["params"].get("protocolVersion", "2025-06-18"),
            "capabilities": {"tools": {}},
            "serverInfo": {"name": "demo1c", "version": "1"},
            "instructions": "Доступ к демонстрационной базе 1С (Бухгалтерия 3.0). Только чтение.",
        }
    elif method == "tools/list":
        r = {"tools": [{**t, "annotations": {"readOnlyHint": True, "destructiveHint": False, "openWorldHint": False}} for t in TOOLS]}
    elif method == "tools/call":
        name, args = m["params"]["name"], m["params"].get("arguments", {})
        r = {"content": [{"type": "text", "text": data.text(name, args)}]}
    elif method == "ping":
        r = {}
    else:
        print(json.dumps({"jsonrpc": "2.0", "id": i, "error": {"code": -32601, "message": "нет такого метода"}}), flush=True)
        continue
    print(json.dumps({"jsonrpc": "2.0", "id": i, "result": r}, ensure_ascii=False), flush=True)
