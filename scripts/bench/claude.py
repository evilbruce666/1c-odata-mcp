"""Claude Code (программа командной строки): те же вопросы, тот же тестовый сервер. Использование: claude.py ПРОГОНОВ.
Нужны: вход в Claude Code (claude auth login), CLAUDE_BIN, CLAUDE_MODEL."""
import json
import os
import subprocess
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common
import data

CLAUDE = os.environ.get("CLAUDE_BIN", "claude")
MODEL = os.environ["CLAUDE_MODEL"]


def ask(question, workdir):
    config = os.path.join(workdir, "mcp.json")
    json.dump({"mcpServers": {"demo1c": {"command": "python3", "args": [f"{common.HERE}/stub.py"]}}}, open(config, "w"))
    cmd = [CLAUDE, "-p", question, "--output-format", "stream-json", "--verbose", "--model", MODEL, "--strict-mcp-config",
           "--mcp-config", config, "--tools", "", "--allowedTools", "mcp__demo1c", "--setting-sources", "",
           "--disable-slash-commands", "--no-session-persistence", "--system-prompt", data.SYSTEM]
    started = time.time()
    proc = subprocess.run(cmd, capture_output=True, text=True, stdin=subprocess.DEVNULL, cwd=workdir)
    events = []
    for line in proc.stdout.splitlines():
        try:
            events.append(json.loads(line))
        except ValueError:
            pass
    result = next((e for e in reversed(events) if e.get("type") == "result"), {})
    tools = [b.get("name") for e in events if e.get("type") == "assistant" for b in e["message"]["content"] if b.get("type") == "tool_use"]
    return {"sec": round(time.time() - started, 1), "rc": proc.returncode, "tools": tools, "answer": result.get("result") or ""}


if __name__ == "__main__":
    with tempfile.TemporaryDirectory() as workdir:
        for run in range(int(sys.argv[1])):
            for number, (question, facts) in enumerate(data.QUESTIONS):
                row = ask(question, workdir)
                row.update(who="claude", model=MODEL, run=run, q=number, ok=data.has(row["answer"], facts))
                common.record(row)
                print(run, number, row["sec"], row["ok"], row["tools"], flush=True)
