"""Codex (программа командной строки): те же вопросы, тот же тестовый сервер. Использование: codex.py ПРОГОНОВ.
Нужны: CODEX_BIN (путь к codex), CODEX_MODEL (по умолчанию из вашего config.toml не берётся — задайте явно)."""
import json
import os
import subprocess
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common
import data

CODEX = os.environ.get("CODEX_BIN", "codex")
MODEL = os.environ["CODEX_MODEL"]
EFFORT = os.environ.get("CODEX_EFFORT", "medium")
OFF = ["shell_tool", "unified_exec", "computer_use", "browser_use", "browser_use_external", "browser_use_full_cdp_access",
       "apps", "plugins", "image_generation", "multi_agent", "memories"]


def ask(question, workdir):
    cmd = [CODEX, "exec", "--json", "--ephemeral", "--skip-git-repo-check", "--ignore-user-config", "--ignore-rules",
           "-s", "read-only", "-C", workdir, "-m", MODEL, "-c", f'model_reasoning_effort="{EFFORT}"', "-c", 'web_search="disabled"',
           "-c", 'mcp_servers.demo1c.command="python3"', "-c", f'mcp_servers.demo1c.args=["{common.HERE}/stub.py"]',
           "-c", 'mcp_servers.demo1c.default_tools_approval_mode="approve"']
    for feature in OFF:
        cmd += ["--disable", feature]
    cmd.append(data.SYSTEM + "\n\n" + question)
    started = time.time()
    proc = subprocess.run(cmd, capture_output=True, text=True, stdin=subprocess.DEVNULL)
    events = []
    for line in proc.stdout.splitlines():
        try:
            events.append(json.loads(line))
        except ValueError:
            pass
    items = [e.get("item", {}) for e in events if e.get("type") == "item.completed"]
    answer = next((i.get("text", "") for i in reversed(items) if i.get("type") == "agent_message"), "")
    return {"sec": round(time.time() - started, 1), "rc": proc.returncode, "tools": [i.get("tool") for i in items if i.get("type") == "mcp_tool_call"], "answer": answer}


if __name__ == "__main__":
    with tempfile.TemporaryDirectory() as workdir:
        for run in range(int(sys.argv[1])):
            for number, (question, facts) in enumerate(data.QUESTIONS):
                row = ask(question, workdir)
                row.update(who="codex", model=MODEL, run=run, q=number, ok=data.has(row["answer"], facts))
                common.record(row)
                print(run, number, row["sec"], row["ok"], row["tools"], flush=True)
