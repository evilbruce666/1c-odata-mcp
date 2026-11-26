"""Локальная модель через API LM Studio: цикл «вопрос → вызовы инструментов → ответ». Использование: lm.py ПРОГОНОВ [МОДЕЛЬ]."""
import json
import os
import sys
import time
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common
import data

URL = os.environ.get("LMSTUDIO_URL", "http://localhost:1234/v1/chat/completions")
MODEL = sys.argv[2] if len(sys.argv) > 2 else os.environ.get("LMSTUDIO_MODEL", "qwen3.5-9b")
TOOLS = [
    {"type": "function", "function": {"name": t["name"].replace(".", "__"), "description": t["description"], "parameters": t["inputSchema"]}}
    for t in data.load_tools()
]


def api(messages):
    body = {"model": MODEL, "messages": messages, "tools": TOOLS, "temperature": 0.1, "max_tokens": 3000}
    request = urllib.request.Request(URL, json.dumps(body).encode(), {"Content-Type": "application/json"})
    return json.load(urllib.request.urlopen(request, timeout=900))


def ask(question):
    messages = [{"role": "system", "content": data.SYSTEM}, {"role": "user", "content": question}]
    started, called = time.time(), []
    for _ in range(8):
        message = api(messages)["choices"][0]["message"]
        calls = message.get("tool_calls") or []
        if not calls:
            return {"sec": round(time.time() - started, 1), "tools": called, "answer": message.get("content") or ""}
        messages.append({"role": "assistant", "content": message.get("content") or "", "tool_calls": calls})
        for call in calls:
            name = call["function"]["name"].replace("__", ".")
            called.append(name)
            messages.append({"role": "tool", "tool_call_id": call["id"], "content": data.text(name, json.loads(call["function"]["arguments"] or "{}"))})
    return {"sec": round(time.time() - started, 1), "tools": called, "answer": "", "error": "больше 8 шагов"}


if __name__ == "__main__":
    for run in range(int(sys.argv[1])):
        for number, (question, facts) in enumerate(data.QUESTIONS):
            row = ask(question)
            row.update(who="lmstudio", model=MODEL, run=run, q=number, ok=data.has(row["answer"], facts))
            common.record(row)
            print(run, number, row["sec"], row["ok"], row["tools"], flush=True)
