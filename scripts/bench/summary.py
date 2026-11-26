"""Сводка out/results.jsonl: сколько ответов верных и сколько секунд от вопроса до ответа (медиана, диапазон)."""
import json
import os
import statistics
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common

rows = [json.loads(line) for line in open(os.path.join(common.OUT, "results.jsonl"), encoding="utf-8")]
for who in dict.fromkeys(r["who"] for r in rows):
    part = [r for r in rows if r["who"] == who and r.get("sec")]
    secs = [r["sec"] for r in part]
    print(f"{who:9} {part[0]['model']:22} прогонов {len(part):3}  верных {sum(r['ok'] for r in part):3}  "
          f"медиана {statistics.median(secs):5.1f} с  от {min(secs):5.1f} до {max(secs):5.1f}")
    by_question = {}
    for r in part:
        by_question.setdefault(r["q"], []).append(r["sec"])
    print("          по вопросам (медиана): " + "  ".join(f"{q + 1}: {statistics.median(v):.1f}" for q, v in sorted(by_question.items())))
