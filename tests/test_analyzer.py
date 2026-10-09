"""Разбор отзывов без обращения к нейросети: ответ модели подменён простым правилом."""
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import analyze_reviews as A  # noqa: E402


def fake_ask(prompt):
    items = re.findall(r"^(\d+)\. (.*)$", prompt.split("Отзывы:")[1], re.M)
    out = []
    for i, text in items:
        t = text.lower()
        tone = "отрицательный" if any(w in t for w in ("ужас", "плох", "разочар")) else "положительный"
        topics = ["Размер и посадка"] if "размер" in t else ["Другое"]
        out.append({"id": int(i), "tone": tone, "topics": topics})
    return "```json\n" + json.dumps(out, ensure_ascii=False) + "\n```"


def test_full_run(tmp_path):
    out = tmp_path / "result.xlsx"
    A.main(str(ROOT / "data" / "отзывы_пример.csv"), out=str(out), ask=fake_ask)
    assert out.exists() and out.stat().st_size > 10_000
