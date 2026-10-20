#!/usr/bin/env python3
"""Copy Suffolk Law's academic calendar from the Suffolk assistant's verified sources into data/.

    python3 scripts/import_calendar.py --source /path/to/suffolk-chatbot/data/current_sources

Maintainers only (the assistant's repo is private). Everyone else uses the committed data file.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
NAME = "law-academic-calendar-2026-27"


def normalize(raw: dict, url: str) -> dict:
    events = []
    for e in raw["entries"]:
        day = dt.datetime.strptime(e["date"], "%B %d, %Y").date()
        events.append({"term": e["term"].replace(" Law ", " "), "event": e["event"], "date": day.isoformat()})
    events.sort(key=lambda e: e["date"])
    return {"source": {"url": url, "verified_at": raw.get("verified_at", "")}, "events": events}


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--source", type=Path, required=True)
    args = ap.parse_args(argv)
    raw = json.loads((args.source / f"{NAME}.json").read_text(encoding="utf-8"))
    url = ""
    txt = args.source / f"{NAME}.txt"
    if txt.exists():
        for line in txt.read_text(encoding="utf-8").splitlines()[:6]:
            if line.startswith("Canonical URL: "):
                url = line.split(": ", 1)[1].strip()
    data = normalize(raw, url)
    out = ROOT / "data" / "academic-calendar.json"
    out.write_text(json.dumps(data, indent=1) + "\n", encoding="utf-8")
    print(f"wrote {len(data['events'])} events to {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
