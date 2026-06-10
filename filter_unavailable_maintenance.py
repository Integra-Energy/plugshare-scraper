#!/usr/bin/env python3
"""Filter the sorted all-status PlugShare export to target unavailable chargers."""

from __future__ import annotations

import argparse
import csv
from collections import Counter
from pathlib import Path

TARGET_STATUSES = {"OUTOFORDER", "UNDER_REPAIR", "UNAVAILABLE"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Filter unavailable and maintenance charger rows.")
    parser.add_argument("input_csv", type=Path)
    parser.add_argument("output_csv", type=Path)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    counts: Counter[str] = Counter()
    args.output_csv.parent.mkdir(parents=True, exist_ok=True)
    with args.input_csv.open(encoding="utf-8-sig", newline="") as source:
        reader = csv.DictReader(source)
        with args.output_csv.open("w", encoding="utf-8-sig", newline="") as output:
            writer = csv.DictWriter(output, fieldnames=reader.fieldnames)
            writer.writeheader()
            for row in reader:
                status = row.get("status", "")
                if status not in TARGET_STATUSES:
                    continue
                writer.writerow(row)
                counts[status] += 1
    print(f"Wrote {sum(counts.values())} target rows to {args.output_csv.resolve()}")
    for status in sorted(counts):
        print(f"{status}: {counts[status]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
