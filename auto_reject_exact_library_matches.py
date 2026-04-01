#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import re
import subprocess
import sys
from pathlib import Path

from excerpt_library import connect_library, find_library_excerpt_match

DEFAULT_API_BASE_URL = (
    "https://script.google.com/macros/s/"
    "AKfycbytGpEIr8CbPdJDDF5tAn7JxH0-kNuBogD_elDlF8ljgb_7ebF2nt-yn45hUIwzRd2Xfg/exec"
)


def run_curl(args: list[str], stdin_text: str | None = None) -> str:
    completed = subprocess.run(
        ["curl", *args],
        input=stdin_text,
        text=True,
        capture_output=True,
        check=False,
    )
    if completed.returncode != 0:
        raise RuntimeError(completed.stderr.strip() or "curl failed")
    return completed.stdout


def fetch_pending_records(api_base_url: str) -> dict:
    payload = run_curl(
        [
            "-L",
            "-sS",
            f"{api_base_url}?action=pendingRecords&callback=weaverPending",
        ]
    )
    start = payload.find("(")
    end = payload.rfind(")")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("Pending records response was not valid JSONP.")
    return json.loads(payload[start + 1 : end])


def batch_save_reviews(api_base_url: str, updates: list[dict]) -> dict:
    payload = run_curl(
        [
            "-L",
            "-sS",
            "-X",
            "POST",
            "-H",
            "Content-Type: application/json",
            f"{api_base_url}?action=saveReviews",
            "--data-binary",
            "@-",
        ],
        stdin_text=json.dumps({"updates": updates}),
    )
    return json.loads(payload or "{}")


def should_auto_reject(match: dict | None) -> bool:
    if not match:
        return False
    return (
        match.get("matchType") == "exact"
        and bool(match.get("formattingMatch"))
        and bool(match.get("lineBreaksMatch"))
    )


def build_update(record: dict) -> dict:
    return {
        "sourceRow": record.get("sourceRow"),
        "recordId": record.get("recordId", ""),
        "approval": "reject",
        "reviewDecision": "reject",
        "correctionNote": record.get("correctionNote", ""),
        "correctedAuthor": record.get("correctedAuthor", ""),
        "correctedTitle": record.get("correctedTitle", ""),
        "correctedBookTitle": record.get("correctedBookTitle", ""),
        "correctedExcerpt": record.get("correctedExcerpt", ""),
        "graphicsQi": "1" if record.get("useForQi") else "0",
        "useForQi": "1" if record.get("useForQi") else "0",
        "photos": "1" if record.get("useForInt") else "0",
        "useForInt": "1" if record.get("useForInt") else "0",
    }


def write_report(path: Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "sourceRow",
                "recordId",
                "author",
                "title",
                "bookTitle",
                "matchSourceRow",
                "matchAuthor",
                "matchTitle",
                "matchBookTitle",
                "excerptPreview",
            ],
        )
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Reject pending Weaver rows that are strict exact matches in the excerpt library."
    )
    parser.add_argument("--api-base-url", default=DEFAULT_API_BASE_URL)
    parser.add_argument("--apply", action="store_true", help="Actually write REJECT decisions back to the sheet.")
    parser.add_argument("--limit", type=int, default=0, help="Optional limit for matched rows.")
    parser.add_argument(
        "--report-csv",
        default="data/exact_library_match_autoreject_report.csv",
        help="Where to write a report of matched rows.",
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=100,
        help="How many updates to send per batch when --apply is used.",
    )
    args = parser.parse_args()

    try:
        pending = fetch_pending_records(args.api_base_url)
    except (RuntimeError, ValueError, json.JSONDecodeError) as error:
        print(f"Failed to load pending records: {error}", file=sys.stderr)
        return 1

    records = pending.get("records") or []
    if not isinstance(records, list):
        print("Pending records payload did not contain a records list.", file=sys.stderr)
        return 1

    connection = connect_library()
    matches: list[dict] = []
    updates: list[dict] = []
    try:
        for record in records:
            match = find_library_excerpt_match(
                connection,
                record.get("excerptText"),
                book_title=record.get("bookTitle"),
                author=record.get("author"),
            )
            if not should_auto_reject(match):
                continue

            matches.append(
                {
                    "sourceRow": record.get("sourceRow"),
                    "recordId": record.get("recordId", ""),
                    "author": record.get("author", ""),
                    "title": record.get("title", ""),
                    "bookTitle": record.get("bookTitle", ""),
                    "matchSourceRow": match.get("sourceRow", ""),
                    "matchAuthor": match.get("author", ""),
                    "matchTitle": match.get("poemTitle", ""),
                    "matchBookTitle": match.get("bookTitle", ""),
                    "excerptPreview": (record.get("excerptText") or "").replace("\n", "\\n")[:200],
                }
            )
            updates.append(build_update(record))

            if args.limit and len(updates) >= args.limit:
                break
    finally:
        connection.close()

    report_path = Path(args.report_csv)
    write_report(report_path, matches)

    print(f"Pending records scanned: {len(records)}")
    print(f"Strict exact library matches: {len(updates)}")
    print(f"Report: {report_path}")

    if not args.apply:
        print("Dry run only. Re-run with --apply to write REJECT decisions.")
        return 0

    if not updates:
        print("No strict exact matches to reject.")
        return 0

    total_saved = 0
    batch_size = max(1, args.batch_size)
    for start in range(0, len(updates), batch_size):
      chunk = updates[start : start + batch_size]
      try:
          result = batch_save_reviews(args.api_base_url, chunk)
      except (RuntimeError, json.JSONDecodeError) as error:
          print(
              f"Batch save failed for rows {start + 1}-{start + len(chunk)}: {error}",
              file=sys.stderr,
          )
          return 1

      if not result.get("ok"):
          print(
              f"Batch save failed for rows {start + 1}-{start + len(chunk)}: "
              f"{result.get('error', 'Unknown error')}",
              file=sys.stderr,
          )
          return 1
      saved_count = int(result.get("savedCount", 0) or 0)
      total_saved += saved_count
      print(f"Saved batch {start + 1}-{start + len(chunk)} ({saved_count} rows).")

    print(f"Saved {total_saved} exact-match rejections.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
