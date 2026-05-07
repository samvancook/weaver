#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import re
from pathlib import Path

RAW_CSV_PATH = Path(__file__).resolve().parent / "Primary Excerpt Database - Excerpt Database.csv"
DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parent / "data" / "int_audit"

LAYOUT_LABEL_PATTERN = re.compile(r"^(full|half|third) page,\s*", re.IGNORECASE)
QUOTED_BG_PATTERN = re.compile(r'^[\"“].*(bg|background)$', re.IGNORECASE)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Audit INT rows and split obvious layout-label rows from excerpt-like interior rows."
    )
    parser.add_argument(
        "--source-csv",
        type=Path,
        default=RAW_CSV_PATH,
        help=f"Raw CSV source path (default: {RAW_CSV_PATH})",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help=f"Audit output directory (default: {DEFAULT_OUTPUT_DIR})",
    )
    return parser.parse_args()


def resolve_content_type(row: dict) -> str:
    return ((row.get("File Type") or "") or (row.get("File Type Helper") or "")).strip().upper()


def classify_int_row(row: dict) -> tuple[str, list[str]]:
    quote = (row.get("Quote") or "").strip()
    full_page_or_excerpt = (row.get("Full page or Excerpt") or "").strip()
    reasons: list[str] = []

    if LAYOUT_LABEL_PATTERN.match(quote):
        reasons.append("layout_label_prefix")
        return "peel_off_layout_int", reasons

    if QUOTED_BG_PATTERN.match(quote):
        reasons.append("quoted_background_variant")
        return "keep_excerpt_linked_int", reasons

    if full_page_or_excerpt.lower() in {"excerpt", "full page", "full poem"}:
        reasons.append(f"full_page_or_excerpt:{full_page_or_excerpt}")
        return "keep_excerpt_linked_int", reasons

    if quote.startswith('"') or quote.startswith("“"):
        reasons.append("quoted_text")
        return "keep_excerpt_linked_int", reasons

    if len(quote.split()) >= 8:
        reasons.append("long_text")
        return "keep_excerpt_linked_int", reasons

    reasons.append("ambiguous_int")
    return "needs_review_int", reasons


def write_csv(path: Path, rows: list[dict], fallback_fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        fieldnames = list(rows[0].keys()) if rows else fallback_fieldnames
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def run_audit(source_csv: Path, output_dir: Path) -> dict:
    buckets = {
        "peel_off_layout_int": [],
        "keep_excerpt_linked_int": [],
        "needs_review_int": [],
    }

    with source_csv.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row_number, row in enumerate(reader, start=2):
            if resolve_content_type(row) != "INT":
                continue
            bucket, reasons = classify_int_row(row)
            buckets[bucket].append(
                {
                    "sourceRow": row_number,
                    "author": (row.get("Author") or "").strip(),
                    "bookTitle": (row.get("Book") or "").strip(),
                    "poemTitle": (row.get("Title") or "").strip(),
                    "excerptText": (row.get("Quote") or "").strip(),
                    "fileName": (row.get("File Name") or "").strip(),
                    "linkUrl": (row.get("Link") or "").strip(),
                    "fullPageOrExcerpt": (row.get("Full page or Excerpt") or "").strip(),
                    "notes": (row.get("Notes") or "").strip(),
                    "intAuditBucket": bucket,
                    "reasons": "; ".join(reasons),
                    "metadataJson": json.dumps(row, ensure_ascii=True),
                }
            )

    output_dir.mkdir(parents=True, exist_ok=True)
    fallback_fields = [
        "sourceRow",
        "author",
        "bookTitle",
        "poemTitle",
        "excerptText",
        "fileName",
        "linkUrl",
        "fullPageOrExcerpt",
        "notes",
        "intAuditBucket",
        "reasons",
        "metadataJson",
    ]

    summary = []
    for bucket, rows in buckets.items():
        output_csv = output_dir / f"{bucket}.csv"
        write_csv(output_csv, rows, fallback_fields)
        summary.append(
            {
                "bucket": bucket,
                "row_count": len(rows),
                "output_csv": str(output_csv),
            }
        )

    manifest = {
        "source_csv": str(source_csv),
        "output_dir": str(output_dir),
        "int_audit": summary,
    }
    (output_dir / "manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    return manifest


def main() -> int:
    args = parse_args()
    result = run_audit(args.source_csv, args.output_dir)
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
