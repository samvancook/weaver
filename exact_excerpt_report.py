#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path

from excerpt_library import (
    DEFAULT_DB_PATH,
    TEXT_COLUMN_FALLBACKS,
    clean_whitespace,
    connect_library,
    get_exact_excerpt_matches,
    resolve_column,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate an exact-match report against the local excerpt library."
    )
    parser.add_argument("input_csv", type=Path, help="CSV of candidate excerpts to check")
    parser.add_argument(
        "--db-path",
        type=Path,
        default=DEFAULT_DB_PATH,
        help=f"SQLite database path (default: {DEFAULT_DB_PATH})",
    )
    parser.add_argument(
        "--text-column",
        help="Column containing the candidate excerpt text",
    )
    parser.add_argument(
        "--output-csv",
        type=Path,
        help="Optional CSV path for row-by-row results",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print JSON summary to stdout",
    )
    return parser.parse_args()


def build_report(input_csv: Path, db_path: Path, text_column: str | None = None) -> dict:
    connection = connect_library(db_path)
    try:
        with input_csv.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            resolved_text_column = resolve_column(
                reader.fieldnames,
                text_column,
                TEXT_COLUMN_FALLBACKS,
            )
            if not resolved_text_column:
                raise ValueError("Could not find an excerpt text column automatically.")

            row_results: list[dict] = []
            for row_number, record in enumerate(reader, start=2):
                excerpt_text = clean_whitespace(record.get(resolved_text_column, ""))
                matches = get_exact_excerpt_matches(connection, excerpt_text)
                row_results.append(
                    {
                        "sourceRow": row_number,
                        "excerptText": excerpt_text,
                        "exactMatchCount": len(matches),
                        "exactMatchFound": bool(matches),
                        "matches": matches,
                    }
                )
    finally:
        connection.close()

    matched_rows = [row for row in row_results if row["exactMatchFound"]]
    return {
        "input_csv": str(input_csv),
        "db_path": str(db_path),
        "total_rows_checked": len(row_results),
        "matched_rows": len(matched_rows),
        "unmatched_rows": len(row_results) - len(matched_rows),
        "row_results": row_results,
    }


def write_output_csv(output_csv: Path, report: dict) -> None:
    output_csv.parent.mkdir(parents=True, exist_ok=True)
    with output_csv.open("w", encoding="utf-8", newline="") as handle:
        fieldnames = [
            "sourceRow",
            "excerptText",
            "exactMatchFound",
            "exactMatchCount",
            "matchedSourceRows",
            "matchedAuthors",
            "matchedBooks",
            "matchedPoemTitles",
        ]
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in report["row_results"]:
            matches = row["matches"]
            writer.writerow(
                {
                    "sourceRow": row["sourceRow"],
                    "excerptText": row["excerptText"],
                    "exactMatchFound": "Y" if row["exactMatchFound"] else "",
                    "exactMatchCount": row["exactMatchCount"],
                    "matchedSourceRows": "; ".join(str(match["sourceRow"]) for match in matches),
                    "matchedAuthors": "; ".join(match["author"] for match in matches if match["author"]),
                    "matchedBooks": "; ".join(match["bookTitle"] for match in matches if match["bookTitle"]),
                    "matchedPoemTitles": "; ".join(match["poemTitle"] for match in matches if match["poemTitle"]),
                }
            )


def main() -> int:
    args = parse_args()
    report = build_report(args.input_csv, args.db_path, text_column=args.text_column)

    if args.output_csv:
        write_output_csv(args.output_csv, report)

    if args.json or not args.output_csv:
        print(json.dumps(report, indent=2))
    else:
        print(
            json.dumps(
                {
                    "input_csv": report["input_csv"],
                    "db_path": report["db_path"],
                    "total_rows_checked": report["total_rows_checked"],
                    "matched_rows": report["matched_rows"],
                    "unmatched_rows": report["unmatched_rows"],
                    "output_csv": str(args.output_csv),
                },
                indent=2,
            )
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
