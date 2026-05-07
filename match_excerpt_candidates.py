#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import sys
from pathlib import Path

from excerpt_library import (
    AUTHOR_COLUMN_FALLBACKS,
    BOOK_COLUMN_FALLBACKS,
    DEFAULT_DB_PATH,
    TEXT_COLUMN_FALLBACKS,
    clean_whitespace,
    connect_library,
    find_library_excerpt_match,
    resolve_column,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Flag candidate excerpts as already-existing or new against the local excerpt library."
    )
    parser.add_argument(
        "input_path",
        nargs="?",
        type=Path,
        help="Optional CSV or JSON file of candidate excerpts. If omitted, reads JSON from stdin.",
    )
    parser.add_argument(
        "--db-path",
        type=Path,
        default=DEFAULT_DB_PATH,
        help=f"SQLite database path (default: {DEFAULT_DB_PATH})",
    )
    parser.add_argument("--text-column", help="Candidate excerpt text column")
    parser.add_argument("--book-column", help="Candidate book title column")
    parser.add_argument("--author-column", help="Candidate author column")
    parser.add_argument(
        "--threshold",
        type=float,
        default=0.72,
        help="Near-match threshold for non-exact matching (default: 0.72)",
    )
    parser.add_argument("--output-csv", type=Path, help="Optional CSV output path")
    return parser.parse_args()


def load_candidates_from_csv(
    input_path: Path,
    text_column: str | None,
    book_column: str | None,
    author_column: str | None,
) -> list[dict]:
    with input_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        resolved_text = resolve_column(reader.fieldnames, text_column, TEXT_COLUMN_FALLBACKS)
        if not resolved_text:
            raise ValueError("Could not find an excerpt text column automatically.")
        resolved_book = resolve_column(reader.fieldnames, book_column, BOOK_COLUMN_FALLBACKS)
        resolved_author = resolve_column(reader.fieldnames, author_column, AUTHOR_COLUMN_FALLBACKS)

        return [
            {
                "sourceRow": row_number,
                "excerptText": clean_whitespace(record.get(resolved_text, "")),
                "bookTitle": clean_whitespace(record.get(resolved_book, "")) if resolved_book else "",
                "author": clean_whitespace(record.get(resolved_author, "")) if resolved_author else "",
                "rawRecord": record,
            }
            for row_number, record in enumerate(reader, start=2)
        ]


def load_candidates_from_json_stream() -> list[dict]:
    payload = json.load(sys.stdin)
    if isinstance(payload, dict):
        records = payload.get("records", [])
    else:
        records = payload
    if not isinstance(records, list):
        raise ValueError("JSON input must be a list of records or an object with a 'records' list.")
    return [
        {
            "sourceRow": record.get("sourceRow"),
            "excerptText": clean_whitespace(record.get("excerptText", "")),
            "bookTitle": clean_whitespace(record.get("bookTitle", "")),
            "author": clean_whitespace(record.get("author", "")),
            "rawRecord": record,
        }
        for record in records
    ]


def build_match_results(candidates: list[dict], db_path: Path, threshold: float) -> dict:
    connection = connect_library(db_path)
    try:
        results = []
        for candidate in candidates:
            match = find_library_excerpt_match(
                connection,
                excerpt_text=candidate["excerptText"],
                book_title=candidate["bookTitle"],
                author=candidate["author"],
                threshold=threshold,
            )
            results.append(
                {
                    "sourceRow": candidate.get("sourceRow"),
                    "excerptText": candidate["excerptText"],
                    "bookTitle": candidate["bookTitle"],
                    "author": candidate["author"],
                    "alreadyExists": bool(match and match.get("matchType") == "exact"),
                    "hasLibraryMatch": bool(match),
                    "libraryMatch": match,
                }
            )
    finally:
        connection.close()

    exact_count = sum(1 for result in results if result["alreadyExists"])
    any_match_count = sum(1 for result in results if result["hasLibraryMatch"])
    return {
        "db_path": str(db_path),
        "total_records": len(results),
        "exact_existing_records": exact_count,
        "records_with_any_library_match": any_match_count,
        "results": results,
    }


def write_results_csv(output_csv: Path, report: dict) -> None:
    output_csv.parent.mkdir(parents=True, exist_ok=True)
    with output_csv.open("w", encoding="utf-8", newline="") as handle:
        fieldnames = [
            "sourceRow",
            "author",
            "bookTitle",
            "excerptText",
            "alreadyExists",
            "hasLibraryMatch",
            "matchType",
            "matchScore",
            "matchedSourceRow",
            "matchedAuthor",
            "matchedBookTitle",
            "matchedPoemTitle",
        ]
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for result in report["results"]:
            match = result["libraryMatch"] or {}
            writer.writerow(
                {
                    "sourceRow": result["sourceRow"] or "",
                    "author": result["author"],
                    "bookTitle": result["bookTitle"],
                    "excerptText": result["excerptText"],
                    "alreadyExists": "Y" if result["alreadyExists"] else "",
                    "hasLibraryMatch": "Y" if result["hasLibraryMatch"] else "",
                    "matchType": match.get("matchType", ""),
                    "matchScore": match.get("score", ""),
                    "matchedSourceRow": match.get("sourceRow", ""),
                    "matchedAuthor": match.get("author", ""),
                    "matchedBookTitle": match.get("bookTitle", ""),
                    "matchedPoemTitle": match.get("poemTitle", ""),
                }
            )


def main() -> int:
    args = parse_args()
    if args.input_path:
        suffix = args.input_path.suffix.lower()
        if suffix == ".csv":
            candidates = load_candidates_from_csv(
                args.input_path,
                text_column=args.text_column,
                book_column=args.book_column,
                author_column=args.author_column,
            )
        elif suffix == ".json":
            with args.input_path.open("r", encoding="utf-8") as handle:
                payload = json.load(handle)
            if isinstance(payload, dict):
                records = payload.get("records", [])
            else:
                records = payload
            candidates = [
                {
                    "sourceRow": record.get("sourceRow"),
                    "excerptText": clean_whitespace(record.get("excerptText", "")),
                    "bookTitle": clean_whitespace(record.get("bookTitle", "")),
                    "author": clean_whitespace(record.get("author", "")),
                    "rawRecord": record,
                }
                for record in records
            ]
        else:
            raise ValueError("Unsupported input file. Use CSV or JSON.")
    else:
        candidates = load_candidates_from_json_stream()

    report = build_match_results(candidates, args.db_path, threshold=args.threshold)

    if args.output_csv:
        write_results_csv(args.output_csv, report)
        print(
            json.dumps(
                {
                    "db_path": report["db_path"],
                    "total_records": report["total_records"],
                    "exact_existing_records": report["exact_existing_records"],
                    "records_with_any_library_match": report["records_with_any_library_match"],
                    "output_csv": str(args.output_csv),
                },
                indent=2,
            )
        )
    else:
        print(json.dumps(report, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
