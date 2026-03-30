#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path

from excerpt_library import (
    AUTHOR_COLUMN_FALLBACKS,
    BOOK_COLUMN_FALLBACKS,
    DEFAULT_DB_PATH,
    TEXT_COLUMN_FALLBACKS,
    TITLE_COLUMN_FALLBACKS,
    clean_whitespace,
    connect_library,
    find_library_excerpt_match,
    resolve_column,
)

REVIEW_DECISION_COLUMN_FALLBACKS = (
    "excerpt_review_decision",
    "Excerpt Review Decision",
    "reviewDecision",
)

APPROVED_FOR_QUOTE_COLUMN_FALLBACKS = (
    "approved_for_quote",
    "Approved for Quote Creation? Y / N",
    "Approved for Quote Creation",
    "approved",
)

EXCLUDE_COLUMN_FALLBACKS = (
    "exclude_from_quote_db",
    "Exclude from Quote DB",
    "exclude",
)

RECORD_ID_COLUMN_FALLBACKS = (
    "recordId",
    "Record ID",
    "ID",
    "id",
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Find accepted excerpts from a sheet export and classify them as already present in "
            "the excerpt library or as new candidates to import."
        )
    )
    parser.add_argument("input_csv", type=Path, help="CSV export of the working sheet")
    parser.add_argument(
        "--db-path",
        type=Path,
        default=DEFAULT_DB_PATH,
        help=f"Excerpt library SQLite path (default: {DEFAULT_DB_PATH})",
    )
    parser.add_argument("--text-column", help="Column containing the excerpt text")
    parser.add_argument("--book-column", help="Column containing the book title")
    parser.add_argument("--title-column", help="Column containing the poem title")
    parser.add_argument("--author-column", help="Column containing the author")
    parser.add_argument("--review-decision-column", help="Column containing excerpt_review_decision")
    parser.add_argument("--approved-column", help="Column containing approved_for_quote / legacy approval")
    parser.add_argument("--exclude-column", help="Column containing exclude_from_quote_db")
    parser.add_argument("--record-id-column", help="Column containing a record ID")
    parser.add_argument(
        "--threshold",
        type=float,
        default=0.72,
        help="Near-match threshold for non-exact library matching (default: 0.72)",
    )
    parser.add_argument("--output-csv", type=Path, help="Optional CSV output path")
    parser.add_argument(
        "--new-candidates-csv",
        type=Path,
        help="Optional CSV output path containing only excerpts that are not already in the library",
    )
    return parser.parse_args()


def normalize_flag(value: str | None) -> str:
    return clean_whitespace(value).upper()


def load_accepted_candidates(args: argparse.Namespace) -> dict:
    with args.input_csv.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        fieldnames = reader.fieldnames

        text_column = resolve_column(fieldnames, args.text_column, TEXT_COLUMN_FALLBACKS)
        if not text_column:
            raise ValueError("Could not find an excerpt text column automatically.")

        book_column = resolve_column(fieldnames, args.book_column, BOOK_COLUMN_FALLBACKS)
        title_column = resolve_column(fieldnames, args.title_column, TITLE_COLUMN_FALLBACKS)
        author_column = resolve_column(fieldnames, args.author_column, AUTHOR_COLUMN_FALLBACKS)
        review_decision_column = resolve_column(
            fieldnames,
            args.review_decision_column,
            REVIEW_DECISION_COLUMN_FALLBACKS,
        )
        approved_column = resolve_column(
            fieldnames,
            args.approved_column,
            APPROVED_FOR_QUOTE_COLUMN_FALLBACKS,
        )
        exclude_column = resolve_column(
            fieldnames,
            args.exclude_column,
            EXCLUDE_COLUMN_FALLBACKS,
        )
        record_id_column = resolve_column(
            fieldnames,
            args.record_id_column,
            RECORD_ID_COLUMN_FALLBACKS,
        )

        candidates: list[dict] = []
        skipped_blank_excerpt = 0
        skipped_excluded = 0
        skipped_not_accepted = 0

        for row_number, record in enumerate(reader, start=2):
            excerpt_text = record.get(text_column, "") if text_column else ""
            if not clean_whitespace(excerpt_text):
                skipped_blank_excerpt += 1
                continue

            excluded = normalize_flag(record.get(exclude_column, "")) == "Y" if exclude_column else False
            if excluded:
                skipped_excluded += 1
                continue

            explicit_decision = normalize_flag(record.get(review_decision_column, "")) if review_decision_column else ""
            approved_for_quote = normalize_flag(record.get(approved_column, "")) if approved_column else ""

            acceptance_source = ""
            if explicit_decision:
                if explicit_decision == "ACCEPT":
                    acceptance_source = "weaver_accept"
                else:
                    skipped_not_accepted += 1
                    continue
            elif approved_for_quote == "Y":
                acceptance_source = "legacy_qi_approved"
            else:
                skipped_not_accepted += 1
                continue

            candidates.append(
                {
                    "sourceRow": row_number,
                    "recordId": clean_whitespace(record.get(record_id_column, "")) if record_id_column else "",
                    "author": clean_whitespace(record.get(author_column, "")) if author_column else "",
                    "bookTitle": clean_whitespace(record.get(book_column, "")) if book_column else "",
                    "poemTitle": clean_whitespace(record.get(title_column, "")) if title_column else "",
                    "excerptText": excerpt_text,
                    "reviewDecision": explicit_decision,
                    "approvedForQuote": approved_for_quote,
                    "acceptanceSource": acceptance_source,
                    "rawRecord": record,
                }
            )

    return {
        "candidates": candidates,
        "resolvedColumns": {
            "text": text_column,
            "book": book_column,
            "title": title_column,
            "author": author_column,
            "reviewDecision": review_decision_column,
            "approvedForQuote": approved_column,
            "exclude": exclude_column,
            "recordId": record_id_column,
        },
        "skipped": {
            "blank_excerpt": skipped_blank_excerpt,
            "excluded": skipped_excluded,
            "not_accepted": skipped_not_accepted,
        },
    }


def build_report(args: argparse.Namespace) -> dict:
    loaded = load_accepted_candidates(args)
    connection = connect_library(args.db_path)

    try:
        results = []
        for candidate in loaded["candidates"]:
            match = find_library_excerpt_match(
                connection,
                excerpt_text=candidate["excerptText"],
                book_title=candidate["bookTitle"],
                author=candidate["author"],
                threshold=args.threshold,
            )
            already_in_library = bool(match and match.get("matchType") == "exact")
            results.append(
                {
                    "sourceRow": candidate["sourceRow"],
                    "recordId": candidate["recordId"],
                    "author": candidate["author"],
                    "bookTitle": candidate["bookTitle"],
                    "poemTitle": candidate["poemTitle"],
                    "excerptText": candidate["excerptText"],
                    "reviewDecision": candidate["reviewDecision"],
                    "approvedForQuote": candidate["approvedForQuote"],
                    "acceptanceSource": candidate["acceptanceSource"],
                    "alreadyInLibrary": already_in_library,
                    "hasLibraryMatch": bool(match),
                    "libraryDisposition": "already_in_library" if already_in_library else "new_candidate",
                    "libraryMatch": match,
                }
            )
    finally:
        connection.close()

    exact_count = sum(1 for row in results if row["alreadyInLibrary"])
    any_match_count = sum(1 for row in results if row["hasLibraryMatch"])
    new_count = sum(1 for row in results if row["libraryDisposition"] == "new_candidate")
    by_source: dict[str, int] = {}
    for row in results:
        by_source[row["acceptanceSource"]] = by_source.get(row["acceptanceSource"], 0) + 1

    return {
        "input_csv": str(args.input_csv),
        "db_path": str(args.db_path),
        "resolved_columns": loaded["resolvedColumns"],
        "skipped": loaded["skipped"],
        "total_accepted_candidates": len(results),
        "already_in_library": exact_count,
        "records_with_any_library_match": any_match_count,
        "new_candidates": new_count,
        "acceptance_source_counts": by_source,
        "results": results,
    }


def write_results_csv(path: Path, results: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        fieldnames = [
            "sourceRow",
            "recordId",
            "author",
            "bookTitle",
            "poemTitle",
            "reviewDecision",
            "approvedForQuote",
            "acceptanceSource",
            "libraryDisposition",
            "alreadyInLibrary",
            "hasLibraryMatch",
            "matchType",
            "matchScore",
            "matchedSourceRow",
            "matchedRecordId",
            "matchedAuthor",
            "matchedBookTitle",
            "matchedPoemTitle",
            "excerptText",
        ]
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for result in results:
            match = result["libraryMatch"] or {}
            writer.writerow(
                {
                    "sourceRow": result["sourceRow"],
                    "recordId": result["recordId"],
                    "author": result["author"],
                    "bookTitle": result["bookTitle"],
                    "poemTitle": result["poemTitle"],
                    "reviewDecision": result["reviewDecision"],
                    "approvedForQuote": result["approvedForQuote"],
                    "acceptanceSource": result["acceptanceSource"],
                    "libraryDisposition": result["libraryDisposition"],
                    "alreadyInLibrary": "Y" if result["alreadyInLibrary"] else "",
                    "hasLibraryMatch": "Y" if result["hasLibraryMatch"] else "",
                    "matchType": match.get("matchType", ""),
                    "matchScore": match.get("score", ""),
                    "matchedSourceRow": match.get("sourceRow", ""),
                    "matchedRecordId": match.get("recordId", ""),
                    "matchedAuthor": match.get("author", ""),
                    "matchedBookTitle": match.get("bookTitle", ""),
                    "matchedPoemTitle": match.get("poemTitle", ""),
                    "excerptText": result["excerptText"],
                }
            )


def main() -> int:
    args = parse_args()
    report = build_report(args)

    if args.output_csv:
        write_results_csv(args.output_csv, report["results"])

    if args.new_candidates_csv:
        write_results_csv(
            args.new_candidates_csv,
            [row for row in report["results"] if row["libraryDisposition"] == "new_candidate"],
        )

    if args.output_csv or args.new_candidates_csv:
        summary = {
            "input_csv": report["input_csv"],
            "db_path": report["db_path"],
            "resolved_columns": report["resolved_columns"],
            "skipped": report["skipped"],
            "total_accepted_candidates": report["total_accepted_candidates"],
            "already_in_library": report["already_in_library"],
            "records_with_any_library_match": report["records_with_any_library_match"],
            "new_candidates": report["new_candidates"],
            "acceptance_source_counts": report["acceptance_source_counts"],
        }
        if args.output_csv:
            summary["output_csv"] = str(args.output_csv)
        if args.new_candidates_csv:
            summary["new_candidates_csv"] = str(args.new_candidates_csv)
        print(json.dumps(summary, indent=2))
    else:
        print(json.dumps(report, indent=2))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
