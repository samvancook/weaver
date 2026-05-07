#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import re
from pathlib import Path

from excerpt_library import DEFAULT_DB_PATH, connect_library, normalize_lookup_text

DEFAULT_OUTPUT_CSV = Path(__file__).resolve().parent / "data" / "excerpt_row_classification.csv"

NON_EXCERPT_EXACT_VALUES = {
    "hi-res",
    "red background",
    "watercolor background",
    "sweet, young, & worried",
    "interior title page",
}

NON_EXCERPT_PATTERNS = [
    re.compile(r"^interior title page", re.IGNORECASE),
    re.compile(r"^stack of .*bookshelf", re.IGNORECASE),
    re.compile(r"^single vertical ", re.IGNORECASE),
    re.compile(r"\bbackground\b", re.IGNORECASE),
    re.compile(r"\bhi-res\b", re.IGNORECASE),
    re.compile(r"\blo-res\b", re.IGNORECASE),
    re.compile(r"\bfront-facing\b", re.IGNORECASE),
    re.compile(r"\bzoomed in\b", re.IGNORECASE),
    re.compile(r"\bcover\b", re.IGNORECASE),
    re.compile(r"\bbookshelf\b", re.IGNORECASE),
]

EXCERPT_HINT_PATTERNS = [
    re.compile(r"^[\"“]"),
    re.compile(r"[.!?][\"”]?$"),
    re.compile(r"\b(i|you|we|they|he|she|love|body|heart|god|world|night|pain)\b", re.IGNORECASE),
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Classify excerpt library rows as likely_excerpt, likely_non_excerpt, or needs_review."
    )
    parser.add_argument(
        "--db-path",
        type=Path,
        default=DEFAULT_DB_PATH,
        help=f"SQLite database path (default: {DEFAULT_DB_PATH})",
    )
    parser.add_argument(
        "--output-csv",
        type=Path,
        default=DEFAULT_OUTPUT_CSV,
        help=f"CSV output path (default: {DEFAULT_OUTPUT_CSV})",
    )
    parser.add_argument(
        "--json-summary",
        action="store_true",
        help="Print JSON summary to stdout",
    )
    parser.add_argument(
        "--sample-size",
        type=int,
        default=10,
        help="How many sample rows to include per class in the JSON summary (default: 10)",
    )
    return parser.parse_args()


def classify_row(row: dict) -> tuple[str, float, list[str]]:
    excerpt_text = (row.get("excerpt_text") or "").strip()
    normalized_text = normalize_lookup_text(excerpt_text)
    poem_title = (row.get("poem_title") or "").strip()
    word_count = int(row.get("word_count") or 0)
    file_type = ((row.get("file_type") or "") or (row.get("file_type_helper") or "")).strip().upper()

    reasons: list[str] = []

    if not normalized_text:
        return "likely_non_excerpt", 1.0, ["blank_text"]

    if normalized_text in NON_EXCERPT_EXACT_VALUES:
        return "likely_non_excerpt", 0.99, [f"exact_non_excerpt_value:{normalized_text}"]

    non_excerpt_hits = [pattern.pattern for pattern in NON_EXCERPT_PATTERNS if pattern.search(excerpt_text)]
    quoted_excerpt = excerpt_text.startswith('"') or excerpt_text.startswith("“")
    has_terminal_punctuation = bool(re.search(r"[.!?][\"”]?$", excerpt_text))
    long_enough_for_excerpt = word_count >= 8
    has_excerpt_hint = any(pattern.search(excerpt_text) for pattern in EXCERPT_HINT_PATTERNS)

    if file_type in {"COV", "ART"}:
        reasons.append(f"content_type:{file_type}")
        if quoted_excerpt and long_enough_for_excerpt:
            reasons.append("quoted_text_overrides_cover_type")
            return "needs_review", 0.7, reasons
        return "likely_non_excerpt", 0.99, reasons

    if file_type in {"QI", "EXC"}:
        reasons.append(f"content_type:{file_type}")
        if quoted_excerpt or long_enough_for_excerpt:
            reasons.append("excerpt_or_quote_content_type")
            return "likely_excerpt", 0.98, reasons

    if file_type == "INT":
        reasons.append("content_type:INT")
        if non_excerpt_hits and not quoted_excerpt:
            reasons.extend([f"non_excerpt_pattern:{hit}" for hit in non_excerpt_hits])
            return "likely_non_excerpt", 0.97, reasons
        if quoted_excerpt and long_enough_for_excerpt:
            reasons.append("quoted_text")
            return "likely_excerpt", 0.82, reasons
        return "needs_review", 0.7, reasons

    if non_excerpt_hits and not quoted_excerpt and word_count <= 8:
        reasons.extend([f"non_excerpt_pattern:{hit}" for hit in non_excerpt_hits])
        if not poem_title:
            reasons.append("missing_poem_title")
        return "likely_non_excerpt", 0.96, reasons

    if non_excerpt_hits and not quoted_excerpt and not has_terminal_punctuation and word_count <= 14:
        reasons.extend([f"non_excerpt_pattern:{hit}" for hit in non_excerpt_hits])
        return "likely_non_excerpt", 0.9, reasons

    if quoted_excerpt and long_enough_for_excerpt:
        reasons.append("quoted_text")
        reasons.append("sufficient_length")
        return "likely_excerpt", 0.95, reasons

    if long_enough_for_excerpt and has_terminal_punctuation and has_excerpt_hint:
        reasons.append("sentence_like_text")
        reasons.append("sufficient_length")
        return "likely_excerpt", 0.88, reasons

    if word_count >= 12:
        reasons.append("long_form_text")
        return "likely_excerpt", 0.75, reasons

    if non_excerpt_hits:
        reasons.extend([f"non_excerpt_pattern:{hit}" for hit in non_excerpt_hits])
        return "needs_review", 0.55, reasons

    if word_count <= 4:
        reasons.append("very_short_text")
        return "needs_review", 0.55, reasons

    reasons.append("ambiguous_text_shape")
    return "needs_review", 0.5, reasons


def build_classification_report(db_path: Path, sample_size: int) -> dict:
    connection = connect_library(db_path)
    try:
        rows = [
            dict(row)
            for row in connection.execute(
                """
                SELECT id, source_id, source_row_number, external_id, author, book_title, poem_title,
                       excerpt_text, word_count, character_count, metadata_json
                FROM excerpt_entries
                ORDER BY source_row_number
                """
            ).fetchall()
        ]
    finally:
        connection.close()

    classified_rows: list[dict] = []
    counts = {
        "likely_excerpt": 0,
        "likely_non_excerpt": 0,
        "needs_review": 0,
    }
    samples = {
        "likely_excerpt": [],
        "likely_non_excerpt": [],
        "needs_review": [],
    }

    for row in rows:
        metadata = json.loads(row["metadata_json"] or "{}")
        file_type = (metadata.get("File Type") or "").strip()
        file_type_helper = (metadata.get("File Type Helper") or "").strip()
        classification, confidence, reasons = classify_row(
            {
                **row,
                "file_type": file_type,
                "file_type_helper": file_type_helper,
            }
        )
        record = {
            "sourceEntryId": row["id"],
            "sourceId": row["source_id"],
            "sourceRow": row["source_row_number"],
            "externalId": row["external_id"] or "",
            "author": row["author"] or "",
            "bookTitle": row["book_title"] or "",
            "poemTitle": row["poem_title"] or "",
            "excerptText": row["excerpt_text"],
            "wordCount": row["word_count"],
            "characterCount": row["character_count"],
            "fileType": file_type,
            "fileTypeHelper": file_type_helper,
            "classification": classification,
            "confidence": round(confidence, 3),
            "reasons": reasons,
        }
        classified_rows.append(record)
        counts[classification] += 1
        if len(samples[classification]) < sample_size:
            samples[classification].append(
                {
                    "sourceRow": record["sourceRow"],
                    "author": record["author"],
                    "bookTitle": record["bookTitle"],
                    "poemTitle": record["poemTitle"],
                    "fileType": record["fileType"],
                    "fileTypeHelper": record["fileTypeHelper"],
                    "excerptText": record["excerptText"],
                    "reasons": record["reasons"],
                }
            )

    return {
        "db_path": str(db_path),
        "counts": counts,
        "samples": samples,
        "rows": classified_rows,
    }


def write_classification_csv(output_csv: Path, report: dict) -> None:
    output_csv.parent.mkdir(parents=True, exist_ok=True)
    with output_csv.open("w", encoding="utf-8", newline="") as handle:
        fieldnames = [
            "sourceEntryId",
            "sourceId",
            "sourceRow",
            "externalId",
            "author",
            "bookTitle",
            "poemTitle",
            "excerptText",
            "wordCount",
            "characterCount",
            "fileType",
            "fileTypeHelper",
            "classification",
            "confidence",
            "reasons",
        ]
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in report["rows"]:
            writer.writerow(
                {
                    **{key: row[key] for key in fieldnames if key != "reasons"},
                    "reasons": "; ".join(row["reasons"]),
                }
            )


def main() -> int:
    args = parse_args()
    report = build_classification_report(args.db_path, args.sample_size)
    write_classification_csv(args.output_csv, report)

    summary = {
        "db_path": report["db_path"],
        "output_csv": str(args.output_csv),
        "counts": report["counts"],
    }
    if args.json_summary:
        summary["samples"] = report["samples"]
    print(json.dumps(summary, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
