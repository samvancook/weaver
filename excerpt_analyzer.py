#!/usr/bin/env python3
"""Annotate excerpt CSV rows with duplicate and overlap signals."""

from __future__ import annotations

import argparse
import csv
import re
from collections import defaultdict
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Iterable


QUOTE_COLUMN_FALLBACKS = (
    "Enter the Quote",
    "Quote",
    "Excerpt",
    "Text",
)


@dataclass
class ExcerptRow:
    row_number: int
    raw_text: str
    normalized_text: str
    tokens: set[str]
    token_count: int
    char_count: int


@dataclass
class MatchResult:
    other_row_number: int
    match_type: str
    score: float
    shared_token_count: int
    other_excerpt: str


def clean_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def normalize_text(text: str) -> str:
    text = text.lower()
    text = text.replace("—", " ")
    text = text.replace("–", " ")
    text = re.sub(r"[\"'“”‘’`.,!?;:()[\]{}]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def tokenize(text: str) -> list[str]:
    return re.findall(r"[a-z0-9]+", normalize_text(text))


def find_quote_column(fieldnames: Iterable[str] | None, preferred: str | None) -> str:
    if not fieldnames:
        raise ValueError("The CSV file does not have a header row.")

    available = list(fieldnames)
    if preferred:
        if preferred not in available:
            raise ValueError(
                f'Column "{preferred}" was not found. Available columns: {", ".join(available)}'
            )
        return preferred

    for candidate in QUOTE_COLUMN_FALLBACKS:
        if candidate in available:
            return candidate

    raise ValueError(
        "Could not find an excerpt column automatically. "
        f"Available columns: {', '.join(available)}"
    )


def sequence_overlap(a: str, b: str) -> float:
    return SequenceMatcher(None, normalize_text(a), normalize_text(b)).ratio()


def token_overlap(a: set[str], b: set[str]) -> tuple[float, int]:
    if not a or not b:
        return 0.0, 0
    intersection = a & b
    union = a | b
    return len(intersection) / len(union), len(intersection)


def build_rows(reader: csv.DictReader, quote_column: str) -> list[tuple[dict[str, str], ExcerptRow]]:
    rows: list[tuple[dict[str, str], ExcerptRow]] = []
    for index, record in enumerate(reader, start=2):
        raw_text = record.get(quote_column, "") or ""
        cleaned = clean_whitespace(raw_text)
        tokens = set(tokenize(cleaned))
        excerpt = ExcerptRow(
            row_number=index,
            raw_text=cleaned,
            normalized_text=normalize_text(cleaned),
            tokens=tokens,
            token_count=len(cleaned.split()) if cleaned else 0,
            char_count=len(cleaned),
        )
        rows.append((record, excerpt))
    return rows


def compute_matches(
    rows: list[tuple[dict[str, str], ExcerptRow]],
    near_duplicate_threshold: float,
) -> dict[int, MatchResult]:
    best_matches: dict[int, MatchResult] = {}
    exact_duplicates: defaultdict[str, list[ExcerptRow]] = defaultdict(list)

    for _, excerpt in rows:
        if excerpt.normalized_text:
            exact_duplicates[excerpt.normalized_text].append(excerpt)

    for matches in exact_duplicates.values():
        if len(matches) < 2:
            continue
        anchor = matches[0]
        for excerpt in matches:
            other = anchor if excerpt.row_number != anchor.row_number else matches[1]
            best_matches[excerpt.row_number] = MatchResult(
                other_row_number=other.row_number,
                match_type="exact_duplicate",
                score=1.0,
                shared_token_count=len(excerpt.tokens),
                other_excerpt=other.raw_text,
            )

    excerpts = [excerpt for _, excerpt in rows]
    for index, current in enumerate(excerpts):
        if not current.raw_text:
            continue

        for other in excerpts[index + 1 :]:
            if not other.raw_text or current.row_number == other.row_number:
                continue

            token_score, shared_token_count = token_overlap(current.tokens, other.tokens)
            sequence_score = sequence_overlap(current.raw_text, other.raw_text)
            combined_score = max(token_score, sequence_score)

            if combined_score < near_duplicate_threshold:
                continue

            if current.normalized_text == other.normalized_text:
                continue

            left_result = MatchResult(
                other_row_number=other.row_number,
                match_type="high_overlap",
                score=combined_score,
                shared_token_count=shared_token_count,
                other_excerpt=other.raw_text,
            )
            right_result = MatchResult(
                other_row_number=current.row_number,
                match_type="high_overlap",
                score=combined_score,
                shared_token_count=shared_token_count,
                other_excerpt=current.raw_text,
            )

            if combined_score > best_matches.get(current.row_number, left_result).score:
                best_matches[current.row_number] = left_result
            if combined_score > best_matches.get(other.row_number, right_result).score:
                best_matches[other.row_number] = right_result

    return best_matches


def annotate_csv(
    input_path: Path,
    output_path: Path,
    quote_column: str | None,
    near_duplicate_threshold: float,
) -> None:
    with input_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        resolved_quote_column = find_quote_column(reader.fieldnames, quote_column)
        rows = build_rows(reader, resolved_quote_column)

    matches = compute_matches(rows, near_duplicate_threshold)
    fieldnames = list(rows[0][0].keys()) if rows else [resolved_quote_column]
    output_fields = fieldnames + [
        "excerpt_word_count",
        "excerpt_character_count",
        "duplicate_status",
        "matched_row_number",
        "overlap_score",
        "shared_token_count",
        "matched_excerpt_preview",
    ]

    with output_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=output_fields)
        writer.writeheader()

        for record, excerpt in rows:
            match = matches.get(excerpt.row_number)
            enriched_record = dict(record)
            enriched_record["excerpt_word_count"] = str(excerpt.token_count)
            enriched_record["excerpt_character_count"] = str(excerpt.char_count)
            enriched_record["duplicate_status"] = match.match_type if match else ""
            enriched_record["matched_row_number"] = str(match.other_row_number) if match else ""
            enriched_record["overlap_score"] = f"{match.score:.3f}" if match else ""
            enriched_record["shared_token_count"] = str(match.shared_token_count) if match else ""
            enriched_record["matched_excerpt_preview"] = (
                match.other_excerpt[:160] if match else ""
            )
            writer.writerow(enriched_record)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Read a CSV of excerpts and add columns for exact duplicates "
            "and strong overlap."
        )
    )
    parser.add_argument("input_csv", type=Path, help="Path to the source CSV file")
    parser.add_argument("output_csv", type=Path, help="Path for the enriched CSV output")
    parser.add_argument(
        "--quote-column",
        help="Name of the column that contains excerpt text",
    )
    parser.add_argument(
        "--near-duplicate-threshold",
        type=float,
        default=0.72,
        help="Minimum similarity score for high-overlap matching (default: 0.72)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    annotate_csv(
        input_path=args.input_csv,
        output_path=args.output_csv,
        quote_column=args.quote_column,
        near_duplicate_threshold=args.near_duplicate_threshold,
    )


if __name__ == "__main__":
    main()
