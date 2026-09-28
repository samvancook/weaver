#!/usr/bin/env python3
"""Search every row in a local operational excerpt-library snapshot."""

from __future__ import annotations

import argparse
import csv
import json
import sqlite3
from difflib import SequenceMatcher
from pathlib import Path

from excerpt_library import normalize_lookup_text


DEFAULT_DB_PATH = Path(__file__).resolve().parent / "data" / "excerpt_library.db"


def possible_variant_score(query: str, candidate: str) -> float:
    query_tokens = query.split()
    candidate_tokens = candidate.split()
    if len(query_tokens) < 5 or not candidate_tokens:
        return 0.0
    if len(set(query_tokens) & set(candidate_tokens)) / len(set(query_tokens)) < 0.8:
        return 0.0
    length = len(query_tokens)
    best = 0.0
    for window_size in range(max(1, length - 2), min(len(candidate_tokens), length + 3) + 1):
        for start in range(len(candidate_tokens) - window_size + 1):
            score = SequenceMatcher(
                None, query_tokens, candidate_tokens[start:start + window_size]
            ).ratio()
            best = max(best, score)
    return round(best, 3) if best >= 0.86 else 0.0


def lookup(
    db_path: Path,
    quote: str,
    author: str = "",
    limit: int = 25,
    index_path: Path | None = None,
) -> dict:
    query = normalize_lookup_text(quote)
    if not query:
        raise ValueError("Quote must contain searchable letters or numbers")
    if index_path is not None and not index_path.is_file():
        raise FileNotFoundError(f"Source index not found: {index_path}")
    if index_path is None and not db_path.is_file():
        raise FileNotFoundError(f"Operational database not found: {db_path}")

    author_query = normalize_lookup_text(author)
    connection = None
    index_handle = None
    if index_path is not None:
        index_handle = index_path.open("r", encoding="utf-8", newline="")
        rows = csv.DictReader(index_handle)
    else:
        connection = sqlite3.connect(db_path.resolve().as_uri() + "?mode=ro", uri=True)
        connection.row_factory = sqlite3.Row
        rows = connection.execute(
            """
            SELECT e.id AS entryId, e.source_row_number AS sourceRow,
                   e.external_id AS sourceRecordId, e.author, e.book_title AS bookTitle,
                   e.poem_title AS poemTitle, e.excerpt_text AS excerptText,
                   e.excerpt_hash AS excerptHash, s.source_name AS sourceName,
                   s.source_kind AS sourceKind
            FROM excerpt_entries AS e
            LEFT JOIN excerpt_sources AS s ON s.id = e.source_id
            ORDER BY e.id
            """
        )
    source_row_count = 0
    try:
        matches = []
        for row in rows:
            source_row_count += 1
            if author_query and author_query not in normalize_lookup_text(row["author"]):
                continue
            candidate = normalize_lookup_text(row["excerptText"])
            if query == candidate:
                match_type = "exact_text"
            elif query in candidate:
                match_type = "quote_within_excerpt"
            else:
                score = possible_variant_score(query, candidate)
                if not score:
                    continue
                match_type = "possible_variant"
            matches.append({
                "matchType": match_type,
                "similarity": score if match_type == "possible_variant" else 1.0,
                "entryId": int(row["entryId"]),
                "sourceName": row["sourceName"] or "",
                "sourceKind": row["sourceKind"] or "",
                "sourceRow": int(row["sourceRow"]),
                "sourceRecordId": row["sourceRecordId"] or "",
                "author": row["author"] or "",
                "bookTitle": row["bookTitle"] or "",
                "poemTitle": row["poemTitle"] or "",
                "excerptText": row["excerptText"],
                "excerptHash": row["excerptHash"],
            })
    finally:
        if connection is not None:
            connection.close()
        if index_handle is not None:
            index_handle.close()

    rank = {"exact_text": 0, "quote_within_excerpt": 1, "possible_variant": 2}
    matches.sort(key=lambda row: (rank[row["matchType"]], -row["similarity"], row["entryId"]))
    return {
        "sourcePath": str((index_path or db_path).resolve()),
        "sourceType": "csv_index" if index_path is not None else "sqlite",
        "scannedRowCount": source_row_count,
        "query": quote,
        "authorFilter": author,
        "matchTypes": ["exact_text", "quote_within_excerpt", "possible_variant"],
        "totalMatches": len(matches),
        "returnedMatches": min(limit, len(matches)),
        "truncated": len(matches) > limit,
        "matches": matches[:limit],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("quote", help="Full excerpt or quote fragment to find")
    parser.add_argument("--author", default="", help="Optional author filter")
    parser.add_argument("--db-path", type=Path, default=DEFAULT_DB_PATH)
    parser.add_argument("--index-path", type=Path, help="Use the exported operational_excerpt_index.csv")
    parser.add_argument("--limit", type=int, default=25)
    args = parser.parse_args()
    if args.limit < 1:
        parser.error("--limit must be at least 1")
    try:
        result = lookup(args.db_path, args.quote, args.author, args.limit, args.index_path)
    except (ValueError, FileNotFoundError) as error:
        parser.error(str(error))
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
