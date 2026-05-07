#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sqlite3
from collections import Counter
from pathlib import Path

from excerpt_library import DEFAULT_DB_PATH, clean_whitespace, connect_library, normalize_lookup_text

DEFAULT_OUTPUT_DB_PATH = Path(__file__).resolve().parent / "data" / "excerpt_library_deduped.db"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Build a deduped excerpt database with canonical excerpts, pull counts, "
            "and preserved occurrence-level metadata."
        )
    )
    parser.add_argument(
        "--source-db",
        type=Path,
        default=DEFAULT_DB_PATH,
        help=f"Source SQLite database path (default: {DEFAULT_DB_PATH})",
    )
    parser.add_argument(
        "--output-db",
        type=Path,
        default=DEFAULT_OUTPUT_DB_PATH,
        help=f"Output SQLite database path (default: {DEFAULT_OUTPUT_DB_PATH})",
    )
    return parser.parse_args()


def ensure_output_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        PRAGMA journal_mode = WAL;

        CREATE TABLE IF NOT EXISTS canonical_excerpts (
            id INTEGER PRIMARY KEY,
            excerpt_hash TEXT NOT NULL UNIQUE,
            canonical_excerpt_text TEXT NOT NULL,
            normalized_excerpt TEXT NOT NULL,
            pull_count INTEGER NOT NULL,
            primary_author TEXT,
            primary_book_title TEXT,
            primary_poem_title TEXT,
            author_variant_count INTEGER NOT NULL,
            book_variant_count INTEGER NOT NULL,
            poem_variant_count INTEGER NOT NULL,
            excerpt_text_variant_count INTEGER NOT NULL,
            has_author_conflict INTEGER NOT NULL,
            has_book_conflict INTEGER NOT NULL,
            has_poem_title_conflict INTEGER NOT NULL,
            author_values_json TEXT NOT NULL,
            book_values_json TEXT NOT NULL,
            poem_title_values_json TEXT NOT NULL,
            excerpt_text_values_json TEXT NOT NULL,
            representative_source_row INTEGER,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        CREATE INDEX IF NOT EXISTS idx_canonical_excerpts_pull_count
            ON canonical_excerpts(pull_count DESC);
        CREATE INDEX IF NOT EXISTS idx_canonical_excerpts_author
            ON canonical_excerpts(primary_author);
        CREATE INDEX IF NOT EXISTS idx_canonical_excerpts_book
            ON canonical_excerpts(primary_book_title);

        CREATE TABLE IF NOT EXISTS canonical_excerpt_occurrences (
            id INTEGER PRIMARY KEY,
            canonical_excerpt_id INTEGER NOT NULL REFERENCES canonical_excerpts(id) ON DELETE CASCADE,
            source_entry_id INTEGER NOT NULL,
            source_id INTEGER NOT NULL,
            source_row_number INTEGER NOT NULL,
            external_id TEXT,
            author TEXT,
            book_title TEXT,
            poem_title TEXT,
            excerpt_text TEXT NOT NULL,
            normalized_author TEXT,
            normalized_book_title TEXT,
            normalized_poem_title TEXT,
            is_representative INTEGER NOT NULL,
            metadata_json TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_canonical_occurrences_canonical_id
            ON canonical_excerpt_occurrences(canonical_excerpt_id);
        CREATE INDEX IF NOT EXISTS idx_canonical_occurrences_source_row
            ON canonical_excerpt_occurrences(source_row_number);

        CREATE VIEW IF NOT EXISTS v_canonical_excerpt_summary AS
        SELECT
            id,
            excerpt_hash,
            canonical_excerpt_text,
            pull_count,
            primary_author,
            primary_book_title,
            primary_poem_title,
            has_author_conflict,
            has_book_conflict,
            has_poem_title_conflict
        FROM canonical_excerpts;
        """
    )


def choose_preferred_value(values: list[str], rows: list[sqlite3.Row], field_name: str) -> str:
    normalized_counter: Counter[str] = Counter()
    display_by_normalized: dict[str, str] = {}

    for row in rows:
        value = clean_whitespace(row[field_name])
        normalized = normalize_lookup_text(value)
        if not normalized:
            continue
        normalized_counter[normalized] += 1
        existing_display = display_by_normalized.get(normalized, "")
        if not existing_display or len(value) > len(existing_display):
            display_by_normalized[normalized] = value

    if normalized_counter:
        best_normalized = max(
            normalized_counter,
            key=lambda key: (
                normalized_counter[key],
                len(display_by_normalized.get(key, "")),
                display_by_normalized.get(key, "").lower(),
            ),
        )
        return display_by_normalized[best_normalized]

    for value in values:
        cleaned = clean_whitespace(value)
        if cleaned:
            return cleaned
    return ""


def choose_representative_row(rows: list[sqlite3.Row]) -> sqlite3.Row:
    def row_score(row: sqlite3.Row) -> tuple[int, int, int, int, int]:
        excerpt_text = clean_whitespace(row["excerpt_text"])
        return (
            1 if clean_whitespace(row["book_title"]) else 0,
            1 if clean_whitespace(row["poem_title"]) else 0,
            1 if clean_whitespace(row["author"]) else 0,
            len(excerpt_text),
            -int(row["source_row_number"]),
        )

    return max(rows, key=row_score)


def collect_distinct_display_values(rows: list[sqlite3.Row], field_name: str) -> list[str]:
    seen: set[str] = set()
    values: list[str] = []
    for row in rows:
        value = clean_whitespace(row[field_name])
        if not value:
            continue
        normalized = normalize_lookup_text(value)
        if normalized in seen:
            continue
        seen.add(normalized)
        values.append(value)
    return values


def build_deduped_database(source_db: Path, output_db: Path) -> dict:
    if output_db.exists():
        output_db.unlink()

    source_connection = connect_library(source_db)
    output_db.parent.mkdir(parents=True, exist_ok=True)
    output_connection = sqlite3.connect(output_db)
    try:
        output_connection.row_factory = sqlite3.Row
        ensure_output_schema(output_connection)

        groups = source_connection.execute(
            """
            SELECT excerpt_hash
            FROM excerpt_entries
            GROUP BY excerpt_hash
            ORDER BY excerpt_hash
            """
        ).fetchall()

        canonical_count = 0
        merged_occurrence_count = 0
        conflict_count = 0

        with output_connection:
            for group in groups:
                excerpt_hash = group["excerpt_hash"]
                rows = source_connection.execute(
                    """
                    SELECT *
                    FROM excerpt_entries
                    WHERE excerpt_hash = ?
                    ORDER BY source_row_number, id
                    """,
                    (excerpt_hash,),
                ).fetchall()
                if not rows:
                    continue

                representative = choose_representative_row(rows)
                author_values = collect_distinct_display_values(rows, "author")
                book_values = collect_distinct_display_values(rows, "book_title")
                poem_title_values = collect_distinct_display_values(rows, "poem_title")
                excerpt_text_values = collect_distinct_display_values(rows, "excerpt_text")

                primary_author = choose_preferred_value(author_values, rows, "author")
                primary_book_title = choose_preferred_value(book_values, rows, "book_title")
                primary_poem_title = choose_preferred_value(poem_title_values, rows, "poem_title")
                canonical_excerpt_text = choose_preferred_value(excerpt_text_values, rows, "excerpt_text")
                normalized_excerpt = normalize_lookup_text(canonical_excerpt_text or representative["excerpt_text"])

                has_author_conflict = 1 if len(author_values) > 1 else 0
                has_book_conflict = 1 if len(book_values) > 1 else 0
                has_poem_title_conflict = 1 if len(poem_title_values) > 1 else 0
                if has_author_conflict or has_book_conflict or has_poem_title_conflict:
                    conflict_count += 1

                cursor = output_connection.execute(
                    """
                    INSERT INTO canonical_excerpts (
                        excerpt_hash,
                        canonical_excerpt_text,
                        normalized_excerpt,
                        pull_count,
                        primary_author,
                        primary_book_title,
                        primary_poem_title,
                        author_variant_count,
                        book_variant_count,
                        poem_variant_count,
                        excerpt_text_variant_count,
                        has_author_conflict,
                        has_book_conflict,
                        has_poem_title_conflict,
                        author_values_json,
                        book_values_json,
                        poem_title_values_json,
                        excerpt_text_values_json,
                        representative_source_row
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        excerpt_hash,
                        canonical_excerpt_text or representative["excerpt_text"],
                        normalized_excerpt,
                        len(rows),
                        primary_author,
                        primary_book_title,
                        primary_poem_title,
                        len(author_values),
                        len(book_values),
                        len(poem_title_values),
                        len(excerpt_text_values),
                        has_author_conflict,
                        has_book_conflict,
                        has_poem_title_conflict,
                        json.dumps(author_values, ensure_ascii=True),
                        json.dumps(book_values, ensure_ascii=True),
                        json.dumps(poem_title_values, ensure_ascii=True),
                        json.dumps(excerpt_text_values, ensure_ascii=True),
                        int(representative["source_row_number"]),
                    ),
                )
                canonical_id = cursor.lastrowid
                canonical_count += 1

                for row in rows:
                    output_connection.execute(
                        """
                        INSERT INTO canonical_excerpt_occurrences (
                            canonical_excerpt_id,
                            source_entry_id,
                            source_id,
                            source_row_number,
                            external_id,
                            author,
                            book_title,
                            poem_title,
                            excerpt_text,
                            normalized_author,
                            normalized_book_title,
                            normalized_poem_title,
                            is_representative,
                            metadata_json
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            canonical_id,
                            int(row["id"]),
                            int(row["source_id"]),
                            int(row["source_row_number"]),
                            row["external_id"] or "",
                            row["author"] or "",
                            row["book_title"] or "",
                            row["poem_title"] or "",
                            row["excerpt_text"],
                            row["normalized_author"] or "",
                            row["normalized_book_title"] or "",
                            row["normalized_poem_title"] or "",
                            1 if int(row["id"]) == int(representative["id"]) else 0,
                            row["metadata_json"],
                        ),
                    )
                    merged_occurrence_count += 1
    finally:
        output_connection.close()
        source_connection.close()

    return {
        "source_db": str(source_db),
        "output_db": str(output_db),
        "canonical_excerpt_count": canonical_count,
        "merged_occurrence_count": merged_occurrence_count,
        "duplicate_rows_removed": merged_occurrence_count - canonical_count,
        "canonical_groups_with_metadata_conflicts": conflict_count,
    }


def main() -> int:
    args = parse_args()
    result = build_deduped_database(args.source_db, args.output_db)
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
