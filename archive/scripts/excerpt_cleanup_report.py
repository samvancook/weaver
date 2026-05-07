#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from excerpt_library import DEFAULT_DB_PATH, connect_library


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate cleanup-oriented summary reports for the excerpt library."
    )
    parser.add_argument(
        "--db-path",
        type=Path,
        default=DEFAULT_DB_PATH,
        help=f"SQLite database path (default: {DEFAULT_DB_PATH})",
    )
    parser.add_argument(
        "--top",
        type=int,
        default=25,
        help="Number of top rows to include per section (default: 25)",
    )
    return parser.parse_args()


def fetch_rows(connection, query: str, params: tuple = ()) -> list[dict]:
    return [dict(row) for row in connection.execute(query, params).fetchall()]


def build_report(db_path: Path, top: int) -> dict:
    connection = connect_library(db_path)
    try:
        overview = fetch_rows(
            connection,
            """
            SELECT
                COUNT(*) AS total_rows,
                COUNT(DISTINCT excerpt_hash) AS distinct_excerpts,
                SUM(CASE WHEN COALESCE(book_title, '') = '' THEN 1 ELSE 0 END) AS blank_book_rows,
                SUM(CASE WHEN COALESCE(poem_title, '') = '' THEN 1 ELSE 0 END) AS blank_title_rows,
                SUM(CASE WHEN length(trim(excerpt_text)) <= 25 THEN 1 ELSE 0 END) AS short_text_rows
            FROM excerpt_entries
            """,
        )[0]

        top_duplicate_groups = fetch_rows(
            connection,
            """
            SELECT
                COUNT(*) AS copies,
                author,
                book_title,
                poem_title,
                substr(excerpt_text, 1, 180) AS excerpt_preview
            FROM excerpt_entries
            GROUP BY excerpt_hash
            HAVING COUNT(*) > 1
            ORDER BY copies DESC, author, book_title, poem_title
            LIMIT ?
            """,
            (top,),
        )

        likely_non_excerpt_rows = fetch_rows(
            connection,
            """
            SELECT
                source_row_number,
                author,
                book_title,
                poem_title,
                excerpt_text
            FROM excerpt_entries
            WHERE
                (
                    excerpt_text IN ('Hi-res', 'Red background', 'Watercolor background')
                    OR lower(excerpt_text) LIKE 'interior title page%'
                    OR lower(excerpt_text) LIKE 'stack of %bookshelf%'
                    OR lower(excerpt_text) LIKE 'single vertical %'
                    OR lower(excerpt_text) LIKE '% background'
                    OR lower(excerpt_text) LIKE '% background%'
                )
                AND excerpt_text NOT LIKE '"%'
                AND excerpt_text NOT LIKE '“%'
            ORDER BY source_row_number
            LIMIT ?
            """,
            (top,),
        )

        blank_book_by_author = fetch_rows(
            connection,
            """
            SELECT
                author,
                COUNT(*) AS blank_book_rows
            FROM excerpt_entries
            WHERE COALESCE(book_title, '') = ''
            GROUP BY author
            ORDER BY blank_book_rows DESC, author
            LIMIT ?
            """,
            (top,),
        )
    finally:
        connection.close()

    return {
        "db_path": str(db_path),
        "overview": overview,
        "top_duplicate_groups": top_duplicate_groups,
        "likely_non_excerpt_rows": likely_non_excerpt_rows,
        "blank_book_by_author": blank_book_by_author,
    }


def main() -> int:
    args = parse_args()
    report = build_report(args.db_path, args.top)
    print(json.dumps(report, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
