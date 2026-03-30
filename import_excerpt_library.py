#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from excerpt_library import DEFAULT_DB_PATH, import_csv_to_library


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Import a large excerpt spreadsheet into a normalized SQLite excerpt library."
    )
    parser.add_argument("input_csv", type=Path, help="Path to the source CSV export")
    parser.add_argument(
        "--db-path",
        type=Path,
        default=DEFAULT_DB_PATH,
        help=f"SQLite database path (default: {DEFAULT_DB_PATH})",
    )
    parser.add_argument("--source-name", help="Human-readable source name for this import")
    parser.add_argument(
        "--source-kind",
        default="spreadsheet",
        help="Source kind label to store with the import (default: spreadsheet)",
    )
    parser.add_argument("--text-column", help="CSV column containing the excerpt text")
    parser.add_argument("--book-column", help="CSV column containing the book title")
    parser.add_argument("--title-column", help="CSV column containing the poem/title field")
    parser.add_argument("--author-column", help="CSV column containing the author")
    parser.add_argument("--id-column", help="CSV column containing the row/external id")
    parser.add_argument(
        "--replace-source",
        action="store_true",
        help="Delete any previously imported rows for the same source before importing",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    result = import_csv_to_library(
        csv_path=args.input_csv,
        db_path=args.db_path,
        source_name=args.source_name,
        source_kind=args.source_kind,
        text_column=args.text_column,
        book_column=args.book_column,
        title_column=args.title_column,
        author_column=args.author_column,
        id_column=args.id_column,
        replace_source=args.replace_source,
    )
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
