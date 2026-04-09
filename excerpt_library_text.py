#!/usr/bin/env python3
from __future__ import annotations

import json
import sqlite3
import sys

from excerpt_library import DEFAULT_DB_PATH, build_library_status


def lookup_library_excerpt(source_row: int | None) -> dict:
    if not source_row:
        return {"ok": False, "error": "sourceRow is required."}

    if not DEFAULT_DB_PATH.exists():
        return {"ok": False, "error": "Excerpt library database is unavailable."}

    connection = sqlite3.connect(DEFAULT_DB_PATH)
    connection.row_factory = sqlite3.Row
    try:
        row = connection.execute(
            """
            SELECT source_row_number, external_id, author, book_title, poem_title, excerpt_text, word_count, metadata_json
            FROM excerpt_entries
            WHERE source_row_number = ?
            LIMIT 1
            """,
            (source_row,),
        ).fetchone()
    finally:
        connection.close()

    if not row:
        return {"ok": False, "error": "Excerpt not found in library."}

    status = build_library_status(row["metadata_json"])

    return {
        "ok": True,
        "sourceRow": row["source_row_number"],
        "recordId": row["external_id"] or "",
        "author": row["author"] or "",
        "bookTitle": row["book_title"] or "",
        "poemTitle": row["poem_title"] or "",
        "text": row["excerpt_text"],
        "wordCount": row["word_count"],
        "libraryStatus": status,
    }


def main() -> int:
    payload = json.load(sys.stdin)
    result = lookup_library_excerpt(payload.get("sourceRow"))
    json.dump(result, sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
