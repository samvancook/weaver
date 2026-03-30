#!/usr/bin/env python3
from __future__ import annotations

import sqlite3
from pathlib import Path

SOURCE_DB_PATH = Path(__file__).resolve().parent / "data" / "excerpt_library.db"
RUNTIME_DB_PATH = Path(__file__).resolve().parent / "data" / "excerpt_library_runtime.db"


def build_runtime_db(source_db_path: Path = SOURCE_DB_PATH, runtime_db_path: Path = RUNTIME_DB_PATH) -> dict:
    if not source_db_path.exists():
        raise FileNotFoundError(f"Source excerpt library not found: {source_db_path}")

    runtime_db_path.parent.mkdir(parents=True, exist_ok=True)
    if runtime_db_path.exists():
        runtime_db_path.unlink()

    connection = sqlite3.connect(runtime_db_path)
    try:
        connection.executescript(
            f"""
            ATTACH DATABASE '{source_db_path.as_posix()}' AS src;

            CREATE TABLE excerpt_entries (
                id INTEGER PRIMARY KEY,
                source_row_number INTEGER NOT NULL,
                external_id TEXT,
                author TEXT,
                normalized_author TEXT,
                book_title TEXT,
                normalized_book_title TEXT,
                poem_title TEXT,
                normalized_poem_title TEXT,
                excerpt_text TEXT NOT NULL,
                normalized_excerpt TEXT NOT NULL,
                excerpt_hash TEXT NOT NULL,
                word_count INTEGER NOT NULL,
                character_count INTEGER NOT NULL
            );

            CREATE INDEX idx_excerpt_entries_hash
                ON excerpt_entries(excerpt_hash);
            CREATE INDEX idx_excerpt_entries_book
                ON excerpt_entries(normalized_book_title);
            CREATE INDEX idx_excerpt_entries_author
                ON excerpt_entries(normalized_author);
            CREATE INDEX idx_excerpt_entries_poem
                ON excerpt_entries(normalized_poem_title);
            CREATE INDEX idx_excerpt_entries_chars
                ON excerpt_entries(character_count);

            INSERT INTO excerpt_entries (
                source_row_number,
                external_id,
                author,
                normalized_author,
                book_title,
                normalized_book_title,
                poem_title,
                normalized_poem_title,
                excerpt_text,
                normalized_excerpt,
                excerpt_hash,
                word_count,
                character_count
            )
            SELECT
                MIN(source_row_number),
                MAX(external_id),
                COALESCE(MAX(author), ''),
                COALESCE(MAX(normalized_author), ''),
                COALESCE(MAX(book_title), ''),
                COALESCE(MAX(normalized_book_title), ''),
                COALESCE(MAX(poem_title), ''),
                COALESCE(MAX(normalized_poem_title), ''),
                MAX(excerpt_text),
                normalized_excerpt,
                excerpt_hash,
                MAX(word_count),
                MAX(character_count)
            FROM src.excerpt_entries
            GROUP BY excerpt_hash, normalized_excerpt;
            """
        )
        connection.commit()
        row_count = connection.execute("SELECT COUNT(*) FROM excerpt_entries").fetchone()[0]
    finally:
        connection.close()

    return {
        "runtime_db_path": str(runtime_db_path),
        "row_count": int(row_count),
        "size_bytes": runtime_db_path.stat().st_size,
    }


def main() -> int:
    result = build_runtime_db()
    print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
