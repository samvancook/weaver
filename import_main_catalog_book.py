#!/usr/bin/env python3
"""Import one verified main-Catalog book response into a Weaver catalog copy.

Input is the JSON returned by GET /books/<book>?include_poems=true&poem_limit=<n>.
The source response is kept outside the repository; only the validated rows are
written to the requested SQLite catalog. Existing books are never overwritten.
"""

from __future__ import annotations

import argparse
import json
import sqlite3
from pathlib import Path

from catalog_validate import DEFAULT_DB_PATH


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def import_book_payload(payload: dict, database: Path) -> dict:
    status = payload["status"]
    books = payload["catalog_books"]
    poems = payload["poems"]
    require(len(books) == 1, "Expected exactly one catalog edition")
    book = books[0]
    book_id = int(status["canonical_book_id"])
    catalog_book_id = int(book["id"])
    title = status["title"].strip()
    require(status["effective_status"] == "catalog_ok", "Book is not catalog_ok")
    require(status["primary_source_format"] in {"epub", "docx"}, "Unsupported source format")
    require(book["title"] == title, "Catalog edition title differs from book status")
    require(book["author"] == status["author"], "Author mismatch")
    require(book["book_shortener"] == status["book_shortener"], "Shortener mismatch")
    require(book["source_format"] == status["primary_source_format"], "Source format mismatch")
    require(len(poems) == int(book["poem_count"]) == int(status["catalog_poem_count"]), "Poem count mismatch")
    require(len(poems) > 0, "No poems to import")
    require(len({poem["id"] for poem in poems}) == len(poems), "Duplicate poem IDs")
    # Catalog poem IDs identify source rows. A book may legitimately reuse a title.
    require(all(poem["book_title"] == title for poem in poems), "Poem book identity mismatch")
    require(all(poem["source_format"] == book["source_format"] for poem in poems), "Poem source mismatch")
    require(all(poem["title"].strip() and (poem.get("served_text") or poem.get("text")) for poem in poems), "Blank title or text")

    connection = sqlite3.connect(database)
    try:
        connection.execute("PRAGMA foreign_keys = ON")
        with connection:
            require(not connection.execute("SELECT 1 FROM canonical_books WHERE id = ? OR lower(title) = lower(?) OR lower(book_shortener) = lower(?)", (book_id, title, status["book_shortener"])).fetchone(), "Book identity already exists")
            require(not connection.execute("SELECT 1 FROM catalog_books WHERE id = ? OR slug = ?", (catalog_book_id, book["slug"])).fetchone(), "Catalog edition already exists")
            poem_ids = [int(poem["id"]) for poem in poems]
            require(not connection.execute(f"SELECT 1 FROM catalog_poems WHERE id IN ({','.join('?' for _ in poem_ids)}) LIMIT 1", poem_ids).fetchone(), "Poem ID already exists")

            connection.execute(
                """INSERT INTO canonical_books
                   (id,title,author,book_shortener,release_catalog,pub_date,book_status,book_release_year)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (book_id,title,status["author"],status["book_shortener"],book["release_catalog"],status["pub_date"],book["book_status"],status["book_release_year"]),
            )
            connection.execute(
                """INSERT INTO book_status
                   (canonical_book_id,title,author,pub_date,book_release_year,book_shortener,csv_book_status,
                    override_status,override_reason,source_file_count,epub_source_count,zip_source_count,
                    manifest_entry_count,max_manifest_poem_count,catalog_book_count,catalog_poem_count,
                    effective_status,primary_source_format,notes)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (book_id,title,status["author"],status["pub_date"],status["book_release_year"],status["book_shortener"],
                 status["csv_book_status"],status["override_status"],status["override_reason"],status["source_file_count"],
                 status["epub_source_count"],status["zip_source_count"],status["manifest_entry_count"],
                 status["max_manifest_poem_count"],status["catalog_book_count"],status["catalog_poem_count"],
                 status["effective_status"],status["primary_source_format"],status["notes"]),
            )
            connection.execute(
                """INSERT INTO catalog_books
                   (id,slug,canonical_book_id,title,author,source_title,source_author,source_path,
                    source_format,poem_count,book_shortener,release_catalog,pub_date,book_status)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (catalog_book_id,book["slug"],book_id,title,book["author"],book["source_title"],
                 book["source_author"],book["source_path"],book["source_format"],book["poem_count"],
                 book["book_shortener"],book["release_catalog"],book["pub_date"],book["book_status"]),
            )
            connection.executemany(
                """INSERT INTO catalog_poems
                   (id,catalog_book_id,title,text,character_count,word_count,extraction_method)
                   VALUES (?,?,?,?,?,?,?)""",
                [(int(poem["id"]),catalog_book_id,poem["title"],poem.get("served_text") or poem["text"],
                  poem["character_count"],poem["word_count"],poem["extraction_method"]) for poem in poems],
            )
            require(connection.execute("PRAGMA foreign_key_check").fetchall() == [], "Foreign-key check failed")
        return {"title": title, "bookShortener": status["book_shortener"], "poemCount": len(poems), "sourcePath": book["source_path"]}
    finally:
        connection.close()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path, help="Main Catalog book JSON response")
    parser.add_argument("--database", type=Path, default=DEFAULT_DB_PATH)
    args = parser.parse_args()
    print(json.dumps(import_book_payload(json.loads(args.source.read_text()), args.database), indent=2))


if __name__ == "__main__":
    main()
