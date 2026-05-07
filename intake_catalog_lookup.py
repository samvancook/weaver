#!/usr/bin/env python3
from __future__ import annotations

import json
import re
import sqlite3
import sys

from catalog_validate import DB_PATH, fetch_book_status, fetch_poems_for_book, load_book_status_rows


def clean_title(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())


def is_probable_poem_title(text: str) -> bool:
    title = clean_title(text)
    if not title:
        return False
    if len(title) > 140:
      return False
    if len(title.split()) > 18:
      return False
    sentence_like_marks = title.count(".") + title.count("?") + title.count("!")
    if sentence_like_marks >= 2:
      return False
    return True


def get_intake_books() -> dict:
    books = []
    for row in load_book_status_rows():
        source_format = (row.get("primary_source_format") or "").strip().lower()
        effective_status = (row.get("effective_status") or "").strip().lower()
        if source_format != "epub":
            continue
        if effective_status == "skip_epub":
            continue
        books.append(
            {
                "title": row.get("title") or "",
                "author": row.get("author") or "",
                "bookShortener": row.get("book_shortener") or "",
                "primarySourceFormat": row.get("primary_source_format") or "",
                "effectiveStatus": row.get("effective_status") or "",
            }
        )

    books.sort(key=lambda book: ((book["title"] or "").lower(), (book["author"] or "").lower()))
    return {"ok": True, "books": books}


def get_poems(book_title: str) -> dict:
    if not (book_title or "").strip():
        return {"ok": False, "error": "bookTitle is required."}

    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    try:
        cursor = connection.cursor()
        book_status = fetch_book_status(cursor, book_title)
        if not book_status:
            return {"ok": False, "error": "Book not found in catalog."}

        poems = fetch_poems_for_book(cursor, int(book_status["canonical_book_id"]))
        titles = sorted(
            {
                clean_title(poem["title"] or "")
                for poem in poems
                if is_probable_poem_title(poem["title"] or "")
            },
            key=str.lower,
        )
        return {
            "ok": True,
            "bookTitle": book_status["title"],
            "author": book_status["author"],
            "primarySourceFormat": book_status.get("primary_source_format") or "",
            "poems": titles,
        }
    finally:
        connection.close()


def main() -> int:
    payload = json.load(sys.stdin)
    action = (payload.get("action") or "").strip().lower()
    if action == "books":
        result = get_intake_books()
    elif action == "poems":
        result = get_poems(str(payload.get("bookTitle") or ""))
    else:
        result = {"ok": False, "error": "Unsupported action."}
    json.dump(result, sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
