#!/usr/bin/env python3
from __future__ import annotations

import json
import re
import sqlite3
import sys

from catalog_validate import DB_PATH, fetch_book_status, fetch_poems_for_book, load_book_status_rows
from catalog_live import fetch_readable_book, list_readable_books


def clean_title(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())


def display_poem_options(poems: list[dict]) -> list[dict]:
    rows = [poem for poem in poems if clean_title(poem["title"] or "")]
    counts: dict[str, int] = {}
    for poem in rows:
        title = clean_title(poem["title"])
        counts[title] = counts.get(title, 0) + 1
    seen: dict[str, int] = {}
    roman = ("I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X")
    options = []
    for poem in rows:
        title = clean_title(poem["title"])
        seen[title] = seen.get(title, 0) + 1
        suffix = roman[seen[title] - 1] if seen[title] <= len(roman) else str(seen[title])
        options.append({
            "title": title,
            "label": f"{title} ({suffix})" if counts[title] > 1 else title,
            "poemKey": poem["poem_key"] or "",
        })
    return sorted(options, key=lambda option: option["title"].lower())


def get_intake_books() -> dict:
    books = []
    try:
        status_rows = list_readable_books()
    except (OSError, ValueError):
        status_rows = load_book_status_rows()
    for row in status_rows:
        source_format = (row.get("primary_source_format") or "").strip().lower()
        effective_status = (row.get("effective_status") or "").strip().lower()
        if source_format not in {"epub", "docx", "pdf"}:
            continue
        if effective_status not in {"catalog_ok", "provisional_source"}:
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

    try:
        live = fetch_readable_book(book_title)
    except (OSError, ValueError):
        live = "unavailable"
    if live != "unavailable":
        if not live:
            return {"ok": False, "error": "Book has no readable poems in the current Catalog."}
        status = live["status"]
        poems = live["poems"]
        return {
            "ok": True,
            "bookTitle": status["title"],
            "author": status["author"],
            "primarySourceFormat": status.get("primary_source_format") or "",
            "poems": sorted({clean_title(poem["title"]) for poem in poems}, key=str.lower),
            "poemOptions": display_poem_options(poems),
        }

    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    try:
        cursor = connection.cursor()
        book_status = fetch_book_status(cursor, book_title)
        if not book_status:
            return {"ok": False, "error": "Book not found in catalog."}

        poems = fetch_poems_for_book(cursor, int(book_status["canonical_book_id"]))
        poem_options = display_poem_options(poems)
        titles = sorted(
            {
                clean_title(poem["title"] or "")
                for poem in poems
                if clean_title(poem["title"] or "")
            },
            key=str.lower,
        )
        if not poem_options:
            return {"ok": False, "error": "Catalog lookup is temporarily unavailable."}
        return {
            "ok": True,
            "bookTitle": book_status["title"],
            "author": book_status["author"],
            "primarySourceFormat": book_status.get("primary_source_format") or "",
            "poems": titles,
            "poemOptions": poem_options,
        }
    finally:
        connection.close()


def verify_poem_selection(book_title: str, poem_title: str, poem_key: str) -> dict:
    if not book_title or not poem_title or not poem_key:
        return {"ok": False, "error": "Book, poem title, and poem key are required."}
    try:
        live = fetch_readable_book(book_title)
    except (OSError, ValueError):
        live = "unavailable"
    if live != "unavailable":
        poems = live["poems"] if live else []
        selected = next((poem for poem in poems if poem["poem_key"] == poem_key), None)
        if not selected or clean_title(selected["title"]) != clean_title(poem_title):
            return {"ok": False, "error": "Selected poem key does not match this book and title."}
        return {"ok": True}
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    try:
        cursor = connection.cursor()
        book_status = fetch_book_status(cursor, book_title)
        if not book_status:
            return {"ok": False, "error": "Book not found in catalog."}
        poems = fetch_poems_for_book(cursor, int(book_status["canonical_book_id"]))
        selected = next((poem for poem in poems if poem["poem_key"] == poem_key), None)
        if not selected or clean_title(selected["title"]) != clean_title(poem_title):
            return {"ok": False, "error": "Selected poem key does not match this book and title."}
        return {"ok": True}
    finally:
        connection.close()


def main() -> int:
    payload = json.load(sys.stdin)
    action = (payload.get("action") or "").strip().lower()
    if action == "books":
        result = get_intake_books()
    elif action == "poems":
        result = get_poems(str(payload.get("bookTitle") or ""))
    elif action == "verify":
        result = verify_poem_selection(
            str(payload.get("bookTitle") or ""),
            str(payload.get("poemTitle") or ""),
            str(payload.get("poemKey") or ""),
        )
    else:
        result = {"ok": False, "error": "Unsupported action."}
    json.dump(result, sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
