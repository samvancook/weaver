#!/usr/bin/env python3
from __future__ import annotations

import json
import re
import sqlite3
import sys
from functools import lru_cache
from pathlib import Path

DB_PATH = Path("/Users/buttonpublishingone/Desktop/CODEX/Social Media Dev/poetry_catalog/formal_catalog.db")


def normalize(text: str | None) -> str:
    if not text:
        return ""
    text = text.lower()
    text = text.replace("—", " ").replace("–", " ")
    text = text.replace("&", " and ")
    text = text.replace("’", "'").replace("‘", "'")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_title(text: str | None) -> str:
    normalized = normalize(text)
    normalized = re.sub(r"[^a-z0-9\s]", " ", normalized)
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized.strip()


def title_aliases(text: str | None) -> set[str]:
    raw = (text or "").strip()
    if not raw:
        return set()

    aliases = {normalize(raw), normalize_title(raw)}

    pieces = re.split(r"\s*[:\-]\s*", raw, maxsplit=1)
    if pieces:
        aliases.add(normalize(pieces[0]))
        aliases.add(normalize_title(pieces[0]))

    simplified = raw.replace("&", "and")
    aliases.add(normalize(simplified))
    aliases.add(normalize_title(simplified))

    return {alias for alias in aliases if alias}


def candidate_snippets(text: str | None) -> list[str]:
    raw = (text or "").strip()
    if not raw:
      return []
    normalized = normalize(raw)
    lines = [normalize(line) for line in raw.splitlines() if normalize(line)]
    snippets: list[str] = []
    for candidate in [
        normalized,
        max(lines, key=len) if lines else "",
        " ".join(normalized.split()[:12]),
        normalized[:120].strip(),
    ]:
        if candidate and candidate not in snippets:
            snippets.append(candidate)
    return snippets


@lru_cache(maxsize=1)
def load_book_status_rows() -> list[dict]:
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    try:
        rows = connection.execute(
            """
            SELECT canonical_book_id, title, author, book_shortener, effective_status
            FROM book_status
            """
        ).fetchall()
        return [dict(row) for row in rows]
    finally:
        connection.close()


def fetch_book_status(_cursor: sqlite3.Cursor, book_title: str) -> dict | None:
    query_aliases = title_aliases(book_title)
    if not query_aliases:
        return None

    best_match: dict | None = None
    best_score = -1

    for row in load_book_status_rows():
        candidate_aliases = title_aliases(row["title"])
        if row.get("book_shortener"):
            candidate_aliases.update(title_aliases(row["book_shortener"]))

        if query_aliases & candidate_aliases:
            # Prefer exact normalized title matches over looser alias/prefix matches.
            score = 2 if normalize(book_title) == normalize(row["title"]) else 1
            if score > best_score:
                best_score = score
                best_match = row

    return best_match


def fetch_poems_for_book(cursor: sqlite3.Cursor, canonical_book_id: int) -> list[dict]:
    return cursor.execute(
        """
        SELECT cp.title, cp.text, cp.word_count, cb.title AS book_title, cb.author
        FROM catalog_poems cp
        JOIN catalog_books cb ON cb.id = cp.catalog_book_id
        WHERE cb.canonical_book_id = ?
        """,
        (canonical_book_id,),
    ).fetchall()


def find_global_excerpt_match(cursor: sqlite3.Cursor, snippets: list[str]) -> dict | None:
    rows = cursor.execute(
        """
        SELECT cp.title, cp.text, cb.title AS book_title, cb.author
        FROM catalog_poems cp
        JOIN catalog_books cb ON cb.id = cp.catalog_book_id
        """
    ).fetchall()
    for row in rows:
        poem_text = normalize(row["text"])
        if any(snippet and snippet in poem_text for snippet in snippets):
            return {
                "book_title": row["book_title"],
                "author": row["author"],
                "poem_title": row["title"],
            }
    return None


def validate_record(cursor: sqlite3.Cursor, record: dict) -> dict:
    book_title = record.get("bookTitle") or ""
    author = record.get("author") or ""
    poem_title = record.get("title") or ""
    snippets = candidate_snippets(record.get("excerptText"))
    book_status = fetch_book_status(cursor, book_title)

    result = {
        "recordId": record.get("recordId"),
        "sourceRow": record.get("sourceRow"),
        "bookFound": bool(book_status),
        "bookEffectiveStatus": None,
        "bookCanonicalTitle": None,
        "bookCanonicalAuthor": None,
        "authorMatchesBook": None,
        "poemTitleMatchesInBook": False,
        "excerptMatchesInBook": False,
        "matchedPoemTitle": None,
        "globalExcerptMatch": None,
        "status": "unvalidated",
    }

    if not book_status:
        result["globalExcerptMatch"] = find_global_excerpt_match(cursor, snippets)
        result["status"] = "book_not_found"
        return result

    result["bookEffectiveStatus"] = book_status["effective_status"]
    result["bookCanonicalTitle"] = book_status["title"]
    result["bookCanonicalAuthor"] = book_status["author"]
    result["authorMatchesBook"] = normalize(author) == normalize(book_status["author"])

    if book_status["effective_status"] == "skip_epub":
        result["status"] = "epub_not_present"
        return result

    poems = fetch_poems_for_book(cursor, int(book_status["canonical_book_id"]))
    normalized_poem_title = normalize(poem_title)

    for poem in poems:
        poem_text = normalize(poem["text"])
        if normalize(poem["title"]) == normalized_poem_title:
            result["poemTitleMatchesInBook"] = True
        if any(snippet and snippet in poem_text for snippet in snippets):
            result["excerptMatchesInBook"] = True
            result["matchedPoemTitle"] = poem["title"]
            break

    if result["excerptMatchesInBook"]:
        title_matches_excerpt = normalize(result["matchedPoemTitle"]) == normalized_poem_title
        if not result["authorMatchesBook"]:
            result["status"] = "author_mismatch"
        elif title_matches_excerpt:
            result["status"] = "catalog_match"
        else:
            result["status"] = "title_mismatch"
    elif result["poemTitleMatchesInBook"]:
        result["status"] = "poem_title_match_only"
    else:
        result["globalExcerptMatch"] = find_global_excerpt_match(cursor, snippets)
        result["status"] = "excerpt_not_found_in_book"

    return result


def main() -> int:
    payload = json.load(sys.stdin)
    records = payload.get("records", [])
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    cursor = connection.cursor()
    try:
        results = [validate_record(cursor, record) for record in records]
    finally:
        connection.close()
    json.dump({"ok": True, "results": results}, sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
