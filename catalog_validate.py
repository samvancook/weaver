#!/usr/bin/env python3
from __future__ import annotations

import json
import re
import sqlite3
import sys
from functools import lru_cache
from pathlib import Path

from excerpt_library import DEFAULT_DB_PATH as EXCERPT_LIBRARY_DB_PATH
from excerpt_library import find_library_excerpt_match

DEFAULT_DB_PATH = Path(__file__).resolve().parent / "data" / "formal_catalog.db"
LEGACY_DB_PATH = Path("/Users/buttonpublishingone/Desktop/CODEX/Social Media Dev/poetry_catalog/formal_catalog.db")
DB_PATH = DEFAULT_DB_PATH if DEFAULT_DB_PATH.exists() else LEGACY_DB_PATH


def normalize(text: str | None) -> str:
    if not text:
        return ""
    text = text.lower()
    text = text.replace("—", " ").replace("–", " ")
    text = text.replace("&", " and ")
    text = text.replace("’", "'").replace("‘", "'")
    text = re.sub(r'[\"“”`]', "", text)
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


@lru_cache(maxsize=1)
def load_author_alias_map() -> dict[int, set[str]]:
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    try:
        rows = connection.execute(
            """
            SELECT canonical_book_id, alias
            FROM author_aliases
            """
        ).fetchall()
    except sqlite3.OperationalError:
        return {}
    finally:
        connection.close()

    alias_map: dict[int, set[str]] = {}
    for row in rows:
        canonical_book_id = int(row["canonical_book_id"])
        alias_map.setdefault(canonical_book_id, set()).add(normalize(row["alias"]))
    return alias_map


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


def load_excerpt_library_connection() -> sqlite3.Connection | None:
    if not EXCERPT_LIBRARY_DB_PATH.exists():
        return None
    connection = sqlite3.connect(EXCERPT_LIBRARY_DB_PATH)
    connection.row_factory = sqlite3.Row
    return connection


def validate_record(
    cursor: sqlite3.Cursor,
    record: dict,
    excerpt_library_connection: sqlite3.Connection | None = None,
) -> dict:
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
        "libraryExcerptMatch": None,
        "status": "unvalidated",
    }

    if excerpt_library_connection is not None:
        result["libraryExcerptMatch"] = find_library_excerpt_match(
            excerpt_library_connection,
            excerpt_text=record.get("excerptText"),
            book_title=book_title,
            author=author,
        )

    if not book_status:
        result["globalExcerptMatch"] = find_global_excerpt_match(cursor, snippets)
        result["status"] = "book_not_found"
        return result

    result["bookEffectiveStatus"] = book_status["effective_status"]
    result["bookCanonicalTitle"] = book_status["title"]
    result["bookCanonicalAuthor"] = book_status["author"]
    canonical_book_id = int(book_status["canonical_book_id"])
    allowed_author_names = {
        normalize(book_status["author"])
    }
    allowed_author_names.update(load_author_alias_map().get(canonical_book_id, set()))
    result["authorMatchesBook"] = normalize(author) in allowed_author_names

    if book_status["effective_status"] == "skip_epub":
        result["status"] = "epub_not_present"
        return result

    poems = fetch_poems_for_book(cursor, int(book_status["canonical_book_id"]))
    normalized_poem_title = normalize_title(poem_title)

    for poem in poems:
        poem_text = normalize(poem["text"])
        if normalize_title(poem["title"]) == normalized_poem_title:
            result["poemTitleMatchesInBook"] = True
        if any(snippet and snippet in poem_text for snippet in snippets):
            result["excerptMatchesInBook"] = True
            result["matchedPoemTitle"] = poem["title"]
            break

    if result["excerptMatchesInBook"]:
        title_matches_excerpt = normalize_title(result["matchedPoemTitle"]) == normalized_poem_title
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
    excerpt_library_connection = load_excerpt_library_connection()
    try:
        results = [
            validate_record(cursor, record, excerpt_library_connection=excerpt_library_connection)
            for record in records
        ]
    finally:
        if excerpt_library_connection is not None:
            excerpt_library_connection.close()
        connection.close()
    json.dump({"ok": True, "results": results}, sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
