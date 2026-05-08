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
    text = re.sub(r"\*([^*\n]+)\*", r"\1", text)
    text = text.replace("*", "")
    text = text.lower()
    text = text.replace("—", " ").replace("–", " ")
    text = text.replace("&", " and ")
    text = text.replace("’", "'").replace("‘", "'")
    text = re.sub(r'[\"“”`]', "", text)
    text = re.sub(r"(?<![a-z0-9])'|'(?![a-z0-9])", "", text)
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


def preserve_excerpt_text(text: str | None) -> str:
    raw = re.sub(r"\*([^*\n]+)\*", r"\1", text or "")
    raw = raw.replace("*", "")
    raw = re.sub(r"(?<![a-z0-9])'|'(?![a-z0-9])", "", raw, flags=re.IGNORECASE)
    raw = raw.replace("\r\n", "\n").replace("\r", "\n")
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in raw.split("\n")]
    while lines and not lines[0]:
        lines.pop(0)
    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines).strip()


def collapse_blank_lines(text: str | None) -> str:
    raw = preserve_excerpt_text(text)
    if not raw:
        return ""
    lines = [line for line in raw.split("\n") if line.strip()]
    return "\n".join(lines).strip()


def strip_wrapping_quotes(text: str | None) -> str:
    raw = (text or "").strip()
    if len(raw) >= 2 and raw[0] in {'"', "“", "”", "'"} and raw[-1] in {'"', "“", "”", "'"}:
        return raw[1:-1].strip()
    return raw


def extract_effective_poem_title_and_text(title: str | None, text: str | None) -> tuple[str, str]:
    raw_title = clean_title = (title or "").strip()
    raw_text = preserve_excerpt_text(text)
    raw_text_for_match = raw_text if len(raw_text) > 60 else ""

    if len(clean_title) > 120:
        tokens = clean_title.split()
        for index, token in enumerate(tokens):
            if index == 0:
                continue
            if re.search(r"[a-z]", token):
                effective_title = " ".join(tokens[:index]).strip()
                effective_text = " ".join(tokens[index:]).strip()
                if effective_title and effective_text:
                    return effective_title, effective_text

    return clean_title, raw_text_for_match


def titles_loosely_match(left: str | None, right: str | None) -> bool:
    normalized_left = normalize_title(left)
    normalized_right = normalize_title(right)
    if not normalized_left or not normalized_right:
        return False
    left_words = normalized_left.split()
    right_words = normalized_right.split()
    return (
        normalized_left == normalized_right
        or (
            len(normalized_left) >= 8
            and len(normalized_right) >= 8
            and len(left_words) >= 2
            and len(right_words) >= 2
            and (
                normalized_left.startswith(normalized_right)
                or normalized_right.startswith(normalized_left)
            )
        )
    )


@lru_cache(maxsize=1)
def load_book_status_rows() -> list[dict]:
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    try:
        rows = connection.execute(
            """
            SELECT canonical_book_id, title, author, book_shortener, effective_status, primary_source_format
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
        "bookPrimarySourceFormat": None,
        "bookCanonicalTitle": None,
        "bookCanonicalAuthor": None,
        "authorMatchesBook": None,
        "poemTitleMatchesInBook": False,
        "excerptMatchesInBook": False,
        "matchedPoemTitle": None,
        "catalogFormattingMatch": None,
        "catalogLineBreaksMatch": None,
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
    result["bookPrimarySourceFormat"] = book_status.get("primary_source_format")
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

    if book_status["effective_status"] != "catalog_ok":
        result["status"] = "catalog_unavailable"
        return result

    poems = fetch_poems_for_book(cursor, int(book_status["canonical_book_id"]))
    if not poems:
        result["status"] = "catalog_unavailable"
        return result

    normalized_poem_title = normalize_title(poem_title)
    query_excerpt_raw = preserve_excerpt_text(record.get("excerptText"))

    for poem in poems:
        effective_poem_title, effective_poem_text = extract_effective_poem_title_and_text(poem["title"], poem["text"])
        poem_text = normalize(effective_poem_text)
        if titles_loosely_match(effective_poem_title, poem_title):
            if preserve_excerpt_text(effective_poem_text):
                result["poemTitleMatchesInBook"] = True
        if any(snippet and snippet in poem_text for snippet in snippets):
            result["excerptMatchesInBook"] = True
            result["matchedPoemTitle"] = effective_poem_title or poem["title"]
            matched_raw_text = preserve_excerpt_text(effective_poem_text)
            matched_raw_text_collapsed = collapse_blank_lines(effective_poem_text)
            query_excerpt_unquoted = strip_wrapping_quotes(query_excerpt_raw)
            query_excerpt_collapsed = collapse_blank_lines(query_excerpt_raw)
            query_excerpt_unquoted_collapsed = collapse_blank_lines(query_excerpt_unquoted)
            result["catalogFormattingMatch"] = bool(
                (matched_raw_text or matched_raw_text_collapsed) and (
                    (query_excerpt_raw and query_excerpt_raw in matched_raw_text)
                    or (query_excerpt_unquoted and query_excerpt_unquoted in matched_raw_text)
                    or (query_excerpt_collapsed and query_excerpt_collapsed in matched_raw_text_collapsed)
                    or (query_excerpt_unquoted_collapsed and query_excerpt_unquoted_collapsed in matched_raw_text_collapsed)
                )
            )
            result["catalogLineBreaksMatch"] = result["catalogFormattingMatch"]
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
