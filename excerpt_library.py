#!/usr/bin/env python3
from __future__ import annotations

import csv
import hashlib
import json
import re
import sqlite3
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Iterable

DEFAULT_DB_PATH = Path(__file__).resolve().parent / "data" / "excerpt_library.db"

TEXT_COLUMN_FALLBACKS = (
    "Enter the Quote",
    "Quote",
    "Excerpt",
    "Text",
    "excerptText",
)

BOOK_COLUMN_FALLBACKS = (
    "Book Title",
    "bookTitle",
    "Book",
    "Collection",
)

TITLE_COLUMN_FALLBACKS = (
    "Title",
    "Poem Title",
    "title",
    "Poem",
)

AUTHOR_COLUMN_FALLBACKS = (
    "Author",
    "author",
)

ID_COLUMN_FALLBACKS = (
    "recordId",
    "Record ID",
    "ID",
    "id",
)


@dataclass(frozen=True)
class ImportedExcerpt:
    row_number: int
    external_id: str
    author: str
    normalized_author: str
    book_title: str
    normalized_book_title: str
    poem_title: str
    normalized_poem_title: str
    excerpt_text: str
    normalized_excerpt: str
    excerpt_hash: str
    word_count: int
    character_count: int
    metadata_json: str


def clean_whitespace(text: str | None) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip()


def normalize_text(text: str | None) -> str:
    if not text:
        return ""
    normalized = text.lower()
    normalized = normalized.replace("—", " ").replace("–", " ")
    normalized = normalized.replace("&", " and ")
    normalized = normalized.replace("’", "'").replace("‘", "'")
    normalized = re.sub(r"[\"“”`]", "", normalized)
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized.strip()


def normalize_lookup_text(text: str | None) -> str:
    normalized = normalize_text(text)
    normalized = re.sub(r"[^a-z0-9\s']", " ", normalized)
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized.strip()


def tokenize(text: str | None) -> set[str]:
    return set(re.findall(r"[a-z0-9']+", normalize_lookup_text(text)))


def fingerprint_excerpt(text: str) -> str:
    return hashlib.sha256(normalize_lookup_text(text).encode("utf-8")).hexdigest()


def sequence_score(left: str, right: str) -> float:
    return SequenceMatcher(None, normalize_lookup_text(left), normalize_lookup_text(right)).ratio()


def token_score(left: str, right: str) -> tuple[float, int]:
    left_tokens = tokenize(left)
    right_tokens = tokenize(right)
    if not left_tokens or not right_tokens:
        return 0.0, 0
    shared = left_tokens & right_tokens
    union = left_tokens | right_tokens
    return len(shared) / len(union), len(shared)


def resolve_column(fieldnames: Iterable[str] | None, explicit: str | None, fallbacks: tuple[str, ...]) -> str:
    available = list(fieldnames or [])
    if not available:
        raise ValueError("The CSV file does not have a header row.")

    if explicit:
        if explicit not in available:
            raise ValueError(
                f'Column "{explicit}" was not found. Available columns: {", ".join(available)}'
            )
        return explicit

    for candidate in fallbacks:
        if candidate in available:
            return candidate

    return ""


def ensure_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        PRAGMA journal_mode = WAL;

        CREATE TABLE IF NOT EXISTS excerpt_sources (
            id INTEGER PRIMARY KEY,
            source_name TEXT NOT NULL,
            source_kind TEXT NOT NULL,
            source_path TEXT,
            imported_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(source_name, source_path)
        );

        CREATE TABLE IF NOT EXISTS excerpt_entries (
            id INTEGER PRIMARY KEY,
            source_id INTEGER NOT NULL REFERENCES excerpt_sources(id) ON DELETE CASCADE,
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
            character_count INTEGER NOT NULL,
            metadata_json TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        CREATE INDEX IF NOT EXISTS idx_excerpt_entries_hash
            ON excerpt_entries(excerpt_hash);
        CREATE INDEX IF NOT EXISTS idx_excerpt_entries_book
            ON excerpt_entries(normalized_book_title);
        CREATE INDEX IF NOT EXISTS idx_excerpt_entries_author
            ON excerpt_entries(normalized_author);
        CREATE INDEX IF NOT EXISTS idx_excerpt_entries_poem
            ON excerpt_entries(normalized_poem_title);
        CREATE INDEX IF NOT EXISTS idx_excerpt_entries_chars
            ON excerpt_entries(character_count);
        """
    )


def build_imported_excerpt(
    row_number: int,
    record: dict[str, str],
    text_column: str,
    book_column: str,
    title_column: str,
    author_column: str,
    id_column: str,
) -> ImportedExcerpt | None:
    excerpt_text = clean_whitespace(record.get(text_column, ""))
    if not excerpt_text:
        return None

    author = clean_whitespace(record.get(author_column, "")) if author_column else ""
    book_title = clean_whitespace(record.get(book_column, "")) if book_column else ""
    poem_title = clean_whitespace(record.get(title_column, "")) if title_column else ""
    external_id = clean_whitespace(record.get(id_column, "")) if id_column else ""

    return ImportedExcerpt(
        row_number=row_number,
        external_id=external_id,
        author=author,
        normalized_author=normalize_lookup_text(author),
        book_title=book_title,
        normalized_book_title=normalize_lookup_text(book_title),
        poem_title=poem_title,
        normalized_poem_title=normalize_lookup_text(poem_title),
        excerpt_text=excerpt_text,
        normalized_excerpt=normalize_lookup_text(excerpt_text),
        excerpt_hash=fingerprint_excerpt(excerpt_text),
        word_count=len(excerpt_text.split()),
        character_count=len(excerpt_text),
        metadata_json=json.dumps(record, ensure_ascii=True),
    )


def import_csv_to_library(
    csv_path: Path,
    db_path: Path = DEFAULT_DB_PATH,
    source_name: str | None = None,
    source_kind: str = "spreadsheet",
    text_column: str | None = None,
    book_column: str | None = None,
    title_column: str | None = None,
    author_column: str | None = None,
    id_column: str | None = None,
    replace_source: bool = False,
) -> dict[str, int | str]:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    source_name = source_name or csv_path.stem

    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        text_column = resolve_column(reader.fieldnames, text_column, TEXT_COLUMN_FALLBACKS)
        if not text_column:
            raise ValueError("Could not find an excerpt text column automatically.")

        book_column = resolve_column(reader.fieldnames, book_column, BOOK_COLUMN_FALLBACKS)
        title_column = resolve_column(reader.fieldnames, title_column, TITLE_COLUMN_FALLBACKS)
        author_column = resolve_column(reader.fieldnames, author_column, AUTHOR_COLUMN_FALLBACKS)
        id_column = resolve_column(reader.fieldnames, id_column, ID_COLUMN_FALLBACKS)

        connection = sqlite3.connect(db_path)
        try:
            ensure_schema(connection)
            with connection:
                cursor = connection.cursor()
                cursor.execute(
                    """
                    INSERT OR IGNORE INTO excerpt_sources (source_name, source_kind, source_path)
                    VALUES (?, ?, ?)
                    """,
                    (source_name, source_kind, str(csv_path)),
                )
                source_id = cursor.execute(
                    """
                    SELECT id
                    FROM excerpt_sources
                    WHERE source_name = ? AND source_path = ?
                    """,
                    (source_name, str(csv_path)),
                ).fetchone()[0]

                if replace_source:
                    cursor.execute("DELETE FROM excerpt_entries WHERE source_id = ?", (source_id,))

                imported_rows = 0
                for row_number, record in enumerate(reader, start=2):
                    excerpt = build_imported_excerpt(
                        row_number=row_number,
                        record=record,
                        text_column=text_column,
                        book_column=book_column,
                        title_column=title_column,
                        author_column=author_column,
                        id_column=id_column,
                    )
                    if not excerpt:
                        continue

                    cursor.execute(
                        """
                        INSERT INTO excerpt_entries (
                            source_id,
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
                            character_count,
                            metadata_json
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            source_id,
                            excerpt.row_number,
                            excerpt.external_id,
                            excerpt.author,
                            excerpt.normalized_author,
                            excerpt.book_title,
                            excerpt.normalized_book_title,
                            excerpt.poem_title,
                            excerpt.normalized_poem_title,
                            excerpt.excerpt_text,
                            excerpt.normalized_excerpt,
                            excerpt.excerpt_hash,
                            excerpt.word_count,
                            excerpt.character_count,
                            excerpt.metadata_json,
                        ),
                    )
                    imported_rows += 1
        finally:
            connection.close()

    return {
        "db_path": str(db_path),
        "source_name": source_name,
        "imported_rows": imported_rows,
        "text_column": text_column,
    }


def find_library_excerpt_match(
    connection: sqlite3.Connection,
    excerpt_text: str | None,
    book_title: str | None = None,
    author: str | None = None,
    threshold: float = 0.72,
) -> dict | None:
    cleaned_excerpt = clean_whitespace(excerpt_text)
    if not cleaned_excerpt:
        return None

    normalized_excerpt = normalize_lookup_text(cleaned_excerpt)
    excerpt_hash = fingerprint_excerpt(cleaned_excerpt)
    book_alias = normalize_lookup_text(book_title)
    author_alias = normalize_lookup_text(author)
    excerpt_len = len(cleaned_excerpt)

    exact_row = connection.execute(
        """
        SELECT source_row_number, external_id, author, book_title, poem_title, excerpt_text
        FROM excerpt_entries
        WHERE excerpt_hash = ?
        LIMIT 1
        """,
        (excerpt_hash,),
    ).fetchone()
    if exact_row:
        return {
            "matchType": "exact",
            "score": 1.0,
            "sourceRow": exact_row[0],
            "recordId": exact_row[1],
            "author": exact_row[2],
            "bookTitle": exact_row[3],
            "poemTitle": exact_row[4],
            "excerptPreview": exact_row[5][:180],
        }

    candidate_rows = connection.execute(
        """
        SELECT source_row_number, external_id, author, normalized_author, book_title,
               normalized_book_title, poem_title, excerpt_text, character_count
        FROM excerpt_entries
        WHERE character_count BETWEEN ? AND ?
        ORDER BY ABS(character_count - ?), source_row_number
        LIMIT 250
        """,
        (max(1, excerpt_len - 160), excerpt_len + 160, excerpt_len),
    ).fetchall()

    best_match: dict | None = None
    best_score = threshold
    for row in candidate_rows:
        candidate_text = row[7]
        normalized_candidate = normalize_lookup_text(candidate_text)
        if not normalized_candidate:
            continue

        if normalized_excerpt in normalized_candidate or normalized_candidate in normalized_excerpt:
            similarity = 0.98
            match_type = "substring"
            shared_token_count = len(tokenize(cleaned_excerpt) & tokenize(candidate_text))
        else:
            overlap_score, shared_token_count = token_score(cleaned_excerpt, candidate_text)
            similarity = max(overlap_score, sequence_score(cleaned_excerpt, candidate_text))
            match_type = "near_duplicate"

        if book_alias and row[5] and row[5] == book_alias:
            similarity += 0.02
        if author_alias and row[3] and row[3] == author_alias:
            similarity += 0.02

        if similarity < best_score:
            continue

        best_score = similarity
        best_match = {
            "matchType": match_type,
            "score": round(min(similarity, 1.0), 3),
            "sharedTokenCount": shared_token_count,
            "sourceRow": row[0],
            "recordId": row[1],
            "author": row[2],
            "bookTitle": row[4],
            "poemTitle": row[6],
            "excerptPreview": candidate_text[:180],
        }

    return best_match


def connect_library(db_path: Path = DEFAULT_DB_PATH) -> sqlite3.Connection:
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    return connection


def get_exact_excerpt_matches(
    connection: sqlite3.Connection,
    excerpt_text: str | None,
) -> list[dict]:
    cleaned_excerpt = clean_whitespace(excerpt_text)
    if not cleaned_excerpt:
        return []

    excerpt_hash = fingerprint_excerpt(cleaned_excerpt)
    rows = connection.execute(
        """
        SELECT source_row_number, external_id, author, book_title, poem_title, excerpt_text
        FROM excerpt_entries
        WHERE excerpt_hash = ?
        ORDER BY source_row_number
        """,
        (excerpt_hash,),
    ).fetchall()
    return [
        {
            "sourceRow": row["source_row_number"],
            "recordId": row["external_id"] or "",
            "author": row["author"] or "",
            "bookTitle": row["book_title"] or "",
            "poemTitle": row["poem_title"] or "",
            "excerptText": row["excerpt_text"],
        }
        for row in rows
    ]
