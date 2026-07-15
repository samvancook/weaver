#!/usr/bin/env python3
from __future__ import annotations

import csv
import hashlib
import json
import re
import sqlite3
from dataclasses import dataclass
from difflib import SequenceMatcher
from functools import lru_cache
from pathlib import Path
from typing import Iterable

RUNTIME_DB_PATH = Path(__file__).resolve().parent / "data" / "excerpt_library_runtime.db"
FULL_DB_PATH = Path(__file__).resolve().parent / "data" / "excerpt_library.db"
QI_STATUS_DB_PATH = Path(__file__).resolve().parent / "data" / "excerpt_library_qi_status.db"
DEFAULT_DB_PATH = FULL_DB_PATH if FULL_DB_PATH.exists() else RUNTIME_DB_PATH

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


def strip_markdown_italics(text: str | None) -> str:
    raw = text or ""
    raw = re.sub(r"\*([^*\n]+)\*", r"\1", raw)
    raw = raw.replace("*", "")
    return raw


def preserve_excerpt_text(text: str | None) -> str:
    raw = strip_markdown_italics(text).replace("\r\n", "\n").replace("\r", "\n")
    raw = re.sub(r"(?<![a-z0-9])'|'(?![a-z0-9])", "", raw, flags=re.IGNORECASE)
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in raw.split("\n")]
    while lines and not lines[0]:
        lines.pop(0)
    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines).strip()


def line_break_signature(text: str | None) -> tuple[int, ...]:
    preserved = preserve_excerpt_text(text)
    if not preserved:
        return ()
    return tuple(len(line) for line in preserved.split("\n"))


def normalize_text(text: str | None) -> str:
    if not text:
        return ""
    normalized = strip_markdown_italics(text).lower()
    normalized = normalized.replace("—", " ").replace("–", " ")
    normalized = normalized.replace("&", " and ")
    normalized = normalized.replace("’", "'").replace("‘", "'")
    normalized = re.sub(r"[\"“”`]", "", normalized)
    normalized = re.sub(r"(?<![a-z0-9])'|'(?![a-z0-9])", "", normalized)
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


def normalize_flag(value: str | None) -> str:
    normalized = clean_whitespace(value).upper()
    if normalized in {"Y", "YES", "TRUE", "1"}:
        return "Y"
    if normalized in {"N", "NO", "FALSE", "0"}:
        return "N"
    return ""


def normalize_bool(value) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return bool(value)
    return normalize_flag(str(value)) == "Y"


def parse_metadata_json(text: str | None) -> dict[str, str]:
    if not text:
        return {}
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        return {}
    return parsed if isinstance(parsed, dict) else {}


def first_metadata_value(metadata: dict[str, str], *keys: str) -> str:
    for key in keys:
        value = metadata.get(key)
        if value is not None and clean_whitespace(value):
            return str(value)
    return ""


def build_library_status(metadata_json: str | None) -> dict[str, str | bool]:
    metadata = parse_metadata_json(metadata_json)
    approved_for_qi = normalize_flag(
        first_metadata_value(
            metadata,
            "Approved for Quote Image",
            "Approved for Quote Creation? Y / N",
            "Approved for Quote Creation",
            "approved_for_quote",
        )
    )
    quote_image_created = normalize_flag(
        first_metadata_value(
            metadata,
            "Quote Image Created",
            "quote_created_qc",
            "Quote Created & QCed (Y/N)",
        )
    )
    graphic_made = normalize_flag(
        first_metadata_value(
            metadata,
            "Graphic Made?",
            "graphic_made",
        )
    )

    made = quote_image_created == "Y" or graphic_made == "Y"
    approved = approved_for_qi == "Y"

    return {
        "hasQiAsset": approved or made,
        "approvedForQi": approved,
        "quoteImageCreated": quote_image_created == "Y",
        "graphicMade": graphic_made == "Y",
        "made": made,
    }


@lru_cache(maxsize=1)
def load_normalized_qi_status_map() -> dict[str, dict[str, int | bool]]:
    if not QI_STATUS_DB_PATH.exists():
        return {}

    connection = sqlite3.connect(QI_STATUS_DB_PATH)
    connection.row_factory = sqlite3.Row
    try:
        rows = connection.execute(
            """
            SELECT excerpt_hash, has_qi_asset, qi_asset_count, qi_linked_asset_count,
                   qi_approved_count, qi_graphic_made_count
            FROM excerpt_qi_status
            """
        ).fetchall()
    finally:
        connection.close()

    return {
        row["excerpt_hash"]: {
            "hasQiAsset": bool(row["has_qi_asset"]),
            "approvedForQi": int(row["qi_approved_count"] or 0) > 0,
            "quoteImageCreated": int(row["qi_graphic_made_count"] or 0) > 0,
            "graphicMade": int(row["qi_graphic_made_count"] or 0) > 0,
            "made": int(row["qi_graphic_made_count"] or 0) > 0,
            "qiAssetCount": int(row["qi_asset_count"] or 0),
            "qiLinkedAssetCount": int(row["qi_linked_asset_count"] or 0),
            "qiApprovedCount": int(row["qi_approved_count"] or 0),
            "qiGraphicMadeCount": int(row["qi_graphic_made_count"] or 0),
        }
        for row in rows
    }


def build_best_library_status(
    *,
    excerpt_hash: str | None = None,
    metadata_json: str | None = None,
) -> dict[str, str | bool | int]:
    if excerpt_hash:
        normalized_status = load_normalized_qi_status_map().get(excerpt_hash)
        if normalized_status is not None:
            return normalized_status
    return build_library_status(metadata_json)


def has_metadata_json_column(connection: sqlite3.Connection) -> bool:
    rows = connection.execute("PRAGMA table_info(excerpt_entries)").fetchall()
    return any((row[1] if isinstance(row, tuple) else row["name"]) == "metadata_json" for row in rows)


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

        CREATE TABLE IF NOT EXISTS raw_import_rows (
            id INTEGER PRIMARY KEY,
            source_name TEXT NOT NULL,
            source_file TEXT NOT NULL,
            source_row_number INTEGER NOT NULL,
            record_id TEXT,
            timestamp_text TEXT,
            email_address TEXT,
            request_type TEXT,
            author TEXT,
            title TEXT,
            book_title TEXT,
            excerpt_text TEXT,
            excerpt_review_decision TEXT,
            exclude_from_quote_db TEXT,
            row_json TEXT NOT NULL,
            imported_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(source_file, source_row_number)
        );

        CREATE INDEX IF NOT EXISTS idx_raw_import_rows_record_id
            ON raw_import_rows(record_id);
        CREATE INDEX IF NOT EXISTS idx_raw_import_rows_timestamp
            ON raw_import_rows(timestamp_text);
        CREATE INDEX IF NOT EXISTS idx_raw_import_rows_author
            ON raw_import_rows(author);

        CREATE TABLE IF NOT EXISTS raw_excerpt_links (
            id INTEGER PRIMARY KEY,
            raw_import_row_id INTEGER NOT NULL UNIQUE REFERENCES raw_import_rows(id) ON DELETE CASCADE,
            excerpt_entry_id INTEGER NOT NULL REFERENCES excerpt_entries(id) ON DELETE CASCADE,
            match_type TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        CREATE INDEX IF NOT EXISTS idx_raw_excerpt_links_excerpt_entry
            ON raw_excerpt_links(excerpt_entry_id);

        CREATE VIEW IF NOT EXISTS v_raw_row_status AS
        SELECT
            r.id AS raw_import_row_id,
            r.source_name,
            r.source_file,
            r.source_row_number,
            r.record_id,
            r.timestamp_text,
            r.email_address,
            r.request_type,
            r.author,
            r.title,
            r.book_title,
            r.excerpt_text,
            r.excerpt_review_decision,
            r.exclude_from_quote_db,
            CASE
                WHEN r.excerpt_review_decision = 'ACCEPT' THEN 1
                ELSE 0
            END AS is_accepted,
            CASE
                WHEN r.excerpt_review_decision = 'REJECT' THEN 1
                ELSE 0
            END AS is_rejected,
            CASE
                WHEN UPPER(COALESCE(r.exclude_from_quote_db, '')) = 'Y' THEN 1
                ELSE 0
            END AS is_excluded,
            CASE
                WHEN COALESCE(TRIM(r.excerpt_text), '') != '' THEN 1
                ELSE 0
            END AS has_excerpt_text,
            l.id AS raw_excerpt_link_id,
            l.match_type,
            l.excerpt_entry_id,
            CASE
                WHEN l.id IS NOT NULL THEN 1
                ELSE 0
            END AS is_linked,
            e.source_id AS excerpt_source_id,
            e.source_row_number AS excerpt_source_row_number,
            e.external_id AS excerpt_external_id,
            e.author AS excerpt_author,
            e.poem_title AS excerpt_poem_title,
            e.book_title AS excerpt_book_title,
            e.excerpt_text AS excerpt_entry_text
        FROM raw_import_rows r
        LEFT JOIN raw_excerpt_links l
            ON l.raw_import_row_id = r.id
        LEFT JOIN excerpt_entries e
            ON e.id = l.excerpt_entry_id;
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
    excerpt_text = preserve_excerpt_text(record.get(text_column, ""))
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


def _get_or_create_source_id(
    connection: sqlite3.Connection,
    source_name: str,
    source_kind: str,
    source_path: str,
) -> int:
    cursor = connection.cursor()
    cursor.execute(
        """
        INSERT OR IGNORE INTO excerpt_sources (source_name, source_kind, source_path)
        VALUES (?, ?, ?)
        """,
        (source_name, source_kind, source_path),
    )
    row = cursor.execute(
        """
        SELECT id
        FROM excerpt_sources
        WHERE source_name = ? AND source_path = ?
        """,
        (source_name, source_path),
    ).fetchone()
    if not row:
        raise RuntimeError("Could not resolve excerpt source id.")
    return int(row[0])


def ingest_weaver_approved_records(
    records: list[dict],
    db_path: Path = DEFAULT_DB_PATH,
    source_name: str = "weaver_approved",
    source_kind: str = "weaver",
    source_path: str = "weaver://approved",
) -> dict[str, int | str]:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(db_path)
    try:
        ensure_schema(connection)
        inserted = 0
        updated = 0
        skipped = 0
        archived_inserted = 0
        archived_updated = 0
        source_cache: dict[tuple[str, str, str], int] = {}

        with connection:
            cursor = connection.cursor()

            for index, record in enumerate(records, start=1):
                approval = record.get("approval") or {}
                status = record.get("status") or {}
                canonical = record.get("canonical") or {}

                review_decision = clean_whitespace(approval.get("reviewDecision")).lower()
                is_excluded = normalize_bool(status.get("excluded"))
                needs_correction = normalize_bool(status.get("needsCorrection"))
                if review_decision and review_decision != "approve":
                    skipped += 1
                    continue
                if is_excluded or needs_correction:
                    skipped += 1
                    continue

                source_kind_value = clean_whitespace(record.get("sourceKind")) or source_kind
                source_name_value = f"{source_name}:{source_kind_value}"
                source_path_value = f"{source_path}:{source_kind_value}"
                source_key = (source_name_value, source_kind_value, source_path_value)
                source_id = source_cache.get(source_key)
                if source_id is None:
                    source_id = _get_or_create_source_id(
                        connection,
                        source_name_value,
                        source_kind_value,
                        source_path_value,
                    )
                    source_cache[source_key] = source_id

                excerpt_text = preserve_excerpt_text(record.get("excerptText"))
                if not excerpt_text:
                    skipped += 1
                    continue

                source_row_number = int(record.get("sourceRow") or index)
                external_id = clean_whitespace(record.get("sourceRecordId") or record.get("recordId"))
                author = clean_whitespace(
                    canonical.get("canonicalAuthor") or record.get("author")
                )
                book_title = clean_whitespace(
                    canonical.get("canonicalBookTitle") or record.get("bookTitle")
                )
                poem_title = clean_whitespace(
                    canonical.get("canonicalPoemTitle")
                    or record.get("poemTitle")
                    or record.get("title")
                )
                metadata_json = json.dumps(record, ensure_ascii=True)
                normalized_author = normalize_lookup_text(author)
                normalized_book_title = normalize_lookup_text(book_title)
                normalized_poem_title = normalize_lookup_text(poem_title)
                normalized_excerpt = normalize_lookup_text(excerpt_text)
                excerpt_hash = fingerprint_excerpt(excerpt_text)

                archive_payload = {
                    "source_name": source_name_value,
                    "source_file": source_path_value,
                    "source_row_number": source_row_number,
                    "record_id": external_id,
                    "timestamp_text": clean_whitespace(
                        record.get("sourceApprovedAt") or record.get("sourceUpdatedAt")
                    ),
                    "email_address": clean_whitespace(record.get("emailAddress")),
                    "request_type": clean_whitespace(record.get("contentType")) or "EXC",
                    "author": author,
                    "title": poem_title,
                    "book_title": book_title,
                    "excerpt_text": excerpt_text,
                    "excerpt_review_decision": review_decision or "approve",
                    "exclude_from_quote_db": "N",
                    "row_json": metadata_json,
                }
                archive_row = cursor.execute(
                    """
                    SELECT id FROM raw_import_rows
                    WHERE source_name = ? AND source_file = ? AND source_row_number = ?
                    LIMIT 1
                    """,
                    (source_name_value, source_path_value, source_row_number),
                ).fetchone()
                if archive_row is None:
                    cursor.execute(
                        """
                        INSERT INTO raw_import_rows (
                            source_name, source_file, source_row_number, record_id,
                            timestamp_text, email_address, request_type, author, title,
                            book_title, excerpt_text, excerpt_review_decision,
                            exclude_from_quote_db, row_json
                        ) VALUES (
                            :source_name, :source_file, :source_row_number, :record_id,
                            :timestamp_text, :email_address, :request_type, :author, :title,
                            :book_title, :excerpt_text, :excerpt_review_decision,
                            :exclude_from_quote_db, :row_json
                        )
                        """,
                        archive_payload,
                    )
                    archived_inserted += 1
                else:
                    cursor.execute(
                        """
                        UPDATE raw_import_rows SET
                            record_id = :record_id,
                            timestamp_text = :timestamp_text,
                            email_address = :email_address,
                            request_type = :request_type,
                            author = :author,
                            title = :title,
                            book_title = :book_title,
                            excerpt_text = :excerpt_text,
                            excerpt_review_decision = :excerpt_review_decision,
                            exclude_from_quote_db = :exclude_from_quote_db,
                            row_json = :row_json
                        WHERE id = :id
                        """,
                        {**archive_payload, "id": int(archive_row[0])},
                    )
                    archived_updated += 1

                existing_row = None
                if external_id:
                    existing_row = cursor.execute(
                        """
                        SELECT id
                        FROM excerpt_entries
                        WHERE source_id = ? AND external_id = ?
                        LIMIT 1
                        """,
                        (source_id, external_id),
                    ).fetchone()
                if existing_row is None:
                    existing_row = cursor.execute(
                        """
                        SELECT id
                        FROM excerpt_entries
                        WHERE source_id = ?
                          AND normalized_author = ?
                          AND normalized_poem_title = ?
                          AND normalized_book_title = ?
                          AND excerpt_hash = ?
                        LIMIT 1
                        """,
                        (
                            source_id,
                            normalized_author,
                            normalized_poem_title,
                            normalized_book_title,
                            excerpt_hash,
                        ),
                    ).fetchone()
                if existing_row is None:
                    existing_row = cursor.execute(
                        """
                        SELECT id
                        FROM excerpt_entries
                        WHERE source_id = ? AND source_row_number = ?
                        LIMIT 1
                        """,
                        (source_id, source_row_number),
                    ).fetchone()

                payload = {
                    "source_id": source_id,
                    "source_row_number": source_row_number,
                    "external_id": external_id,
                    "author": author,
                    "normalized_author": normalized_author,
                    "book_title": book_title,
                    "normalized_book_title": normalized_book_title,
                    "poem_title": poem_title,
                    "normalized_poem_title": normalized_poem_title,
                    "excerpt_text": excerpt_text,
                    "normalized_excerpt": normalized_excerpt,
                    "excerpt_hash": excerpt_hash,
                    "word_count": len(excerpt_text.split()),
                    "character_count": len(excerpt_text),
                    "metadata_json": metadata_json,
                }

                if existing_row is not None:
                    cursor.execute(
                        """
                        UPDATE excerpt_entries
                        SET
                            source_row_number = :source_row_number,
                            external_id = :external_id,
                            author = :author,
                            normalized_author = :normalized_author,
                            book_title = :book_title,
                            normalized_book_title = :normalized_book_title,
                            poem_title = :poem_title,
                            normalized_poem_title = :normalized_poem_title,
                            excerpt_text = :excerpt_text,
                            normalized_excerpt = :normalized_excerpt,
                            excerpt_hash = :excerpt_hash,
                            word_count = :word_count,
                            character_count = :character_count,
                            metadata_json = :metadata_json
                        WHERE id = :id
                        """,
                        {
                            **payload,
                            "id": int(existing_row[0]),
                        },
                    )
                    updated += 1
                else:
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
                        ) VALUES (
                            :source_id,
                            :source_row_number,
                            :external_id,
                            :author,
                            :normalized_author,
                            :book_title,
                            :normalized_book_title,
                            :poem_title,
                            :normalized_poem_title,
                            :excerpt_text,
                            :normalized_excerpt,
                            :excerpt_hash,
                            :word_count,
                            :character_count,
                            :metadata_json
                        )
                        """,
                        payload,
                    )
                    inserted += 1
    finally:
        connection.close()

    return {
        "db_path": str(db_path),
        "source_name": source_name,
        "source_kind": source_kind,
        "source_path": source_path,
        "inserted": inserted,
        "updated": updated,
        "skipped": skipped,
        "archived_inserted": archived_inserted,
        "archived_updated": archived_updated,
    }


def find_library_excerpt_match(
    connection: sqlite3.Connection,
    excerpt_text: str | None,
    book_title: str | None = None,
    author: str | None = None,
    threshold: float = 0.72,
) -> dict | None:
    cleaned_excerpt = clean_whitespace(excerpt_text)
    preserved_excerpt = preserve_excerpt_text(excerpt_text)
    if not cleaned_excerpt:
        return None

    normalized_excerpt = normalize_lookup_text(cleaned_excerpt)
    excerpt_hash = fingerprint_excerpt(cleaned_excerpt)
    book_alias = normalize_lookup_text(book_title)
    author_alias = normalize_lookup_text(author)
    excerpt_len = len(cleaned_excerpt)
    query_line_break_signature = line_break_signature(preserved_excerpt)
    include_metadata = has_metadata_json_column(connection)

    exact_sql = """
        SELECT source_row_number, external_id, author, book_title, poem_title, excerpt_text
        {metadata_select}
        FROM excerpt_entries
        WHERE excerpt_hash = ?
        LIMIT 1
    """.format(
        metadata_select=", metadata_json" if include_metadata else ""
    )
    exact_row = connection.execute(exact_sql, (excerpt_hash,)).fetchone()
    if not exact_row and normalized_excerpt:
        normalized_exact_sql = """
            SELECT source_row_number, external_id, author, book_title, poem_title, excerpt_text
            {metadata_select}
            FROM excerpt_entries
            WHERE normalized_excerpt = ?
            LIMIT 1
        """.format(
            metadata_select=", metadata_json" if include_metadata else ""
        )
        exact_row = connection.execute(normalized_exact_sql, (normalized_excerpt,)).fetchone()
    if exact_row:
        candidate_excerpt = exact_row[5] or ""
        candidate_line_break_signature = line_break_signature(candidate_excerpt)
        status = build_best_library_status(
            excerpt_hash=excerpt_hash,
            metadata_json=exact_row[6] if include_metadata else None,
        )
        return {
            "matchType": "exact",
            "score": 1.0,
            "sourceRow": exact_row[0],
            "recordId": exact_row[1],
            "author": exact_row[2],
            "bookTitle": exact_row[3],
            "poemTitle": exact_row[4],
            "excerptPreview": candidate_excerpt[:180],
            "formattingMatch": preserve_excerpt_text(candidate_excerpt) == preserved_excerpt,
            "lineBreaksMatch": candidate_line_break_signature == query_line_break_signature,
            "queryLineCount": len(query_line_break_signature),
            "matchedLineCount": len(candidate_line_break_signature),
            "libraryStatus": status,
        }

    candidate_sql = """
        SELECT source_row_number, external_id, author, normalized_author, book_title,
               normalized_book_title, poem_title, excerpt_text, character_count
               {metadata_select}
        FROM excerpt_entries
        WHERE character_count BETWEEN ? AND ?
        ORDER BY ABS(character_count - ?), source_row_number
        LIMIT 250
    """.format(
        metadata_select=", metadata_json" if include_metadata else ""
    )
    candidate_rows = connection.execute(
        candidate_sql,
        (max(1, excerpt_len - 160), excerpt_len + 160, excerpt_len),
    ).fetchall()

    for row in candidate_rows:
        candidate_text = row[7] or ""
        if normalize_lookup_text(candidate_text) != normalized_excerpt:
            continue
        candidate_line_break_signature = line_break_signature(candidate_text)
        return {
            "matchType": "exact",
            "score": 1.0,
            "sourceRow": row[0],
            "recordId": row[1],
            "author": row[2],
            "bookTitle": row[4],
            "poemTitle": row[6],
            "excerptPreview": candidate_text[:180],
            "formattingMatch": preserve_excerpt_text(candidate_text) == preserved_excerpt,
            "lineBreaksMatch": candidate_line_break_signature == query_line_break_signature,
            "queryLineCount": len(query_line_break_signature),
            "matchedLineCount": len(candidate_line_break_signature),
            "libraryStatus": build_best_library_status(
                excerpt_hash=fingerprint_excerpt(candidate_text),
                metadata_json=row[9] if include_metadata else None,
            ),
        }

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
            "formattingMatch": False,
            "lineBreaksMatch": False,
            "libraryStatus": build_best_library_status(
                excerpt_hash=fingerprint_excerpt(candidate_text),
                metadata_json=row[9] if include_metadata else None,
            ),
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
