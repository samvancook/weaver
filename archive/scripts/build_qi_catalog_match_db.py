#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sqlite3
from pathlib import Path

from build_qi_image_database import build_database, load_rows
from catalog_validate import candidate_snippets, normalize, normalize_title, title_aliases


DEFAULT_QI_DB_PATH = Path("data/excerpt_library_normalized.db")
DEFAULT_CATALOG_DB_PATH = Path(
    "/Users/buttonpublishingone/Desktop/CODEX/Social Media Dev/poetry_catalog/formal_catalog.db"
)
DEFAULT_OUTPUT_DB_PATH = Path("data/qi_catalog_match.db")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a local SQLite database of quote-image assets matched to the formal catalog."
    )
    parser.add_argument("--qi-db-path", type=Path, default=DEFAULT_QI_DB_PATH)
    parser.add_argument("--catalog-db-path", type=Path, default=DEFAULT_CATALOG_DB_PATH)
    parser.add_argument("--output-db", type=Path, default=DEFAULT_OUTPUT_DB_PATH)
    return parser.parse_args()


def clean(value: str | None) -> str:
    return (value or "").strip()


def choose(*values: str) -> str:
    for value in values:
        if clean(value):
            return clean(value)
    return ""


def ensure_schema(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        PRAGMA journal_mode = WAL;

        CREATE TABLE image_assets (
            asset_key TEXT PRIMARY KEY,
            release_year TEXT,
            author_name TEXT,
            book_shortener TEXT,
            book_title TEXT,
            poem_title TEXT,
            excerpt TEXT,
            file_name TEXT,
            link_url TEXT,
            low_res_file_name TEXT,
            low_res_link_url TEXT,
            folder_link TEXT,
            file_type TEXT,
            file_type_helper TEXT,
            book_release_catalog TEXT,
            graphic_made_status TEXT,
            approved_for_quote_image TEXT,
            used_status TEXT,
            notes TEXT,
            alternate_filename TEXT,
            source_row_number TEXT,
            source_row_numbers TEXT,
            excerpt_id TEXT,
            excerpt_ids TEXT,
            source_row_count INTEGER,
            release_year_values TEXT,
            any_graphic_made_status_y TEXT,
            any_approved_for_quote_image_y TEXT,
            missing_core_fields TEXT,
            normalized_author_name TEXT,
            normalized_book_shortener TEXT,
            normalized_book_title TEXT,
            normalized_poem_title TEXT,
            catalog_canonical_book_id INTEGER,
            catalog_book_id INTEGER,
            catalog_poem_id INTEGER,
            book_match_method TEXT,
            book_match_confidence INTEGER,
            poem_match_method TEXT,
            poem_match_confidence INTEGER,
            matched_catalog_author TEXT,
            matched_catalog_book_title TEXT,
            matched_catalog_book_shortener TEXT,
            matched_catalog_release_year TEXT,
            matched_catalog_release_catalog TEXT,
            matched_catalog_poem_title TEXT,
            resolved_author_name TEXT,
            resolved_book_shortener TEXT,
            resolved_book_title TEXT,
            resolved_poem_title TEXT,
            resolved_release_year TEXT
        );

        CREATE INDEX idx_image_assets_book_id ON image_assets(catalog_canonical_book_id);
        CREATE INDEX idx_image_assets_poem_id ON image_assets(catalog_poem_id);
        CREATE INDEX idx_image_assets_norm_book ON image_assets(normalized_book_title);
        CREATE INDEX idx_image_assets_norm_poem ON image_assets(normalized_poem_title);
        """
    )


def load_catalog_books(conn: sqlite3.Connection) -> tuple[list[dict], dict[int, list[dict]]]:
    conn.row_factory = sqlite3.Row
    alias_rows = conn.execute(
        "SELECT canonical_book_id, alias FROM author_aliases"
    ).fetchall()
    author_aliases: dict[int, set[str]] = {}
    for row in alias_rows:
        author_aliases.setdefault(int(row["canonical_book_id"]), set()).add(normalize(row["alias"]))

    canonical_rows = conn.execute(
        """
        SELECT id, title, author, book_shortener, release_catalog, pub_date, book_status, book_release_year
        FROM canonical_books
        """
    ).fetchall()

    catalog_book_rows = conn.execute(
        """
        SELECT id, canonical_book_id, title, author, book_shortener, release_catalog
        FROM catalog_books
        """
    ).fetchall()
    by_canonical: dict[int, list[dict]] = {}
    for row in catalog_book_rows:
        if row["canonical_book_id"] is None:
            continue
        by_canonical.setdefault(int(row["canonical_book_id"]), []).append(dict(row))

    books: list[dict] = []
    for row in canonical_rows:
        canonical_id = int(row["id"])
        aliases = set()
        aliases.update(title_aliases(row["title"]))
        aliases.update(title_aliases(row["book_shortener"]))
        for catalog_book in by_canonical.get(canonical_id, []):
            aliases.update(title_aliases(catalog_book["title"]))
            aliases.update(title_aliases(catalog_book["book_shortener"]))

        author_names = {normalize(row["author"])}
        author_names.update(author_aliases.get(canonical_id, set()))
        books.append(
            {
                "canonical_book_id": canonical_id,
                "title": clean(row["title"]),
                "author": clean(row["author"]),
                "book_shortener": clean(row["book_shortener"]),
                "release_catalog": clean(row["release_catalog"]),
                "book_release_year": clean(row["book_release_year"]),
                "aliases": {alias for alias in aliases if alias},
                "author_aliases": {alias for alias in author_names if alias},
                "catalog_books": by_canonical.get(canonical_id, []),
            }
        )
    return books, by_canonical


def load_catalog_poems(conn: sqlite3.Connection) -> dict[int, list[dict]]:
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        """
        SELECT
            cp.id,
            cp.catalog_book_id,
            cp.title,
            cp.text,
            cb.canonical_book_id
        FROM catalog_poems cp
        JOIN catalog_books cb ON cb.id = cp.catalog_book_id
        """
    ).fetchall()
    poems: dict[int, list[dict]] = {}
    for row in rows:
        if row["canonical_book_id"] is None:
            continue
        poems.setdefault(int(row["canonical_book_id"]), []).append(
            {
                "catalog_poem_id": int(row["id"]),
                "catalog_book_id": int(row["catalog_book_id"]),
                "title": clean(row["title"]),
                "normalized_title": normalize_title(row["title"]),
                "normalized_text": normalize(row["text"]),
            }
        )
    return poems


def score_book_match(record: dict[str, str], book: dict) -> tuple[int, str]:
    score = 0
    method = ""
    book_shortener = normalize_title(record["book_shortener"])
    book_title = normalize_title(record["book_title"])
    author = normalize(record["author_name"])

    if book_shortener and book_shortener == normalize_title(book["book_shortener"]):
        score += 100
        method = "book_shortener_exact"
    elif book_title and book_title == normalize_title(book["title"]):
        score += 85
        method = "book_title_exact"
    elif book_title and book_title in book["aliases"]:
        score += 75
        method = "book_title_alias"

    if author and author == normalize(book["author"]):
        score += 10
        method = f"{method}+author_exact" if method else "author_exact"
    elif author and author in book["author_aliases"]:
        score += 8
        method = f"{method}+author_alias" if method else "author_alias"

    return score, method


def match_book(record: dict[str, str], books: list[dict]) -> dict | None:
    scored = []
    for book in books:
        score, method = score_book_match(record, book)
        if score > 0:
            scored.append((score, method, book))
    if not scored:
        return None
    scored.sort(key=lambda item: (item[0], len(item[2]["title"])), reverse=True)
    best_score, best_method, best_book = scored[0]
    return {
        "score": best_score,
        "method": best_method,
        "book": best_book,
    }


def match_poem(record: dict[str, str], poems: list[dict]) -> dict | None:
    normalized_poem_title = normalize_title(record["poem_title"])
    snippets = candidate_snippets(record["excerpt"])
    scored = []
    for poem in poems:
        score = 0
        method = ""
        if normalized_poem_title and normalized_poem_title == poem["normalized_title"]:
            score = 100
            method = "poem_title_exact"
        elif normalized_poem_title and normalized_poem_title in poem["normalized_title"]:
            score = 85
            method = "poem_title_contains"
        elif snippets and any(snippet in poem["normalized_text"] for snippet in snippets):
            score = 70
            method = "excerpt_snippet"
        if score > 0:
            scored.append((score, method, poem))
    if not scored:
        return None
    scored.sort(key=lambda item: (item[0], len(item[2]["title"])), reverse=True)
    best_score, best_method, best_poem = scored[0]
    return {"score": best_score, "method": best_method, "poem": best_poem}


def resolved_value(record_value: str, matched_value: str) -> str:
    return clean(record_value) or clean(matched_value)


def insert_records(
    conn: sqlite3.Connection,
    records: list[dict[str, str]],
    books: list[dict],
    poems_by_canonical: dict[int, list[dict]],
) -> None:
    insert_sql = """
        INSERT INTO image_assets (
            asset_key, release_year, author_name, book_shortener, book_title, poem_title, excerpt,
            file_name, link_url, low_res_file_name, low_res_link_url, folder_link, file_type,
            file_type_helper, book_release_catalog, graphic_made_status, approved_for_quote_image,
            used_status, notes, alternate_filename, source_row_number, source_row_numbers, excerpt_id,
            excerpt_ids, source_row_count, release_year_values, any_graphic_made_status_y,
            any_approved_for_quote_image_y, missing_core_fields, normalized_author_name,
            normalized_book_shortener, normalized_book_title, normalized_poem_title,
            catalog_canonical_book_id, catalog_book_id, catalog_poem_id, book_match_method,
            book_match_confidence, poem_match_method, poem_match_confidence, matched_catalog_author,
            matched_catalog_book_title, matched_catalog_book_shortener, matched_catalog_release_year,
            matched_catalog_release_catalog, matched_catalog_poem_title, resolved_author_name,
            resolved_book_shortener, resolved_book_title, resolved_poem_title, resolved_release_year
        ) VALUES (
            :asset_key, :release_year, :author_name, :book_shortener, :book_title, :poem_title, :excerpt,
            :file_name, :link_url, :low_res_file_name, :low_res_link_url, :folder_link, :file_type,
            :file_type_helper, :book_release_catalog, :graphic_made_status, :approved_for_quote_image,
            :used_status, :notes, :alternate_filename, :source_row_number, :source_row_numbers, :excerpt_id,
            :excerpt_ids, :source_row_count, :release_year_values, :any_graphic_made_status_y,
            :any_approved_for_quote_image_y, :missing_core_fields, :normalized_author_name,
            :normalized_book_shortener, :normalized_book_title, :normalized_poem_title,
            :catalog_canonical_book_id, :catalog_book_id, :catalog_poem_id, :book_match_method,
            :book_match_confidence, :poem_match_method, :poem_match_confidence, :matched_catalog_author,
            :matched_catalog_book_title, :matched_catalog_book_shortener, :matched_catalog_release_year,
            :matched_catalog_release_catalog, :matched_catalog_poem_title, :resolved_author_name,
            :resolved_book_shortener, :resolved_book_title, :resolved_poem_title, :resolved_release_year
        )
    """

    payloads = []
    for record in records:
        book_match = match_book(record, books)
        matched_book = book_match["book"] if book_match else {}
        poem_match = None
        if book_match:
            poem_match = match_poem(
                record,
                poems_by_canonical.get(matched_book["canonical_book_id"], []),
            )
        matched_poem = poem_match["poem"] if poem_match else {}
        payloads.append(
            {
                **record,
                "normalized_author_name": normalize(record["author_name"]),
                "normalized_book_shortener": normalize_title(record["book_shortener"]),
                "normalized_book_title": normalize_title(record["book_title"]),
                "normalized_poem_title": normalize_title(record["poem_title"]),
                "catalog_canonical_book_id": matched_book.get("canonical_book_id"),
                "catalog_book_id": matched_poem.get("catalog_book_id"),
                "catalog_poem_id": matched_poem.get("catalog_poem_id"),
                "book_match_method": book_match["method"] if book_match else "",
                "book_match_confidence": book_match["score"] if book_match else 0,
                "poem_match_method": poem_match["method"] if poem_match else "",
                "poem_match_confidence": poem_match["score"] if poem_match else 0,
                "matched_catalog_author": matched_book.get("author", ""),
                "matched_catalog_book_title": matched_book.get("title", ""),
                "matched_catalog_book_shortener": matched_book.get("book_shortener", ""),
                "matched_catalog_release_year": matched_book.get("book_release_year", ""),
                "matched_catalog_release_catalog": matched_book.get("release_catalog", ""),
                "matched_catalog_poem_title": matched_poem.get("title", ""),
                "resolved_author_name": resolved_value(record["author_name"], matched_book.get("author", "")),
                "resolved_book_shortener": resolved_value(record["book_shortener"], matched_book.get("book_shortener", "")),
                "resolved_book_title": resolved_value(record["book_title"], matched_book.get("title", "")),
                "resolved_poem_title": resolved_value(record["poem_title"], matched_poem.get("title", "")),
                "resolved_release_year": resolved_value(record["release_year"], matched_book.get("book_release_year", "")),
            }
        )
    conn.executemany(insert_sql, payloads)


def main() -> None:
    args = parse_args()
    rows = load_rows(args.qi_db_path)
    records = build_database(rows)

    if args.output_db.exists():
        args.output_db.unlink()

    catalog_conn = sqlite3.connect(args.catalog_db_path)
    output_conn = sqlite3.connect(args.output_db)
    try:
        ensure_schema(output_conn)
        books, _catalog_books = load_catalog_books(catalog_conn)
        poems_by_canonical = load_catalog_poems(catalog_conn)
        insert_records(output_conn, records, books, poems_by_canonical)
        output_conn.commit()

        total = output_conn.execute("SELECT COUNT(*) FROM image_assets").fetchone()[0]
        matched_books = output_conn.execute(
            "SELECT COUNT(*) FROM image_assets WHERE catalog_canonical_book_id IS NOT NULL"
        ).fetchone()[0]
        matched_poems = output_conn.execute(
            "SELECT COUNT(*) FROM image_assets WHERE catalog_poem_id IS NOT NULL"
        ).fetchone()[0]
        print(f"Wrote {total} image asset rows to {args.output_db}")
        print(f"Matched books: {matched_books}")
        print(f"Matched poems: {matched_poems}")
    finally:
        catalog_conn.close()
        output_conn.close()


if __name__ == "__main__":
    main()
