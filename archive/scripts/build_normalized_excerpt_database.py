#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sqlite3
from collections import Counter, defaultdict
from pathlib import Path

from audit_int_rows import classify_int_row
from classify_excerpt_rows import classify_row
from excerpt_library import DEFAULT_DB_PATH, clean_whitespace, connect_library, normalize_lookup_text

DEFAULT_OUTPUT_DB_PATH = Path(__file__).resolve().parent / "data" / "excerpt_library_normalized.db"
PEELED_OFF_CONTENT_TYPES = {"COV", "ART"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Build a normalized excerpt database where canonical excerpts are primary records "
            "and QI/INT/COV rows are linked assets or source-row metadata."
        )
    )
    parser.add_argument(
        "--source-db",
        type=Path,
        default=DEFAULT_DB_PATH,
        help=f"Source SQLite database path (default: {DEFAULT_DB_PATH})",
    )
    parser.add_argument(
        "--output-db",
        type=Path,
        default=DEFAULT_OUTPUT_DB_PATH,
        help=f"Output SQLite database path (default: {DEFAULT_OUTPUT_DB_PATH})",
    )
    return parser.parse_args()


def ensure_output_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        PRAGMA journal_mode = WAL;

        CREATE TABLE IF NOT EXISTS excerpts (
            id INTEGER PRIMARY KEY,
            excerpt_hash TEXT NOT NULL UNIQUE,
            canonical_excerpt_text TEXT NOT NULL,
            normalized_excerpt TEXT NOT NULL,
            pull_count INTEGER NOT NULL,
            primary_author TEXT,
            primary_book_title TEXT,
            primary_poem_title TEXT,
            classification_basis TEXT NOT NULL,
            source_row_count INTEGER NOT NULL,
            qi_asset_count INTEGER NOT NULL DEFAULT 0,
            qi_linked_asset_count INTEGER NOT NULL DEFAULT 0,
            qi_approved_count INTEGER NOT NULL DEFAULT 0,
            qi_graphic_made_count INTEGER NOT NULL DEFAULT 0,
            int_asset_count INTEGER NOT NULL DEFAULT 0,
            cov_asset_count INTEGER NOT NULL DEFAULT 0,
            other_asset_count INTEGER NOT NULL DEFAULT 0,
            has_qi_asset INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        CREATE INDEX IF NOT EXISTS idx_excerpts_pull_count
            ON excerpts(pull_count DESC);

        CREATE TABLE IF NOT EXISTS source_rows (
            id INTEGER PRIMARY KEY,
            source_entry_id INTEGER NOT NULL UNIQUE,
            source_id INTEGER NOT NULL,
            source_row_number INTEGER NOT NULL,
            excerpt_id INTEGER REFERENCES excerpts(id) ON DELETE SET NULL,
            author TEXT,
            book_title TEXT,
            poem_title TEXT,
            excerpt_text TEXT NOT NULL,
            content_type TEXT,
            content_type_helper TEXT,
            classification TEXT NOT NULL,
            classification_confidence REAL NOT NULL,
            classification_reasons_json TEXT NOT NULL,
            metadata_json TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_source_rows_excerpt_id
            ON source_rows(excerpt_id);
        CREATE INDEX IF NOT EXISTS idx_source_rows_content_type
            ON source_rows(content_type);
        CREATE INDEX IF NOT EXISTS idx_source_rows_classification
            ON source_rows(classification);

        CREATE TABLE IF NOT EXISTS excerpt_assets (
            id INTEGER PRIMARY KEY,
            excerpt_id INTEGER REFERENCES excerpts(id) ON DELETE SET NULL,
            source_row_id INTEGER NOT NULL UNIQUE REFERENCES source_rows(id) ON DELETE CASCADE,
            asset_type TEXT NOT NULL,
            graphic_made_status TEXT,
            approved_for_quote_image TEXT,
            used_status TEXT,
            file_name TEXT,
            link_url TEXT,
            low_res_file_name TEXT,
            low_res_link_url TEXT,
            notes TEXT
        );

        CREATE INDEX IF NOT EXISTS idx_excerpt_assets_excerpt_id
            ON excerpt_assets(excerpt_id);
        CREATE INDEX IF NOT EXISTS idx_excerpt_assets_type
            ON excerpt_assets(asset_type);

        CREATE TABLE IF NOT EXISTS peeled_off_rows (
            id INTEGER PRIMARY KEY,
            source_entry_id INTEGER NOT NULL UNIQUE,
            source_id INTEGER NOT NULL,
            source_row_number INTEGER NOT NULL,
            peeled_off_category TEXT NOT NULL,
            author TEXT,
            book_title TEXT,
            poem_title TEXT,
            excerpt_text TEXT NOT NULL,
            content_type TEXT,
            metadata_json TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_peeled_off_rows_category
            ON peeled_off_rows(peeled_off_category);
        """
    )


def choose_preferred_value(rows: list[sqlite3.Row], field_name: str) -> str:
    counter: Counter[str] = Counter()
    display_by_norm: dict[str, str] = {}
    for row in rows:
        value = clean_whitespace(row[field_name])
        norm = normalize_lookup_text(value)
        if not norm:
            continue
        counter[norm] += 1
        if norm not in display_by_norm or len(value) > len(display_by_norm[norm]):
            display_by_norm[norm] = value

    if not counter:
        return ""

    best_norm = max(
        counter,
        key=lambda key: (
            counter[key],
            len(display_by_norm[key]),
            display_by_norm[key].lower(),
        ),
    )
    return display_by_norm[best_norm]


def classify_entry(row: sqlite3.Row, metadata: dict) -> tuple[str, float, list[str]]:
    file_type = clean_whitespace(metadata.get("File Type", ""))
    file_type_helper = clean_whitespace(metadata.get("File Type Helper", ""))
    return classify_row(
        {
            "excerpt_text": row["excerpt_text"],
            "poem_title": row["poem_title"],
            "word_count": row["word_count"],
            "file_type": file_type,
            "file_type_helper": file_type_helper,
        }
    )


def resolve_content_type(metadata: dict) -> str:
    primary = clean_whitespace(metadata.get("File Type", ""))
    helper = clean_whitespace(metadata.get("File Type Helper", ""))
    return (primary or helper).upper()


def row_has_asset_metadata(metadata: dict) -> bool:
    return any(
        clean_whitespace(metadata.get(key, ""))
        for key in [
            "File Name",
            "Link",
            "Low Resolution File Name",
            "(Low Resolution) Direct Link to file (for easy downloading)",
        ]
    )


def resolve_peeled_off_category(metadata: dict) -> str | None:
    content_type = resolve_content_type(metadata)
    if content_type in PEELED_OFF_CONTENT_TYPES:
        return content_type

    if content_type == "INT":
        bucket, _reasons = classify_int_row(
            {
                "Quote": metadata.get("Quote", ""),
                "Full page or Excerpt": metadata.get("Full page or Excerpt", ""),
            }
        )
        if bucket == "peel_off_layout_int":
            return "INT_LAYOUT"

    return None


def build_normalized_database(source_db: Path, output_db: Path) -> dict:
    if output_db.exists():
        output_db.unlink()

    source_connection = connect_library(source_db)
    output_db.parent.mkdir(parents=True, exist_ok=True)
    output_connection = sqlite3.connect(output_db)
    try:
        output_connection.row_factory = sqlite3.Row
        ensure_output_schema(output_connection)

        raw_rows = source_connection.execute(
            """
            SELECT *
            FROM excerpt_entries
            ORDER BY source_row_number, id
            """
        ).fetchall()

        prepared_rows: list[dict] = []
        group_rows: dict[str, list[sqlite3.Row]] = defaultdict(list)
        excluded_non_excerpt_rows = 0
        peeled_off_row_count = 0

        for row in raw_rows:
            metadata = json.loads(row["metadata_json"] or "{}")
            classification, confidence, reasons = classify_entry(row, metadata)
            content_type = resolve_content_type(metadata)
            peeled_off_category = resolve_peeled_off_category(metadata)

            prepared_rows.append(
                {
                    "row": row,
                    "metadata": metadata,
                    "classification": classification,
                    "confidence": confidence,
                    "reasons": reasons,
                    "content_type": content_type,
                    "peeled_off_category": peeled_off_category,
                }
            )

            if peeled_off_category:
                peeled_off_row_count += 1
            elif classification != "likely_non_excerpt":
                group_rows[row["excerpt_hash"]].append(row)
            else:
                excluded_non_excerpt_rows += 1

        excerpt_id_by_hash: dict[str, int] = {}
        asset_counts_by_excerpt_id: dict[int, Counter[str]] = defaultdict(Counter)
        source_row_id_by_source_entry_id: dict[int, int] = {}

        with output_connection:
            for excerpt_hash, rows in group_rows.items():
                if not rows:
                    continue
                primary_author = choose_preferred_value(rows, "author")
                primary_book_title = choose_preferred_value(rows, "book_title")
                primary_poem_title = choose_preferred_value(rows, "poem_title")
                canonical_excerpt_text = choose_preferred_value(rows, "excerpt_text") or rows[0]["excerpt_text"]

                cursor = output_connection.execute(
                    """
                    INSERT INTO excerpts (
                        excerpt_hash,
                        canonical_excerpt_text,
                        normalized_excerpt,
                        pull_count,
                        primary_author,
                        primary_book_title,
                        primary_poem_title,
                        classification_basis,
                        source_row_count
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        excerpt_hash,
                        canonical_excerpt_text,
                        normalize_lookup_text(canonical_excerpt_text),
                        len(rows),
                        primary_author,
                        primary_book_title,
                        primary_poem_title,
                        "exclude_likely_non_excerpt_rows",
                        len(rows),
                    ),
                )
                excerpt_id_by_hash[excerpt_hash] = cursor.lastrowid

            for prepared in prepared_rows:
                row = prepared["row"]
                metadata = prepared["metadata"]
                content_type = prepared["content_type"]
                peeled_off_category = prepared["peeled_off_category"]

                if peeled_off_category:
                    output_connection.execute(
                        """
                        INSERT INTO peeled_off_rows (
                            source_entry_id,
                            source_id,
                            source_row_number,
                            peeled_off_category,
                            author,
                            book_title,
                            poem_title,
                            excerpt_text,
                            content_type,
                            metadata_json
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            int(row["id"]),
                            int(row["source_id"]),
                            int(row["source_row_number"]),
                            peeled_off_category,
                            row["author"] or "",
                            row["book_title"] or "",
                            row["poem_title"] or "",
                            row["excerpt_text"],
                            clean_whitespace(metadata.get("File Type", "")),
                            row["metadata_json"],
                        ),
                    )
                    continue

                excerpt_id = excerpt_id_by_hash.get(row["excerpt_hash"])

                cursor = output_connection.execute(
                    """
                    INSERT INTO source_rows (
                        source_entry_id,
                        source_id,
                        source_row_number,
                        excerpt_id,
                        author,
                        book_title,
                        poem_title,
                        excerpt_text,
                        content_type,
                        content_type_helper,
                        classification,
                        classification_confidence,
                        classification_reasons_json,
                        metadata_json
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        int(row["id"]),
                        int(row["source_id"]),
                        int(row["source_row_number"]),
                        excerpt_id,
                        row["author"] or "",
                        row["book_title"] or "",
                        row["poem_title"] or "",
                        row["excerpt_text"],
                        clean_whitespace(metadata.get("File Type", "")),
                        clean_whitespace(metadata.get("File Type Helper", "")),
                        prepared["classification"],
                        float(prepared["confidence"]),
                        json.dumps(prepared["reasons"], ensure_ascii=True),
                        row["metadata_json"],
                    ),
                )
                source_row_id = cursor.lastrowid
                source_row_id_by_source_entry_id[int(row["id"])] = source_row_id

                is_asset_type = content_type in {"QI", "INT", "COV", "ART"} or row_has_asset_metadata(metadata)
                if not is_asset_type:
                    continue

                output_connection.execute(
                    """
                    INSERT INTO excerpt_assets (
                        excerpt_id,
                        source_row_id,
                        asset_type,
                        graphic_made_status,
                        approved_for_quote_image,
                        used_status,
                        file_name,
                        link_url,
                        low_res_file_name,
                        low_res_link_url,
                        notes
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        excerpt_id,
                        source_row_id,
                        content_type or "ASSET",
                        clean_whitespace(metadata.get("Graphic Made?", "")),
                        clean_whitespace(metadata.get("Approved for Quote Image", "")),
                        clean_whitespace(metadata.get("USED", "")),
                        clean_whitespace(metadata.get("File Name", "")),
                        clean_whitespace(metadata.get("Link", "")),
                        clean_whitespace(metadata.get("Low Resolution File Name", "")),
                        clean_whitespace(metadata.get("(Low Resolution) Direct Link to file (for easy downloading)", "")),
                        clean_whitespace(metadata.get("Notes", "")),
                    ),
                )
                if excerpt_id is not None:
                    asset_counts_by_excerpt_id[excerpt_id][content_type or "ASSET"] += 1
                    if content_type == "QI":
                        if clean_whitespace(metadata.get("Link", "")):
                            asset_counts_by_excerpt_id[excerpt_id]["QI_LINKED"] += 1
                        if clean_whitespace(metadata.get("Approved for Quote Image", "")) == "Y":
                            asset_counts_by_excerpt_id[excerpt_id]["QI_APPROVED"] += 1
                        if clean_whitespace(metadata.get("Graphic Made?", "")):
                            asset_counts_by_excerpt_id[excerpt_id]["QI_GRAPHIC_MADE"] += 1

            for excerpt_id, counts in asset_counts_by_excerpt_id.items():
                qi_count = counts.get("QI", 0)
                qi_linked_count = counts.get("QI_LINKED", 0)
                qi_approved_count = counts.get("QI_APPROVED", 0)
                qi_graphic_made_count = counts.get("QI_GRAPHIC_MADE", 0)
                int_count = counts.get("INT", 0)
                cov_count = counts.get("COV", 0)
                other_count = sum(
                    count
                    for key, count in counts.items()
                    if key not in {"QI", "QI_LINKED", "QI_APPROVED", "QI_GRAPHIC_MADE", "INT", "COV"}
                )
                output_connection.execute(
                    """
                    UPDATE excerpts
                    SET
                        qi_asset_count = ?,
                        qi_linked_asset_count = ?,
                        qi_approved_count = ?,
                        qi_graphic_made_count = ?,
                        int_asset_count = ?,
                        cov_asset_count = ?,
                        other_asset_count = ?,
                        has_qi_asset = ?
                    WHERE id = ?
                    """,
                    (
                        qi_count,
                        qi_linked_count,
                        qi_approved_count,
                        qi_graphic_made_count,
                        int_count,
                        cov_count,
                        other_count,
                        1 if qi_count > 0 else 0,
                        excerpt_id,
                    ),
                )
    finally:
        output_connection.close()
        source_connection.close()

    normalized_connection = sqlite3.connect(output_db)
    try:
        counts = {
            "excerpt_count": normalized_connection.execute("SELECT COUNT(*) FROM excerpts").fetchone()[0],
            "source_row_count": normalized_connection.execute("SELECT COUNT(*) FROM source_rows").fetchone()[0],
            "asset_count": normalized_connection.execute("SELECT COUNT(*) FROM excerpt_assets").fetchone()[0],
            "excerpts_with_qi_asset": normalized_connection.execute(
                "SELECT COUNT(*) FROM excerpts WHERE has_qi_asset = 1"
            ).fetchone()[0],
            "unlinked_assets": normalized_connection.execute(
                "SELECT COUNT(*) FROM excerpt_assets WHERE excerpt_id IS NULL"
            ).fetchone()[0],
            "excluded_non_excerpt_rows": excluded_non_excerpt_rows,
            "peeled_off_row_count": normalized_connection.execute(
                "SELECT COUNT(*) FROM peeled_off_rows"
            ).fetchone()[0],
        }
    finally:
        normalized_connection.close()

    return {
        "source_db": str(source_db),
        "output_db": str(output_db),
        **counts,
    }


def main() -> int:
    args = parse_args()
    result = build_normalized_database(args.source_db, args.output_db)
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
