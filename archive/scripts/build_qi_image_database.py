#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import sqlite3
from collections import defaultdict
from pathlib import Path


DEFAULT_DB_PATH = Path("data/excerpt_library_normalized.db")
DEFAULT_OUTPUT_PATH = Path("data/qi_image_database.csv")
DEFAULT_MISSING_OUTPUT_PATH = Path("data/qi_image_database_missing_core_fields.csv")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Build a deduplicated quote-image asset database from the normalized excerpt library."
        )
    )
    parser.add_argument(
        "--db-path",
        type=Path,
        default=DEFAULT_DB_PATH,
        help=f"Normalized excerpt DB path (default: {DEFAULT_DB_PATH})",
    )
    parser.add_argument(
        "--output-csv",
        type=Path,
        default=DEFAULT_OUTPUT_PATH,
        help=f"Output CSV path (default: {DEFAULT_OUTPUT_PATH})",
    )
    parser.add_argument(
        "--missing-csv",
        type=Path,
        default=DEFAULT_MISSING_OUTPUT_PATH,
        help=f"Output CSV path for rows missing core fields (default: {DEFAULT_MISSING_OUTPUT_PATH})",
    )
    return parser.parse_args()


def load_rows(db_path: Path) -> list[sqlite3.Row]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        return conn.execute(
            """
            SELECT
                sr.source_row_number,
                sr.excerpt_id,
                sr.author,
                sr.book_title,
                sr.poem_title,
                sr.excerpt_text,
                sr.metadata_json,
                ea.graphic_made_status,
                ea.approved_for_quote_image,
                ea.used_status,
                ea.file_name,
                ea.link_url,
                ea.low_res_file_name,
                ea.low_res_link_url,
                ea.notes
            FROM source_rows sr
            JOIN excerpt_assets ea
              ON ea.source_row_id = sr.id
            WHERE ea.asset_type = 'QI'
              AND (
                COALESCE(ea.file_name, '') != ''
                OR COALESCE(ea.link_url, '') != ''
                OR COALESCE(ea.low_res_file_name, '') != ''
                OR COALESCE(ea.low_res_link_url, '') != ''
              )
            ORDER BY sr.source_row_number, sr.id
            """
        ).fetchall()
    finally:
        conn.close()


def parse_metadata(metadata_json: str | None) -> dict[str, str]:
    if not metadata_json:
        return {}
    try:
        raw = json.loads(metadata_json)
    except json.JSONDecodeError:
        return {}
    return {
        str(key).strip(): "" if value is None else str(value).strip()
        for key, value in raw.items()
    }


def clean(value: str | None) -> str:
    return (value or "").strip()


def choose(*values: str) -> str:
    for value in values:
        if clean(value):
            return clean(value)
    return ""


def asset_key(row: sqlite3.Row) -> str:
    return choose(
        row["link_url"],
        row["low_res_link_url"],
        row["file_name"],
        row["low_res_file_name"],
        f"source-row-{row['source_row_number']}",
    )


def completeness_score(record: dict[str, str]) -> tuple[int, int]:
    core_fields = [
        "release_year",
        "author_name",
        "book_shortener",
        "book_title",
        "poem_title",
        "excerpt",
        "file_name",
        "link_url",
        "folder_link",
    ]
    filled = sum(1 for field in core_fields if record.get(field))
    excerpt_len = len(record.get("excerpt", ""))
    return (filled, excerpt_len)


def row_to_record(row: sqlite3.Row) -> dict[str, str]:
    metadata = parse_metadata(row["metadata_json"])
    return {
        "release_year": choose(metadata.get("Book Release Year")),
        "author_name": choose(row["author"], metadata.get("Author")),
        "book_shortener": choose(metadata.get("Book Shortener")),
        "book_title": choose(row["book_title"], metadata.get("Book")),
        "poem_title": choose(
            row["poem_title"],
            metadata.get("Title w/o quotes"),
            metadata.get("New Title"),
            metadata.get("Title"),
        ),
        "excerpt": choose(row["excerpt_text"], metadata.get("Quote")),
        "file_name": choose(row["file_name"], metadata.get("File Name")),
        "link_url": choose(row["link_url"], metadata.get("Link")),
        "low_res_file_name": choose(
            row["low_res_file_name"], metadata.get("Low Resolution File Name")
        ),
        "low_res_link_url": choose(
            row["low_res_link_url"],
            metadata.get("(Low Resolution) Direct Link to file (for easy downloading)"),
        ),
        "folder_link": choose(metadata.get("Link to folder on drive")),
        "file_type": choose(metadata.get("File Type")),
        "file_type_helper": choose(metadata.get("File Type Helper")),
        "book_release_catalog": choose(metadata.get("Book Release Catalog")),
        "graphic_made_status": clean(row["graphic_made_status"]),
        "approved_for_quote_image": clean(row["approved_for_quote_image"]),
        "used_status": clean(row["used_status"]),
        "notes": choose(row["notes"], metadata.get("Notes")),
        "alternate_filename": choose(metadata.get("Alternate Filename?")),
        "source_row_number": str(row["source_row_number"]),
        "excerpt_id": str(row["excerpt_id"]),
    }


def build_database(rows: list[sqlite3.Row]) -> list[dict[str, str]]:
    grouped: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in rows:
        grouped[asset_key(row)].append(row_to_record(row))

    output: list[dict[str, str]] = []
    for key, records in grouped.items():
        primary = max(records, key=completeness_score)
        source_rows = sorted({record["source_row_number"] for record in records if record["source_row_number"]}, key=int)
        excerpt_ids = sorted({record["excerpt_id"] for record in records if record["excerpt_id"]}, key=int)
        release_years = sorted({record["release_year"] for record in records if record["release_year"]})

        merged = dict(primary)
        merged["asset_key"] = key
        merged["release_year"] = primary["release_year"] or (release_years[0] if len(release_years) == 1 else "")
        merged["source_row_numbers"] = ", ".join(source_rows)
        merged["excerpt_ids"] = ", ".join(excerpt_ids)
        merged["source_row_count"] = str(len(source_rows))
        merged["release_year_values"] = ", ".join(release_years)
        merged["any_graphic_made_status_y"] = "Y" if any(
            record["graphic_made_status"].upper() == "Y" for record in records
        ) else ""
        merged["any_approved_for_quote_image_y"] = "Y" if any(
            record["approved_for_quote_image"].upper() == "Y" for record in records
        ) else ""
        merged["missing_core_fields"] = ", ".join(
            field
            for field, value in (
                ("release_year", merged["release_year"]),
                ("author_name", merged["author_name"]),
                ("book_shortener", merged["book_shortener"]),
                ("book_title", merged["book_title"]),
                ("poem_title", merged["poem_title"]),
                ("excerpt", merged["excerpt"]),
            )
            if not value
        )
        output.append(merged)

    return sorted(
        output,
        key=lambda row: (
            row["release_year"] or "9999",
            row["author_name"].lower(),
            row["book_title"].lower(),
            row["poem_title"].lower(),
            row["file_name"].lower(),
        ),
    )


def write_csv(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "release_year",
        "author_name",
        "book_shortener",
        "book_title",
        "poem_title",
        "excerpt",
        "file_name",
        "link_url",
        "low_res_file_name",
        "low_res_link_url",
        "folder_link",
        "file_type",
        "file_type_helper",
        "book_release_catalog",
        "graphic_made_status",
        "approved_for_quote_image",
        "used_status",
        "notes",
        "alternate_filename",
        "source_row_number",
        "source_row_numbers",
        "excerpt_id",
        "excerpt_ids",
        "source_row_count",
        "release_year_values",
        "any_graphic_made_status_y",
        "any_approved_for_quote_image_y",
        "missing_core_fields",
        "asset_key",
    ]
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def write_missing_csv(path: Path, rows: list[dict[str, str]]) -> None:
    missing_rows = [row for row in rows if row.get("missing_core_fields")]
    write_csv(path, missing_rows)


def main() -> None:
    args = parse_args()
    rows = load_rows(args.db_path)
    output_rows = build_database(rows)
    write_csv(args.output_csv, output_rows)
    write_missing_csv(args.missing_csv, output_rows)
    print(f"Wrote {len(output_rows)} rows to {args.output_csv}")
    print(
        f"Wrote {sum(1 for row in output_rows if row.get('missing_core_fields'))} "
        f"rows with missing core fields to {args.missing_csv}"
    )


if __name__ == "__main__":
    main()
