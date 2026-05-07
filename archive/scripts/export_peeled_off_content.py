#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path

from excerpt_library import DEFAULT_DB_PATH, connect_library

DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parent / "data" / "peeled_off"
DEFAULT_SOURCE_CSV = Path(__file__).resolve().parent / "Primary Excerpt Database - Excerpt Database.csv"
PEELED_OFF_CONTENT_TYPES = {"COV", "ART"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export content types that are being peeled off from the excerpt tool."
    )
    parser.add_argument(
        "--source-db",
        type=Path,
        default=DEFAULT_DB_PATH,
        help=f"Source SQLite database path (default: {DEFAULT_DB_PATH})",
    )
    parser.add_argument(
        "--source-csv",
        type=Path,
        default=DEFAULT_SOURCE_CSV if DEFAULT_SOURCE_CSV.exists() else None,
        help="Optional raw CSV source for peeled-off export",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help=f"Output directory for peeled-off exports (default: {DEFAULT_OUTPUT_DIR})",
    )
    return parser.parse_args()


def resolve_content_type(metadata: dict) -> str:
    return ((metadata.get("File Type") or "") or (metadata.get("File Type Helper") or "")).strip().upper()


def write_grouped_exports(output_dir: Path, grouped_rows: dict[str, list[dict]], source_label: str) -> dict:
    output_dir.mkdir(parents=True, exist_ok=True)
    summary_rows = []
    for content_type in sorted(PEELED_OFF_CONTENT_TYPES):
        export_rows = grouped_rows.get(content_type, [])
        output_csv = output_dir / f"{content_type.lower()}_rows.csv"
        with output_csv.open("w", encoding="utf-8", newline="") as handle:
            fieldnames = list(export_rows[0].keys()) if export_rows else [
                "sourceRow",
                "author",
                "bookTitle",
                "poemTitle",
                "excerptText",
                "contentType",
                "fileName",
                "linkUrl",
                "lowResFileName",
                "lowResLinkUrl",
                "graphicMadeStatus",
                "notes",
                "metadataJson",
            ]
            writer = csv.DictWriter(handle, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(export_rows)

        summary_rows.append(
            {
                "content_type": content_type,
                "row_count": len(export_rows),
                "output_csv": str(output_csv),
            }
        )

    manifest = {
        "source": source_label,
        "output_dir": str(output_dir),
        "peeled_off_content": summary_rows,
    }

    manifest_path = output_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    return manifest


def export_from_db(source_db: Path, output_dir: Path) -> dict:
    connection = connect_library(source_db)
    try:
        rows = connection.execute(
            """
            SELECT *
            FROM excerpt_entries
            ORDER BY source_row_number
            """
        ).fetchall()
    finally:
        connection.close()

    grouped_rows: dict[str, list[dict]] = {content_type: [] for content_type in PEELED_OFF_CONTENT_TYPES}

    for row in rows:
        metadata = json.loads(row["metadata_json"] or "{}")
        content_type = resolve_content_type(metadata)
        if content_type not in grouped_rows:
            continue
        grouped_rows[content_type].append(
            {
                "sourceEntryId": row["id"],
                "sourceId": row["source_id"],
                "sourceRow": row["source_row_number"],
                "author": row["author"] or "",
                "bookTitle": row["book_title"] or "",
                "poemTitle": row["poem_title"] or "",
                "excerptText": row["excerpt_text"],
                "contentType": content_type,
                "fileName": (metadata.get("File Name") or "").strip(),
                "linkUrl": (metadata.get("Link") or "").strip(),
                "lowResFileName": (metadata.get("Low Resolution File Name") or "").strip(),
                "lowResLinkUrl": (metadata.get("(Low Resolution) Direct Link to file (for easy downloading)") or "").strip(),
                "graphicMadeStatus": (metadata.get("Graphic Made?") or "").strip(),
                "notes": (metadata.get("Notes") or "").strip(),
                "metadataJson": row["metadata_json"],
            }
        )

    return write_grouped_exports(output_dir, grouped_rows, source_label=str(source_db))


def export_from_csv(source_csv: Path, output_dir: Path) -> dict:
    grouped_rows: dict[str, list[dict]] = {content_type: [] for content_type in PEELED_OFF_CONTENT_TYPES}

    with source_csv.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row_number, row in enumerate(reader, start=2):
            content_type = resolve_content_type(row)
            if content_type not in grouped_rows:
                continue
            grouped_rows[content_type].append(
                {
                    "sourceRow": row_number,
                    "author": (row.get("Author") or "").strip(),
                    "bookTitle": (row.get("Book") or "").strip(),
                    "poemTitle": (row.get("Title") or "").strip(),
                    "excerptText": (row.get("Quote") or "").strip(),
                    "contentType": content_type,
                    "fileName": (row.get("File Name") or "").strip(),
                    "linkUrl": (row.get("Link") or "").strip(),
                    "lowResFileName": (row.get("Low Resolution File Name") or "").strip(),
                    "lowResLinkUrl": (row.get("(Low Resolution) Direct Link to file (for easy downloading)") or "").strip(),
                    "graphicMadeStatus": (row.get("Graphic Made?") or "").strip(),
                    "notes": (row.get("Notes") or "").strip(),
                    "metadataJson": json.dumps(row, ensure_ascii=True),
                }
            )

    return write_grouped_exports(output_dir, grouped_rows, source_label=str(source_csv))


def export_peeled_off_content(source_db: Path, source_csv: Path | None, output_dir: Path) -> dict:
    result = {}
    db_dir = output_dir / "from_db"
    result["from_db"] = export_from_db(source_db, db_dir)
    if source_csv and source_csv.exists():
        csv_dir = output_dir / "from_raw_csv"
        result["from_raw_csv"] = export_from_csv(source_csv, csv_dir)
    else:
        result["from_raw_csv"] = None
    return result


def main() -> int:
    args = parse_args()
    result = export_peeled_off_content(args.source_db, args.source_csv, args.output_dir)
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
