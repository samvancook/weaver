#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import sqlite3
from pathlib import Path

from build_normalized_excerpt_database import DEFAULT_OUTPUT_DB_PATH


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Build a reconciliation report for approved quote-image rows by matching them to "
            "existing linked graphic assets in the normalized excerpt database."
        )
    )
    parser.add_argument(
        "--db-path",
        type=Path,
        default=DEFAULT_OUTPUT_DB_PATH,
        help=f"Normalized excerpt database path (default: {DEFAULT_OUTPUT_DB_PATH})",
    )
    parser.add_argument(
        "--output-csv",
        type=Path,
        default=Path("data/qi_graphic_reconciliation.csv"),
        help="Output CSV path for the detailed reconciliation report",
    )
    parser.add_argument(
        "--unmatched-csv",
        type=Path,
        help="Optional CSV path for only rows with no detected linked graphic",
    )
    parser.add_argument(
        "--sheet-update-csv",
        type=Path,
        help=(
            "Optional compact CSV of recommended updates keyed by approved source row number "
            "for syncing back into a sheet"
        ),
    )
    return parser.parse_args()


def connect_database(db_path: Path) -> sqlite3.Connection:
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    return connection


def has_asset_file(row: sqlite3.Row) -> bool:
    return any(
        (row[key] or "").strip()
        for key in ("file_name", "link_url", "low_res_file_name", "low_res_link_url")
    )


def normalize_flag(value: str | None) -> str:
    return (value or "").strip().upper()


def excerpt_preview(text: str, length: int = 160) -> str:
    flattened = " ".join((text or "").replace("\r", "\n").split())
    if len(flattened) <= length:
        return flattened
    return flattened[: length - 1].rstrip() + "…"


def choose_primary_match(asset_rows: list[sqlite3.Row]) -> sqlite3.Row | None:
    if not asset_rows:
        return None
    return sorted(
        asset_rows,
        key=lambda row: (
            0 if normalize_flag(row["graphic_made_status"]) == "Y" else 1,
            0 if (row["link_url"] or "").strip() else 1,
            0 if (row["file_name"] or "").strip() else 1,
            row["source_row_number"],
        ),
    )[0]


def fetch_approved_qi_rows(connection: sqlite3.Connection) -> list[sqlite3.Row]:
    return connection.execute(
        """
        SELECT
            sr.id AS source_row_id,
            sr.source_row_number,
            sr.excerpt_id,
            sr.author,
            sr.book_title,
            sr.poem_title,
            sr.excerpt_text,
            ea.asset_type,
            ea.approved_for_quote_image,
            ea.graphic_made_status,
            ea.file_name,
            ea.link_url,
            ea.low_res_file_name,
            ea.low_res_link_url,
            ea.notes
        FROM source_rows sr
        JOIN excerpt_assets ea
          ON ea.source_row_id = sr.id
        WHERE ea.asset_type = 'QI'
          AND ea.approved_for_quote_image = 'Y'
        ORDER BY sr.source_row_number, sr.id
        """
    ).fetchall()


def fetch_linked_qi_assets_by_excerpt(connection: sqlite3.Connection) -> dict[int, list[sqlite3.Row]]:
    rows = connection.execute(
        """
        SELECT
            sr.id AS source_row_id,
            sr.source_row_number,
            sr.excerpt_id,
            sr.author,
            sr.book_title,
            sr.poem_title,
            sr.excerpt_text,
            ea.asset_type,
            ea.approved_for_quote_image,
            ea.graphic_made_status,
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
        ORDER BY sr.excerpt_id, sr.source_row_number, sr.id
        """
    ).fetchall()

    results: dict[int, list[sqlite3.Row]] = {}
    for row in rows:
        excerpt_id = int(row["excerpt_id"])
        results.setdefault(excerpt_id, []).append(row)
    return results


def build_reconciliation_rows(connection: sqlite3.Connection) -> list[dict[str, str]]:
    approved_rows = fetch_approved_qi_rows(connection)
    linked_assets_by_excerpt = fetch_linked_qi_assets_by_excerpt(connection)

    output_rows: list[dict[str, str]] = []
    for approved in approved_rows:
        excerpt_id = int(approved["excerpt_id"])
        direct_file = has_asset_file(approved)
        linked_assets = linked_assets_by_excerpt.get(excerpt_id, [])
        sibling_assets = [
            row for row in linked_assets if int(row["source_row_id"]) != int(approved["source_row_id"])
        ]

        if direct_file:
            match_bucket = "direct_on_approved_row"
            match_confidence = "high"
            primary_match = approved
        elif sibling_assets:
            match_bucket = "sibling_qi_asset_same_excerpt"
            match_confidence = "high"
            primary_match = choose_primary_match(sibling_assets)
        else:
            match_bucket = "no_detected_graphic"
            match_confidence = "none"
            primary_match = None

        matched_assets = sibling_assets if match_bucket == "sibling_qi_asset_same_excerpt" else linked_assets
        matched_assets_json = json.dumps(
            [
                {
                    "sourceRow": int(row["source_row_number"]),
                    "graphicMadeStatus": (row["graphic_made_status"] or "").strip(),
                    "fileName": (row["file_name"] or "").strip(),
                    "linkUrl": (row["link_url"] or "").strip(),
                    "lowResFileName": (row["low_res_file_name"] or "").strip(),
                    "lowResLinkUrl": (row["low_res_link_url"] or "").strip(),
                }
                for row in matched_assets
            ],
            ensure_ascii=True,
        )

        output_rows.append(
            {
                "approvedSourceRow": str(approved["source_row_number"]),
                "excerptId": str(excerpt_id),
                "author": (approved["author"] or "").strip(),
                "bookTitle": (approved["book_title"] or "").strip(),
                "poemTitle": (approved["poem_title"] or "").strip(),
                "excerptPreview": excerpt_preview(approved["excerpt_text"] or ""),
                "matchBucket": match_bucket,
                "matchConfidence": match_confidence,
                "approvedRowGraphicMadeStatus": (approved["graphic_made_status"] or "").strip(),
                "approvedRowHasDirectFile": "Y" if direct_file else "",
                "approvedRowFileName": (approved["file_name"] or "").strip(),
                "approvedRowLinkUrl": (approved["link_url"] or "").strip(),
                "approvedRowLowResFileName": (approved["low_res_file_name"] or "").strip(),
                "approvedRowLowResLinkUrl": (approved["low_res_link_url"] or "").strip(),
                "matchedAssetCount": str(len(matched_assets)),
                "matchedSourceRow": str(primary_match["source_row_number"]) if primary_match else "",
                "matchedGraphicMadeStatus": (primary_match["graphic_made_status"] or "").strip()
                if primary_match
                else "",
                "matchedFileName": (primary_match["file_name"] or "").strip() if primary_match else "",
                "matchedLinkUrl": (primary_match["link_url"] or "").strip() if primary_match else "",
                "matchedLowResFileName": (primary_match["low_res_file_name"] or "").strip()
                if primary_match
                else "",
                "matchedLowResLinkUrl": (primary_match["low_res_link_url"] or "").strip()
                if primary_match
                else "",
                "recommendMarkGraphicMade": "Y"
                if match_bucket in {"direct_on_approved_row", "sibling_qi_asset_same_excerpt"}
                else "",
                "recommendBackfillLinkToApprovedRow": "Y"
                if match_bucket == "sibling_qi_asset_same_excerpt"
                else "",
                "matchedAssetsJson": matched_assets_json,
            }
        )

    return output_rows


def write_csv(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "approvedSourceRow",
        "excerptId",
        "author",
        "bookTitle",
        "poemTitle",
        "excerptPreview",
        "matchBucket",
        "matchConfidence",
        "approvedRowGraphicMadeStatus",
        "approvedRowHasDirectFile",
        "approvedRowFileName",
        "approvedRowLinkUrl",
        "approvedRowLowResFileName",
        "approvedRowLowResLinkUrl",
        "matchedAssetCount",
        "matchedSourceRow",
        "matchedGraphicMadeStatus",
        "matchedFileName",
        "matchedLinkUrl",
        "matchedLowResFileName",
        "matchedLowResLinkUrl",
        "recommendMarkGraphicMade",
        "recommendBackfillLinkToApprovedRow",
        "matchedAssetsJson",
    ]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def build_sheet_update_rows(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    output: list[dict[str, str]] = []
    for row in rows:
        if row["recommendMarkGraphicMade"] != "Y":
            continue
        output.append(
            {
                "approvedSourceRow": row["approvedSourceRow"],
                "matchBucket": row["matchBucket"],
                "matchConfidence": row["matchConfidence"],
                "recommendGraphicMade": "Y",
                "recommendFileName": row["matchedFileName"] or row["approvedRowFileName"],
                "recommendLinkUrl": row["matchedLinkUrl"] or row["approvedRowLinkUrl"],
                "recommendLowResFileName": row["matchedLowResFileName"] or row["approvedRowLowResFileName"],
                "recommendLowResLinkUrl": row["matchedLowResLinkUrl"] or row["approvedRowLowResLinkUrl"],
                "matchedSourceRow": row["matchedSourceRow"],
                "excerptId": row["excerptId"],
                "author": row["author"],
                "bookTitle": row["bookTitle"],
                "poemTitle": row["poemTitle"],
                "excerptPreview": row["excerptPreview"],
            }
        )
    return output


def write_sheet_update_csv(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "approvedSourceRow",
        "matchBucket",
        "matchConfidence",
        "recommendGraphicMade",
        "recommendFileName",
        "recommendLinkUrl",
        "recommendLowResFileName",
        "recommendLowResLinkUrl",
        "matchedSourceRow",
        "excerptId",
        "author",
        "bookTitle",
        "poemTitle",
        "excerptPreview",
    ]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def summarize(rows: list[dict[str, str]]) -> dict[str, int]:
    summary = {
        "approved_rows": len(rows),
        "direct_on_approved_row": 0,
        "sibling_qi_asset_same_excerpt": 0,
        "no_detected_graphic": 0,
        "recommend_mark_graphic_made": 0,
        "recommend_backfill_link_to_approved_row": 0,
    }
    for row in rows:
        summary[row["matchBucket"]] += 1
        if row["recommendMarkGraphicMade"] == "Y":
            summary["recommend_mark_graphic_made"] += 1
        if row["recommendBackfillLinkToApprovedRow"] == "Y":
            summary["recommend_backfill_link_to_approved_row"] += 1
    return summary


def main() -> int:
    args = parse_args()
    connection = connect_database(args.db_path)
    try:
        rows = build_reconciliation_rows(connection)
    finally:
        connection.close()

    write_csv(args.output_csv, rows)

    if args.unmatched_csv:
        unmatched = [row for row in rows if row["matchBucket"] == "no_detected_graphic"]
        write_csv(args.unmatched_csv, unmatched)

    if args.sheet_update_csv:
        write_sheet_update_csv(args.sheet_update_csv, build_sheet_update_rows(rows))

    payload = {
        "db_path": str(args.db_path),
        "output_csv": str(args.output_csv),
        "unmatched_csv": str(args.unmatched_csv) if args.unmatched_csv else None,
        "sheet_update_csv": str(args.sheet_update_csv) if args.sheet_update_csv else None,
        "summary": summarize(rows),
    }
    print(json.dumps(payload, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
