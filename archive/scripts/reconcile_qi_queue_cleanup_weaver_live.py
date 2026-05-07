#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import sqlite3
import subprocess
from collections import Counter, defaultdict
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from excerpt_library import fingerprint_excerpt, normalize_lookup_text
from reconcile_qi_queue_cleanup import DEFAULT_WORKBOOK_PATH, XML_NS, col_to_num, load_shared_strings
from sync_qi_queue_cleanup_updates import fetch_live_rows, refresh_access_token, api_request, SPREADSHEET_ID, SHEET_NAME

DEFAULT_QI_DATABASE_CSV = Path("data/qi_image_database.csv")
DEFAULT_NORMALIZED_DB = Path("data/excerpt_library_normalized.db")
DEFAULT_OUTPUT_CSV = Path("data/qi_queue_cleanup_weaver_live_matches.csv")
DEFAULT_UPDATES_CSV = Path("data/qi_queue_cleanup_weaver_live_updates.csv")
WEAVER_QUEUE_URL = "https://weaver.buttonpoetry.com/api/pig/graphics-request-records?filter=all"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Match Weaver's current live graphics queue against the QI database and mark safe "
            "rows as already made by setting AQ=Y on the live source sheet."
        )
    )
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK_PATH)
    parser.add_argument("--qi-csv", type=Path, default=DEFAULT_QI_DATABASE_CSV)
    parser.add_argument("--normalized-db", type=Path, default=DEFAULT_NORMALIZED_DB)
    parser.add_argument("--output-csv", type=Path, default=DEFAULT_OUTPUT_CSV)
    parser.add_argument("--updates-csv", type=Path, default=DEFAULT_UPDATES_CSV)
    parser.add_argument("--min-meta-score", type=int, default=2, choices=[0, 1, 2, 3])
    parser.add_argument("--apply", action="store_true", help="Write AQ=Y for verified safe matches.")
    return parser.parse_args()


def clean_flag(value: str | None) -> str:
    return (value or "").strip().upper()


def clean_text(value: str | None) -> str:
    return (value or "").replace("\r\n", "\n").replace("\r", "\n").strip()


def normalize_text(value: str | None) -> str:
    text = clean_text(value).lower()
    text = text.replace("’", "'").replace("‘", "'")
    text = text.replace("“", '"').replace("”", '"')
    text = text.replace("—", "-").replace("–", "-")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def iter_sheet_rows(archive: ZipFile, sheet_path: str) -> Iterable[tuple[int, dict[int, str]]]:
    shared_strings = load_shared_strings(archive)
    root = ET.fromstring(archive.read(sheet_path))

    for row in root.findall(".//a:sheetData/a:row", XML_NS):
        row_number = int(row.attrib["r"])
        values: dict[int, str] = {}
        for cell in row.findall("a:c", XML_NS):
            reference = cell.attrib.get("r", "")
            match = re.match(r"([A-Z]+)(\d+)", reference)
            if not match:
                continue

            column_number = col_to_num(match.group(1))
            cell_type = cell.attrib.get("t")
            if cell_type == "inlineStr":
                node = cell.find("a:is/a:t", XML_NS)
                value = node.text if node is not None else ""
            elif cell_type == "s":
                node = cell.find("a:v", XML_NS)
                value = shared_strings[int(node.text)] if node is not None and node.text else ""
            else:
                node = cell.find("a:v", XML_NS)
                value = node.text if node is not None else ""
            values[column_number] = value
        yield row_number, values


def build_local_source_rows(workbook_path: Path) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    with ZipFile(workbook_path) as archive:
        for row_number, row in iter_sheet_rows(archive, "xl/worksheets/sheet2.xml"):
            if row_number == 1:
                continue
            author = clean_text(row.get(3))
            poem_title = clean_text(row.get(4))
            book_title = clean_text(row.get(6))
            excerpt = clean_text(row.get(7))
            if not book_title or not excerpt:
                continue
            rows.append(
                {
                    "sheetRow": str(row_number),
                    "author": author,
                    "poemTitle": poem_title,
                    "bookTitle": book_title,
                    "excerpt": excerpt,
                    "aqOverride": clean_flag(row.get(43)),
                }
            )
    return rows


def build_source_lookup(rows: list[dict[str, str]]) -> dict[str, list[dict[str, str]]]:
    by_key: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in rows:
        key = build_row_key(row["author"], row["poemTitle"], row["bookTitle"], row["excerpt"])
        by_key[key].append(row)
    return by_key


def build_row_key(author: str, poem_title: str, book_title: str, excerpt: str) -> str:
    payload = "\u241f".join(
        [
            normalize_text(author),
            normalize_text(poem_title),
            normalize_text(book_title),
            normalize_text(excerpt),
        ]
    )
    return hashlib.sha1(payload.encode("utf-8")).hexdigest()


def fetch_weaver_queue_records() -> list[dict[str, str]]:
    result = subprocess.run(
        ["curl", "-sS", WEAVER_QUEUE_URL],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or f"curl failed with exit {result.returncode}")
    payload = json.loads(result.stdout or "{}")
    return payload.get("requests", [])


def load_qi_rows(qi_csv_path: Path) -> dict[str, list[dict[str, str]]]:
    rows_by_hash: dict[str, list[dict[str, str]]] = defaultdict(list)
    with qi_csv_path.open(newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            if not (
                row.get("link_url")
                or row.get("file_name")
                or row.get("low_res_link_url")
                or row.get("low_res_file_name")
            ):
                continue
            excerpt_hash = fingerprint_excerpt(row.get("excerpt", ""))
            rows_by_hash[excerpt_hash].append(row)
    return rows_by_hash


def load_normalized_qi_rows(db_path: Path) -> dict[str, list[dict[str, str]]]:
    if not db_path.exists():
        return {}

    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    try:
        rows = connection.execute(
            """
            SELECT
              e.excerpt_hash,
              e.primary_author,
              e.primary_book_title,
              e.primary_poem_title,
              e.qi_asset_count,
              e.qi_linked_asset_count,
              e.qi_graphic_made_count,
              a.file_name,
              a.link_url,
              a.low_res_file_name,
              a.low_res_link_url,
              a.graphic_made_status,
              a.approved_for_quote_image
            FROM excerpts e
            LEFT JOIN excerpt_assets a
              ON a.excerpt_id = e.id
             AND a.asset_type = 'QI'
            WHERE e.qi_linked_asset_count > 0 OR e.qi_graphic_made_count > 0
            """
        ).fetchall()
    finally:
        connection.close()

    rows_by_hash: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in rows:
        rows_by_hash[row["excerpt_hash"]].append(
            {
                "author_name": row["primary_author"] or "",
                "book_title": row["primary_book_title"] or "",
                "poem_title": row["primary_poem_title"] or "",
                "file_name": row["file_name"] or row["low_res_file_name"] or "",
                "link_url": row["link_url"] or row["low_res_link_url"] or "",
                "low_res_file_name": row["low_res_file_name"] or "",
                "low_res_link_url": row["low_res_link_url"] or "",
                "graphic_made_status": row["graphic_made_status"] or "",
                "approved_for_quote_image": row["approved_for_quote_image"] or "",
                "qiAssetCount": str(row["qi_asset_count"] or 0),
                "qiLinkedAssetCount": str(row["qi_linked_asset_count"] or 0),
                "qiGraphicMadeCount": str(row["qi_graphic_made_count"] or 0),
                "matchSource": "normalized_db",
            }
        )
    return rows_by_hash


def score_metadata(candidate: dict[str, str], qi_row: dict[str, str]) -> tuple[int, dict[str, bool]]:
    author_match = bool(
        normalize_lookup_text(candidate["author"])
        and normalize_lookup_text(candidate["author"]) == normalize_lookup_text(qi_row.get("author_name", ""))
    )
    book_match = bool(
        normalize_lookup_text(candidate["bookTitle"])
        and normalize_lookup_text(candidate["bookTitle"]) == normalize_lookup_text(qi_row.get("book_title", ""))
    )
    poem_match = bool(
        normalize_lookup_text(candidate["poemTitle"])
        and normalize_lookup_text(candidate["poemTitle"]) == normalize_lookup_text(qi_row.get("poem_title", ""))
    )
    return int(author_match) + int(book_match) + int(poem_match), {
        "authorMatch": bool(author_match),
        "bookMatch": bool(book_match),
        "poemMatch": bool(poem_match),
    }


def verify_live_rows(candidate_rows: list[dict[str, str]]) -> tuple[list[dict[str, str]], Counter]:
    access_token = refresh_access_token()
    live_rows = fetch_live_rows(access_token, [int(row["sheetRow"]) for row in candidate_rows])
    verified: list[dict[str, str]] = []
    counts: Counter = Counter()

    for row in candidate_rows:
        live = live_rows.get(int(row["sheetRow"]))
        if not live:
            counts["missingLiveRow"] += 1
            continue
        if clean_flag(live.get("aqOverride")) == "Y":
            counts["alreadyOverridden"] += 1
            continue

        checks = [
            normalize_text(row["author"]) == normalize_text(live.get("author")),
            normalize_text(row["bookTitle"]) == normalize_text(live.get("bookTitle")),
            normalize_text(row["poemTitle"]) == normalize_text(live.get("poemTitle")),
            normalize_text(row["excerpt"]) == normalize_text(live.get("excerptText")),
        ]
        if not all(checks):
            counts["liveMismatch"] += 1
            continue
        verified.append(row)

    return verified, counts


def apply_updates(rows: list[dict[str, str]]) -> dict:
    access_token = refresh_access_token()
    payload = {
        "valueInputOption": "USER_ENTERED",
        "data": [
            {
                "range": f"'{SHEET_NAME}'!AQ{row['sheetRow']}",
                "majorDimension": "ROWS",
                "values": [["Y"]],
            }
            for row in rows
        ],
    }
    return api_request(
        access_token,
        "POST",
        f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}/values:batchUpdate",
        payload=payload,
    )


def write_csv(path: Path, rows: list[dict[str, str]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    args = parse_args()
    queue_records = fetch_weaver_queue_records()
    source_lookup = build_source_lookup(build_local_source_rows(args.workbook))
    qi_rows_by_hash = load_qi_rows(args.qi_csv)
    normalized_qi_rows_by_hash = load_normalized_qi_rows(args.normalized_db)

    audit_rows: list[dict[str, str]] = []
    update_candidates: list[dict[str, str]] = []
    counts: Counter = Counter()

    for record in queue_records:
        key = build_row_key(
            record.get("author", ""),
            record.get("poemTitle", ""),
            record.get("bookTitle", ""),
            record.get("quoteText", ""),
        )
        source_matches = source_lookup.get(key, [])
        if not source_matches:
            counts["noLocalSourceMatch"] += 1
            continue
        if len(source_matches) > 1:
            counts["ambiguousLocalSourceMatch"] += 1
            continue

        source_row = source_matches[0]
        excerpt_hash = fingerprint_excerpt(record.get("quoteText", ""))
        qi_rows = qi_rows_by_hash.get(excerpt_hash, [])
        match_source = "qi_csv"
        if not qi_rows:
            qi_rows = normalized_qi_rows_by_hash.get(excerpt_hash, [])
            match_source = "normalized_db"
        if not qi_rows:
            counts["noExactExcerpt"] += 1
            continue

        best_score = -1
        best_flags = {"authorMatch": False, "bookMatch": False, "poemMatch": False}
        best_row: dict[str, str] | None = None
        for qi_row in qi_rows:
            score, flags = score_metadata(
                {
                    "author": record.get("author", ""),
                    "bookTitle": record.get("bookTitle", ""),
                    "poemTitle": record.get("poemTitle", ""),
                },
                qi_row,
            )
            if score > best_score:
                best_score = score
                best_flags = flags
                best_row = qi_row

        counts[f"metaScore{best_score}"] += 1
        audit_rows.append(
            {
                "queueSheetRow": str(record.get("queueSheetRow", "")),
                "sheetRow": source_row["sheetRow"],
                "graphicsRequestId": str(record.get("graphicsRequestId", "")),
                "recordId": str(record.get("recordId", "")),
                "author": record.get("author", ""),
                "bookTitle": record.get("bookTitle", ""),
                "poemTitle": record.get("poemTitle", ""),
                "excerpt": record.get("quoteText", ""),
                "excerptHash": excerpt_hash,
                "matchCount": str(len(qi_rows)),
                "bestMetaScore": str(best_score),
                "authorMatch": "Y" if best_flags["authorMatch"] else "",
                "bookMatch": "Y" if best_flags["bookMatch"] else "",
                "poemMatch": "Y" if best_flags["poemMatch"] else "",
                "qiAuthor": (best_row or {}).get("author_name", ""),
                "qiBookTitle": (best_row or {}).get("book_title", ""),
                "qiPoemTitle": (best_row or {}).get("poem_title", ""),
                "qiFileName": (best_row or {}).get("file_name", ""),
                "qiLinkUrl": (best_row or {}).get("link_url", ""),
                "matchSource": match_source,
            }
        )

        if best_score >= args.min_meta_score:
            update_candidates.append(
                {
                    "sheetRow": source_row["sheetRow"],
                    "queueSheetRow": str(record.get("queueSheetRow", "")),
                    "graphicsRequestId": str(record.get("graphicsRequestId", "")),
                    "recordId": str(record.get("recordId", "")),
                    "author": record.get("author", ""),
                    "bookTitle": record.get("bookTitle", ""),
                    "poemTitle": record.get("poemTitle", ""),
                    "excerpt": record.get("quoteText", ""),
                    "bestMetaScore": str(best_score),
                    "qiFileName": (best_row or {}).get("file_name", ""),
                    "qiLinkUrl": (best_row or {}).get("link_url", ""),
                    "matchSource": match_source,
                }
            )

    verified_updates, verify_counts = verify_live_rows(update_candidates)
    counts.update(verify_counts)

    audit_fieldnames = [
        "queueSheetRow",
        "sheetRow",
        "graphicsRequestId",
        "recordId",
        "author",
        "bookTitle",
        "poemTitle",
        "excerpt",
        "excerptHash",
        "matchCount",
        "bestMetaScore",
        "authorMatch",
        "bookMatch",
        "poemMatch",
        "qiAuthor",
        "qiBookTitle",
        "qiPoemTitle",
        "qiFileName",
        "qiLinkUrl",
        "matchSource",
    ]
    write_csv(args.output_csv, audit_rows, audit_fieldnames)

    update_fieldnames = [
        "sheetRow",
        "queueSheetRow",
        "graphicsRequestId",
        "recordId",
        "author",
        "bookTitle",
        "poemTitle",
        "bestMetaScore",
        "qiFileName",
        "qiLinkUrl",
        "matchSource",
    ]
    write_csv(args.updates_csv, verified_updates, update_fieldnames)

    summary = {
        "queueRows": len(queue_records),
        "matchedAuditRows": len(audit_rows),
        "safeUpdates": len(verified_updates),
        "counts": dict(counts),
        "auditCsv": str(args.output_csv),
        "updatesCsv": str(args.updates_csv),
        "applied": False,
    }

    if args.apply and verified_updates:
        response = apply_updates(verified_updates)
        summary["applied"] = True
        summary["totalUpdatedCells"] = response.get("totalUpdatedCells")
        summary["totalUpdatedRows"] = response.get("totalUpdatedRows")
        summary["totalUpdatedSheets"] = response.get("totalUpdatedSheets")

    print(json.dumps(summary, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
