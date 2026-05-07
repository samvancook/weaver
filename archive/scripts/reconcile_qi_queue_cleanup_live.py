#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import urllib.parse
from collections import Counter, defaultdict
from pathlib import Path

from excerpt_library import fingerprint_excerpt, normalize_lookup_text
from sync_qi_queue_cleanup_updates import SPREADSHEET_ID, SHEET_NAME, api_request, refresh_access_token

DEFAULT_QI_DATABASE_CSV = Path("data/qi_image_database.csv")
DEFAULT_OUTPUT_CSV = Path("data/qi_queue_cleanup_live_matches.csv")
DEFAULT_UPDATES_CSV = Path("data/qi_queue_cleanup_live_updates.csv")
QUEUE_SHEET_NAME = "Queue - Needs Graphics"
QUEUE_BATCH_SIZE = 500
SOURCE_BATCH_SIZE = 1000


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Match the current live 'New - Quote Creation Tool Database' rows against the "
            "QI image database so we can mark already-made graphics out of Weaver."
        )
    )
    parser.add_argument("--qi-csv", type=Path, default=DEFAULT_QI_DATABASE_CSV)
    parser.add_argument("--output-csv", type=Path, default=DEFAULT_OUTPUT_CSV)
    parser.add_argument("--updates-csv", type=Path, default=DEFAULT_UPDATES_CSV)
    parser.add_argument("--min-meta-score", type=int, default=2, choices=[0, 1, 2, 3])
    parser.add_argument("--apply", action="store_true", help="Write AQ=Y for verified safe matches.")
    return parser.parse_args()


def clean_flag(value: str | None) -> str:
    return (value or "").strip().upper()


def clean_text(value: str | None) -> str:
    return (value or "").replace("\r\n", "\n").replace("\r", "\n").strip()


def fetch_rows(access_token: str, sheet_name: str, start_col: str, end_col: str, batch_size: int) -> list[list[str]]:
    all_rows: list[list[str]] = []
    start_row = 2

    while True:
        end_row = start_row + batch_size - 1
        url = (
            f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}/values/"
            + urllib.parse.quote(f"'{sheet_name}'!{start_col}{start_row}:{end_col}{end_row}")
            + "?majorDimension=ROWS"
        )
        payload = api_request(access_token, "GET", url)
        rows = payload.get("values", [])
        if not rows:
            break

        all_rows.extend(rows)
        if len(rows) < batch_size:
            break
        start_row += batch_size

    return all_rows


def fetch_source_row_lookup(access_token: str) -> dict[str, dict[str, str]]:
    rows = fetch_rows(access_token, SHEET_NAME, "AP", "AQ", SOURCE_BATCH_SIZE)
    lookup: dict[str, dict[str, str]] = {}

    for index, row in enumerate(rows, start=2):
        padded = row + [""] * max(0, 2 - len(row))
        record_id = clean_text(padded[0])
        if not record_id:
            continue
        lookup[record_id] = {
            "sheetRow": str(index),
            "recordId": record_id,
            "aqOverride": clean_flag(padded[1]),
        }

    return lookup


def fetch_live_candidate_rows(access_token: str) -> list[dict[str, str]]:
    queue_rows = fetch_rows(access_token, QUEUE_SHEET_NAME, "A", "I", QUEUE_BATCH_SIZE)
    source_lookup = fetch_source_row_lookup(access_token)
    candidates: list[dict[str, str]] = []

    for row in queue_rows:
        padded = row + [""] * max(0, 9 - len(row))
        record_id = clean_text(padded[8])
        if not record_id:
            continue

        source = source_lookup.get(record_id)
        if not source or source["aqOverride"] == "Y":
            continue

        author = clean_text(padded[0])
        poem_title = clean_text(padded[1])
        book_title = clean_text(padded[2])
        excerpt = clean_text(padded[3])
        if not book_title or not excerpt:
            continue

        candidates.append(
            {
                "sheetRow": source["sheetRow"],
                "queueRecordId": record_id,
                "recordId": record_id,
                "sourceRow": "",
                "author": author,
                "bookTitle": book_title,
                "poemTitle": poem_title,
                "excerpt": excerpt,
            }
        )

    return candidates


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


def build_matches(
    candidates: list[dict[str, str]],
    qi_rows_by_hash: dict[str, list[dict[str, str]]],
) -> tuple[list[dict[str, str]], Counter]:
    audit_rows: list[dict[str, str]] = []
    counts: Counter = Counter()

    for candidate in candidates:
        excerpt_hash = fingerprint_excerpt(candidate["excerpt"])
        qi_rows = qi_rows_by_hash.get(excerpt_hash, [])
        if not qi_rows:
            counts["noExactExcerpt"] += 1
            continue

        best_score = -1
        best_flags = {"authorMatch": False, "bookMatch": False, "poemMatch": False}
        best_row: dict[str, str] | None = None
        for qi_row in qi_rows:
            score, flags = score_metadata(candidate, qi_row)
            if score > best_score:
                best_score = score
                best_flags = flags
                best_row = qi_row

        counts[f"metaScore{best_score}"] += 1
        audit_rows.append(
            {
                "sheetRow": candidate["sheetRow"],
                "recordId": candidate["recordId"],
                "sourceRow": candidate["sourceRow"],
                "author": candidate["author"],
                "bookTitle": candidate["bookTitle"],
                "poemTitle": candidate["poemTitle"],
                "excerpt": candidate["excerpt"],
                "excerptHash": excerpt_hash,
                "matchCount": str(len(qi_rows)),
                "bestMetaScore": str(best_score),
                "authorMatch": "Y" if best_flags["authorMatch"] else "",
                "bookMatch": "Y" if best_flags["bookMatch"] else "",
                "poemMatch": "Y" if best_flags["poemMatch"] else "",
                "qiAuthor": (best_row or {}).get("author_name", ""),
                "qiBookTitle": (best_row or {}).get("book_title", ""),
                "qiPoemTitle": (best_row or {}).get("poem_title", ""),
                "qiExcerpt": (best_row or {}).get("excerpt", ""),
                "qiFileName": (best_row or {}).get("file_name", ""),
                "qiLinkUrl": (best_row or {}).get("link_url", ""),
                "qiLowResFileName": (best_row or {}).get("low_res_file_name", ""),
                "qiLowResLinkUrl": (best_row or {}).get("low_res_link_url", ""),
                "qiGraphicMadeStatus": (best_row or {}).get("graphic_made_status", ""),
                "qiApprovedForQuoteImage": (best_row or {}).get("approved_for_quote_image", ""),
            }
        )

    return audit_rows, counts


def write_csv(path: Path, rows: list[dict[str, str]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def apply_updates(access_token: str, rows: list[dict[str, str]]) -> dict:
    data = [
        {
            "range": f"'{SHEET_NAME}'!AQ{row['sheetRow']}",
            "majorDimension": "ROWS",
            "values": [["Y"]],
        }
        for row in rows
    ]
    payload = {
        "valueInputOption": "USER_ENTERED",
        "data": data,
    }
    return api_request(
        access_token,
        "POST",
        f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}/values:batchUpdate",
        payload=payload,
    )


def main() -> int:
    args = parse_args()
    access_token = refresh_access_token()
    candidates = fetch_live_candidate_rows(access_token)
    qi_rows_by_hash = load_qi_rows(args.qi_csv)
    audit_rows, counts = build_matches(candidates, qi_rows_by_hash)

    audit_fieldnames = [
        "sheetRow",
        "recordId",
        "sourceRow",
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
        "qiExcerpt",
        "qiFileName",
        "qiLinkUrl",
        "qiLowResFileName",
        "qiLowResLinkUrl",
        "qiGraphicMadeStatus",
        "qiApprovedForQuoteImage",
    ]
    write_csv(args.output_csv, audit_rows, audit_fieldnames)

    update_rows = [
        {
            "sheetRow": row["sheetRow"],
            "recordId": row["recordId"],
            "sourceRow": row["sourceRow"],
            "author": row["author"],
            "bookTitle": row["bookTitle"],
            "poemTitle": row["poemTitle"],
            "bestMetaScore": row["bestMetaScore"],
            "setDriveGraphicMatchOverride": "Y",
            "qiFileName": row["qiFileName"],
            "qiLinkUrl": row["qiLinkUrl"],
        }
        for row in audit_rows
        if int(row["bestMetaScore"]) >= args.min_meta_score
    ]
    update_fieldnames = [
        "sheetRow",
        "recordId",
        "sourceRow",
        "author",
        "bookTitle",
        "poemTitle",
        "bestMetaScore",
        "setDriveGraphicMatchOverride",
        "qiFileName",
        "qiLinkUrl",
    ]
    write_csv(args.updates_csv, update_rows, update_fieldnames)

    summary = {
        "candidateRows": len(candidates),
        "safeUpdates": len(update_rows),
        "counts": dict(counts),
        "auditCsv": str(args.output_csv),
        "updatesCsv": str(args.updates_csv),
        "applied": False,
    }

    if args.apply and update_rows:
        response = apply_updates(access_token, update_rows)
        summary["applied"] = True
        summary["totalUpdatedCells"] = response.get("totalUpdatedCells")
        summary["totalUpdatedRows"] = response.get("totalUpdatedRows")
        summary["totalUpdatedSheets"] = response.get("totalUpdatedSheets")

    print(json.dumps(summary, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
