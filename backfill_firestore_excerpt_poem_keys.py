"""Backfill verified catalog poem keys into Weaver Firestore excerptRecords."""

import argparse
import json
import sqlite3
import urllib.parse
from collections import Counter
from pathlib import Path

from excerpt_library import fingerprint_excerpt
from weaver_runtime_db import (
    FIRESTORE_EXCERPT_RECORDS_COLLECTION,
    connect_firestore_ledger,
    firestore_encode_value,
)


PROJECT_ID = "button-weaver-internal"
DATABASE_ID = "weaverledger"
KEY_SOURCE = "catalog_shared_poem_key_seed_2026-10-15"
SEED_PATH = Path(__file__).parent / "data" / "catalog_poem_key_seed_2026-10-15.json"


def local_rows_by_source_id(db_path: Path) -> dict[str, list[dict]]:
    connection = sqlite3.connect(f"file:{db_path.resolve()}?mode=ro", uri=True)
    connection.row_factory = sqlite3.Row
    try:
        rows = connection.execute(
            "SELECT external_id, excerpt_text, poem_key FROM excerpt_entries "
            "WHERE NULLIF(TRIM(external_id), '') IS NOT NULL"
        )
        by_id: dict[str, list[dict]] = {}
        for row in rows:
            by_id.setdefault(row["external_id"].strip(), []).append({
                "hash": fingerprint_excerpt(row["excerpt_text"]),
                "poemKey": row["poem_key"] or "",
            })
        return by_id
    finally:
        connection.close()


def plan_backfill(records: list[dict], local_rows: dict[str, list[dict]], seed: dict[str, str]) -> tuple[list[dict], list[dict]]:
    planned = []
    skipped = []
    seen_ids = set()
    for record in records:
        source_id = str(record.get("sourceRecordId") or "").strip()
        excerpt_text = str(record.get("excerptText") or "")
        reason = ""
        key = ""
        excerpt_hash = ""
        if not source_id or source_id in seen_ids:
            reason = "missing_or_duplicate_source_id"
        elif not excerpt_text.strip():
            reason = "missing_excerpt_text"
        else:
            excerpt_hash = fingerprint_excerpt(excerpt_text)
            seed_key = seed.get(excerpt_hash, "")
            matching_local = local_rows.get(source_id, [])
            if matching_local and any(row["hash"] != excerpt_hash for row in matching_local):
                reason = "source_id_text_conflict"
            elif len({row["poemKey"] for row in matching_local if row["poemKey"]}) > 1:
                reason = "local_key_conflict"
            else:
                local_key = next((row["poemKey"] for row in matching_local if row["poemKey"]), "")
                if seed_key and local_key and seed_key != local_key:
                    reason = "seed_local_key_conflict"
                else:
                    key = local_key or seed_key
                    if not key:
                        reason = "no_verified_key"
        if source_id:
            seen_ids.add(source_id)
        existing_key = str(record.get("poemKey") or "").strip()
        if not reason and existing_key and existing_key != key:
            reason = "firestore_key_conflict"
        elif not reason and existing_key == key:
            reason = "already_keyed"
        if reason:
            skipped.append({"sourceRecordId": source_id, "reason": reason})
        else:
            planned.append({"sourceRecordId": source_id, "poemKey": key, "excerptHash": excerpt_hash})
    return planned, skipped


def run(db_path: Path, apply: bool = False, limit: int | None = None) -> dict:
    seed = json.loads(SEED_PATH.read_text(encoding="utf-8"))
    local_rows = local_rows_by_source_id(db_path)
    client = connect_firestore_ledger(PROJECT_ID, DATABASE_ID)
    records = client.list_raw_documents(FIRESTORE_EXCERPT_RECORDS_COLLECTION)
    planned, skipped = plan_backfill(records, local_rows, seed)
    summary = {
        "firestoreRecords": len(records),
        "planned": len(planned),
        "plannedIds": [item["sourceRecordId"] for item in planned[:20]],
        "skippedByReason": dict(sorted(Counter(row["reason"] for row in skipped).items())),
        "written": 0,
        "verified": 0,
        "writeConflicts": 0,
        "writeErrors": 0,
        "writtenIds": [],
    }
    if not apply:
        return summary
    for item in planned[:limit]:
        source_id = item["sourceRecordId"]
        try:
            current, update_time = client.get_raw_document_versioned(
                FIRESTORE_EXCERPT_RECORDS_COLLECTION, source_id
            )
            if not current or not update_time or str(current.get("sourceRecordId") or "").strip() != source_id:
                summary["writeConflicts"] += 1
                continue
            if current.get("poemKey") or fingerprint_excerpt(current.get("excerptText") or "") != item["excerptHash"]:
                summary["writeConflicts"] += 1
                continue
            query = urllib.parse.urlencode([
                ("updateMask.fieldPaths", "poemKey"),
                ("updateMask.fieldPaths", "poemKeySource"),
                ("currentDocument.updateTime", update_time),
            ])
            url = f"{client.document_url_for(FIRESTORE_EXCERPT_RECORDS_COLLECTION, source_id)}?{query}"
            client.request("PATCH", url, {"fields": {
                "poemKey": firestore_encode_value(item["poemKey"]),
                "poemKeySource": firestore_encode_value(KEY_SOURCE),
            }})
            summary["written"] += 1
            summary["writtenIds"].append(source_id)
            saved = client.get_raw_document(FIRESTORE_EXCERPT_RECORDS_COLLECTION, source_id)
            if saved and saved.get("poemKey") == item["poemKey"] and saved.get("poemKeySource") == KEY_SOURCE:
                summary["verified"] += 1
            else:
                summary["writeErrors"] += 1
        except RuntimeError as error:
            if "HTTP 412" in str(error):
                summary["writeConflicts"] += 1
            else:
                summary["writeErrors"] += 1
    return summary


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("db_path", type=Path)
    parser.add_argument("--apply", action="store_true")
    parser.add_argument("--limit", type=int)
    args = parser.parse_args()
    if args.limit is not None and args.limit < 1:
        parser.error("--limit must be positive")
    print(json.dumps(run(args.db_path, args.apply, args.limit), sort_keys=True))
