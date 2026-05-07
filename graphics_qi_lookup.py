#!/usr/bin/env python3
from __future__ import annotations

import csv
import json
from pathlib import Path

from excerpt_library import fingerprint_excerpt, normalize_lookup_text

QI_IMAGE_DATABASE_PATH = Path(__file__).resolve().parent / "data" / "qi_image_database.csv"
QUEUE_CLEANUP_UPDATES_PATH = Path(__file__).resolve().parent / "data" / "qi_queue_cleanup_updates.csv"


def load_cleanup_override_links() -> dict[str, dict[str, str]]:
    if not QUEUE_CLEANUP_UPDATES_PATH.exists():
      return {}

    overrides: dict[str, dict[str, str]] = {}
    with QUEUE_CLEANUP_UPDATES_PATH.open() as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            record_id = (row.get("recordId") or "").strip()
            if not record_id:
                continue
            overrides[record_id] = {
                "matchType": "cleanup_override",
                "linkUrl": (row.get("qiLinkUrl") or "").strip(),
                "fileName": (row.get("qiFileName") or "").strip(),
            }
    return overrides


def score_metadata(record: dict[str, str], asset: dict[str, str]) -> int:
    score = 0
    if normalize_lookup_text(record.get("author")) == normalize_lookup_text(asset.get("author_name")):
        score += 1
    if normalize_lookup_text(record.get("bookTitle")) == normalize_lookup_text(asset.get("book_title")):
        score += 1
    if normalize_lookup_text(record.get("poemTitle")) == normalize_lookup_text(asset.get("poem_title")):
        score += 1
    return score


def load_qi_assets() -> dict[str, list[dict[str, str]]]:
    if not QI_IMAGE_DATABASE_PATH.exists():
        return {}

    by_excerpt_hash: dict[str, list[dict[str, str]]] = {}
    with QI_IMAGE_DATABASE_PATH.open() as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            excerpt = (row.get("excerpt") or "").strip()
            link = (row.get("link_url") or "").strip()
            if not excerpt or not link:
                continue
            excerpt_hash = fingerprint_excerpt(excerpt)
            by_excerpt_hash.setdefault(excerpt_hash, []).append(row)
    return by_excerpt_hash


def lookup_graphics_assets(records: list[dict[str, str]]) -> dict[str, dict[str, str] | None]:
    cleanup_overrides = load_cleanup_override_links()
    qi_assets_by_hash = load_qi_assets()
    results: dict[str, dict[str, str] | None] = {}

    for record in records:
        record_id = (record.get("recordId") or "").strip()
        lookup_key = record_id or str(record.get("sheetRow") or "")
        if not lookup_key:
            continue

        if record_id and record_id in cleanup_overrides:
            results[lookup_key] = cleanup_overrides[record_id]
            continue

        excerpt = (record.get("quoteText") or "").strip()
        if not excerpt:
            results[lookup_key] = None
            continue

        excerpt_hash = fingerprint_excerpt(excerpt)
        candidates = qi_assets_by_hash.get(excerpt_hash, [])
        if not candidates:
            results[lookup_key] = None
            continue

        best = max(candidates, key=lambda asset: (
            score_metadata(record, asset),
            1 if (asset.get("link_url") or "").strip() else 0,
            1 if (asset.get("low_res_link_url") or "").strip() else 0,
        ))
        best_score = score_metadata(record, best)

        # Be conservative for QC links. If we cannot match at least two of
        # author/book/poem on top of the excerpt hash, don't surface a Drive
        # link in Weaver yet.
        if best_score < 2:
            results[lookup_key] = None
            continue

        results[lookup_key] = {
            "matchType": "qi_database",
            "metadataScore": str(best_score),
            "linkUrl": (best.get("link_url") or "").strip(),
            "lowResLinkUrl": (best.get("low_res_link_url") or "").strip(),
            "folderLink": (best.get("folder_link") or "").strip(),
            "fileName": (best.get("file_name") or "").strip(),
            "poemTitle": (best.get("poem_title") or "").strip(),
            "bookTitle": (best.get("book_title") or "").strip(),
            "author": (best.get("author_name") or "").strip(),
        }

    return results


def main() -> int:
    payload = json.loads(input() or "{}")
    records = payload.get("records") or []
    if not isinstance(records, list):
        records = []

    print(json.dumps({
        "ok": True,
        "matches": lookup_graphics_assets(records),
    }))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
