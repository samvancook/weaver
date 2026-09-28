#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

from excerpt_library import normalize_lookup_text


DEFAULT_DB_PATH = Path("data/excerpt_library_normalized.db")
DEFAULT_OPERATIONAL_DB_PATH = Path("data/excerpt_library.db")
DEFAULT_OUTPUT_DIR = Path("data/exports/weaver_excerpt_library")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export a stable excerpt-library artifact for Weaver."
    )
    parser.add_argument(
        "--db-path",
        type=Path,
        default=DEFAULT_DB_PATH,
        help=f"Normalized excerpt DB path (default: {DEFAULT_DB_PATH})",
    )
    parser.add_argument(
        "--operational-db-path",
        type=Path,
        default=DEFAULT_OPERATIONAL_DB_PATH,
        help=f"Operational excerpt DB used for event metadata (default: {DEFAULT_OPERATIONAL_DB_PATH})",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help=f"Output directory (default: {DEFAULT_OUTPUT_DIR})",
    )
    return parser.parse_args()


def fetch_event_metadata(db_path: Path) -> dict[str, dict[str, list[str]]]:
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    try:
        rows = connection.execute(
            "SELECT excerpt_hash, metadata_json FROM excerpt_entries"
        ).fetchall()
    finally:
        connection.close()

    events_by_hash: dict[str, dict[str, set[str]]] = {}
    for row in rows:
        try:
            metadata = json.loads(row["metadata_json"] or "{}")
        except json.JSONDecodeError:
            continue
        event = str(metadata.get("sourceEvent") or "").strip()
        label = str(metadata.get("sourceEventLabel") or "").strip()
        if not event and not label:
            continue
        bucket = events_by_hash.setdefault(
            row["excerpt_hash"], {"sourceEvents": set(), "sourceEventLabels": set()}
        )
        if event:
            bucket["sourceEvents"].add(event)
        if label:
            bucket["sourceEventLabels"].add(label)

    return {
        key: {
            "sourceEvents": sorted(value["sourceEvents"]),
            "sourceEventLabels": sorted(value["sourceEventLabels"]),
        }
        for key, value in events_by_hash.items()
    }


def fetch_rows(db_path: Path, operational_db_path: Path) -> list[dict]:
    events_by_hash = fetch_event_metadata(operational_db_path)
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    try:
        rows = connection.execute(
            """
            SELECT
                id,
                excerpt_hash,
                canonical_excerpt_text,
                normalized_excerpt,
                pull_count,
                primary_author,
                primary_book_title,
                primary_poem_title,
                classification_basis,
                source_row_count,
                has_qi_asset,
                qi_asset_count,
                qi_linked_asset_count,
                qi_approved_count,
                qi_graphic_made_count,
                int_asset_count,
                other_asset_count
            FROM excerpts
            ORDER BY id
            """
        ).fetchall()
    finally:
        connection.close()

    output: list[dict] = []
    for row in rows:
        event_metadata = events_by_hash.get(
            row["excerpt_hash"], {"sourceEvents": [], "sourceEventLabels": []}
        )
        output.append({
            "excerptId": int(row["id"]),
            "excerptHash": row["excerpt_hash"],
            "excerptText": row["canonical_excerpt_text"],
            "normalizedExcerpt": row["normalized_excerpt"],
            "pullCount": int(row["pull_count"] or 0),
            "author": row["primary_author"] or "",
            "bookTitle": row["primary_book_title"] or "",
            "poemTitle": row["primary_poem_title"] or "",
            "classificationBasis": row["classification_basis"] or "",
            "sourceRowCount": int(row["source_row_count"] or 0),
            "hasQiAsset": int(row["has_qi_asset"] or 0),
            "qiAssetCount": int(row["qi_asset_count"] or 0),
            "qiLinkedAssetCount": int(row["qi_linked_asset_count"] or 0),
            "qiApprovedCount": int(row["qi_approved_count"] or 0),
            "qiGraphicMadeCount": int(row["qi_graphic_made_count"] or 0),
            "intAssetCount": int(row["int_asset_count"] or 0),
            "otherAssetCount": int(row["other_asset_count"] or 0),
            "sourceEvents": event_metadata["sourceEvents"],
            "sourceEventLabels": event_metadata["sourceEventLabels"],
        })
    return output


def write_csv(path: Path, rows: list[dict]) -> None:
    fieldnames = [
        "excerptId",
        "excerptHash",
        "excerptText",
        "normalizedExcerpt",
        "pullCount",
        "author",
        "bookTitle",
        "poemTitle",
        "classificationBasis",
        "sourceRowCount",
        "hasQiAsset",
        "qiAssetCount",
        "qiLinkedAssetCount",
        "qiApprovedCount",
        "qiGraphicMadeCount",
        "intAssetCount",
        "otherAssetCount",
        "sourceEvents",
        "sourceEventLabels",
    ]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            csv_row = dict(row)
            csv_row["sourceEvents"] = " | ".join(row["sourceEvents"])
            csv_row["sourceEventLabels"] = " | ".join(row["sourceEventLabels"])
            writer.writerow(csv_row)


def write_source_index(path: Path, db_path: Path) -> int:
    """Export every operational row, including rows rolled into canonical excerpts."""
    connection = sqlite3.connect(db_path.resolve().as_uri() + "?mode=ro", uri=True)
    connection.row_factory = sqlite3.Row
    fieldnames = [
        "entryId", "sourceName", "sourceKind", "sourceRow", "sourceRecordId",
        "author", "bookTitle", "poemTitle", "excerptText", "normalizedLookupText",
        "excerptHash",
    ]
    count = 0
    try:
        with path.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=fieldnames)
            writer.writeheader()
            for row in connection.execute(
                """
                SELECT e.id, s.source_name, s.source_kind, e.source_row_number,
                       e.external_id, e.author, e.book_title, e.poem_title,
                       e.excerpt_text, e.excerpt_hash
                FROM excerpt_entries AS e
                JOIN excerpt_sources AS s ON s.id = e.source_id
                ORDER BY e.id
                """
            ):
                writer.writerow({
                    "entryId": row["id"],
                    "sourceName": row["source_name"],
                    "sourceKind": row["source_kind"],
                    "sourceRow": row["source_row_number"],
                    "sourceRecordId": row["external_id"] or "",
                    "author": row["author"] or "",
                    "bookTitle": row["book_title"] or "",
                    "poemTitle": row["poem_title"] or "",
                    "excerptText": row["excerpt_text"],
                    "normalizedLookupText": normalize_lookup_text(row["excerpt_text"]),
                    "excerptHash": row["excerpt_hash"],
                })
                count += 1
    finally:
        connection.close()
    return count


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def reconcile_coverage(operational_db: Path, normalized_db: Path) -> dict:
    operational = sqlite3.connect(operational_db.resolve().as_uri() + "?mode=ro", uri=True)
    normalized = sqlite3.connect(normalized_db.resolve().as_uri() + "?mode=ro", uri=True)
    try:
        source_rows = operational.execute("SELECT id, excerpt_hash FROM excerpt_entries").fetchall()
        canonical_hashes = {
            row[0] for row in normalized.execute("SELECT excerpt_hash FROM excerpts")
        }
        source_hashes = {row[1] for row in source_rows}
        peeled = dict(normalized.execute(
            "SELECT source_entry_id, peeled_off_category FROM peeled_off_rows"
        ))
        classified = dict(normalized.execute(
            "SELECT source_entry_id, classification FROM source_rows WHERE excerpt_id IS NULL"
        ))
        linked_ids = {
            row[0] for row in normalized.execute("SELECT source_entry_id FROM source_rows")
        }
    finally:
        operational.close()
        normalized.close()

    source_ids = {row[0] for row in source_rows}
    if source_ids != linked_ids | set(peeled):
        raise ValueError("Normalized source rows do not account for every operational entry")
    if not canonical_hashes <= source_hashes:
        raise ValueError("Canonical excerpts contain hashes absent from the operational database")

    missing = [(entry_id, excerpt_hash) for entry_id, excerpt_hash in source_rows
               if excerpt_hash not in canonical_hashes]
    dispositions: dict[str, int] = {}
    for entry_id, _ in missing:
        if entry_id in peeled:
            label = f"peeled:{peeled[entry_id]}"
        elif entry_id in classified:
            label = f"classified:{classified[entry_id]}"
        else:
            raise ValueError(f"Operational entry {entry_id} is absent from the canonical export without a disposition")
        dispositions[label] = dispositions.get(label, 0) + 1

    return {
        "operationalRows": len(source_rows),
        "operationalDistinctHashes": len(source_hashes),
        "canonicalRows": len(canonical_hashes),
        "sourceRowsOutsideCanonical": len(missing),
        "distinctHashesOutsideCanonical": len({row[1] for row in missing}),
        "outsideCanonicalDispositionRows": dict(sorted(dispositions.items())),
        "unaccountedOperationalRows": 0,
    }


def main() -> int:
    args = parse_args()
    for db_path in (args.db_path, args.operational_db_path):
        if not db_path.is_file():
            raise FileNotFoundError(f"Database not found: {db_path}")
    source_stats_before = {
        str(path): (path.stat().st_size, path.stat().st_mtime_ns)
        for path in (args.db_path, args.operational_db_path)
    }
    coverage = reconcile_coverage(args.operational_db_path, args.db_path)
    rows = fetch_rows(args.db_path, args.operational_db_path)

    args.output_dir.mkdir(parents=True, exist_ok=True)
    json_path = args.output_dir / "weaver_excerpt_library.json"
    csv_path = args.output_dir / "weaver_excerpt_library.csv"
    manifest_path = args.output_dir / "manifest.json"
    source_index_path = args.output_dir / "operational_excerpt_index.csv"

    json_path.write_text(json.dumps({"records": rows}, indent=2, ensure_ascii=True) + "\n", encoding="utf-8")
    write_csv(csv_path, rows)
    source_row_count = write_source_index(source_index_path, args.operational_db_path)
    if source_row_count != coverage["operationalRows"] or len(rows) != coverage["canonicalRows"]:
        raise ValueError("Export counts changed during generation")
    source_stats_after = {
        str(path): (path.stat().st_size, path.stat().st_mtime_ns)
        for path in (args.db_path, args.operational_db_path)
    }
    if source_stats_before != source_stats_after:
        raise ValueError("Source databases changed during export; rerun from a stable snapshot")

    manifest = {
        "dbPath": str(args.db_path),
        "operationalDbPath": str(args.operational_db_path),
        "recordCount": len(rows),
        "operationalRowCount": source_row_count,
        "primaryArtifact": str(json_path),
        "secondaryArtifact": str(csv_path),
        "sourceIndexArtifact": str(source_index_path),
        "artifactSha256": {
            "json": sha256_file(json_path),
            "csv": sha256_file(csv_path),
            "sourceIndexCsv": sha256_file(source_index_path),
        },
        "artifactFiles": {
            "json": json_path.name,
            "csv": csv_path.name,
            "sourceIndexCsv": source_index_path.name,
        },
        "sourceDatabaseSha256": {
            "normalized": sha256_file(args.db_path),
            "operational": sha256_file(args.operational_db_path),
        },
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "coverage": coverage,
        "schemaVersion": 4,
        "scope": {
            "canonicalExport": "One row per normalized excerpt; source rows may be merged or peeled off.",
            "sourceIndex": "Every row in the local operational excerpt_entries table at export time.",
            "freshness": "Local snapshot only. This manifest does not prove live cloud sync or consumer refresh.",
        },
        "matchKey": "normalizedExcerpt",
        "preferredMatchWorkflow": [
            "Normalize Weaver candidate excerpt text the same way the library does.",
            "Match on normalizedExcerpt for exact-text duplicate detection.",
            "Use excerptHash as a stable canonical excerpt identifier inside the export.",
            "Treat hasQiAsset and qi* fields as informational only, not match criteria.",
        ],
        "fields": {
            "excerptId": "Local canonical excerpt row id in normalized DB.",
            "excerptHash": "Stable hash of the canonical excerpt text.",
            "excerptText": "Canonical excerpt text.",
            "normalizedExcerpt": "Normalized excerpt text for exact duplicate matching.",
            "pullCount": "How many source rows rolled up into this excerpt.",
            "author": "Preferred author value for the excerpt.",
            "bookTitle": "Preferred book title value for the excerpt.",
            "poemTitle": "Preferred poem/source title value for the excerpt.",
            "hasQiAsset": "1 if at least one linked QI asset exists.",
            "qiAssetCount": "Total number of linked QI asset rows.",
            "qiLinkedAssetCount": "QI assets with links present.",
            "qiApprovedCount": "QI assets marked approved.",
            "qiGraphicMadeCount": "QI assets with a graphic-made status.",
            "sourceEvents": "Optional event identifiers associated with this canonical excerpt.",
            "sourceEventLabels": "Optional human-readable event labels associated with this canonical excerpt.",
        },
    }
    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=True) + "\n", encoding="utf-8")

    print(
        json.dumps(
            {
                "recordCount": len(rows),
                "jsonPath": str(json_path),
                "csvPath": str(csv_path),
                "sourceIndexPath": str(source_index_path),
                "operationalRowCount": source_row_count,
                "manifestPath": str(manifest_path),
            },
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
