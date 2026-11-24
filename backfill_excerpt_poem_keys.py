"""Apply catalog-issued poem keys to existing excerpts by exact excerpt hash."""

import argparse
import csv
import json
import re
import sqlite3
from pathlib import Path

from excerpt_library import ensure_schema, fingerprint_excerpt


SOURCE = "catalog_shared_poem_key_seed_2026-10-15"


def backfill(db_path: Path, seed_path: Path, apply: bool = False) -> dict:
    seed = {}
    excluded_hashes = set()
    counts = {"seedUnique": 0, "matched": 0, "updated": 0, "alreadySet": 0,
              "conflicts": 0, "unmatched": 0, "staleStoredHashes": 0,
              "hashesUpdated": 0}
    if seed_path.suffix.lower() == ".json":
        with seed_path.open(encoding="utf-8-sig") as handle:
            entries = json.load(handle)
        if not isinstance(entries, dict):
            raise ValueError("JSON seed must map excerpt hashes to poem keys")
        rows = ({"excerpt_hash": excerpt_hash, "poem_key": poem_key, "match": "unique"}
                for excerpt_hash, poem_key in entries.items())
    else:
        with seed_path.open(newline="", encoding="utf-8-sig") as handle:
            rows = list(csv.DictReader(handle))
    for row in rows:
        excerpt_hash = (row.get("excerpt_hash") or "").strip().lower()
        poem_key = (row.get("poem_key") or "").strip()
        if row.get("match") != "unique":
            excluded_hashes.add(excerpt_hash)
            continue
        if not re.fullmatch(r"[0-9a-f]{64}", excerpt_hash) or not re.fullmatch(r"BPP-\d{5}", poem_key):
            raise ValueError("Invalid unique seed hash or poem key")
        if excerpt_hash in seed and seed[excerpt_hash] != poem_key:
            raise ValueError("Conflicting keys for one excerpt hash")
        seed[excerpt_hash] = poem_key
        counts["seedUnique"] += 1
    for excerpt_hash in excluded_hashes:
        seed.pop(excerpt_hash, None)

    uri = f"file:{db_path.resolve()}?mode=ro" if not apply else str(db_path)
    connection = sqlite3.connect(uri, uri=not apply)
    try:
        if apply:
            ensure_schema(connection)
        columns = {row[1] for row in connection.execute("PRAGMA table_info(excerpt_entries)")}
        has_key = "poem_key" in columns
        query = "SELECT id, excerpt_hash, excerpt_text" + (
            ", poem_key" if has_key else ", NULL AS poem_key"
        ) + " FROM excerpt_entries"
        updates = []
        hash_updates = []
        matched_hashes = set()
        for row_id, stored_hash, excerpt_text, existing_key in connection.execute(query):
            actual_hash = fingerprint_excerpt(excerpt_text)
            if stored_hash != actual_hash:
                counts["staleStoredHashes"] += 1
                hash_updates.append((actual_hash, row_id))
            poem_key = seed.get(stored_hash) if stored_hash == actual_hash else None
            poem_key = poem_key or seed.get(actual_hash)
            if not poem_key:
                continue
            matched_hashes.add(actual_hash if actual_hash in seed else stored_hash)
            counts["matched"] += 1
            if existing_key == poem_key:
                counts["alreadySet"] += 1
            elif existing_key:
                counts["conflicts"] += 1
            else:
                updates.append((poem_key, SOURCE, 1.0, row_id))
        counts["unmatched"] = len(seed.keys() - matched_hashes)
        counts["updated"] = len(updates)
        counts["hashesUpdated"] = len(hash_updates)
        if apply:
            with connection:
                connection.executemany(
                    "UPDATE excerpt_entries SET excerpt_hash = ? WHERE id = ?",
                    hash_updates,
                )
                connection.executemany(
                    "UPDATE excerpt_entries SET poem_key = ?, poem_key_source = ?, "
                    "poem_key_confidence = ? WHERE id = ? AND NULLIF(TRIM(poem_key), '') IS NULL",
                    updates,
                )
    finally:
        connection.close()
    return counts


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("db_path", type=Path)
    parser.add_argument("seed_path", type=Path)
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()
    print(json.dumps(backfill(args.db_path, args.seed_path, args.apply), sort_keys=True))
