import csv
import sqlite3
import tempfile
import unittest
from pathlib import Path

from backfill_excerpt_poem_keys import backfill
from excerpt_library import ensure_schema, fingerprint_excerpt


class PoemKeyBackfillTest(unittest.TestCase):
    def test_unique_hash_only_and_idempotent(self):
        with tempfile.TemporaryDirectory() as directory:
            db_path = Path(directory) / "excerpts.db"
            seed_path = Path(directory) / "seed.csv"
            connection = sqlite3.connect(db_path)
            ensure_schema(connection)
            connection.execute(
                "INSERT INTO excerpt_sources (id, source_name, source_kind) VALUES (1, 'test', 'test')"
            )
            for row_id, text in ((1, "One quoted line"), (2, "Ambiguous line"),
                                 (3, "One quoted line")):
                connection.execute(
                    "INSERT INTO excerpt_entries (id, source_id, source_row_number, excerpt_text, "
                    "normalized_excerpt, excerpt_hash, word_count, character_count, metadata_json) "
                    "VALUES (?, 1, ?, ?, ?, ?, 3, ?, '{}')",
                    (row_id, row_id, text, text,
                     "0" * 64 if row_id == 3 else fingerprint_excerpt(text), len(text)),
                )
            connection.commit()
            connection.close()
            with seed_path.open("w", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=["excerpt_hash", "poem_key", "match"])
                writer.writeheader()
                writer.writerow({"excerpt_hash": fingerprint_excerpt("One quoted line"),
                                 "poem_key": "BPP-00001", "match": "unique"})
                writer.writerow({"excerpt_hash": fingerprint_excerpt("Ambiguous line"),
                                 "poem_key": "", "match": "ambiguous"})

            self.assertEqual(backfill(db_path, seed_path)["updated"], 2)
            self.assertEqual(backfill(db_path, seed_path, apply=True)["updated"], 2)
            self.assertEqual(backfill(db_path, seed_path, apply=True)["updated"], 0)
            connection = sqlite3.connect(db_path)
            self.assertEqual(connection.execute(
                "SELECT poem_key FROM excerpt_entries WHERE id = 1"
            ).fetchone()[0], "BPP-00001")
            self.assertEqual(connection.execute(
                "SELECT poem_key FROM excerpt_entries WHERE id = 3"
            ).fetchone()[0], "BPP-00001")
            self.assertIsNone(connection.execute(
                "SELECT poem_key FROM excerpt_entries WHERE id = 2"
            ).fetchone()[0])
            connection.close()


if __name__ == "__main__":
    unittest.main()
