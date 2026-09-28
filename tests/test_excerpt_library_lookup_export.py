import sqlite3
import tempfile
import unittest
from pathlib import Path

from export_weaver_excerpt_library import reconcile_coverage, write_source_index
from lookup_excerpt_library import lookup


class ExcerptLibraryLookupExportTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        root = Path(self.temp_dir.name)
        self.operational = root / "operational.db"
        self.normalized = root / "normalized.db"

        connection = sqlite3.connect(self.operational)
        connection.executescript("""
            CREATE TABLE excerpt_sources (id INTEGER PRIMARY KEY, source_name TEXT, source_kind TEXT);
            CREATE TABLE excerpt_entries (
                id INTEGER PRIMARY KEY, source_id INTEGER, source_row_number INTEGER,
                external_id TEXT, author TEXT, book_title TEXT, poem_title TEXT,
                excerpt_text TEXT, excerpt_hash TEXT
            );
            INSERT INTO excerpt_sources VALUES (1, 'Primary Database', 'spreadsheet');
            INSERT INTO excerpt_entries VALUES
                (16001, 1, 36138, '', 'Rudy Francisco', '', 'Complainers',
                 'Tragedy and silence have the exact same address.', 'canonical'),
                (16002, 1, 36112, '', 'Rudy Francisco', 'Helium', 'Complainers (NPS 2014)',
                 'Tragedy and silence often have the exact same address.', 'variant'),
                (16003, 1, 36113, '', 'Other', '', '', 'Hi-res image', 'nonexcerpt');
        """)
        connection.close()

        connection = sqlite3.connect(self.normalized)
        connection.executescript("""
            CREATE TABLE excerpts (id INTEGER PRIMARY KEY, excerpt_hash TEXT);
            CREATE TABLE source_rows (source_entry_id INTEGER, excerpt_id INTEGER, classification TEXT);
            CREATE TABLE peeled_off_rows (source_entry_id INTEGER, peeled_off_category TEXT);
            INSERT INTO excerpts VALUES (1, 'canonical');
            INSERT INTO source_rows VALUES (16001, 1, 'likely_excerpt');
            INSERT INTO source_rows VALUES (16002, NULL, 'likely_non_excerpt');
            INSERT INTO peeled_off_rows VALUES (16003, 'COV');
        """)
        connection.close()

    def test_lookup_scans_all_rows_and_labels_variant(self):
        result = lookup(
            self.operational,
            "Tragedy and silence have the exact same address",
            author="Rudy Francisco",
            limit=1,
        )
        self.assertEqual(result["scannedRowCount"], 3)
        self.assertEqual(result["totalMatches"], 2)
        self.assertTrue(result["truncated"])
        self.assertEqual(result["matches"][0]["sourceRow"], 36138)
        self.assertEqual(result["matches"][0]["matchType"], "exact_text")

        all_matches = lookup(self.operational, "Tragedy and silence have the exact same address")
        self.assertEqual(all_matches["matches"][1]["matchType"], "possible_variant")
        self.assertEqual(all_matches["matches"][1]["bookTitle"], "Helium")

        index_path = Path(self.temp_dir.name) / "index.csv"
        self.assertEqual(write_source_index(index_path, self.operational), 3)
        from_index = lookup(
            self.operational, "Tragedy and silence have the exact same address",
            index_path=index_path,
        )
        self.assertEqual(from_index["totalMatches"], all_matches["totalMatches"])
        self.assertEqual(from_index["matches"], all_matches["matches"])
        self.assertEqual(from_index["sourceType"], "csv_index")

    def test_every_omitted_hash_has_disposition(self):
        coverage = reconcile_coverage(self.operational, self.normalized)
        self.assertEqual(coverage["operationalRows"], 3)
        self.assertEqual(coverage["canonicalRows"], 1)
        self.assertEqual(coverage["sourceRowsOutsideCanonical"], 2)
        self.assertEqual(coverage["distinctHashesOutsideCanonical"], 2)
        self.assertEqual(coverage["outsideCanonicalDispositionRows"], {
            "classified:likely_non_excerpt": 1,
            "peeled:COV": 1,
        })


if __name__ == "__main__":
    unittest.main()
