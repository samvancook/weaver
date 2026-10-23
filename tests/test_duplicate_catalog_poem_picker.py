import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import catalog_poem_text
import catalog_validate
import intake_catalog_lookup


class DuplicateCatalogPoemPickerTest(unittest.TestCase):
    def test_duplicate_titles_keep_distinct_keys_and_text(self):
        with tempfile.TemporaryDirectory() as directory:
            database = Path(directory) / "catalog.db"
            with sqlite3.connect(database) as connection:
                connection.execute(
                    "CREATE TABLE catalog_books (id INTEGER PRIMARY KEY, canonical_book_id INTEGER, title TEXT, author TEXT)"
                )
                connection.execute(
                    "CREATE TABLE catalog_poems (id INTEGER PRIMARY KEY, catalog_book_id INTEGER, poem_key TEXT, title TEXT, text TEXT, word_count INTEGER)"
                )
                connection.execute("INSERT INTO catalog_books VALUES (1, 2, 'Example Book', 'Example Author')")
                connection.executemany(
                    "INSERT INTO catalog_poems VALUES (?, 1, ?, 'Shared Title', ?, 2)",
                    [(10, "BPP-00010", "First distinct poem with enough additional words to be treated as a complete source poem in validation."),
                     (11, "BPP-00011", "Second distinct poem with enough additional words to be treated as a complete source poem in validation.")],
                )
            book = {
                "canonical_book_id": 2,
                "title": "Example Book",
                "author": "Example Author",
                "primary_source_format": "docx",
            }
            with patch.object(intake_catalog_lookup, "DB_PATH", database), \
                 patch.object(intake_catalog_lookup, "fetch_book_status", return_value=book), \
                 patch.object(catalog_poem_text, "DB_PATH", database), \
                 patch.object(catalog_poem_text, "fetch_book_status", return_value=book):
                options = intake_catalog_lookup.get_poems("Example Book")["poemOptions"]
                self.assertEqual([option["label"] for option in options], ["Shared Title (I)", "Shared Title (II)"])
                self.assertEqual([option["poemKey"] for option in options], ["BPP-00010", "BPP-00011"])
                self.assertTrue(intake_catalog_lookup.verify_poem_selection("Example Book", "Shared Title", "BPP-00011")["ok"])
                self.assertFalse(intake_catalog_lookup.verify_poem_selection("Example Book", "Other Title", "BPP-00011")["ok"])
                first = catalog_poem_text.lookup_poem_text("Example Book", "Shared Title", poem_key="BPP-00010")
                second = catalog_poem_text.lookup_poem_text("Example Book", "Shared Title", poem_key="BPP-00011")
                self.assertTrue(first["text"].startswith("First distinct poem"))
                self.assertTrue(second["text"].startswith("Second distinct poem"))
                self.assertFalse(catalog_poem_text.lookup_poem_text("Example Book", "Shared Title")["ok"])
                with sqlite3.connect(database) as connection, \
                     patch.object(catalog_validate, "fetch_book_status", return_value={**book, "effective_status": "catalog_ok"}), \
                     patch.object(catalog_validate, "fetch_readable_book", return_value=None), \
                     patch.object(catalog_validate, "load_author_alias_map", return_value={}):
                    connection.row_factory = sqlite3.Row
                    record = {
                        "bookTitle": "Example Book", "author": "Example Author",
                        "title": "Shared Title", "excerptText": "First distinct poem",
                    }
                    first_validation = catalog_validate.validate_record(
                        connection.cursor(), {**record, "poemKey": "BPP-00010"}
                    )
                    second_validation = catalog_validate.validate_record(
                        connection.cursor(), {**record, "poemKey": "BPP-00011"}
                    )
                    self.assertEqual(first_validation["status"], "catalog_match")
                    self.assertFalse(second_validation["excerptMatchesInBook"])


if __name__ == "__main__":
    unittest.main()
