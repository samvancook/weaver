import sqlite3
import unittest
from unittest.mock import patch

import catalog_poem_text
import catalog_validate
import intake_catalog_lookup


class LiveCatalogBookIntakeTest(unittest.TestCase):
    def test_newly_readable_book_overrides_stale_local_status(self):
        title = "Roads"
        poem_title = "A newly readable poem"
        text = "A distinct beginning followed by enough words to make this live Catalog poem usable for excerpt validation."
        live = {
            "status": {
                "canonical_book_id": 99,
                "title": title,
                "author": "Matt Mason",
                "book_shortener": "ROADS",
                "effective_status": "catalog_ok",
                "primary_source_format": "epub",
            },
            "poems": [{"id": 1, "poem_key": "BPP-00001", "title": poem_title,
                       "text": text, "word_count": 16}],
        }
        with patch.object(intake_catalog_lookup, "fetch_readable_book", return_value=live):
            options = intake_catalog_lookup.get_poems(title)["poemOptions"]
            self.assertEqual([(option["label"], option["poemKey"]) for option in options],
                             [(poem_title, "BPP-00001")])
            self.assertTrue(intake_catalog_lookup.verify_poem_selection(
                title, poem_title, "BPP-00001")["ok"])
        with patch.object(catalog_poem_text, "fetch_readable_book", return_value=live):
            preview = catalog_poem_text.lookup_poem_text(title, poem_title, poem_key="BPP-00001")
            self.assertEqual(preview["text"], text)
        with sqlite3.connect(":memory:") as connection, \
             patch.object(catalog_validate, "fetch_readable_book", return_value=live), \
             patch.object(catalog_validate, "fetch_book_status", return_value={
                 **live["status"], "effective_status": "set_aside"
             }), \
             patch.object(catalog_validate, "load_author_alias_map", return_value={}):
            result = catalog_validate.validate_record(connection.cursor(), {
                "recordId": "gathering-preview", "sourceRow": 0, "bookTitle": title,
                "author": "Matt Mason", "title": poem_title, "poemKey": "BPP-00001",
                "excerptText": "A distinct beginning followed by enough words",
            })
            self.assertEqual(result["status"], "catalog_match")
            saved = catalog_validate.validate_record(connection.cursor(), {
                "recordId": "weaver:row-2", "sourceRow": 2, "bookTitle": title,
                "author": "Matt Mason", "title": poem_title, "poemKey": "BPP-00001",
                "excerptText": "A distinct beginning followed by enough words",
            })
            self.assertEqual(saved["status"], "catalog_match")


if __name__ == "__main__":
    unittest.main()
