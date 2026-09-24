import sqlite3
import unittest
from pathlib import Path


CATALOG_DB = Path(__file__).resolve().parents[1] / "data" / "formal_catalog.db"


class CatalogIdentityTest(unittest.TestCase):
    def test_ayanna_florence_book_uses_unreliable_narrator(self):
        with sqlite3.connect(f"file:{CATALOG_DB}?mode=ro", uri=True) as connection:
            canonical = connection.execute(
                "SELECT title, book_shortener FROM canonical_books WHERE id = 5"
            ).fetchone()
            status = connection.execute(
                "SELECT title, book_shortener, effective_status FROM book_status WHERE canonical_book_id = 5"
            ).fetchone()
            override = connection.execute(
                "SELECT title, status FROM manual_overrides WHERE canonical_book_id = 5"
            ).fetchone()

        self.assertEqual(canonical, ("unreliable narrator", "UN"))
        self.assertEqual(status, ("unreliable narrator", "UN", "set_aside"))
        self.assertEqual(override, ("unreliable narrator", "set_aside"))


if __name__ == "__main__":
    unittest.main()
