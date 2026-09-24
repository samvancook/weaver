import sqlite3
import unittest
from pathlib import Path

from catalog_validate import DB_PATH


CATALOG_DB = Path(__file__).resolve().parents[1] / "data" / "formal_catalog.db"


class CatalogIdentityTest(unittest.TestCase):
    def test_local_validation_uses_repository_catalog(self):
        self.assertEqual(DB_PATH, CATALOG_DB)

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

    def test_un_catalog_has_distinct_poems_without_prologue(self):
        with sqlite3.connect(f"file:{CATALOG_DB}?mode=ro", uri=True) as connection:
            book = connection.execute(
                "SELECT id, poem_count FROM catalog_books WHERE canonical_book_id = 5"
            ).fetchone()
            poems = connection.execute(
                "SELECT title, text FROM catalog_poems WHERE catalog_book_id = ?", (book[0],)
            ).fetchall()
            status = connection.execute(
                "SELECT effective_status, primary_source_format FROM book_status WHERE canonical_book_id = 5"
            ).fetchone()

        self.assertEqual(book[1], 46)
        self.assertEqual(len(poems), 46)
        titles = {title for title, _ in poems}
        self.assertEqual(len(titles), 46)
        self.assertNotIn("prologue", titles)
        self.assertIn("cycle", titles)
        self.assertIn("namesake", titles)
        self.assertNotIn("cycle\n", dict(poems)["universal truths"].lower())
        self.assertNotIn("namesake\n", dict(poems)["self love"].lower())
        self.assertEqual(status, ("set_aside", None))


if __name__ == "__main__":
    unittest.main()
