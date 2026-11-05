import unittest
from unittest.mock import patch

import catalog_live
import intake_catalog_lookup


def book(book_id, **changes):
    return {"book_id": book_id, "title": f"Book {book_id}", "author": "Author",
            "shortener": f"B{book_id}", "effective_status": "catalog_ok",
            "readable": True, "source_format": "docx", "poem_key_count": 2, **changes}


def page(books, next_after=None, revision="catalog-revision"):
    return {"books": books, "next_after": next_after, "catalog_revision": revision}


class CatalogBooksListingTest(unittest.TestCase):
    def test_all_pages_use_stable_cursor_and_preserve_intake_contract(self):
        with patch.object(catalog_live, "_get_json", side_effect=[
            page([book(2)], 2), page([book(5, effective_status="provisional_source")])
        ]) as get:
            result = intake_catalog_lookup.get_intake_books()
        self.assertEqual([c.args[0] for c in get.call_args_list], [
            "/books?readable=true&limit=500&after=0",
            "/books?readable=true&limit=500&after=2"])
        self.assertEqual(result["catalogLookupStatus"], "live")
        self.assertEqual(result["catalogRevision"], "catalog-revision")
        self.assertEqual(result["books"], [
            {"title": "Book 2", "author": "Author", "bookShortener": "B2",
             "primarySourceFormat": "docx", "effectiveStatus": "catalog_ok"},
            {"title": "Book 5", "author": "Author", "bookShortener": "B5",
             "primarySourceFormat": "docx", "effectiveStatus": "provisional_source"}])

    def test_readiness_rules_and_key_count_are_not_text_counts(self):
        with patch.object(catalog_live, "_get_json", return_value=page([
            book(1), book(2, readable=False), book(3, effective_status="set_aside"),
            book(4, source_format="txt"), book(5, poem_key_count=0)
        ])):
            rows = catalog_live.list_readable_books()
        self.assertEqual([r["canonical_book_id"] for r in rows], [1])
        self.assertEqual(rows[0]["poem_key_count"], 2)
        self.assertNotIn("catalog_poem_count", rows[0])

    def test_invalid_pagination_never_returns_partial_live_list(self):
        cases = [
            [page([book(1)], 1), page([book(2)], revision="changed")],
            [page([book(1)], 1), page([book(1)])],
            [page([book(1)], 9)],
            [page([], 1)],
            [{"books": [], "catalog_revision": "rev"}],
            [page([book(1)], revision="")],
            [page([book(1, poem_key_count="bad")])],
            [page([book(True)])],
        ]
        for pages in cases:
            with self.subTest(pages=pages), patch.object(catalog_live, "_get_json", side_effect=pages):
                with self.assertRaises(ValueError):
                    catalog_live.fetch_readable_books()

    def test_mid_listing_failure_uses_explicit_cached_fallback_without_live_revision(self):
        cached = {"title": "Cached Book", "author": "Author", "book_shortener": "CB",
                  "primary_source_format": "epub", "effective_status": "catalog_ok"}
        with patch.object(catalog_live, "_get_json", side_effect=[page([book(1)], 1), OSError("offline")]), \
             patch.object(intake_catalog_lookup, "load_book_status_rows", return_value=[cached]):
            result = intake_catalog_lookup.get_intake_books()
        self.assertEqual(result["catalogLookupStatus"], "cached_fallback")
        self.assertEqual(result["catalogRevision"], "")
        self.assertEqual([b["title"] for b in result["books"]], ["Cached Book"])

    def test_empty_live_listing_keeps_revision_and_does_not_fall_back(self):
        with patch.object(catalog_live, "_get_json", return_value=page([])), \
             patch.object(intake_catalog_lookup, "load_book_status_rows") as fallback:
            result = intake_catalog_lookup.get_intake_books()
        self.assertEqual(result["books"], [])
        self.assertEqual(result["catalogRevision"], "catalog-revision")
        fallback.assert_not_called()


if __name__ == "__main__":
    unittest.main()
