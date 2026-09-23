import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from weaver_runtime_sync import upsert_curation_review  # noqa: E402


class FakeFirestoreLedger:
    def __init__(self):
        self.documents = {}

    def get_raw_document(self, collection, document_id):
        return self.documents.get((collection, document_id))

    def write_raw_document(self, collection, document_id, record):
        self.documents[(collection, document_id)] = dict(record)
        return dict(record)


class LegacyCurationScoreTests(unittest.TestCase):
    def setUp(self):
        self.connection = FakeFirestoreLedger()
        self.identity = {
            "prioritySetId": "bpl-charm-city-2026",
            "sourceFileId": "drive-video-1",
            "reviewerEmail": "reviewer@buttonpoetry.com",
        }

    def test_numeric_score_is_converted_and_preserved(self):
        record = upsert_curation_review(self.connection, {"review": {
            **self.identity,
            "legacyImport": True,
            "legacyScore": 7.5,
            "legacyNotes": "Historical review",
            "legacySourceSpreadsheetId": "legacy-sheet",
        }})["record"]
        self.assertEqual(record["rating"], "like")
        self.assertEqual(record["legacyRating"], "like")
        self.assertEqual(record["legacyScore"], 7.5)
        self.assertEqual(record["legacySourceSpreadsheetId"], "legacy-sheet")

    def test_import_does_not_replace_a_current_vote(self):
        current = upsert_curation_review(self.connection, {"review": {
            **self.identity,
            "rating": "moved_me",
            "notes": "Current review",
        }})["record"]
        imported = upsert_curation_review(self.connection, {"review": {
            **self.identity,
            "legacyImport": True,
            "legacyScore": 7.5,
            "legacyNotes": "Historical review",
        }})["record"]
        self.assertEqual(imported["reviewId"], current["reviewId"])
        self.assertEqual(imported["rating"], "moved_me")
        self.assertEqual(imported["notes"], "Current review")
        self.assertEqual(imported["legacyScore"], 7.5)

    def test_score_outside_ten_point_scale_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "between 0 and 10"):
            upsert_curation_review(self.connection, {"review": {
                **self.identity,
                "legacyImport": True,
                "legacyScore": 11,
            }})


if __name__ == "__main__":
    unittest.main()
