import unittest

from backfill_firestore_excerpt_poem_keys import plan_backfill
from excerpt_library import fingerprint_excerpt


class FirestoreExcerptPoemKeyBackfillTest(unittest.TestCase):
    def test_only_verified_unchanged_records_are_planned(self):
        text = "A verified excerpt"
        excerpt_hash = fingerprint_excerpt(text)
        records = [
            {"sourceRecordId": "safe", "excerptText": text},
            {"sourceRecordId": "changed", "excerptText": text},
            {"sourceRecordId": "existing", "excerptText": text, "poemKey": "BPP-99999"},
            {"sourceRecordId": "blank", "excerptText": ""},
        ]
        local_rows = {"changed": [{"hash": fingerprint_excerpt("Different text"), "poemKey": "BPP-00001"}]}
        planned, skipped = plan_backfill(records, local_rows, {excerpt_hash: "BPP-00001"})
        self.assertEqual([item["sourceRecordId"] for item in planned], ["safe"])
        self.assertEqual(
            {item["sourceRecordId"]: item["reason"] for item in skipped},
            {"changed": "source_id_text_conflict", "existing": "firestore_key_conflict",
             "blank": "missing_excerpt_text"},
        )


if __name__ == "__main__":
    unittest.main()
