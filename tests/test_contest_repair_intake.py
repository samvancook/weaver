import unittest
from unittest.mock import patch

from weaver_runtime_sync import upsert_repair_requests


class FakeFirestore:
    def __init__(self):
        self.documents = {}

    def get_raw_document(self, collection, document_id):
        return self.documents.get((collection, document_id))

    def write_raw_document(self, collection, document_id, record):
        self.documents[(collection, document_id)] = record
        return record


class ContestRepairIntakeTest(unittest.TestCase):
    def setUp(self):
        self.connection = FakeFirestore()
        self.request = {
            "id": "Xn8OMGMVpU3nOfumcMPY",
            "status": "requested",
            "imageId": "PP-CONTEST-QI-aNT7jMmPfZ7eoyPxOsYr",
            "contentType": "QI",
            "author": "Leslie Kaup",
            "book": "",
            "title": "",
            "originalContent": {
                "sourceSubmissionId": "aNT7jMmPfZ7eoyPxOsYr",
                "sourceProgramId": "fall-2026-contest",
                "sourceRecordId": "poetryplease:contest:fall-2026-contest:aNT7jMmPfZ7eoyPxOsYr",
                "sourceEvent": "Button Poetry's Fall 2026 Short Form Contest",
                "imageUrl": "https://example.com/original.png",
                "misc": "weaverSourceRequestId=weaver:qi:original | sourceExcerpt=first line\nsecond line",
            },
        }

    def test_verified_contest_keeps_blank_metadata_and_poem_line_breaks(self):
        handoffs = []

        def save_handoff(_connection, handoff):
            saved = {**handoff, "isActionable": True}
            handoffs.append(saved)
            return saved

        with patch("weaver_runtime_sync.upsert_graphics_handoff_request", side_effect=save_handoff), \
             patch("weaver_runtime_sync.get_graphics_handoff", side_effect=lambda _connection, _id: handoffs[0]):
            first = upsert_repair_requests(self.connection, {"requests": [self.request]})
            second = upsert_repair_requests(self.connection, {"requests": [self.request]})

        self.assertEqual(first["blockedCount"], 0)
        self.assertEqual(second["duplicateCount"], 1)
        self.assertEqual(len(handoffs), 1)
        payload = handoffs[0]["sourcePayload"]
        self.assertEqual(payload["book"], "")
        self.assertEqual(payload["title"], "")
        self.assertEqual(payload["quoteText"], "first line\nsecond line")
        self.assertEqual(payload["sourceRecordId"], self.request["originalContent"]["sourceRecordId"])
        self.assertEqual(payload["originalContentId"], self.request["imageId"])
        self.assertEqual(first["records"][0]["pigJobId"], handoffs[0]["graphicsRequestId"])

    def test_unverified_contest_still_requires_book_and_title(self):
        self.request["originalContent"]["sourceRecordId"] = "wrong-source"
        with patch("weaver_runtime_sync.upsert_graphics_handoff_request") as save_handoff:
            result = upsert_repair_requests(self.connection, {"requests": [self.request]})
        self.assertEqual(result["blockedCount"], 1)
        self.assertIn("book,title", result["records"][0]["blockedReason"])
        save_handoff.assert_not_called()


if __name__ == "__main__":
    unittest.main()
