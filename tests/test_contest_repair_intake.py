import unittest
from unittest.mock import patch

from weaver_runtime_db import queue_card_record
from weaver_runtime_sync import original_repair_editable_project, upsert_repair_requests


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

    def test_compact_repair_card_keeps_short_form_identity(self):
        card = queue_card_record({
            "graphicsRequestId": "weaver:repair:existing",
            "contentType": "QI",
            "sourcePayload": {
                "sourceRecordId": self.request["originalContent"]["sourceRecordId"],
                "sourceProgramId": self.request["originalContent"]["sourceProgramId"],
                "sourceEvent": self.request["originalContent"]["sourceEvent"],
                "sourceEventLabel": self.request["originalContent"]["sourceEvent"],
            },
        })
        self.assertEqual(card["programId"], "fall-2026-contest")
        self.assertEqual(card["sourceEvent"], "Button Poetry's Fall 2026 Short Form Contest")
        self.assertEqual(card["sourceRecordId"], self.request["originalContent"]["sourceRecordId"])

    def test_claimed_repair_inherits_only_verified_original_project(self):
        original_id = "weaver:qi:original"
        completion_id = "pig-original-completion"
        project = {
            "graphicsRequestId": original_id,
            "sourceCompletionId": completion_id,
            "contentType": "QI",
            "author": "Leslie Kaup",
            "quoteText": "first line\nsecond line",
            "editableProjectValidationStatus": "verified",
            "pigProjectId": "project-original",
            "editableProjectFileId": "drive-original-json",
            "editableProjectUrl": "https://drive.google.com/file/d/drive-original-json/view",
            "editableProjectKind": "pig.editableProject",
            "editableProjectSchemaVersion": 1,
        }
        self.request["originalContent"]["misc"] = (
            f"weaverSourceRequestId={original_id} | "
            f"weaverSourceCompletionId={completion_id} | sourceExcerpt=first line\nsecond line"
        )
        existing_handoff = {
            "graphicsRequestId": "weaver:repair:existing",
            "sourcePayload": {"repairRequestId": self.request["id"]},
            "handoffStatus": "claimed",
            "pigStatus": "claimed",
            "isActionable": True,
        }
        with patch("weaver_runtime_sync.get_graphics_handoff", side_effect=lambda _db, key: {
            original_id: project,
            existing_handoff["graphicsRequestId"]: existing_handoff,
        }.get(key)), patch("weaver_runtime_sync.upsert_graphics_handoff_request", return_value=existing_handoff):
            first = upsert_repair_requests(self.connection, {"requests": [self.request]})
        self.assertEqual(first["createdCount"], 1)
        self.assertEqual(first["records"][0]["pigJobId"], existing_handoff["graphicsRequestId"])

        with patch("weaver_runtime_sync.get_graphics_handoff", side_effect=lambda _db, key: {
            original_id: project,
            existing_handoff["graphicsRequestId"]: existing_handoff,
        }.get(key)), patch("weaver_runtime_sync.update_graphics_handoff", return_value=existing_handoff) as update, patch(
            "weaver_runtime_sync.upsert_graphics_handoff_request"
        ) as create:
            second = upsert_repair_requests(self.connection, {"requests": [self.request]})
        self.assertEqual(second["duplicateCount"], 1)
        create.assert_not_called()
        update.assert_called_once()
        self.assertEqual(update.call_args.args[1], existing_handoff["graphicsRequestId"])
        self.assertEqual(update.call_args.args[2]["editableProjectFileId"], "drive-original-json")
        self.assertEqual(update.call_args.args[2]["originalGraphicsRequestId"], original_id)
        self.assertNotIn("handoffStatus", update.call_args.args[2])

        project["quoteText"] = "another poem"
        with patch("weaver_runtime_sync.get_graphics_handoff", return_value=project):
            self.assertEqual(original_repair_editable_project(self.connection, self.request, "first line\nsecond line"), {})


if __name__ == "__main__":
    unittest.main()
