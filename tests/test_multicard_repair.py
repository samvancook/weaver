import unittest
from unittest.mock import patch

from weaver_runtime_db import FirestoreLedgerClient, normalize_handoff_record, queue_card_record
from weaver_runtime_sync import (
    FIRESTORE_REPAIR_REQUESTS_COLLECTION, build_repair_handoff_request,
    original_repair_editable_project, repair_card_context, repair_document_id,
    sync_completions, reconcile_repair_request, resolve_content_type,
)


class FakeFirestore:
    def __init__(self):
        self.documents = {}

    def get_raw_document(self, collection, document_id):
        return self.documents.get((collection, document_id))

    def write_raw_document(self, collection, document_id, record):
        self.documents[(collection, document_id)] = record
        return record


class MultiCardRepairTest(unittest.TestCase):
    def setUp(self):
        self.assets = [{"index": i, "role": "card", "assetUrl": f"https://example.com/{i}.png",
                        "assetPreviewUrl": f"https://example.com/{i}-preview.png",
                        "assetFileId": f"synthetic-{i}", "fileName": f"{i}.png"} for i in range(1, 4)]
        self.request = {
            "id": "synthetic-repair", "contentType": "FPI", "author": "Synthetic Author",
            "sourceText": "Synthetic poem", "affectedCardIndex": 2,
            "originalContent": {"assetSet": self.assets, "sourceRecordId": "original-source",
                                "sourceRequestId": "original-request", "sourceCompletionId": "original-completion"},
        }

    def test_repair_reopens_verified_whole_project_with_target_card(self):
        original = {
            "contentType": "FPI", "author": "Synthetic Author", "quoteText": "Synthetic poem",
            "sourceCompletionId": "original-completion", "editableProjectValidationStatus": "verified",
            "pigProjectId": "whole-project", "editableProjectFileId": "whole-project-json",
            "editableProjectUrl": "https://example.com/project.json",
        }
        with patch("weaver_runtime_sync.get_graphics_handoff", return_value=original):
            project = original_repair_editable_project(None, self.request, "Synthetic poem")
        handoff = build_repair_handoff_request(self.request, self.request["id"], "repair-job", project)
        card = queue_card_record(normalize_handoff_record(handoff))
        self.assertEqual(card["assetSet"], self.assets)
        self.assertEqual(card["previousAssetSet"], self.assets)
        self.assertEqual(card["affectedCardIndex"], 2)
        self.assertEqual(card["pigProjectId"], "whole-project")
        self.assertEqual(card["editableProjectFileId"], "whole-project-json")
        self.assertEqual(card["revisionOf"], "original-completion")
        self.assertEqual(card["contentType"], "FPI")
        self.assertEqual(card["sourceRecordId"], "original-source")

    def test_replacement_completion_returns_full_set_and_preserves_target(self):
        connection = FakeFirestore()
        connection.write_raw_document(FIRESTORE_REPAIR_REQUESTS_COLLECTION,
                                      repair_document_id(self.request["id"]), {
                                          **repair_card_context(self.request), "pigJobId": "repair-job"})
        replacement_set = [dict(asset) for asset in self.assets]
        replacement_set[1].update(assetFileId="replacement-2", assetUrl="https://example.com/replaced-2.png")
        completion = {
            "requestId": "repair-job", "completionId": "replacement-completion", "contentType": "FPI",
            "repairRequestId": self.request["id"], "assetSet": replacement_set,
            "assetUrl": self.assets[0]["assetUrl"], "assetFileId": self.assets[0]["assetFileId"],
            "pigProjectId": "whole-project", "editableProjectFileId": "whole-project-json",
        }
        with patch("weaver_runtime_sync.upsert_graphics_request"), \
             patch("weaver_runtime_sync.replace_graphics_request_items"), \
             patch("weaver_runtime_sync.insert_graphics_completion"), \
             patch("weaver_runtime_sync.upsert_graphics_handoff_request", return_value={}), \
             patch("weaver_runtime_sync.update_graphics_handoff"):
            result = sync_completions(connection, {"completions": [completion]})
        self.assertEqual(result["written"], 1)
        self.assertEqual(len(result["repairReturns"]), 1)
        returned = result["repairReturns"][0]
        self.assertEqual(returned["assetSet"], replacement_set)
        self.assertEqual(returned["assetSet"][0], self.assets[0])
        self.assertEqual(returned["assetSet"][2], self.assets[2])
        self.assertEqual(returned["affectedCardIndex"], 2)
        self.assertEqual(returned["sourceRecordId"], "original-source")
        self.assertEqual(returned["replacementAssetLink"], self.assets[0]["assetUrl"])
        self.assertEqual(returned["pigProjectId"], "whole-project")

    def test_persisted_handoff_updates_set_instead_of_leaving_original_payload(self):
        class MemoryHandoff(FirestoreLedgerClient):
            def __init__(self, record):
                self.record = normalize_handoff_record(record)

            def get_handoff(self, _request_id):
                return dict(self.record)

            def write_record(self, record):
                self.record = normalize_handoff_record(record)
                return self.record

        original = build_repair_handoff_request(self.request, self.request["id"], "repair-job")
        connection = MemoryHandoff(original)
        replacements = [dict(asset) for asset in self.assets]
        replacements[1]["assetFileId"] = "replacement-2"
        saved = connection.update_handoff("repair-job", {"assetSet": replacements})
        self.assertEqual(queue_card_record(saved)["assetSet"], replacements)
        self.assertEqual(queue_card_record(saved)["previousAssetSet"], self.assets)
        self.assertEqual(queue_card_record(saved)["affectedCardIndex"], 2)

    def test_rejected_return_reopens_same_project_and_full_returned_set(self):
        connection = FakeFirestore()
        connection.write_raw_document(FIRESTORE_REPAIR_REQUESTS_COLLECTION,
                                      repair_document_id(self.request["id"]), {
                                          **repair_card_context(self.request),
                                          "pigJobId": "repair-job", "weaverRepairStatus": "returned",
                                          "replacementAssetId": "synthetic-1",
                                          "replacementAssetLink": self.assets[0]["assetUrl"],
                                          "returnedAt": "2026-10-26T10:00:00Z",
                                      })
        handoff = {"graphicsRequestId": "repair-job", "pigProjectId": "whole-project"}
        with patch("weaver_runtime_sync.get_graphics_handoff", return_value=handoff), \
             patch("weaver_runtime_sync.update_graphics_handoff") as update:
            result = reconcile_repair_request(connection, {
                "repairRequestId": self.request["id"],
                "poetryPleaseRequest": {"status": "returned", "returnReviewStatus": "rejected",
                                        "returnReviewNote": "Card 2 still needs correction",
                                        "returnReviewedAt": "2026-10-26T11:00:00Z"},
            })
        self.assertEqual(result["transition"], "reopened")
        self.assertEqual(update.call_args.args[1], "repair-job")
        self.assertEqual(update.call_args.args[2]["previousAssetSet"], self.assets)
        self.assertEqual(update.call_args.args[2]["affectedCardIndex"], 2)
        self.assertEqual(update.call_args.args[2]["nextAction"], "rework")

    def test_invalid_target_is_rejected(self):
        for invalid in (0, 4, True, "2"):
            with self.assertRaises(ValueError):
                repair_card_context({**self.request, "affectedCardIndex": invalid})

    def test_full_poem_graphic_completion_uses_fpi_output_type(self):
        self.assertEqual(resolve_content_type({"contentType": "FP", "assetSet": self.assets}), "FPI")


if __name__ == "__main__":
    unittest.main()
