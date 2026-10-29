import unittest

from weaver_runtime_sync import claim_video_curation_handoff, finish_video_curation_handoff


class FakeGateConnection:
    def __init__(self, record):
        self.record = record
        self.version = "2026-09-29T00:00:00.000000Z"

    def get_raw_document_versioned(self, _collection, _document_id):
        return dict(self.record), self.version

    def write_raw_document_if_unchanged(self, _collection, _document_id, record, update_time):
        if update_time != self.version:
            return None
        self.record = record
        self.version = "2026-09-29T00:00:01.000000Z"
        return record


class VideoCurationClaimTest(unittest.TestCase):
    def test_only_first_claim_can_start_handoff(self):
        connection = FakeGateConnection({
            "decision": "ready_for_poetry_please",
            "poetryPleaseHandoff": {},
        })
        payload = {"prioritySetId": "bpl-charm-city-2026", "sourceFileId": "video-1"}

        first = claim_video_curation_handoff(connection, payload)
        second = claim_video_curation_handoff(connection, payload)

        self.assertTrue(first["claimed"])
        self.assertFalse(second["claimed"])
        self.assertEqual(second["reason"], "handoff_already_started")
        self.assertEqual(connection.record["poetryPleaseHandoff"]["status"], "sending")
        self.assertTrue(connection.record["poetryPleaseHandoff"]["attemptId"])

    def test_only_claimed_attempt_can_finish_without_rewriting_decision(self):
        connection = FakeGateConnection({
            "decision": "ready_for_poetry_please",
            "decidedBy": "curator@example.com",
            "decidedAt": "2026-10-01T00:00:00Z",
            "sourceReviewFileId": "reviewed-file",
            "publishableAssetUrl": "https://drive.google.com/file/d/final-file/view",
            "poetryPleaseHandoff": {},
        })
        identity = {"prioritySetId": "mpmu-2026-finals", "sourceFileId": "final-file"}
        claim = claim_video_curation_handoff(connection, identity)
        attempt_id = claim["record"]["poetryPleaseHandoff"]["attemptId"]
        result = finish_video_curation_handoff(connection, {
            **identity,
            "attemptId": attempt_id,
            "poetryPleaseHandoff": {
                "status": "sent_to_poetry_please",
                "sourceRecordId": "weaver:video:reviewed-file",
                "canonicalVideoId": "WEAVER-VV-123",
                "canonicalVideoUrl": "https://example.com/video/123",
            },
        })
        self.assertTrue(result["saved"])
        self.assertEqual(connection.record["decidedBy"], "curator@example.com")
        self.assertEqual(connection.record["decidedAt"], "2026-10-01T00:00:00Z")
        self.assertEqual(connection.record["sourceReviewFileId"], "reviewed-file")
        self.assertEqual(connection.record["publishableAssetUrl"], "https://drive.google.com/file/d/final-file/view")
        self.assertFalse(finish_video_curation_handoff(connection, {
            **identity, "attemptId": attempt_id, "poetryPleaseHandoff": {"status": "failed"}
        })["saved"])

    def test_failed_handoff_remains_held_for_receipt_reconciliation(self):
        connection = FakeGateConnection({
            "decision": "ready_for_poetry_please",
            "poetryPleaseHandoff": {"status": "failed"},
        })
        result = claim_video_curation_handoff(connection, {
            "prioritySetId": "mpmu-2026-finals", "sourceFileId": "final-file"
        })
        self.assertFalse(result["claimed"])
        self.assertEqual(result["reason"], "handoff_already_started")

    def test_blocked_receiver_result_is_saved_and_cannot_be_reclaimed(self):
        connection = FakeGateConnection({
            "decision": "ready_for_poetry_please",
            "poetryPleaseHandoff": {},
        })
        identity = {"prioritySetId": "bpl-charm-city-2026", "sourceFileId": "fv-file"}
        claim = claim_video_curation_handoff(connection, identity)
        attempt_id = claim["record"]["poetryPleaseHandoff"]["attemptId"]
        result = finish_video_curation_handoff(connection, {
            **identity,
            "attemptId": attempt_id,
            "poetryPleaseHandoff": {
                "status": "blocked",
                "sourceRecordId": "weaver:video:fv-file",
                "error": "final_asset_matches_raw_source",
            },
        })
        self.assertTrue(result["saved"])
        self.assertEqual(connection.record["handoffStatus"], "blocked")
        self.assertEqual(connection.record["poetryPleaseHandoff"]["error"], "final_asset_matches_raw_source")
        self.assertFalse(claim_video_curation_handoff(connection, identity)["claimed"])

    def test_unready_gate_cannot_be_claimed(self):
        connection = FakeGateConnection({"decision": "send_to_editing"})
        result = claim_video_curation_handoff(
            connection, {"prioritySetId": "bpl-charm-city-2026", "sourceFileId": "video-1"}
        )
        self.assertFalse(result["claimed"])
        self.assertEqual(result["reason"], "gate_not_ready")


if __name__ == "__main__":
    unittest.main()
