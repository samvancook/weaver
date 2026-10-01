import unittest

from weaver_runtime_sync import claim_video_curation_handoff


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

    def test_unready_gate_cannot_be_claimed(self):
        connection = FakeGateConnection({"decision": "send_to_editing"})
        result = claim_video_curation_handoff(
            connection, {"prioritySetId": "bpl-charm-city-2026", "sourceFileId": "video-1"}
        )
        self.assertFalse(result["claimed"])
        self.assertEqual(result["reason"], "gate_not_ready")


if __name__ == "__main__":
    unittest.main()
