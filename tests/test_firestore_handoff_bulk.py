import unittest
from unittest.mock import patch

from weaver_runtime_db import FirestoreLedgerClient
from weaver_runtime_sync import fetch_handoff_queue


class FirestoreHandoffBulkTest(unittest.TestCase):
    def test_full_page_exposes_next_cursor_and_second_page(self):
        client = FirestoreLedgerClient("test-project", "test-db")
        records = [
            {"graphicsRequestId": f"request-{i:03d}", "isActionable": True,
             "createdAt": f"2026-10-30T00:{i // 60:02d}:{i % 60:02d}Z"}
            for i in range(501)
        ]
        with patch.object(client, "list_handoffs", return_value=records):
            first = fetch_handoff_queue(client, {"limit": 500, "cursor": 0})
            second = fetch_handoff_queue(client, {"limit": 500, "cursor": first["nextCursor"]})

        self.assertEqual(len(first["records"]), 500)
        self.assertEqual(first["nextCursor"], "500")
        self.assertEqual([row["graphicsRequestId"] for row in second["records"]], ["request-500"])
        self.assertEqual(second["nextCursor"], "")

    def test_reads_requested_handoffs_in_bulk_and_resolves_aliases(self):
        client = FirestoreLedgerClient("test-project", "test-db")
        records = [
            {"graphicsRequestId": "one", "handoffStatus": "requested"},
            {"graphicsRequestId": "canonical-two", "handoffStatus": "claimed"},
        ]
        aliases = [{"graphicsRequestId": "two", "canonicalGraphicsRequestId": "canonical-two"}]
        with patch.object(client, "list_handoffs", return_value=records) as list_handoffs, \
             patch.object(client, "list_raw_documents", return_value=aliases) as list_aliases, \
             patch.object(client, "get_handoff") as get_handoff:
            result = client.get_handoffs(["one", "two", "missing"])

        self.assertEqual([row["graphicsRequestId"] for row in result], ["one", "two"])
        self.assertEqual([row["handoffStatus"] for row in result], ["requested", "claimed"])
        list_handoffs.assert_called_once_with()
        list_aliases.assert_called_once()
        get_handoff.assert_not_called()


if __name__ == "__main__":
    unittest.main()
