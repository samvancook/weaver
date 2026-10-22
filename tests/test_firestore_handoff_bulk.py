import unittest
from unittest.mock import patch

from weaver_runtime_db import FirestoreLedgerClient


class FirestoreHandoffBulkTest(unittest.TestCase):
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
