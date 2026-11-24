import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from weaver_runtime_sync import upsert_excerpt_records


class Ledger:
    def __init__(self, existing=None):
        self.existing = existing
        self.writes = []

    def get_raw_document(self, collection, document_id):
        return self.existing

    def write_raw_document(self, collection, document_id, record):
        self.writes.append((collection, document_id, record))
        self.existing = record
        return record


class BookIntakeReplayTest(unittest.TestCase):
    def test_replay_never_rewrites_later_approval_correction_or_identity(self):
        for decision in ("ACCEPT", "ACCEPT_SKIP_GRAPHIC", "NEEDS_CORRECTION", "REJECT"):
            with self.subTest(decision=decision):
                existing = {"sourceRecordId": "source-one", "reviewDecision": decision,
                            "approvedForUse": decision.startswith("ACCEPT"),
                            "correctedExcerptText": "Corrected text", "correctionNote": "Staff note",
                            "poemKey": "KEY-1", "excluded": True, "updatedAt": "later-review"}
                connection = Ledger(existing)
                result = upsert_excerpt_records(connection, {"intakeOnly": True, "records": [{
                    "sourceRecordId": "source-one", "intakeMode": "book", "excerptText": "Original intake",
                    "reviewDecision": "", "poemKey": "old-key", "excluded": False,
                }]})
                self.assertEqual(connection.writes, [])
                self.assertEqual(result["records"], [existing])
                self.assertEqual(connection.existing["updatedAt"], "later-review")

    def test_first_intake_still_creates_the_keyed_record(self):
        connection = Ledger()
        result = upsert_excerpt_records(connection, {"intakeOnly": True, "records": [{
            "sourceRecordId": "source-one", "intakeMode": "book", "poemKey": "KEY-1",
            "poemKeySource": "catalog", "excerptText": "An actual submitted passage",
        }]})
        self.assertEqual(len(connection.writes), 1)
        self.assertEqual(result["records"][0]["poemKey"], "KEY-1")

    def test_normal_review_updates_still_save_without_the_intake_flag(self):
        connection = Ledger({"sourceRecordId": "source-one", "reviewDecision": "", "poemKey": "KEY-1"})
        result = upsert_excerpt_records(connection, {"records": [{
            "sourceRecordId": "source-one", "reviewDecision": "ACCEPT_SKIP_GRAPHIC", "approvedForUse": True,
        }]})
        self.assertEqual(len(connection.writes), 1)
        self.assertEqual(result["records"][0]["reviewDecision"], "ACCEPT_SKIP_GRAPHIC")
        self.assertEqual(result["records"][0]["poemKey"], "KEY-1")
