import sqlite3
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from weaver_runtime_db import (  # noqa: E402
    claim_graphics_handoff,
    ensure_runtime_schema,
    get_graphics_handoff,
    get_graphics_handoff_queue,
    insert_graphics_completion,
    insert_graphics_qc_review,
    update_graphics_handoff,
    upsert_graphics_handoff_request,
    upsert_graphics_request,
)
from weaver_runtime_sync import sync_qc_reviews  # noqa: E402


def memory_db() -> sqlite3.Connection:
    connection = sqlite3.connect(":memory:")
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    ensure_runtime_schema(connection)
    return connection


def seed_request(connection: sqlite3.Connection, request_id: str = "weaver:row-2130") -> dict:
    return upsert_graphics_handoff_request(
        connection,
        {
            "graphicsRequestId": request_id,
            "sourceSystem": "weaver",
            "sourceStatus": "needs_graphics",
            "sourcePayload": {
                "sourceSheetRow": 2130,
                "bookTitle": "Test Book",
                "poemTitle": "Test Poem",
            },
        },
    )


class GraphicsHandoffLedgerTest(unittest.TestCase):
    def test_create_request(self):
        with memory_db() as connection:
            record = seed_request(connection)

        self.assertEqual(record["graphicsRequestId"], "weaver:row-2130")
        self.assertEqual(record["handoffStatus"], "requested")
        self.assertEqual(record["pigStatus"], "not_started")
        self.assertEqual(record["qcStatus"], "not_sent")

    def test_upsert_same_request_twice_does_not_duplicate(self):
        with memory_db() as connection:
            seed_request(connection)
            seed_request(connection)
            count = connection.execute("SELECT COUNT(*) FROM graphics_handoff_ledger").fetchone()[0]

        self.assertEqual(count, 1)

    def test_claim_request(self):
        with memory_db() as connection:
            seed_request(connection)
            record = claim_graphics_handoff(connection, "weaver:row-2130", "pig-worker")

        self.assertEqual(record["handoffStatus"], "claimed")
        self.assertEqual(record["pigStatus"], "claimed")
        self.assertEqual(record["claimedBy"], "pig-worker")
        self.assertTrue(record["claimedAt"])

    def test_mark_generated_and_uploaded(self):
        with memory_db() as connection:
            seed_request(connection)
            generated = update_graphics_handoff(
                connection,
                "weaver:row-2130",
                {"handoffStatus": "generated", "pigStatus": "generated"},
            )
            uploaded = update_graphics_handoff(
                connection,
                "weaver:row-2130",
                {
                    "handoffStatus": "uploaded",
                    "pigStatus": "uploaded",
                    "assetUrl": "https://drive.google.com/file/d/abc/view",
                    "driveFileId": "abc",
                    "driveFileName": "test.png",
                    "mimeType": "image/png",
                    "exportType": "png",
                    "variant": "default",
                    "version": "1",
                },
            )

        self.assertTrue(generated["generatedAt"])
        self.assertEqual(uploaded["assetUrl"], "https://drive.google.com/file/d/abc/view")
        self.assertEqual(uploaded["driveFileId"], "abc")
        self.assertTrue(uploaded["uploadedAt"])

    def test_send_to_qc(self):
        with memory_db() as connection:
            seed_request(connection)
            record = update_graphics_handoff(
                connection,
                "weaver:row-2130",
                {
                    "handoffStatus": "sent_to_weaver_qc",
                    "pigStatus": "uploaded",
                    "qcStatus": "pending",
                    "assetUrl": "https://drive.google.com/file/d/abc/view",
                },
            )

        self.assertEqual(record["handoffStatus"], "sent_to_weaver_qc")
        self.assertEqual(record["qcStatus"], "pending")
        self.assertTrue(record["sentToQcAt"])

    def test_approval_and_rejection(self):
        with memory_db() as connection:
            seed_request(connection, "weaver:approve")
            approved = update_graphics_handoff(
                connection,
                "weaver:approve",
                {"handoffStatus": "approved", "qcStatus": "approved"},
            )
            seed_request(connection, "weaver:reject")
            rejected = update_graphics_handoff(
                connection,
                "weaver:reject",
                {"handoffStatus": "rejected", "qcStatus": "needs_revision"},
            )

        self.assertTrue(approved["approvedAt"])
        self.assertTrue(rejected["rejectedAt"])
        self.assertEqual(rejected["qcStatus"], "needs_revision")

    def test_duplicate_prevention_primary_key(self):
        with memory_db() as connection:
            seed_request(connection)
            with self.assertRaises(sqlite3.IntegrityError):
                connection.execute(
                    """
                    INSERT INTO graphics_handoff_ledger (
                        graphics_request_id, created_at, updated_at
                    ) VALUES ('weaver:row-2130', '2026-01-01T00:00:00Z', '2026-01-01T00:00:00Z')
                    """
                )

    def test_completed_request_is_not_returned_as_open_queue_work(self):
        with memory_db() as connection:
            upsert_graphics_request(
                connection,
                {
                    "id": "weaver:row-2130",
                    "book_title": "Test Book",
                    "poem_title": "Test Poem",
                    "author": "Test Author",
                    "quote_text": "Test quote",
                },
            )
            seed_request(connection)
            insert_graphics_completion(
                connection,
                {
                    "id": "pig-weaver-row-2130",
                    "graphics_request_id": "weaver:row-2130",
                    "asset_url": "https://drive.google.com/file/d/abc/view",
                },
            )
            update_graphics_handoff(
                connection,
                "weaver:row-2130",
                {
                    "handoffStatus": "sent_to_weaver_qc",
                    "pigStatus": "uploaded",
                    "qcStatus": "pending",
                    "assetUrl": "https://drive.google.com/file/d/abc/view",
                },
            )
            queue = get_graphics_handoff_queue(connection)

        self.assertEqual(queue, [])
        self.assertEqual(get_graphics_handoff(connection, "weaver:row-2130")["pigStatus"], "uploaded")

    def test_only_revision_rejects_return_to_pig_queue(self):
        with memory_db() as connection:
            for request_id, completion_id in [
                ("weaver:revision", "pig-revision"),
                ("weaver:final", "pig-final"),
            ]:
                upsert_graphics_request(
                    connection,
                    {
                        "id": request_id,
                        "book_title": "Test Book",
                        "poem_title": "Test Poem",
                        "author": "Test Author",
                        "quote_text": "Test quote",
                    },
                )
                seed_request(connection, request_id)
                insert_graphics_completion(
                    connection,
                    {
                        "id": completion_id,
                        "graphics_request_id": request_id,
                        "asset_url": "https://drive.google.com/file/d/abc/view",
                    },
                )
                update_graphics_handoff(
                    connection,
                    request_id,
                    {
                        "handoffStatus": "sent_to_weaver_qc",
                        "pigStatus": "uploaded",
                        "qcStatus": "pending",
                        "assetUrl": "https://drive.google.com/file/d/abc/view",
                    },
                )

            sync_qc_reviews(
                connection,
                {
                    "reviews": [
                        {
                            "storageTarget": "pig_sheet",
                            "graphicsRequestId": "weaver:revision",
                            "pigCompletionId": "pig-revision",
                            "qcDecision": "reject",
                            "rejectReason": "correct_and_recreate",
                        },
                        {
                            "storageTarget": "pig_sheet",
                            "graphicsRequestId": "weaver:final",
                            "pigCompletionId": "pig-final",
                            "qcDecision": "reject",
                            "rejectReason": "final_reject",
                        },
                    ]
                },
            )
            queue_ids = {record["graphicsRequestId"] for record in get_graphics_handoff_queue(connection)}
            revision = get_graphics_handoff(connection, "weaver:revision")
            final = get_graphics_handoff(connection, "weaver:final")

        self.assertIn("weaver:revision", queue_ids)
        self.assertNotIn("weaver:final", queue_ids)
        self.assertEqual(revision["qcStatus"], "needs_revision")
        self.assertEqual(revision["pigStatus"], "not_started")
        self.assertEqual(final["qcStatus"], "rejected")


if __name__ == "__main__":
    unittest.main()
