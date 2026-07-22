import sqlite3
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from weaver_runtime_db import (  # noqa: E402
    build_graphics_qc_queue_card,
    canonical_qi_content_id,
    canonical_qi_graphics_request_id,
    claim_graphics_handoff,
    ensure_runtime_schema,
    get_excerpt_handoff,
    get_graphics_handoff,
    get_graphics_handoff_queue,
    insert_graphics_completion,
    insert_graphics_qc_review,
    normalize_handoff_record,
    update_graphics_handoff,
    upsert_excerpt_handoff,
    upsert_graphics_handoff_request,
    upsert_graphics_request,
)
from weaver_runtime_sync import sync_excerpt_handoffs, sync_qc_reviews  # noqa: E402


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
            "contentType": "QI",
            "sourcePayload": {
                "sourceSheetRow": 2130,
                "bookTitle": "Test Book",
                "poemTitle": "Test Poem",
                "quoteText": "Test quote",
            },
        },
    )


class GraphicsHandoffLedgerTest(unittest.TestCase):
    def test_rework_explicitly_reports_missing_editable_project(self):
        record = normalize_handoff_record({
            "graphicsRequestId": "weaver:qi:test",
            "contentType": "QI",
            "imageType": "QI",
            "handoffStatus": "rejected",
            "pigStatus": "not_started",
            "qcStatus": "needs_revision",
            "quoteText": "Test quote",
        })

        self.assertFalse(record["editableProjectAvailable"])

    def test_rework_reports_complete_editable_project_identity(self):
        record = normalize_handoff_record({
            "graphicsRequestId": "weaver:qi:test",
            "contentType": "QI",
            "imageType": "QI",
            "handoffStatus": "rejected",
            "pigStatus": "not_started",
            "qcStatus": "needs_revision",
            "quoteText": "Test quote",
            "pigProjectId": "project-real",
            "editableProjectFileId": "drive-file-real",
            "editableProjectUrl": "https://drive.google.com/file/d/drive-file-real/view",
        })

        self.assertTrue(record["editableProjectAvailable"])

    def test_qc_queue_card_preserves_editable_project_identity(self):
        card = build_graphics_qc_queue_card({
            "id": "pig-completion-real",
            "graphicsRequestId": "weaver:qi:test",
            "contentType": "QI",
            "sourcePayload": {
                "pigProjectId": "project-real",
                "editableProjectFileId": "drive-file-real",
                "editableProjectUrl": "https://drive.google.com/file/d/drive-file-real/view",
            },
        })

        self.assertEqual(card["pigProjectId"], "project-real")
        self.assertEqual(card["editableProjectFileId"], "drive-file-real")
        self.assertTrue(card["editableProjectAvailable"])

    def test_qi_identity_does_not_depend_on_sheet_row(self):
        first = {
            "graphicsRequestId": "weaver:row-258",
            "contentType": "QI",
            "sourcePayload": {
                "queueSheetRow": 258,
                "author": "Angela Spinzig",
                "poemTitle": "In a Gaza refugee camp",
                "bookTitle": "Short Form Contest May 2026",
                "quoteText": "The same excerpt text",
            },
        }
        moved = {
            **first,
            "graphicsRequestId": "weaver:row-353",
            "sourcePayload": {**first["sourcePayload"], "queueSheetRow": 353},
        }

        first_content_id = canonical_qi_content_id(first, first["sourcePayload"])
        moved_content_id = canonical_qi_content_id(moved, moved["sourcePayload"])

        self.assertEqual(first_content_id, moved_content_id)
        self.assertEqual(
            canonical_qi_graphics_request_id(first_content_id),
            canonical_qi_graphics_request_id(moved_content_id),
        )

    def test_qi_identity_distinguishes_different_excerpts(self):
        request = {
            "contentType": "QI",
            "sourcePayload": {
                "author": "Test Author",
                "poemTitle": "Test Poem",
                "bookTitle": "Test Book",
                "quoteText": "First excerpt",
            },
        }
        changed = {
            **request,
            "sourcePayload": {**request["sourcePayload"], "quoteText": "Second excerpt"},
        }

        self.assertNotEqual(
            canonical_qi_content_id(request, request["sourcePayload"]),
            canonical_qi_content_id(changed, changed["sourcePayload"]),
        )

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

    def test_generated_exported_uploaded_or_sent_records_leave_handoff_queue(self):
        terminal_updates = [
            ("weaver:generated", {"handoffStatus": "generated", "pigStatus": "generated"}),
            ("weaver:exported", {"handoffStatus": "exported", "pigStatus": "exported", "exportType": "download_png"}),
            ("weaver:uploaded", {"handoffStatus": "uploaded", "pigStatus": "uploaded"}),
            ("weaver:sent", {"handoffStatus": "sent_to_weaver_qc", "pigStatus": "uploaded", "qcStatus": "pending"}),
        ]
        with memory_db() as connection:
            for request_id, update in terminal_updates:
                seed_request(connection, request_id)
                update_graphics_handoff(connection, request_id, update)
            queue = get_graphics_handoff_queue(connection)

        self.assertEqual(queue, [])

    def test_only_revision_rejects_return_to_pig_queue(self):
        with memory_db() as connection:
            for request_id, completion_id in [
                ("weaver:revision", "pig-revision"),
                ("weaver:final", "pig-final"),
                ("weaver:mismatch", "pig-mismatch"),
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
                            "contentType": "QI",
                            "qcDecision": "reject",
                            "rejectReason": "correct_and_recreate",
                        },
                        {
                            "storageTarget": "pig_sheet",
                            "graphicsRequestId": "weaver:final",
                            "pigCompletionId": "pig-final",
                            "contentType": "QI",
                            "qcDecision": "reject",
                            "rejectReason": "final_reject",
                        },
                        {
                            "storageTarget": "pig_sheet",
                            "graphicsRequestId": "weaver:mismatch",
                            "pigCompletionId": "pig-mismatch",
                            "contentType": "QI",
                            "qcDecision": "reject",
                            "rejectReason": "mismatched_graphic",
                        },
                    ]
                },
            )
            queue_ids = {record["graphicsRequestId"] for record in get_graphics_handoff_queue(connection)}
            revision = get_graphics_handoff(connection, "weaver:revision")
            final = get_graphics_handoff(connection, "weaver:final")
            mismatch = get_graphics_handoff(connection, "weaver:mismatch")

        self.assertIn("weaver:revision", queue_ids)
        self.assertNotIn("weaver:final", queue_ids)
        self.assertNotIn("weaver:mismatch", queue_ids)
        self.assertEqual(revision["qcStatus"], "needs_revision")
        self.assertEqual(revision["pigStatus"], "not_started")
        self.assertEqual(final["qcStatus"], "rejected")
        self.assertEqual(mismatch["qcStatus"], "rejected")

    def test_qc_approve_removes_request_from_queue(self):
        with memory_db() as connection:
            upsert_graphics_request(
                connection,
                {
                    "id": "weaver:approve",
                    "book_title": "Test Book",
                    "poem_title": "Test Poem",
                    "author": "Test Author",
                    "quote_text": "Test quote",
                },
            )
            seed_request(connection, "weaver:approve")
            insert_graphics_completion(
                connection,
                {
                    "id": "pig-approve",
                    "graphics_request_id": "weaver:approve",
                    "asset_url": "https://drive.google.com/file/d/abc/view",
                },
            )
            update_graphics_handoff(
                connection,
                "weaver:approve",
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
                            "graphicsRequestId": "weaver:approve",
                            "pigCompletionId": "pig-approve",
                            "contentType": "QI",
                            "qcDecision": "approve",
                        }
                    ]
                },
            )
            queue_ids = {record["graphicsRequestId"] for record in get_graphics_handoff_queue(connection)}
            approved = get_graphics_handoff(connection, "weaver:approve")

        self.assertNotIn("weaver:approve", queue_ids)
        self.assertEqual(approved["handoffStatus"], "approved")
        self.assertEqual(approved["qcStatus"], "approved")
        self.assertTrue(approved["approvedAt"])

    def test_blank_request_text_never_appears_in_handoff_queue(self):
        with memory_db() as connection:
            record = upsert_graphics_handoff_request(
                connection,
                {
                    "graphicsRequestId": "weaver:row-3",
                    "sourceSystem": "weaver",
                    "sourceStatus": "needs_graphics",
                    "contentType": "QI",
                    "sourcePayload": {
                        "sourceSheetRow": 3,
                        "bookTitle": "Blank Book",
                        "poemTitle": "Blank Poem",
                        "quoteText": "",
                    },
                },
            )
            queue_ids = {item["graphicsRequestId"] for item in get_graphics_handoff_queue(connection)}
            claimed = claim_graphics_handoff(connection, "weaver:row-3", "pig-worker")

        self.assertEqual(record["handoffStatus"], "blocked")
        self.assertEqual(record["pigStatus"], "failed")
        self.assertEqual(record["blockedReason"], "blank_request_text")
        self.assertNotIn("weaver:row-3", queue_ids)
        self.assertEqual(claimed["handoffStatus"], "blocked")

    def test_excerpt_handoff_uses_stable_record_id(self):
        with memory_db() as connection:
            record = upsert_excerpt_handoff(
                connection,
                {
                    "recordId": "weaver-exc-8722",
                    "sourceRecordId": "20260514131148-8722",
                    "author": "Buddy Wakefield",
                    "bookTitle": "Stunt Water",
                    "poemTitle": "Flockprinter",
                    "excerpt": "Even before we met.",
                    "bookShortener": "SWTWOB",
                    "handoffStatus": "queued",
                },
            )
            again = upsert_excerpt_handoff(
                connection,
                {
                    "recordId": "weaver-exc-8722",
                    "sourceRecordId": "20260514131148-8722",
                    "author": "Buddy Wakefield",
                    "bookTitle": "Stunt Water",
                    "poemTitle": "Flockprinter",
                    "excerpt": "Even before we met.",
                    "handoffStatus": "sent",
                },
            )
            count = connection.execute("SELECT COUNT(*) FROM excerpt_handoff_ledger").fetchone()[0]

        self.assertEqual(count, 1)
        self.assertEqual(record["recordId"], "weaver-exc-8722")
        self.assertEqual(again["handoffStatus"], "sent")

    def test_excerpt_handoff_sync_roundtrip(self):
        with memory_db() as connection:
            sync_excerpt_handoffs(
                connection,
                {
                    "handoff": {
                        "recordId": "weaver-exc-row-77",
                        "contentType": "EXC",
                        "sourceSystem": "weaver",
                        "sourceRecordId": "weaver:row-77",
                        "author": "Gigi Bella",
                        "bookTitle": "Without the Frills",
                        "poemTitle": "the ikea poem",
                        "excerpt": "only as long as i remember",
                        "approvedAt": "2026-05-15T10:00:00Z",
                        "payload": {"sourceRow": 77},
                    }
                },
            )
            record = get_excerpt_handoff(connection, "weaver-exc-row-77")

        self.assertIsNotNone(record)
        assert record is not None
        self.assertEqual(record["contentType"], "EXC")
        self.assertEqual(record["sourceRecordId"], "weaver:row-77")
        self.assertEqual(record["poemTitle"], "the ikea poem")
        self.assertEqual(record["payload"], {"sourceRow": 77})


if __name__ == "__main__":
    unittest.main()
