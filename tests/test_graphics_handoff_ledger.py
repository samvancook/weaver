import sqlite3
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from weaver_runtime_db import (  # noqa: E402
    assert_completion_identity_consistent,
    assert_editable_project_file_not_cross_linked,
    build_graphics_qc_queue_card,
    canonical_qi_content_id,
    canonical_qi_graphics_request_id,
    claim_graphics_handoff,
    ensure_runtime_schema,
    find_editable_project_link_conflicts,
    FirestoreLedgerClient,
    get_excerpt_handoff,
    get_graphics_handoff,
    get_graphics_handoff_queue,
    insert_graphics_completion,
    insert_graphics_qc_review,
    normalize_handoff_record,
    queue_card_record,
    update_graphics_handoff,
    upsert_excerpt_handoff,
    upsert_graphics_handoff_request,
    upsert_graphics_request,
)
from weaver_runtime_sync import (  # noqa: E402
    reconcile_repair_request,
    sync_completions,
    sync_excerpt_handoffs,
    sync_qc_reviews,
    update_repair_request,
    upsert_repair_requests,
)


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


class MemoryFirestoreLedger(FirestoreLedgerClient):
    def __init__(self):
        self.documents = {}
        self.handoffs = {}
        self.handoff_calls = []
        self.handoff_update_calls = []

    def get_raw_document(self, collection, document_id):
        return self.documents.get((collection, document_id))

    def write_raw_document(self, collection, document_id, record):
        written = dict(record)
        self.documents[(collection, document_id)] = written
        return written

    def list_raw_documents(self, collection, page_size=500):
        return [
            record for (stored_collection, _), record in self.documents.items()
            if stored_collection == collection
        ]

    def upsert_handoff_request(self, request):
        request_id = request["graphicsRequestId"]
        self.handoff_calls.append(request_id)
        self.handoffs[request_id] = dict(request)
        return dict(request)

    def get_handoff(self, graphics_request_id):
        return self.handoffs.get(graphics_request_id)

    def update_handoff(self, graphics_request_id, update):
        self.handoff_update_calls.append(graphics_request_id)
        current = dict(self.handoffs[graphics_request_id])
        current.update(update)
        self.handoffs[graphics_request_id] = current
        return dict(current)

    def insert_graphics_completion(self, completion):
        completion_id = completion["id"]
        self.documents[("graphicsCompletions", completion_id)] = dict(completion)
        return dict(completion)


class RepairIdentityFirestoreLedger(FirestoreLedgerClient):
    def __init__(self):
        self.records = {}
        self.canonical_query_count = 0

    def get_handoff(self, graphics_request_id):
        return self.records.get(graphics_request_id)

    def query_raw_documents(self, collection, field, value, page_size=100):
        self.canonical_query_count += 1
        return []

    def write_record(self, record):
        self.records[record["graphicsRequestId"]] = dict(record)
        return dict(record)

    def write_raw_document(self, collection, document_id, record):
        return dict(record)


class GraphicsHandoffLedgerTest(unittest.TestCase):
    def test_external_repair_preserves_stable_job_id_without_canonical_collapse(self):
        connection = RepairIdentityFirestoreLedger()
        repair_job_id = "weaver:repair:stable-canary"
        written = connection.upsert_handoff_request({
            "graphicsRequestId": repair_job_id,
            "sourceSystem": "poetry_please_repair",
            "sourceStatus": "repair_requested",
            "contentType": "QI",
            "repairRequestId": "repair-canary-stable",
            "sourcePayload": {
                "repairRequestId": "repair-canary-stable",
                "bookTitle": "Test Book",
                "poemTitle": "Test Poem",
                "author": "Test Author",
                "quoteText": "Canonical excerpt text",
                "queueView": "rework",
            },
        })

        self.assertEqual(written["graphicsRequestId"], repair_job_id)
        self.assertEqual(written["canonicalContentId"], "REPAIR:repair-canary-stable")
        self.assertEqual(connection.canonical_query_count, 0)

    def test_poetry_please_repair_ingest_is_idempotent_and_routes_by_type(self):
        connection = MemoryFirestoreLedger()
        qi_request = {
            "id": "repair-canary-1",
            "status": "requested",
            "action": "recreate",
            "sourceFlagId": "flag-1",
            "imageId": "original-qi-1",
            "contentType": "QI",
            "author": "Test Author",
            "book": "Test Book",
            "title": "Test Poem",
            "issueReason": "Incorrect text",
            "requestNote": "Recreate from canonical excerpt",
            "originalContent": {"quoteText": "Canonical excerpt text"},
        }

        first = upsert_repair_requests(connection, {"requests": [qi_request]})
        second = upsert_repair_requests(connection, {"requests": [qi_request]})
        exc = upsert_repair_requests(connection, {"requests": [{
            **qi_request,
            "id": "repair-canary-exc",
            "imageId": "original-exc-1",
            "contentType": "EXC",
        }]})

        self.assertEqual(first["createdCount"], 1)
        self.assertEqual(first["records"][0]["repairDestination"], "pig")
        self.assertTrue(first["records"][0]["pigJobId"].startswith("weaver:repair:"))
        self.assertEqual(second["duplicateCount"], 1)
        self.assertEqual(connection.handoff_calls.count(first["records"][0]["pigJobId"]), 1)
        queue_card = queue_card_record(normalize_handoff_record({
            **connection.handoffs[first["records"][0]["pigJobId"]],
            "sourcePayload": connection.handoffs[first["records"][0]["pigJobId"]]["sourcePayload"],
        }))
        self.assertEqual(queue_card["repairRequestId"], "repair-canary-1")
        self.assertEqual(queue_card["sourceFlagId"], "flag-1")
        self.assertEqual(queue_card["originalContentId"], "original-qi-1")
        self.assertEqual(queue_card["repairInstructions"], "Recreate from canonical excerpt")
        self.assertEqual(exc["createdCount"], 1)
        self.assertEqual(exc["records"][0]["repairDestination"], "weaver_review")
        self.assertFalse(exc["records"][0]["pigJobId"])

    def test_repair_completion_returns_structured_replacement_identity(self):
        connection = MemoryFirestoreLedger()
        request = {
            "id": "repair-return-1",
            "status": "requested",
            "action": "recreate",
            "sourceFlagId": "flag-return-1",
            "imageId": "original-return-1",
            "contentType": "FPI",
            "author": "Test Author",
            "book": "Test Book",
            "title": "Test Poem",
            "originalContent": {"imageUrl": "https://example.com/original.png"},
        }
        imported = upsert_repair_requests(connection, {"requests": [request]})
        weaver_job_id = imported["records"][0]["pigJobId"]

        completed = sync_completions(connection, {"completions": [{
            "completionId": "pig-repair-return-1",
            "graphicsRequestId": weaver_job_id,
            "repairRequestId": "repair-return-1",
            "contentType": "FPI",
            "assetFileId": "replacement-file-1",
            "assetUrl": "https://example.com/replacement.png",
        }]})

        self.assertEqual(completed["repairReturns"][0]["replacementAssetId"], "replacement-file-1")
        self.assertEqual(completed["repairReturns"][0]["replacementAssetLink"], "https://example.com/replacement.png")
        self.assertEqual(completed["repairReturns"][0]["weaverJobId"], weaver_job_id)
        self.assertEqual(completed["repairReturns"][0]["pigJobId"], "pig-repair-return-1")

    def test_accepted_repair_resolves_once(self):
        connection = MemoryFirestoreLedger()
        request = {
            "id": "repair-accepted-1",
            "status": "requested",
            "action": "recreate",
            "sourceFlagId": "flag-accepted-1",
            "imageId": "original-accepted-1",
            "contentType": "QI",
            "author": "Test Author",
            "book": "Test Book",
            "title": "Test Poem",
            "originalContent": {"quoteText": "Canonical excerpt"},
        }
        upsert_repair_requests(connection, {"requests": [request]})
        update_repair_request(connection, {
            "repairRequestId": "repair-accepted-1",
            "update": {
                "weaverRepairStatus": "returned",
                "replacementAssetId": "replacement-accepted-1",
                "replacementAssetLink": "https://example.com/accepted.png",
            },
        })
        poetry_please_request = {
            "id": "repair-accepted-1",
            "status": "resolved",
            "returnReviewStatus": "accepted",
            "returnReviewedAt": "2026-08-03T12:00:00Z",
            "returnReviewedBy": "admin@buttonpoetry.com",
            "replacementAssetId": "replacement-accepted-1",
            "replacementAssetLink": "https://example.com/accepted.png",
        }

        first = reconcile_repair_request(connection, {
            "repairRequestId": "repair-accepted-1",
            "poetryPleaseRequest": poetry_please_request,
            "replacementAssetAvailable": True,
        })
        second = reconcile_repair_request(connection, {
            "repairRequestId": "repair-accepted-1",
            "poetryPleaseRequest": poetry_please_request,
            "replacementAssetAvailable": True,
        })

        self.assertEqual(first["transition"], "resolved")
        self.assertEqual(first["record"]["weaverRepairStatus"], "resolved")
        self.assertEqual(first["record"]["resolvedAt"], "2026-08-03T12:00:00Z")
        self.assertEqual(second["transition"], "none")

    def test_returned_repair_records_pending_review_once(self):
        connection = MemoryFirestoreLedger()
        request = {
            "id": "repair-pending-1",
            "status": "requested",
            "action": "recreate",
            "sourceFlagId": "flag-pending-1",
            "imageId": "original-pending-1",
            "contentType": "QI",
            "author": "Test Author",
            "book": "Test Book",
            "title": "Test Poem",
            "originalContent": {"quoteText": "Canonical excerpt"},
        }
        upsert_repair_requests(connection, {"requests": [request]})
        update_repair_request(connection, {
            "repairRequestId": "repair-pending-1",
            "update": {
                "weaverRepairStatus": "returned",
                "replacementAssetId": "replacement-pending-1",
                "replacementAssetLink": "https://example.com/pending.png",
            },
        })
        poetry_please_request = {
            "id": "repair-pending-1",
            "status": "returned",
            "returnReviewStatus": "pending",
            "replacementAssetId": "replacement-pending-1",
            "replacementAssetLink": "https://example.com/pending.png",
        }

        first = reconcile_repair_request(connection, {
            "repairRequestId": "repair-pending-1",
            "poetryPleaseRequest": poetry_please_request,
            "replacementAssetAvailable": True,
        })
        second = reconcile_repair_request(connection, {
            "repairRequestId": "repair-pending-1",
            "poetryPleaseRequest": poetry_please_request,
            "replacementAssetAvailable": True,
        })

        self.assertEqual(first["transition"], "observed")
        self.assertEqual(first["record"]["weaverRepairStatus"], "returned")
        self.assertEqual(first["record"]["returnReviewStatus"], "pending")
        self.assertEqual(second["transition"], "none")

    def test_rejected_repair_reopens_same_pig_job_once(self):
        connection = MemoryFirestoreLedger()
        request = {
            "id": "repair-rejected-1",
            "status": "requested",
            "action": "recreate",
            "sourceFlagId": "flag-rejected-1",
            "imageId": "original-rejected-1",
            "contentType": "FPI",
            "author": "Test Author",
            "book": "Test Book",
            "title": "Test Poem",
            "originalContent": {"imageUrl": "https://example.com/original.png"},
        }
        imported = upsert_repair_requests(connection, {"requests": [request]})
        pig_job_id = imported["records"][0]["pigJobId"]
        update_repair_request(connection, {
            "repairRequestId": "repair-rejected-1",
            "update": {
                "weaverRepairStatus": "returned",
                "replacementAssetId": "replacement-rejected-1",
                "replacementAssetLink": "https://example.com/rejected.png",
            },
        })
        poetry_please_request = {
            "id": "repair-rejected-1",
            "status": "returned",
            "returnReviewStatus": "rejected",
            "returnReviewNote": "Attribution still needs correction.",
        }

        first = reconcile_repair_request(connection, {
            "repairRequestId": "repair-rejected-1",
            "poetryPleaseRequest": poetry_please_request,
            "replacementAssetAvailable": True,
        })
        second = reconcile_repair_request(connection, {
            "repairRequestId": "repair-rejected-1",
            "poetryPleaseRequest": poetry_please_request,
            "replacementAssetAvailable": True,
        })

        self.assertEqual(first["transition"], "reopened")
        self.assertEqual(first["record"]["weaverRepairStatus"], "waiting_on_pig")
        self.assertEqual(first["record"]["returnReviewNote"], "Attribution still needs correction.")
        self.assertEqual(connection.handoffs[pig_job_id]["handoffStatus"], "requested")
        self.assertEqual(connection.handoffs[pig_job_id]["pigStatus"], "not_started")
        self.assertEqual(connection.handoff_update_calls, [pig_job_id])
        self.assertEqual(second["transition"], "none")
        self.assertEqual(connection.handoff_calls.count(pig_job_id), 1)

    def test_qc_card_preserves_structured_drive_validation_error(self):
        error = {
            "failingService": "google_drive",
            "operation": "drive.files.get?alt=media&supportsAllDrives=true",
            "fileId": "drive-json-file",
            "httpStatus": 404,
        }
        card = build_graphics_qc_queue_card({
            "id": "pig-completion-drive-warning",
            "graphicsRequestId": "weaver:fpi:drive-warning",
            "contentType": "FPI",
            "assetUrl": "https://drive.google.com/file/d/png-file/view",
            "editableProjectAvailable": False,
            "editableProjectValidationStatus": "drive_inaccessible",
            "editableProjectValidationError": error,
        })

        self.assertEqual(card["assetLinkUrl"], "https://drive.google.com/file/d/png-file/view")
        self.assertFalse(card["editableProjectAvailable"])
        self.assertEqual(card["editableProjectValidationError"], error)

    def test_rework_queue_splits_qi_and_fpi_lanes(self):
        with memory_db() as connection:
            for request_id, content_type in (("weaver:qi:test", "QI"), ("weaver:fpi:test", "FPI")):
                upsert_graphics_handoff_request(connection, {
                    "graphicsRequestId": request_id,
                    "contentType": content_type,
                    "imageType": content_type,
                    "sourcePayload": {"quoteText": f"{content_type} text"},
                })
                update_graphics_handoff(connection, request_id, {
                    "handoffStatus": "rejected",
                    "pigStatus": "not_started",
                    "qcStatus": "needs_revision",
                })

            qi_records = get_graphics_handoff_queue(connection, filter_mode="rework", content_type="QI")
            fpi_records = get_graphics_handoff_queue(connection, filter_mode="rework", content_type="FPI")

        self.assertEqual([record["graphicsRequestId"] for record in qi_records], ["weaver:qi:test"])
        self.assertEqual([record["graphicsRequestId"] for record in fpi_records], ["weaver:fpi:test"])
        self.assertEqual(qi_records[0]["queueLane"], "QI")
        self.assertEqual(fpi_records[0]["reworkLane"], "FPI")

    def test_completion_rejects_conflicting_request_identity(self):
        with self.assertRaisesRegex(ValueError, "conflicting graphicsRequestId"):
            assert_completion_identity_consistent({
                "graphicsRequestId": "weaver:fpi:book-a-poem-a",
                "sourcePayload": {"requestId": "weaver:fpi:book-b-poem-b"},
            })

    def test_completion_rejects_conflicting_text_hash(self):
        with self.assertRaisesRegex(ValueError, "textHash"):
            assert_completion_identity_consistent({
                "graphicsRequestId": "weaver:fpi:book-a-poem-a",
                "textHash": "0" * 64,
                "sourcePayload": {"quoteText": "Actual poem text"},
            })

    def test_consecutive_fpi_completions_cannot_cross_link_editable_projects(self):
        existing = [{
            "id": "pig-completion-a",
            "graphicsRequestId": "weaver:fpi:book-a-poem-a",
            "editableProjectFileId": "drive-editable-shared",
        }]
        candidate = {
            "id": "pig-completion-b",
            "graphicsRequestId": "weaver:fpi:book-b-poem-b",
            "editableProjectFileId": "drive-editable-shared",
        }

        with self.assertRaisesRegex(ValueError, "refusing cross-link"):
            assert_editable_project_file_not_cross_linked(existing, candidate)

    def test_editable_project_reconciliation_reports_conflicting_identities(self):
        conflicts = find_editable_project_link_conflicts([
            {
                "id": "pig-completion-a",
                "graphicsRequestId": "weaver:fpi:book-a-poem-a",
                "editableProjectFileId": "drive-editable-shared",
            },
            {
                "id": "pig-completion-b",
                "graphicsRequestId": "weaver:fpi:book-b-poem-b",
                "editableProjectFileId": "drive-editable-shared",
            },
        ])

        self.assertEqual(len(conflicts), 1)
        self.assertEqual(conflicts[0]["editableProjectFileId"], "drive-editable-shared")

    def test_handoff_queue_preserves_canonical_book_key(self):
        connection = memory_db()
        upsert_graphics_handoff_request(connection, {
            "graphicsRequestId": "weaver:qi:book-key",
            "sourceSystem": "weaver",
            "sourceStatus": "open",
            "contentType": "QI",
            "sourcePayload": {
                "bookTitle": "Book Display Title",
                "bookKey": "canonical-book-key",
                "quoteText": "Test quote",
            },
        })

        records = get_graphics_handoff_queue(connection, limit=10, filter_mode="current_titles")

        self.assertEqual(records[0]["bookKey"], "canonical-book-key")

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
