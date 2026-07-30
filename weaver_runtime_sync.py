#!/usr/bin/env python3
from __future__ import annotations

import json
import hashlib
import re
import os
import sys
from typing import Any

from weaver_runtime_db import (
    FIRESTORE_EXCERPT_RECORDS_COLLECTION,
    FIRESTORE_GRAPHICS_COMPLETIONS_COLLECTION,
    FIRESTORE_GRAPHICS_QC_QUEUE_COLLECTION,
    connect_firestore_ledger,
    connect_runtime_db,
    ensure_runtime_schema,
    find_editable_project_link_conflicts,
    claim_graphics_handoff,
    get_excerpt_handoffs,
    get_graphics_handoff,
    get_graphics_handoffs,
    get_graphics_handoff_queue,
    get_latest_graphics_qc_reviews,
    get_latest_poetry_please_handoffs,
    get_pending_graphics_qc_records,
    rebuild_graphics_qc_queue_cards,
    insert_graphics_completion,
    insert_graphics_qc_review,
    insert_poetry_please_handoff,
    replace_graphics_request_items,
    upsert_excerpt_handoff,
    update_graphics_handoff,
    upsert_graphics_handoff_request,
    upsert_graphics_request,
    utc_now_iso,
)


FIRESTORE_HANDOFF_ACTIONS = {
    "upsert_completions",
    "insert_qc_reviews",
    "insert_poetry_please_handoffs",
    "get_graphics_state",
    "get_pending_graphics_qc",
    "rebuild_graphics_qc_queue",
    "upsert_graphics_qc_queue_cards",
    "upsert_handoff_requests",
    "get_handoff_queue",
    "claim_handoff_request",
    "patch_handoff_request",
    "get_handoff_request",
    "get_handoff_requests",
    "get_graphics_completion",
    "audit_editable_project_links",
    "recover_stalled_pig_requests",
    "upsert_repair_requests",
    "list_repair_requests",
    "update_repair_request",
    "upsert_excerpt_records",
    "get_excerpt_records",
}

EXPECTED_FIRESTORE_PROJECT_ID = "button-weaver-internal"
EXPECTED_FIRESTORE_DATABASE_ID = "weaverledger"
FIRESTORE_REPAIR_REQUESTS_COLLECTION = "poetryPleaseRepairRequests"
PIG_REPAIR_CONTENT_TYPES = {"QI", "FPI"}


def validate_production_handoff_backend(action: str) -> None:
    if not os.environ.get("K_SERVICE") or action not in FIRESTORE_HANDOFF_ACTIONS:
        return
    expected = {
        "WEAVER_LEDGER_BACKEND": "firestore",
        "WEAVER_FIRESTORE_PROJECT_ID": EXPECTED_FIRESTORE_PROJECT_ID,
        "WEAVER_FIRESTORE_DATABASE_ID": EXPECTED_FIRESTORE_DATABASE_ID,
    }
    invalid = [
        f"{name}={os.environ.get(name, '')!r}"
        for name, value in expected.items()
        if os.environ.get(name, "").strip() != value
    ]
    if invalid:
        raise RuntimeError(
            "Production graphics handoff storage is misconfigured: " + ", ".join(invalid)
        )


def load_payload() -> dict[str, Any]:
    raw = sys.stdin.read().strip()
    if not raw:
        return {}
    return json.loads(raw)


def normalize_text(value: Any) -> str:
    return str(value or "").strip()


def normalize_content_type(value: Any, default: str = "QI") -> str:
    normalized = normalize_text(value).upper()
    if normalized in {"QI", "QUOTE IMAGE"}:
        return "QI"
    if normalized in {"FP", "FULL POEM"}:
        return "FP"
    if normalized in {"FPI", "FULL POEM IMAGE", "FULL POEM IMAGE-BACKED", "FULL POEM IMAGE BACKED"}:
        return "FPI"
    return default


def resolve_content_type(record: dict[str, Any], default: str = "") -> str:
    content_type = normalize_content_type(
        record.get("contentType")
        or record.get("imageType")
        or record.get("content_type")
        or record.get("image_type"),
        "",
    )
    if content_type:
        return "FPI" if content_type == "FP" else content_type
    notes = str(record.get("productionNotes") or record.get("notes") or "")
    for line in notes.splitlines():
        if ":" not in line:
            continue
        label, value = line.split(":", 1)
        if normalize_text(label).lower() in {"content type", "image type", "type"}:
            resolved = normalize_content_type(value, default)
            return "FPI" if resolved == "FP" else resolved
    identity_values = [
        record.get("requestId"),
        record.get("graphicsRequestId"),
        record.get("sourceRecordId"),
        record.get("recordId"),
        record.get("canonicalPoemId"),
        record.get("poemId"),
    ]
    if any(re.search(r"(^|[-:])FPI?($|[-:])", normalize_text(value), re.IGNORECASE) for value in identity_values):
        return "FPI"
    if any(re.search(r"(^|[-:])QI($|[-:])", normalize_text(value), re.IGNORECASE) for value in identity_values):
        return "QI"
    return default


def sync_completions(connection, payload: dict[str, Any]) -> dict[str, Any]:
    completions = payload.get("completions") or []
    written = 0
    request_ids: list[str] = []
    repair_returns: list[dict[str, str]] = []

    for completion in completions:
        request_id = normalize_text(completion.get("requestId") or completion.get("graphicsRequestId"))
        if not request_id:
            continue
        content_type = resolve_content_type(completion)
        if not content_type:
            raise ValueError(f"contentType is required for completion {request_id}")

        request_record = {
            "id": request_id,
            "request_status": "COMPLETED_RETURNED",
            "source_type": "weaver_sheet_queue",
            "content_type": content_type,
            "image_type": content_type,
            "book_title": normalize_text(completion.get("bookTitle")),
            "poem_title": normalize_text(completion.get("poemTitle")),
            "author": normalize_text(completion.get("author")),
            "quote_text": str(completion.get("quoteText") or ""),
            "source_record_id": normalize_text(completion.get("sourceRecordId") or completion.get("recordId")),
            "source_sheet_name": "Queue - Needs Graphics",
            "source_sheet_row": completion.get("sourceSheetRow") or completion.get("sheetRow"),
            "source_payload": completion,
            "updated_at": normalize_text(completion.get("completedAt")) or utc_now_iso(),
        }
        upsert_graphics_request(connection, request_record)
        replace_graphics_request_items(connection, request_id, [{
            "source_record_id": request_record["source_record_id"],
            "source_sheet_name": request_record["source_sheet_name"],
            "source_sheet_row": request_record["source_sheet_row"],
            "book_title": request_record["book_title"],
            "poem_title": request_record["poem_title"],
            "author": request_record["author"],
            "quote_text": request_record["quote_text"],
            "created_at": normalize_text(completion.get("completedAt")) or utc_now_iso(),
        }])

        insert_graphics_completion(connection, {
            "id": normalize_text(completion.get("completionId") or completion.get("id")),
            "graphics_request_id": request_id,
            "source_tool": normalize_text(completion.get("sourceTool")) or "P.I.G.",
            "content_type": content_type,
            "image_type": content_type,
            "asset_url": normalize_text(completion.get("assetUrl") or completion.get("assetLinkUrl") or completion.get("driveUrl")),
            "asset_preview_url": normalize_text(completion.get("assetPreviewUrl") or completion.get("previewUrl") or completion.get("thumbnailUrl")),
            "production_notes": str(completion.get("productionNotes") or completion.get("notes") or ""),
            "completion_status": "RETURNED",
            "completed_at": normalize_text(completion.get("completedAt")) or utc_now_iso(),
            "source_payload": completion,
        })
        handoff = upsert_graphics_handoff_request(connection, {
            "graphicsRequestId": request_id,
            "sourceSystem": "weaver",
            "sourceStatus": "needs_graphics",
            "contentType": content_type,
            "imageType": content_type,
            "sourceCompletionId": normalize_text(completion.get("completionId") or completion.get("id")),
            "sourcePayload": request_record,
        })
        ledger_request_id = normalize_text(handoff.get("graphicsRequestId")) or request_id
        update_graphics_handoff(connection, ledger_request_id, {
            "contentType": content_type,
            "imageType": content_type,
            "sourceCompletionId": normalize_text(completion.get("completionId") or completion.get("id")),
            "pigStatus": "uploaded",
            "handoffStatus": "sent_to_weaver_qc",
            "qcStatus": "pending",
            "assetUrl": completion.get("assetUrl") or completion.get("assetLinkUrl") or completion.get("driveUrl"),
            "assetPreviewUrl": completion.get("assetPreviewUrl") or completion.get("previewUrl") or completion.get("thumbnailUrl"),
            "driveFileId": completion.get("driveFileId") or completion.get("fileId"),
            "driveFileName": completion.get("driveFileName") or completion.get("fileName"),
            "mimeType": completion.get("mimeType"),
            "exportType": completion.get("exportType"),
            "variant": completion.get("variant"),
            "version": completion.get("version"),
            "pigProjectId": completion.get("pigProjectId"),
            "editableProjectFileId": completion.get("editableProjectFileId") or completion.get("projectFileId"),
            "editableProjectUrl": completion.get("editableProjectUrl"),
            "editableProjectAvailable": bool(completion.get("editableProjectAvailable")),
            "candidatePigProjectId": completion.get("candidatePigProjectId"),
            "candidateEditableProjectFileId": completion.get("candidateEditableProjectFileId"),
            "candidateEditableProjectUrl": completion.get("candidateEditableProjectUrl"),
            "editableProjectValidationStatus": completion.get("editableProjectValidationStatus"),
            "editableProjectValidationError": completion.get("editableProjectValidationError") or {},
            "pigPayload": completion,
        })
        written += 1
        request_ids.append(request_id)
        completion_payload = completion.get("sourcePayload") if isinstance(completion.get("sourcePayload"), dict) else {}
        repair_request_id = normalize_text(
            completion.get("repairRequestId")
            or completion_payload.get("repairRequestId")
        )
        replacement_asset_link = normalize_text(
            completion.get("assetUrl")
            or completion.get("assetLinkUrl")
            or completion.get("driveUrl")
        )
        replacement_asset_id = normalize_text(
            completion.get("assetFileId")
            or completion.get("driveFileId")
            or completion.get("fileId")
            or completion.get("contentId")
            or completion.get("imageId")
        )
        if repair_request_id and replacement_asset_link:
            repair_update = update_repair_request(connection, {
                "repairRequestId": repair_request_id,
                "update": {
                    "weaverRepairStatus": "returned_pending_sync",
                    "replacementAssetId": replacement_asset_id,
                    "replacementAssetLink": replacement_asset_link,
                    "replacementPigCompletionId": normalize_text(
                        completion.get("completionId") or completion.get("id")
                    ),
                    "replacementAvailableAt": normalize_text(completion.get("completedAt")) or utc_now_iso(),
                    "retryable": True,
                    "historyEvent": "replacement_asset_received",
                    "historyNote": request_id,
                },
            })
            if repair_update.get("ok"):
                repair_returns.append({
                    "repairRequestId": repair_request_id,
                    "replacementAssetId": replacement_asset_id,
                    "replacementAssetLink": replacement_asset_link,
                    "pigJobId": request_id,
                    "pigCompletionId": normalize_text(completion.get("completionId") or completion.get("id")),
                })

    return {
        "ok": True,
        "written": written,
        "request_ids": request_ids,
        "repairReturns": repair_returns,
    }


def sync_qc_reviews(connection, payload: dict[str, Any]) -> dict[str, Any]:
    reviews = payload.get("reviews") or []
    written = 0
    skipped: list[dict[str, Any]] = []

    for review in reviews:
        storage_target = normalize_text(review.get("storageTarget")).lower()
        completion_id = normalize_text(review.get("pigCompletionId") or review.get("graphicsCompletionId"))
        request_id = normalize_text(review.get("graphicsRequestId"))
        if storage_target == "cleanup_sheet" and not completion_id and request_id:
            completion_id = f"cleanup:{request_id}"
        if storage_target not in {"firestore", "pig_sheet", "cleanup_sheet"} or not completion_id:
            skipped.append({
                "reason": "not_qc_queue_backed",
                "recordId": normalize_text(review.get("recordId")),
                "sheetRow": review.get("sheetRow"),
            })
            continue

        if not request_id:
            skipped.append({
                "reason": "missing_request_id",
                "recordId": normalize_text(review.get("recordId")),
                "sheetRow": review.get("sheetRow"),
                "completionId": completion_id,
            })
            continue
        content_type = resolve_content_type(review)
        if not content_type:
            raise ValueError(f"contentType is required for QC review {request_id}")

        if storage_target == "firestore":
            completion = connection.get_raw_document(
                FIRESTORE_GRAPHICS_COMPLETIONS_COLLECTION,
                completion_id,
            ) or {}
            completion_request_id = normalize_text(completion.get("graphicsRequestId"))
            if not completion_request_id:
                raise ValueError(f"Firestore completion {completion_id} was not found")
            if completion_request_id != request_id:
                raise ValueError(
                    f"Firestore completion {completion_id} belongs to {completion_request_id}, not {request_id}"
                )
            completion_payload = (
                completion.get("sourcePayload")
                if isinstance(completion.get("sourcePayload"), dict)
                else {}
            )
            pig_project_id = normalize_text(
                completion.get("pigProjectId") or completion_payload.get("pigProjectId")
            )
            editable_project_file_id = normalize_text(
                completion.get("editableProjectFileId")
                or completion_payload.get("editableProjectFileId")
                or completion_payload.get("projectFileId")
            )
            editable_project_url = normalize_text(
                completion.get("editableProjectUrl") or completion_payload.get("editableProjectUrl")
            )
            source_identity = completion.get("sourceIdentity") or completion_payload.get("sourceIdentity")
            if not isinstance(source_identity, dict):
                source_identity = {}
            content_id = normalize_text(
                completion.get("contentId")
                or completion.get("imageId")
                or completion_payload.get("contentId")
                or completion_payload.get("imageId")
                or completion_payload.get("sourceRecordId")
            )
            image_id = normalize_text(
                completion.get("imageId") or completion_payload.get("imageId") or content_id
            )
            text_hash = normalize_text(
                completion.get("textHash")
                or completion_payload.get("textHash")
                or source_identity.get("textHash")
            )
            editable_project_available = bool(
                pig_project_id and editable_project_file_id and editable_project_url
            )
            review = {
                **review,
                "pigProjectId": pig_project_id if editable_project_available else "",
                "editableProjectFileId": editable_project_file_id if editable_project_available else "",
                "editableProjectUrl": editable_project_url if editable_project_available else "",
                "editableProjectAvailable": editable_project_available,
                "contentId": content_id,
                "imageId": image_id,
                "sourceIdentity": source_identity,
                "textHash": text_hash,
            }

        if storage_target != "firestore":
            upsert_graphics_request(connection, {
                "id": request_id,
                "request_status": "COMPLETED_RETURNED",
                "source_type": "weaver_sheet_queue",
                "content_type": content_type,
                "image_type": content_type,
                "book_title": normalize_text(review.get("bookTitle")),
                "poem_title": normalize_text(review.get("poemTitle")),
                "author": normalize_text(review.get("author")),
                "quote_text": str(review.get("quoteText") or ""),
                "source_record_id": normalize_text(review.get("recordId")),
                "source_sheet_name": str(review.get("storageTarget") or ""),
                "source_sheet_row": review.get("sheetRow") or 0,
                "source_payload": review,
                "latest_completion_id": completion_id,
            })

            insert_graphics_completion(connection, {
                "id": completion_id,
                "graphics_request_id": request_id,
                "source_tool": normalize_text(review.get("sourceTool") or "P.I.G."),
                "content_type": content_type,
                "image_type": content_type,
                "asset_url": str(review.get("assetUrl") or ""),
                "asset_preview_url": str(review.get("assetPreviewUrl") or review.get("assetUrl") or ""),
                "production_notes": str(review.get("notes") or ""),
                "completion_status": "RETURNED",
                "completed_at": normalize_text(review.get("completedAt")) or utc_now_iso(),
                "source_payload": review,
            })

        insert_graphics_qc_review(connection, {
            "graphics_completion_id": completion_id,
            "decision": normalize_text(review.get("qcDecision")).upper(),
            "metadata_issue": normalize_text(review.get("metadataIssue")),
            "aesthetic_issue": normalize_text(review.get("aestheticIssue")),
            "note": str(review.get("qcNote") or ""),
            "reviewed_by": normalize_text(review.get("reviewedBy")),
            "reviewed_at": utc_now_iso(),
            "source_payload": review,
        })
        decision = normalize_text(review.get("qcDecision")).lower()
        if request_id:
            revision_reject = decision == "reject" and is_revision_reject(review)
            handoff = upsert_graphics_handoff_request(connection, {
                "graphicsRequestId": request_id,
                "sourceSystem": "weaver",
                "sourceStatus": "qc_review",
                "contentType": content_type,
                "imageType": content_type,
                "sourceCompletionId": completion_id,
                "sourcePayload": review,
                "handoffStatus": "sent_to_weaver_qc",
                "pigStatus": "uploaded",
                "qcStatus": "pending",
            })
            ledger_request_id = normalize_text(handoff.get("graphicsRequestId")) or request_id
            update_graphics_handoff(connection, ledger_request_id, {
                "contentType": content_type,
                "imageType": content_type,
                "sourceCompletionId": completion_id,
                "handoffStatus": "approved" if decision == "approve" else "rejected",
                "pigStatus": "not_started" if revision_reject else None,
                "qcStatus": "approved" if decision == "approve" else ("needs_revision" if revision_reject else "rejected"),
                "assetUrl": str(review.get("assetUrl") or review.get("assetLinkUrl") or ""),
                "assetPreviewUrl": str(review.get("assetPreviewUrl") or ""),
                "pigProjectId": str(review.get("pigProjectId") or ""),
                "editableProjectFileId": str(review.get("editableProjectFileId") or ""),
                "editableProjectUrl": str(review.get("editableProjectUrl") or ""),
                "editableProjectAvailable": bool(review.get("editableProjectAvailable")),
                "contentId": str(review.get("contentId") or ""),
                "imageId": str(review.get("imageId") or ""),
                "sourceIdentity": review.get("sourceIdentity") or {},
                "textHash": str(review.get("textHash") or ""),
                "qcPayload": review,
            })
        written += 1

    return {
        "ok": True,
        "written": written,
        "skipped": skipped,
    }


def upsert_graphics_qc_queue_cards(connection, payload: dict[str, Any]) -> dict[str, Any]:
    cards = payload.get("cards") or []
    written = 0
    for card in cards:
        request_id = normalize_text(card.get("graphicsRequestId"))
        completion_id = normalize_text(card.get("pigCompletionId"))
        if not completion_id and normalize_text(card.get("storageTarget")).lower() == "cleanup_sheet":
            completion_id = f"cleanup:{request_id}"
        if not completion_id:
            continue
        connection.write_raw_document(FIRESTORE_GRAPHICS_QC_QUEUE_COLLECTION, completion_id, {
            **card,
            "pigCompletionId": completion_id,
            "isPendingQc": True,
            "updatedAt": utc_now_iso(),
        })
        written += 1
    return {"ok": True, "written": written}


def is_revision_reject(review: dict[str, Any]) -> bool:
    reject_reason = normalize_text(
        review.get("rejectReason")
        or review.get("qcRejectReason")
        or review.get("revisionReason")
        or review.get("reworkReason")
        or review.get("qcStatus")
    ).lower()
    if reject_reason in {"needs_revision", "correct_and_recreate"}:
        return True
    qc_note = normalize_text(review.get("qcNote") or review.get("graphicsQcNote")).lower()
    return "reject reason: correct and recreate" in qc_note


def sync_poetry_please_handoffs(connection, payload: dict[str, Any]) -> dict[str, Any]:
    handoffs = payload.get("handoffs") or []
    written = 0

    for handoff in handoffs:
      completion_id = normalize_text(handoff.get("graphics_completion_id") or handoff.get("graphicsCompletionId"))
      if not completion_id:
          continue

      insert_poetry_please_handoff(connection, {
          "graphics_completion_id": completion_id,
          "handoff_status": normalize_text(handoff.get("handoff_status") or handoff.get("handoffStatus")),
          "handoff_mode": normalize_text(handoff.get("handoff_mode") or handoff.get("handoffMode")) or "auto",
          "handed_off_at": normalize_text(handoff.get("handed_off_at") or handoff.get("handedOffAt")) or utc_now_iso(),
          "poetry_please_item_id": normalize_text(handoff.get("poetry_please_item_id") or handoff.get("poetryPleaseItemId")),
          "payload": handoff.get("payload") or {},
      })
      written += 1

    return {
        "ok": True,
        "written": written,
    }


def sync_excerpt_handoffs(connection, payload: dict[str, Any]) -> dict[str, Any]:
    handoffs = payload.get("handoffs") or []
    if payload.get("handoff"):
        handoffs = [payload["handoff"]]
    records = [upsert_excerpt_handoff(connection, handoff) for handoff in handoffs]
    return {"ok": True, "records": records, "count": len(records)}


def fetch_excerpt_handoffs(connection, payload: dict[str, Any]) -> dict[str, Any]:
    record_ids = payload.get("recordIds") or []
    return {
        "ok": True,
        "records": get_excerpt_handoffs(connection, record_ids),
    }


def normalize_excerpt_record(record: dict[str, Any]) -> dict[str, Any]:
    source_record_id = normalize_text(record.get("sourceRecordId") or record.get("recordId"))
    if not source_record_id:
        raise ValueError("sourceRecordId is required")
    now = utc_now_iso()
    return {
        "recordId": source_record_id,
        "sourceKind": normalize_text(record.get("sourceKind")) or "weaver",
        "sourceRecordId": source_record_id,
        "sourceRow": int(record.get("sourceRow") or 0),
        "intakeMode": normalize_text(record.get("intakeMode")),
        "intakeLabel": normalize_text(record.get("intakeLabel")),
        "submittedBy": normalize_text(record.get("submittedBy") or record.get("email")),
        "author": normalize_text(record.get("author")),
        "poemTitle": normalize_text(record.get("poemTitle") or record.get("title")),
        "bookTitle": normalize_text(record.get("bookTitle")),
        "excerptText": str(record.get("excerptText") or record.get("quote") or ""),
        "sourceEvent": normalize_text(record.get("sourceEvent") or record.get("eventName")),
        "notes": str(record.get("notes") or ""),
        "contentType": normalize_text(record.get("contentType")) or "EXC",
        "releaseCatalog": normalize_text(record.get("releaseCatalog")),
        "bookShortener": normalize_text(record.get("bookShortener")),
        "socialMediaHandle": normalize_text(record.get("socialMediaHandle")),
        "reviewDecision": normalize_text(record.get("reviewDecision")).upper(),
        "approvedForUse": bool(record.get("approvedForUse")),
        "approvedForQuoteImage": bool(record.get("approvedForQuoteImage")),
        "approvedForGraphics": bool(record.get("approvedForGraphics")),
        "useForInt": bool(record.get("useForInt")),
        "excluded": bool(record.get("excluded")),
        "needsCorrection": bool(record.get("needsCorrection")),
        "correctionNote": str(record.get("correctionNote") or ""),
        "correctedAuthor": normalize_text(record.get("correctedAuthor")),
        "correctedPoemTitle": normalize_text(record.get("correctedPoemTitle") or record.get("correctedTitle")),
        "correctedBookTitle": normalize_text(record.get("correctedBookTitle")),
        "correctedExcerptText": str(record.get("correctedExcerptText") or record.get("correctedExcerpt") or ""),
        "duplicateGroupId": normalize_text(record.get("duplicateGroupId")),
        "validation": record.get("validation") if isinstance(record.get("validation"), dict) else {},
        "sourcePayload": record.get("sourcePayload") if isinstance(record.get("sourcePayload"), dict) else record,
        "createdAt": normalize_text(record.get("createdAt")) or now,
        "updatedAt": normalize_text(record.get("updatedAt")) or now,
    }


def upsert_excerpt_records(connection, payload: dict[str, Any]) -> dict[str, Any]:
    records = payload.get("records") or []
    if payload.get("record"):
        records = [payload["record"]]
    written = []
    for source in records:
        record = normalize_excerpt_record(source)
        existing = connection.get_raw_document(
            FIRESTORE_EXCERPT_RECORDS_COLLECTION,
            record["sourceRecordId"],
        ) or {}
        merged = {
            **existing,
            **record,
            "createdAt": normalize_text(existing.get("createdAt")) or record["createdAt"],
        }
        written.append(connection.write_raw_document(
            FIRESTORE_EXCERPT_RECORDS_COLLECTION,
            record["sourceRecordId"],
            merged,
        ))
    return {"ok": True, "records": written, "count": len(written)}


def fetch_excerpt_records(connection, payload: dict[str, Any]) -> dict[str, Any]:
    record_ids = {
        normalize_text(value)
        for value in (payload.get("recordIds") or [])
        if normalize_text(value)
    }
    records = connection.list_raw_documents(FIRESTORE_EXCERPT_RECORDS_COLLECTION)
    if record_ids:
        records = [record for record in records if normalize_text(record.get("sourceRecordId")) in record_ids]
    records.sort(key=lambda record: normalize_text(record.get("createdAt")))
    return {"ok": True, "records": records, "count": len(records)}


def fetch_graphics_state(connection, payload: dict[str, Any]) -> dict[str, Any]:
    completion_ids = [
        normalize_text(value)
        for value in (payload.get("completionIds") or [])
        if normalize_text(value)
    ]
    return {
        "ok": True,
        "qc_reviews": get_latest_graphics_qc_reviews(connection, completion_ids),
        "handoffs": get_latest_poetry_please_handoffs(connection, completion_ids),
    }


def upsert_handoff_requests(connection, payload: dict[str, Any]) -> dict[str, Any]:
    requests = payload.get("requests") or []
    if payload.get("request"):
        requests = [payload["request"]]
    records = [upsert_graphics_handoff_request(connection, request) for request in requests]
    return {"ok": True, "records": records, "count": len(records)}


def fetch_handoff_queue(connection, payload: dict[str, Any]) -> dict[str, Any]:
    limit = int(payload.get("limit") or 100)
    cursor = int(payload.get("cursor") or 0)
    records = get_graphics_handoff_queue(
        connection,
        limit + 1,
        normalize_text(payload.get("filter") or "all"),
        cursor,
        normalize_text(payload.get("contentType")),
    )
    has_more = len(records) > limit
    return {
        "ok": True,
        "contentType": normalize_text(payload.get("contentType")).upper(),
        "records": records[:limit],
        "nextCursor": str(cursor + limit) if has_more else "",
    }


def claim_handoff_request(connection, payload: dict[str, Any]) -> dict[str, Any]:
    graphics_request_id = normalize_text(payload.get("graphicsRequestId"))
    if not graphics_request_id:
        return {"ok": False, "error": "graphicsRequestId is required"}
    return {
        "ok": True,
        "record": claim_graphics_handoff(connection, graphics_request_id, normalize_text(payload.get("claimedBy"))),
    }


def patch_handoff_request(connection, payload: dict[str, Any]) -> dict[str, Any]:
    graphics_request_id = normalize_text(payload.get("graphicsRequestId"))
    if not graphics_request_id:
        return {"ok": False, "error": "graphicsRequestId is required"}
    update = payload.get("update") or {}
    return {
        "ok": True,
        "record": update_graphics_handoff(connection, graphics_request_id, update),
    }


def fetch_handoff_request(connection, payload: dict[str, Any]) -> dict[str, Any]:
    graphics_request_id = normalize_text(payload.get("graphicsRequestId"))
    if not graphics_request_id:
        return {"ok": False, "error": "graphicsRequestId is required"}
    record = get_graphics_handoff(connection, graphics_request_id)
    return {"ok": bool(record), "record": record, "error": "" if record else "not_found"}


def fetch_handoff_requests(connection, payload: dict[str, Any]) -> dict[str, Any]:
    request_ids = [
        normalize_text(value)
        for value in (payload.get("graphicsRequestIds") or payload.get("requestIds") or [])
        if normalize_text(value)
    ]
    records = get_graphics_handoffs(connection, request_ids)
    return {"ok": True, "records": records, "count": len(records)}


def repair_document_id(repair_request_id: str) -> str:
    return hashlib.sha256(repair_request_id.encode("utf-8")).hexdigest()


def repair_source_value(request: dict[str, Any], *keys: str) -> Any:
    original = request.get("originalContent") if isinstance(request.get("originalContent"), dict) else {}
    for key in keys:
        value = request.get(key)
        if value not in (None, "", [], {}):
            return value
        value = original.get(key)
        if value not in (None, "", [], {}):
            return value
    return ""


def append_repair_history(
    history: Any,
    status: str,
    event: str,
    note: str = "",
    at: str = "",
) -> list[dict[str, Any]]:
    entries = [entry for entry in (history or []) if isinstance(entry, dict)]
    next_entry = {
        "at": at or utc_now_iso(),
        "status": status,
        "event": event,
        "note": note,
    }
    if entries and all(entries[-1].get(key) == next_entry.get(key) for key in ("status", "event", "note")):
        return entries[-99:]
    entries.append(next_entry)
    return entries[-100:]


def build_repair_handoff_request(
    request: dict[str, Any],
    repair_request_id: str,
    pig_job_id: str,
) -> dict[str, Any]:
    content_type = normalize_text(request.get("contentType")).upper()
    original_content = request.get("originalContent") if isinstance(request.get("originalContent"), dict) else {}
    source_text = normalize_text(repair_source_value(
        request,
        "quoteText",
        "sourceText",
        "text",
        "excerpt",
        "fullText",
        "ocrText",
    ))
    original_asset_link = normalize_text(repair_source_value(
        request,
        "originalAssetLink",
        "assetUrl",
        "assetLinkUrl",
        "imageUrl",
        "driveLink",
        "previewUrl",
    ))
    issue_reason = normalize_text(request.get("issueReason"))
    request_note = normalize_text(request.get("requestNote"))
    return {
        "graphicsRequestId": pig_job_id,
        "sourceSystem": "poetry_please_repair",
        "sourceStatus": "repair_requested",
        "contentType": content_type,
        "imageType": content_type,
        "repairRequestId": repair_request_id,
        "handoffStatus": "requested",
        "pigStatus": "not_started",
        "qcStatus": "not_sent",
        "queueView": "rework",
        "reworkReason": issue_reason or "poetry_please_flag",
        "requestedChanges": request_note or issue_reason,
        "sourcePayload": {
            "repairRequestId": repair_request_id,
            "sourceFlagId": normalize_text(request.get("sourceFlagId")),
            "originalContentId": normalize_text(request.get("imageId")),
            "originalDocId": normalize_text(request.get("originalDocId")),
            "originalCollection": normalize_text(request.get("originalCollection")),
            "contentType": content_type,
            "imageType": content_type,
            "author": normalize_text(request.get("author")),
            "bookTitle": normalize_text(request.get("book")),
            "book": normalize_text(request.get("book")),
            "poemTitle": normalize_text(request.get("title")),
            "title": normalize_text(request.get("title")),
            "releaseCatalog": normalize_text(request.get("releaseCatalog")),
            "bookShortener": normalize_text(request.get("bookShortener")),
            "quoteText": source_text,
            "sourceText": source_text,
            "previousAssetUrl": original_asset_link,
            "assetUrl": original_asset_link,
            "issueReason": issue_reason,
            "requestNote": request_note,
            "requestedChanges": request_note or issue_reason,
            "originalContent": original_content,
            "existingPigHistory": original_content.get("pigHistory") or original_content.get("pig") or {},
            "existingWeaverHistory": original_content.get("weaverHistory") or original_content.get("weaver") or {},
            "queueView": "rework",
            "nextAction": "rework",
        },
    }


def upsert_repair_requests(connection, payload: dict[str, Any]) -> dict[str, Any]:
    if not hasattr(connection, "write_raw_document"):
        return {"ok": False, "error": "Repair request ingestion requires Firestore."}
    requests = payload.get("requests") or []
    created_count = 0
    updated_count = 0
    duplicate_count = 0
    blocked_count = 0
    error_count = 0
    accepted_for_status_sync: list[dict[str, Any]] = []
    records: list[dict[str, Any]] = []
    errors: list[dict[str, str]] = []

    for request in requests:
        repair_request_id = normalize_text(request.get("id"))
        if not repair_request_id:
            error_count += 1
            errors.append({"repairRequestId": "", "error": "missing_repair_request_id"})
            continue
        try:
            now = utc_now_iso()
            document_id = repair_document_id(repair_request_id)
            existing = connection.get_raw_document(FIRESTORE_REPAIR_REQUESTS_COLLECTION, document_id) or {}
            content_type = normalize_text(request.get("contentType")).upper()
            destination = "pig" if content_type in PIG_REPAIR_CONTENT_TYPES else "weaver_review"
            original_content = request.get("originalContent") if isinstance(request.get("originalContent"), dict) else {}
            original_asset_link = normalize_text(repair_source_value(
                request,
                "originalAssetLink",
                "assetUrl",
                "assetLinkUrl",
                "imageUrl",
                "driveLink",
                "previewUrl",
            ))
            source_text = normalize_text(repair_source_value(
                request,
                "quoteText",
                "sourceText",
                "text",
                "excerpt",
                "fullText",
                "ocrText",
            ))
            missing_fields = [
                label for label, value in (
                    ("imageId", request.get("imageId")),
                    ("contentType", content_type),
                    ("author", request.get("author")),
                    ("book", request.get("book")),
                    ("title", request.get("title")),
                )
                if not normalize_text(value)
            ]
            if destination == "pig" and not (source_text or original_asset_link):
                missing_fields.append("sourceTextOrAsset")
            blocked_reason = (
                f"missing_required_metadata:{','.join(missing_fields)}"
                if missing_fields
                else ""
            )
            pig_job_id = normalize_text(existing.get("pigJobId"))
            handed_off_at = normalize_text(existing.get("handedOffAt"))
            if destination == "pig" and not blocked_reason and not pig_job_id:
                pig_job_id = f"weaver:repair:{hashlib.sha256(repair_request_id.encode('utf-8')).hexdigest()[:32]}"
                handoff = upsert_graphics_handoff_request(
                    connection,
                    build_repair_handoff_request(request, repair_request_id, pig_job_id),
                )
                pig_job_id = normalize_text(handoff.get("graphicsRequestId")) or pig_job_id
                handed_off_at = now

            if blocked_reason:
                repair_status = "blocked"
            elif destination == "pig":
                repair_status = "waiting_on_pig"
            else:
                repair_status = "in_progress"

            terminal_status = normalize_text(existing.get("weaverRepairStatus")).lower()
            if terminal_status in {"returned", "resolved"}:
                repair_status = terminal_status
            source_snapshot = json.dumps(request, sort_keys=True, separators=(",", ":"))
            unchanged = bool(
                existing
                and existing.get("sourceSnapshot") == source_snapshot
                and normalize_text(existing.get("pigJobId")) == pig_job_id
                and normalize_text(existing.get("blockedReason")) == blocked_reason
            )
            if unchanged:
                duplicate_count += 1
                records.append(existing)
                if (
                    not blocked_reason
                    and normalize_text(existing.get("poetryPleaseStatus")).lower() == "requested"
                    and normalize_text(existing.get("weaverRepairStatus")).lower() not in {"returned", "resolved"}
                ):
                    accepted_for_status_sync.append({
                        "repairRequestId": repair_request_id,
                        "destination": destination,
                        "pigJobId": pig_job_id,
                    })
                continue

            event = "repair_imported" if not existing else "repair_metadata_updated"
            record = {
                **existing,
                "id": document_id,
                "poetryPleaseRepairRequestId": repair_request_id,
                "sourceFlagId": normalize_text(request.get("sourceFlagId")),
                "originalContentId": normalize_text(request.get("imageId")),
                "originalDocId": normalize_text(request.get("originalDocId")),
                "originalCollection": normalize_text(request.get("originalCollection")),
                "contentType": content_type,
                "author": normalize_text(request.get("author")),
                "canonicalBook": normalize_text(request.get("book")),
                "poemTitle": normalize_text(request.get("title")),
                "releaseCatalog": normalize_text(request.get("releaseCatalog")),
                "bookShortener": normalize_text(request.get("bookShortener")),
                "issueReason": normalize_text(request.get("issueReason")),
                "repairInstructions": normalize_text(request.get("requestNote")),
                "originalAssetLink": original_asset_link,
                "originalAssetLinks": {
                    key: value for key, value in {
                        "assetUrl": normalize_text(original_content.get("assetUrl")),
                        "assetLinkUrl": normalize_text(original_content.get("assetLinkUrl")),
                        "imageUrl": normalize_text(original_content.get("imageUrl")),
                        "driveLink": normalize_text(original_content.get("driveLink")),
                    }.items() if value
                },
                "originalMetadata": original_content,
                "sourceText": source_text,
                "sourceRequest": request,
                "sourceCreatedAt": request.get("createdAt") or "",
                "sourceSnapshot": source_snapshot,
                "poetryPleaseStatus": normalize_text(existing.get("poetryPleaseStatus") or request.get("status") or "requested"),
                "weaverRepairStatus": repair_status,
                "repairDestination": destination,
                "repairDestinationReason": (
                    "image_backed_content_type"
                    if destination == "pig"
                    else f"unsupported_graphics_content_type:{content_type or 'missing'}"
                ),
                "pigJobId": pig_job_id,
                "blockedReason": blocked_reason,
                "retryable": bool(blocked_reason),
                "createdAt": existing.get("createdAt") or now,
                "updatedAt": now,
                "handedOffAt": handed_off_at,
                "returnedAt": existing.get("returnedAt") or "",
                "resolvedAt": existing.get("resolvedAt") or "",
                "replacementAssetId": existing.get("replacementAssetId") or "",
                "replacementAssetLink": existing.get("replacementAssetLink") or "",
                "latestPoetryPleaseResponse": existing.get("latestPoetryPleaseResponse") or {
                    "ok": True,
                    "operation": "fetch_requested_repairs",
                    "at": now,
                },
                "statusHistory": append_repair_history(
                    existing.get("statusHistory"),
                    repair_status,
                    event,
                    blocked_reason or destination,
                    now,
                ),
            }
            written = connection.write_raw_document(
                FIRESTORE_REPAIR_REQUESTS_COLLECTION,
                document_id,
                record,
            )
            records.append(written)
            if existing:
                updated_count += 1
            else:
                created_count += 1
            if blocked_reason:
                blocked_count += 1
            else:
                accepted_for_status_sync.append({
                    "repairRequestId": repair_request_id,
                    "destination": destination,
                    "pigJobId": pig_job_id,
                })
        except Exception as error:
            error_count += 1
            errors.append({"repairRequestId": repair_request_id, "error": str(error)})

    return {
        "ok": error_count == 0,
        "createdCount": created_count,
        "updatedCount": updated_count,
        "duplicateCount": duplicate_count,
        "blockedCount": blocked_count,
        "errorCount": error_count,
        "acceptedForStatusSync": accepted_for_status_sync,
        "records": records,
        "errors": errors,
    }


def list_repair_requests(connection, payload: dict[str, Any]) -> dict[str, Any]:
    records = connection.list_raw_documents(FIRESTORE_REPAIR_REQUESTS_COLLECTION)
    status_filter = {
        normalize_text(value).lower()
        for value in (payload.get("statuses") or [])
        if normalize_text(value)
    }
    if status_filter:
        records = [
            record for record in records
            if normalize_text(record.get("weaverRepairStatus")).lower() in status_filter
        ]
    records.sort(key=lambda record: normalize_text(record.get("updatedAt")), reverse=True)
    limit = max(1, min(int(payload.get("limit") or 200), 500))
    return {"ok": True, "records": records[:limit], "count": min(len(records), limit)}


def update_repair_request(connection, payload: dict[str, Any]) -> dict[str, Any]:
    repair_request_id = normalize_text(payload.get("repairRequestId"))
    if not repair_request_id:
        return {"ok": False, "error": "repairRequestId is required"}
    document_id = repair_document_id(repair_request_id)
    existing = connection.get_raw_document(FIRESTORE_REPAIR_REQUESTS_COLLECTION, document_id)
    if not existing:
        return {"ok": False, "error": "repair_request_not_found"}
    now = utc_now_iso()
    update = payload.get("update") if isinstance(payload.get("update"), dict) else {}
    next_status = normalize_text(update.get("weaverRepairStatus") or existing.get("weaverRepairStatus"))
    event = normalize_text(update.get("historyEvent") or "repair_status_updated")
    note = normalize_text(update.get("historyNote"))
    record = {
        **existing,
        **update,
        "id": document_id,
        "poetryPleaseRepairRequestId": repair_request_id,
        "updatedAt": now,
        "statusHistory": append_repair_history(
            existing.get("statusHistory"),
            next_status,
            event,
            note,
            now,
        ),
    }
    written = connection.write_raw_document(FIRESTORE_REPAIR_REQUESTS_COLLECTION, document_id, record)
    return {"ok": True, "record": written}


def recover_stalled_pig_requests(connection, payload: dict[str, Any]) -> dict[str, Any]:
    if not hasattr(connection, "list_handoffs") or not hasattr(connection, "list_raw_documents"):
        return {"ok": False, "error": "Stalled P.I.G. recovery requires Firestore."}

    start_at = normalize_text(payload.get("startAt"))
    end_before = normalize_text(payload.get("endBefore"))
    apply_changes = bool(payload.get("apply"))
    requested_ids = {
        normalize_text(value)
        for value in (payload.get("graphicsRequestIds") or [])
        if normalize_text(value)
    }
    if not start_at or not end_before:
        return {"ok": False, "error": "startAt and endBefore are required"}

    completion_request_ids = {
        normalize_text(record.get("graphicsRequestId"))
        for record in connection.list_raw_documents(FIRESTORE_GRAPHICS_COMPLETIONS_COLLECTION)
        if normalize_text(record.get("graphicsRequestId"))
    }
    qc_request_ids = {
        normalize_text(record.get("graphicsRequestId"))
        for record in connection.list_raw_documents(FIRESTORE_GRAPHICS_QC_QUEUE_COLLECTION)
        if normalize_text(record.get("graphicsRequestId"))
    }

    candidates: list[dict[str, Any]] = []
    protected_qc_count = 0
    protected_completion_count = 0
    inspected = 0
    for record in connection.list_handoffs():
        request_id = normalize_text(record.get("graphicsRequestId"))
        if not request_id or (requested_ids and request_id not in requested_ids):
            continue
        inspected += 1
        activity_at = normalize_text(record.get("claimedAt") or record.get("updatedAt"))
        if not activity_at or activity_at < start_at or activity_at >= end_before:
            continue
        if request_id in qc_request_ids:
            protected_qc_count += 1
            continue
        if request_id in completion_request_ids:
            protected_completion_count += 1
            continue

        handoff_status = normalize_text(record.get("handoffStatus")).lower()
        pig_status = normalize_text(record.get("pigStatus")).lower()
        qc_status = normalize_text(record.get("qcStatus")).lower()
        if qc_status != "not_sent":
            continue
        if handoff_status in {"approved", "rejected", "sent_to_weaver_qc", "blocked"}:
            continue
        if record.get("isActionable"):
            continue
        if (
            handoff_status not in {"claimed", "generated", "exported", "uploaded", "errored", "requested"}
            and pig_status not in {"claimed", "generating", "generated", "exported", "uploaded", "failed"}
        ):
            continue

        source_payload = record.get("sourcePayload") if isinstance(record.get("sourcePayload"), dict) else {}
        candidates.append({
            "graphicsRequestId": request_id,
            "contentType": normalize_text(record.get("contentType") or record.get("imageType")).upper(),
            "author": normalize_text(record.get("author") or source_payload.get("author")),
            "poemTitle": normalize_text(record.get("poemTitle") or source_payload.get("poemTitle")),
            "bookTitle": normalize_text(record.get("bookTitle") or source_payload.get("bookTitle")),
            "activityAt": activity_at,
            "previousHandoffStatus": handoff_status,
            "previousPigStatus": pig_status,
            "previousQcStatus": qc_status,
        })

    repaired: list[dict[str, Any]] = []
    if apply_changes:
        now = utc_now_iso()
        for candidate in candidates:
            request_id = candidate["graphicsRequestId"]
            record = connection.get_handoff(request_id)
            if not record:
                continue
            transition_log = [
                entry for entry in (record.get("transitionLog") or [])
                if isinstance(entry, dict)
            ]
            transition_log.append({
                "at": now,
                "event": "stalled_completion_reset",
                "reason": "completion_failed_before_graphics_qc",
                "previousHandoffStatus": candidate["previousHandoffStatus"],
                "previousPigStatus": candidate["previousPigStatus"],
            })
            record.update({
                "handoffStatus": "requested",
                "pigStatus": "not_started",
                "qcStatus": "not_sent",
                "queueView": (
                    record.get("queueView")
                    if record.get("queueView") in {"current_titles", "coverage_needs"}
                    else "current_titles"
                ),
                "sourceCompletionId": "",
                "claimedBy": "",
                "claimedAt": "",
                "assetUrl": "",
                "assetPreviewUrl": "",
                "driveFileId": "",
                "driveFileName": "",
                "mimeType": "",
                "exportType": "",
                "variant": "",
                "pigProjectId": "",
                "editableProjectFileId": "",
                "editableProjectUrl": "",
                "editableProjectAvailable": False,
                "candidatePigProjectId": "",
                "candidateEditableProjectFileId": "",
                "candidateEditableProjectUrl": "",
                "editableProjectValidationStatus": "",
                "editableProjectValidationError": {},
                "errorMessage": "",
                "blockedReason": "",
                "generatedAt": "",
                "uploadedAt": "",
                "sentToQcAt": "",
                "updatedAt": now,
                "transitionLog": transition_log,
            })
            repaired.append(connection.write_record(record))

    return {
        "ok": True,
        "dryRun": not apply_changes,
        "startAt": start_at,
        "endBefore": end_before,
        "inspectedCount": inspected,
        "eligibleCount": len(candidates),
        "repairedCount": len(repaired),
        "protectedCompletionCount": protected_completion_count,
        "protectedQcCount": protected_qc_count,
        "candidates": candidates,
        "repairedGraphicsRequestIds": [
            normalize_text(record.get("graphicsRequestId"))
            for record in repaired
        ],
    }


def main() -> int:
    payload = load_payload()
    action = normalize_text(payload.get("action"))
    validate_production_handoff_backend(action)
    use_firestore = (
        os.environ.get("WEAVER_LEDGER_BACKEND", "").strip().lower() == "firestore"
        and action in FIRESTORE_HANDOFF_ACTIONS
    )
    connection = connect_firestore_ledger() if use_firestore else connect_runtime_db()
    try:
        if not use_firestore:
            ensure_runtime_schema(connection)
        if action == "upsert_completions":
            result = sync_completions(connection, payload)
        elif action == "insert_qc_reviews":
            result = sync_qc_reviews(connection, payload)
        elif action == "insert_poetry_please_handoffs":
            result = sync_poetry_please_handoffs(connection, payload)
        elif action == "upsert_excerpt_handoffs":
            result = sync_excerpt_handoffs(connection, payload)
        elif action == "get_excerpt_handoffs":
            result = fetch_excerpt_handoffs(connection, payload)
        elif action == "get_graphics_state":
            result = fetch_graphics_state(connection, payload)
        elif action == "get_pending_graphics_qc":
            records = get_pending_graphics_qc_records(
                connection,
                include_cleanup=bool(payload.get("includeCleanup")),
            )
            result = {"ok": True, "records": records, "count": len(records)}
        elif action == "rebuild_graphics_qc_queue":
            result = rebuild_graphics_qc_queue_cards(connection)
        elif action == "upsert_graphics_qc_queue_cards":
            result = upsert_graphics_qc_queue_cards(connection, payload)
        elif action == "upsert_handoff_requests":
            result = upsert_handoff_requests(connection, payload)
        elif action == "get_handoff_queue":
            result = fetch_handoff_queue(connection, payload)
        elif action == "claim_handoff_request":
            result = claim_handoff_request(connection, payload)
        elif action == "patch_handoff_request":
            result = patch_handoff_request(connection, payload)
        elif action == "get_handoff_request":
            result = fetch_handoff_request(connection, payload)
        elif action == "get_handoff_requests":
            result = fetch_handoff_requests(connection, payload)
        elif action == "get_graphics_completion":
            completion_id = normalize_text(payload.get("completionId") or payload.get("pigCompletionId"))
            record = connection.get_raw_document(
                FIRESTORE_GRAPHICS_COMPLETIONS_COLLECTION,
                completion_id,
            ) if completion_id else None
            result = {"ok": True, "record": record or {}}
        elif action == "audit_editable_project_links":
            records = connection.list_raw_documents(FIRESTORE_GRAPHICS_COMPLETIONS_COLLECTION)
            conflicts = find_editable_project_link_conflicts(records)
            result = {
                "ok": True,
                "completionCount": len(records),
                "conflictCount": len(conflicts),
                "conflicts": conflicts,
            }
        elif action == "recover_stalled_pig_requests":
            result = recover_stalled_pig_requests(connection, payload)
        elif action == "upsert_repair_requests":
            result = upsert_repair_requests(connection, payload)
        elif action == "list_repair_requests":
            result = list_repair_requests(connection, payload)
        elif action == "update_repair_request":
            result = update_repair_request(connection, payload)
        elif action == "upsert_excerpt_records":
            result = upsert_excerpt_records(connection, payload)
        elif action == "get_excerpt_records":
            result = fetch_excerpt_records(connection, payload)
        else:
            result = {
                "ok": False,
                "error": f"Unsupported action: {action or '(blank)'}",
            }
    finally:
        connection.close()

    print(json.dumps(result))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
