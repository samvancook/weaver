#!/usr/bin/env python3
from __future__ import annotations

import json
import os
import sys
from typing import Any

from weaver_runtime_db import (
    connect_firestore_ledger,
    connect_runtime_db,
    ensure_runtime_schema,
    claim_graphics_handoff,
    get_excerpt_handoffs,
    get_graphics_handoff,
    get_graphics_handoffs,
    get_graphics_handoff_queue,
    get_latest_graphics_qc_reviews,
    get_latest_poetry_please_handoffs,
    get_pending_graphics_qc_records,
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
    "upsert_handoff_requests",
    "get_handoff_queue",
    "claim_handoff_request",
    "patch_handoff_request",
    "get_handoff_request",
    "get_handoff_requests",
}

EXPECTED_FIRESTORE_PROJECT_ID = "button-weaver-internal"
EXPECTED_FIRESTORE_DATABASE_ID = "weaverledger"


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


def resolve_content_type(record: dict[str, Any], default: str = "QI") -> str:
    content_type = normalize_content_type(
        record.get("contentType")
        or record.get("imageType")
        or record.get("content_type")
        or record.get("image_type"),
        "",
    )
    if content_type:
        return content_type
    notes = str(record.get("productionNotes") or record.get("notes") or "")
    for line in notes.splitlines():
        if ":" not in line:
            continue
        label, value = line.split(":", 1)
        if normalize_text(label).lower() in {"content type", "image type", "type"}:
            return normalize_content_type(value, default)
    return default


def sync_completions(connection, payload: dict[str, Any]) -> dict[str, Any]:
    completions = payload.get("completions") or []
    written = 0
    request_ids: list[str] = []

    for completion in completions:
        request_id = normalize_text(completion.get("requestId") or completion.get("graphicsRequestId"))
        if not request_id:
            continue
        content_type = resolve_content_type(completion)

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
        upsert_graphics_handoff_request(connection, {
            "graphicsRequestId": request_id,
            "sourceSystem": "weaver",
            "sourceStatus": "needs_graphics",
            "contentType": content_type,
            "imageType": content_type,
            "sourceCompletionId": normalize_text(completion.get("completionId") or completion.get("id")),
            "sourcePayload": request_record,
        })
        update_graphics_handoff(connection, request_id, {
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
            "pigPayload": completion,
        })
        written += 1
        request_ids.append(request_id)

    return {
        "ok": True,
        "written": written,
        "request_ids": request_ids,
    }


def sync_qc_reviews(connection, payload: dict[str, Any]) -> dict[str, Any]:
    reviews = payload.get("reviews") or []
    written = 0
    skipped: list[dict[str, Any]] = []

    for review in reviews:
        storage_target = normalize_text(review.get("storageTarget")).lower()
        completion_id = normalize_text(review.get("pigCompletionId") or review.get("graphicsCompletionId"))
        request_id = normalize_text(review.get("graphicsRequestId"))
        if storage_target != "pig_sheet" or not completion_id:
            skipped.append({
                "reason": "not_pig_backed",
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
            upsert_graphics_handoff_request(connection, {
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
            update_graphics_handoff(connection, request_id, {
                "contentType": content_type,
                "imageType": content_type,
                "sourceCompletionId": completion_id,
                "handoffStatus": "approved" if decision == "approve" else "rejected",
                "pigStatus": "not_started" if revision_reject else None,
                "qcStatus": "approved" if decision == "approve" else ("needs_revision" if revision_reject else "rejected"),
                "qcPayload": review,
            })
        written += 1

    return {
        "ok": True,
        "written": written,
        "skipped": skipped,
    }


def is_revision_reject(review: dict[str, Any]) -> bool:
    reject_reason = normalize_text(
        review.get("rejectReason")
        or review.get("qcRejectReason")
        or review.get("revisionReason")
        or review.get("qcStatus")
    ).lower()
    return reject_reason in {"needs_revision", "correct_and_recreate"}


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
    )
    has_more = len(records) > limit
    return {
        "ok": True,
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
            records = get_pending_graphics_qc_records(connection)
            result = {"ok": True, "records": records, "count": len(records)}
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
