#!/usr/bin/env python3
from __future__ import annotations

import json
import sqlite3
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
DEFAULT_DB_PATH = ROOT / "data" / "weaver_runtime.db"
SCHEMA_PATH = ROOT / "db" / "weaver_runtime_schema.sql"


def utc_now_iso() -> str:
    return datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def normalize_key(value: str) -> str:
    return " ".join((value or "").strip().lower().split())


def count_words(text: str) -> int:
    return len([part for part in (text or "").split() if part.strip()])


def connect_runtime_db(db_path: Path = DEFAULT_DB_PATH) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def ensure_runtime_schema(connection: sqlite3.Connection, schema_path: Path = SCHEMA_PATH) -> None:
    connection.executescript(schema_path.read_text())
    connection.commit()


def get_runtime_db_summary(connection: sqlite3.Connection) -> dict[str, Any]:
    tables = [
        "graphics_requests",
        "graphics_request_items",
        "graphics_completions",
        "graphics_handoff_ledger",
        "graphics_qc_reviews",
        "poetry_please_handoffs",
    ]
    counts = {
        table: int(connection.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0])
        for table in tables
    }
    return {
        "counts": counts,
        "views": {
            "qc_approved_graphics": int(connection.execute("SELECT COUNT(*) FROM v_qc_approved_graphics").fetchone()[0]),
            "latest_qc_reviews": int(connection.execute("SELECT COUNT(*) FROM v_latest_graphics_qc_reviews").fetchone()[0]),
        },
    }


def upsert_graphics_request(connection: sqlite3.Connection, request: dict[str, Any]) -> str:
    request_id = str(request["id"]).strip()
    created_at = request.get("created_at") or utc_now_iso()
    updated_at = request.get("updated_at") or created_at
    quote_text = str(request.get("quote_text") or "")
    payload_json = json.dumps(request.get("source_payload") or {}, ensure_ascii=True, sort_keys=True)

    connection.execute(
        """
        INSERT INTO graphics_requests (
            id, request_status, source_type, book_title, poem_title, author, quote_text,
            normalized_book_key, normalized_poem_key, normalized_author_key, normalized_quote_key,
            word_count, source_record_id, source_sheet_name, source_sheet_row,
            source_payload_json, created_at, updated_at, latest_completion_id
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
            request_status = excluded.request_status,
            source_type = excluded.source_type,
            book_title = excluded.book_title,
            poem_title = excluded.poem_title,
            author = excluded.author,
            quote_text = excluded.quote_text,
            normalized_book_key = excluded.normalized_book_key,
            normalized_poem_key = excluded.normalized_poem_key,
            normalized_author_key = excluded.normalized_author_key,
            normalized_quote_key = excluded.normalized_quote_key,
            word_count = excluded.word_count,
            source_record_id = excluded.source_record_id,
            source_sheet_name = excluded.source_sheet_name,
            source_sheet_row = excluded.source_sheet_row,
            source_payload_json = excluded.source_payload_json,
            updated_at = excluded.updated_at,
            latest_completion_id = COALESCE(excluded.latest_completion_id, graphics_requests.latest_completion_id)
        """,
        (
            request_id,
            str(request.get("request_status") or "OPEN"),
            str(request.get("source_type") or "weaver_sheet_queue"),
            str(request.get("book_title") or ""),
            str(request.get("poem_title") or ""),
            str(request.get("author") or ""),
            quote_text,
            normalize_key(str(request.get("book_title") or "")),
            normalize_key(str(request.get("poem_title") or "")),
            normalize_key(str(request.get("author") or "")),
            normalize_key(quote_text),
            int(request.get("word_count") or count_words(quote_text)),
            str(request.get("source_record_id") or ""),
            str(request.get("source_sheet_name") or ""),
            int(request.get("source_sheet_row") or 0) or None,
            payload_json,
            created_at,
            updated_at,
            str(request.get("latest_completion_id") or ""),
        ),
    )
    connection.commit()
    return request_id


HANDOFF_STATUSES = {
    "requested",
    "claimed",
    "generated",
    "exported",
    "uploaded",
    "sent_to_weaver_qc",
    "approved",
    "rejected",
    "blocked",
    "errored",
}
PIG_STATUSES = {
    "not_started",
    "claimed",
    "generating",
    "generated",
    "exported",
    "uploaded",
    "failed",
}
QC_STATUSES = {
    "not_sent",
    "pending",
    "approved",
    "rejected",
    "needs_revision",
}


def normalize_enum(value: Any, allowed: set[str], default: str) -> str:
    normalized = str(value or "").strip().lower()
    return normalized if normalized in allowed else default


def extract_handoff_text(payload: dict[str, Any]) -> str:
    for key in ("quoteText", "quote_text", "sourceText", "source_text", "text", "excerpt", "correctedExcerpt"):
        value = str(payload.get(key) or "").strip()
        if value:
            return value
    return ""


def row_to_handoff(row: sqlite3.Row) -> dict[str, Any]:
    source_payload = json.loads(row["source_payload_json"] or "{}")
    quote_text = extract_handoff_text(source_payload)
    return {
        "graphicsRequestId": str(row["graphics_request_id"] or ""),
        "sourceSystem": str(row["source_system"] or ""),
        "sourceStatus": str(row["source_status"] or ""),
        "pigStatus": str(row["pig_status"] or ""),
        "handoffStatus": str(row["handoff_status"] or ""),
        "qcStatus": str(row["qc_status"] or ""),
        "assetUrl": str(row["asset_url"] or ""),
        "assetPreviewUrl": str(row["asset_preview_url"] or ""),
        "driveFileId": str(row["drive_file_id"] or ""),
        "driveFileName": str(row["drive_file_name"] or ""),
        "mimeType": str(row["mime_type"] or ""),
        "exportType": str(row["export_type"] or ""),
        "variant": str(row["variant"] or ""),
        "version": str(row["version"] or ""),
        "claimedBy": str(row["claimed_by"] or ""),
        "errorMessage": str(row["error_message"] or ""),
        "blockedReason": str(row["blocked_reason"] or ""),
        "sourceSheetRow": source_payload.get("sourceSheetRow") or source_payload.get("source_sheet_row") or source_payload.get("sheetRow") or "",
        "queueSheetRow": source_payload.get("queueSheetRow") or source_payload.get("sourceSheetRow") or source_payload.get("sheetRow") or "",
        "author": str(source_payload.get("author") or ""),
        "poemTitle": str(source_payload.get("poemTitle") or source_payload.get("title") or ""),
        "bookTitle": str(source_payload.get("bookTitle") or ""),
        "quoteText": quote_text,
        "text": quote_text,
        "sourcePayload": source_payload,
        "pigPayload": json.loads(row["pig_payload_json"] or "{}"),
        "qcPayload": json.loads(row["qc_payload_json"] or "{}"),
        "transitionLog": json.loads(row["transition_log_json"] or "[]"),
        "createdAt": str(row["created_at"] or ""),
        "updatedAt": str(row["updated_at"] or ""),
        "claimedAt": str(row["claimed_at"] or ""),
        "generatedAt": str(row["generated_at"] or ""),
        "uploadedAt": str(row["uploaded_at"] or ""),
        "sentToQcAt": str(row["sent_to_qc_at"] or ""),
        "approvedAt": str(row["approved_at"] or ""),
        "rejectedAt": str(row["rejected_at"] or ""),
    }


def append_transition_log(existing_json: str, event: dict[str, Any]) -> str:
    try:
        entries = json.loads(existing_json or "[]")
        if not isinstance(entries, list):
            entries = []
    except json.JSONDecodeError:
        entries = []
    entries.append(event)
    return json.dumps(entries[-50:], ensure_ascii=True, sort_keys=True)


def get_graphics_handoff(connection: sqlite3.Connection, graphics_request_id: str) -> dict[str, Any] | None:
    row = connection.execute(
        "SELECT * FROM graphics_handoff_ledger WHERE graphics_request_id = ?",
        (graphics_request_id,),
    ).fetchone()
    return row_to_handoff(row) if row else None


def get_graphics_handoffs(connection: sqlite3.Connection, graphics_request_ids: list[str]) -> list[dict[str, Any]]:
    ids = [str(value or "").strip() for value in graphics_request_ids if str(value or "").strip()]
    if not ids:
        return []
    placeholders = ",".join("?" for _ in ids)
    rows = connection.execute(
        f"SELECT * FROM graphics_handoff_ledger WHERE graphics_request_id IN ({placeholders})",
        ids,
    ).fetchall()
    return [row_to_handoff(row) for row in rows]


def upsert_graphics_handoff_request(connection: sqlite3.Connection, request: dict[str, Any]) -> dict[str, Any]:
    graphics_request_id = str(request.get("graphicsRequestId") or request.get("id") or "").strip()
    if not graphics_request_id:
        raise ValueError("graphicsRequestId is required")

    now = utc_now_iso()
    existing = connection.execute(
        "SELECT * FROM graphics_handoff_ledger WHERE graphics_request_id = ?",
        (graphics_request_id,),
    ).fetchone()
    source_payload = request.get("sourcePayload") or request.get("payload") or request
    has_text = bool(extract_handoff_text(source_payload))
    handoff_status = normalize_enum(request.get("handoffStatus"), HANDOFF_STATUSES, "requested")
    pig_status = normalize_enum(request.get("pigStatus"), PIG_STATUSES, "not_started")
    qc_status = normalize_enum(request.get("qcStatus"), QC_STATUSES, "not_sent")
    blocked_reason = str(request.get("blockedReason") or "")
    if not has_text and handoff_status in {"requested", "claimed"}:
        handoff_status = "blocked"
        pig_status = "failed"
        blocked_reason = blocked_reason or "blank_request_text"
    log_json = append_transition_log(
        existing["transition_log_json"] if existing else "[]",
        {
            "at": now,
            "event": "request_upserted",
            "handoffStatus": handoff_status,
            "pigStatus": pig_status,
            "qcStatus": qc_status,
            "blockedReason": blocked_reason,
        },
    )

    connection.execute(
        """
        INSERT INTO graphics_handoff_ledger (
            graphics_request_id, source_system, source_status, pig_status, handoff_status, qc_status,
            source_payload_json, blocked_reason, transition_log_json, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(graphics_request_id) DO UPDATE SET
            source_system = excluded.source_system,
            source_status = excluded.source_status,
            pig_status = excluded.pig_status,
            handoff_status = excluded.handoff_status,
            qc_status = excluded.qc_status,
            source_payload_json = excluded.source_payload_json,
            blocked_reason = excluded.blocked_reason,
            transition_log_json = excluded.transition_log_json,
            updated_at = excluded.updated_at
        """,
        (
            graphics_request_id,
            str(request.get("sourceSystem") or "weaver"),
            str(request.get("sourceStatus") or "needs_graphics"),
            pig_status,
            handoff_status,
            qc_status,
            json.dumps(source_payload, ensure_ascii=True, sort_keys=True),
            blocked_reason,
            log_json,
            existing["created_at"] if existing else now,
            now,
        ),
    )
    connection.commit()
    return get_graphics_handoff(connection, graphics_request_id) or {}


def claim_graphics_handoff(connection: sqlite3.Connection, graphics_request_id: str, claimed_by: str = "") -> dict[str, Any]:
    existing = get_graphics_handoff(connection, graphics_request_id)
    if not existing:
        raise KeyError(f"Unknown graphicsRequestId: {graphics_request_id}")
    if existing["handoffStatus"] in {"generated", "exported", "uploaded", "sent_to_weaver_qc", "approved", "blocked", "errored"}:
        return existing

    now = utc_now_iso()
    row = connection.execute(
        "SELECT transition_log_json FROM graphics_handoff_ledger WHERE graphics_request_id = ?",
        (graphics_request_id,),
    ).fetchone()
    connection.execute(
        """
        UPDATE graphics_handoff_ledger
        SET pig_status = 'claimed',
            handoff_status = 'claimed',
            claimed_by = ?,
            claimed_at = COALESCE(claimed_at, ?),
            updated_at = ?,
            transition_log_json = ?
        WHERE graphics_request_id = ?
        """,
        (
            claimed_by,
            now,
            now,
            append_transition_log(row["transition_log_json"], {"at": now, "event": "claimed", "claimedBy": claimed_by}),
            graphics_request_id,
        ),
    )
    connection.commit()
    return get_graphics_handoff(connection, graphics_request_id) or {}


def update_graphics_handoff(connection: sqlite3.Connection, graphics_request_id: str, update: dict[str, Any]) -> dict[str, Any]:
    existing = get_graphics_handoff(connection, graphics_request_id)
    if not existing:
        raise KeyError(f"Unknown graphicsRequestId: {graphics_request_id}")

    now = utc_now_iso()
    handoff_status = normalize_enum(update.get("handoffStatus"), HANDOFF_STATUSES, existing["handoffStatus"])
    pig_status = normalize_enum(update.get("pigStatus"), PIG_STATUSES, existing["pigStatus"])
    qc_status = normalize_enum(update.get("qcStatus"), QC_STATUSES, existing["qcStatus"])
    asset_url = str(update.get("assetUrl") or update.get("assetLinkUrl") or update.get("driveUrl") or existing["assetUrl"] or "")
    uploaded = handoff_status in {"uploaded", "sent_to_weaver_qc", "approved"} or pig_status == "uploaded"
    generated = uploaded or handoff_status in {"generated", "exported", "sent_to_weaver_qc", "approved"} or pig_status in {"generated", "exported", "uploaded"}
    sent_to_qc = handoff_status in {"sent_to_weaver_qc", "approved", "rejected"} or qc_status in {"pending", "approved", "rejected", "needs_revision"}
    approved = handoff_status == "approved" or qc_status == "approved"
    rejected = handoff_status == "rejected" or qc_status in {"rejected", "needs_revision"}

    row = connection.execute(
        "SELECT transition_log_json FROM graphics_handoff_ledger WHERE graphics_request_id = ?",
        (graphics_request_id,),
    ).fetchone()
    event = {
        "at": now,
        "event": "updated",
        "handoffStatus": handoff_status,
        "pigStatus": pig_status,
        "qcStatus": qc_status,
    }
    connection.execute(
        """
        UPDATE graphics_handoff_ledger
        SET source_status = ?,
            pig_status = ?,
            handoff_status = ?,
            qc_status = ?,
            asset_url = ?,
            asset_preview_url = ?,
            drive_file_id = ?,
            drive_file_name = ?,
            mime_type = ?,
            export_type = ?,
            variant = ?,
            version = ?,
            error_message = ?,
            blocked_reason = ?,
            pig_payload_json = ?,
            qc_payload_json = ?,
            updated_at = ?,
            generated_at = CASE WHEN ? THEN COALESCE(generated_at, ?) ELSE generated_at END,
            uploaded_at = CASE WHEN ? THEN COALESCE(uploaded_at, ?) ELSE uploaded_at END,
            sent_to_qc_at = CASE WHEN ? THEN COALESCE(sent_to_qc_at, ?) ELSE sent_to_qc_at END,
            approved_at = CASE WHEN ? THEN COALESCE(approved_at, ?) ELSE approved_at END,
            rejected_at = CASE WHEN ? THEN COALESCE(rejected_at, ?) ELSE rejected_at END,
            transition_log_json = ?
        WHERE graphics_request_id = ?
        """,
        (
            str(update.get("sourceStatus") or existing["sourceStatus"]),
            pig_status,
            handoff_status,
            qc_status,
            asset_url,
            str(update.get("assetPreviewUrl") or update.get("previewUrl") or update.get("thumbnailUrl") or existing["assetPreviewUrl"] or ""),
            str(update.get("driveFileId") or update.get("fileId") or existing["driveFileId"] or ""),
            str(update.get("driveFileName") or update.get("fileName") or existing["driveFileName"] or ""),
            str(update.get("mimeType") or existing["mimeType"] or ""),
            str(update.get("exportType") or existing["exportType"] or ""),
            str(update.get("variant") or existing["variant"] or ""),
            str(update.get("version") or existing["version"] or ""),
            str(update.get("errorMessage") or existing["errorMessage"] or ""),
            str(update.get("blockedReason") or existing["blockedReason"] or ""),
            json.dumps(update.get("pigPayload") or update, ensure_ascii=True, sort_keys=True),
            json.dumps(update.get("qcPayload") or update, ensure_ascii=True, sort_keys=True),
            now,
            generated,
            now,
            uploaded,
            now,
            sent_to_qc,
            now,
            approved,
            now,
            rejected,
            now,
            append_transition_log(row["transition_log_json"], event),
            graphics_request_id,
        ),
    )
    connection.commit()
    return get_graphics_handoff(connection, graphics_request_id) or {}


def get_graphics_handoff_queue(connection: sqlite3.Connection, limit: int = 100) -> list[dict[str, Any]]:
    rows = connection.execute(
        """
        SELECT *
        FROM graphics_handoff_ledger
        WHERE handoff_status IN ('requested', 'claimed', 'rejected')
          AND pig_status NOT IN ('generated', 'exported', 'uploaded', 'failed')
          AND qc_status IN ('not_sent', 'needs_revision')
        ORDER BY created_at ASC
        LIMIT ?
        """,
        (max(1, min(int(limit or 100), 500)),),
    ).fetchall()
    return [
        record
        for record in (row_to_handoff(row) for row in rows)
        if record["quoteText"].strip()
    ]


def replace_graphics_request_items(connection: sqlite3.Connection, graphics_request_id: str, items: list[dict[str, Any]]) -> None:
    connection.execute("DELETE FROM graphics_request_items WHERE graphics_request_id = ?", (graphics_request_id,))
    for position, item in enumerate(items, start=1):
        quote_text = str(item.get("quote_text") or "")
        connection.execute(
            """
            INSERT INTO graphics_request_items (
                graphics_request_id, item_position, source_record_id, source_sheet_name, source_sheet_row,
                book_title, poem_title, author, quote_text, normalized_quote_key, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                graphics_request_id,
                position,
                str(item.get("source_record_id") or ""),
                str(item.get("source_sheet_name") or ""),
                int(item.get("source_sheet_row") or 0) or None,
                str(item.get("book_title") or ""),
                str(item.get("poem_title") or ""),
                str(item.get("author") or ""),
                quote_text,
                normalize_key(quote_text),
                str(item.get("created_at") or utc_now_iso()),
            ),
        )
    connection.commit()


def insert_graphics_completion(connection: sqlite3.Connection, completion: dict[str, Any]) -> str:
    completion_id = str(completion["id"]).strip()
    graphics_request_id = str(completion["graphics_request_id"]).strip()
    ingested_at = completion.get("ingested_at") or utc_now_iso()
    payload_json = json.dumps(completion.get("source_payload") or {}, ensure_ascii=True, sort_keys=True)

    connection.execute(
        """
        INSERT INTO graphics_completions (
            id, graphics_request_id, source_tool, asset_url, asset_preview_url,
            production_notes, completion_status, completed_at, ingested_at, source_payload_json
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
            graphics_request_id = excluded.graphics_request_id,
            source_tool = excluded.source_tool,
            asset_url = excluded.asset_url,
            asset_preview_url = excluded.asset_preview_url,
            production_notes = excluded.production_notes,
            completion_status = excluded.completion_status,
            completed_at = excluded.completed_at,
            ingested_at = excluded.ingested_at,
            source_payload_json = excluded.source_payload_json
        """,
        (
            completion_id,
            graphics_request_id,
            str(completion.get("source_tool") or "P.I.G."),
            str(completion.get("asset_url") or ""),
            str(completion.get("asset_preview_url") or ""),
            str(completion.get("production_notes") or ""),
            str(completion.get("completion_status") or "RETURNED"),
            str(completion.get("completed_at") or utc_now_iso()),
            ingested_at,
            payload_json,
        ),
    )
    connection.execute(
        """
        UPDATE graphics_requests
        SET latest_completion_id = ?, request_status = ?, updated_at = ?
        WHERE id = ?
        """,
        (completion_id, "COMPLETED_RETURNED", utc_now_iso(), graphics_request_id),
    )
    connection.commit()
    return completion_id


def insert_graphics_qc_review(connection: sqlite3.Connection, review: dict[str, Any]) -> int:
    reviewed_at = review.get("reviewed_at") or utc_now_iso()
    payload_json = json.dumps(review.get("source_payload") or {}, ensure_ascii=True, sort_keys=True)
    cursor = connection.execute(
        """
        INSERT INTO graphics_qc_reviews (
            graphics_completion_id, decision, metadata_issue, aesthetic_issue,
            note, reviewed_by, reviewed_at, source_payload_json
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            str(review["graphics_completion_id"]).strip(),
            str(review["decision"]).strip(),
            str(review.get("metadata_issue") or ""),
            str(review.get("aesthetic_issue") or ""),
            str(review.get("note") or ""),
            str(review.get("reviewed_by") or ""),
            reviewed_at,
            payload_json,
        ),
    )
    connection.commit()
    return int(cursor.lastrowid)


def insert_poetry_please_handoff(connection: sqlite3.Connection, handoff: dict[str, Any]) -> int:
    cursor = connection.execute(
        """
        INSERT INTO poetry_please_handoffs (
            graphics_completion_id, handoff_status, handoff_mode,
            handed_off_at, poetry_please_item_id, payload_json
        ) VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            str(handoff["graphics_completion_id"]).strip(),
            str(handoff.get("handoff_status") or "QUEUED"),
            str(handoff.get("handoff_mode") or "preview"),
            str(handoff.get("handed_off_at") or utc_now_iso()),
            str(handoff.get("poetry_please_item_id") or ""),
            json.dumps(handoff.get("payload") or {}, ensure_ascii=True, sort_keys=True),
        ),
    )
    connection.commit()
    return int(cursor.lastrowid)


def get_latest_graphics_qc_reviews(
    connection: sqlite3.Connection,
    completion_ids: list[str] | None = None,
) -> dict[str, dict[str, Any]]:
    params: list[Any] = []
    sql = """
        SELECT graphics_completion_id, decision, metadata_issue, aesthetic_issue,
               note, reviewed_by, reviewed_at
        FROM v_latest_graphics_qc_reviews
    """
    if completion_ids:
        placeholders = ",".join("?" for _ in completion_ids)
        sql += f" WHERE graphics_completion_id IN ({placeholders})"
        params.extend(completion_ids)

    rows = connection.execute(sql, params).fetchall()
    return {
        str(row["graphics_completion_id"]): {
            "decision": str(row["decision"] or ""),
            "metadata_issue": str(row["metadata_issue"] or ""),
            "aesthetic_issue": str(row["aesthetic_issue"] or ""),
            "note": str(row["note"] or ""),
            "reviewed_by": str(row["reviewed_by"] or ""),
            "reviewed_at": str(row["reviewed_at"] or ""),
        }
        for row in rows
    }


def get_latest_poetry_please_handoffs(
    connection: sqlite3.Connection,
    completion_ids: list[str] | None = None,
) -> dict[str, dict[str, Any]]:
    params: list[Any] = []
    sql = """
        SELECT handoff.graphics_completion_id, handoff.handoff_status, handoff.handoff_mode,
               handoff.handed_off_at, handoff.poetry_please_item_id, handoff.payload_json
        FROM poetry_please_handoffs AS handoff
        JOIN (
            SELECT graphics_completion_id, MAX(handed_off_at) AS max_handed_off_at, MAX(id) AS max_id
            FROM poetry_please_handoffs
            GROUP BY graphics_completion_id
        ) AS latest
          ON latest.graphics_completion_id = handoff.graphics_completion_id
         AND latest.max_handed_off_at = handoff.handed_off_at
         AND latest.max_id = handoff.id
    """
    if completion_ids:
        placeholders = ",".join("?" for _ in completion_ids)
        sql += f" WHERE handoff.graphics_completion_id IN ({placeholders})"
        params.extend(completion_ids)

    rows = connection.execute(sql, params).fetchall()
    return {
        str(row["graphics_completion_id"]): {
            "handoff_status": str(row["handoff_status"] or ""),
            "handoff_mode": str(row["handoff_mode"] or ""),
            "handed_off_at": str(row["handed_off_at"] or ""),
            "poetry_please_item_id": str(row["poetry_please_item_id"] or ""),
            "payload_json": str(row["payload_json"] or ""),
        }
        for row in rows
    }
