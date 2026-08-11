#!/usr/bin/env python3
from __future__ import annotations

import json
import hashlib
import os
import re
import ssl
import sqlite3
import subprocess
import urllib.error
import urllib.parse
import urllib.request
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
DEFAULT_DB_PATH = ROOT / "data" / "weaver_runtime.db"
SCHEMA_PATH = ROOT / "db" / "weaver_runtime_schema.sql"
FIRESTORE_COLLECTION = "graphicsHandoffLedger"
FIRESTORE_HANDOFF_ALIASES_COLLECTION = "graphicsHandoffAliases"
FIRESTORE_GRAPHICS_REQUESTS_COLLECTION = "graphicsRequests"
FIRESTORE_GRAPHICS_COMPLETIONS_COLLECTION = "graphicsCompletions"
FIRESTORE_GRAPHICS_QC_REVIEWS_COLLECTION = "graphicsQcReviews"
FIRESTORE_GRAPHICS_QC_QUEUE_COLLECTION = "graphicsQcQueueCards"
FIRESTORE_POETRY_PLEASE_HANDOFFS_COLLECTION = "poetryPleaseHandoffs"
FIRESTORE_EXCERPT_RECORDS_COLLECTION = "excerptRecords"
FIRESTORE_DEFAULT_DATABASE_ID = "weaverledger"
FIRESTORE_DEFAULT_PROJECT_ID = "button-weaver-internal"
FIRESTORE_DEFAULT_SERVICE_ACCOUNT = (
    "weaver-deployer@button-weaver-internal.iam.gserviceaccount.com"
)
HANDOFF_TRANSITION_LOG_LIMIT = 12
HANDOFF_LIST_TRANSITION_LOG_LIMIT = 3
_FIRESTORE_ACCESS_TOKEN_CACHE: tuple[str, float] | None = None


def utc_now_iso() -> str:
    return datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def normalize_key(value: str) -> str:
    return " ".join((value or "").strip().lower().split())


def clean_durable_value(value: Any) -> str:
    if value is None:
        return ""
    cleaned = str(value).strip()
    return "" if cleaned.lower() in {"none", "null", "undefined"} else cleaned


def first_durable_value(*values: Any) -> str:
    for value in values:
        cleaned = clean_durable_value(value)
        if cleaned:
            return cleaned
    return ""


def count_words(text: str) -> int:
    return len([part for part in (text or "").split() if part.strip()])


def normalize_content_type(value: Any, default: str = "QI") -> str:
    normalized = str(value or "").strip().upper()
    if normalized in {"QI", "QUOTE IMAGE"}:
        return "QI"
    if normalized in {"FP", "FULL POEM"}:
        return "FP"
    if normalized in {"FPI", "FULL POEM IMAGE", "FULL POEM IMAGE-BACKED", "FULL POEM IMAGE BACKED"}:
        return "FPI"
    return default


def build_graphics_qc_queue_card(completion: dict[str, Any]) -> dict[str, Any]:
    payload = completion.get("sourcePayload") if isinstance(completion.get("sourcePayload"), dict) else {}
    completion_id = str(completion.get("id") or payload.get("pigCompletionId") or "").strip()
    graphics_request_id = str(
        completion.get("graphicsRequestId") or payload.get("graphicsRequestId") or ""
    ).strip()
    asset_url = str(completion.get("assetUrl") or payload.get("assetUrl") or "").strip()
    pig_project_id = str(completion.get("pigProjectId") or payload.get("pigProjectId") or "").strip()
    editable_project_file_id = str(
        completion.get("editableProjectFileId")
        or completion.get("projectFileId")
        or payload.get("editableProjectFileId")
        or payload.get("projectFileId")
        or ""
    ).strip()
    editable_project_url = str(
        completion.get("editableProjectUrl") or payload.get("editableProjectUrl") or ""
    ).strip()
    content_id = str(
        completion.get("contentId")
        or completion.get("imageId")
        or payload.get("contentId")
        or payload.get("imageId")
        or payload.get("sourceRecordId")
        or ""
    ).strip()
    text_hash = str(completion.get("textHash") or payload.get("textHash") or "").strip()
    if not text_hash and payload.get("quoteText"):
        text_hash = stable_text_hash(payload.get("quoteText"))
    source_identity = completion.get("sourceIdentity") or payload.get("sourceIdentity")
    if not isinstance(source_identity, dict):
        source_identity = {}
    source_identity = {
        **source_identity,
        "graphicsRequestId": str(source_identity.get("graphicsRequestId") or graphics_request_id),
        "contentId": str(source_identity.get("contentId") or content_id),
        "imageId": str(source_identity.get("imageId") or content_id),
        "textHash": str(source_identity.get("textHash") or text_hash),
    }
    return {
        "pigCompletionId": completion_id,
        "graphicsRequestId": graphics_request_id,
        "recordId": str(payload.get("recordId") or graphics_request_id or completion_id),
        "sheetRow": payload.get("sheetRow") or 0,
        "storageTarget": str(payload.get("storageTarget") or "firestore"),
        "contentType": normalize_content_type(
            completion.get("contentType") or payload.get("contentType")
        ),
        "author": str(payload.get("author") or ""),
        "poemTitle": str(payload.get("poemTitle") or ""),
        "bookTitle": str(payload.get("bookTitle") or ""),
        "quoteText": str(payload.get("quoteText") or ""),
        "assetLinkUrl": asset_url,
        "assetPreviewUrl": str(
            completion.get("assetPreviewUrl") or payload.get("assetPreviewUrl") or asset_url
        ).strip(),
        "pigProjectId": pig_project_id,
        "editableProjectFileId": editable_project_file_id,
        "editableProjectUrl": editable_project_url,
        "editableProjectAvailable": bool(
            pig_project_id and editable_project_file_id and editable_project_url
        ),
        "editableProjectValidationStatus": str(
            completion.get("editableProjectValidationStatus")
            or payload.get("editableProjectValidationStatus")
            or ""
        ),
        "editableProjectValidationError": (
            completion.get("editableProjectValidationError")
            or payload.get("editableProjectValidationError")
            or {}
        ),
        "contentId": content_id,
        "imageId": content_id,
        "sourceIdentity": source_identity,
        "textHash": text_hash,
        "completedAt": str(completion.get("completedAt") or payload.get("completedAt") or ""),
        "graphicsQcDecision": "",
        "graphicsQcNote": "",
        "graphicsQcUpdatedAt": "",
        "poetryPleaseStatus": "",
        "poetryPleaseUpdatedAt": "",
        "poetryPleaseNote": "",
        "isPendingQc": True,
        "updatedAt": utc_now_iso(),
    }


def normalize_identity_text(value: Any) -> str:
    return " ".join(str(value or "").split()).casefold()


def stable_text_hash(value: Any) -> str:
    return hashlib.sha256(normalize_identity_text(value).encode("utf-8")).hexdigest()


def assert_completion_identity_consistent(record: dict[str, Any]) -> None:
    payload = record.get("sourcePayload") if isinstance(record.get("sourcePayload"), dict) else {}
    source_identity = record.get("sourceIdentity") or payload.get("sourceIdentity")
    if not isinstance(source_identity, dict):
        source_identity = {}
    request_ids = {
        str(value).strip().casefold()
        for value in (
            record.get("graphicsRequestId"),
            payload.get("graphicsRequestId"),
            payload.get("requestId"),
            source_identity.get("graphicsRequestId"),
        )
        if str(value or "").strip()
    }
    if len(request_ids) > 1:
        raise ValueError("Completion contains conflicting graphicsRequestId values")

    text_hashes = {
        str(value).strip().casefold()
        for value in (
            record.get("textHash"),
            payload.get("textHash"),
            source_identity.get("textHash"),
        )
        if len(str(value or "").strip()) == 64
    }
    quote_text = payload.get("quoteText") or record.get("quoteText")
    if quote_text:
        text_hashes.add(stable_text_hash(quote_text))
    if len(text_hashes) > 1:
        raise ValueError("Completion contains a textHash that conflicts with its quote text")


def completion_canonical_identity(record: dict[str, Any]) -> str:
    payload = record.get("sourcePayload") if isinstance(record.get("sourcePayload"), dict) else {}
    source_identity = record.get("sourceIdentity") or payload.get("sourceIdentity")
    if not isinstance(source_identity, dict):
        source_identity = {}
    for value in (
        source_identity.get("graphicsRequestId"),
        record.get("graphicsRequestId"),
        payload.get("graphicsRequestId"),
        payload.get("requestId"),
        source_identity.get("contentId"),
        source_identity.get("imageId"),
        record.get("contentId"),
        record.get("imageId"),
        payload.get("contentId"),
        payload.get("imageId"),
        payload.get("sourceRecordId"),
        source_identity.get("textHash"),
        record.get("textHash"),
        payload.get("textHash"),
    ):
        normalized = str(value or "").strip()
        if normalized:
            return normalized.casefold()
    quote_text = payload.get("quoteText") or record.get("quoteText")
    return stable_text_hash(quote_text) if quote_text else ""


def assert_editable_project_file_not_cross_linked(
    existing_records: list[dict[str, Any]], candidate: dict[str, Any]
) -> None:
    file_id = str(candidate.get("editableProjectFileId") or "").strip()
    candidate_identity = completion_canonical_identity(candidate)
    if not file_id or not candidate_identity:
        return
    for existing in existing_records:
        existing_id = str(existing.get("id") or "").strip()
        if existing_id and existing_id == str(candidate.get("id") or "").strip():
            continue
        existing_identity = completion_canonical_identity(existing)
        if existing_identity and existing_identity != candidate_identity:
            raise ValueError(
                f"Editable project file {file_id} is already linked to canonical identity "
                f"{existing_identity}; refusing cross-link to {candidate_identity}."
            )


def find_editable_project_link_conflicts(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by_file: dict[str, list[dict[str, str]]] = {}
    for record in records:
        file_id = str(record.get("editableProjectFileId") or "").strip()
        identity = completion_canonical_identity(record)
        if not file_id or not identity:
            continue
        by_file.setdefault(file_id, []).append({
            "completionId": str(record.get("id") or ""),
            "graphicsRequestId": str(record.get("graphicsRequestId") or ""),
            "canonicalIdentity": identity,
        })
    return [
        {
            "editableProjectFileId": file_id,
            "canonicalIdentities": sorted({entry["canonicalIdentity"] for entry in entries}),
            "records": entries,
        }
        for file_id, entries in sorted(by_file.items())
        if len({entry["canonicalIdentity"] for entry in entries}) > 1
    ]


def connect_runtime_db(db_path: Path = DEFAULT_DB_PATH) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def ensure_runtime_schema(connection: sqlite3.Connection, schema_path: Path = SCHEMA_PATH) -> None:
    connection.executescript(schema_path.read_text())
    ensure_runtime_migrations(connection)
    connection.commit()


def ensure_columns(connection: sqlite3.Connection, table: str, columns: dict[str, str]) -> None:
    existing = {
        str(row["name"])
        for row in connection.execute(f"PRAGMA table_info({table})").fetchall()
    }
    for name, definition in columns.items():
        if name not in existing:
            connection.execute(f"ALTER TABLE {table} ADD COLUMN {name} {definition}")


def ensure_runtime_migrations(connection: sqlite3.Connection) -> None:
    ensure_columns(connection, "graphics_requests", {
        "content_type": "TEXT NOT NULL DEFAULT 'QI'",
        "image_type": "TEXT NOT NULL DEFAULT 'QI'",
    })
    ensure_columns(connection, "graphics_completions", {
        "content_type": "TEXT NOT NULL DEFAULT 'QI'",
        "image_type": "TEXT NOT NULL DEFAULT 'QI'",
    })
    ensure_columns(connection, "graphics_handoff_ledger", {
        "content_type": "TEXT NOT NULL DEFAULT 'QI'",
        "image_type": "TEXT NOT NULL DEFAULT 'QI'",
        "source_completion_id": "TEXT NOT NULL DEFAULT ''",
        "revision_of": "TEXT NOT NULL DEFAULT ''",
        "original_graphics_request_id": "TEXT NOT NULL DEFAULT ''",
        "review_status": "TEXT NOT NULL DEFAULT ''",
        "ocr_text": "TEXT NOT NULL DEFAULT ''",
    })
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_graphics_handoff_ledger_content "
        "ON graphics_handoff_ledger(content_type, image_type, review_status)"
    )


def get_runtime_db_summary(connection: sqlite3.Connection) -> dict[str, Any]:
    tables = [
        "graphics_requests",
        "graphics_request_items",
        "graphics_completions",
        "graphics_handoff_ledger",
        "graphics_qc_reviews",
        "poetry_please_handoffs",
        "excerpt_handoff_ledger",
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
    if is_firestore_ledger(connection):
        return connection.upsert_graphics_request(request)
    request_id = str(request["id"]).strip()
    created_at = request.get("created_at") or utc_now_iso()
    updated_at = request.get("updated_at") or created_at
    quote_text = str(request.get("quote_text") or "")
    payload_json = json.dumps(request.get("source_payload") or {}, ensure_ascii=True, sort_keys=True)
    content_type = normalize_content_type(request.get("content_type") or request.get("contentType") or request.get("imageType"))
    image_type = normalize_content_type(request.get("image_type") or request.get("imageType") or request.get("contentType"), content_type)

    connection.execute(
        """
        INSERT INTO graphics_requests (
            id, request_status, source_type, content_type, image_type, book_title, poem_title, author, quote_text,
            normalized_book_key, normalized_poem_key, normalized_author_key, normalized_quote_key,
            word_count, source_record_id, source_sheet_name, source_sheet_row,
            source_payload_json, created_at, updated_at, latest_completion_id
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
            request_status = excluded.request_status,
            source_type = excluded.source_type,
            content_type = excluded.content_type,
            image_type = excluded.image_type,
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
            content_type,
            image_type,
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
TERMINAL_HANDOFF_STATUSES = {
    "generated",
    "exported",
    "uploaded",
    "sent_to_weaver_qc",
    "approved",
    "blocked",
    "errored",
}
SOURCE_QUEUE_REFRESH_STATUSES = {"requested", "claimed", "rejected"}
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


def extract_nested_payload(payload: dict[str, Any]) -> dict[str, Any]:
    nested = payload.get("source_payload") if isinstance(payload.get("source_payload"), dict) else {}
    if not nested:
        nested = payload.get("sourcePayload") if isinstance(payload.get("sourcePayload"), dict) else {}
    return nested


def extract_handoff_value(payload: dict[str, Any], *keys: str) -> str:
    nested = extract_nested_payload(payload)
    for key in keys:
        value = str(payload.get(key) or nested.get(key) or "").strip()
        if value:
            return value
    return ""


def has_fpi_source_asset(payload: dict[str, Any]) -> bool:
    return bool(extract_handoff_value(
        payload,
        "driveLink",
        "imageUrl",
        "assetUrl",
        "assetLinkUrl",
        "assetPreviewUrl",
        "previousAssetUrl",
        "previousAssetPreviewUrl",
    ))


def infer_graphics_content_type(
    request: dict[str, Any],
    source_payload: dict[str, Any],
    existing: dict[str, Any] | None = None,
) -> str:
    raw_type = (
        request.get("contentType")
        or request.get("imageType")
        or extract_handoff_value(source_payload, "contentType", "imageType")
    )
    explicit_type = normalize_content_type(raw_type, "")
    if explicit_type == "FP":
        explicit_type = "FPI"

    identity_values = [
        request.get("graphicsRequestId"),
        request.get("id"),
        request.get("sourceRecordId"),
        request.get("canonicalPoemId"),
        request.get("poemId"),
        extract_handoff_value(
            source_payload,
            "graphicsRequestId",
            "sourceRecordId",
            "canonicalPoemId",
            "poemId",
            "fullPoemId",
        ),
    ]
    has_fpi_identity = any(
        re.search(r"(^|[-:])FPI?($|[-:])", str(value or ""), re.IGNORECASE)
        for value in identity_values
    )
    has_qi_identity = any(
        re.search(r"(^|[-:])QI($|[-:])", str(value or ""), re.IGNORECASE)
        for value in identity_values
    )

    if explicit_type == "QI" and has_fpi_identity:
        raise ValueError("contentType QI conflicts with an FP/FPI request identity")
    if explicit_type:
        return explicit_type
    if existing and existing.get("contentType"):
        existing_type = normalize_content_type(existing.get("contentType"), "")
        return "FPI" if existing_type == "FP" else existing_type
    if has_fpi_identity:
        return "FPI"
    if has_qi_identity:
        return "QI"
    raise ValueError("contentType is required for a new graphics handoff request")


def slug_identity_token(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]+", "-", str(value or "").strip().upper()).strip("-")


def canonical_fpi_graphics_request_id(canonical_content_id: str) -> str:
    digest = hashlib.sha256(canonical_content_id.encode("utf-8")).hexdigest()[:32]
    return f"weaver:fpi:{digest}"


def canonical_qi_content_id(request: dict[str, Any], source_payload: dict[str, Any]) -> str:
    book_title = str(
        request.get("bookTitle")
        or extract_handoff_value(source_payload, "bookTitle", "book_title", "book")
        or ""
    ).strip()
    author = str(
        request.get("author")
        or extract_handoff_value(source_payload, "author", "author_name")
        or ""
    ).strip()
    poem_title = str(
        request.get("poemTitle")
        or request.get("title")
        or extract_handoff_value(source_payload, "poemTitle", "poem_title", "title")
        or ""
    ).strip()
    quote_text = str(
        request.get("quoteText")
        or request.get("text")
        or extract_handoff_text(source_payload)
        or ""
    ).strip()
    if not quote_text or not (book_title or poem_title):
        raise ValueError("QI requests require excerpt text and book or poem identity")
    identity = "\n".join(normalize_key(value) for value in (
        book_title,
        author,
        poem_title,
        quote_text,
    ))
    digest = hashlib.sha256(identity.encode("utf-8")).hexdigest()[:32]
    return f"QI:EXCERPT:{digest}"


def canonical_qi_graphics_request_id(canonical_content_id: str) -> str:
    digest = str(canonical_content_id or "").rsplit(":", 1)[-1].strip().lower()
    if not re.fullmatch(r"[0-9a-f]{32}", digest):
        digest = hashlib.sha256(canonical_content_id.encode("utf-8")).hexdigest()[:32]
    return f"weaver:qi:{digest}"


def canonical_fpi_content_id(request: dict[str, Any], source_payload: dict[str, Any]) -> str:
    book_shortener = str(
        request.get("bookShortener")
        or extract_handoff_value(source_payload, "bookShortener", "book_shortener")
        or ""
    ).strip()
    poem_title = str(
        request.get("poemTitle")
        or request.get("title")
        or extract_handoff_value(source_payload, "poemTitle", "poem_title", "title")
        or ""
    ).strip()
    if book_shortener and poem_title:
        return f"FPI:{slug_identity_token(book_shortener)}-FPI-{slug_identity_token(poem_title)}"

    book_title = str(
        request.get("bookTitle")
        or request.get("book")
        or extract_handoff_value(source_payload, "bookTitle", "book_title", "book")
        or ""
    ).strip()
    if book_title and poem_title:
        digest = hashlib.sha256(
            f"{normalize_key(book_title)}\n{normalize_key(poem_title)}".encode("utf-8")
        ).hexdigest()[:24]
        return f"FPI:BOOK-POEM-{digest}"

    canonical_poem_id = str(
        request.get("canonicalPoemId")
        or request.get("poemId")
        or request.get("fullPoemId")
        or extract_handoff_value(source_payload, "canonicalPoemId", "poemId", "fullPoemId")
        or ""
    ).strip()
    if canonical_poem_id:
        canonical_token = re.sub(
            r"(^|-)FP(?=-|$)",
            r"\1FPI",
            slug_identity_token(canonical_poem_id),
        )
        return f"FPI:{canonical_token}"

    source_record_id = str(
        request.get("sourceRecordId")
        or extract_handoff_value(source_payload, "sourceRecordId", "source_record_id")
        or request.get("graphicsRequestId")
        or ""
    ).strip()
    if source_record_id and re.search(r"(^|[-:])FPI?($|[-:])", source_record_id, re.IGNORECASE):
        return f"FPI:{slug_identity_token(source_record_id)}"
    raise ValueError("FPI requests require canonicalPoemId, book and poem identity, or a canonical FP/FPI sourceRecordId")


def assert_compatible_graphics_content_type(existing: dict[str, Any] | None, content_type: str) -> None:
    if not existing:
        return
    existing_type = normalize_content_type(existing.get("contentType") or existing.get("imageType"), "")
    if existing_type == "FP":
        existing_type = "FPI"
    if existing_type and existing_type != content_type:
        raise ValueError(
            f"graphicsRequestId already belongs to contentType {existing_type}, not {content_type}"
        )


def row_to_handoff(row: sqlite3.Row) -> dict[str, Any]:
    source_payload = json.loads(row["source_payload_json"] or "{}")
    nested_source = extract_nested_payload(source_payload)
    quote_text = extract_handoff_text(source_payload)
    if not quote_text and nested_source:
        quote_text = extract_handoff_text(nested_source)
    record = {
        "graphicsRequestId": str(row["graphics_request_id"] or ""),
        "sourceSystem": str(row["source_system"] or ""),
        "sourceStatus": str(row["source_status"] or ""),
        "contentType": str(row["content_type"] or ""),
        "imageType": str(row["image_type"] or ""),
        "sourceCompletionId": str(row["source_completion_id"] or ""),
        "revisionOf": str(row["revision_of"] or ""),
        "originalGraphicsRequestId": str(row["original_graphics_request_id"] or ""),
        "reviewStatus": str(row["review_status"] or ""),
        "ocrText": str(row["ocr_text"] or ""),
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
        "sourceSheetRow": source_payload.get("sourceSheetRow") or source_payload.get("source_sheet_row") or source_payload.get("sheetRow") or nested_source.get("sourceSheetRow") or nested_source.get("source_sheet_row") or nested_source.get("sheetRow") or "",
        "queueSheetRow": source_payload.get("queueSheetRow") or source_payload.get("sourceSheetRow") or source_payload.get("sheetRow") or nested_source.get("queueSheetRow") or nested_source.get("sourceSheetRow") or nested_source.get("sheetRow") or "",
        "author": str(source_payload.get("author") or source_payload.get("author_name") or nested_source.get("author") or nested_source.get("author_name") or ""),
        "poemTitle": str(source_payload.get("poemTitle") or source_payload.get("title") or source_payload.get("poem_title") or nested_source.get("poemTitle") or nested_source.get("title") or nested_source.get("poem_title") or ""),
        "bookTitle": str(source_payload.get("bookTitle") or source_payload.get("book_title") or nested_source.get("bookTitle") or nested_source.get("book_title") or ""),
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
    return normalize_handoff_record(record)


def default_handoff_record(graphics_request_id: str) -> dict[str, Any]:
    return {
        "graphicsRequestId": graphics_request_id,
        "sourceSystem": "",
        "sourceStatus": "",
        "queueView": "",
        "isActionable": False,
        "nextAction": "",
        "statusLabel": "",
        "contentType": "",
        "imageType": "",
        "targetCount": 0,
        "approvedCount": 0,
        "pendingQcCount": 0,
        "inProgressCount": 0,
        "reworkCount": 0,
        "remainingApprovedNeeded": 0,
        "remainingActionableNeeded": 0,
        "priorityTier": "",
        "priorityScore": 0,
        "sourceCompletionId": "",
        "revisionOf": "",
        "originalGraphicsRequestId": "",
        "reviewStatus": "",
        "ocrText": "",
        "pigStatus": "",
        "handoffStatus": "",
        "qcStatus": "",
        "assetUrl": "",
        "assetPreviewUrl": "",
        "assetFileId": "",
        "driveFileId": "",
        "driveFileName": "",
        "mimeType": "",
        "exportType": "",
        "variant": "",
        "version": "",
        "pigProjectId": "",
        "editableProjectFileId": "",
        "editableProjectUrl": "",
        "editableProjectKind": "",
        "editableProjectSchemaVersion": "",
        "editableProjectAvailable": False,
        "reworkReason": "",
        "metadataIssue": "",
        "aestheticIssue": "",
        "qcNote": "",
        "requestedChanges": "",
        "claimedBy": "",
        "errorMessage": "",
        "blockedReason": "",
        "sourceSheetRow": "",
        "queueSheetRow": "",
        "author": "",
        "poemTitle": "",
        "bookTitle": "",
        "bookKey": "",
        "quoteText": "",
        "text": "",
        "sourcePayload": {},
        "pigPayload": {},
        "qcPayload": {},
        "transitionLog": [],
        "createdAt": "",
        "updatedAt": "",
        "claimedAt": "",
        "generatedAt": "",
        "uploadedAt": "",
        "sentToQcAt": "",
        "approvedAt": "",
        "rejectedAt": "",
    }


def normalize_handoff_record(record: dict[str, Any]) -> dict[str, Any]:
    graphics_request_id = str(record.get("graphicsRequestId") or record.get("id") or "").strip()
    normalized = default_handoff_record(graphics_request_id)
    normalized.update(record)
    normalized["graphicsRequestId"] = graphics_request_id
    for key in ("sourcePayload", "pigPayload", "qcPayload"):
        if not isinstance(normalized.get(key), dict):
            normalized[key] = {}
    if not isinstance(normalized.get("transitionLog"), list):
        normalized["transitionLog"] = []

    source_payload = normalized["sourcePayload"]
    nested_source = extract_nested_payload(source_payload)
    qc_payload = normalized["qcPayload"]
    pig_payload = normalized["pigPayload"]
    payload_sources = [
        source_payload,
        nested_source,
        qc_payload,
        extract_nested_payload(qc_payload),
        pig_payload,
        extract_nested_payload(pig_payload),
    ]
    quote_text = str(normalized.get("quoteText") or extract_handoff_text(source_payload) or "").strip()
    if not quote_text and nested_source:
        quote_text = extract_handoff_text(nested_source)
    normalized["quoteText"] = quote_text
    normalized["text"] = str(normalized.get("text") or quote_text)
    normalized["sourceSheetRow"] = (
        normalized.get("sourceSheetRow")
        or source_payload.get("sourceSheetRow")
        or source_payload.get("source_sheet_row")
        or source_payload.get("sheetRow")
        or nested_source.get("sourceSheetRow")
        or nested_source.get("source_sheet_row")
        or nested_source.get("sheetRow")
        or ""
    )
    normalized["queueSheetRow"] = (
        normalized.get("queueSheetRow")
        or source_payload.get("queueSheetRow")
        or source_payload.get("sourceSheetRow")
        or source_payload.get("sheetRow")
        or nested_source.get("queueSheetRow")
        or nested_source.get("sourceSheetRow")
        or nested_source.get("sheetRow")
        or ""
    )
    normalized["author"] = str(normalized.get("author") or source_payload.get("author") or source_payload.get("author_name") or nested_source.get("author") or nested_source.get("author_name") or "")
    normalized["poemTitle"] = str(normalized.get("poemTitle") or source_payload.get("poemTitle") or source_payload.get("title") or source_payload.get("poem_title") or nested_source.get("poemTitle") or nested_source.get("title") or nested_source.get("poem_title") or "")
    normalized["bookTitle"] = str(normalized.get("bookTitle") or source_payload.get("bookTitle") or source_payload.get("book_title") or nested_source.get("bookTitle") or nested_source.get("book_title") or "")
    normalized["bookKey"] = str(normalized.get("bookKey") or source_payload.get("bookKey") or source_payload.get("book_key") or nested_source.get("bookKey") or nested_source.get("book_key") or "")
    payload_fields = {
        "assetFileId": ("assetFileId", "driveFileId", "fileId"),
        "driveFileId": ("driveFileId", "assetFileId", "fileId"),
        "assetUrl": ("assetUrl", "assetLinkUrl", "driveUrl"),
        "assetPreviewUrl": ("assetPreviewUrl", "previewUrl", "thumbnailUrl"),
        "pigProjectId": ("pigProjectId", "pig_project_id"),
        "editableProjectFileId": ("editableProjectFileId", "projectFileId", "editable_project_file_id"),
        "editableProjectUrl": ("editableProjectUrl", "editable_project_url"),
        "editableProjectKind": ("editableProjectKind", "editable_project_kind"),
        "editableProjectSchemaVersion": ("editableProjectSchemaVersion", "editable_project_schema_version"),
        "originalGraphicsRequestId": ("originalGraphicsRequestId", "original_graphics_request_id"),
        "revisionOf": ("revisionOf", "revision_of"),
        "version": ("version",),
        "imageType": ("imageType", "contentType", "image_type", "content_type"),
        "reworkReason": ("reworkReason", "rejectReason", "rejectedReason"),
        "metadataIssue": ("metadataIssue", "metadata_issue"),
        "aestheticIssue": ("aestheticIssue", "aesthetic_issue"),
        "qcNote": ("qcNote", "graphicsQcNote", "qc_note"),
        "requestedChanges": ("requestedChanges", "requested_changes", "qcNote", "graphicsQcNote", "qc_note"),
    }
    for field, aliases in payload_fields.items():
        if clean_durable_value(normalized.get(field)):
            continue
        for alias in aliases:
            for payload_source in payload_sources:
                value = payload_source.get(alias)
                if clean_durable_value(value):
                    normalized[field] = clean_durable_value(value)
                    break
            if clean_durable_value(normalized.get(field)):
                break
    for field in (
        "assetFileId", "driveFileId", "assetUrl", "assetPreviewUrl", "pigProjectId",
        "editableProjectFileId", "editableProjectUrl", "editableProjectKind",
        "editableProjectSchemaVersion",
    ):
        normalized[field] = clean_durable_value(normalized.get(field))
    normalized["editableProjectAvailable"] = bool(
        str(normalized.get("pigProjectId") or "").strip()
        and str(normalized.get("editableProjectFileId") or "").strip()
        and str(normalized.get("editableProjectUrl") or "").strip()
    )
    if normalized.get("qcStatus") == "needs_revision" or normalized.get("handoffStatus") == "rejected":
        normalized["originalGraphicsRequestId"] = str(
            normalized.get("originalGraphicsRequestId") or graphics_request_id
        )
        normalized["revisionOf"] = str(
            normalized.get("revisionOf") or normalized.get("sourceCompletionId") or ""
        )
        normalized["version"] = normalized.get("version") or 2
        normalized["reworkReason"] = str(
            normalized.get("reworkReason") or "correct_and_recreate"
        )
        normalized["previousAssetUrl"] = str(
            normalized.get("previousAssetUrl") or normalized.get("assetUrl") or ""
        )
        normalized["previousAssetPreviewUrl"] = str(
            normalized.get("previousAssetPreviewUrl")
            or normalized.get("assetPreviewUrl")
            or normalized.get("previousAssetUrl")
            or ""
        )
    apply_handoff_queue_contract(normalized)
    return normalized


def handoff_is_actionable(record: dict[str, Any]) -> bool:
    return bool(
        record.get("handoffStatus") in {"requested", "claimed", "rejected"}
        and record.get("pigStatus") not in {"generated", "exported", "uploaded", "failed"}
        and record.get("qcStatus") in {"not_sent", "needs_revision"}
        and (
            str(record.get("quoteText") or "").strip()
            or (
                record.get("contentType") == "FPI"
                and (str(record.get("assetUrl") or "").strip() or has_fpi_source_asset(record.get("sourcePayload") or {}))
            )
        )
    )


def apply_handoff_queue_contract(record: dict[str, Any]) -> None:
    source_payload = record.get("sourcePayload") if isinstance(record.get("sourcePayload"), dict) else {}
    nested_source = extract_nested_payload(source_payload)
    queue_view = str(
        record.get("queueView")
        or source_payload.get("queueView")
        or nested_source.get("queueView")
        or ""
    ).strip().lower()
    source_system = str(record.get("sourceSystem") or "").strip().lower()
    source_status = str(record.get("sourceStatus") or "").strip().lower()
    if str(record.get("qcStatus") or "").strip().lower() == "needs_revision":
        queue_view = "rework"
    elif not queue_view:
        if source_system == "weaver_qc_rework" or source_status.startswith("rework"):
            queue_view = "rework"
        elif source_system == "coverage_needs" or source_status == "coverage_needs":
            queue_view = "coverage_needs"
        else:
            queue_view = "current_titles"

    record["queueView"] = queue_view
    record["isActionable"] = handoff_is_actionable(record)
    record["nextAction"] = "rework" if queue_view == "rework" or record.get("qcStatus") == "needs_revision" else "generate"
    if record["isActionable"]:
        record["statusLabel"] = "Ready for rework" if record["nextAction"] == "rework" else "Ready for P.I.G."
    elif record.get("handoffStatus") == "approved" or record.get("qcStatus") == "approved":
        record["statusLabel"] = "Approved"
    elif record.get("handoffStatus") == "blocked":
        record["statusLabel"] = "Blocked"
    else:
        record["statusLabel"] = str(record.get("handoffStatus") or "")

    for key in (
        "targetCount",
        "approvedCount",
        "pendingQcCount",
        "inProgressCount",
        "reworkCount",
        "remainingApprovedNeeded",
        "remainingActionableNeeded",
        "priorityScore",
    ):
        try:
            record[key] = int(record.get(key) or source_payload.get(key) or nested_source.get(key) or 0)
        except (TypeError, ValueError):
            record[key] = 0
    record["priorityTier"] = str(record.get("priorityTier") or source_payload.get("priorityTier") or nested_source.get("priorityTier") or "")


def append_transition_log(existing_json: str, event: dict[str, Any]) -> str:
    try:
        entries = json.loads(existing_json or "[]")
        if not isinstance(entries, list):
            entries = []
    except json.JSONDecodeError:
        entries = []
    return json.dumps(append_transition_entries(entries, event), ensure_ascii=True, sort_keys=True)


def transition_event_key(event: dict[str, Any]) -> tuple[str, str, str, str, str, str]:
    return (
        str(event.get("event") or ""),
        str(event.get("handoffStatus") or ""),
        str(event.get("pigStatus") or ""),
        str(event.get("qcStatus") or ""),
        str(event.get("blockedReason") or ""),
        str(event.get("claimedBy") or ""),
    )


def append_transition_entries(existing: Any, event: dict[str, Any]) -> list[dict[str, Any]]:
    entries = existing if isinstance(existing, list) else []
    entries = [entry for entry in entries if isinstance(entry, dict)]
    if entries and transition_event_key(entries[-1]) == transition_event_key(event):
        entries[-1] = event
        return entries[-HANDOFF_TRANSITION_LOG_LIMIT:]
    entries.append(event)
    return entries[-HANDOFF_TRANSITION_LOG_LIMIT:]


def compact_handoff_record(record: dict[str, Any]) -> dict[str, Any]:
    compact = dict(record)
    transition_log = compact.get("transitionLog")
    if isinstance(transition_log, list):
        compact["transitionLog"] = [
            entry for entry in transition_log[-HANDOFF_LIST_TRANSITION_LOG_LIMIT:]
            if isinstance(entry, dict)
        ]
    compact.pop("pigPayload", None)
    return compact


def queue_card_record(record: dict[str, Any]) -> dict[str, Any]:
    content_type = normalize_content_type(record.get("contentType") or record.get("imageType"))
    is_rework = str(record.get("nextAction") or "").strip().lower() == "rework"
    source_payload = record.get("sourcePayload") if isinstance(record.get("sourcePayload"), dict) else {}
    return {
        "graphicsRequestId": record.get("graphicsRequestId") or "",
        "queueView": record.get("queueView") or "",
        "isActionable": bool(record.get("isActionable")),
        "nextAction": record.get("nextAction") or "",
        "statusLabel": record.get("statusLabel") or "",
        "handoffStatus": record.get("handoffStatus") or "",
        "pigStatus": record.get("pigStatus") or "",
        "qcStatus": record.get("qcStatus") or "",
        "imageType": content_type,
        "contentType": content_type,
        "queueLane": content_type,
        "reworkLane": content_type if is_rework else "",
        "sourceSystem": record.get("sourceSystem") or "",
        "sourceStatus": record.get("sourceStatus") or "",
        "repairRequestId": source_payload.get("repairRequestId") or "",
        "sourceFlagId": source_payload.get("sourceFlagId") or "",
        "originalContentId": source_payload.get("originalContentId") or "",
        "originalDocId": source_payload.get("originalDocId") or "",
        "originalCollection": source_payload.get("originalCollection") or "",
        "releaseCatalog": source_payload.get("releaseCatalog") or "",
        "bookShortener": source_payload.get("bookShortener") or "",
        "issueReason": source_payload.get("issueReason") or "",
        "repairInstructions": source_payload.get("requestNote") or source_payload.get("requestedChanges") or "",
        "originalAssetLink": source_payload.get("previousAssetUrl") or source_payload.get("assetUrl") or "",
        "originalContent": source_payload.get("originalContent") or {},
        "existingPigHistory": source_payload.get("existingPigHistory") or {},
        "existingWeaverHistory": source_payload.get("existingWeaverHistory") or {},
        "sourceCompletionId": record.get("sourceCompletionId") or "",
        "revisionOf": record.get("revisionOf") or "",
        "originalGraphicsRequestId": record.get("originalGraphicsRequestId") or "",
        "version": record.get("version") or "",
        "pigProjectId": record.get("pigProjectId") or "",
        "editableProjectFileId": record.get("editableProjectFileId") or "",
        "editableProjectUrl": record.get("editableProjectUrl") or "",
        "editableProjectAvailable": bool(record.get("editableProjectAvailable")),
        "editableProjectValidationStatus": record.get("editableProjectValidationStatus") or "",
        "editableProjectValidationError": record.get("editableProjectValidationError") or {},
        "assetUrl": record.get("assetUrl") or "",
        "assetPreviewUrl": record.get("assetPreviewUrl") or "",
        "previousAssetUrl": record.get("previousAssetUrl") or record.get("assetUrl") or "",
        "previousAssetPreviewUrl": record.get("previousAssetPreviewUrl") or record.get("assetPreviewUrl") or "",
        "reworkReason": record.get("reworkReason") or "",
        "rejectReason": record.get("rejectReason") or "",
        "rejectedReason": record.get("rejectedReason") or "",
        "metadataIssue": record.get("metadataIssue") or "",
        "aestheticIssue": record.get("aestheticIssue") or "",
        "qcNote": record.get("qcNote") or "",
        "requestedChanges": record.get("requestedChanges") or "",
        "notes": record.get("notes") or "",
        "qcPayload": record.get("qcPayload") if isinstance(record.get("qcPayload"), dict) else {},
        "queueSheetRow": record.get("queueSheetRow") or "",
        "author": record.get("author") or "",
        "poemTitle": record.get("poemTitle") or "",
        "bookTitle": record.get("bookTitle") or "",
        "bookKey": record.get("bookKey") or "",
        "quoteText": record.get("quoteText") or "",
        "targetCount": int(record.get("targetCount") or 0),
        "approvedCount": int(record.get("approvedCount") or 0),
        "pendingQcCount": int(record.get("pendingQcCount") or 0),
        "inProgressCount": int(record.get("inProgressCount") or 0),
        "reworkCount": int(record.get("reworkCount") or 0),
        "remainingApprovedNeeded": int(record.get("remainingApprovedNeeded") or 0),
        "remainingActionableNeeded": int(record.get("remainingActionableNeeded") or 0),
        "priorityTier": record.get("priorityTier") or "",
        "priorityScore": int(record.get("priorityScore") or 0),
        "createdAt": record.get("createdAt") or "",
        "updatedAt": record.get("updatedAt") or "",
    }


def firestore_encode_value(value: Any) -> dict[str, Any]:
    if value is None:
        return {"nullValue": None}
    if isinstance(value, bool):
        return {"booleanValue": value}
    if isinstance(value, int) and not isinstance(value, bool):
        return {"integerValue": str(value)}
    if isinstance(value, float):
        return {"doubleValue": value}
    if isinstance(value, dict):
        return {"mapValue": {"fields": {str(key): firestore_encode_value(item) for key, item in value.items()}}}
    if isinstance(value, list):
        return {"arrayValue": {"values": [firestore_encode_value(item) for item in value]}}
    return {"stringValue": str(value)}


def firestore_decode_value(value: dict[str, Any]) -> Any:
    if "nullValue" in value:
        return None
    if "booleanValue" in value:
        return value["booleanValue"]
    if "integerValue" in value:
        try:
            return int(value["integerValue"])
        except (TypeError, ValueError):
            return value["integerValue"]
    if "doubleValue" in value:
        return value["doubleValue"]
    if "stringValue" in value:
        return value["stringValue"]
    if "timestampValue" in value:
        return value["timestampValue"]
    if "mapValue" in value:
        return {
            key: firestore_decode_value(item)
            for key, item in (value.get("mapValue", {}).get("fields") or {}).items()
        }
    if "arrayValue" in value:
        return [firestore_decode_value(item) for item in value.get("arrayValue", {}).get("values", [])]
    return None


def firestore_stable_document_id(prefix: str, payload: dict[str, Any]) -> str:
    raw = json.dumps(payload, ensure_ascii=True, sort_keys=True, default=str)
    digest = hashlib.sha1(raw.encode("utf-8")).hexdigest()[:20]
    return f"{prefix}-{digest}"


def firestore_access_token(project_id: str) -> str:
    global _FIRESTORE_ACCESS_TOKEN_CACHE
    env_token = os.environ.get("GOOGLE_OAUTH_ACCESS_TOKEN", "").strip()
    if env_token:
        return env_token

    now = datetime.now(UTC).timestamp()
    if _FIRESTORE_ACCESS_TOKEN_CACHE and _FIRESTORE_ACCESS_TOKEN_CACHE[1] > now:
        return _FIRESTORE_ACCESS_TOKEN_CACHE[0]

    metadata_request = urllib.request.Request(
        "http://metadata.google.internal/computeMetadata/v1/instance/service-accounts/default/token",
        headers={"Metadata-Flavor": "Google"},
    )
    try:
        with urllib.request.urlopen(metadata_request, timeout=2) as response:
            payload = json.loads(response.read().decode("utf-8"))
            token = str(payload.get("access_token") or "").strip()
            if token:
                ttl = max(60, int(payload.get("expires_in") or 300) - 60)
                _FIRESTORE_ACCESS_TOKEN_CACHE = (token, now + ttl)
                return token
    except Exception:
        pass

    service_account = (
        os.environ.get("WEAVER_FIRESTORE_SERVICE_ACCOUNT")
        or FIRESTORE_DEFAULT_SERVICE_ACCOUNT
    ).strip()
    try:
        result = subprocess.run(
            [
                "gcloud",
                "auth",
                "print-access-token",
                service_account,
                f"--project={project_id}",
            ],
            check=True,
            capture_output=True,
            text=True,
            timeout=10,
        )
        token = result.stdout.strip()
        if token:
            _FIRESTORE_ACCESS_TOKEN_CACHE = (token, now + 240)
            return token
    except Exception:
        pass

    raise RuntimeError(
        "Unable to obtain a Firestore access token from Cloud Run metadata or "
        f"the configured Weaver service account ({service_account})"
    )


def firestore_ssl_context() -> ssl.SSLContext:
    cafile = os.environ.get("SSL_CERT_FILE", "").strip()
    if cafile:
        return ssl.create_default_context(cafile=cafile)
    for candidate in (
        "/etc/ssl/cert.pem",
        "/opt/homebrew/etc/ca-certificates/cert.pem",
        "/usr/local/etc/openssl@3/cert.pem",
    ):
        if Path(candidate).exists():
            return ssl.create_default_context(cafile=candidate)
    return ssl.create_default_context()


class FirestoreLedgerClient:
    def __init__(self, project_id: str, database_id: str = FIRESTORE_DEFAULT_DATABASE_ID) -> None:
        self.project_id = project_id
        self.database_id = database_id or FIRESTORE_DEFAULT_DATABASE_ID

    def close(self) -> None:
        return None

    @property
    def documents_url(self) -> str:
        project = urllib.parse.quote(self.project_id, safe="")
        database = urllib.parse.quote(self.database_id, safe="()")
        return f"https://firestore.googleapis.com/v1/projects/{project}/databases/{database}/documents"

    @property
    def collection_url(self) -> str:
        return f"{self.documents_url}/{FIRESTORE_COLLECTION}"

    def collection_url_for(self, collection: str) -> str:
        return f"{self.documents_url}/{urllib.parse.quote(collection, safe='')}"

    def document_url(self, graphics_request_id: str) -> str:
        return f"{self.collection_url}/{urllib.parse.quote(graphics_request_id, safe='')}"

    def document_url_for(self, collection: str, document_id: str) -> str:
        return f"{self.collection_url_for(collection)}/{urllib.parse.quote(document_id, safe='')}"

    def request(self, method: str, url: str, body: dict[str, Any] | None = None) -> Any:
        data = json.dumps(body).encode("utf-8") if body is not None else None
        request = urllib.request.Request(
            url,
            data=data,
            method=method,
            headers={
                "Authorization": f"Bearer {firestore_access_token(self.project_id)}",
                "Content-Type": "application/json",
            },
        )
        try:
            with urllib.request.urlopen(request, timeout=20, context=firestore_ssl_context()) as response:
                raw = response.read().decode("utf-8")
                return json.loads(raw) if raw else {}
        except urllib.error.HTTPError as exc:
            if exc.code == 404:
                return None
            detail = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"Firestore {method} failed with HTTP {exc.code}: {detail}") from exc

    def decode_document(self, document: dict[str, Any] | None) -> dict[str, Any] | None:
        if not document:
            return None
        record = {
            key: firestore_decode_value(value)
            for key, value in (document.get("fields") or {}).items()
        }
        return normalize_handoff_record(record)

    def decode_raw_document(self, document: dict[str, Any] | None) -> dict[str, Any] | None:
        if not document:
            return None
        return {
            key: firestore_decode_value(value)
            for key, value in (document.get("fields") or {}).items()
        }

    def write_raw_document(self, collection: str, document_id: str, record: dict[str, Any]) -> dict[str, Any]:
        body = {
            "fields": {
                key: firestore_encode_value(value)
                for key, value in record.items()
            }
        }
        document = self.request("PATCH", self.document_url_for(collection, document_id), body)
        return self.decode_raw_document(document) or record

    def get_raw_document(self, collection: str, document_id: str) -> dict[str, Any] | None:
        document = self.request("GET", self.document_url_for(collection, document_id))
        return self.decode_raw_document(document)

    def list_raw_documents(self, collection: str, page_size: int = 500) -> list[dict[str, Any]]:
        records: list[dict[str, Any]] = []
        page_token = ""
        while True:
            query = urllib.parse.urlencode({
                "pageSize": str(max(1, min(int(page_size or 500), 500))),
                **({"pageToken": page_token} if page_token else {}),
            })
            response = self.request("GET", f"{self.collection_url_for(collection)}?{query}") or {}
            for document in response.get("documents", []):
                record = self.decode_raw_document(document)
                if record:
                    records.append(record)
            page_token = str(response.get("nextPageToken") or "")
            if not page_token:
                break
        return records

    def query_raw_documents(
        self,
        collection: str,
        field: str,
        value: Any,
        page_size: int = 100,
    ) -> list[dict[str, Any]]:
        records: list[dict[str, Any]] = []
        cursor_name = ""
        limit = max(1, min(int(page_size or 100), 500))
        while True:
            query: dict[str, Any] = {
                "from": [{"collectionId": collection}],
                "where": {
                    "fieldFilter": {
                        "field": {"fieldPath": field},
                        "op": "EQUAL",
                        "value": firestore_encode_value(value),
                    }
                },
                "orderBy": [{"field": {"fieldPath": "__name__"}, "direction": "ASCENDING"}],
                "limit": limit,
            }
            if cursor_name:
                query["startAt"] = {
                    "values": [{"referenceValue": cursor_name}],
                    "before": False,
                }
            response = self.request("POST", f"{self.documents_url}:runQuery", {"structuredQuery": query}) or []
            documents = [entry.get("document") for entry in response if entry.get("document")]
            for document in documents:
                record = self.decode_raw_document(document)
                if record:
                    records.append(record)
            if len(documents) < limit:
                break
            cursor_name = str(documents[-1].get("name") or "")
            if not cursor_name:
                break
        return records

    def write_record(self, record: dict[str, Any]) -> dict[str, Any]:
        normalized = normalize_handoff_record(record)
        transition_log = normalized.get("transitionLog")
        if isinstance(transition_log, list):
            normalized["transitionLog"] = [
                entry for entry in transition_log[-HANDOFF_TRANSITION_LOG_LIMIT:]
                if isinstance(entry, dict)
            ]
        body = {
            "fields": {
                key: firestore_encode_value(value)
                for key, value in normalized.items()
            }
        }
        document = self.request("PATCH", self.document_url(normalized["graphicsRequestId"]), body)
        return self.decode_document(document) or normalized

    def get_handoff(self, graphics_request_id: str) -> dict[str, Any] | None:
        graphics_request_id = str(graphics_request_id or "").strip()
        if not graphics_request_id:
            return None
        alias = self.get_raw_document(
            FIRESTORE_HANDOFF_ALIASES_COLLECTION,
            hashlib.sha256(graphics_request_id.encode("utf-8")).hexdigest(),
        )
        canonical_request_id = str((alias or {}).get("canonicalGraphicsRequestId") or "").strip()
        if canonical_request_id and canonical_request_id != graphics_request_id:
            canonical = self.decode_document(
                self.request("GET", self.document_url(canonical_request_id))
            )
            if canonical:
                return canonical
        return self.decode_document(self.request("GET", self.document_url(graphics_request_id)))

    def get_handoffs(self, graphics_request_ids: list[str]) -> list[dict[str, Any]]:
        records: list[dict[str, Any]] = []
        for request_id in graphics_request_ids:
            record = self.get_handoff(request_id)
            if record:
                records.append(compact_handoff_record(record))
        return records

    def list_handoffs(self, page_size: int = 500) -> list[dict[str, Any]]:
        records: list[dict[str, Any]] = []
        page_token = ""
        while True:
            query = urllib.parse.urlencode({
                "pageSize": str(max(1, min(int(page_size or 500), 500))),
                **({"pageToken": page_token} if page_token else {}),
            })
            response = self.request("GET", f"{self.collection_url}?{query}") or {}
            for document in response.get("documents", []):
                record = self.decode_document(document)
                if record:
                    records.append(compact_handoff_record(record))
            page_token = str(response.get("nextPageToken") or "")
            if not page_token:
                break
        return records

    def upsert_handoff_request(self, request: dict[str, Any]) -> dict[str, Any]:
        graphics_request_id = str(request.get("graphicsRequestId") or request.get("id") or "").strip()
        if not graphics_request_id:
            raise ValueError("graphicsRequestId is required")
        supplied_request_id = graphics_request_id

        now = utc_now_iso()
        existing = self.get_handoff(graphics_request_id)
        source_payload = request.get("sourcePayload") or request.get("payload") or request
        content_type = infer_graphics_content_type(request, source_payload, existing)
        assert_compatible_graphics_content_type(existing, content_type)
        repair_request_id = str(
            request.get("repairRequestId")
            or extract_handoff_value(source_payload, "repairRequestId")
            or ""
        ).strip()
        is_external_repair_request = bool(
            repair_request_id
            and str(request.get("sourceSystem") or "").strip().lower() == "poetry_please_repair"
        )
        canonical_content_id = ""
        if is_external_repair_request:
            canonical_content_id = f"REPAIR:{repair_request_id}"
        elif content_type == "QI":
            canonical_content_id = canonical_qi_content_id(request, source_payload)
            canonical_matches = self.query_raw_documents(
                FIRESTORE_COLLECTION,
                "canonicalContentId",
                canonical_content_id,
                page_size=2,
            )
            if len(canonical_matches) > 1:
                raise ValueError(f"Duplicate QI canonical identity: {canonical_content_id}")
            if canonical_matches:
                canonical_record = normalize_handoff_record(canonical_matches[0])
                graphics_request_id = str(canonical_record.get("graphicsRequestId") or graphics_request_id)
                existing = canonical_record
            elif not existing:
                graphics_request_id = canonical_qi_graphics_request_id(canonical_content_id)
                existing = self.get_handoff(graphics_request_id)
        elif content_type == "FPI":
            canonical_content_id = canonical_fpi_content_id(request, source_payload)
            canonical_matches = self.query_raw_documents(
                FIRESTORE_COLLECTION,
                "canonicalContentId",
                canonical_content_id,
                page_size=2,
            )
            if len(canonical_matches) > 1:
                raise ValueError(f"Duplicate FPI canonical identity: {canonical_content_id}")
            if canonical_matches:
                canonical_record = normalize_handoff_record(canonical_matches[0])
                graphics_request_id = str(canonical_record.get("graphicsRequestId") or graphics_request_id)
                existing = canonical_record
            elif not existing:
                graphics_request_id = canonical_fpi_graphics_request_id(canonical_content_id)
                existing = self.get_handoff(graphics_request_id)
        existing = existing or default_handoff_record(graphics_request_id)
        image_type = normalize_content_type(
            request.get("imageType")
            or request.get("contentType")
            or extract_handoff_value(source_payload, "imageType", "contentType"),
            content_type,
        )
        if image_type == "FP":
            image_type = "FPI"
        has_text = bool(extract_handoff_text(source_payload))
        has_required_source = has_text or (content_type == "FPI" and has_fpi_source_asset(source_payload))
        if not has_required_source and existing.get("sourcePayload"):
            source_payload = existing["sourcePayload"]
            has_text = bool(extract_handoff_text(source_payload))
            has_required_source = has_text or (content_type == "FPI" and has_fpi_source_asset(source_payload))
        handoff_status = normalize_enum(request.get("handoffStatus"), HANDOFF_STATUSES, "requested")
        pig_status = normalize_enum(request.get("pigStatus"), PIG_STATUSES, "not_started")
        qc_status = normalize_enum(request.get("qcStatus"), QC_STATUSES, "not_sent")
        source_status = str(request.get("sourceStatus") or "needs_graphics")
        blocked_reason = str(request.get("blockedReason") or "")
        existing_handoff_status = str(existing.get("handoffStatus") or "")
        is_rework_request = bool(
            is_external_repair_request
            or
            request.get("revisionOf")
            or request.get("originalGraphicsRequestId")
            or extract_handoff_value(source_payload, "revisionOf", "originalGraphicsRequestId")
            or source_status.lower().startswith(("rework", "manual_rework"))
        )
        terminal_source_refresh = False
        if (
            not is_rework_request
            and existing_handoff_status in TERMINAL_HANDOFF_STATUSES
            and handoff_status in SOURCE_QUEUE_REFRESH_STATUSES
        ):
            terminal_source_refresh = True
            source_status = str(existing.get("sourceStatus") or source_status)
            handoff_status = existing_handoff_status
            pig_status = normalize_enum(existing.get("pigStatus"), PIG_STATUSES, pig_status)
            qc_status = normalize_enum(existing.get("qcStatus"), QC_STATUSES, qc_status)
            blocked_reason = str(existing.get("blockedReason") or blocked_reason)
        elif not has_required_source and handoff_status in {"requested", "claimed"}:
            handoff_status = "blocked"
            pig_status = "failed"
            blocked_reason = blocked_reason or ("blank_fpi_source_asset" if content_type == "FPI" else "blank_request_text")

        record = {
            **existing,
            "graphicsRequestId": graphics_request_id,
            "sourceSystem": str(request.get("sourceSystem") or "weaver"),
            "sourceStatus": source_status,
            "contentType": content_type,
            "imageType": image_type,
            "canonicalContentId": canonical_content_id or existing.get("canonicalContentId") or "",
            "sourceCompletionId": str(request.get("sourceCompletionId") or extract_handoff_value(source_payload, "sourceCompletionId") or ""),
            "revisionOf": str(request.get("revisionOf") or extract_handoff_value(source_payload, "revisionOf") or ""),
            "originalGraphicsRequestId": str(request.get("originalGraphicsRequestId") or extract_handoff_value(source_payload, "originalGraphicsRequestId") or ""),
            "reviewStatus": str(request.get("reviewStatus") or extract_handoff_value(source_payload, "reviewStatus") or ""),
            "ocrText": str(request.get("ocrText") or extract_handoff_value(source_payload, "ocrText") or ""),
            "version": str(request.get("version") or extract_handoff_value(source_payload, "version") or existing.get("version") or ""),
            "assetFileId": first_durable_value(request.get("assetFileId"), request.get("driveFileId"), request.get("fileId"), extract_handoff_value(source_payload, "assetFileId", "driveFileId", "fileId"), existing.get("assetFileId"), existing.get("driveFileId")),
            "driveFileId": first_durable_value(request.get("driveFileId"), request.get("assetFileId"), request.get("fileId"), extract_handoff_value(source_payload, "driveFileId", "assetFileId", "fileId"), existing.get("driveFileId"), existing.get("assetFileId")),
            "assetUrl": first_durable_value(request.get("assetUrl"), request.get("assetLinkUrl"), request.get("driveUrl"), extract_handoff_value(source_payload, "assetUrl", "assetLinkUrl", "driveUrl"), existing.get("assetUrl")),
            "assetPreviewUrl": first_durable_value(request.get("assetPreviewUrl"), request.get("previewUrl"), request.get("thumbnailUrl"), extract_handoff_value(source_payload, "assetPreviewUrl", "previewUrl", "thumbnailUrl"), existing.get("assetPreviewUrl")),
            "pigProjectId": first_durable_value(request.get("pigProjectId"), extract_handoff_value(source_payload, "pigProjectId", "pig_project_id"), existing.get("pigProjectId")),
            "editableProjectFileId": first_durable_value(request.get("editableProjectFileId"), request.get("projectFileId"), extract_handoff_value(source_payload, "editableProjectFileId", "projectFileId", "editable_project_file_id"), existing.get("editableProjectFileId")),
            "editableProjectUrl": first_durable_value(request.get("editableProjectUrl"), extract_handoff_value(source_payload, "editableProjectUrl", "editable_project_url"), existing.get("editableProjectUrl")),
            "editableProjectKind": first_durable_value(request.get("editableProjectKind"), extract_handoff_value(source_payload, "editableProjectKind", "editable_project_kind"), existing.get("editableProjectKind")),
            "editableProjectSchemaVersion": first_durable_value(request.get("editableProjectSchemaVersion"), extract_handoff_value(source_payload, "editableProjectSchemaVersion", "editable_project_schema_version"), existing.get("editableProjectSchemaVersion")),
            "reworkReason": str(request.get("reworkReason") or extract_handoff_value(source_payload, "reworkReason", "rejectReason", "rejectedReason") or existing.get("reworkReason") or ""),
            "metadataIssue": str(request.get("metadataIssue") or extract_handoff_value(source_payload, "metadataIssue", "metadata_issue") or existing.get("metadataIssue") or ""),
            "aestheticIssue": str(request.get("aestheticIssue") or extract_handoff_value(source_payload, "aestheticIssue", "aesthetic_issue") or existing.get("aestheticIssue") or ""),
            "qcNote": str(request.get("qcNote") or extract_handoff_value(source_payload, "qcNote", "graphicsQcNote", "qc_note") or existing.get("qcNote") or ""),
            "requestedChanges": str(request.get("requestedChanges") or extract_handoff_value(source_payload, "requestedChanges", "requested_changes") or existing.get("requestedChanges") or ""),
            "pigStatus": pig_status,
            "handoffStatus": handoff_status,
            "qcStatus": qc_status,
            "sourcePayload": source_payload,
            "blockedReason": blocked_reason,
            "createdAt": existing.get("createdAt") or now,
            "updatedAt": now,
            "transitionLog": existing.get("transitionLog") if terminal_source_refresh else append_transition_entries(existing.get("transitionLog"), {
                "at": now,
                "event": "request_upserted",
                "handoffStatus": handoff_status,
                "pigStatus": pig_status,
                "qcStatus": qc_status,
                "blockedReason": blocked_reason,
            }),
        }
        written = self.write_record(record)
        if supplied_request_id != graphics_request_id:
            self.write_raw_document(
                FIRESTORE_HANDOFF_ALIASES_COLLECTION,
                hashlib.sha256(supplied_request_id.encode("utf-8")).hexdigest(),
                {
                    "graphicsRequestId": supplied_request_id,
                    "canonicalGraphicsRequestId": graphics_request_id,
                    "canonicalContentId": canonical_content_id,
                    "updatedAt": now,
                },
            )
        return written

    def claim_handoff(self, graphics_request_id: str, claimed_by: str = "") -> dict[str, Any]:
        existing = self.get_handoff(graphics_request_id)
        if not existing:
            raise KeyError(f"Unknown graphicsRequestId: {graphics_request_id}")
        if existing["handoffStatus"] in {"generated", "exported", "uploaded", "sent_to_weaver_qc", "approved", "blocked", "errored"}:
            return existing

        now = utc_now_iso()
        existing.update({
            "pigStatus": "claimed",
            "handoffStatus": "claimed",
            "claimedBy": claimed_by,
            "claimedAt": existing.get("claimedAt") or now,
            "updatedAt": now,
            "transitionLog": append_transition_entries(existing.get("transitionLog"), {"at": now, "event": "claimed", "claimedBy": claimed_by}),
        })
        return self.write_record(existing)

    def update_handoff(self, graphics_request_id: str, update: dict[str, Any]) -> dict[str, Any]:
        existing = self.get_handoff(graphics_request_id)
        if not existing:
            raise KeyError(f"Unknown graphicsRequestId: {graphics_request_id}")

        now = utc_now_iso()
        handoff_status = normalize_enum(update.get("handoffStatus"), HANDOFF_STATUSES, existing["handoffStatus"])
        pig_status = normalize_enum(update.get("pigStatus"), PIG_STATUSES, existing["pigStatus"])
        qc_status = normalize_enum(update.get("qcStatus"), QC_STATUSES, existing["qcStatus"])
        content_type = normalize_content_type(update.get("contentType") or update.get("imageType") or existing.get("contentType"), existing.get("contentType") or "QI")
        image_type = normalize_content_type(update.get("imageType") or update.get("contentType") or existing.get("imageType"), content_type)
        if content_type == "FP":
            content_type = "FPI"
        if image_type == "FP":
            image_type = "FPI"
        asset_url = first_durable_value(update.get("assetUrl"), update.get("assetLinkUrl"), update.get("driveUrl"), existing.get("assetUrl"))
        uploaded = handoff_status in {"uploaded", "sent_to_weaver_qc", "approved"} or pig_status == "uploaded"
        generated = uploaded or handoff_status in {"generated", "exported", "sent_to_weaver_qc", "approved"} or pig_status in {"generated", "exported", "uploaded"}
        sent_to_qc = handoff_status in {"sent_to_weaver_qc", "approved", "rejected"} or qc_status in {"pending", "approved", "rejected", "needs_revision"}
        approved = handoff_status == "approved" or qc_status == "approved"
        rejected = handoff_status == "rejected" or qc_status in {"rejected", "needs_revision"}

        existing.update({
            "sourceStatus": str(update.get("sourceStatus") or existing["sourceStatus"]),
            "contentType": content_type,
            "imageType": image_type,
            "sourceCompletionId": str(update.get("sourceCompletionId") or existing.get("sourceCompletionId") or ""),
            "revisionOf": str(update.get("revisionOf") or existing.get("revisionOf") or ""),
            "originalGraphicsRequestId": str(update.get("originalGraphicsRequestId") or existing.get("originalGraphicsRequestId") or ""),
            "reviewStatus": str(update.get("reviewStatus") or existing.get("reviewStatus") or ""),
            "ocrText": str(update.get("ocrText") or existing.get("ocrText") or ""),
            "pigStatus": pig_status,
            "handoffStatus": handoff_status,
            "qcStatus": qc_status,
            "assetUrl": asset_url,
            "assetPreviewUrl": first_durable_value(update.get("assetPreviewUrl"), update.get("previewUrl"), update.get("thumbnailUrl"), existing.get("assetPreviewUrl")),
            "assetFileId": first_durable_value(update.get("assetFileId"), update.get("driveFileId"), update.get("fileId"), existing.get("assetFileId"), existing.get("driveFileId")),
            "driveFileId": first_durable_value(update.get("driveFileId"), update.get("assetFileId"), update.get("fileId"), existing.get("driveFileId"), existing.get("assetFileId")),
            "driveFileName": str(update.get("driveFileName") or update.get("fileName") or existing["driveFileName"] or ""),
            "mimeType": str(update.get("mimeType") or existing["mimeType"] or ""),
            "exportType": str(update.get("exportType") or existing["exportType"] or ""),
            "variant": str(update.get("variant") or existing["variant"] or ""),
            "version": str(update.get("version") or existing["version"] or ""),
            "pigProjectId": first_durable_value(update.get("pigProjectId"), existing.get("pigProjectId")),
            "editableProjectFileId": first_durable_value(update.get("editableProjectFileId"), update.get("projectFileId"), existing.get("editableProjectFileId")),
            "editableProjectUrl": first_durable_value(update.get("editableProjectUrl"), existing.get("editableProjectUrl")),
            "editableProjectKind": first_durable_value(update.get("editableProjectKind"), existing.get("editableProjectKind")),
            "editableProjectSchemaVersion": first_durable_value(update.get("editableProjectSchemaVersion"), existing.get("editableProjectSchemaVersion")),
            "candidatePigProjectId": str(
                update.get("candidatePigProjectId") or existing.get("candidatePigProjectId") or ""
            ),
            "candidateEditableProjectFileId": str(
                update.get("candidateEditableProjectFileId")
                or existing.get("candidateEditableProjectFileId")
                or ""
            ),
            "candidateEditableProjectUrl": str(
                update.get("candidateEditableProjectUrl")
                or existing.get("candidateEditableProjectUrl")
                or ""
            ),
            "editableProjectValidationStatus": str(
                update.get("editableProjectValidationStatus")
                or existing.get("editableProjectValidationStatus")
                or ""
            ),
            "editableProjectValidationError": (
                update.get("editableProjectValidationError")
                or existing.get("editableProjectValidationError")
                or {}
            ),
            "reworkReason": str(update.get("reworkReason") or update.get("rejectReason") or existing.get("reworkReason") or ""),
            "metadataIssue": str(update.get("metadataIssue") or existing.get("metadataIssue") or ""),
            "aestheticIssue": str(update.get("aestheticIssue") or existing.get("aestheticIssue") or ""),
            "qcNote": str(update.get("qcNote") or existing.get("qcNote") or ""),
            "requestedChanges": str(update.get("requestedChanges") or existing.get("requestedChanges") or ""),
            "errorMessage": str(update.get("errorMessage") or existing["errorMessage"] or ""),
            "blockedReason": str(update.get("blockedReason") or existing["blockedReason"] or ""),
            "pigPayload": update.get("pigPayload") or update,
            "qcPayload": update.get("qcPayload") or update,
            "updatedAt": now,
            "generatedAt": existing.get("generatedAt") or (now if generated else ""),
            "uploadedAt": existing.get("uploadedAt") or (now if uploaded else ""),
            "sentToQcAt": existing.get("sentToQcAt") or (now if sent_to_qc else ""),
            "approvedAt": existing.get("approvedAt") or (now if approved else ""),
            "rejectedAt": existing.get("rejectedAt") or (now if rejected else ""),
            "transitionLog": append_transition_entries(existing.get("transitionLog"), {
                "at": now,
                "event": "updated",
                "handoffStatus": handoff_status,
                "pigStatus": pig_status,
                "qcStatus": qc_status,
            }),
        })
        existing["editableProjectAvailable"] = bool(
            existing.get("pigProjectId")
            and existing.get("editableProjectFileId")
            and existing.get("editableProjectUrl")
        )
        return self.write_record(existing)

    def get_handoff_queue(
        self,
        limit: int = 100,
        filter_mode: str = "all",
        cursor: int = 0,
        content_type: str = "",
    ) -> list[dict[str, Any]]:
        normalized_filter = str(filter_mode or "all").strip().lower()
        normalized_content_type = str(content_type or "").strip().upper()
        if normalized_content_type not in {"", "QI", "FPI"}:
            raise ValueError("contentType must be QI or FPI")
        offset = max(0, int(cursor or 0))
        records = [
            record for record in self.list_handoffs()
            if record.get("isActionable")
            and (
                not normalized_content_type
                or normalize_content_type(record.get("contentType") or record.get("imageType"))
                == normalized_content_type
            )
            and (
                normalized_filter in {"all", ""}
                or record.get("queueView") == normalized_filter
                or (normalized_filter == "rework" and record.get("nextAction") == "rework")
            )
        ]
        records.sort(key=lambda item: str(item.get("createdAt") or ""))
        page_limit = max(1, min(int(limit or 100), 500))
        return [queue_card_record(record) for record in records[offset:offset + page_limit]]

    def upsert_graphics_request(self, request: dict[str, Any]) -> str:
        request_id = str(request["id"]).strip()
        if not request_id:
            raise ValueError("graphics request id is required")
        created_at = str(request.get("created_at") or request.get("createdAt") or utc_now_iso())
        updated_at = str(request.get("updated_at") or request.get("updatedAt") or created_at)
        quote_text = str(request.get("quote_text") or request.get("quoteText") or "")
        content_type = normalize_content_type(request.get("content_type") or request.get("contentType") or request.get("imageType"))
        image_type = normalize_content_type(request.get("image_type") or request.get("imageType") or request.get("contentType"), content_type)
        self.write_raw_document(FIRESTORE_GRAPHICS_REQUESTS_COLLECTION, request_id, {
            "id": request_id,
            "requestStatus": str(request.get("request_status") or request.get("requestStatus") or "OPEN"),
            "sourceType": str(request.get("source_type") or request.get("sourceType") or "weaver_sheet_queue"),
            "contentType": content_type,
            "imageType": image_type,
            "bookTitle": str(request.get("book_title") or request.get("bookTitle") or ""),
            "poemTitle": str(request.get("poem_title") or request.get("poemTitle") or ""),
            "author": str(request.get("author") or ""),
            "quoteText": quote_text,
            "wordCount": int(request.get("word_count") or request.get("wordCount") or count_words(quote_text)),
            "sourceRecordId": str(request.get("source_record_id") or request.get("sourceRecordId") or ""),
            "sourceSheetName": str(request.get("source_sheet_name") or request.get("sourceSheetName") or ""),
            "sourceSheetRow": request.get("source_sheet_row") or request.get("sourceSheetRow") or 0,
            "sourcePayload": request.get("source_payload") or request.get("sourcePayload") or {},
            "latestCompletionId": str(request.get("latest_completion_id") or request.get("latestCompletionId") or ""),
            "createdAt": created_at,
            "updatedAt": updated_at,
        })
        return request_id

    def replace_graphics_request_items(self, graphics_request_id: str, items: list[dict[str, Any]]) -> None:
        request_id = str(graphics_request_id or "").strip()
        if not request_id:
            return
        normalized_items = []
        for position, item in enumerate(items, start=1):
            quote_text = str(item.get("quote_text") or item.get("quoteText") or "")
            normalized_items.append({
                "itemPosition": position,
                "sourceRecordId": str(item.get("source_record_id") or item.get("sourceRecordId") or ""),
                "sourceSheetName": str(item.get("source_sheet_name") or item.get("sourceSheetName") or ""),
                "sourceSheetRow": item.get("source_sheet_row") or item.get("sourceSheetRow") or 0,
                "bookTitle": str(item.get("book_title") or item.get("bookTitle") or ""),
                "poemTitle": str(item.get("poem_title") or item.get("poemTitle") or ""),
                "author": str(item.get("author") or ""),
                "quoteText": quote_text,
                "createdAt": str(item.get("created_at") or item.get("createdAt") or utc_now_iso()),
            })
        self.write_raw_document(FIRESTORE_GRAPHICS_REQUESTS_COLLECTION, request_id, {
            "id": request_id,
            "requestItems": normalized_items,
            "updatedAt": utc_now_iso(),
        })

    def insert_graphics_completion(self, completion: dict[str, Any]) -> str:
        completion_id = str(completion["id"]).strip()
        graphics_request_id = str(completion["graphics_request_id"]).strip()
        if not completion_id or not graphics_request_id:
            raise ValueError("completion id and graphics_request_id are required")
        handoff = self.get_handoff(graphics_request_id)
        if handoff and handoff.get("graphicsRequestId"):
            graphics_request_id = str(handoff["graphicsRequestId"])
        content_type = normalize_content_type(completion.get("content_type") or completion.get("contentType") or completion.get("imageType"))
        image_type = normalize_content_type(completion.get("image_type") or completion.get("imageType") or completion.get("contentType"), content_type)
        ingested_at = str(completion.get("ingested_at") or completion.get("ingestedAt") or utc_now_iso())
        source_payload = completion.get("source_payload") or completion.get("sourcePayload") or {}
        pig_project_id = str(
            completion.get("pig_project_id")
            or completion.get("pigProjectId")
            or extract_handoff_value(source_payload, "pigProjectId", "pig_project_id")
            or ""
        ).strip()
        editable_project_file_id = str(
            completion.get("editable_project_file_id")
            or completion.get("editableProjectFileId")
            or completion.get("projectFileId")
            or extract_handoff_value(
                source_payload,
                "editableProjectFileId",
                "projectFileId",
                "editable_project_file_id",
            )
            or ""
        ).strip()
        editable_project_url = str(
            completion.get("editable_project_url")
            or completion.get("editableProjectUrl")
            or extract_handoff_value(source_payload, "editableProjectUrl", "editable_project_url")
            or ""
        ).strip()
        asset_file_id = first_durable_value(
            completion.get("asset_file_id"), completion.get("assetFileId"),
            completion.get("drive_file_id"), completion.get("driveFileId"), completion.get("fileId"),
            extract_handoff_value(source_payload, "assetFileId", "driveFileId", "fileId"),
        )
        editable_project_kind = first_durable_value(
            completion.get("editable_project_kind"), completion.get("editableProjectKind"),
            extract_handoff_value(source_payload, "editableProjectKind", "editable_project_kind"),
        )
        editable_project_schema_version = first_durable_value(
            completion.get("editable_project_schema_version"), completion.get("editableProjectSchemaVersion"),
            extract_handoff_value(source_payload, "editableProjectSchemaVersion", "editable_project_schema_version"),
        )
        content_id = str(
            completion.get("content_id")
            or completion.get("contentId")
            or completion.get("imageId")
            or extract_handoff_value(source_payload, "contentId", "imageId", "sourceRecordId")
            or ""
        ).strip()
        text_hash = str(
            completion.get("text_hash")
            or completion.get("textHash")
            or extract_handoff_value(source_payload, "textHash")
            or ""
        ).strip()
        if not text_hash:
            quote_text = extract_handoff_value(source_payload, "quoteText", "text")
            text_hash = stable_text_hash(quote_text) if quote_text else ""
        source_identity = completion.get("sourceIdentity") or source_payload.get("sourceIdentity")
        if not isinstance(source_identity, dict):
            source_identity = {}
        source_identity = {
            **source_identity,
            "graphicsRequestId": str(source_identity.get("graphicsRequestId") or graphics_request_id),
            "contentId": str(source_identity.get("contentId") or content_id),
            "imageId": str(source_identity.get("imageId") or content_id),
            "textHash": str(source_identity.get("textHash") or text_hash),
        }
        completion_record = {
            "id": completion_id,
            "graphicsRequestId": graphics_request_id,
            "sourceTool": str(completion.get("source_tool") or completion.get("sourceTool") or "P.I.G."),
            "contentType": content_type,
            "imageType": image_type,
            "assetUrl": str(completion.get("asset_url") or completion.get("assetUrl") or ""),
            "assetPreviewUrl": str(completion.get("asset_preview_url") or completion.get("assetPreviewUrl") or ""),
            "assetFileId": asset_file_id,
            "driveFileId": asset_file_id,
            "pigProjectId": pig_project_id,
            "editableProjectFileId": editable_project_file_id,
            "editableProjectUrl": editable_project_url,
            "editableProjectKind": editable_project_kind,
            "editableProjectSchemaVersion": editable_project_schema_version,
            "editableProjectAvailable": bool(
                pig_project_id and editable_project_file_id and editable_project_url
            ),
            "candidatePigProjectId": str(source_payload.get("candidatePigProjectId") or ""),
            "candidateEditableProjectFileId": str(
                source_payload.get("candidateEditableProjectFileId") or ""
            ),
            "candidateEditableProjectUrl": str(source_payload.get("candidateEditableProjectUrl") or ""),
            "editableProjectValidationStatus": str(
                source_payload.get("editableProjectValidationStatus") or ""
            ),
            "editableProjectValidationError": source_payload.get("editableProjectValidationError") or {},
            "contentId": content_id,
            "imageId": content_id,
            "sourceIdentity": source_identity,
            "textHash": text_hash,
            "productionNotes": str(completion.get("production_notes") or completion.get("productionNotes") or ""),
            "completionStatus": str(completion.get("completion_status") or completion.get("completionStatus") or "RETURNED"),
            "completedAt": str(completion.get("completed_at") or completion.get("completedAt") or utc_now_iso()),
            "ingestedAt": ingested_at,
            "sourcePayload": source_payload,
        }
        assert_completion_identity_consistent(completion_record)
        if editable_project_file_id:
            assert_editable_project_file_not_cross_linked(
                self.query_raw_documents(
                    FIRESTORE_GRAPHICS_COMPLETIONS_COLLECTION,
                    "editableProjectFileId",
                    editable_project_file_id,
                ),
                completion_record,
            )
        pending_candidates = [completion_record]
        for existing in self.query_raw_documents(
            FIRESTORE_GRAPHICS_COMPLETIONS_COLLECTION,
            "graphicsRequestId",
            graphics_request_id,
        ):
            existing_id = str(existing.get("id") or "").strip()
            if not existing_id or existing_id == completion_id:
                continue
            queue_card = self.get_raw_document(FIRESTORE_GRAPHICS_QC_QUEUE_COLLECTION, existing_id) or {}
            if queue_card.get("isPendingQc") is True:
                pending_candidates.append(existing)

        winner = max(
            pending_candidates,
            key=lambda record: (
                str(record.get("completedAt") or ""),
                str(record.get("id") or ""),
            ),
        )
        winner_id = str(winner.get("id") or completion_id)
        if completion_id != winner_id:
            completion_record.update({
                "completionStatus": "SUPERSEDED",
                "supersededBy": winner_id,
                "supersededAt": utc_now_iso(),
            })
        self.write_raw_document(FIRESTORE_GRAPHICS_COMPLETIONS_COLLECTION, completion_id, completion_record)
        self.write_raw_document(
            FIRESTORE_GRAPHICS_QC_QUEUE_COLLECTION,
            completion_id,
            {
                **build_graphics_qc_queue_card(completion_record),
                "isPendingQc": completion_id == winner_id,
                **({
                    "graphicsQcDecision": "SUPERSEDED",
                    "supersededBy": winner_id,
                } if completion_id != winner_id else {}),
            },
        )
        for existing in pending_candidates:
            existing_id = str(existing.get("id") or "").strip()
            if not existing_id or existing_id in {completion_id, winner_id}:
                continue
            superseded_at = utc_now_iso()
            self.write_raw_document(FIRESTORE_GRAPHICS_COMPLETIONS_COLLECTION, existing_id, {
                **existing,
                "completionStatus": "SUPERSEDED",
                "supersededBy": winner_id,
                "supersededAt": superseded_at,
            })
            queue_card = self.get_raw_document(FIRESTORE_GRAPHICS_QC_QUEUE_COLLECTION, existing_id) or {
                "pigCompletionId": existing_id,
                "graphicsRequestId": graphics_request_id,
            }
            self.write_raw_document(FIRESTORE_GRAPHICS_QC_QUEUE_COLLECTION, existing_id, {
                **queue_card,
                "isPendingQc": False,
                "graphicsQcDecision": "SUPERSEDED",
                "supersededBy": winner_id,
                "updatedAt": superseded_at,
            })
        return completion_id

    def insert_graphics_qc_review(self, review: dict[str, Any]) -> int:
        reviewed_at = str(review.get("reviewed_at") or review.get("reviewedAt") or utc_now_iso())
        completion_id = str(review["graphics_completion_id"]).strip()
        record = {
            "graphicsCompletionId": completion_id,
            "decision": str(review["decision"]).strip(),
            "metadataIssue": str(review.get("metadata_issue") or review.get("metadataIssue") or ""),
            "aestheticIssue": str(review.get("aesthetic_issue") or review.get("aestheticIssue") or ""),
            "note": str(review.get("note") or ""),
            "reviewedBy": str(review.get("reviewed_by") or review.get("reviewedBy") or ""),
            "reviewedAt": reviewed_at,
            "sourcePayload": review.get("source_payload") or review.get("sourcePayload") or {},
        }
        document_id = firestore_stable_document_id("qc", record)
        self.write_raw_document(FIRESTORE_GRAPHICS_QC_REVIEWS_COLLECTION, document_id, {**record, "id": document_id})
        queue_card = self.get_raw_document(FIRESTORE_GRAPHICS_QC_QUEUE_COLLECTION, completion_id) or {
            "pigCompletionId": completion_id,
        }
        self.write_raw_document(FIRESTORE_GRAPHICS_QC_QUEUE_COLLECTION, completion_id, {
            **queue_card,
            "isPendingQc": False,
            "graphicsQcDecision": record["decision"],
            "graphicsQcNote": record["note"],
            "graphicsQcUpdatedAt": reviewed_at,
            "updatedAt": reviewed_at,
        })
        return 1

    def insert_poetry_please_handoff(self, handoff: dict[str, Any]) -> int:
        handed_off_at = str(handoff.get("handed_off_at") or handoff.get("handedOffAt") or utc_now_iso())
        completion_id = str(handoff["graphics_completion_id"]).strip()
        record = {
            "graphicsCompletionId": completion_id,
            "handoffStatus": str(handoff.get("handoff_status") or handoff.get("handoffStatus") or "QUEUED"),
            "handoffMode": str(handoff.get("handoff_mode") or handoff.get("handoffMode") or "preview"),
            "handedOffAt": handed_off_at,
            "poetryPleaseItemId": str(handoff.get("poetry_please_item_id") or handoff.get("poetryPleaseItemId") or ""),
            "payload": handoff.get("payload") or {},
        }
        document_id = firestore_stable_document_id("pph", record)
        self.write_raw_document(FIRESTORE_POETRY_PLEASE_HANDOFFS_COLLECTION, document_id, {**record, "id": document_id})
        return 1

    def get_latest_graphics_qc_reviews(self, completion_ids: list[str] | None = None) -> dict[str, dict[str, Any]]:
        allowed = {str(value or "").strip() for value in (completion_ids or []) if str(value or "").strip()}
        latest: dict[str, dict[str, Any]] = {}
        for record in self.list_raw_documents(FIRESTORE_GRAPHICS_QC_REVIEWS_COLLECTION):
            completion_id = str(record.get("graphicsCompletionId") or "")
            if allowed and completion_id not in allowed:
                continue
            existing = latest.get(completion_id)
            if existing and str(existing.get("reviewed_at") or "") >= str(record.get("reviewedAt") or ""):
                continue
            latest[completion_id] = {
                "decision": str(record.get("decision") or ""),
                "metadata_issue": str(record.get("metadataIssue") or ""),
                "aesthetic_issue": str(record.get("aestheticIssue") or ""),
                "note": str(record.get("note") or ""),
                "reviewed_by": str(record.get("reviewedBy") or ""),
                "reviewed_at": str(record.get("reviewedAt") or ""),
            }
        return latest

    def get_latest_poetry_please_handoffs(self, completion_ids: list[str] | None = None) -> dict[str, dict[str, Any]]:
        allowed = {str(value or "").strip() for value in (completion_ids or []) if str(value or "").strip()}
        latest: dict[str, dict[str, Any]] = {}
        for record in self.list_raw_documents(FIRESTORE_POETRY_PLEASE_HANDOFFS_COLLECTION):
            completion_id = str(record.get("graphicsCompletionId") or "")
            if allowed and completion_id not in allowed:
                continue
            existing = latest.get(completion_id)
            if existing and str(existing.get("handed_off_at") or "") >= str(record.get("handedOffAt") or ""):
                continue
            latest[completion_id] = {
                "handoff_status": str(record.get("handoffStatus") or ""),
                "handoff_mode": str(record.get("handoffMode") or ""),
                "handed_off_at": str(record.get("handedOffAt") or ""),
                "poetry_please_item_id": str(record.get("poetryPleaseItemId") or ""),
                "payload_json": json.dumps(record.get("payload") or {}, ensure_ascii=True, sort_keys=True),
            }
        return latest


def connect_firestore_ledger(
    project_id: str | None = None,
    database_id: str | None = None,
) -> FirestoreLedgerClient:
    resolved_project_id = (
        project_id
        or os.environ.get("WEAVER_FIRESTORE_PROJECT_ID")
        or FIRESTORE_DEFAULT_PROJECT_ID
    ).strip()
    if not resolved_project_id:
        raise RuntimeError("WEAVER_FIRESTORE_PROJECT_ID is required for Firestore ledger backend")
    resolved_database_id = (database_id or os.environ.get("WEAVER_FIRESTORE_DATABASE_ID") or FIRESTORE_DEFAULT_DATABASE_ID).strip()
    return FirestoreLedgerClient(resolved_project_id, resolved_database_id)


def is_firestore_ledger(connection: Any) -> bool:
    return isinstance(connection, FirestoreLedgerClient)



def get_graphics_handoff(connection: Any, graphics_request_id: str) -> dict[str, Any] | None:
    if is_firestore_ledger(connection):
        return connection.get_handoff(graphics_request_id)
    row = connection.execute(
        "SELECT * FROM graphics_handoff_ledger WHERE graphics_request_id = ?",
        (graphics_request_id,),
    ).fetchone()
    return row_to_handoff(row) if row else None


def get_graphics_handoffs(connection: Any, graphics_request_ids: list[str]) -> list[dict[str, Any]]:
    ids = [str(value or "").strip() for value in graphics_request_ids if str(value or "").strip()]
    if not ids:
        return []
    if is_firestore_ledger(connection):
        return connection.get_handoffs(ids)
    placeholders = ",".join("?" for _ in ids)
    rows = connection.execute(
        f"SELECT * FROM graphics_handoff_ledger WHERE graphics_request_id IN ({placeholders})",
        ids,
    ).fetchall()
    return [row_to_handoff(row) for row in rows]


def upsert_graphics_handoff_request(connection: Any, request: dict[str, Any]) -> dict[str, Any]:
    if is_firestore_ledger(connection):
        return connection.upsert_handoff_request(request)
    graphics_request_id = str(request.get("graphicsRequestId") or request.get("id") or "").strip()
    if not graphics_request_id:
        raise ValueError("graphicsRequestId is required")

    now = utc_now_iso()
    existing = connection.execute(
        "SELECT * FROM graphics_handoff_ledger WHERE graphics_request_id = ?",
        (graphics_request_id,),
    ).fetchone()
    source_payload = request.get("sourcePayload") or request.get("payload") or request
    existing_record = row_to_handoff(existing) if existing else None
    content_type = infer_graphics_content_type(request, source_payload, existing_record)
    assert_compatible_graphics_content_type(existing_record, content_type)
    if content_type == "FPI":
        canonical_content_id = canonical_fpi_content_id(request, source_payload)
        canonical_matches = []
        for candidate in connection.execute(
            "SELECT * FROM graphics_handoff_ledger WHERE content_type IN ('FP', 'FPI')"
        ).fetchall():
            candidate_record = row_to_handoff(candidate)
            candidate_payload = candidate_record.get("sourcePayload") or {}
            try:
                candidate_identity = canonical_fpi_content_id(candidate_record, candidate_payload)
            except ValueError:
                continue
            if candidate_identity == canonical_content_id:
                canonical_matches.append((candidate, candidate_record))
        if len(canonical_matches) > 1:
            raise ValueError(f"Duplicate FPI canonical identity: {canonical_content_id}")
        if canonical_matches:
            existing, candidate_record = canonical_matches[0]
            graphics_request_id = str(candidate_record.get("graphicsRequestId") or graphics_request_id)
    image_type = normalize_content_type(
        request.get("imageType")
        or request.get("contentType")
        or extract_handoff_value(source_payload, "imageType", "contentType"),
        content_type,
    )
    if image_type == "FP":
        image_type = "FPI"
    has_text = bool(extract_handoff_text(source_payload))
    has_required_source = has_text or (content_type == "FPI" and has_fpi_source_asset(source_payload))
    if not has_required_source and existing:
        existing_payload = json.loads(existing["source_payload_json"] or "{}")
        if existing_payload:
            source_payload = existing_payload
            has_text = bool(extract_handoff_text(source_payload))
            has_required_source = has_text or (content_type == "FPI" and has_fpi_source_asset(source_payload))
    handoff_status = normalize_enum(request.get("handoffStatus"), HANDOFF_STATUSES, "requested")
    pig_status = normalize_enum(request.get("pigStatus"), PIG_STATUSES, "not_started")
    qc_status = normalize_enum(request.get("qcStatus"), QC_STATUSES, "not_sent")
    source_status = str(request.get("sourceStatus") or "needs_graphics")
    blocked_reason = str(request.get("blockedReason") or "")
    existing_handoff_status = str(existing["handoff_status"] if existing else "")
    is_rework_request = bool(
        request.get("revisionOf")
        or request.get("originalGraphicsRequestId")
        or extract_handoff_value(source_payload, "revisionOf", "originalGraphicsRequestId")
        or source_status.lower().startswith(("rework", "manual_rework"))
    )
    terminal_source_refresh = False
    if (
        not is_rework_request
        and existing_handoff_status in TERMINAL_HANDOFF_STATUSES
        and handoff_status in SOURCE_QUEUE_REFRESH_STATUSES
    ):
        terminal_source_refresh = True
        source_status = str(existing["source_status"] or source_status)
        handoff_status = existing_handoff_status
        pig_status = normalize_enum(existing["pig_status"], PIG_STATUSES, pig_status)
        qc_status = normalize_enum(existing["qc_status"], QC_STATUSES, qc_status)
        blocked_reason = str(existing["blocked_reason"] or blocked_reason)
    elif not has_required_source and handoff_status in {"requested", "claimed"}:
        handoff_status = "blocked"
        pig_status = "failed"
        blocked_reason = blocked_reason or ("blank_fpi_source_asset" if content_type == "FPI" else "blank_request_text")
    log_json = existing["transition_log_json"] if terminal_source_refresh and existing else append_transition_log(
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
            graphics_request_id, source_system, source_status, content_type, image_type,
            source_completion_id, revision_of, original_graphics_request_id, review_status, ocr_text,
            pig_status, handoff_status, qc_status,
            source_payload_json, blocked_reason, transition_log_json, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(graphics_request_id) DO UPDATE SET
            source_system = excluded.source_system,
            source_status = excluded.source_status,
            content_type = excluded.content_type,
            image_type = excluded.image_type,
            source_completion_id = excluded.source_completion_id,
            revision_of = excluded.revision_of,
            original_graphics_request_id = excluded.original_graphics_request_id,
            review_status = excluded.review_status,
            ocr_text = excluded.ocr_text,
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
            source_status,
            content_type,
            image_type,
            str(request.get("sourceCompletionId") or extract_handoff_value(source_payload, "sourceCompletionId") or ""),
            str(request.get("revisionOf") or extract_handoff_value(source_payload, "revisionOf") or ""),
            str(request.get("originalGraphicsRequestId") or extract_handoff_value(source_payload, "originalGraphicsRequestId") or ""),
            str(request.get("reviewStatus") or extract_handoff_value(source_payload, "reviewStatus") or ""),
            str(request.get("ocrText") or extract_handoff_value(source_payload, "ocrText") or ""),
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


def claim_graphics_handoff(connection: Any, graphics_request_id: str, claimed_by: str = "") -> dict[str, Any]:
    if is_firestore_ledger(connection):
        return connection.claim_handoff(graphics_request_id, claimed_by)
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


def update_graphics_handoff(connection: Any, graphics_request_id: str, update: dict[str, Any]) -> dict[str, Any]:
    if is_firestore_ledger(connection):
        return connection.update_handoff(graphics_request_id, update)
    existing = get_graphics_handoff(connection, graphics_request_id)
    if not existing:
        raise KeyError(f"Unknown graphicsRequestId: {graphics_request_id}")

    now = utc_now_iso()
    handoff_status = normalize_enum(update.get("handoffStatus"), HANDOFF_STATUSES, existing["handoffStatus"])
    pig_status = normalize_enum(update.get("pigStatus"), PIG_STATUSES, existing["pigStatus"])
    qc_status = normalize_enum(update.get("qcStatus"), QC_STATUSES, existing["qcStatus"])
    content_type = normalize_content_type(update.get("contentType") or update.get("imageType") or existing.get("contentType"), existing.get("contentType") or "QI")
    image_type = normalize_content_type(update.get("imageType") or update.get("contentType") or existing.get("imageType"), content_type)
    if content_type == "FP":
        content_type = "FPI"
    if image_type == "FP":
        image_type = "FPI"
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
            content_type = ?,
            image_type = ?,
            source_completion_id = ?,
            revision_of = ?,
            original_graphics_request_id = ?,
            review_status = ?,
            ocr_text = ?,
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
            content_type,
            image_type,
            str(update.get("sourceCompletionId") or existing.get("sourceCompletionId") or ""),
            str(update.get("revisionOf") or existing.get("revisionOf") or ""),
            str(update.get("originalGraphicsRequestId") or existing.get("originalGraphicsRequestId") or ""),
            str(update.get("reviewStatus") or existing.get("reviewStatus") or ""),
            str(update.get("ocrText") or existing.get("ocrText") or ""),
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


def get_graphics_handoff_queue(
    connection: Any,
    limit: int = 100,
    filter_mode: str = "all",
    cursor: int = 0,
    content_type: str = "",
) -> list[dict[str, Any]]:
    if is_firestore_ledger(connection):
        return connection.get_handoff_queue(limit, filter_mode, cursor, content_type)
    normalized_filter = str(filter_mode or "all").strip().lower()
    normalized_content_type = str(content_type or "").strip().upper()
    if normalized_content_type not in {"", "QI", "FPI"}:
        raise ValueError("contentType must be QI or FPI")
    offset = max(0, int(cursor or 0))
    rows = connection.execute(
        """
        SELECT *
        FROM graphics_handoff_ledger
        WHERE handoff_status IN ('requested', 'claimed', 'rejected')
          AND pig_status NOT IN ('generated', 'exported', 'uploaded', 'failed')
          AND qc_status IN ('not_sent', 'needs_revision')
        ORDER BY created_at ASC
        LIMIT ? OFFSET ?
        """,
        (max(1, min(int(limit or 100), 500)), offset),
    ).fetchall()
    return [
        queue_card_record(record)
        for record in (
            record
            for record in (row_to_handoff(row) for row in rows)
            if record.get("isActionable")
            and (
                not normalized_content_type
                or normalize_content_type(record.get("contentType") or record.get("imageType"))
                == normalized_content_type
            )
            and (
                normalized_filter in {"all", ""}
                or record.get("queueView") == normalized_filter
                or (normalized_filter == "rework" and record.get("nextAction") == "rework")
            )
        )
    ]


def replace_graphics_request_items(connection: sqlite3.Connection, graphics_request_id: str, items: list[dict[str, Any]]) -> None:
    if is_firestore_ledger(connection):
        connection.replace_graphics_request_items(graphics_request_id, items)
        return
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
    if is_firestore_ledger(connection):
        return connection.insert_graphics_completion(completion)
    completion_id = str(completion["id"]).strip()
    graphics_request_id = str(completion["graphics_request_id"]).strip()
    ingested_at = completion.get("ingested_at") or utc_now_iso()
    payload_json = json.dumps(completion.get("source_payload") or {}, ensure_ascii=True, sort_keys=True)
    content_type = normalize_content_type(completion.get("content_type") or completion.get("contentType") or completion.get("imageType"))
    image_type = normalize_content_type(completion.get("image_type") or completion.get("imageType") or completion.get("contentType"), content_type)

    connection.execute(
        """
        INSERT INTO graphics_completions (
            id, graphics_request_id, source_tool, content_type, image_type, asset_url, asset_preview_url,
            production_notes, completion_status, completed_at, ingested_at, source_payload_json
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
            graphics_request_id = excluded.graphics_request_id,
            source_tool = excluded.source_tool,
            content_type = excluded.content_type,
            image_type = excluded.image_type,
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
            content_type,
            image_type,
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
    if is_firestore_ledger(connection):
        return connection.insert_graphics_qc_review(review)
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
    if is_firestore_ledger(connection):
        return connection.insert_poetry_please_handoff(handoff)
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


def row_to_excerpt_handoff(row: sqlite3.Row) -> dict[str, Any]:
    return {
        "recordId": str(row["record_id"] or ""),
        "contentType": str(row["content_type"] or ""),
        "sourceSystem": str(row["source_system"] or ""),
        "sourceRecordId": str(row["source_record_id"] or ""),
        "author": str(row["author"] or ""),
        "bookTitle": str(row["book_title"] or ""),
        "poemTitle": str(row["poem_title"] or ""),
        "excerpt": str(row["excerpt_text"] or ""),
        "pageNumber": str(row["page_number"] or ""),
        "bookLink": str(row["book_link"] or ""),
        "releaseCatalog": str(row["release_catalog"] or ""),
        "bookShortener": str(row["book_shortener"] or ""),
        "driveLink": str(row["drive_link"] or ""),
        "sourceUrl": str(row["source_url"] or ""),
        "handoffStatus": str(row["handoff_status"] or ""),
        "handoffMode": str(row["handoff_mode"] or ""),
        "approvedAt": str(row["approved_at"] or ""),
        "updatedAt": str(row["updated_at"] or ""),
        "handedOffAt": str(row["handed_off_at"] or ""),
        "poetryPleaseItemId": str(row["poetry_please_item_id"] or ""),
        "errorMessage": str(row["error_message"] or ""),
        "payload": json.loads(row["payload_json"] or "{}"),
        "createdAt": str(row["created_at"] or ""),
    }


def upsert_excerpt_handoff(connection: sqlite3.Connection, handoff: dict[str, Any]) -> dict[str, Any]:
    record_id = str(handoff.get("recordId") or "").strip()
    if not record_id:
        raise ValueError("recordId is required")

    existing = connection.execute(
        "SELECT created_at FROM excerpt_handoff_ledger WHERE record_id = ?",
        (record_id,),
    ).fetchone()
    created_at = str(existing["created_at"] or "") if existing else utc_now_iso()
    updated_at = str(handoff.get("updatedAt") or utc_now_iso())

    connection.execute(
        """
        INSERT INTO excerpt_handoff_ledger (
            record_id, content_type, source_system, source_record_id, author, book_title, poem_title,
            excerpt_text, page_number, book_link, release_catalog, book_shortener, drive_link, source_url,
            handoff_status, handoff_mode, approved_at, updated_at, handed_off_at, poetry_please_item_id,
            error_message, payload_json, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(record_id) DO UPDATE SET
            content_type = excluded.content_type,
            source_system = excluded.source_system,
            source_record_id = excluded.source_record_id,
            author = excluded.author,
            book_title = excluded.book_title,
            poem_title = excluded.poem_title,
            excerpt_text = excluded.excerpt_text,
            page_number = excluded.page_number,
            book_link = excluded.book_link,
            release_catalog = excluded.release_catalog,
            book_shortener = excluded.book_shortener,
            drive_link = excluded.drive_link,
            source_url = excluded.source_url,
            handoff_status = excluded.handoff_status,
            handoff_mode = excluded.handoff_mode,
            approved_at = excluded.approved_at,
            updated_at = excluded.updated_at,
            handed_off_at = excluded.handed_off_at,
            poetry_please_item_id = excluded.poetry_please_item_id,
            error_message = excluded.error_message,
            payload_json = excluded.payload_json
        """,
        (
            record_id,
            str(handoff.get("contentType") or "EXC"),
            str(handoff.get("sourceSystem") or "weaver"),
            str(handoff.get("sourceRecordId") or ""),
            str(handoff.get("author") or ""),
            str(handoff.get("bookTitle") or ""),
            str(handoff.get("poemTitle") or ""),
            str(handoff.get("excerpt") or handoff.get("quoteText") or ""),
            str(handoff.get("pageNumber") or ""),
            str(handoff.get("bookLink") or ""),
            str(handoff.get("releaseCatalog") or ""),
            str(handoff.get("bookShortener") or ""),
            str(handoff.get("driveLink") or ""),
            str(handoff.get("sourceUrl") or ""),
            str(handoff.get("handoffStatus") or "queued"),
            str(handoff.get("handoffMode") or "auto"),
            str(handoff.get("approvedAt") or ""),
            updated_at,
            str(handoff.get("handedOffAt") or ""),
            str(handoff.get("poetryPleaseItemId") or ""),
            str(handoff.get("errorMessage") or ""),
            json.dumps(handoff.get("payload") or {}, ensure_ascii=True, sort_keys=True),
            created_at,
        ),
    )
    connection.commit()
    return get_excerpt_handoff(connection, record_id) or {}


def get_excerpt_handoff(connection: sqlite3.Connection, record_id: str) -> dict[str, Any] | None:
    row = connection.execute(
        "SELECT * FROM excerpt_handoff_ledger WHERE record_id = ?",
        (record_id,),
    ).fetchone()
    return row_to_excerpt_handoff(row) if row else None


def get_excerpt_handoffs(connection: sqlite3.Connection, record_ids: list[str] | None = None) -> list[dict[str, Any]]:
    params: list[Any] = []
    sql = "SELECT * FROM excerpt_handoff_ledger"
    if record_ids:
        ids = [str(value or "").strip() for value in record_ids if str(value or "").strip()]
        if not ids:
            return []
        placeholders = ",".join("?" for _ in ids)
        sql += f" WHERE record_id IN ({placeholders})"
        params.extend(ids)
    sql += " ORDER BY updated_at DESC, created_at DESC"
    rows = connection.execute(sql, params).fetchall()
    return [row_to_excerpt_handoff(row) for row in rows]


def build_pending_graphics_qc_records_from_ledger(connection: Any) -> list[dict[str, Any]]:
    if not is_firestore_ledger(connection):
        raise RuntimeError("Pending Graphics QC read model requires Firestore")

    completions = connection.list_raw_documents(FIRESTORE_GRAPHICS_COMPLETIONS_COLLECTION)
    completion_ids = [str(record.get("id") or "").strip() for record in completions]
    qc_reviews = connection.get_latest_graphics_qc_reviews(completion_ids)
    handoffs = connection.get_latest_poetry_please_handoffs(completion_ids)
    pending: list[dict[str, Any]] = []

    for completion in completions:
        payload = completion.get("sourcePayload") if isinstance(completion.get("sourcePayload"), dict) else {}
        completion_id = str(
            completion.get("id") or payload.get("pigCompletionId") or ""
        ).strip()
        if not completion_id:
            continue
        qc = qc_reviews.get(completion_id) or {}
        if str(qc.get("decision") or payload.get("qcDecision") or "").strip():
            continue
        handoff = handoffs.get(completion_id) or {}
        graphics_request_id = str(
            completion.get("graphicsRequestId") or payload.get("graphicsRequestId") or ""
        ).strip()
        asset_url = str(completion.get("assetUrl") or payload.get("assetUrl") or "").strip()
        asset_preview_url = str(
            completion.get("assetPreviewUrl") or payload.get("assetPreviewUrl") or asset_url
        ).strip()
        pending.append({
            "pigCompletionId": completion_id,
            "graphicsRequestId": graphics_request_id,
            "recordId": str(payload.get("recordId") or graphics_request_id or completion_id),
            "sheetRow": payload.get("sheetRow") or 0,
            "storageTarget": str(payload.get("storageTarget") or "firestore"),
            "contentType": normalize_content_type(
                completion.get("contentType") or payload.get("contentType")
            ),
            "author": str(payload.get("author") or ""),
            "poemTitle": str(payload.get("poemTitle") or ""),
            "bookTitle": str(payload.get("bookTitle") or ""),
            "quoteText": str(payload.get("quoteText") or ""),
            "assetLinkUrl": asset_url,
            "assetPreviewUrl": asset_preview_url,
            "completedAt": str(completion.get("completedAt") or payload.get("completedAt") or ""),
            "graphicsQcDecision": "",
            "graphicsQcNote": "",
            "graphicsQcUpdatedAt": "",
            "poetryPleaseStatus": str(handoff.get("handoff_status") or ""),
            "poetryPleaseUpdatedAt": str(handoff.get("handed_off_at") or ""),
            "poetryPleaseNote": (
                f"item={handoff.get('poetry_please_item_id')}"
                if handoff.get("poetry_please_item_id") else ""
            ),
        })

    return sorted(
        pending,
        key=lambda record: (
            str(record.get("completedAt") or ""),
            str(record.get("pigCompletionId") or ""),
        ),
    )


def rebuild_graphics_qc_queue_cards(connection: Any) -> dict[str, Any]:
    records = build_pending_graphics_qc_records_from_ledger(connection)
    for record in records:
        completion_id = str(record.get("pigCompletionId") or "").strip()
        if completion_id:
            connection.write_raw_document(FIRESTORE_GRAPHICS_QC_QUEUE_COLLECTION, completion_id, {
                **record,
                "isPendingQc": True,
                "updatedAt": utc_now_iso(),
            })
    return {"ok": True, "written": len(records)}


def get_pending_graphics_qc_records(connection: Any, include_cleanup: bool = False) -> list[dict[str, Any]]:
    if not is_firestore_ledger(connection):
        raise RuntimeError("Pending Graphics QC read model requires Firestore")
    return sorted(
        [
            {key: value for key, value in record.items() if key not in {"isPendingQc", "updatedAt"}}
            for record in connection.query_raw_documents(
                FIRESTORE_GRAPHICS_QC_QUEUE_COLLECTION,
                "isPendingQc",
                True,
            )
            if include_cleanup or str(record.get("storageTarget") or "").lower() != "cleanup_sheet"
        ],
        key=lambda record: (
            str(record.get("completedAt") or ""),
            str(record.get("pigCompletionId") or ""),
        ),
    )


def get_latest_graphics_qc_reviews(
    connection: sqlite3.Connection,
    completion_ids: list[str] | None = None,
) -> dict[str, dict[str, Any]]:
    if is_firestore_ledger(connection):
        return connection.get_latest_graphics_qc_reviews(completion_ids)
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
    if is_firestore_ledger(connection):
        return connection.get_latest_poetry_please_handoffs(completion_ids)
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
