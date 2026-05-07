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
