"""Read current, approved poem text from the main Catalog service."""

from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request


CATALOG_URL = os.environ.get(
    "BUTTON_POETRY_CATALOG_URL",
    "https://button-poetry-catalog-350789123099.us-central1.run.app",
).rstrip("/")
READABLE_STATUSES = {"catalog_ok", "provisional_source"}


def _get_json(path: str):
    request = urllib.request.Request(
        CATALOG_URL + path,
        headers={"Accept": "application/json", "User-Agent": "Weaver-live-catalog/1"},
    )
    with urllib.request.urlopen(request, timeout=15) as response:
        return json.load(response)


def list_readable_books() -> list[dict]:
    # Catalog search uses SQL LIKE; % lists its current book statuses in one call.
    rows = _get_json("/search?query=%25&limit=1000")
    if not isinstance(rows, list):
        raise ValueError("Catalog book search returned an invalid response")
    return [
        row for row in rows
        if isinstance(row, dict)
        and row.get("effective_status") in READABLE_STATUSES
        and row.get("primary_source_format") in {"epub", "docx", "pdf"}
        and int(row.get("catalog_poem_count") or 0) > 0
    ]


def fetch_readable_book(title_or_shortener: str) -> dict | None:
    path = "/books/" + urllib.parse.quote(title_or_shortener, safe="")
    path += "?include_poems=true&poem_limit=1000"
    try:
        payload = _get_json(path)
    except urllib.error.HTTPError as error:
        if error.code in {404, 409}:
            return None
        raise
    if not isinstance(payload, dict):
        raise ValueError("Catalog book lookup returned an invalid response")
    status = payload.get("status") or {}
    poems = payload.get("poems") or []
    if not isinstance(status, dict) or not isinstance(poems, list):
        raise ValueError("Catalog book lookup returned an invalid bundle")
    if status.get("effective_status") not in READABLE_STATUSES:
        return None
    if not poems:
        return None
    return {
        "status": status,
        "poems": [
            {
                "id": poem.get("id"),
                "poem_key": poem.get("poem_key"),
                "title": poem.get("title") or "",
                "text": poem.get("served_text") or poem.get("text") or "",
                "word_count": poem.get("word_count"),
                "book_title": status.get("title") or "",
                "author": status.get("author") or "",
            }
            for poem in poems
            if poem.get("title") and (poem.get("served_text") or poem.get("text"))
        ],
    }
