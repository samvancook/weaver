"""Read current, approved poem text from the main Catalog service."""

from __future__ import annotations

import json
import os
from functools import lru_cache
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


def fetch_readable_books() -> dict:
    """Read a complete, single-revision Catalog listing using its stable cursor."""
    books = []
    revision = None
    after = 0
    while True:
        query = urllib.parse.urlencode({"readable": "true", "limit": 500, "after": after})
        page = _get_json("/books?" + query)
        if not isinstance(page, dict) or not isinstance(page.get("books"), list):
            raise ValueError("Catalog book list returned an invalid response")
        page_revision = page.get("catalog_revision")
        if not isinstance(page_revision, str) or not page_revision.strip():
            raise ValueError("Catalog book list has no revision")
        if revision is not None and page_revision != revision:
            raise ValueError("Catalog changed during book listing; retry the lookup")
        revision = page_revision
        last_id = after
        for row in page["books"]:
            if not isinstance(row, dict):
                raise ValueError("Catalog book list contains an invalid book")
            book_id = row.get("book_id")
            if type(book_id) is not int or book_id <= last_id:
                raise ValueError("Catalog book IDs are not strictly increasing")
            last_id = book_id
            count = row.get("poem_key_count")
            if type(count) is not int or count < 0:
                raise ValueError("Catalog book has an invalid poem-key count")
            if (row.get("readable") is True
                    and row.get("effective_status") in READABLE_STATUSES
                    and row.get("source_format") in {"epub", "docx", "pdf"}
                    and count > 0):
                books.append({
                    **row,
                    "canonical_book_id": book_id,
                    "book_shortener": row.get("shortener") or "",
                    "primary_source_format": row["source_format"],
                })
        if "next_after" not in page:
            raise ValueError("Catalog book list has no pagination marker")
        next_after = page["next_after"]
        if next_after is None:
            return {"books": books, "catalog_revision": revision}
        if type(next_after) is not int or next_after <= after or next_after != last_id:
            raise ValueError("Catalog book list returned an invalid cursor")
        after = next_after


def list_readable_books() -> list[dict]:
    return fetch_readable_books()["books"]


@lru_cache(maxsize=128)
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
