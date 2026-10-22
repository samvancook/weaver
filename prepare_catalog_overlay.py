#!/usr/bin/env python3
"""Build a private runtime catalog overlay from the live main Catalog service."""

from __future__ import annotations

import json
import os
import shutil
import sqlite3
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

from catalog_validate import DEFAULT_DB_PATH
from import_main_catalog_book import import_book_payload


CATALOG_URL = "https://button-poetry-catalog-350789123099.us-central1.run.app"
RUNTIME_DB = Path(os.environ.get("WEAVER_CATALOG_DB", "/tmp/weaver-formal-catalog.db"))
BOOKS = ("BBFR", "WG")


def fetch_book(shortener: str) -> dict | None:
    url = f"{CATALOG_URL}/books/{urllib.parse.quote(shortener)}?include_poems=true&poem_limit=1000"
    request = urllib.request.Request(url, headers={"Accept": "application/json", "User-Agent": "Weaver-catalog-overlay/1"})
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            return json.load(response)
    except urllib.error.HTTPError as error:
        if error.code == 404:
            return None
        raise


def prepare() -> list[dict]:
    RUNTIME_DB.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(DEFAULT_DB_PATH, RUNTIME_DB)
    imported = []
    for shortener in BOOKS:
        payload = fetch_book(shortener)
        if payload is None:
            print(f"Catalog overlay: {shortener} absent from main Catalog", file=sys.stderr)
            continue
        imported.append(import_book_payload(payload, RUNTIME_DB))
    with sqlite3.connect(RUNTIME_DB) as connection:
        if connection.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
            raise RuntimeError("Catalog overlay integrity check failed")
    return imported


if __name__ == "__main__":
    print(json.dumps({"catalogOverlay": prepare()}))
