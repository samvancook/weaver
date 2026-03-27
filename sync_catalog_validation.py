#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import ssl
import sqlite3
import sys
import urllib.error
import urllib.parse
import urllib.request
from typing import Any

from catalog_validate import DB_PATH, validate_record


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Validate pending Weaver records against the poetry catalog and write results back to the sheet."
    )
    parser.add_argument(
        "--url",
        default=os.environ.get("WEAVER_APPS_SCRIPT_URL", "").strip(),
        help="Weaver Apps Script web app URL. Defaults to WEAVER_APPS_SCRIPT_URL.",
    )
    parser.add_argument(
        "--mode",
        choices=["stale", "force"],
        default="stale",
        help="Validate only stale/unvalidated rows or force all pending rows.",
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=50,
        help="How many validation updates to send per write batch.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="Optional cap on how many queued records to process in this run.",
    )
    return parser.parse_args()


def fetch_json(url: str, action: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
    query = {"action": action}
    if params:
        query.update(params)
    full_url = f"{url}?{urllib.parse.urlencode(query)}"
    with urllib.request.urlopen(full_url, context=get_ssl_context()) as response:
        return json.loads(response.read().decode("utf-8"))


def post_json(url: str, action: str, payload: dict[str, Any]) -> dict[str, Any]:
    full_url = f"{url}?{urllib.parse.urlencode({'action': action})}"
    body = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(
        full_url,
        data=body,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )
    with urllib.request.urlopen(request, context=get_ssl_context()) as response:
        return json.loads(response.read().decode("utf-8"))


def get_ssl_context() -> ssl.SSLContext | None:
    if os.environ.get("WEAVER_INSECURE_SSL", "").strip() == "1":
      return ssl._create_unverified_context()
    return None


def chunked(items: list[dict[str, Any]], size: int) -> list[list[dict[str, Any]]]:
    return [items[index:index + size] for index in range(0, len(items), size)]


def main() -> int:
    args = parse_args()
    if not args.url:
        print("Missing --url or WEAVER_APPS_SCRIPT_URL.", file=sys.stderr)
        return 1

    try:
        queue = fetch_json(args.url, "validationQueue", {"mode": args.mode})
    except urllib.error.URLError as error:
        print(f"Failed to fetch validation queue: {error}", file=sys.stderr)
        return 1

    if not queue.get("ok"):
        print(f"Validation queue error: {queue.get('error', 'unknown error')}", file=sys.stderr)
        return 1

    records = queue.get("records", [])
    if not records:
        print("No records need validation.")
        return 0

    if args.limit and args.limit > 0:
        records = records[:args.limit]

    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    cursor = connection.cursor()
    try:
        updates = [validate_record(cursor, record) for record in records]
    finally:
        connection.close()

    total_saved = 0
    batches = chunked(updates, max(args.batch_size, 1))
    for index, batch in enumerate(batches, start=1):
        payload = {"updates": batch}
        try:
            result = post_json(args.url, "saveValidationBatch", payload)
        except urllib.error.URLError as error:
            print(f"Failed to write validation batch: {error}", file=sys.stderr)
            return 1

        if not result.get("ok"):
            print(f"Validation save error: {result.get('error', 'unknown error')}", file=sys.stderr)
            return 1

        total_saved += int(result.get("savedCount", 0))
        print(
            json.dumps(
                {
                    "progress": f"{index}/{len(batches)}",
                    "saved_so_far": total_saved,
                }
            ),
            flush=True,
        )

    statuses: dict[str, int] = {}
    for update in updates:
        status = (update.get("status") or "unknown").strip()
        statuses[status] = statuses.get(status, 0) + 1

    print(
        json.dumps(
            {
                "ok": True,
                "mode": args.mode,
                "queued": len(records),
                "saved": total_saved,
                "statuses": statuses,
            },
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
