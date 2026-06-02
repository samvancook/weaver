#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import subprocess
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

SPREADSHEET_ID = "1yTCRQKAavimDEJka1-Ice4xlJ1mCm8hq0-G1PTQkTLM"
SHEET_NAME = "Queue - Needs Graphics"
CLASPRC_PATH = Path.home() / ".clasprc.json"
DEFAULT_SOURCE_JSON = Path("/private/tmp/weaver-stunt-water-requests-20260602.json")
DEFAULT_BACKUP_JSON = Path("/private/tmp/stunt-water-duplicate-cleanup-backup-20260602.json")
LEDGER_BASE_URL = "https://weaver.buttonpoetry.com/graphics-handoff"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Suppress later exact-duplicate Weaver graphics queue rows and block their handoff ledger records."
    )
    parser.add_argument("--source-json", type=Path, default=DEFAULT_SOURCE_JSON)
    parser.add_argument("--book-title", default="Stunt Water")
    parser.add_argument("--backup-json", type=Path, default=DEFAULT_BACKUP_JSON)
    parser.add_argument("--apply", action="store_true", help="Write suppression statuses after verification.")
    parser.add_argument(
        "--block-ledger-only",
        action="store_true",
        help="Block the matching Weaver handoff records after the sheet statuses were updated separately.",
    )
    return parser.parse_args()


def normalize_text(value: str | None) -> str:
    return (
        " ".join(str(value or "").split())
        .replace("’", "'")
        .replace("‘", "'")
        .replace("“", '"')
        .replace("”", '"')
        .replace("—", "-")
        .replace("–", "-")
        .lower()
        .strip()
    )


def queue_row(request_row: dict) -> int:
    rows = request_row.get("queueSheetRows") or []
    excerpts = request_row.get("excerpts") or []
    value = rows[0] if rows else (excerpts[0].get("queueSheetRow") if excerpts else 0)
    return int(value or 0)


def quote_text(request_row: dict) -> str:
    excerpts = request_row.get("excerpts") or []
    return str((excerpts[0].get("quoteText") if excerpts else request_row.get("quoteText")) or "")


def load_duplicate_targets(source_json: Path, book_title: str) -> list[dict]:
    payload = json.loads(source_json.read_text(encoding="utf-8"))
    requests = [
        row
        for row in payload.get("requests", [])
        if normalize_text(row.get("bookTitle")).startswith(normalize_text(book_title))
    ]
    grouped: dict[str, list[dict]] = {}
    for row in requests:
        identity = normalize_text(quote_text(row))
        if identity:
            grouped.setdefault(identity, []).append(row)

    targets: list[dict] = []
    for rows in grouped.values():
        ordered = sorted(rows, key=queue_row)
        if len(ordered) < 2:
            continue
        keep_row = queue_row(ordered[0])
        for duplicate in ordered[1:]:
            targets.append(
                {
                    "sheetRow": queue_row(duplicate),
                    "keepRow": keep_row,
                    "graphicsRequestId": str(duplicate.get("graphicsRequestId") or f"weaver:row-{queue_row(duplicate)}"),
                    "author": str(duplicate.get("author") or ""),
                    "poemTitle": str(duplicate.get("poemTitle") or ""),
                    "bookTitle": str(duplicate.get("bookTitle") or ""),
                    "quoteText": quote_text(duplicate),
                }
            )
    return sorted(targets, key=lambda row: row["sheetRow"])


def load_clasp_credentials() -> dict[str, str]:
    payload = json.loads(CLASPRC_PATH.read_text(encoding="utf-8"))
    token = payload["tokens"]["default"]
    return {
        "client_id": token["client_id"],
        "client_secret": token["client_secret"],
        "refresh_token": token["refresh_token"],
    }


def use_curl_json(method: str, url: str, *, headers: dict[str, str] | None = None, body: bytes | None = None) -> dict:
    command = ["curl", "-sS", "-X", method]
    for key, value in (headers or {}).items():
        command.extend(["-H", f"{key}: {value}"])
    if body is not None:
        command.extend(["--data-binary", "@-"])
    command.append(url)
    result = subprocess.run(command, input=body, capture_output=True, check=False)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.decode("utf-8", errors="replace").strip() or f"curl failed with {result.returncode}")
    content = result.stdout.decode("utf-8", errors="replace")
    payload = json.loads(content) if content.strip() else {}
    if isinstance(payload, dict) and payload.get("error"):
        raise RuntimeError(f"{method} {url} failed: {json.dumps(payload['error'])}")
    return payload


def should_fallback_to_curl(error: Exception) -> bool:
    message = str(error)
    return "CERTIFICATE_VERIFY_FAILED" in message or "unable to get local issuer certificate" in message


def json_request(method: str, url: str, *, headers: dict[str, str] | None = None, payload: dict | None = None) -> dict:
    body = json.dumps(payload).encode("utf-8") if payload is not None else None
    request_headers = dict(headers or {})
    if body is not None:
        request_headers["Content-Type"] = "application/json"
    request = urllib.request.Request(url, data=body, headers=request_headers, method=method)
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            content = response.read().decode("utf-8")
    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"{method} {url} failed ({error.code}): {detail}") from error
    except urllib.error.URLError as error:
        if not should_fallback_to_curl(error):
            raise
        return use_curl_json(method, url, headers=request_headers, body=body)
    payload = json.loads(content) if content.strip() else {}
    if isinstance(payload, dict) and payload.get("error"):
        raise RuntimeError(f"{method} {url} failed: {json.dumps(payload['error'])}")
    return payload


def refresh_access_token() -> str:
    credentials = load_clasp_credentials()
    body = urllib.parse.urlencode(
        {
            **credentials,
            "grant_type": "refresh_token",
        }
    ).encode("utf-8")
    request = urllib.request.Request(
        "https://oauth2.googleapis.com/token",
        data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.URLError as error:
        if not should_fallback_to_curl(error):
            raise
        payload = use_curl_json(
            "POST",
            "https://oauth2.googleapis.com/token",
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            body=body,
        )
    access_token = str(payload.get("access_token") or "")
    if not access_token:
        raise RuntimeError("Could not refresh Google access token.")
    return access_token


def fetch_live_rows(access_token: str, sheet_rows: list[int]) -> dict[int, list[str]]:
    live_rows: dict[int, list[str]] = {}
    for index in range(0, len(sheet_rows), 50):
        chunk = sheet_rows[index : index + 50]
        ranges = [f"'{SHEET_NAME}'!A{row}:I{row}" for row in chunk]
        url = (
            f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}/values:batchGet?"
            + urllib.parse.urlencode([("ranges", item) for item in ranges])
            + "&majorDimension=ROWS"
        )
        payload = json_request("GET", url, headers={"Authorization": f"Bearer {access_token}"})
        for value_range in payload.get("valueRanges", []):
            range_name = str(value_range.get("range") or "")
            row_number = int(range_name.split("!")[-1].split(":")[0][1:])
            values = value_range.get("values") or [[]]
            live_rows[row_number] = list(values[0] if values else []) + [""] * 9
    return live_rows


def verify_targets(targets: list[dict], live_rows: dict[int, list[str]]) -> tuple[list[dict], list[dict]]:
    verified: list[dict] = []
    mismatches: list[dict] = []
    for target in targets:
        row_number = target["sheetRow"]
        live = live_rows.get(row_number)
        if not live:
            mismatches.append({"sheetRow": row_number, "reason": "missing_live_row"})
            continue
        checks = [
            ("author", target["author"], live[0]),
            ("poemTitle", target["poemTitle"], live[1]),
            ("bookTitle", target["bookTitle"], live[2]),
            ("quoteText", target["quoteText"], live[3]),
        ]
        mismatch = next((field for field, expected, actual in checks if normalize_text(expected) != normalize_text(actual)), "")
        if mismatch:
            mismatches.append({"sheetRow": row_number, "reason": f"mismatch_{mismatch}"})
            continue
        verified.append({**target, "liveValues": live[:9], "alreadySuppressed": normalize_text(live[7]) == "duplicate_suppressed"})
    return verified, mismatches


def update_sheet(access_token: str, targets: list[dict]) -> dict:
    data = [
        {
            "range": f"'{SHEET_NAME}'!H{target['sheetRow']}",
            "majorDimension": "ROWS",
            "values": [["duplicate_suppressed"]],
        }
        for target in targets
        if not target["alreadySuppressed"]
    ]
    if not data:
        return {}
    return json_request(
        "POST",
        f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}/values:batchUpdate",
        headers={"Authorization": f"Bearer {access_token}"},
        payload={"valueInputOption": "USER_ENTERED", "data": data},
    )


def block_ledger_targets(targets: list[dict]) -> None:
    for target in targets:
        request_id = urllib.parse.quote(target["graphicsRequestId"], safe="")
        json_request(
            "PATCH",
            f"{LEDGER_BASE_URL}/{request_id}",
            payload={
                "sourceStatus": "duplicate_suppressed",
                "handoffStatus": "blocked",
                "pigStatus": "failed",
                "blockedReason": f"exact_duplicate_of_queue_row_{target['keepRow']}",
            },
        )


def main() -> int:
    args = parse_args()
    targets = load_duplicate_targets(args.source_json, args.book_title)
    if args.block_ledger_only:
        block_ledger_targets(targets)
        print(json.dumps({"blockedLedgerRecords": len(targets)}, indent=2))
        return 0

    access_token = refresh_access_token()
    live_rows = fetch_live_rows(access_token, [target["sheetRow"] for target in targets])
    verified, mismatches = verify_targets(targets, live_rows)
    summary = {
        "bookTitle": args.book_title,
        "candidateDuplicateRows": len(targets),
        "verifiedRows": len(verified),
        "alreadySuppressed": sum(1 for row in verified if row["alreadySuppressed"]),
        "mismatches": len(mismatches),
        "apply": bool(args.apply),
        "rows": [row["sheetRow"] for row in verified],
    }
    if mismatches:
        print(json.dumps({"summary": summary, "mismatches": mismatches}, indent=2))
        return 1
    if not args.apply:
        print(json.dumps({"summary": summary}, indent=2))
        return 0

    args.backup_json.write_text(
        json.dumps(
            {
                "createdAt": datetime.now(timezone.utc).isoformat(),
                "spreadsheetId": SPREADSHEET_ID,
                "sheetName": SHEET_NAME,
                "rows": verified,
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    response = update_sheet(access_token, verified)
    block_ledger_targets(verified)
    print(
        json.dumps(
            {
                "summary": summary,
                "backupJson": str(args.backup_json),
                "updatedSheetCells": response.get("totalUpdatedCells", 0),
                "blockedLedgerRecords": len(verified),
            },
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
