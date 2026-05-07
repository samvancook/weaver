#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

SPREADSHEET_ID = "1yTCRQKAavimDEJka1-Ice4xlJ1mCm8hq0-G1PTQkTLM"
SHEET_NAME = "New - Quote Creation Tool Database"
CLASPRC_PATH = Path.home() / ".clasprc.json"
DEFAULT_UPDATE_CSV = Path("data/qi_queue_cleanup_updates.csv")
SHEETS_BATCHGET_CHUNK_SIZE = 50


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Verify and optionally apply QI cleanup override updates to the live "
            "'New - Quote Creation Tool Database' sheet so Weaver stops showing "
            "already-made graphics in Needs graphics."
        )
    )
    parser.add_argument(
        "--update-csv",
        type=Path,
        default=DEFAULT_UPDATE_CSV,
        help=f"Cleanup override update CSV (default: {DEFAULT_UPDATE_CSV})",
    )
    parser.add_argument(
        "--verify-only",
        action="store_true",
        help="Only verify matching live rows without writing AQ overrides.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        help="Optional limit for testing with only the first N update rows.",
    )
    return parser.parse_args()


def clean_whitespace(text: str | None) -> str:
    return " ".join((text or "").split()).strip()


def normalize_text(text: str | None) -> str:
    normalized = clean_whitespace(text).lower()
    normalized = normalized.replace("’", "'").replace("‘", "'")
    normalized = normalized.replace("“", '"').replace("”", '"')
    normalized = normalized.replace("—", "-").replace("–", "-")
    return normalized


def excerpt_prefix(text: str | None, max_length: int = 120) -> str:
    return clean_whitespace(text)[:max_length]


def load_update_rows(path: Path, limit: int | None = None) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8", newline="") as handle:
        rows = list(csv.DictReader(handle))
    if limit is not None:
        rows = rows[:limit]
    return rows


def load_clasp_credentials() -> dict[str, str]:
    payload = json.loads(CLASPRC_PATH.read_text(encoding="utf-8"))
    token = payload["tokens"]["default"]
    return {
        "client_id": token["client_id"],
        "client_secret": token["client_secret"],
        "refresh_token": token["refresh_token"],
    }


def use_curl_json(
    method: str,
    url: str,
    headers: dict[str, str] | None = None,
    body: bytes | None = None,
) -> dict:
    command = ["curl", "-sS", "-X", method]
    for key, value in (headers or {}).items():
        command.extend(["-H", f"{key}: {value}"])
    if body is not None:
        command.extend(["--data-binary", "@-"])
    command.append(url)

    result = subprocess.run(command, input=body, capture_output=True, check=False)
    if result.returncode != 0:
        stderr = result.stderr.decode("utf-8", errors="replace").strip()
        raise RuntimeError(f"curl request failed: {stderr or f'exit {result.returncode}'}")

    stdout = result.stdout.decode("utf-8", errors="replace")
    if not stdout.strip():
        return {}

    try:
        return json.loads(stdout)
    except json.JSONDecodeError as error:
        raise RuntimeError(f"curl returned non-JSON response: {stdout[:400]}") from error


def should_fallback_to_curl(error: Exception) -> bool:
    message = str(error)
    return "CERTIFICATE_VERIFY_FAILED" in message or "unable to get local issuer certificate" in message


def refresh_access_token() -> str:
    credentials = load_clasp_credentials()
    body = urllib.parse.urlencode(
        {
            "client_id": credentials["client_id"],
            "client_secret": credentials["client_secret"],
            "refresh_token": credentials["refresh_token"],
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

    access_token = payload.get("access_token")
    if not access_token:
        raise RuntimeError("Could not refresh Google access token.")
    return access_token


def api_request(access_token: str, method: str, url: str, payload: dict | None = None) -> dict:
    body = None
    headers = {"Authorization": f"Bearer {access_token}"}
    if payload is not None:
        body = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"

    request = urllib.request.Request(url, data=body, headers=headers, method=method)
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as error:
        message = error.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Google API request failed ({error.code}): {message}") from error
    except urllib.error.URLError as error:
        if not should_fallback_to_curl(error):
            raise
        return use_curl_json(method, url, headers=headers, body=body)


def fetch_live_rows(access_token: str, sheet_rows: list[int]) -> dict[int, dict[str, str]]:
    live_by_row: dict[int, dict[str, str]] = {}
    for index in range(0, len(sheet_rows), SHEETS_BATCHGET_CHUNK_SIZE):
        chunk = sheet_rows[index : index + SHEETS_BATCHGET_CHUNK_SIZE]
        ranges = [f"'{SHEET_NAME}'!C{row}:AQ{row}" for row in chunk]
        url = (
            f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}/values:batchGet?"
            + urllib.parse.urlencode([("ranges", item) for item in ranges])
            + "&majorDimension=ROWS"
        )
        payload = api_request(access_token, "GET", url)

        for value_range in payload.get("valueRanges", []):
            range_name = value_range.get("range", "")
            row_number = int(range_name.split("!")[-1].split(":")[0][1:])
            values = value_range.get("values", [[]])
            row = values[0] if values else []

            padded = row + [""] * max(0, 41 - len(row))
            live_by_row[row_number] = {
                "author": padded[0],
                "poemTitle": padded[1],
                "bookTitle": padded[3],
                "excerptText": padded[4],
                "approved": padded[10],
                "created": padded[13],
                "sourceRow": padded[35],
                "recordId": padded[36],
                "aqOverride": padded[40],
            }
    return live_by_row


def row_matches(expected: dict[str, str], live: dict[str, str]) -> tuple[bool, str]:
    checks = [
        ("author", normalize_text(expected["author"]), normalize_text(live["author"])),
        ("bookTitle", normalize_text(expected["bookTitle"]), normalize_text(live["bookTitle"])),
        ("poemTitle", normalize_text(expected["poemTitle"]), normalize_text(live["poemTitle"])),
        (
            "excerptText",
            normalize_text(excerpt_prefix(expected["excerpt"])),
            normalize_text(excerpt_prefix(live["excerptText"])),
        ),
    ]
    for field, expected_value, live_value in checks:
        if expected_value != live_value:
            return False, field
    return True, ""


def build_update_payload(verified_rows: list[dict[str, str]]) -> dict:
    data = []
    for row in verified_rows:
        sheet_row = row["sheetRow"]
        data.append(
            {
                "range": f"'{SHEET_NAME}'!AQ{sheet_row}",
                "majorDimension": "ROWS",
                "values": [["Y"]],
            }
        )

    return {
        "valueInputOption": "USER_ENTERED",
        "data": data,
    }


def main() -> int:
    args = parse_args()
    update_rows = load_update_rows(args.update_csv, limit=args.limit)
    access_token = refresh_access_token()
    live_by_row = fetch_live_rows(access_token, [int(row["sheetRow"]) for row in update_rows])

    verified_rows: list[dict[str, str]] = []
    mismatches: list[dict[str, str]] = []
    already_done = 0

    for row in update_rows:
        sheet_row = int(row["sheetRow"])
        live = live_by_row.get(sheet_row)
        if not live:
            mismatches.append({"sheetRow": str(sheet_row), "reason": "missing_live_row"})
            continue

        matches, mismatch_field = row_matches(row, live)
        if not matches:
            mismatches.append({"sheetRow": str(sheet_row), "reason": f"mismatch_{mismatch_field}"})
            continue

        if normalize_text(row["recordId"]) and normalize_text(row["recordId"]) != normalize_text(live["recordId"]):
            mismatches.append({"sheetRow": str(sheet_row), "reason": "mismatch_recordId"})
            continue

        if normalize_text(row["sourceRow"]) and normalize_text(row["sourceRow"]) != normalize_text(live["sourceRow"]):
            mismatches.append({"sheetRow": str(sheet_row), "reason": "mismatch_sourceRow"})
            continue

        if clean_whitespace(live["aqOverride"]).upper() == "Y":
            already_done += 1
            continue

        verified_rows.append(row)

    summary = {
        "spreadsheetId": SPREADSHEET_ID,
        "sheetName": SHEET_NAME,
        "candidate_rows": len(update_rows),
        "verified_rows_to_update": len(verified_rows),
        "already_done": already_done,
        "mismatches": len(mismatches),
        "verify_only": bool(args.verify_only),
    }

    if args.verify_only:
        print(json.dumps({"summary": summary, "mismatches": mismatches[:25]}, indent=2))
        return 0

    if not verified_rows:
        print(json.dumps({"summary": summary, "message": "No verified rows require updates."}, indent=2))
        return 0

    update_payload = build_update_payload(verified_rows)
    response = api_request(
        access_token,
        "POST",
        f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}/values:batchUpdate",
        payload=update_payload,
    )

    print(
        json.dumps(
            {
                "summary": summary,
                "updatedDataRanges": len(update_payload["data"]),
                "totalUpdatedCells": response.get("totalUpdatedCells"),
                "totalUpdatedRows": response.get("totalUpdatedRows"),
                "totalUpdatedSheets": response.get("totalUpdatedSheets"),
            },
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(json.dumps({"ok": False, "error": str(error)}), file=sys.stderr)
        raise
