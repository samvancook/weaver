#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sqlite3
import subprocess
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path


DEFAULT_DB_PATH = Path("data/qi_catalog_match.db")
CLASPRC_PATH = Path.home() / ".clasprc.json"
BATCH_SIZE = 200
DEFAULT_UNRESOLVED_REPORT = Path("data/folder_link_backfill_unresolved.json")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Backfill image_assets.folder_link in qi_catalog_match.db by asking the "
            "Google Drive API for each file's parent folder."
        )
    )
    parser.add_argument(
        "--db-path",
        type=Path,
        default=DEFAULT_DB_PATH,
        help=f"SQLite DB to update (default: {DEFAULT_DB_PATH})",
    )
    parser.add_argument(
        "--limit",
        type=int,
        help="Optional limit for testing only the first N rows missing folder_link.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Resolve folder links without writing updates to the DB.",
    )
    parser.add_argument(
        "--unresolved-report",
        type=Path,
        default=DEFAULT_UNRESOLVED_REPORT,
        help=f"Where to write unresolved rows (default: {DEFAULT_UNRESOLVED_REPORT})",
    )
    return parser.parse_args()


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


def api_request(access_token: str, method: str, url: str) -> dict:
    headers = {"Authorization": f"Bearer {access_token}"}
    request = urllib.request.Request(url, headers=headers, method=method)
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as error:
        message = error.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Google API request failed ({error.code}): {message}") from error
    except urllib.error.URLError as error:
        if not should_fallback_to_curl(error):
            raise
        return use_curl_json(method, url, headers=headers)


def extract_drive_file_id(url: str | None) -> str | None:
    if not url:
        return None
    parsed = urllib.parse.urlparse(url)
    segments = [segment for segment in parsed.path.split("/") if segment]

    if "file" in segments and "d" in segments:
        try:
            return segments[segments.index("d") + 1]
        except IndexError:
            return None

    if "open" in segments:
        return urllib.parse.parse_qs(parsed.query).get("id", [None])[0]

    return None


def load_rows_missing_folder_link(
    connection: sqlite3.Connection,
    limit: int | None = None,
) -> list[sqlite3.Row]:
    sql = """
        SELECT
            asset_key,
            file_name,
            link_url,
            folder_link,
            resolved_release_year,
            resolved_book_shortener,
            resolved_book_title
        FROM image_assets
        WHERE link_url IS NOT NULL
          AND TRIM(link_url) <> ''
          AND (folder_link IS NULL OR TRIM(folder_link) = '')
        ORDER BY resolved_release_year DESC, resolved_book_shortener, file_name
    """
    if limit is not None:
        sql += f" LIMIT {int(limit)}"
    return connection.execute(sql).fetchall()


def write_unresolved_report(
    path: Path,
    unresolved: list[dict[str, str]],
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(unresolved, ensure_ascii=True, indent=2), encoding="utf-8")


def resolve_parent_folder_link(access_token: str, file_id: str) -> str | None:
    url = (
        f"https://www.googleapis.com/drive/v3/files/{urllib.parse.quote(file_id)}?"
        "fields=id,parents&supportsAllDrives=true"
    )
    payload = api_request(access_token, "GET", url)
    parents = payload.get("parents") or []
    if not parents:
        return None
    return f"https://drive.google.com/drive/folders/{parents[0]}"


def apply_updates(connection: sqlite3.Connection, updates: list[tuple[str, str]]) -> None:
    connection.executemany(
        "UPDATE image_assets SET folder_link = ? WHERE asset_key = ?",
        [(folder_link, asset_key) for asset_key, folder_link in updates],
    )
    connection.commit()


def main() -> int:
    args = parse_args()

    connection = sqlite3.connect(args.db_path)
    connection.row_factory = sqlite3.Row
    try:
        rows = load_rows_missing_folder_link(connection, args.limit)
        if not rows:
            print("No rows need folder_link backfill.")
            return 0

        access_token = refresh_access_token()
        updates: list[tuple[str, str]] = []
        updated_count = 0
        unresolved: list[dict[str, str]] = []
        skipped_no_file_id = 0

        for row in rows:
            file_id = extract_drive_file_id(row["link_url"])
            if not file_id:
                skipped_no_file_id += 1
                unresolved.append(
                    {
                        "asset_key": row["asset_key"],
                        "file_name": row["file_name"] or "",
                        "link_url": row["link_url"] or "",
                        "resolved_release_year": row["resolved_release_year"] or "",
                        "resolved_book_shortener": row["resolved_book_shortener"] or "",
                        "resolved_book_title": row["resolved_book_title"] or "",
                        "reason": "missing_file_id",
                    }
                )
                continue

            try:
                folder_link = resolve_parent_folder_link(access_token, file_id)
            except Exception as error:  # noqa: BLE001
                unresolved.append(
                    {
                        "asset_key": row["asset_key"],
                        "file_name": row["file_name"] or "",
                        "link_url": row["link_url"] or "",
                        "resolved_release_year": row["resolved_release_year"] or "",
                        "resolved_book_shortener": row["resolved_book_shortener"] or "",
                        "resolved_book_title": row["resolved_book_title"] or "",
                        "reason": str(error),
                    }
                )
                continue

            if folder_link:
                updates.append((row["asset_key"], folder_link))
            else:
                unresolved.append(
                    {
                        "asset_key": row["asset_key"],
                        "file_name": row["file_name"] or "",
                        "link_url": row["link_url"] or "",
                        "resolved_release_year": row["resolved_release_year"] or "",
                        "resolved_book_shortener": row["resolved_book_shortener"] or "",
                        "resolved_book_title": row["resolved_book_title"] or "",
                        "reason": "no_parent_folder",
                    }
                )

            if not args.dry_run and len(updates) >= BATCH_SIZE:
                apply_updates(connection, updates)
                updated_count += len(updates)
                updates.clear()

        if not args.dry_run and updates:
            apply_updates(connection, updates)
            updated_count += len(updates)

        write_unresolved_report(args.unresolved_report, unresolved)

        print(
            json.dumps(
                {
                    "scanned": len(rows),
                    "updated": len(updates) if args.dry_run else updated_count,
                    "unresolved": len(unresolved),
                    "skipped_no_file_id": skipped_no_file_id,
                    "dry_run": args.dry_run,
                    "sample_unresolved": unresolved[:10],
                    "unresolved_report": str(args.unresolved_report),
                },
                ensure_ascii=True,
                indent=2,
            )
        )
        return 0
    finally:
        connection.close()


if __name__ == "__main__":
    raise SystemExit(main())
