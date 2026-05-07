#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sqlite3
from pathlib import Path


DEFAULT_DB_PATH = Path("data/qi_catalog_match.db")

PATTERN_UPDATES = [
    ("Andrea - LOTB - Quote Image - %", "https://drive.google.com/drive/folders/1oOchQZH_Uh6qDtyhndGptOB2PBnVz4Tl"),
    ("Micheal Lee - TOWWK - Quote Image - %", "https://drive.google.com/drive/folders/1npGbiSdAUhmVp2IGuO2I1ltbUT_ajgqT"),
    ("DEMULDER - EPHE - Quote Images - %", "https://drive.google.com/drive/folders/1y1nz9Bj0gYS3XmLUHjOF6ZSMeDhrr3u_"),
    ("Freeman - URBA - Quote Image - %", "https://drive.google.com/drive/folders/1cRN4Yw69UKHDSkSZNG23xR6SXln-Aj3j"),
    ("Holmon - Quote Image - %", "https://drive.google.com/drive/folders/1oIxhf63K6ZfcyoXv1B1bYcIQ0zh3rbRJ"),
    ("Olayiwola - ISST - Quote Image - %", "https://drive.google.com/drive/folders/1vSbBnNOK0lyK1w8QDs0iadpEjmPH-MMW"),
    ("SCHMINKEY - WIDT - Quote Images - %", "https://drive.google.com/drive/folders/1iQB6pFHxR-R-JdjeR2THLCekl5y2kjsZ"),
    ("Schminkey - DDJ - Quote Image - %", "https://drive.google.com/drive/folders/1TC0904jJp0EMq4crYpXtNGpiwLZTnn5W"),
    ("%PBCA - Quote Images - %", "https://drive.google.com/drive/folders/1Y4dHA_Kdxd6J-HW-z18S3kEgKL36BQ7J"),
    ("%PBC - Quote Images - %", "https://drive.google.com/drive/folders/1Y4dHA_Kdxd6J-HW-z18S3kEgKL36BQ7J"),
    ("Thorngate-Rein -PBC- Quote images%", "https://drive.google.com/drive/folders/1Y4dHA_Kdxd6J-HW-z18S3kEgKL36BQ7J"),
    ("WINTERS - POMB - Quote Images - %", "https://drive.google.com/drive/folders/10InTWstSzVAcnPQChqJaaSWb0GBg-C1Z"),
    ("William - SCDMDH - Quote Image - %", "https://drive.google.com/drive/folders/16q6qdE073gAOsD2asZzSr0TMUcUn9WSF"),
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Backfill remaining folder_link values from known filename patterns."
    )
    parser.add_argument(
        "--db-path",
        type=Path,
        default=DEFAULT_DB_PATH,
        help=f"SQLite DB to update (default: {DEFAULT_DB_PATH})",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be updated without writing to the DB.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    connection = sqlite3.connect(args.db_path)
    try:
        updated_rows = 0
        matched_patterns: list[dict[str, object]] = []
        for pattern, folder_link in PATTERN_UPDATES:
            count = connection.execute(
                """
                SELECT COUNT(*)
                FROM image_assets
                WHERE (folder_link IS NULL OR TRIM(folder_link) = '')
                  AND file_name LIKE ?
                """,
                (pattern,),
            ).fetchone()[0]
            if count == 0:
                continue
            matched_patterns.append(
                {
                    "pattern": pattern,
                    "folder_link": folder_link,
                    "rows": count,
                }
            )
            updated_rows += count
            if not args.dry_run:
                connection.execute(
                    """
                    UPDATE image_assets
                    SET folder_link = ?
                    WHERE (folder_link IS NULL OR TRIM(folder_link) = '')
                      AND file_name LIKE ?
                    """,
                    (folder_link, pattern),
                )
        if not args.dry_run:
            connection.commit()
        print(
            json.dumps(
                {
                    "dry_run": args.dry_run,
                    "updated_rows": updated_rows,
                    "matched_patterns": matched_patterns,
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
