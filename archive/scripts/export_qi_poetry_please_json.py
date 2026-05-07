#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import sqlite3
import unicodedata
from collections import defaultdict
from pathlib import Path

DB_PATH = Path("data/qi_catalog_match.db")
CATALOG_DB_PATH = Path("/Users/buttonpublishingone/Desktop/CODEX/Social Media Dev/poetry_catalog/formal_catalog.db")
BOOK_LINK_OVERRIDES_PATH = Path("data/book_link_overrides.json")
DEFAULT_OUTPUT = Path("data/poetry_please_qi_spring_2026.json")
DEFAULT_REVIEW_OUTPUT = Path("data/poetry_please_qi_spring_2026_review.json")
ALL_DRIVE_OUTPUT = Path("data/poetry_please_qi_all_drive.json")
ALL_DRIVE_REVIEW_OUTPUT = Path("data/poetry_please_qi_all_drive_review.json")
SPRING_2026_FOLDER_LINKS = {
    "https://drive.google.com/drive/folders/1ZzxVBThoJ1AONiTwLDKhuB_Bdqa-NMak",
    "https://drive.google.com/drive/folders/1SqOXOmebxQDIMy4XkCbnB2Qm6IxavFgx",
    "https://drive.google.com/drive/folders/1eFqwDjT1wRv5OtAshLtxbf6Y2-28D-s8",
    "https://drive.google.com/drive/folders/1-nQIwYb-_8CADQQENKQHIfJq8mHZ56-j",
    "https://drive.google.com/drive/folders/18LuwQZxD5KTFyihVed24tCThifD4T1Dz",
    "https://drive.google.com/drive/folders/1mQYgJ24OBclsmBScphKzdJh4prxGl1Je",
    "https://drive.google.com/drive/folders/1jUsz8_OBOsYurVDivXlTptHJ6aAAB81X",
    "https://drive.google.com/drive/folders/1OmjNUknTxfrNCUZeKe3XWtb6rK1wDcGu",
    "https://drive.google.com/drive/folders/1HvLHSgDo8sXkPz_C88Bgef8vBVYrr7Qq",
}
BOOK_LINK_OVERRIDES = {
    ("a choir of honest killers", "achk"): "https://buttonpoetry.com/product/a-choir-of-honest-killers/",
    ("stunt water", "sw"): "https://buttonpoetry.com/product/stunt-water/",
    ("stunt water: the work of buddy wakefield", "sw"): "https://buttonpoetry.com/product/stunt-water/",
}


def slugify(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value or "")
    ascii_text = normalized.encode("ascii", "ignore").decode("ascii")
    ascii_text = ascii_text.upper()
    ascii_text = re.sub(r"[^A-Z0-9]+", "-", ascii_text)
    return ascii_text.strip("-")


def build_qi_id(book_shortener: str, poem_title: str) -> str:
    prefix = (book_shortener or "QI").upper()
    title_part = slugify(poem_title or "") or "UNTITLED"
    return f"{prefix}-QI-{title_part}"


def extract_drive_file_id(url: str) -> str:
    match = re.search(r"/file/d/([^/]+)", url or "")
    return match.group(1) if match else ""


def build_drive_image_url(drive_link: str) -> str:
    file_id = extract_drive_file_id(drive_link)
    if not file_id:
        return ""
    return f"https://drive.google.com/thumbnail?id={file_id}&sz=w2000"


def load_book_links(catalog_db_path: Path) -> dict[tuple[str, str], str]:
    conn = sqlite3.connect(catalog_db_path)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            """
            SELECT title, book_shortener, buttonpoetry_link
            FROM canonical_books
            """
        ).fetchall()
    finally:
        conn.close()

    mapping: dict[tuple[str, str], str] = {}
    for row in rows:
        title_key = (row["title"] or "").strip().casefold()
        shortener_key = (row["book_shortener"] or "").strip().casefold()
        mapping[(title_key, shortener_key)] = row["buttonpoetry_link"] or ""
    mapping.update(BOOK_LINK_OVERRIDES)
    if BOOK_LINK_OVERRIDES_PATH.exists():
        override_rows = json.loads(BOOK_LINK_OVERRIDES_PATH.read_text(encoding="utf-8"))
        for key, value in override_rows.items():
            shortener, title = key.split("||", 1)
            mapping[(title.strip().casefold(), shortener.strip().casefold())] = value
    return mapping


def build_misc(parts: dict[str, str]) -> str:
    return " | ".join(f"{key}={value}" for key, value in parts.items() if value)


def main() -> int:
    parser = argparse.ArgumentParser(description="Export Spring 2026 QI assets into a Poetry Please JSON staging payload.")
    parser.add_argument("--db", default=str(DB_PATH))
    parser.add_argument("--catalog-db", default=str(CATALOG_DB_PATH))
    parser.add_argument("--scope", choices=["spring_2026_drive", "all_drive"], default="spring_2026_drive")
    parser.add_argument("--output")
    parser.add_argument("--review-output")
    args = parser.parse_args()

    if args.scope == "all_drive":
        output_default = ALL_DRIVE_OUTPUT
        review_output_default = ALL_DRIVE_REVIEW_OUTPUT
    else:
        output_default = DEFAULT_OUTPUT
        review_output_default = DEFAULT_REVIEW_OUTPUT

    output_path = Path(args.output) if args.output else output_default
    review_output_path = Path(args.review_output) if args.review_output else review_output_default

    conn = sqlite3.connect(args.db)
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        """
        SELECT
            asset_key,
            file_name,
            link_url,
            folder_link,
            notes,
            excerpt,
            catalog_poem_id,
            resolved_author_name,
            resolved_book_title,
            resolved_poem_title,
            resolved_book_shortener,
            matched_catalog_release_catalog,
            book_release_catalog,
            author_name,
            book_title,
            poem_title
        FROM image_assets
        WHERE link_url IS NOT NULL
          AND TRIM(link_url) <> ''
        ORDER BY resolved_book_title, resolved_poem_title, file_name
        """
    ).fetchall()
    conn.close()
    book_links = load_book_links(Path(args.catalog_db))
    if args.scope == "spring_2026_drive":
        rows = [
            row
            for row in rows
            if (row["folder_link"] or "") in SPRING_2026_FOLDER_LINKS
            and (row["matched_catalog_release_catalog"] or "") == "Spring 2026"
        ]

    payload: list[dict[str, object]] = []
    review_payload: list[dict[str, object]] = []
    id_counts: defaultdict[str, int] = defaultdict(int)

    for row in rows:
        title = row["resolved_poem_title"] or row["poem_title"] or ""
        author = row["resolved_author_name"] or row["author_name"] or ""
        book = row["resolved_book_title"] or row["book_title"] or ""
        shortener = row["resolved_book_shortener"] or ""
        release_catalog = row["matched_catalog_release_catalog"] or row["book_release_catalog"] or ""
        base_id = build_qi_id(
            shortener,
            title or row["file_name"] or "",
        )
        id_counts[base_id] += 1
        image_id = base_id if id_counts[base_id] == 1 else f"{base_id}-V{id_counts[base_id]}"
        drive_link = row["link_url"] or ""
        book_link = book_links.get(
            (book.strip().casefold(), shortener.strip().casefold()),
            "",
        )
        misc = build_misc(
            {
                "source": "qi_catalog_match_db",
                "sourceAssetKey": row["asset_key"] or "",
                "sourceFileName": row["file_name"] or "",
                "sourceFolderLink": row["folder_link"] or "",
                "catalogPoemId": str(row["catalog_poem_id"] or ""),
                "matchStatus": "ready" if row["catalog_poem_id"] else "needs_review",
                "notes": row["notes"] or "",
            }
        )
        record = {
            "docId": image_id,
            "imageId": image_id,
            "imageType": "QI",
            "title": title,
            "author": author,
            "book": book,
            "imageUrl": build_drive_image_url(drive_link),
            "driveLink": drive_link,
            "bookLink": book_link,
            "releaseCatalog": release_catalog,
            "misc": misc,
        }
        payload.append(record)
        if row["catalog_poem_id"] is None:
            review_payload.append(record)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    review_output_path.parent.mkdir(parents=True, exist_ok=True)
    review_output_path.write_text(json.dumps(review_payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    print(
        json.dumps(
            {
                "exported": len(payload),
                "needs_review": len(review_payload),
                "output": str(output_path),
                "review_output": str(review_output_path),
            }
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
