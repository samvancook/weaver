#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from excerpt_library import fingerprint_excerpt, normalize_lookup_text


DEFAULT_WORKBOOK_PATH = Path("Excerpt Gathering & Quote Image Creation Database (Responses) copy.xlsx")
DEFAULT_QI_DATABASE_CSV = Path("data/qi_image_database.csv")
DEFAULT_OUTPUT_CSV = Path("data/qi_queue_cleanup_matches.csv")
DEFAULT_UPDATES_CSV = Path("data/qi_queue_cleanup_updates.csv")

XML_NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Match approved, not-yet-created rows in the quote creation workbook against the "
            "QI image database so we can mark already-made graphics out of Weaver."
        )
    )
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK_PATH)
    parser.add_argument("--qi-csv", type=Path, default=DEFAULT_QI_DATABASE_CSV)
    parser.add_argument("--output-csv", type=Path, default=DEFAULT_OUTPUT_CSV)
    parser.add_argument("--updates-csv", type=Path, default=DEFAULT_UPDATES_CSV)
    parser.add_argument(
        "--min-meta-score",
        type=int,
        default=2,
        choices=[0, 1, 2, 3],
        help="Minimum count of matching author/book/poem fields required for a live AQ override update.",
    )
    return parser.parse_args()


def col_to_num(label: str) -> int:
    number = 0
    for char in label:
        if char.isalpha():
            number = number * 26 + ord(char.upper()) - 64
    return number


def load_shared_strings(archive: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []

    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for item in root.findall("a:si", XML_NS):
        strings.append("".join(node.text or "" for node in item.iterfind(".//a:t", XML_NS)))
    return strings


def iter_sheet_rows(archive: ZipFile, sheet_path: str) -> Iterable[tuple[int, dict[int, str]]]:
    shared_strings = load_shared_strings(archive)
    root = ET.fromstring(archive.read(sheet_path))

    for row in root.findall(".//a:sheetData/a:row", XML_NS):
        row_number = int(row.attrib["r"])
        values: dict[int, str] = {}
        for cell in row.findall("a:c", XML_NS):
            reference = cell.attrib.get("r", "")
            match = re.match(r"([A-Z]+)(\d+)", reference)
            if not match:
                continue

            column_number = col_to_num(match.group(1))
            cell_type = cell.attrib.get("t")
            if cell_type == "inlineStr":
                node = cell.find("a:is/a:t", XML_NS)
                value = node.text if node is not None else ""
            elif cell_type == "s":
                node = cell.find("a:v", XML_NS)
                value = shared_strings[int(node.text)] if node is not None and node.text else ""
            else:
                node = cell.find("a:v", XML_NS)
                value = node.text if node is not None else ""
            values[column_number] = value
        yield row_number, values


def clean_flag(value: str | None) -> str:
    return (value or "").strip().upper()


def clean_text(value: str | None) -> str:
    return (value or "").replace("\r\n", "\n").replace("\r", "\n").strip()


def load_candidate_rows(workbook_path: Path) -> list[dict[str, str]]:
    with ZipFile(workbook_path) as archive:
        rows = iter_sheet_rows(archive, "xl/worksheets/sheet2.xml")
        candidates: list[dict[str, str]] = []
        for row_number, row in rows:
            if row_number == 1:
                continue

            approved = clean_flag(row.get(13)) == "Y"
            created = clean_flag(row.get(16)) == "Y"
            overridden = clean_flag(row.get(43)) == "Y"
            if not approved or created or overridden:
                continue

            candidates.append(
                {
                    "sheetRow": str(row_number),
                    "author": clean_text(row.get(3)),
                    "poemTitle": clean_text(row.get(4)),
                    "bookTitle": clean_text(row.get(6)),
                    "excerpt": clean_text(row.get(7)),
                    "sourceRow": clean_text(row.get(38)),
                    "recordId": clean_text(row.get(39)),
                }
            )
    return candidates


def load_qi_rows(qi_csv_path: Path) -> dict[str, list[dict[str, str]]]:
    rows_by_hash: dict[str, list[dict[str, str]]] = defaultdict(list)
    with qi_csv_path.open(newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            if not (
                row.get("link_url")
                or row.get("file_name")
                or row.get("low_res_link_url")
                or row.get("low_res_file_name")
            ):
                continue
            excerpt_hash = fingerprint_excerpt(row.get("excerpt", ""))
            rows_by_hash[excerpt_hash].append(row)
    return rows_by_hash


def score_metadata(candidate: dict[str, str], qi_row: dict[str, str]) -> tuple[int, dict[str, bool]]:
    author_match = bool(
        normalize_lookup_text(candidate["author"])
        and normalize_lookup_text(candidate["author"]) == normalize_lookup_text(qi_row.get("author_name", ""))
    )
    book_match = bool(
        normalize_lookup_text(candidate["bookTitle"])
        and normalize_lookup_text(candidate["bookTitle"]) == normalize_lookup_text(qi_row.get("book_title", ""))
    )
    poem_match = bool(
        normalize_lookup_text(candidate["poemTitle"])
        and normalize_lookup_text(candidate["poemTitle"]) == normalize_lookup_text(qi_row.get("poem_title", ""))
    )
    return int(author_match) + int(book_match) + int(poem_match), {
        "authorMatch": bool(author_match),
        "bookMatch": bool(book_match),
        "poemMatch": bool(poem_match),
    }


def build_matches(
    candidates: list[dict[str, str]],
    qi_rows_by_hash: dict[str, list[dict[str, str]]],
) -> tuple[list[dict[str, str]], Counter]:
    audit_rows: list[dict[str, str]] = []
    counts: Counter = Counter()

    for candidate in candidates:
        excerpt_hash = fingerprint_excerpt(candidate["excerpt"])
        qi_rows = qi_rows_by_hash.get(excerpt_hash, [])
        if not qi_rows:
            counts["noExactExcerpt"] += 1
            continue

        best_score = -1
        best_flags = {"authorMatch": False, "bookMatch": False, "poemMatch": False}
        best_row: dict[str, str] | None = None
        for qi_row in qi_rows:
            score, flags = score_metadata(candidate, qi_row)
            if score > best_score:
                best_score = score
                best_flags = flags
                best_row = qi_row

        counts[f"metaScore{best_score}"] += 1
        audit_rows.append(
            {
                "sheetRow": candidate["sheetRow"],
                "recordId": candidate["recordId"],
                "sourceRow": candidate["sourceRow"],
                "author": candidate["author"],
                "bookTitle": candidate["bookTitle"],
                "poemTitle": candidate["poemTitle"],
                "excerpt": candidate["excerpt"],
                "excerptHash": excerpt_hash,
                "matchCount": str(len(qi_rows)),
                "bestMetaScore": str(best_score),
                "authorMatch": "Y" if best_flags["authorMatch"] else "",
                "bookMatch": "Y" if best_flags["bookMatch"] else "",
                "poemMatch": "Y" if best_flags["poemMatch"] else "",
                "qiAuthor": (best_row or {}).get("author_name", ""),
                "qiBookTitle": (best_row or {}).get("book_title", ""),
                "qiPoemTitle": (best_row or {}).get("poem_title", ""),
                "qiExcerpt": (best_row or {}).get("excerpt", ""),
                "qiFileName": (best_row or {}).get("file_name", ""),
                "qiLinkUrl": (best_row or {}).get("link_url", ""),
                "qiLowResFileName": (best_row or {}).get("low_res_file_name", ""),
                "qiLowResLinkUrl": (best_row or {}).get("low_res_link_url", ""),
                "qiGraphicMadeStatus": (best_row or {}).get("graphic_made_status", ""),
                "qiApprovedForQuoteImage": (best_row or {}).get("approved_for_quote_image", ""),
            }
        )

    return audit_rows, counts


def write_csv(path: Path, rows: list[dict[str, str]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    args = parse_args()
    candidates = load_candidate_rows(args.workbook)
    qi_rows_by_hash = load_qi_rows(args.qi_csv)
    audit_rows, counts = build_matches(candidates, qi_rows_by_hash)

    audit_fieldnames = [
        "sheetRow",
        "recordId",
        "sourceRow",
        "author",
        "bookTitle",
        "poemTitle",
        "excerpt",
        "excerptHash",
        "matchCount",
        "bestMetaScore",
        "authorMatch",
        "bookMatch",
        "poemMatch",
        "qiAuthor",
        "qiBookTitle",
        "qiPoemTitle",
        "qiExcerpt",
        "qiFileName",
        "qiLinkUrl",
        "qiLowResFileName",
        "qiLowResLinkUrl",
        "qiGraphicMadeStatus",
        "qiApprovedForQuoteImage",
    ]
    write_csv(args.output_csv, audit_rows, audit_fieldnames)

    update_rows = [
        {
            "sheetRow": row["sheetRow"],
            "recordId": row["recordId"],
            "sourceRow": row["sourceRow"],
            "author": row["author"],
            "bookTitle": row["bookTitle"],
            "poemTitle": row["poemTitle"],
            "bestMetaScore": row["bestMetaScore"],
            "setDriveGraphicMatchOverride": "Y",
            "qiFileName": row["qiFileName"],
            "qiLinkUrl": row["qiLinkUrl"],
        }
        for row in audit_rows
        if int(row["bestMetaScore"]) >= args.min_meta_score
    ]
    update_fieldnames = [
        "sheetRow",
        "recordId",
        "sourceRow",
        "author",
        "bookTitle",
        "poemTitle",
        "bestMetaScore",
        "setDriveGraphicMatchOverride",
        "qiFileName",
        "qiLinkUrl",
    ]
    write_csv(args.updates_csv, update_rows, update_fieldnames)

    print(f"candidateRows={len(candidates)}")
    print(f"safeUpdates={len(update_rows)}")
    for key in sorted(counts):
        print(f"{key}={counts[key]}")
    print(f"auditCsv={args.output_csv}")
    print(f"updatesCsv={args.updates_csv}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
