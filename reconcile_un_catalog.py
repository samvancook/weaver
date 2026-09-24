#!/usr/bin/env python3
"""Import corrected UN poem text without changing its set-aside status."""

from __future__ import annotations

import argparse
import sqlite3
from pathlib import Path


BOOK_ID = 5
REPLACED_TITLES = {"universal truths", "self love"}
LOCAL_TITLES = ("universal truths", "cycle", "self love", "namesake")
CLASSIFICATION_COLUMNS = (
    "content_kind", "form_kind", "is_likely_poem", "review_status", "confidence",
    "reason", "cleaned_text", "title_stripped_excerpt", "classified_by",
)


def rows(connection: sqlite3.Connection, query: str, parameters: tuple = ()) -> list[sqlite3.Row]:
    return connection.execute(query, parameters).fetchall()


def normalized(text: str) -> str:
    return " ".join(text.lower().split())


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--target", type=Path, required=True)
    parser.add_argument("--workbench", type=Path, required=True)
    parser.add_argument("--legacy", type=Path, required=True)
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()

    with sqlite3.connect(f"file:{args.workbench}?mode=ro", uri=True) as workbench, \
         sqlite3.connect(f"file:{args.legacy}?mode=ro", uri=True) as legacy:
        workbench.row_factory = sqlite3.Row
        legacy.row_factory = sqlite3.Row
        book = rows(workbench, "SELECT * FROM catalog_books WHERE canonical_book_id = ?", (BOOK_ID,))
        assert len(book) == 1 and book[0]["title"] == "unreliable narrator"
        source_poems = rows(workbench, "SELECT * FROM catalog_poems WHERE catalog_book_id = ? ORDER BY id", (book[0]["id"],))
        assert len(source_poems) == 45
        merged = {poem["title"]: poem for poem in source_poems if poem["title"] in REPLACED_TITLES}
        assert len(merged) == 2
        selected = [(workbench, poem) for poem in source_poems
                    if poem["title"] != "prologue" and poem["title"] not in REPLACED_TITLES]
        assert len(selected) == 42
        for title in LOCAL_TITLES:
            local = rows(legacy, "SELECT * FROM catalog_poems WHERE catalog_book_id = ? AND title = ?", (book[0]["id"], title))
            assert len(local) == 1, title
            if title in REPLACED_TITLES:
                assert normalized(local[0]["text"]) in normalized(merged[title]["text"]), title
            else:
                parent = "universal truths" if title == "cycle" else "self love"
                assert normalized(local[0]["text"]) in normalized(merged[parent]["text"]), title
            selected.append((legacy, local[0]))
        assert len(selected) == 46
        assert len({poem["title"].casefold() for _, poem in selected}) == 46

        with sqlite3.connect(args.target) as target:
            target.row_factory = sqlite3.Row
            status = rows(target, "SELECT * FROM book_status WHERE canonical_book_id = ?", (BOOK_ID,))
            assert len(status) == 1 and status[0]["effective_status"] == "set_aside"
            assert status[0]["primary_source_format"] in (None, "")
            existing = rows(target, "SELECT id FROM catalog_books WHERE canonical_book_id = ?", (BOOK_ID,))
            if existing:
                assert len(existing) == 1
                poems = rows(target, "SELECT title, text FROM catalog_poems WHERE catalog_book_id = ?", (existing[0]["id"],))
                assert {(row["title"], row["text"]) for row in poems} == {
                    (poem["title"], poem["text"]) for _, poem in selected
                }
                print("UN catalog already reconciled: 46 poems; status set_aside")
                return
            print("UN catalog plan: 42 workbench rows + 4 corrected local rows = 46 poems")
            if not args.apply:
                return

            book_columns = [column for column in book[0].keys() if column != "id"]
            book_values = [book[0][column] for column in book_columns]
            book_values[book_columns.index("poem_count")] = 46
            placeholders = ", ".join("?" for _ in book_columns)
            cursor = target.execute(
                f"INSERT INTO catalog_books ({', '.join(book_columns)}) VALUES ({placeholders})",
                book_values,
            )
            new_book_id = cursor.lastrowid
            for source, poem in selected:
                cursor = target.execute(
                    "INSERT INTO catalog_poems (catalog_book_id, title, text, character_count, word_count, extraction_method) "
                    "VALUES (?, ?, ?, ?, ?, ?)",
                    (new_book_id, poem["title"], poem["text"], poem["character_count"],
                     poem["word_count"], poem["extraction_method"]),
                )
                classification = rows(source, "SELECT * FROM poem_content_classifications WHERE catalog_poem_id = ?", (poem["id"],))
                assert len(classification) == 1, poem["title"]
                target.execute(
                    f"INSERT INTO poem_content_classifications (catalog_poem_id, {', '.join(CLASSIFICATION_COLUMNS)}) "
                    f"VALUES ({', '.join('?' for _ in range(len(CLASSIFICATION_COLUMNS) + 1))})",
                    (cursor.lastrowid, *(classification[0][column] for column in CLASSIFICATION_COLUMNS)),
                )
            target.execute(
                "UPDATE book_status SET catalog_book_count = 1, catalog_poem_count = 46 "
                "WHERE canonical_book_id = ? AND effective_status = 'set_aside'",
                (BOOK_ID,),
            )
            assert target.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
            print("UN catalog imported: 46 poems; status set_aside")


if __name__ == "__main__":
    main()
