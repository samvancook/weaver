#!/usr/bin/env python3
from __future__ import annotations

import json
import sqlite3
import sys

from catalog_validate import (
    DB_PATH,
    candidate_snippets,
    fetch_book_status,
    fetch_poems_for_book,
    normalize,
    title_aliases,
)


def lookup_poem_text(book_title: str, poem_title: str, excerpt_text: str = "") -> dict:
    if not book_title:
        return {"ok": False, "error": "bookTitle is required."}

    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    try:
      cursor = connection.cursor()
      book_status = fetch_book_status(cursor, book_title)
      if not book_status:
          return {"ok": False, "error": "Book not found in catalog."}

      poems = fetch_poems_for_book(cursor, int(book_status["canonical_book_id"]))
      requested_aliases = title_aliases(poem_title)
      requested_normalized = normalize(poem_title)
      excerpt_snippets = candidate_snippets(excerpt_text)

      best_match = None
      best_score = -1

      for poem in poems:
          poem_aliases = title_aliases(poem["title"])
          score = 0
          if requested_normalized and normalize(poem["title"]) == requested_normalized:
              score = 3
          elif requested_aliases & poem_aliases:
              score = 2
          elif requested_normalized and requested_normalized in normalize(poem["title"]):
              score = 1

          if score > best_score:
              best_score = score
              best_match = poem

      if not best_match or best_score < 1:
          fallback_match = None
          fallback_score = -1

          for poem in poems:
              poem_text = normalize(poem["text"])
              score = 0

              if any(snippet and snippet in poem_text for snippet in excerpt_snippets):
                  score = 2
              elif excerpt_snippets:
                  longest = max(excerpt_snippets, key=len)
                  if longest and longest[:80] and longest[:80] in poem_text:
                      score = 1

              if score > fallback_score:
                  fallback_score = score
                  fallback_match = poem

          if not fallback_match or fallback_score < 1:
              return {"ok": False, "error": "Poem not found in catalog for that book."}

          best_match = fallback_match

      return {
          "ok": True,
          "bookTitle": book_status["title"],
          "author": book_status["author"],
          "poemTitle": best_match["title"],
          "text": best_match["text"],
          "wordCount": best_match["word_count"],
      }
    finally:
      connection.close()


def main() -> int:
    payload = json.load(sys.stdin)
    result = lookup_poem_text(
        str(payload.get("bookTitle") or ""),
        str(payload.get("poemTitle") or ""),
        str(payload.get("excerptText") or ""),
    )
    json.dump(result, sys.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
