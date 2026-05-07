#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from weaver_runtime_db import DEFAULT_DB_PATH, connect_runtime_db, ensure_runtime_schema, get_runtime_db_summary


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Initialize the Weaver runtime database.")
    parser.add_argument("--db-path", type=Path, default=DEFAULT_DB_PATH, help="Output SQLite database path.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    connection = connect_runtime_db(args.db_path)
    try:
        ensure_runtime_schema(connection)
        summary = get_runtime_db_summary(connection)
    finally:
        connection.close()

    result = {
        "db_path": str(args.db_path),
        "size_bytes": args.db_path.stat().st_size if args.db_path.exists() else 0,
        **summary,
    }
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
