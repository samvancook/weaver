#!/bin/sh
set -eu

export WEAVER_CATALOG_DB="/tmp/weaver-formal-catalog.db"
python3 prepare_catalog_overlay.py
exec npm start
