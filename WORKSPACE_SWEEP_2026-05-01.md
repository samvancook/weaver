# Workspace Sweep 2026-05-01

Backup created before cleanup:
- [workspace-pre-cleanup-2026-04-30.tar.gz](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/backups/workspace-pre-cleanup-2026-04-30.tar.gz)

## Moved
- Moved loose branding images out of the repo root into:
  - [assets/branding](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/assets/branding)

## Keep and track
- [cloudbuild.weaver.yaml](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/cloudbuild.weaver.yaml)
- [db](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/db)
- [WEAVER_DATABASE_PLAN.md](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/WEAVER_DATABASE_PLAN.md)
- [EXCERPT_GATHERING_MODULE_PLAN.md](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/EXCERPT_GATHERING_MODULE_PLAN.md)
- [PEELED_OFF_CONTENT.md](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/PEELED_OFF_CONTENT.md)
- [WORKSPACE_SWEEP_2026-05-01.md](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/WORKSPACE_SWEEP_2026-05-01.md)
- Runtime/support scripts that appear to be part of current workflow:
  - [intake_catalog_lookup.py](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/intake_catalog_lookup.py)
  - [graphics_qi_lookup.py](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/graphics_qi_lookup.py)
  - [weaver_runtime_db.py](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/weaver_runtime_db.py)
  - [weaver_runtime_sync.py](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/weaver_runtime_sync.py)
  - [init_weaver_runtime_db.py](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/init_weaver_runtime_db.py)
  - [ingest_weaver_approved_excerpts.py](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/ingest_weaver_approved_excerpts.py)
  - [match_excerpt_candidates.py](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/match_excerpt_candidates.py)
  - [data/book_link_overrides.json](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/data/book_link_overrides.json)
  - [data/qi_image_database.csv](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/data/qi_image_database.csv)
  - [data/qi_image_database_missing_core_fields.csv](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/data/qi_image_database_missing_core_fields.csv)
  - [assets/branding](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/assets/branding)

## Legacy cleanup completed
- `.firebaserc`
- `firebase.json`
- `functions/`
- `weaver-apps-script/`

These have already been removed from the workspace and remain represented as deletions in Git status. They should be finalized together in Git and not reintroduced.

## Keep but archive elsewhere
- One-off audit/export scripts that may still be useful, but probably do not belong in the long-term product root:
  - moved to [archive/scripts](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/archive/scripts)

## Probably delete
- None moved or deleted in this pass.
- I do not currently see any untracked code/doc file that is obviously safe to delete without a human decision.
