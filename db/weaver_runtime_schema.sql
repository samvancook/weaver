PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS graphics_requests (
    id TEXT PRIMARY KEY,
    request_status TEXT NOT NULL DEFAULT 'OPEN',
    source_type TEXT NOT NULL DEFAULT 'weaver_sheet_queue',
    book_title TEXT NOT NULL,
    poem_title TEXT NOT NULL,
    author TEXT NOT NULL,
    quote_text TEXT NOT NULL,
    normalized_book_key TEXT NOT NULL,
    normalized_poem_key TEXT NOT NULL,
    normalized_author_key TEXT NOT NULL,
    normalized_quote_key TEXT NOT NULL,
    word_count INTEGER NOT NULL DEFAULT 0,
    source_record_id TEXT,
    source_sheet_name TEXT,
    source_sheet_row INTEGER,
    source_payload_json TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    latest_completion_id TEXT
);

CREATE INDEX IF NOT EXISTS idx_graphics_requests_status
    ON graphics_requests(request_status);
CREATE INDEX IF NOT EXISTS idx_graphics_requests_book
    ON graphics_requests(normalized_book_key);

CREATE TABLE IF NOT EXISTS graphics_request_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    graphics_request_id TEXT NOT NULL,
    item_position INTEGER NOT NULL DEFAULT 1,
    source_record_id TEXT,
    source_sheet_name TEXT,
    source_sheet_row INTEGER,
    book_title TEXT NOT NULL,
    poem_title TEXT NOT NULL,
    author TEXT NOT NULL,
    quote_text TEXT NOT NULL,
    normalized_quote_key TEXT NOT NULL,
    created_at TEXT NOT NULL,
    FOREIGN KEY (graphics_request_id) REFERENCES graphics_requests(id) ON DELETE CASCADE
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_graphics_request_items_unique_source
    ON graphics_request_items(graphics_request_id, source_sheet_name, source_sheet_row, item_position);

CREATE TABLE IF NOT EXISTS graphics_completions (
    id TEXT PRIMARY KEY,
    graphics_request_id TEXT NOT NULL,
    source_tool TEXT NOT NULL DEFAULT 'P.I.G.',
    asset_url TEXT NOT NULL,
    asset_preview_url TEXT,
    production_notes TEXT NOT NULL DEFAULT '',
    completion_status TEXT NOT NULL DEFAULT 'RETURNED',
    completed_at TEXT NOT NULL,
    ingested_at TEXT NOT NULL,
    source_payload_json TEXT NOT NULL DEFAULT '{}',
    FOREIGN KEY (graphics_request_id) REFERENCES graphics_requests(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_graphics_completions_request
    ON graphics_completions(graphics_request_id);

CREATE TABLE IF NOT EXISTS graphics_handoff_ledger (
    graphics_request_id TEXT PRIMARY KEY,
    source_system TEXT NOT NULL DEFAULT 'weaver',
    source_status TEXT NOT NULL DEFAULT 'needs_graphics',
    pig_status TEXT NOT NULL DEFAULT 'not_started',
    handoff_status TEXT NOT NULL DEFAULT 'requested',
    qc_status TEXT NOT NULL DEFAULT 'not_sent',
    asset_url TEXT NOT NULL DEFAULT '',
    asset_preview_url TEXT NOT NULL DEFAULT '',
    drive_file_id TEXT NOT NULL DEFAULT '',
    drive_file_name TEXT NOT NULL DEFAULT '',
    mime_type TEXT NOT NULL DEFAULT '',
    export_type TEXT NOT NULL DEFAULT '',
    variant TEXT NOT NULL DEFAULT '',
    version TEXT NOT NULL DEFAULT '',
    claimed_by TEXT NOT NULL DEFAULT '',
    error_message TEXT NOT NULL DEFAULT '',
    blocked_reason TEXT NOT NULL DEFAULT '',
    source_payload_json TEXT NOT NULL DEFAULT '{}',
    pig_payload_json TEXT NOT NULL DEFAULT '{}',
    qc_payload_json TEXT NOT NULL DEFAULT '{}',
    transition_log_json TEXT NOT NULL DEFAULT '[]',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    claimed_at TEXT,
    generated_at TEXT,
    uploaded_at TEXT,
    sent_to_qc_at TEXT,
    approved_at TEXT,
    rejected_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_graphics_handoff_ledger_queue
    ON graphics_handoff_ledger(handoff_status, pig_status, qc_status, updated_at);
CREATE INDEX IF NOT EXISTS idx_graphics_handoff_ledger_source_status
    ON graphics_handoff_ledger(source_system, source_status);

CREATE TABLE IF NOT EXISTS graphics_qc_reviews (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    graphics_completion_id TEXT NOT NULL,
    decision TEXT NOT NULL,
    metadata_issue TEXT NOT NULL DEFAULT '',
    aesthetic_issue TEXT NOT NULL DEFAULT '',
    note TEXT NOT NULL DEFAULT '',
    reviewed_by TEXT NOT NULL DEFAULT '',
    reviewed_at TEXT NOT NULL,
    source_payload_json TEXT NOT NULL DEFAULT '{}',
    FOREIGN KEY (graphics_completion_id) REFERENCES graphics_completions(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_graphics_qc_reviews_completion
    ON graphics_qc_reviews(graphics_completion_id, reviewed_at DESC);

CREATE TABLE IF NOT EXISTS poetry_please_handoffs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    graphics_completion_id TEXT NOT NULL,
    handoff_status TEXT NOT NULL,
    handoff_mode TEXT NOT NULL DEFAULT 'preview',
    handed_off_at TEXT NOT NULL,
    poetry_please_item_id TEXT NOT NULL DEFAULT '',
    payload_json TEXT NOT NULL DEFAULT '{}',
    FOREIGN KEY (graphics_completion_id) REFERENCES graphics_completions(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_poetry_please_handoffs_completion
    ON poetry_please_handoffs(graphics_completion_id, handed_off_at DESC);

CREATE TABLE IF NOT EXISTS excerpt_handoff_ledger (
    record_id TEXT PRIMARY KEY,
    content_type TEXT NOT NULL DEFAULT 'EXC',
    source_system TEXT NOT NULL DEFAULT 'weaver',
    source_record_id TEXT NOT NULL DEFAULT '',
    author TEXT NOT NULL DEFAULT '',
    book_title TEXT NOT NULL DEFAULT '',
    poem_title TEXT NOT NULL DEFAULT '',
    excerpt_text TEXT NOT NULL DEFAULT '',
    page_number TEXT NOT NULL DEFAULT '',
    book_link TEXT NOT NULL DEFAULT '',
    release_catalog TEXT NOT NULL DEFAULT '',
    book_shortener TEXT NOT NULL DEFAULT '',
    drive_link TEXT NOT NULL DEFAULT '',
    source_url TEXT NOT NULL DEFAULT '',
    handoff_status TEXT NOT NULL DEFAULT 'queued',
    handoff_mode TEXT NOT NULL DEFAULT 'auto',
    approved_at TEXT NOT NULL DEFAULT '',
    updated_at TEXT NOT NULL DEFAULT '',
    handed_off_at TEXT NOT NULL DEFAULT '',
    poetry_please_item_id TEXT NOT NULL DEFAULT '',
    error_message TEXT NOT NULL DEFAULT '',
    payload_json TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_excerpt_handoff_ledger_status
    ON excerpt_handoff_ledger(handoff_status, updated_at DESC);
CREATE INDEX IF NOT EXISTS idx_excerpt_handoff_ledger_source
    ON excerpt_handoff_ledger(source_system, source_record_id);

CREATE VIEW IF NOT EXISTS v_latest_graphics_qc_reviews AS
SELECT review.*
FROM graphics_qc_reviews AS review
JOIN (
    SELECT graphics_completion_id, MAX(reviewed_at) AS max_reviewed_at, MAX(id) AS max_id
    FROM graphics_qc_reviews
    GROUP BY graphics_completion_id
) AS latest
  ON latest.graphics_completion_id = review.graphics_completion_id
 AND latest.max_reviewed_at = review.reviewed_at
 AND latest.max_id = review.id;

CREATE VIEW IF NOT EXISTS v_qc_approved_graphics AS
SELECT
    completion.id AS graphics_completion_id,
    completion.graphics_request_id,
    request.book_title,
    request.poem_title,
    request.author,
    request.quote_text,
    completion.asset_url,
    completion.asset_preview_url,
    completion.production_notes,
    completion.completed_at,
    review.reviewed_at AS qc_approved_at,
    review.note AS qc_note
FROM graphics_completions AS completion
JOIN graphics_requests AS request
  ON request.id = completion.graphics_request_id
JOIN v_latest_graphics_qc_reviews AS review
  ON review.graphics_completion_id = completion.id
WHERE review.decision = 'APPROVE';
