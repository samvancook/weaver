const WEAVER_CONFIG = {
  version: "2026-03-30-explicit-review-only-1",
  spreadsheetId: "1yTCRQKAavimDEJka1-Ice4xlJ1mCm8hq0-G1PTQkTLM",
  sourceSheetName: "Excerpt Tool 1.20",
  startRow: 2,
  columnMap: {
    recordId: 25,
    author: 26,
    title: 27,
    excerpt: 28,
    bookTitle: 29,
    exclude: 30,
    duplicateGroupId: 32,
    exactPullCount: 34,
    approved: 35,
    statusIndicator: 36,
    quoteCreatedQc: 37,
    graphicsQi: 44,
    photos: 45,
    correctionNote: 46,
    validationStatus: 47,
    validationCanonicalBook: 48,
    validationCanonicalAuthor: 49,
    validationMatchedPoemTitle: 50,
    validationGlobalMatchBook: 51,
    validationGlobalMatchAuthor: 52,
    validationGlobalMatchPoem: 53,
    validationValidatedAt: 54,
    excerptReviewDecision: 55,
    useForInt: 56,
    correctedAuthor: 57,
    correctedTitle: 58,
    correctedBookTitle: 59,
    correctedExcerpt: 60
  }
};

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "";
  const callback = (e && e.parameter && e.parameter.callback) || "";
  let payload;

  if (!action) {
    payload = { ok: true, message: "Weaver Apps Script API is running.", version: WEAVER_CONFIG.version };
  } else if (action === "books") {
    payload = getWeaverBooks_();
  } else if (action === "excerpts") {
    payload = getWeaverExcerptsForBook_((e.parameter.bookTitle || "").toString());
  } else if (action === "correctionBooks") {
    payload = getWeaverCorrectionBooks_();
  } else if (action === "corrections") {
    payload = getWeaverCorrectionsForBook_((e.parameter.bookTitle || "").toString());
  } else if (action === "pendingRecords") {
    payload = getAllPendingWeaverRecords_();
  } else if (action === "validationQueue") {
    payload = getValidationQueue_((e.parameter.mode || "").toString());
  } else if (action === "saveReview") {
    payload = saveSingleWeaverReview_(e.parameter || {});
  } else if (action === "debugRecord") {
    payload = debugWeaverRecord_(e.parameter || {});
  } else {
    payload = { ok: false, error: "Unknown action" };
  }

  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(payload) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const action = (e && e.parameter && e.parameter.action) || "";

  if (action === "saveReviews") {
    const payloadText =
      (e && e.parameter && e.parameter.payload) ||
      (e && e.postData && e.postData.contents) ||
      "{}";
    const parsed = safeParseJson_(payloadText);
    const result = saveWeaverReviews_(parsed);

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } else if (action === "saveValidationBatch") {
    const payloadText =
      (e && e.postData && e.postData.contents) ||
      (e && e.parameter && e.parameter.payload) ||
      "{}";
    const parsed = safeParseJson_(payloadText);
    const result = saveValidationBatch_(parsed);

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: "Unknown action" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function safeParseJson_(text) {
  try {
    return JSON.parse(text);
  } catch (_error) {
    return {};
  }
}

function getWeaverBooks_() {
  const config = WEAVER_CONFIG;
  const sourceSheet = getSourceSheet_();
  ensureWeaverReviewColumns_(sourceSheet);
  const lastRow = sourceSheet.getLastRow();

  if (lastRow < config.startRow) {
    return { ok: true, books: [] };
  }

  const rowCount = lastRow - config.startRow + 1;
  const width = config.columnMap.correctedExcerpt;
  const values = sourceSheet.getRange(config.startRow, 1, rowCount, width).getValues();
  const counts = {};

  values.forEach(function(row) {
    const bookTitle = cleanWhitespace_(row[config.columnMap.bookTitle - 1]);
    const excerpt = cleanWhitespace_(row[config.columnMap.excerpt - 1]);
    const excluded = isYes_(row[config.columnMap.exclude - 1]);
    const reviewDecision = getExcerptReviewDecisionFromRow_(row, config);
    if (!bookTitle || !excerpt || excluded || !isPendingReview_(reviewDecision)) return;
    counts[bookTitle] = (counts[bookTitle] || 0) + 1;
  });

  const books = Object.keys(counts)
    .sort(function(a, b) {
      return a.localeCompare(b);
    })
    .map(function(title) {
      return {
        title: title,
        count: counts[title]
      };
    });

  return { ok: true, version: config.version, books: books };
}

function getWeaverExcerptsForBook_(bookTitle) {
  const config = WEAVER_CONFIG;
  const sourceSheet = getSourceSheet_();
  ensureWeaverReviewColumns_(sourceSheet);
  const resolvedBookTitle = cleanWhitespace_(bookTitle);

  if (!resolvedBookTitle) {
    return { ok: true, bookTitle: "", excerpts: [] };
  }

  const lastRow = sourceSheet.getLastRow();
  if (lastRow < config.startRow) {
    return { ok: true, bookTitle: resolvedBookTitle, excerpts: [] };
  }

  const rowCount = lastRow - config.startRow + 1;
  const width = config.columnMap.correctedExcerpt;
  const values = sourceSheet.getRange(config.startRow, 1, rowCount, width).getValues();
  const excerpts = [];

  values.forEach(function(row, index) {
    const currentBookTitle = cleanWhitespace_(row[config.columnMap.bookTitle - 1]);
    const excerptText = (row[config.columnMap.excerpt - 1] || "").toString();
    const cleanedExcerptText = cleanWhitespace_(excerptText);
    const excluded = isYes_(row[config.columnMap.exclude - 1]);
    const approved = (row[config.columnMap.approved - 1] || "").toString();
    const statusIndicator = cleanWhitespace_(row[config.columnMap.statusIndicator - 1]);
    const reviewDecision = getExcerptReviewDecisionFromRow_(row, config);
    if (
      currentBookTitle !== resolvedBookTitle ||
      !cleanedExcerptText ||
      excluded ||
      !isPendingReview_(reviewDecision)
    ) {
      return;
    }

    excerpts.push(buildExcerptPayload_(row, config.startRow + index, config));
  });

  return {
    ok: true,
    version: config.version,
    bookTitle: resolvedBookTitle,
    excerpts: excerpts
  };
}

function getWeaverCorrectionBooks_() {
  const config = WEAVER_CONFIG;
  const sourceSheet = getSourceSheet_();
  ensureWeaverReviewColumns_(sourceSheet);
  const lastRow = sourceSheet.getLastRow();

  if (lastRow < config.startRow) {
    return { ok: true, version: config.version, books: [] };
  }

  const rowCount = lastRow - config.startRow + 1;
  const width = config.columnMap.correctedExcerpt;
  const values = sourceSheet.getRange(config.startRow, 1, rowCount, width).getValues();
  const counts = {};

  values.forEach(function(row) {
    const bookTitle = cleanWhitespace_(row[config.columnMap.bookTitle - 1]);
    const excerpt = cleanWhitespace_(row[config.columnMap.excerpt - 1]);
    const excluded = isYes_(row[config.columnMap.exclude - 1]);
    const reviewDecision = getExcerptReviewDecisionFromRow_(row, config);
    if (!bookTitle || !excerpt || excluded || reviewDecision !== "NEEDS_CORRECTION") return;
    counts[bookTitle] = (counts[bookTitle] || 0) + 1;
  });

  const books = Object.keys(counts)
    .sort(function(a, b) {
      return a.localeCompare(b);
    })
    .map(function(title) {
      return {
        title: title,
        count: counts[title]
      };
    });

  return { ok: true, version: config.version, books: books };
}

function getWeaverCorrectionsForBook_(bookTitle) {
  const config = WEAVER_CONFIG;
  const sourceSheet = getSourceSheet_();
  ensureWeaverReviewColumns_(sourceSheet);
  const resolvedBookTitle = cleanWhitespace_(bookTitle);

  if (!resolvedBookTitle) {
    return { ok: true, version: config.version, bookTitle: "", excerpts: [] };
  }

  const lastRow = sourceSheet.getLastRow();
  if (lastRow < config.startRow) {
    return { ok: true, version: config.version, bookTitle: resolvedBookTitle, excerpts: [] };
  }

  const rowCount = lastRow - config.startRow + 1;
  const width = config.columnMap.correctedExcerpt;
  const values = sourceSheet.getRange(config.startRow, 1, rowCount, width).getValues();
  const excerpts = [];

  values.forEach(function(row, index) {
    const currentBookTitle = cleanWhitespace_(row[config.columnMap.bookTitle - 1]);
    const cleanedExcerptText = cleanWhitespace_(row[config.columnMap.excerpt - 1]);
    const excluded = isYes_(row[config.columnMap.exclude - 1]);
    const reviewDecision = getExcerptReviewDecisionFromRow_(row, config);

    if (
      currentBookTitle !== resolvedBookTitle ||
      !cleanedExcerptText ||
      excluded ||
      reviewDecision !== "NEEDS_CORRECTION"
    ) {
      return;
    }

    excerpts.push(buildExcerptPayload_(row, config.startRow + index, config));
  });

  return {
    ok: true,
    version: config.version,
    bookTitle: resolvedBookTitle,
    excerpts: excerpts
  };
}

function getAllPendingWeaverRecords_() {
  const config = WEAVER_CONFIG;
  const sourceSheet = getSourceSheet_();
  ensureWeaverReviewColumns_(sourceSheet);
  const lastRow = sourceSheet.getLastRow();

  if (lastRow < config.startRow) {
    return { ok: true, version: config.version, records: [] };
  }

  const rowCount = lastRow - config.startRow + 1;
  const width = config.columnMap.useForInt;
  const values = sourceSheet.getRange(config.startRow, 1, rowCount, width).getValues();
  const records = [];

  values.forEach(function(row, index) {
    const excerptText = (row[config.columnMap.excerpt - 1] || "").toString();
    const cleanedExcerptText = cleanWhitespace_(excerptText);
    const excluded = isYes_(row[config.columnMap.exclude - 1]);
    const reviewDecision = getExcerptReviewDecisionFromRow_(row, config);
    const bookTitle = cleanWhitespace_(row[config.columnMap.bookTitle - 1]);

    if (!bookTitle || !cleanedExcerptText || excluded || !isPendingReview_(reviewDecision)) {
      return;
    }

    records.push({
      sourceRow: config.startRow + index,
      recordId: (row[config.columnMap.recordId - 1] || "").toString(),
      author: (row[config.columnMap.author - 1] || "").toString(),
      title: (row[config.columnMap.title - 1] || "").toString(),
      bookTitle: bookTitle,
      excerptText: excerptText,
      catalogValidation: buildCatalogValidationPayload_(row, config)
    });
  });

  return {
    ok: true,
    version: config.version,
    records: records
  };
}

function buildExcerptPayload_(row, sourceRow, config) {
  const rawExcerptText = (row[config.columnMap.excerpt - 1] || "").toString();
  const correctedExcerptText = (row[config.columnMap.correctedExcerpt - 1] || "").toString();
  const excerptText = correctedExcerptText || rawExcerptText;
  const cleanedExcerptText = cleanWhitespace_(excerptText);
  const approved = (row[config.columnMap.approved - 1] || "").toString();
  const reviewDecision = getExcerptReviewDecisionFromRow_(row, config);
  const rawAuthor = (row[config.columnMap.author - 1] || "").toString();
  const rawTitle = (row[config.columnMap.title - 1] || "").toString();
  const rawBookTitle = cleanWhitespace_(row[config.columnMap.bookTitle - 1]);
  const correctedAuthor = (row[config.columnMap.correctedAuthor - 1] || "").toString();
  const correctedTitle = (row[config.columnMap.correctedTitle - 1] || "").toString();
  const correctedBookTitle = (row[config.columnMap.correctedBookTitle - 1] || "").toString();

  return {
    sourceRow: sourceRow,
    recordId: (row[config.columnMap.recordId - 1] || "").toString(),
    author: correctedAuthor || rawAuthor,
    title: correctedTitle || rawTitle,
    bookTitle: cleanWhitespace_(correctedBookTitle || rawBookTitle),
    excerptText: excerptText,
    wordCount: countWords_(cleanedExcerptText),
    approved: approved,
    statusIndicator: cleanWhitespace_(row[config.columnMap.statusIndicator - 1]),
    quoteCreatedQc: (row[config.columnMap.quoteCreatedQc - 1] || "").toString(),
    correctionNote: (row[config.columnMap.correctionNote - 1] || "").toString(),
    excludeRaw: (row[config.columnMap.exclude - 1] || "").toString(),
    duplicateGroupId: (row[config.columnMap.duplicateGroupId - 1] || "").toString(),
    exactPullCount: parseInteger_(row[config.columnMap.exactPullCount - 1]),
    pending: isPendingReview_(reviewDecision),
    excerptReviewDecision: reviewDecision,
    useForQi: approved === "Y",
    useForInt: isYes_(row[config.columnMap.useForInt - 1]),
    useForGraphicsQi: approved === "Y",
    useForPhotos: isYes_(row[config.columnMap.useForInt - 1]),
    correctedAuthor: correctedAuthor,
    correctedTitle: correctedTitle,
    correctedBookTitle: correctedBookTitle,
    correctedExcerpt: correctedExcerptText,
    rawAuthor: rawAuthor,
    rawTitle: rawTitle,
    rawBookTitle: rawBookTitle,
    rawExcerptText: rawExcerptText,
    catalogValidation: buildCatalogValidationPayload_(row, config)
  };
}

function getValidationQueue_(mode) {
  const queueMode = cleanWhitespace_(mode).toLowerCase();
  const pending = getAllPendingWeaverRecords_();
  if (!pending.ok) return pending;

  const records = pending.records.filter(function(record) {
    const validation = record.catalogValidation || {};
    const hasStatus = !!cleanWhitespace_(validation.status);
    if (queueMode === "force") return true;
    if (queueMode === "stale") return !hasStatus;
    return !hasStatus;
  });

  return {
    ok: true,
    version: WEAVER_CONFIG.version,
    records: records
  };
}

function saveWeaverReviews_(payload) {
  const config = WEAVER_CONFIG;
  const sourceSheet = getSourceSheet_();
  ensureWeaverReviewColumns_(sourceSheet);
  const updates = Array.isArray(payload && payload.updates) ? payload.updates : [];

  if (!updates.length) {
    return { ok: false, error: "No updates provided." };
  }

  const lastRow = sourceSheet.getLastRow();
  if (lastRow < config.startRow) {
    return { ok: false, error: "Source sheet has no data rows." };
  }

  const rowCount = lastRow - config.startRow + 1;
  const stateRange = sourceSheet.getRange(config.startRow, config.columnMap.approved, rowCount, config.columnMap.correctedExcerpt - config.columnMap.approved + 1);
  const stateValues = stateRange.getValues();

  let savedCount = 0;
  updates.forEach(function(update) {
    const sourceRow = parseInt(update.sourceRow, 10);
    if (!sourceRow) return;

    const sourceIndex = sourceRow - config.startRow;
    if (sourceIndex < 0 || sourceIndex >= stateValues.length) return;

    applyReviewStateToRow_(stateValues[sourceIndex], update);
    savedCount++;
  });

  stateRange.setValues(stateValues);
  SpreadsheetApp.flush();

  return {
    ok: true,
    version: config.version,
    savedCount: savedCount
  };
}

function saveSingleWeaverReview_(params) {
  const config = WEAVER_CONFIG;
  const sourceSheet = getSourceSheet_();
  ensureWeaverReviewColumns_(sourceSheet);
  const recordId = cleanWhitespace_(params.recordId);
  const sourceRow = parseInt(params.sourceRow, 10);
  let targetRow = 0;

  if (recordId) {
    targetRow = findSourceRowByRecordId_(sourceSheet, recordId, config);
  }

  if (!targetRow && sourceRow) {
    targetRow = sourceRow;
  }

  const lastRow = sourceSheet.getLastRow();
  if (!targetRow) {
    return { ok: false, error: "Missing source row target." };
  }

  if (targetRow < config.startRow || targetRow > lastRow) {
    return { ok: false, error: "source row is out of bounds." };
  }

  const stateRange = sourceSheet.getRange(targetRow, config.columnMap.approved, 1, config.columnMap.correctedExcerpt - config.columnMap.approved + 1);
  const stateRow = stateRange.getValues()[0];
  applyReviewStateToRow_(stateRow, {
    reviewDecision: params.reviewDecision || params.approval,
    useForQi: isTruthyParam_(params.useForQi || params.graphicsQi),
    useForInt: isTruthyParam_(params.useForInt || params.photos),
    correctionNote: params.correctionNote || "",
    correctedAuthor: params.correctedAuthor || "",
    correctedTitle: params.correctedTitle || "",
    correctedBookTitle: params.correctedBookTitle || "",
    correctedExcerpt: params.correctedExcerpt || ""
  });
  stateRange.setValues([stateRow]);
  SpreadsheetApp.flush();

  return {
    ok: true,
    version: config.version,
    sourceRow: targetRow,
    recordId: recordId
  };
}

function saveValidationBatch_(payload) {
  const config = WEAVER_CONFIG;
  const sourceSheet = getSourceSheet_();
  ensureWeaverReviewColumns_(sourceSheet);
  const updates = Array.isArray(payload && payload.updates) ? payload.updates : [];

  if (!updates.length) {
    return { ok: false, error: "No validation updates provided." };
  }

  let savedCount = 0;

  updates.forEach(function(update) {
    const recordId = cleanWhitespace_(update.recordId);
    const sourceRow = parseInt(update.sourceRow, 10);
    let targetRow = 0;

    if (recordId) {
      targetRow = findSourceRowByRecordId_(sourceSheet, recordId, config);
    }
    if (!targetRow && sourceRow) {
      targetRow = sourceRow;
    }
    if (!targetRow) return;

    sourceSheet.getRange(targetRow, config.columnMap.validationStatus, 1, 8).setValues([[
      cleanWhitespace_(update.status),
      cleanWhitespace_(update.bookCanonicalTitle),
      cleanWhitespace_(update.bookCanonicalAuthor),
      cleanWhitespace_(update.matchedPoemTitle),
      cleanWhitespace_(update.globalExcerptMatch && update.globalExcerptMatch.book_title),
      cleanWhitespace_(update.globalExcerptMatch && update.globalExcerptMatch.author),
      cleanWhitespace_(update.globalExcerptMatch && update.globalExcerptMatch.poem_title),
      new Date()
    ]]);
    savedCount++;
  });

  SpreadsheetApp.flush();

  return {
    ok: true,
    version: config.version,
    savedCount: savedCount
  };
}

function debugWeaverRecord_(params) {
  const config = WEAVER_CONFIG;
  const sourceSheet = getSourceSheet_();
  ensureWeaverReviewColumns_(sourceSheet);
  const recordId = cleanWhitespace_(params.recordId);
  const sourceRowParam = parseInt(params.sourceRow, 10);
  const lastRow = sourceSheet.getLastRow();

  if (lastRow < config.startRow) {
    return { ok: false, error: "Source sheet has no data rows." };
  }

  const width = config.columnMap.correctedExcerpt;

  if (recordId) {
    const foundRow = findSourceRowByRecordId_(sourceSheet, recordId, config);
    if (!foundRow) {
      return { ok: false, error: "Record not found.", recordId: recordId };
    }

    const row = sourceSheet.getRange(foundRow, 1, 1, width).getValues()[0];
    return buildDebugPayload_(row, foundRow, config);
  }

  if (sourceRowParam) {
    if (sourceRowParam < config.startRow || sourceRowParam > lastRow) {
      return { ok: false, error: "sourceRow is out of bounds.", sourceRow: sourceRowParam };
    }

    const row = sourceSheet.getRange(sourceRowParam, 1, 1, width).getValues()[0];
    return buildDebugPayload_(row, sourceRowParam, config);
  }

  return { ok: false, error: "Provide recordId or sourceRow." };
}

function getSourceSheet_() {
  const config = WEAVER_CONFIG;
  const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
  const sourceSheet = spreadsheet.getSheetByName(config.sourceSheetName);

  if (!sourceSheet) {
    throw new Error('Source sheet "' + config.sourceSheetName + '" not found.');
  }

  return sourceSheet;
}

function ensureWeaverReviewColumns_(sourceSheet) {
  const headers = [
    { column: WEAVER_CONFIG.columnMap.correctionNote, header: "correction_note" },
    { column: WEAVER_CONFIG.columnMap.validationStatus, header: "catalog_validation_status" },
    { column: WEAVER_CONFIG.columnMap.validationCanonicalBook, header: "catalog_canonical_book" },
    { column: WEAVER_CONFIG.columnMap.validationCanonicalAuthor, header: "catalog_canonical_author" },
    { column: WEAVER_CONFIG.columnMap.validationMatchedPoemTitle, header: "catalog_matched_poem_title" },
    { column: WEAVER_CONFIG.columnMap.validationGlobalMatchBook, header: "catalog_global_match_book" },
    { column: WEAVER_CONFIG.columnMap.validationGlobalMatchAuthor, header: "catalog_global_match_author" },
    { column: WEAVER_CONFIG.columnMap.validationGlobalMatchPoem, header: "catalog_global_match_poem" },
    { column: WEAVER_CONFIG.columnMap.validationValidatedAt, header: "catalog_validated_at" },
    { column: WEAVER_CONFIG.columnMap.excerptReviewDecision, header: "excerpt_review_decision" },
    { column: WEAVER_CONFIG.columnMap.useForInt, header: "use_for_int" },
    { column: WEAVER_CONFIG.columnMap.correctedAuthor, header: "corrected_author" },
    { column: WEAVER_CONFIG.columnMap.correctedTitle, header: "corrected_title" },
    { column: WEAVER_CONFIG.columnMap.correctedBookTitle, header: "corrected_book_title" },
    { column: WEAVER_CONFIG.columnMap.correctedExcerpt, header: "corrected_excerpt" }
  ];

  headers.forEach(function(definition) {
    const cell = sourceSheet.getRange(1, definition.column);
    const existing = cleanWhitespace_(cell.getValue());
    if (!existing) {
      cell.setValue(definition.header);
    }
  });

  backfillUseForInt_(sourceSheet, WEAVER_CONFIG);
}

function cleanWhitespace_(text) {
  return (text || "").toString().replace(/\s+/g, " ").trim();
}

function isYes_(value) {
  return (value || "").toString().trim().toUpperCase() === "Y";
}

function normalizeApprovalValue_(value) {
  const normalized = (value || "").toString().trim().toLowerCase();
  if (normalized === "accept") return "Y";
  if (normalized === "reject") return "N";
  return "";
}

function isPendingReview_(reviewDecision) {
  const decision = cleanWhitespace_(reviewDecision).toUpperCase();
  return !decision;
}

function isTruthyParam_(value) {
  const normalized = (value || "").toString().trim().toLowerCase();
  return normalized === "1" || normalized === "true" || normalized === "y" || normalized === "yes";
}

function countWords_(text) {
  const cleaned = cleanWhitespace_(text);
  if (!cleaned) return 0;
  return cleaned.split(/\s+/).length;
}

function parseInteger_(value) {
  const parsed = parseInt(value, 10);
  return isNaN(parsed) ? 0 : parsed;
}

function applyReviewStateToRow_(row, update) {
  const reviewDecision = normalizeReviewDecisionValue_(update.reviewDecision);
  const needsCorrection = reviewDecision === "NEEDS_CORRECTION";

  row[0] = update.useForQi ? "Y" : "N";
  row[1] = needsCorrection ? "NEEDS_CORRECTION" : "";
  row[11] = needsCorrection ? (update.correctionNote || "").toString().trim() : "";
  row[20] = reviewDecision;
  row[21] = update.useForInt ? "Y" : "";
  row[22] = (update.correctedAuthor || "").toString().trim();
  row[23] = (update.correctedTitle || "").toString().trim();
  row[24] = (update.correctedBookTitle || "").toString().trim();
  row[25] = (update.correctedExcerpt || "").toString();
}

function normalizeDecision_(value) {
  const normalized = (value || "").toString().trim().toLowerCase();
  if (normalized === "accept") return "accept";
  if (normalized === "reject") return "reject";
  if (normalized === "needs_correction") return "needs_correction";
  return "";
}

function normalizeReviewDecisionValue_(value) {
  const normalized = normalizeDecision_(value);
  if (normalized === "accept") return "ACCEPT";
  if (normalized === "reject") return "REJECT";
  if (normalized === "needs_correction") return "NEEDS_CORRECTION";
  return "";
}

function findSourceRowByRecordId_(sourceSheet, recordId, config) {
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < config.startRow) return 0;

  const values = sourceSheet
    .getRange(config.startRow, config.columnMap.recordId, lastRow - config.startRow + 1, 1)
    .getValues();

  for (let index = 0; index < values.length; index++) {
    if (cleanWhitespace_(values[index][0]) === recordId) {
      return config.startRow + index;
    }
  }

  return 0;
}

function buildCatalogValidationPayload_(row, config) {
  const status = cleanWhitespace_(row[config.columnMap.validationStatus - 1]);
  if (!status) {
    return null;
  }

  const globalBook = cleanWhitespace_(row[config.columnMap.validationGlobalMatchBook - 1]);
  const globalAuthor = cleanWhitespace_(row[config.columnMap.validationGlobalMatchAuthor - 1]);
  const globalPoem = cleanWhitespace_(row[config.columnMap.validationGlobalMatchPoem - 1]);

  return {
    status: status,
    bookCanonicalTitle: cleanWhitespace_(row[config.columnMap.validationCanonicalBook - 1]),
    bookCanonicalAuthor: cleanWhitespace_(row[config.columnMap.validationCanonicalAuthor - 1]),
    matchedPoemTitle: cleanWhitespace_(row[config.columnMap.validationMatchedPoemTitle - 1]),
    globalExcerptMatch: globalBook || globalAuthor || globalPoem
      ? {
          book_title: globalBook,
          author: globalAuthor,
          poem_title: globalPoem
        }
      : null,
    validatedAt: cleanWhitespace_(row[config.columnMap.validationValidatedAt - 1])
  };
}

function getExcerptReviewDecisionFromRow_(row, config) {
  const explicitDecision = cleanWhitespace_(row[config.columnMap.excerptReviewDecision - 1]).toUpperCase();
  if (explicitDecision) return explicitDecision;

  const statusIndicator = cleanWhitespace_(row[config.columnMap.statusIndicator - 1]).toUpperCase();
  if (statusIndicator === "NEEDS_CORRECTION") return "NEEDS_CORRECTION";

  return "";
}

function backfillUseForInt_(sourceSheet, config) {
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < config.startRow) return;

  const rowCount = lastRow - config.startRow + 1;
  const values = sourceSheet
    .getRange(config.startRow, config.columnMap.photos, rowCount, config.columnMap.useForInt - config.columnMap.photos + 1)
    .getValues();

  const updates = [];
  values.forEach(function(row, index) {
    const existingUseForInt = cleanWhitespace_(row[config.columnMap.useForInt - config.columnMap.photos]).toUpperCase();
    if (existingUseForInt) return;

    const legacyPhotosFlag = cleanWhitespace_(row[0]).toUpperCase();
    if (legacyPhotosFlag !== "Y") return;

    updates.push(config.startRow + index);
  });

  updates.forEach(function(rowNumber) {
    sourceSheet.getRange(rowNumber, config.columnMap.useForInt).setValue("Y");
  });
}

function buildDebugPayload_(row, sourceRow, config) {
  return {
    ok: true,
    sourceRow: sourceRow,
    recordId: cleanWhitespace_(row[config.columnMap.recordId - 1]),
    title: cleanWhitespace_(row[config.columnMap.title - 1]),
    author: cleanWhitespace_(row[config.columnMap.author - 1]),
    bookTitle: cleanWhitespace_(row[config.columnMap.bookTitle - 1]),
    excerpt: cleanWhitespace_(row[config.columnMap.excerpt - 1]),
    exclude: cleanWhitespace_(row[config.columnMap.exclude - 1]),
    approved: cleanWhitespace_(row[config.columnMap.approved - 1]),
    excerptReviewDecision: getExcerptReviewDecisionFromRow_(row, config),
    graphicsQi: cleanWhitespace_(row[config.columnMap.graphicsQi - 1]),
    photos: cleanWhitespace_(row[config.columnMap.photos - 1]),
    useForInt: cleanWhitespace_(row[config.columnMap.useForInt - 1]),
    correctedAuthor: cleanWhitespace_(row[config.columnMap.correctedAuthor - 1]),
    correctedTitle: cleanWhitespace_(row[config.columnMap.correctedTitle - 1]),
    correctedBookTitle: cleanWhitespace_(row[config.columnMap.correctedBookTitle - 1]),
    correctedExcerpt: cleanWhitespace_(row[config.columnMap.correctedExcerpt - 1]),
    pending: isPendingReview_(getExcerptReviewDecisionFromRow_(row, config)),
    catalogValidation: buildCatalogValidationPayload_(row, config)
  };
}
