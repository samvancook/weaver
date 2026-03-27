const WEAVER_CONFIG = {
  version: "2026-03-27-correction-2",
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
    graphicsQi: 44,
    photos: 45,
    correctionNote: 46
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
  } else if (action === "pendingRecords") {
    payload = getAllPendingWeaverRecords_();
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
  const width = config.columnMap.photos;
  const values = sourceSheet.getRange(config.startRow, 1, rowCount, width).getValues();
  const counts = {};

  values.forEach(function(row) {
    const bookTitle = cleanWhitespace_(row[config.columnMap.bookTitle - 1]);
    const excerpt = cleanWhitespace_(row[config.columnMap.excerpt - 1]);
    const excluded = isYes_(row[config.columnMap.exclude - 1]);
    const approved = (row[config.columnMap.approved - 1] || "").toString();
    const statusIndicator = cleanWhitespace_(row[config.columnMap.statusIndicator - 1]);
    if (!bookTitle || !excerpt || excluded || !isPendingReview_(approved, statusIndicator)) return;
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
  const width = config.columnMap.correctionNote;
  const values = sourceSheet.getRange(config.startRow, 1, rowCount, width).getValues();
  const excerpts = [];

  values.forEach(function(row, index) {
    const currentBookTitle = cleanWhitespace_(row[config.columnMap.bookTitle - 1]);
    const excerptText = (row[config.columnMap.excerpt - 1] || "").toString();
    const cleanedExcerptText = cleanWhitespace_(excerptText);
    const excluded = isYes_(row[config.columnMap.exclude - 1]);
    const approved = (row[config.columnMap.approved - 1] || "").toString();
    const statusIndicator = cleanWhitespace_(row[config.columnMap.statusIndicator - 1]);
    if (
      currentBookTitle !== resolvedBookTitle ||
      !cleanedExcerptText ||
      excluded ||
      !isPendingReview_(approved, statusIndicator)
    ) {
      return;
    }

    excerpts.push({
      sourceRow: config.startRow + index,
      recordId: (row[config.columnMap.recordId - 1] || "").toString(),
      author: (row[config.columnMap.author - 1] || "").toString(),
      title: (row[config.columnMap.title - 1] || "").toString(),
      bookTitle: currentBookTitle,
      excerptText: excerptText,
      wordCount: countWords_(cleanedExcerptText),
      approved: (row[config.columnMap.approved - 1] || "").toString(),
      statusIndicator: statusIndicator,
      correctionNote: (row[config.columnMap.correctionNote - 1] || "").toString(),
      excludeRaw: (row[config.columnMap.exclude - 1] || "").toString(),
      duplicateGroupId: (row[config.columnMap.duplicateGroupId - 1] || "").toString(),
      exactPullCount: parseInteger_(row[config.columnMap.exactPullCount - 1]),
      pending: isPendingApproval_(approved),
      useForGraphicsQi: isYes_(row[config.columnMap.graphicsQi - 1]),
      useForPhotos: isYes_(row[config.columnMap.photos - 1])
    });
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
  const width = config.columnMap.correctionNote;
  const values = sourceSheet.getRange(config.startRow, 1, rowCount, width).getValues();
  const records = [];

  values.forEach(function(row, index) {
    const excerptText = (row[config.columnMap.excerpt - 1] || "").toString();
    const cleanedExcerptText = cleanWhitespace_(excerptText);
    const excluded = isYes_(row[config.columnMap.exclude - 1]);
    const approved = (row[config.columnMap.approved - 1] || "").toString();
    const statusIndicator = cleanWhitespace_(row[config.columnMap.statusIndicator - 1]);
    const bookTitle = cleanWhitespace_(row[config.columnMap.bookTitle - 1]);

    if (!bookTitle || !cleanedExcerptText || excluded || !isPendingReview_(approved, statusIndicator)) {
      return;
    }

    records.push({
      sourceRow: config.startRow + index,
      recordId: (row[config.columnMap.recordId - 1] || "").toString(),
      author: (row[config.columnMap.author - 1] || "").toString(),
      title: (row[config.columnMap.title - 1] || "").toString(),
      bookTitle: bookTitle,
      excerptText: excerptText
    });
  });

  return {
    ok: true,
    version: config.version,
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
  const stateRange = sourceSheet.getRange(config.startRow, config.columnMap.approved, rowCount, 12);
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

  const stateRange = sourceSheet.getRange(targetRow, config.columnMap.approved, 1, 12);
  const stateRow = stateRange.getValues()[0];
  applyReviewStateToRow_(stateRow, {
    approval: params.approval,
    useForGraphicsQi: isTruthyParam_(params.graphicsQi),
    useForPhotos: isTruthyParam_(params.photos),
    correctionNote: params.correctionNote || ""
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

  const width = config.columnMap.photos;

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
  const headerCell = sourceSheet.getRange(1, WEAVER_CONFIG.columnMap.correctionNote);
  const existing = cleanWhitespace_(headerCell.getValue());
  if (!existing) {
    headerCell.setValue("correction_note");
  }
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

function isPendingReview_(approvalValue, statusIndicator) {
  const approved = (approvalValue || "").toString().trim().toUpperCase();
  const status = (statusIndicator || "").toString().trim().toUpperCase();
  return approved !== "Y" && approved !== "N" && status !== "NEEDS_CORRECTION";
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
  const approval = normalizeApprovalValue_(update.approval);
  const needsCorrection = normalizeDecision_(update.approval) === "needs_correction";

  row[0] = approval;
  row[1] = needsCorrection ? "NEEDS_CORRECTION" : "";
  row[9] = update.useForGraphicsQi ? "Y" : "";
  row[10] = update.useForPhotos ? "Y" : "";
  row[11] = needsCorrection ? (update.correctionNote || "").toString().trim() : "";
}

function normalizeDecision_(value) {
  const normalized = (value || "").toString().trim().toLowerCase();
  if (normalized === "accept") return "accept";
  if (normalized === "reject") return "reject";
  if (normalized === "needs_correction") return "needs_correction";
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
    graphicsQi: cleanWhitespace_(row[config.columnMap.graphicsQi - 1]),
    photos: cleanWhitespace_(row[config.columnMap.photos - 1]),
    pending: isPendingApproval_(row[config.columnMap.approved - 1])
  };
}
