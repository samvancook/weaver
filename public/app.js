let runtimeConfig = window.WEAVER_CONFIG || {};

const elements = {
  showGatheringModule: document.getElementById("show-gathering-module"),
  showReviewModule: document.getElementById("show-review-module"),
  showWeirdModule: document.getElementById("show-weird-module"),
  showCorrectionsModule: document.getElementById("show-corrections-module"),
  showGraphicsModule: document.getElementById("show-graphics-module"),
  gatheringTabBook: document.getElementById("gathering-tab-book"),
  gatheringTabVideo: document.getElementById("gathering-tab-video"),
  gatheringTabFix: document.getElementById("gathering-tab-fix"),
  gatheringModule: document.getElementById("gathering-module"),
  reviewModule: document.getElementById("review-module"),
  reviewQueuePanel: document.getElementById("review-queue-panel"),
  weirdModule: document.getElementById("weird-module"),
  weirdQueuePanel: document.getElementById("weird-queue-panel"),
  correctionsModule: document.getElementById("corrections-module"),
  graphicsModule: document.getElementById("graphics-module"),
  loadBooks: document.getElementById("load-books"),
  bookSelect: document.getElementById("book-select"),
  weirdBookSelect: document.getElementById("weird-book-select"),
  reviewFilter: document.getElementById("review-filter"),
  reviewDisplayMode: document.getElementById("review-display-mode"),
  weirdReviewFilter: document.getElementById("weird-review-filter"),
  loadExcerpts: document.getElementById("load-excerpts"),
  loadWeirdExcerpts: document.getElementById("load-weird-excerpts"),
  submitReview: document.getElementById("submit-review"),
  submitWeirdReview: document.getElementById("submit-weird-review"),
  excerptList: document.getElementById("excerpt-list"),
  weirdExcerptList: document.getElementById("weird-excerpt-list"),
  loadCorrectionBooks: document.getElementById("load-correction-books"),
  correctionBookSelect: document.getElementById("correction-book-select"),
  loadCorrections: document.getElementById("load-corrections"),
  submitCorrections: document.getElementById("submit-corrections"),
  autoApplyCorrections: document.getElementById("auto-apply-corrections"),
  correctionList: document.getElementById("correction-list"),
  graphicsMode: document.getElementById("graphics-mode"),
  graphicsFilter: document.getElementById("graphics-filter"),
  graphicsBookSelect: document.getElementById("graphics-book-select"),
  loadGraphicsBooks: document.getElementById("load-graphics-books"),
  loadGraphicsRecords: document.getElementById("load-graphics-records"),
  graphicsFolderImportPanel: document.getElementById("graphics-folder-import-panel"),
  graphicsFolderUrl: document.getElementById("graphics-folder-url"),
  previewGraphicsFolderImport: document.getElementById("preview-graphics-folder-import"),
  applyGraphicsFolderImport: document.getElementById("apply-graphics-folder-import"),
  graphicsFolderImportResults: document.getElementById("graphics-folder-import-results"),
  exportGraphicsSheet: document.getElementById("export-graphics-sheet"),
  submitGraphicsQc: document.getElementById("submit-graphics-qc"),
  graphicsList: document.getElementById("graphics-list"),
  gatheringMode: document.getElementById("gathering-mode"),
  gatheringEmail: document.getElementById("gathering-email"),
  gatheringBookFields: document.getElementById("gathering-book-fields"),
  gatheringBookAuthor: document.getElementById("gathering-book-author"),
  gatheringBookTitle: document.getElementById("gathering-book-title"),
  gatheringBookBook: document.getElementById("gathering-book-book"),
  gatheringBookQuote: document.getElementById("gathering-book-quote"),
  gatheringBookQuoteMeta: document.getElementById("gathering-book-quote-meta"),
  gatheringBookSourceHint: document.getElementById("gathering-book-source-hint"),
  gatheringPrevCatalogPoem: document.getElementById("gathering-prev-catalog-poem"),
  gatheringNextCatalogPoem: document.getElementById("gathering-next-catalog-poem"),
  gatheringViewCatalogPoem: document.getElementById("gathering-view-catalog-poem"),
  gatheringBookNotes: document.getElementById("gathering-book-notes"),
  gatheringBookItalics: document.getElementById("gathering-book-italics"),
  gatheringBookReaction: document.getElementById("gathering-book-reaction"),
  gatheringVideoFields: document.getElementById("gathering-video-fields"),
  gatheringVideoAuthor: document.getElementById("gathering-video-author"),
  gatheringVideoTitle: document.getElementById("gathering-video-title"),
  gatheringVideoBook: document.getElementById("gathering-video-book"),
  gatheringVideoEvent: document.getElementById("gathering-video-event"),
  gatheringVideoQuote: document.getElementById("gathering-video-quote"),
  gatheringVideoQuoteMeta: document.getElementById("gathering-video-quote-meta"),
  gatheringFixFields: document.getElementById("gathering-fix-fields"),
  gatheringFixPart: document.getElementById("gathering-fix-part"),
  gatheringFixAuthor: document.getElementById("gathering-fix-author"),
  gatheringFixIncorrect: document.getElementById("gathering-fix-incorrect"),
  gatheringFixCorrect: document.getElementById("gathering-fix-correct"),
  gatheringAuthorOptions: document.getElementById("gathering-author-options"),
  gatheringBookOptions: document.getElementById("gathering-book-options"),
  gatheringBookPoemOptions: document.getElementById("gathering-book-poem-options"),
  loadGatheringOptions: document.getElementById("load-gathering-options"),
  submitGathering: document.getElementById("submit-gathering"),
  statusOutput: document.getElementById("status-output"),
  appModeBadge: document.getElementById("app-mode-badge"),
  gatheringModeBadge: document.getElementById("gathering-mode-badge"),
  bookCountBadge: document.getElementById("book-count-badge"),
  excerptCountBadge: document.getElementById("excerpt-count-badge"),
  weirdBookCountBadge: document.getElementById("weird-book-count-badge"),
  weirdExcerptCountBadge: document.getElementById("weird-excerpt-count-badge"),
  correctionBookCountBadge: document.getElementById("correction-book-count-badge"),
  correctionExcerptCountBadge: document.getElementById("correction-excerpt-count-badge"),
  graphicsBookCountBadge: document.getElementById("graphics-book-count-badge"),
  graphicsCountBadge: document.getElementById("graphics-count-badge")
};

let currentExcerpts = [];
let currentWeirdExcerpts = [];
let currentCorrectionExcerpts = [];
let currentPendingRecords = [];
let currentGraphicsRecords = [];
let currentGraphicsBookSummaries = [];
let currentGraphicsAssetMatches = new Map();
let currentGraphicsFolderImportPreview = null;
let isSaving = false;
let currentValidationByRecordId = new Map();
let currentModule = "review";
let intakeOptionsLoaded = false;
let currentIntakeCatalogBooks = [];
let currentIntakeCatalogBooksByKey = new Map();
let currentIntakeLegacyBooks = [];
let currentIntakePublishingBooks = [];
let currentIntakePublishingBooksByKey = new Map();
let currentGatheringBookPoems = [];
let googleSheetsTokenClient = null;
let googleSheetsAccessToken = "";
let reviewVisibleCount = 1;
let reviewPinnedRowOrder = [];
let weirdVisibleCount = 25;
let weirdPinnedRowOrder = [];
let currentReviewBookSummaries = [];
let currentWeirdBookSummaries = [];
let reviewBookSummaryByKey = new Map();
let weirdBookSummaryByKey = new Map();
let graphicsBookSummaryByKey = new Map();

const REVIEW_SINGLE_BATCH_SIZE = 1;
const REVIEW_MULTI_BATCH_SIZE = 25;
const EXTRA_REVIEW_BATCH_SIZE = 1;
const GOOGLE_SHEETS_SCOPES = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file";
const INTAKE_MODE_LABELS = {
  book: "Add a quote from a book",
  video: "Add a quote from a video",
  fix: "Fix a quote in one of the quote tools"
};
const SHEET_SOURCE_CONFIG = {
  startRow: 2,
  sourceRange: "A2:BK",
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
    correctedExcerpt: 60,
    validationPrimarySourceFormat: 61
  }
};
const GRAPHICS_SHEET_CONFIG = {
  queueSheetName: "Queue - Needs Graphics",
  cleanupSheetName: "Cleanup - Created Graphics",
  reviewSheetName: "New - Quote Creation Tool Database",
  recordsRange: "A2:I",
  qcStateRange: "AM2:AU"
};
const GRAPHICS_QC_METADATA_OPTIONS = [
  { value: "", label: "No metadata correction" },
  { value: "wrong_title", label: "Wrong title" },
  { value: "wrong_book", label: "Wrong book" },
  { value: "wrong_author", label: "Wrong author" },
  { value: "wrong_excerpt_text", label: "Wrong excerpt text" },
  { value: "missing_element", label: "Missing element" }
];
const GRAPHICS_QC_REJECT_REASON_OPTIONS = [
  { value: "", label: "Choose a reject reason" },
  { value: "mismatched_graphic", label: "Mismatched graphic" },
  { value: "correct_and_recreate", label: "Correct and recreate" },
  { value: "final_reject", label: "Final reject" }
];
const GRAPHICS_QC_AESTHETIC_OPTIONS = [
  { value: "", label: "No aesthetic adjustment" },
  { value: "better_line_breaks", label: "Better line breaks" },
  { value: "different_color_palette", label: "Different color palette" },
  { value: "text_too_small", label: "Text too small" },
  { value: "different_template", label: "Different template" },
  { value: "different_background_image", label: "Different background image" },
  { value: "other", label: "Other" }
];
const GRAPHICS_QC_DEFAULT_NOTES = {
  mismatched_graphic: "Mismatched graphic: the image and excerpt are not the same piece.",
  correct_and_recreate: "Correct and recreate: choose the metadata and/or aesthetic issue, then add any needed details.",
  final_reject: "Final reject: this graphic should not move forward."
};

function normalizeGraphicsQcDecisionClient(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "approve") return "approve";
  if (
    normalized === "mismatch" ||
    normalized === "mismatched_graphic" ||
    normalized === "correct" ||
    normalized === "recreate" ||
    normalized === "correct_and_recreate" ||
    normalized === "reject"
  ) {
    return "reject";
  }
  return "";
}

function getGraphicsQcOptionLabel(options, value) {
  return options.find(option => option.value === value)?.label || "";
}

function renderGraphicsQcOptions(options, selectedValue) {
  return options.map(option => (
    `<option value="${escapeAttribute(option.value)}" ${option.value === selectedValue ? "selected" : ""}>${escapeHtml(option.label)}</option>`
  )).join("");
}

function parseGraphicsQcStructuredNote(note) {
  const parsed = {
    rejectReason: "",
    metadataIssue: "",
    aestheticIssue: "",
    details: ""
  };
  const text = String(note || "").trim();
  if (!text) return parsed;

  const lines = text.split(/\r?\n/).map(line => line.trim()).filter(Boolean);
  const detailLines = [];

  lines.forEach(line => {
    if (line.startsWith("Reject reason: ")) {
      const label = line.slice("Reject reason: ".length).trim();
      const match = GRAPHICS_QC_REJECT_REASON_OPTIONS.find(option => option.label === label);
      if (match) {
        parsed.rejectReason = match.value;
        return;
      }
    }
    if (line.startsWith("Metadata issue: ")) {
      const label = line.slice("Metadata issue: ".length).trim();
      const match = GRAPHICS_QC_METADATA_OPTIONS.find(option => option.label === label);
      if (match) {
        parsed.metadataIssue = match.value;
        return;
      }
    }
    if (line.startsWith("Aesthetic issue: ")) {
      const label = line.slice("Aesthetic issue: ".length).trim();
      const match = GRAPHICS_QC_AESTHETIC_OPTIONS.find(option => option.label === label);
      if (match) {
        parsed.aestheticIssue = match.value;
        return;
      }
    }
    if (line.startsWith("Details: ")) {
      detailLines.push(line.slice("Details: ".length).trim());
      return;
    }
    detailLines.push(line);
  });

  parsed.details = detailLines.join("\n").trim();
  return parsed;
}

function getGraphicsQcRejectReasonFromLegacyDecision(decision, note) {
  const normalizedDecision = String(decision || "").trim().toLowerCase();
  const parsed = parseGraphicsQcStructuredNote(note);
  if (parsed.rejectReason) {
    return parsed.rejectReason;
  }
  if (normalizedDecision === "mismatch" || normalizedDecision === "mismatched_graphic") {
    return "mismatched_graphic";
  }
  if (normalizedDecision === "correct" || normalizedDecision === "recreate" || normalizedDecision === "correct_and_recreate") {
    return "correct_and_recreate";
  }
  if (normalizedDecision === "reject") {
    return "final_reject";
  }
  return "";
}

function buildGraphicsQcNotePayload(decision, rejectReason, metadataIssue, aestheticIssue, detailNote) {
  const normalizedDecision = normalizeGraphicsQcDecisionClient(decision);
  const normalizedRejectReason = String(rejectReason || "").trim();
  const details = String(detailNote || "").trim();

  if (normalizedDecision === "reject") {
    const lines = [];
    const rejectReasonLabel = getGraphicsQcOptionLabel(GRAPHICS_QC_REJECT_REASON_OPTIONS, normalizedRejectReason);
    const metadataLabel = getGraphicsQcOptionLabel(GRAPHICS_QC_METADATA_OPTIONS, metadataIssue);
    const aestheticLabel = getGraphicsQcOptionLabel(GRAPHICS_QC_AESTHETIC_OPTIONS, aestheticIssue);
    if (rejectReasonLabel) lines.push(`Reject reason: ${rejectReasonLabel}`);
    if (metadataLabel) lines.push(`Metadata issue: ${metadataLabel}`);
    if (aestheticLabel) lines.push(`Aesthetic issue: ${aestheticLabel}`);
    if (details) lines.push(`Details: ${details}`);
    return lines.join("\n").trim();
  }

  return details;
}

function setStatus(message, details) {
  elements.statusOutput.textContent = details
    ? `${message}\n\n${JSON.stringify(details, null, 2)}`
    : message;
}

function showSheetLinkModal(url) {
  const existing = document.getElementById("sheet-export-modal");
  if (existing) existing.remove();

  const backdrop = document.createElement("div");
  backdrop.id = "sheet-export-modal";
  backdrop.className = "sheet-modal-backdrop";
  backdrop.innerHTML = `
    <div class="sheet-modal" role="dialog" aria-modal="true" aria-labelledby="sheet-modal-title">
      <h2 id="sheet-modal-title">Your sheet is ready</h2>
      <p>Your browser blocked the new tab, but the export succeeded.</p>
      <p><a href="${escapeAttribute(url)}" target="_blank" rel="noopener noreferrer">${escapeHtml(url)}</a></p>
      <div class="sheet-modal-actions">
        <button id="sheet-modal-open" type="button" class="button">Open sheet</button>
        <button id="sheet-modal-close" type="button" class="button button--secondary">Close</button>
      </div>
    </div>
  `;

  document.body.appendChild(backdrop);
  backdrop.querySelector("#sheet-modal-open")?.addEventListener("click", () => {
    window.open(url, "_blank", "noopener,noreferrer");
  });

  const close = () => backdrop.remove();
  backdrop.querySelector("#sheet-modal-close")?.addEventListener("click", close);
  backdrop.addEventListener("click", event => {
    if (event.target === backdrop) {
      close();
    }
  });
}

function setSubmitState(isBusy, label) {
  isSaving = isBusy;
  elements.submitReview.disabled = isBusy;
  if (elements.submitWeirdReview) {
    elements.submitWeirdReview.disabled = isBusy;
  }
  if (elements.submitCorrections) {
    elements.submitCorrections.disabled = isBusy;
  }
  if (elements.autoApplyCorrections) {
    elements.autoApplyCorrections.disabled = isBusy;
  }
  if (elements.submitGraphicsQc) {
    elements.submitGraphicsQc.disabled = isBusy;
  }
  elements.submitReview.textContent = label || (isBusy ? "Saving..." : "Submit Decisions");
  if (elements.submitWeirdReview) {
    elements.submitWeirdReview.textContent = label || (isBusy ? "Saving..." : "Submit Decisions");
  }
  if (elements.submitCorrections) {
    elements.submitCorrections.textContent = label || (isBusy ? "Saving..." : "Save Corrections");
  }
  if (elements.autoApplyCorrections) {
    elements.autoApplyCorrections.textContent = isBusy ? "Applying..." : "Apply Auto-Fixes";
  }
  if (elements.submitGraphicsQc) {
    elements.submitGraphicsQc.textContent = label || (isBusy ? "Saving..." : "Save QC Decisions");
  }
}

function getSelectedReviewDisplayMode() {
  return elements.reviewDisplayMode?.value || "single";
}

function getReviewBatchSize() {
  return getSelectedReviewDisplayMode() === "batch"
    ? REVIEW_MULTI_BATCH_SIZE
    : REVIEW_SINGLE_BATCH_SIZE;
}

function normalizeBookKey(text) {
  return (text || "")
    .trim()
    .replace(/[’‘]/g, "'")
    .replace(/[“”]/g, "\"")
    .replace(/[–—]/g, "-")
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function cleanSheetWhitespace(text) {
  return (text || "").toString().replace(/\s+/g, " ").trim();
}

function isSheetYes(value) {
  return cleanSheetWhitespace(value).toUpperCase() === "Y";
}

function parseSheetInteger(value) {
  const parsed = parseInt(value, 10);
  return Number.isNaN(parsed) ? 0 : parsed;
}

function getSheetExcerptReviewDecision(row) {
  const config = SHEET_SOURCE_CONFIG.columnMap;
  const explicitDecision = cleanSheetWhitespace(row[config.excerptReviewDecision - 1]).toUpperCase();
  if (explicitDecision) return explicitDecision;

  const statusIndicator = cleanSheetWhitespace(row[config.statusIndicator - 1]).toUpperCase();
  if (statusIndicator === "NEEDS_CORRECTION") return "NEEDS_CORRECTION";

  return "";
}

function isPendingSheetReview(reviewDecision) {
  return !cleanSheetWhitespace(reviewDecision);
}

function buildBrowserCatalogValidationPayload(row) {
  const config = SHEET_SOURCE_CONFIG.columnMap;
  const status = cleanSheetWhitespace(row[config.validationStatus - 1]);
  if (!status) {
    return null;
  }

  const globalBook = cleanSheetWhitespace(row[config.validationGlobalMatchBook - 1]);
  const globalAuthor = cleanSheetWhitespace(row[config.validationGlobalMatchAuthor - 1]);
  const globalPoem = cleanSheetWhitespace(row[config.validationGlobalMatchPoem - 1]);

  return {
    status,
    bookCanonicalTitle: cleanSheetWhitespace(row[config.validationCanonicalBook - 1]),
    bookCanonicalAuthor: cleanSheetWhitespace(row[config.validationCanonicalAuthor - 1]),
    matchedPoemTitle: cleanSheetWhitespace(row[config.validationMatchedPoemTitle - 1]),
    bookPrimarySourceFormat: cleanSheetWhitespace(row[config.validationPrimarySourceFormat - 1]),
    globalExcerptMatch: globalBook || globalAuthor || globalPoem
      ? {
          book_title: globalBook,
          author: globalAuthor,
          poem_title: globalPoem
        }
      : null,
    validatedAt: cleanSheetWhitespace(row[config.validationValidatedAt - 1])
  };
}

function buildPendingRecordFromSheetRow(row, index) {
  const config = SHEET_SOURCE_CONFIG.columnMap;
  const excerptText = (row[config.excerpt - 1] || "").toString();
  const cleanedExcerptText = cleanSheetWhitespace(excerptText);
  const excluded = isSheetYes(row[config.exclude - 1]);
  const reviewDecision = getSheetExcerptReviewDecision(row);
  const bookTitle = cleanSheetWhitespace(row[config.bookTitle - 1]);

  if (!bookTitle || !cleanedExcerptText || excluded || !isPendingSheetReview(reviewDecision)) {
    return null;
  }

  return {
    sourceRow: SHEET_SOURCE_CONFIG.startRow + index,
    recordId: (row[config.recordId - 1] || "").toString(),
    author: (row[config.author - 1] || "").toString(),
    title: (row[config.title - 1] || "").toString(),
    bookTitle,
    excerptText,
    wordCount: countWordsFromText(excerptText),
    approved: (row[config.approved - 1] || "").toString(),
    quoteCreatedQc: (row[config.quoteCreatedQc - 1] || "").toString(),
    correctionNote: (row[config.correctionNote - 1] || "").toString(),
    duplicateGroupId: (row[config.duplicateGroupId - 1] || "").toString(),
    exactPullCount: parseSheetInteger(row[config.exactPullCount - 1]),
    excerptReviewDecision: reviewDecision,
    useForInt: isSheetYes(row[config.useForInt - 1]),
    bookPrimarySourceFormat: cleanSheetWhitespace(row[config.validationPrimarySourceFormat - 1]),
    catalogValidation: buildBrowserCatalogValidationPayload(row)
  };
}

function buildExcerptPayloadFromSheetRow(row, index) {
  const config = SHEET_SOURCE_CONFIG.columnMap;
  const rawExcerptText = (row[config.excerpt - 1] || "").toString();
  const correctedExcerptText = (row[config.correctedExcerpt - 1] || "").toString();
  const excerptText = correctedExcerptText || rawExcerptText;
  const cleanedExcerptText = cleanSheetWhitespace(excerptText);
  const approved = (row[config.approved - 1] || "").toString();
  const reviewDecision = getSheetExcerptReviewDecision(row);
  const rawAuthor = (row[config.author - 1] || "").toString();
  const rawTitle = (row[config.title - 1] || "").toString();
  const rawBookTitle = cleanSheetWhitespace(row[config.bookTitle - 1]);
  const correctedAuthor = (row[config.correctedAuthor - 1] || "").toString();
  const correctedTitle = (row[config.correctedTitle - 1] || "").toString();
  const correctedBookTitle = (row[config.correctedBookTitle - 1] || "").toString();

  if (!cleanedExcerptText || !rawBookTitle) {
    return null;
  }

  return {
    sourceRow: SHEET_SOURCE_CONFIG.startRow + index,
    recordId: (row[config.recordId - 1] || "").toString(),
    author: correctedAuthor || rawAuthor,
    title: correctedTitle || rawTitle,
    bookTitle: cleanSheetWhitespace(correctedBookTitle || rawBookTitle),
    excerptText,
    wordCount: countWordsFromText(cleanedExcerptText),
    approved,
    statusIndicator: cleanSheetWhitespace(row[config.statusIndicator - 1]),
    quoteCreatedQc: (row[config.quoteCreatedQc - 1] || "").toString(),
    correctionNote: (row[config.correctionNote - 1] || "").toString(),
    excludeRaw: (row[config.exclude - 1] || "").toString(),
    duplicateGroupId: (row[config.duplicateGroupId - 1] || "").toString(),
    exactPullCount: parseSheetInteger(row[config.exactPullCount - 1]),
    pending: isPendingSheetReview(reviewDecision),
    excerptReviewDecision: reviewDecision,
    useForQi: approved === "Y",
    useForInt: isSheetYes(row[config.useForInt - 1]),
    useForGraphicsQi: approved === "Y",
    useForPhotos: isSheetYes(row[config.useForInt - 1]),
    correctedAuthor,
    correctedTitle,
    correctedBookTitle,
    correctedExcerpt: correctedExcerptText,
    rawAuthor,
    rawTitle,
    rawBookTitle,
    rawExcerptText,
    bookPrimarySourceFormat: cleanSheetWhitespace(row[config.validationPrimarySourceFormat - 1]),
    catalogValidation: buildBrowserCatalogValidationPayload(row)
  };
}

function summarizeSimpleBooks(records) {
  const counts = new Map();
  records.forEach(record => {
    const bookTitle = cleanSheetWhitespace(record.bookTitle);
    const bookKey = normalizeBookKey(bookTitle);
    if (!bookKey) return;
    if (!counts.has(bookKey)) {
      counts.set(bookKey, { title: bookTitle, count: 0 });
    }
    const summary = counts.get(bookKey);
    summary.title = choosePreferredBookTitle(summary.title, bookTitle);
    summary.count += 1;
  });
  return Array.from(counts.entries())
    .map(([key, summary]) => ({ key, title: summary.title, count: summary.count }))
    .sort((left, right) => left.title.localeCompare(right.title));
}

async function loadCorrectionBooksFromSheetsFallback() {
  const sheetName = runtimeConfig.sourceSheetName || SHEET_SOURCE_CONFIG.sourceRange;
  const range = `'${sheetName.replace(/'/g, "''")}'!${SHEET_SOURCE_CONFIG.sourceRange}`;
  const values = await fetchSheetValues(range);
  const config = SHEET_SOURCE_CONFIG.columnMap;
  const records = values
    .map((row, index) => {
      const record = buildExcerptPayloadFromSheetRow(row, index);
      if (!record) return null;
      if (isSheetYes(row[config.exclude - 1])) return null;
      if (record.excerptReviewDecision !== "NEEDS_CORRECTION") return null;
      return record;
    })
    .filter(Boolean);

  return {
    ok: true,
    version: `${runtimeConfig.appVersion || "unknown"}-sheets-fallback`,
    books: summarizeSimpleBooks(records)
  };
}

async function loadCorrectionsFromSheetsFallback(bookTitle) {
  const bookKey = normalizeBookKey(bookTitle);
  if (!bookKey) {
    return { ok: true, version: `${runtimeConfig.appVersion || "unknown"}-sheets-fallback`, bookTitle: "", excerpts: [] };
  }

  const sheetName = runtimeConfig.sourceSheetName || SHEET_SOURCE_CONFIG.sourceRange;
  const range = `'${sheetName.replace(/'/g, "''")}'!${SHEET_SOURCE_CONFIG.sourceRange}`;
  const values = await fetchSheetValues(range);
  const excerpts = [];
  let preferredBookTitle = cleanSheetWhitespace(bookTitle);

  values.forEach((row, index) => {
    const record = buildExcerptPayloadFromSheetRow(row, index);
    if (!record) return;
    if (record.excerptReviewDecision !== "NEEDS_CORRECTION") return;
    if (normalizeBookKey(record.bookTitle) !== bookKey) return;
    preferredBookTitle = choosePreferredBookTitle(preferredBookTitle, record.bookTitle);
    excerpts.push(record);
  });

  return {
    ok: true,
    version: `${runtimeConfig.appVersion || "unknown"}-sheets-fallback`,
    bookTitle: preferredBookTitle,
    excerpts
  };
}

function getGraphicsSheetName(mode) {
  return mode === "cleanup"
    ? GRAPHICS_SHEET_CONFIG.cleanupSheetName
    : GRAPHICS_SHEET_CONFIG.queueSheetName;
}

async function loadGraphicsQcStateMapFromSheetsFallback() {
  const range = `'${GRAPHICS_SHEET_CONFIG.reviewSheetName.replace(/'/g, "''")}'!${GRAPHICS_SHEET_CONFIG.qcStateRange}`;
  const values = await fetchSheetValues(range);
  const state = new Map();
  values.forEach(row => {
    const recordId = cleanSheetWhitespace(row[0]);
    if (!recordId) return;
    state.set(recordId, {
      decision: cleanSheetWhitespace(row[6]),
      note: (row[7] || "").toString(),
      updatedAt: cleanSheetWhitespace(row[8])
    });
  });
  return state;
}

async function loadGraphicsBooksFromSheetsFallback(mode) {
  const sheetName = getGraphicsSheetName(mode);
  const range = `'${sheetName.replace(/'/g, "''")}'!${GRAPHICS_SHEET_CONFIG.recordsRange}`;
  const values = await fetchSheetValues(range);
  const records = values.map(row => ({
    author: (row[0] || "").toString(),
    poemTitle: (row[1] || "").toString(),
    bookTitle: cleanSheetWhitespace(row[2]),
    quoteText: (row[3] || "").toString(),
    notes: (row[4] || "").toString(),
    approved: cleanSheetWhitespace(row[5]),
    created: cleanSheetWhitespace(row[6]),
    workflowStatus: cleanSheetWhitespace(row[7]),
    recordId: (row[8] || "").toString()
  })).filter(record => record.bookTitle);

  return {
    ok: true,
    version: `${runtimeConfig.appVersion || "unknown"}-sheets-fallback`,
    mode,
    books: summarizeSimpleBooks(records)
  };
}

async function loadGraphicsRecordsFromSheetsFallback(bookTitle, mode) {
  const bookKey = normalizeBookKey(bookTitle);
  if (!bookKey) {
    return { ok: true, version: `${runtimeConfig.appVersion || "unknown"}-sheets-fallback`, mode, bookTitle: "", records: [] };
  }

  const sheetName = getGraphicsSheetName(mode);
  const range = `'${sheetName.replace(/'/g, "''")}'!${GRAPHICS_SHEET_CONFIG.recordsRange}`;
  const [values, qcState] = await Promise.all([
    fetchSheetValues(range),
    loadGraphicsQcStateMapFromSheetsFallback()
  ]);

  const records = [];
  let preferredBookTitle = cleanSheetWhitespace(bookTitle);

  values.forEach((row, index) => {
    const currentBookTitle = cleanSheetWhitespace(row[2]);
    if (normalizeBookKey(currentBookTitle) !== bookKey) return;

    preferredBookTitle = choosePreferredBookTitle(preferredBookTitle, currentBookTitle);
    const recordId = (row[8] || "").toString();
    const qc = qcState.get(recordId) || {};
    records.push({
      sheetRow: index + 2,
      author: (row[0] || "").toString(),
      poemTitle: (row[1] || "").toString(),
      bookTitle: choosePreferredBookTitle(preferredBookTitle, currentBookTitle),
      quoteText: (row[3] || "").toString(),
      notes: (row[4] || "").toString(),
      approved: cleanSheetWhitespace(row[5]),
      created: cleanSheetWhitespace(row[6]),
      workflowStatus: cleanSheetWhitespace(row[7]),
      recordId,
      graphicsQcDecision: cleanSheetWhitespace(qc.decision),
      graphicsQcNote: qc.note || "",
      graphicsQcUpdatedAt: cleanSheetWhitespace(qc.updatedAt)
    });
  });

  return {
    ok: true,
    version: `${runtimeConfig.appVersion || "unknown"}-sheets-fallback`,
    mode,
    bookTitle: preferredBookTitle,
    records
  };
}

function getReviewQueueIncludeSet() {
  const titles = Array.isArray(runtimeConfig.reviewQueueIncludeTitles)
    ? runtimeConfig.reviewQueueIncludeTitles
    : [];
  return new Set(titles.map(normalizeBookKey).filter(Boolean));
}

function getVisibleReviewBookSummaries() {
  const reviewQueueIncludeSet = getReviewQueueIncludeSet();
  if (getSelectedReviewFilter() !== "current_titles" || !reviewQueueIncludeSet.size) {
    return currentReviewBookSummaries;
  }
  return currentReviewBookSummaries.filter(book => reviewQueueIncludeSet.has(book.key));
}

function refreshReviewBookSelect(preserveSelection = true) {
  const previousSelection = preserveSelection ? elements.bookSelect?.value || "" : "";
  const visibleBooks = getVisibleReviewBookSummaries();
  populateBookSelect(elements.bookSelect, visibleBooks, previousSelection, book => `${book.title} (${book.standardCount})`);
  elements.bookCountBadge.textContent = `${visibleBooks.length} Books`;
}

function getBookTitleDisplayScore(title) {
  const text = (title || "").trim().replace(/\s+/g, " ");
  if (!text) return -Infinity;

  let score = 0;
  if (text === (title || "")) score += 2;
  if (/[a-z]/.test(text)) score += 3;
  if (/^[A-Z0-9\s&'"?!:;.,()\/|-]+$/.test(text) && !/[a-z]/.test(text)) score -= 2;
  score -= Math.max(0, text.length - text.trim().length);
  return score;
}

function choosePreferredBookTitle(currentTitle, candidateTitle) {
  if (!currentTitle) return candidateTitle;
  if (!candidateTitle) return currentTitle;

  const currentBase = getIntakeBookBaseTitle(currentTitle);
  const candidateBase = getIntakeBookBaseTitle(candidateTitle);
  if (candidateBase === candidateTitle && currentBase !== currentTitle) return candidateTitle;
  if (currentBase === currentTitle && candidateBase !== candidateTitle) return currentTitle;

  const currentScore = getBookTitleDisplayScore(currentTitle);
  const candidateScore = getBookTitleDisplayScore(candidateTitle);
  if (candidateScore > currentScore) return candidateTitle;
  if (candidateScore < currentScore) return currentTitle;

  return candidateTitle.length < currentTitle.length ? candidateTitle : currentTitle;
}

function getIntakeBookBaseTitle(title) {
  const cleanedTitle = cleanSheetWhitespace(title);
  if (!cleanedTitle) return "";

  if (cleanedTitle.includes(":")) {
    const colonBase = cleanedTitle.split(/\s*:\s*/, 1)[0]?.trim() || "";
    if (colonBase) {
      return colonBase;
    }
  }

  const editionMatch = cleanedTitle.match(/^(.*?)\s*-\s*(limited edition|special edition|re-?release)$/i);
  if (editionMatch?.[1]) {
    return cleanSheetWhitespace(editionMatch[1]);
  }

  const parenEditionMatch = cleanedTitle.match(/^(.*?)\s*\((limited edition|special edition|re-?release)\)$/i);
  if (parenEditionMatch?.[1]) {
    return cleanSheetWhitespace(parenEditionMatch[1]);
  }

  return cleanedTitle;
}

function buildIntakeBookMergeKeys(title) {
  const cleanedTitle = cleanSheetWhitespace(title);
  if (!cleanedTitle) return [];

  const keys = new Set([normalizeBookKey(cleanedTitle)]);
  const baseTitle = getIntakeBookBaseTitle(cleanedTitle);
  if (baseTitle && baseTitle !== cleanedTitle) {
    keys.add(normalizeBookKey(baseTitle));
  }

  return Array.from(keys).filter(Boolean);
}

function indexBookSummariesByKey(summaries) {
  return new Map(summaries.map(summary => [summary.key, summary]));
}

function getPendingRecordsForBookKey(bookKey, records = currentPendingRecords) {
  return records.filter(record => normalizeBookKey(record.bookTitle) === bookKey);
}

async function requestMergedBookRecords(sourceExcerpts) {
  const rawBookTitles = Array.from(
    new Set(
      sourceExcerpts
        .map(excerpt => normalizeCorrectionNote(excerpt.bookTitle))
        .filter(Boolean)
    )
  );

  if (!rawBookTitles.length) {
    return { ok: true, records: [] };
  }

  const responses = await Promise.all(
    rawBookTitles.map(bookTitle => requestReviewApi("/api/review/excerpts", { bookTitle }))
  );

  const failed = responses.find(response => !response.ok);
  if (failed) {
    throw new Error(failed.error || "Book verification reload failed.");
  }

  const records = [];
  const seenSourceRows = new Set();

  responses.forEach(response => {
    const excerpts = Array.isArray(response.excerpts) ? response.excerpts : [];
    excerpts.forEach(excerpt => {
      const sourceRow = Number(excerpt.sourceRow);
      if (sourceRow && seenSourceRows.has(sourceRow)) {
        return;
      }
      if (sourceRow) {
        seenSourceRows.add(sourceRow);
      }
      records.push(excerpt);
    });
  });

  return { ok: true, records };
}

function refreshBookCountsInBackground() {
  requestReviewApi("/api/review/pending-records")
    .then(data => {
      if (!data.ok) return;
      const records = Array.isArray(data.records) ? data.records : [];
      applyPendingBookData(records, { preserveSelection: true });
    })
    .catch(() => {
      // Background refresh is best-effort only.
    });
}

function applyPendingBookData(records, { preserveSelection = false } = {}) {
  const previousWeirdSelection = preserveSelection ? elements.weirdBookSelect?.value || "" : "";

  currentPendingRecords = records;
  const allBookSummaries = summarizePendingBooks(records);
  currentReviewBookSummaries = allBookSummaries.filter(book => book.standardCount > 0);
  currentWeirdBookSummaries = allBookSummaries.filter(book => book.needsCheckingCount > 0);
  reviewBookSummaryByKey = indexBookSummariesByKey(currentReviewBookSummaries);
  weirdBookSummaryByKey = indexBookSummariesByKey(currentWeirdBookSummaries);

  refreshReviewBookSelect(preserveSelection);
  populateBookSelect(elements.weirdBookSelect, currentWeirdBookSummaries, previousWeirdSelection, book => `${book.title} (${book.needsCheckingCount})`);

  if (elements.weirdBookCountBadge) {
    elements.weirdBookCountBadge.textContent = `${currentWeirdBookSummaries.length} Books`;
  }

  return allBookSummaries;
}

async function loadRuntimeConfig() {
  try {
    const response = await fetch("/api/bootstrap", {
      headers: {
        Accept: "application/json"
      }
    });
    if (!response.ok) {
      throw new Error(`/api/bootstrap returned ${response.status}`);
    }
    const data = await response.json();
    runtimeConfig = {
      ...runtimeConfig,
      ...data
    };
    window.WEAVER_CONFIG = runtimeConfig;
  } catch (error) {
    setStatus(`Bootstrap config load failed: ${error.message}`);
  }
}

function applyRuntimeMode() {
  if (elements.appModeBadge) {
    elements.appModeBadge.textContent = "Hosted";
  }
}

function getSelectedGatheringMode() {
  return elements.gatheringMode?.value || "book";
}

function updateGatheringModeUi() {
  const mode = getSelectedGatheringMode();
  elements.gatheringBookFields?.toggleAttribute("hidden", mode !== "book");
  elements.gatheringVideoFields?.toggleAttribute("hidden", mode !== "video");
  elements.gatheringFixFields?.toggleAttribute("hidden", mode !== "fix");
  elements.gatheringTabBook?.classList.toggle("gathering-tab--active", mode === "book");
  elements.gatheringTabVideo?.classList.toggle("gathering-tab--active", mode === "video");
  elements.gatheringTabFix?.classList.toggle("gathering-tab--active", mode === "fix");
  if (elements.gatheringModeBadge) {
    elements.gatheringModeBadge.textContent = (
      mode === "book" ? "Book Excerpts" :
      mode === "video" ? "Video Excerpts" :
      "Fix Existing Quote"
    );
  }
  updateGatheringQuoteMeta();
}

function setGatheringMode(mode) {
  if (!elements.gatheringMode) return;
  elements.gatheringMode.value = mode;
  updateGatheringModeUi();
}

function setActiveModule(moduleName) {
  currentModule = ["gathering", "review", "weird", "corrections", "graphics"].includes(moduleName) ? moduleName : "review";
  elements.gatheringModule?.classList.toggle("module-panel--active", currentModule === "gathering");
  elements.reviewModule?.classList.toggle("module-panel--active", currentModule === "review");
  elements.reviewQueuePanel?.classList.toggle("module-panel--active", currentModule === "review");
  elements.weirdModule?.classList.toggle("module-panel--active", currentModule === "weird");
  elements.weirdQueuePanel?.classList.toggle("module-panel--active", currentModule === "weird");
  elements.correctionsModule?.classList.toggle("module-panel--active", currentModule === "corrections");
  elements.graphicsModule?.classList.toggle("module-panel--active", currentModule === "graphics");
  elements.showGatheringModule?.classList.toggle("hero-pill--active", currentModule === "gathering");
  elements.showReviewModule?.classList.toggle("hero-pill--active", currentModule === "review");
  elements.showWeirdModule?.classList.toggle("hero-pill--active", currentModule === "weird");
  elements.showCorrectionsModule?.classList.toggle("hero-pill--active", currentModule === "corrections");
  elements.showGraphicsModule?.classList.toggle("hero-pill--active", currentModule === "graphics");
}

async function requestReviewApi(path, params = {}) {
  const url = new URL(path, window.location.origin);
  Object.entries(params).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== "") {
      url.searchParams.set(key, String(value));
    }
  });

  const response = await fetch(url, {
    headers: {
      Accept: "application/json"
    }
  });

  let data;
  try {
    data = await response.json();
  } catch (_error) {
    throw new Error(`${path} returned unreadable JSON.`);
  }

  if (!response.ok || !data.ok) {
    throw new Error(data.error || `${path} failed.`);
  }

  return data;
}

async function postReviewApi(path, payload) {
  const response = await fetch(path, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json"
    },
    body: JSON.stringify(payload || {})
  });

  let data;
  try {
    data = await response.json();
  } catch (_error) {
    throw new Error(`${path} returned unreadable JSON.`);
  }

  if (!response.ok || !data.ok) {
    throw new Error(data.error || `${path} failed.`);
  }

  return data;
}

function populateDatalist(element, values) {
  if (!element) return;
  element.innerHTML = values
    .filter(Boolean)
    .map(value => `<option value="${escapeAttribute(value)}"></option>`)
    .join("");
}

function mergeIntakeBookSuggestions({ legacyBooks = [], publishingBooks = [], catalogBooks = [] } = {}) {
  const byKey = new Map();

  function upsert(title, sourcePriority) {
    const cleanedTitle = cleanSheetWhitespace(title);
    const keys = buildIntakeBookMergeKeys(cleanedTitle);
    if (!keys.length) return;

    let current = null;
    for (const key of keys) {
      if (byKey.has(key)) {
        current = byKey.get(key);
        break;
      }
    }

    if (!current) {
      current = {
        title: cleanedTitle,
        sourcePriority
      };
    }

    const preferredTitle = choosePreferredBookTitle(current.title, cleanedTitle);
    const preferredPriority = Math.max(current.sourcePriority, sourcePriority);
    const mergedEntry = {
      title: preferredPriority > current.sourcePriority
        ? cleanedTitle
        : preferredTitle,
      sourcePriority: preferredPriority
    };

    keys.forEach(key => {
      byKey.set(key, mergedEntry);
    });
  }

  legacyBooks.forEach(title => upsert(title, 1));
  publishingBooks.forEach(book => upsert(book?.title, 2));
  catalogBooks.forEach(book => upsert(book?.title, 3));

  return Array.from(new Set(Array.from(byKey.values())))
    .map(entry => entry.title)
    .sort((left, right) => left.localeCompare(right));
}

async function loadGatheringOptions({ force = false } = {}) {
  if (intakeOptionsLoaded && !force) {
    return;
  }

  setStatus("Loading excerpt gathering intake options...");
  const data = await requestReviewApi("/api/intake/options");
  populateDatalist(elements.gatheringAuthorOptions, data.authors || []);
  currentIntakeLegacyBooks = Array.isArray(data.books) ? data.books : [];
  currentIntakeCatalogBooks = Array.isArray(data.catalogBooks) ? data.catalogBooks : [];
  currentIntakeCatalogBooksByKey = new Map(
    currentIntakeCatalogBooks.map(book => [normalizeBookKey(book.title), book])
  );
  currentIntakePublishingBooks = Array.isArray(data.publishingBooks) ? data.publishingBooks : [];
  currentIntakePublishingBooksByKey = new Map(
    currentIntakePublishingBooks.map(book => [normalizeBookKey(book.title), book])
  );
  populateDatalist(
    elements.gatheringBookOptions,
    mergeIntakeBookSuggestions({
      legacyBooks: currentIntakeLegacyBooks,
      publishingBooks: currentIntakePublishingBooks,
      catalogBooks: currentIntakeCatalogBooks
    })
  );
  intakeOptionsLoaded = true;
  const publishingStatusNote = data.publishingBooksError
    ? " Publishing-order titles are temporarily unavailable until that sheet is shared with Weaver."
    : "";
  setStatus(`Loaded ${data.authors?.length || 0} legacy author suggestions, ${currentIntakeLegacyBooks.length} legacy book suggestions, ${currentIntakePublishingBooks.length} publishing-order books, and ${currentIntakeCatalogBooks.length} EPUB-backed catalog books for intake.${publishingStatusNote}`);
}

async function loadGatheringPoemsForBook(bookTitle, { preserveTitle = false } = {}) {
  const cleanedBookTitle = (bookTitle || "").trim();
  if (!cleanedBookTitle || !elements.gatheringBookTitle) {
    populateDatalist(elements.gatheringBookPoemOptions, []);
    currentGatheringBookPoems = [];
    if (elements.gatheringBookSourceHint) {
      elements.gatheringBookSourceHint.textContent = "Select a book to load poem titles from the catalog/EPUB source, or keep going manually for non-EPUB titles.";
    }
    updateGatheringCatalogPreviewState();
    return;
  }

  const previousValue = preserveTitle ? elements.gatheringBookTitle.value.trim() : "";
  const data = await requestReviewApi("/api/intake/catalog/poems", { bookTitle: cleanedBookTitle });
  const poems = Array.isArray(data.poems) ? data.poems : [];
  currentGatheringBookPoems = poems;
  populateDatalist(elements.gatheringBookPoemOptions, poems);
  if (previousValue && poems.includes(previousValue)) {
    elements.gatheringBookTitle.value = previousValue;
  } else if (!preserveTitle) {
    elements.gatheringBookTitle.value = "";
  }
  if (elements.gatheringBookBook) {
    elements.gatheringBookBook.value = data.bookTitle || cleanedBookTitle;
  }
  if (elements.gatheringBookAuthor) {
    elements.gatheringBookAuthor.value = data.author || "";
  }
  if (elements.gatheringBookSourceHint) {
    elements.gatheringBookSourceHint.textContent = `${data.primarySourceFormat || "Catalog"} source connected. ${poems.length} poem titles loaded for ${data.bookTitle || cleanedBookTitle}. You can also type a manual poem title if needed.`;
  }
  updateGatheringCatalogPreviewState();
}

function updateGatheringCatalogPreviewState() {
  const bookTitle = elements.gatheringBookBook?.value.trim() || "";
  const poemTitle = elements.gatheringBookTitle?.value.trim() || "";
  const hasCatalogBook = !!currentIntakeCatalogBooksByKey.get(normalizeBookKey(bookTitle));
  const currentIndex = currentGatheringBookPoems.indexOf(poemTitle);
  const hasIndexedPoem = currentIndex >= 0;

  if (elements.gatheringViewCatalogPoem) {
    elements.gatheringViewCatalogPoem.disabled = !(hasCatalogBook && bookTitle && poemTitle);
  }
  if (elements.gatheringPrevCatalogPoem) {
    elements.gatheringPrevCatalogPoem.disabled = !(hasCatalogBook && hasIndexedPoem && currentIndex > 0);
  }
  if (elements.gatheringNextCatalogPoem) {
    elements.gatheringNextCatalogPoem.disabled = !(hasCatalogBook && hasIndexedPoem && currentIndex < currentGatheringBookPoems.length - 1);
  }
}

function setGatheringBookManualMode(bookTitle) {
  populateDatalist(elements.gatheringBookPoemOptions, []);
  currentGatheringBookPoems = [];
  if (elements.gatheringBookSourceHint) {
    const cleanedBookTitle = cleanSheetWhitespace(bookTitle);
    const normalizedBook = normalizeBookKey(cleanedBookTitle);
    const isLegacySuggested = currentIntakeLegacyBooks.some(value => normalizeBookKey(value) === normalizedBook);
    const publishingBook = currentIntakePublishingBooksByKey.get(normalizedBook);
    if (publishingBook && elements.gatheringBookAuthor && !elements.gatheringBookAuthor.value.trim()) {
      elements.gatheringBookAuthor.value = publishingBook.author || "";
    }
    if (publishingBook) {
      const catalogLabel = publishingBook.releaseCatalog || "Publishing order";
      elements.gatheringBookSourceHint.textContent = `${cleanedBookTitle} is coming from the publishing-order sheet (${catalogLabel}) and does not have EPUB support yet. Author is prefilled when available; enter the poem title manually.`;
    } else if (cleanedBookTitle && isLegacySuggested) {
      elements.gatheringBookSourceHint.textContent = `${cleanedBookTitle} is available from the older intake suggestions, but not from an EPUB-backed catalog source yet. Enter the poem title and author manually.`;
    } else if (cleanedBookTitle) {
      elements.gatheringBookSourceHint.textContent = `${cleanedBookTitle} is not EPUB-backed in Weaver yet. You can still enter the author, poem title, and quote manually.`;
    } else {
      elements.gatheringBookSourceHint.textContent = "Select a book to load poem titles from the catalog/EPUB source, or keep going manually for non-EPUB titles.";
    }
  }
  if (elements.gatheringViewCatalogPoem) {
    updateGatheringCatalogPreviewState();
  }
}

function navigateGatheringCatalogPoem(direction) {
  if (!currentGatheringBookPoems.length || !elements.gatheringBookTitle) {
    return;
  }

  const currentTitle = elements.gatheringBookTitle.value.trim();
  const currentIndex = currentGatheringBookPoems.indexOf(currentTitle);
  if (currentIndex < 0) {
    return;
  }

  const nextIndex = currentIndex + direction;
  if (nextIndex < 0 || nextIndex >= currentGatheringBookPoems.length) {
    return;
  }

  elements.gatheringBookTitle.value = currentGatheringBookPoems[nextIndex];
  updateGatheringCatalogPreviewState();
}

async function handleGatheringBookSelectionChange({ preserveTitle = false } = {}) {
  const bookTitle = elements.gatheringBookBook?.value.trim() || "";
  if (!bookTitle) {
    if (elements.gatheringBookAuthor) elements.gatheringBookAuthor.value = "";
    if (!preserveTitle && elements.gatheringBookTitle) elements.gatheringBookTitle.value = "";
    setGatheringBookManualMode("");
    return;
  }

  const catalogBook = currentIntakeCatalogBooksByKey.get(normalizeBookKey(bookTitle));
  if (!catalogBook) {
    if (!preserveTitle && elements.gatheringBookTitle) elements.gatheringBookTitle.value = "";
    setGatheringBookManualMode(bookTitle);
    return;
  }

  await loadGatheringPoemsForBook(catalogBook.title, { preserveTitle });
}

function resetGatheringForm() {
  [
    elements.gatheringBookAuthor,
    elements.gatheringBookBook,
    elements.gatheringBookQuote,
    elements.gatheringBookNotes,
    elements.gatheringBookItalics,
    elements.gatheringVideoAuthor,
    elements.gatheringVideoTitle,
    elements.gatheringVideoBook,
    elements.gatheringVideoEvent,
    elements.gatheringVideoQuote,
    elements.gatheringFixPart,
    elements.gatheringFixAuthor,
    elements.gatheringFixIncorrect,
    elements.gatheringFixCorrect
  ].forEach(field => {
    if (field) field.value = "";
  });
  populateDatalist(elements.gatheringBookPoemOptions, []);
  setGatheringBookManualMode("");
  updateGatheringCatalogPreviewState();
  updateGatheringQuoteMeta();
}

function resetGatheringAfterSubmit(mode) {
  if (mode === "book") {
    if (elements.gatheringBookQuote) elements.gatheringBookQuote.value = "";
    if (elements.gatheringBookNotes) elements.gatheringBookNotes.value = "";
    if (elements.gatheringBookItalics) elements.gatheringBookItalics.value = "";
    if (elements.gatheringBookReaction) elements.gatheringBookReaction.value = "";
    updateGatheringCatalogPreviewState();
    updateGatheringQuoteMeta();
    elements.gatheringBookQuote?.focus();
    return;
  }

  if (mode === "video") {
    if (elements.gatheringVideoQuote) elements.gatheringVideoQuote.value = "";
    updateGatheringQuoteMeta();
    elements.gatheringVideoQuote?.focus();
    return;
  }

  if (mode === "fix") {
    if (elements.gatheringFixIncorrect) elements.gatheringFixIncorrect.value = "";
    if (elements.gatheringFixCorrect) elements.gatheringFixCorrect.value = "";
    elements.gatheringFixIncorrect?.focus();
    return;
  }

  resetGatheringForm();
}

function describeGatheringQuoteRange(wordCount) {
  if (wordCount <= 0) return "Word count: 0. Rough sweet spot: 10–25 words.";
  if (wordCount < 10) return `Word count: ${wordCount}. This is on the short side for the usual target range.`;
  if (wordCount <= 25) return `Word count: ${wordCount}. This sits in the target 10–25 word range.`;
  if (wordCount <= 40) return `Word count: ${wordCount}. Slightly long, but still workable if the line really earns it.`;
  return `Word count: ${wordCount}. This is long, so it’s worth sanity-checking before you keep it.`;
}

function updateGatheringQuoteMeta() {
  if (elements.gatheringBookQuoteMeta) {
    elements.gatheringBookQuoteMeta.textContent = describeGatheringQuoteRange(
      countWordsFromText(elements.gatheringBookQuote?.value || "")
    );
  }
  if (elements.gatheringVideoQuoteMeta) {
    elements.gatheringVideoQuoteMeta.textContent = describeGatheringQuoteRange(
      countWordsFromText(elements.gatheringVideoQuote?.value || "")
    );
  }
}

function handleGatheringQuickSubmit(event) {
  if ((event.metaKey || event.ctrlKey) && event.key === "Enter") {
    event.preventDefault();
    submitGathering();
  }
}

function buildGatheringPayload() {
  const mode = getSelectedGatheringMode();
  const email = elements.gatheringEmail?.value.trim() || "";
  if (!email) {
    throw new Error("Add an email before submitting an excerpt gathering row.");
  }

  if (mode === "book") {
    const author = elements.gatheringBookAuthor?.value.trim() || "";
    const title = elements.gatheringBookTitle?.value.trim() || "";
    const quote = elements.gatheringBookQuote?.value.trim() || "";
    const bookTitle = elements.gatheringBookBook?.value.trim() || "";
    const notes = elements.gatheringBookNotes?.value.trim() || "";
    const italicsMarkup = elements.gatheringBookItalics?.value.trim() || "";
    const reaction = elements.gatheringBookReaction?.value.trim() || "";
    if (!author || !title || !quote) {
      throw new Error("Book intake needs an author, poem title, and quote.");
    }
    const combinedNotes = [
      notes,
      italicsMarkup ? `Italics markup: ${italicsMarkup}` : "",
      reaction ? `Full-poem reaction: ${reaction}` : ""
    ]
      .filter(Boolean)
      .join("\n\n");
    return {
      mode,
      email,
      author,
      title,
      quote,
      bookTitle,
      notes: combinedNotes
    };
  }

  if (mode === "video") {
    const author = elements.gatheringVideoAuthor?.value.trim() || "";
    const title = elements.gatheringVideoTitle?.value.trim() || "";
    const quote = elements.gatheringVideoQuote?.value.trim() || "";
    const bookTitle = elements.gatheringVideoBook?.value.trim() || "";
    const eventName = elements.gatheringVideoEvent?.value.trim() || "";
    if (!author || !quote) {
      throw new Error("Video intake needs an author and quote.");
    }
    return {
      mode,
      email,
      author,
      title,
      quote,
      bookTitle,
      eventName
    };
  }

  const wrongPart = elements.gatheringFixPart?.value.trim() || "";
  const incorrectText = elements.gatheringFixIncorrect?.value.trim() || "";
  const correctedText = elements.gatheringFixCorrect?.value.trim() || "";
  const author = elements.gatheringFixAuthor?.value.trim() || "";
  if (!wrongPart || !incorrectText || !correctedText) {
    throw new Error("Fix intake needs the wrong-part label, incorrect text, and corrected text.");
  }
  return {
    mode,
    email,
    wrongPart,
    incorrectText,
    correctedText,
    author
  };
}

async function submitGathering() {
  try {
    const payload = buildGatheringPayload();
    setStatus(`Submitting ${INTAKE_MODE_LABELS[payload.mode] || "excerpt gathering"} row...`);
    const result = await postReviewApi("/api/intake/submit", payload);
    resetGatheringAfterSubmit(payload.mode);
    setStatus(`Saved excerpt gathering row ${result.rowNumber}.`, result);
    if (currentModule === "review") {
      loadBooks();
    }
  } catch (error) {
    setStatus(`Excerpt gathering submit failed: ${error.message}`);
  }
}

function openGatheringCatalogPoem() {
  const bookTitle = elements.gatheringBookBook?.value.trim() || "";
  const poemTitle = elements.gatheringBookTitle?.value.trim() || "";
  const excerptText = elements.gatheringBookQuote?.value.trim() || "";
  if (!bookTitle || !poemTitle) {
    setStatus("Choose a catalog-backed book and poem before opening EPUB context.");
    return;
  }
  const url = new URL("/catalog-poem", window.location.origin);
  url.searchParams.set("bookTitle", bookTitle);
  url.searchParams.set("poemTitle", poemTitle);
  if (excerptText) {
    url.searchParams.set("excerptText", excerptText);
  }
  openComparisonWindow(url.toString(), "weaverGatheringCatalogPoem");
}

function ensureGoogleSheetsTokenClient() {
  if (googleSheetsTokenClient) {
    return googleSheetsTokenClient;
  }

  const clientId = runtimeConfig.googleOAuthClientId || "";
  if (!clientId) {
    throw new Error("Google Sheets export is not configured yet.");
  }

  if (!window.google?.accounts?.oauth2) {
    throw new Error("Google export is still loading. Try again in a moment.");
  }

  googleSheetsTokenClient = window.google.accounts.oauth2.initTokenClient({
    client_id: clientId,
    scope: GOOGLE_SHEETS_SCOPES,
    callback: () => {}
  });

  return googleSheetsTokenClient;
}

async function getGoogleSheetsAccessToken() {
  const tokenClient = ensureGoogleSheetsTokenClient();

  return await new Promise((resolve, reject) => {
    tokenClient.callback = response => {
      if (response?.error) {
        reject(new Error(response.error));
        return;
      }

      googleSheetsAccessToken = response?.access_token || "";
      if (!googleSheetsAccessToken) {
        reject(new Error("missing_google_access_token"));
        return;
      }

      resolve(googleSheetsAccessToken);
    };

    tokenClient.requestAccessToken({
      prompt: googleSheetsAccessToken ? "" : "consent"
    });
  });
}

async function fetchSheetValues(range, { valueRenderOption = "FORMATTED_VALUE" } = {}) {
  const token = await getGoogleSheetsAccessToken();
  const spreadsheetId = runtimeConfig.spreadsheetId || "";
  if (!spreadsheetId) {
    throw new Error("Spreadsheet ID is not configured.");
  }

  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}?valueRenderOption=${encodeURIComponent(valueRenderOption)}`,
    {
      headers: {
        Authorization: `Bearer ${token}`
      }
    }
  );

  if (!response.ok) {
    throw new Error(`sheets_read_failed:${response.status}:${await response.text().catch(() => "")}`);
  }

  const data = await response.json();
  return Array.isArray(data.values) ? data.values : [];
}

async function loadPendingRecordsFromSheetsFallback() {
  if (!runtimeConfig.sheetReadFallbackEnabled) {
    throw new Error("Sheet fallback is not enabled.");
  }

  const sheetName = runtimeConfig.sourceSheetName || SHEET_SOURCE_CONFIG.sourceRange;
  const range = `'${sheetName.replace(/'/g, "''")}'!${SHEET_SOURCE_CONFIG.sourceRange}`;
  const values = await fetchSheetValues(range);
  const records = values
    .map((row, index) => buildPendingRecordFromSheetRow(row, index))
    .filter(Boolean);

  return {
    ok: true,
    version: `${runtimeConfig.appVersion || "unknown"}-sheets-fallback`,
    records
  };
}

function buildGraphicsExportHeaders() {
  return [
    "Label",
    "Graphics Request ID",
    "Book",
    "Poem Title",
    "Quote",
    "Author",
    "Record ID",
    "Sheet Row"
  ];
}

function getAuthorLastNameLabel(author) {
  const parts = cleanSheetWhitespace(author).split(" ").filter(Boolean);
  return (parts[parts.length - 1] || "UNKNOWN").toUpperCase();
}

function getBookShortenerForTitle(bookTitle) {
  const normalizedBook = normalizeBookKey(bookTitle);
  const publishingBook = currentIntakePublishingBooksByKey.get(normalizedBook);
  if (publishingBook?.bookShortener) {
    return cleanSheetWhitespace(publishingBook.bookShortener).toUpperCase();
  }
  const catalogBook = currentIntakeCatalogBooksByKey.get(normalizedBook);
  if (catalogBook?.bookShortener) {
    return cleanSheetWhitespace(catalogBook.bookShortener).toUpperCase();
  }
  return "BOOK";
}

function buildGraphicsExportLabel(record) {
  const lastName = getAuthorLastNameLabel(record.author || "");
  const bookShortener = getBookShortenerForTitle(record.bookTitle || "");
  const poemTitle = cleanSheetWhitespace(record.poemTitle || "UNTITLED POEM").toUpperCase();
  return `${lastName} - ${bookShortener} - QI - ${poemTitle}`;
}

function buildGraphicsExportRows(records) {
  return records.map(record => [
    buildGraphicsExportLabel(record),
    record.graphicsRequestId || "",
    record.bookTitle || "",
    record.poemTitle || "",
    record.quoteText || "",
    record.author || "",
    record.recordId || "",
    record.sheetRow || ""
  ]);
}

function buildGraphicsExportTitle(bookTitle) {
  const today = new Date().toISOString().slice(0, 10);
  const normalizedBookTitle = (bookTitle || "Graphics Export")
    .trim()
    .replace(/\s+/g, " ")
    .slice(0, 80);
  return `Weaver Graphics Export - ${normalizedBookTitle} - ${today}`;
}

async function createGoogleSheetFromRows({ title, headers, rows }) {
  const token = await getGoogleSheetsAccessToken();

  const createResponse = await fetch("https://sheets.googleapis.com/v4/spreadsheets", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      properties: { title },
      sheets: [{ properties: { title: "Graphics Export" } }]
    })
  });

  if (!createResponse.ok) {
    throw new Error(`sheet_create_failed:${createResponse.status}:${await createResponse.text().catch(() => "")}`);
  }

  const created = await createResponse.json();
  const spreadsheetId = created.spreadsheetId;
  if (!spreadsheetId) {
    throw new Error("missing_spreadsheet_id");
  }

  const writeResponse = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/A1:append?valueInputOption=RAW`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      majorDimension: "ROWS",
      values: [headers, ...rows]
    })
  });

  if (!writeResponse.ok) {
    throw new Error(`sheet_write_failed:${writeResponse.status}:${await writeResponse.text().catch(() => "")}`);
  }

  return {
    spreadsheetId,
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`
  };
}

function openComparisonWindow(href, windowName) {
  const width = 820;
  const height = Math.max(720, Math.min(960, (window.screen?.availHeight || window.outerHeight || 960) - 80));
  const screenLeft = typeof window.screenX === "number" ? window.screenX : 0;
  const screenTop = typeof window.screenY === "number" ? window.screenY : 0;
  const outerWidth = window.outerWidth || window.innerWidth || width;
  const availWidth = window.screen?.availWidth || screenLeft + outerWidth + width + 40;
  const rightSideLeft = screenLeft + outerWidth + 20;
  const leftSideLeft = screenLeft - width - 20;
  let left = rightSideLeft;

  if (rightSideLeft + width > availWidth - 20) {
    left = leftSideLeft >= 20 ? leftSideLeft : Math.max(20, availWidth - width - 20);
  }

  const top = Math.max(20, screenTop + 40);
  const features = [
    "popup=yes",
    `width=${width}`,
    `height=${height}`,
    `left=${Math.round(left)}`,
    `top=${Math.round(top)}`,
    "toolbar=no",
    "menubar=no",
    "location=no",
    "status=no",
    "scrollbars=yes",
    "resizable=yes"
  ].join(",");
  const popup = window.open("", windowName, features);

  if (popup) {
    try {
      popup.moveTo(Math.round(left), Math.round(top));
      popup.resizeTo(width, height);
    } catch (_error) {
      // Ignore browser restrictions on popup positioning.
    }
    popup.location.replace(href);
    popup.focus();
  } else {
    window.open(href, "_blank", "noopener,noreferrer");
  }
}

function extractGoogleDriveFileId(url) {
  const text = String(url || "").trim();
  if (!text) return "";

  const fileMatch = text.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (fileMatch) return fileMatch[1];

  try {
    const parsed = new URL(text);
    return parsed.searchParams.get("id") || "";
  } catch (_error) {
    return "";
  }
}

function buildGraphicsPreviewUrl(assetMatch) {
  if (!assetMatch) return "";

  const directUrl = assetMatch.lowResLinkUrl || assetMatch.linkUrl || "";
  const driveFileId = extractGoogleDriveFileId(directUrl);
  if (driveFileId) {
    return `/api/drive-image?fileId=${encodeURIComponent(driveFileId)}`;
  }

  return directUrl;
}

function buildReviewSavePayload(update) {
  return {
    sourceRow: update.sourceRow,
    recordId: update.recordId,
    approval: update.reviewDecision,
    reviewDecision: update.reviewDecision,
    correctionNote: update.correctionNote,
    correctedAuthor: update.correctedAuthor,
    correctedTitle: update.correctedTitle,
    correctedBookTitle: update.correctedBookTitle,
    correctedExcerpt: update.correctedExcerpt,
    graphicsQi: update.useForQi ? "1" : "0",
    useForQi: update.useForQi ? "1" : "0",
    photos: update.useForInt ? "1" : "0",
    useForInt: update.useForInt ? "1" : "0",
    duplicateGroupId: update.duplicateGroupId || ""
  };
}

function getEffectiveReviewDecision(excerpt) {
  const explicitDecision = normalizeDecision(excerpt?.excerptReviewDecision);
  if (explicitDecision) {
    return explicitDecision;
  }

  return Boolean(excerpt?.useForQi ?? excerpt?.useForGraphicsQi) ? "accept" : "";
}

function buildBatchReviewSavePayload(update) {
  return {
    sourceRow: update.sourceRow,
    recordId: update.recordId,
    approval: update.reviewDecision,
    reviewDecision: update.reviewDecision,
    correctionNote: update.correctionNote,
    correctedAuthor: update.correctedAuthor,
    correctedTitle: update.correctedTitle,
    correctedBookTitle: update.correctedBookTitle,
    correctedExcerpt: update.correctedExcerpt,
    graphicsQi: Boolean(update.useForQi),
    useForQi: Boolean(update.useForQi),
    photos: Boolean(update.useForInt),
    useForInt: Boolean(update.useForInt),
    duplicateGroupId: update.duplicateGroupId || ""
  };
}

async function requestBatchSave(updates) {
  const response = await fetch("/api/save-reviews", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      updates: updates.map(buildBatchReviewSavePayload)
    })
  });

  let data;
  try {
    data = await response.json();
  } catch (_error) {
    throw new Error("Batch save returned an unreadable response.");
  }

  if (!response.ok || !data.ok) {
    throw new Error(data.error || "Batch save failed.");
  }

  return data;
}

async function saveReviewsSequentially(changedUpdates) {
  for (const update of changedUpdates) {
    const response = await fetch("/api/save-review-single", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        update: buildReviewSavePayload(update)
      })
    });

    let data;
    try {
      data = await response.json();
    } catch (_error) {
      throw new Error(`Save returned unreadable JSON for row ${update.sourceRow}.`);
    }

    if (!response.ok || !data.ok) {
      throw new Error(data.error || `Save failed for row ${update.sourceRow}.`);
    }
  }
}

async function loadBooks(options = {}) {
  const preserveSelection = options.preserveSelection !== false;

  try {
    setStatus("Loading book titles...");
    let data;
    try {
      data = await requestReviewApi("/api/review/pending-records");
    } catch (primaryError) {
      if (!runtimeConfig.sheetReadFallbackEnabled) {
        throw primaryError;
      }
      setStatus(`Primary review backend is unavailable. Loading books directly from Google Sheets...`, {
        error: primaryError.message
      });
      data = await loadPendingRecordsFromSheetsFallback();
    }

    const records = Array.isArray(data.records) ? data.records : [];
    const allBookSummaries = applyPendingBookData(records, {
      preserveSelection,
    });
    setStatus(
      `Loaded ${allBookSummaries.length} books. Backend ${data.version || "unknown"}.`,
      allBookSummaries.slice(0, 10)
    );

    loadCatalogValidation(records)
      .then(() => {
        const refreshedSummaries = applyPendingBookData(records, { preserveSelection: true });
        setStatus(
          `Loaded ${refreshedSummaries.length} books. Backend ${data.version || "unknown"}.`,
          refreshedSummaries.slice(0, 10)
        );
      })
      .catch(() => {
        // Keep the queue usable even if validation refresh fails.
      });
  } catch (error) {
    setStatus(`Book load failed: ${error.message}`);
  }
}

function summarizePendingBooks(records) {
  const byTitle = new Map();

  records.forEach(record => {
    const title = (record.bookTitle || "").trim().replace(/\s+/g, " ");
    if (!title) return;
    const titleKey = normalizeBookKey(title);
    if (!titleKey) return;
    const key = record.recordId || String(record.sourceRow);
    if (record.catalogValidation?.status) {
      currentValidationByRecordId.set(key, record.catalogValidation);
    }
    const validation = currentValidationByRecordId.get(key) || record.catalogValidation || null;

    if (!byTitle.has(titleKey)) {
      byTitle.set(titleKey, {
        key: titleKey,
        title,
        totalCount: 0,
        standardCount: 0,
        goodCount: 0,
        needsCheckingCount: 0,
        pdfOnlyCount: 0,
        variants: new Set([title])
      });
    }

    const summary = byTitle.get(titleKey);
    summary.title = choosePreferredBookTitle(summary.title, title);
    summary.variants.add(title);
    summary.totalCount += 1;
    const isPdfOnly = isPdfOnlyCatalogValidation(validation, record);
    if (isPdfOnly) {
      summary.pdfOnlyCount += 1;
    }
    if (!isPdfOnly) {
      summary.standardCount += 1;
      summary.goodCount += 1;
    }
    if (isExtraReviewRecord(record, validation)) {
      summary.needsCheckingCount += 1;
    }
  });

  return Array.from(byTitle.values())
    .map(summary => ({
      ...summary,
      variants: Array.from(summary.variants).sort((left, right) => left.localeCompare(right))
    }))
    .sort((left, right) => left.title.localeCompare(right.title));
}

function populateBookSelect(select, books, previousSelection, labelBuilder) {
  if (!select) return;
  select.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Choose a book";
  select.appendChild(placeholder);

  books.forEach(book => {
    const option = document.createElement("option");
    option.value = book.key;
    option.textContent = labelBuilder(book);
    select.appendChild(option);
  });

  if (previousSelection) {
    const hasPreviousSelection = books.some(book => book.key === previousSelection);
    if (hasPreviousSelection) {
      select.value = previousSelection;
    }
  }
}

function getSelectedGraphicsMode() {
  return elements.graphicsMode?.value || "queue";
}

function getSelectedGraphicsFilter() {
  return elements.graphicsFilter?.value || "current_titles";
}

function getGraphicsModeLabel(mode = getSelectedGraphicsMode()) {
  if (mode === "cleanup") return "graphics QC";
  if (mode === "mismatch") return "mismatch pairing";
  if (mode === "handoff") return "Poetry Please handoff";
  return "graphics creation";
}

function getVisibleGraphicsBookSummaries() {
  if (getSelectedGraphicsMode() !== "queue" || getSelectedGraphicsFilter() !== "current_titles") {
    return currentGraphicsBookSummaries;
  }

  const reviewQueueIncludeSet = getReviewQueueIncludeSet();
  if (!reviewQueueIncludeSet.size) {
    return currentGraphicsBookSummaries;
  }

  return currentGraphicsBookSummaries.filter(book => reviewQueueIncludeSet.has(normalizeBookKey(book.title)));
}

function refreshGraphicsBookSelect(preserveSelection = true) {
  const previousSelection = preserveSelection ? elements.graphicsBookSelect?.value || "" : "";
  const visibleBooks = getVisibleGraphicsBookSummaries();

  populateBookSelect(
    elements.graphicsBookSelect,
    visibleBooks.map(book => ({ ...book, key: normalizeBookKey(book.title) })),
    previousSelection,
    book => `${book.title} (${book.count})`
  );

  if (elements.graphicsBookCountBadge) {
    elements.graphicsBookCountBadge.textContent = `${visibleBooks.length} Books`;
  }
}

function setGraphicsFolderImportApplyEnabled(enabled) {
  if (elements.applyGraphicsFolderImport) {
    elements.applyGraphicsFolderImport.disabled = !enabled;
  }
}

function hasGraphicsFolderManualSelections() {
  return Object.keys(collectGraphicsFolderManualSelections()).length > 0;
}

function renderGraphicsFolderImportPreview(preview) {
  if (!elements.graphicsFolderImportResults) {
    return;
  }

  currentGraphicsFolderImportPreview = preview || null;
  setGraphicsFolderImportApplyEnabled(Boolean(preview?.matches?.length));
  elements.graphicsFolderImportResults.innerHTML = "";

  if (!preview) {
    const empty = document.createElement("p");
    empty.className = "empty-state";
    empty.textContent = "Preview a Drive folder to see which open graphics requests can be matched safely.";
    elements.graphicsFolderImportResults.appendChild(empty);
    return;
  }

  const summary = document.createElement("div");
  summary.className = "folder-import-summary";
  const sourceLabel = preview.sourceType === "file" ? "File" : "Folder";
  summary.innerHTML = `
    <span class="badge badge--muted">${sourceLabel} ${escapeHtml(preview.folderName || preview.fileId || preview.folderId || "unknown")}</span>
    <span class="badge badge--muted">${Number(preview.imageCount || 0)} image file${Number(preview.imageCount || 0) === 1 ? "" : "s"}</span>
    <span class="badge badge--signal">${Number(preview.matches?.length || 0)} safe match${Number(preview.matches?.length || 0) === 1 ? "" : "es"}</span>
    <span class="badge badge--warn">${Number(preview.unmatched?.length || 0)} unmatched</span>
  `;
  elements.graphicsFolderImportResults.appendChild(summary);

  const matchedSection = document.createElement("section");
  matchedSection.className = "folder-import-section";
  matchedSection.innerHTML = `<h4>Safe Matches</h4>`;
  const matchedList = document.createElement("div");
  matchedList.className = "folder-import-list";
  const matches = Array.isArray(preview.matches) ? preview.matches : [];
    if (!matches.length) {
      const empty = document.createElement("p");
      empty.className = "empty-state";
      empty.textContent = "No safe matches found yet. This usually means the filename does not line up cleanly with the open queue.";
      matchedList.appendChild(empty);
    } else {
    matches.forEach(match => {
      const card = document.createElement("article");
      card.className = "folder-import-card";
      card.innerHTML = `
        <h5 class="folder-import-card__title">${escapeHtml(match.fileName || "Unnamed file")}</h5>
        ${match.assetPreviewUrl ? `
          <div class="folder-import-card__preview">
            <img class="folder-import-card__image" src="${escapeAttribute(match.assetPreviewUrl)}" alt="Preview of ${escapeAttribute(match.fileName || "matched graphic")}" loading="lazy" />
          </div>
        ` : ""}
        <p class="folder-import-card__meta"><strong>${escapeHtml(match.poemTitle || "Untitled poem")}</strong> · ${escapeHtml(match.author || "Unknown author")}</p>
        <p class="folder-import-card__meta">${escapeHtml(match.bookTitle || "")}</p>
        ${match.quoteText ? `<blockquote class="folder-import-card__quote">${escapeHtml(match.quoteText)}</blockquote>` : ""}
        <p class="folder-import-card__meta">Queue row ${escapeHtml(String(match.queueSheetRow || "—"))} · Request ${escapeHtml(match.graphicsRequestId || "—")}</p>
      `;
      matchedList.appendChild(card);
    });
  }
  matchedSection.appendChild(matchedList);
  elements.graphicsFolderImportResults.appendChild(matchedSection);

  const unmatched = Array.isArray(preview.unmatched) ? preview.unmatched : [];
  const unmatchedRequests = Array.isArray(preview.unmatchedRequests) ? preview.unmatchedRequests : [];
  const unmatchedRequestOptions = unmatchedRequests.map(item => `
    <option value="${escapeAttribute(item.graphicsRequestId || "")}">
      ${escapeHtml(`${item.poemTitle || "Untitled poem"} · ${item.author || "Unknown author"} · ${item.bookTitle || ""}`)}
    </option>
  `).join("");
  if (unmatched.length) {
    const unmatchedSection = document.createElement("details");
    unmatchedSection.className = "folder-import-section";
    unmatchedSection.innerHTML = `<summary>Unmatched Files (${unmatched.length})</summary>`;
    const unmatchedList = document.createElement("div");
    unmatchedList.className = "folder-import-list";
    unmatched.slice(0, 50).forEach(item => {
      const card = document.createElement("article");
      card.className = "folder-import-card";
      card.dataset.fileId = item.fileId || "";
      card.innerHTML = `
        <h5 class="folder-import-card__title">${escapeHtml(item.fileName || "Unnamed file")}</h5>
        ${item.assetPreviewUrl ? `
          <div class="folder-import-card__preview">
            <img class="folder-import-card__image" src="${escapeAttribute(item.assetPreviewUrl)}" alt="Preview of ${escapeAttribute(item.fileName || "unmatched graphic")}" loading="lazy" />
          </div>
        ` : ""}
        <p class="folder-import-card__meta">${escapeHtml(item.reason || "No safe open-queue match was found.")}</p>
      `;
      if (Array.isArray(item.candidates) && item.candidates.length) {
        const candidates = document.createElement("div");
        candidates.className = "folder-import-candidates";
        candidates.innerHTML = `
          <p class="hint">Choose the correct match if you want to import this one manually:</p>
          ${item.candidates.map(candidate => `
            <div class="folder-import-candidate">
              <label>
                <span><input type="radio" name="folder-import-match-${escapeAttribute(item.fileId || "unknown")}" value="${escapeAttribute(candidate.graphicsRequestId || "")}" /> <strong>${escapeHtml(candidate.poemTitle || "Untitled poem")}</strong> · ${escapeHtml(candidate.author || "Unknown author")}</span>
                <span>${escapeHtml(candidate.bookTitle || "")}</span>
                ${candidate.quoteText ? `<blockquote class="folder-import-card__quote folder-import-card__quote--candidate">${escapeHtml(candidate.quoteText)}</blockquote>` : ""}
                <span>Request ${escapeHtml(candidate.graphicsRequestId || "—")} · score ${escapeHtml(String(candidate.score || ""))} · ${escapeHtml((candidate.reasons || []).join(", "))}</span>
              </label>
            </div>
          `).join("")}
        `;
        card.appendChild(candidates);
      }
      if (unmatchedRequests.length) {
        const picker = document.createElement("div");
        picker.className = "folder-import-candidates";
        picker.innerHTML = `
          <p class="hint">Or pair this file with any unmatched excerpt from this preview:</p>
          <label class="folder-import-manual-pairing">
            <select class="folder-import-manual-select" data-file-id="${escapeAttribute(item.fileId || "")}">
              <option value="">Choose unmatched excerpt</option>
              ${unmatchedRequestOptions}
            </select>
          </label>
        `;
        card.appendChild(picker);
      }
      unmatchedList.appendChild(card);
    });
    unmatchedSection.appendChild(unmatchedList);
    elements.graphicsFolderImportResults.appendChild(unmatchedSection);
  }

  if (unmatchedRequests.length) {
    const requestsSection = document.createElement("details");
    requestsSection.className = "folder-import-section";
    requestsSection.innerHTML = `<summary>Unmatched Excerpts (${unmatchedRequests.length})</summary>`;
    const requestsList = document.createElement("div");
    requestsList.className = "folder-import-list";
    unmatchedRequests.slice(0, 75).forEach(item => {
      const card = document.createElement("article");
      card.className = "folder-import-card";
      card.innerHTML = `
        <h5 class="folder-import-card__title">${escapeHtml(item.poemTitle || "Untitled poem")}</h5>
        <p class="folder-import-card__meta"><strong>${escapeHtml(item.author || "Unknown author")}</strong> · ${escapeHtml(item.bookTitle || "")}</p>
        ${item.quoteText ? `<blockquote class="folder-import-card__quote">${escapeHtml(item.quoteText)}</blockquote>` : ""}
        <p class="folder-import-card__meta">Queue row ${escapeHtml(String(item.queueSheetRow || "—"))} · Request ${escapeHtml(item.graphicsRequestId || "—")}</p>
      `;
      requestsList.appendChild(card);
    });
    requestsSection.appendChild(requestsList);
    elements.graphicsFolderImportResults.appendChild(requestsSection);
  }

  setGraphicsFolderImportApplyEnabled(Boolean(matches.length || hasGraphicsFolderManualSelections()));
}

function collectGraphicsFolderManualSelections() {
  if (!elements.graphicsFolderImportResults) {
    return {};
  }

  const selections = {};
  elements.graphicsFolderImportResults.querySelectorAll(".folder-import-card[data-file-id]").forEach(card => {
    const fileId = card.dataset.fileId || "";
    if (!fileId) return;
    const selected = card.querySelector('input[type="radio"]:checked');
    if (selected?.value) {
      selections[fileId] = selected.value;
      return;
    }
    const manualSelect = card.querySelector(".folder-import-manual-select");
    if (manualSelect?.value) {
      selections[fileId] = manualSelect.value;
    }
  });
  return selections;
}

function refreshGraphicsFolderImportApplyState() {
  setGraphicsFolderImportApplyEnabled(Boolean(currentGraphicsFolderImportPreview?.matches?.length || hasGraphicsFolderManualSelections()));
}

function refreshGraphicsFolderImportVisibility() {
  const isQueueMode = getSelectedGraphicsMode() === "queue";
  if (elements.graphicsFolderImportPanel) {
    elements.graphicsFolderImportPanel.hidden = !isQueueMode;
  }
}

async function previewGraphicsFolderImport() {
  const folderUrl = elements.graphicsFolderUrl?.value?.trim() || "";
  if (!folderUrl) {
    setStatus("Paste a Google Drive folder or file link first.");
    return;
  }

  try {
    if (elements.previewGraphicsFolderImport) {
      elements.previewGraphicsFolderImport.disabled = true;
      elements.previewGraphicsFolderImport.textContent = "Previewing...";
    }
    setStatus("Scanning the Drive link and previewing safe matches...");
    const response = await fetch("/api/graphics/folder-import/preview", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ folderUrl })
    });
    const result = await response.json();
    if (!response.ok || !result.ok) {
      throw new Error(result.error || `/api/graphics/folder-import/preview returned ${response.status}`);
    }
    renderGraphicsFolderImportPreview(result);
    setStatus(`Previewed ${result.imageCount || 0} image file${Number(result.imageCount || 0) === 1 ? "" : "s"}. Found ${result.matches?.length || 0} safe matches and ${result.unmatchedRequests?.length || 0} unmatched excerpts.`);
  } catch (error) {
    renderGraphicsFolderImportPreview(null);
    setStatus(`Folder import preview failed: ${error.message}`);
  } finally {
    if (elements.previewGraphicsFolderImport) {
      elements.previewGraphicsFolderImport.disabled = false;
      elements.previewGraphicsFolderImport.textContent = "Preview Matches";
    }
  }
}

async function applyGraphicsFolderImport() {
  const folderUrl = elements.graphicsFolderUrl?.value?.trim() || "";
  if (!folderUrl) {
    setStatus("Paste a Google Drive folder or file link first.");
    return;
  }

  if (!currentGraphicsFolderImportPreview?.matches?.length && !hasGraphicsFolderManualSelections()) {
    setStatus("Preview the folder first, then select at least one safe or manual match to import.");
    return;
  }

  try {
    if (elements.applyGraphicsFolderImport) {
      elements.applyGraphicsFolderImport.disabled = true;
      elements.applyGraphicsFolderImport.textContent = "Importing...";
    }
    const manualSelections = collectGraphicsFolderManualSelections();
    const importCount = Number(currentGraphicsFolderImportPreview.matches?.length || 0) + Object.keys(manualSelections).length;
    setStatus(`Importing up to ${importCount} matched graphics into QC...`);
    const response = await fetch("/api/graphics/folder-import/apply", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        folderUrl,
        manualSelections,
        options: {
          expectedSafeMatchCount: Number(currentGraphicsFolderImportPreview.matches?.length || 0)
        }
      })
    });
    const result = await response.json();
    if (!response.ok || !result.ok) {
      throw new Error(result.error || `/api/graphics/folder-import/apply returned ${response.status}`);
    }
    renderGraphicsFolderImportPreview(null);
    await loadGraphicsBooks();
    if (elements.graphicsBookSelect?.value) {
      await loadGraphicsRecords();
    }
    setStatus(`Attempted ${result.attemptedCount || 0}; imported ${result.importedCount || 0}; skipped ${result.skippedCount || 0}.`);
  } catch (error) {
    setGraphicsFolderImportApplyEnabled(Boolean(currentGraphicsFolderImportPreview?.matches?.length || hasGraphicsFolderManualSelections()));
    setStatus(`Drive folder import failed: ${error.message}`);
  } finally {
    if (elements.applyGraphicsFolderImport) {
      elements.applyGraphicsFolderImport.textContent = "Import Matched Graphics";
      elements.applyGraphicsFolderImport.disabled = !(currentGraphicsFolderImportPreview?.matches?.length || hasGraphicsFolderManualSelections());
    }
  }
}

async function loadGraphicsBooks(options = {}) {
  const preserveSelection = options.preserveSelection !== false;
  const mode = getSelectedGraphicsMode();

  try {
    setStatus(`Loading ${getGraphicsModeLabel(mode)} books...`);
    let data;
    try {
      data = await requestReviewApi("/api/review/graphics-books", { mode });
    } catch (primaryError) {
      if (!runtimeConfig.sheetReadFallbackEnabled) {
        throw primaryError;
      }
      setStatus(`Primary graphics backend is unavailable. Loading ${getGraphicsModeLabel(mode)} books directly from Google Sheets...`, {
        error: primaryError.message
      });
      data = await loadGraphicsBooksFromSheetsFallback(mode);
    }

    currentGraphicsBookSummaries = Array.isArray(data.books) ? data.books : [];
    graphicsBookSummaryByKey = new Map(
      currentGraphicsBookSummaries.map(book => [normalizeBookKey(book.title), book])
    );
    refreshGraphicsBookSelect(preserveSelection);

    setStatus(
      `Loaded ${getVisibleGraphicsBookSummaries().length} ${getGraphicsModeLabel(mode)} books. Backend ${data.version || "unknown"}.`,
      getVisibleGraphicsBookSummaries().slice(0, 10)
    );
  } catch (error) {
    setStatus(`Graphics book load failed: ${error.message}`);
  }
}

async function loadGraphicsRecords() {
  const mode = getSelectedGraphicsMode();
  const bookKey = elements.graphicsBookSelect?.value || "";
  if (!bookKey) {
    setStatus("Choose a graphics book title first.");
    return;
  }

  const summary = graphicsBookSummaryByKey.get(bookKey);
  const bookTitle = summary?.title || bookKey;

  try {
    setStatus(`Loading ${getGraphicsModeLabel(mode)} rows for "${bookTitle}"...`);
    let data;
    try {
      data = await requestReviewApi("/api/review/graphics-records", { mode, bookTitle });
    } catch (primaryError) {
      if (!runtimeConfig.sheetReadFallbackEnabled) {
        throw primaryError;
      }
      setStatus(`Primary graphics backend is unavailable. Loading ${getGraphicsModeLabel(mode)} rows directly from Google Sheets...`, {
        error: primaryError.message
      });
      data = await loadGraphicsRecordsFromSheetsFallback(bookTitle, mode);
    }

    currentGraphicsRecords = Array.isArray(data.records) ? data.records : [];
    currentGraphicsAssetMatches = await loadGraphicsAssetMatches(currentGraphicsRecords);
    renderGraphicsRecords(currentGraphicsRecords);
    setStatus(
      `Loaded ${currentGraphicsRecords.length} ${getGraphicsModeLabel(mode)} rows for "${bookTitle}".`,
      currentGraphicsRecords.slice(0, 5)
    );
  } catch (error) {
    setStatus(`Graphics record load failed: ${error.message}`);
  }
}

async function loadGraphicsAssetMatches(records) {
  if (!records.length) {
    return new Map();
  }

  try {
    const response = await fetch("/api/graphics/links", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ records })
    });
    if (!response.ok) {
      throw new Error(`/api/graphics/links returned ${response.status}`);
    }
    const data = await response.json();
    const matches = data && typeof data.matches === "object" && data.matches ? data.matches : {};
    return new Map(Object.entries(matches).filter(([, value]) => value));
  } catch (error) {
    setStatus(`Graphics asset link load failed: ${error.message}`);
    return new Map();
  }
}

function isGoodValidation(validation) {
  return Boolean(validation && validation.status === "catalog_match");
}

async function loadExcerpts() {
  const bookKey = elements.bookSelect.value;
  if (!bookKey) {
    setStatus("Choose a book title first.");
    return;
  }

  try {
    const summary = reviewBookSummaryByKey.get(bookKey);
    const bookTitle = summary?.title || bookKey;
    setStatus(`Loading excerpts for "${bookTitle}"...`);
    currentExcerpts = getPendingRecordsForBookKey(bookKey);
    reviewVisibleCount = getReviewBatchSize();
    reviewPinnedRowOrder = [];
    await loadCatalogValidation(currentExcerpts);
    renderCurrentExcerpts();
    setStatus(`Loaded ${currentExcerpts.length} excerpts for "${bookTitle}".`);
  } catch (error) {
    setStatus(`Excerpt load failed: ${error.message}`);
  }
}

async function loadWeirdExcerpts() {
  const bookKey = elements.weirdBookSelect?.value;
  if (!bookKey) {
    setStatus("Choose a book title first.");
    return;
  }

  try {
    const summary = weirdBookSummaryByKey.get(bookKey) || reviewBookSummaryByKey.get(bookKey);
    const bookTitle = summary?.title || bookKey;
    setStatus(`Loading extra-review excerpts for "${bookTitle}"...`);
    currentWeirdExcerpts = getPendingRecordsForBookKey(bookKey);
    weirdVisibleCount = EXTRA_REVIEW_BATCH_SIZE;
    weirdPinnedRowOrder = [];
    await loadCatalogValidation(currentWeirdExcerpts);
    renderWeirdCurrentExcerpts();
    setStatus(`Loaded ${applyExtraReviewFilter(currentWeirdExcerpts).length} extra-review excerpts for "${bookTitle}".`);
  } catch (error) {
    setStatus(`Extra review load failed: ${error.message}`);
  }
}

async function loadCorrectionBooks() {
  const previousSelection = elements.correctionBookSelect?.value || "";

  try {
    setStatus("Loading correction books...");
    let data;
    try {
      data = await requestReviewApi("/api/review/correction-books");
    } catch (primaryError) {
      if (!runtimeConfig.sheetReadFallbackEnabled) {
        throw primaryError;
      }
      setStatus("Primary corrections backend is unavailable. Loading correction books directly from Google Sheets...", {
        error: primaryError.message
      });
      data = await loadCorrectionBooksFromSheetsFallback();
    }

    elements.correctionBookSelect.innerHTML = "";
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Choose a correction book";
    elements.correctionBookSelect.appendChild(placeholder);

    data.books.forEach(book => {
      const option = document.createElement("option");
      option.value = book.title;
      option.textContent = `${book.title} (${book.count})`;
      elements.correctionBookSelect.appendChild(option);
    });

    if (previousSelection) {
      const hasPreviousSelection = data.books.some(book => book.title === previousSelection);
      if (hasPreviousSelection) {
        elements.correctionBookSelect.value = previousSelection;
      }
    }

    elements.correctionBookCountBadge.textContent = `${data.books.length} Books`;
    setStatus(`Loaded ${data.books.length} correction books. Backend ${data.version || "unknown"}.`, data.books.slice(0, 10));
  } catch (error) {
    setStatus(`Correction book load failed: ${error.message}`);
  }
}

async function loadCorrections() {
  const bookTitle = elements.correctionBookSelect?.value;
  if (!bookTitle) {
    setStatus("Choose a correction book title first.");
    return;
  }

  try {
    setStatus(`Loading corrections for "${bookTitle}"...`);
    let data;
    try {
      data = await requestReviewApi("/api/review/corrections", { bookTitle });
    } catch (primaryError) {
      if (!runtimeConfig.sheetReadFallbackEnabled) {
        throw primaryError;
      }
      setStatus(`Primary corrections backend is unavailable. Loading corrections directly from Google Sheets...`, {
        error: primaryError.message
      });
      data = await loadCorrectionsFromSheetsFallback(bookTitle);
    }

    currentCorrectionExcerpts = data.excerpts;
    await loadCatalogValidation(data.excerpts);
    renderCorrectionExcerpts(data.excerpts);
    setStatus(`Loaded ${data.excerpts.length} correction records for "${bookTitle}". Backend ${data.version || "unknown"}.`);
  } catch (error) {
    setStatus(`Correction load failed: ${error.message}`);
  }
}

async function loadCatalogValidation(excerpts) {
  if (!excerpts.length) return;

  const recordsNeedingLiveValidation = [];

  excerpts.forEach(excerpt => {
    const key = excerpt.recordId || String(excerpt.sourceRow);
    if (excerpt.catalogValidation && excerpt.catalogValidation.status) {
      currentValidationByRecordId.set(key, excerpt.catalogValidation);
    } else if (currentValidationByRecordId.has(key)) {
      return;
    } else {
      recordsNeedingLiveValidation.push({
        sourceRow: excerpt.sourceRow,
        recordId: excerpt.recordId,
        author: excerpt.author,
        title: excerpt.title,
        bookTitle: excerpt.bookTitle,
        excerptText: excerpt.excerptText
      });
    }
  });

  if (!recordsNeedingLiveValidation.length) return;

  try {
    const response = await fetch("/api/catalog/validate", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        records: recordsNeedingLiveValidation
      })
    });
    const data = await response.json();
    if (!data.ok) return;
    data.results.forEach(result => {
      currentValidationByRecordId.set(result.recordId || String(result.sourceRow), result);
    });
  } catch (_error) {
    // Keep the queue usable even if local validation is unavailable.
  }
}

function renderExcerpts(excerpts) {
  const totalMatching = excerpts.length;
  const visibleExcerpts = orderReviewExcerptsForRender(excerpts).slice(0, reviewVisibleCount);
  renderExcerptCollection(visibleExcerpts, elements.excerptList, elements.excerptCountBadge, getEmptyStateMessage(), {
    totalMatching,
    visibleCount: visibleExcerpts.length,
    batchSize: getReviewBatchSize(),
    canShowMore: totalMatching > visibleExcerpts.length,
    onShowMore: () => {
      reviewVisibleCount += getReviewBatchSize();
      renderCurrentExcerpts();
    }
  });
}

function renderCorrectionExcerpts(excerpts) {
  renderExcerptCollection(
    excerpts,
    elements.correctionList,
    elements.correctionExcerptCountBadge,
    "No correction records remain for this book."
  );
}

function renderGraphicsRecords(records) {
  refreshGraphicsFolderImportVisibility();
  if (elements.graphicsCountBadge) {
    elements.graphicsCountBadge.textContent = `${records.length} Rows`;
  }
  if (elements.submitGraphicsQc) {
    elements.submitGraphicsQc.hidden = getSelectedGraphicsMode() !== "cleanup";
  }
  if (elements.exportGraphicsSheet) {
    elements.exportGraphicsSheet.hidden = !runtimeConfig.graphicsExportEnabled || getSelectedGraphicsMode() !== "queue";
  }

  if (!elements.graphicsList) {
    return;
  }

  elements.graphicsList.innerHTML = "";

  if (!records.length) {
    const empty = document.createElement("p");
    empty.className = "empty-state";
    empty.textContent = getSelectedGraphicsMode() === "cleanup"
      ? "No QC rows remain for this book."
      : getSelectedGraphicsMode() === "mismatch"
        ? "No mismatched graphics remain for this book."
      : "No graphics creation rows remain for this book.";
    elements.graphicsList.appendChild(empty);
    return;
  }

  records.forEach(record => {
    elements.graphicsList.appendChild(buildGraphicsCard(record));
  });
}

function renderExcerptCollection(excerpts, container, countBadge, emptyMessage, options = {}) {
  const totalMatching = options.totalMatching ?? excerpts.length;
  const visibleCount = options.visibleCount ?? excerpts.length;
  container.innerHTML = "";
  countBadge.textContent = totalMatching > visibleCount
    ? `${visibleCount} of ${totalMatching} Excerpts`
    : `${visibleCount} Excerpts`;

  if (!visibleCount) {
    const empty = document.createElement("p");
    empty.className = "empty-state";
    empty.textContent = emptyMessage;
    container.appendChild(empty);
    return;
  }

  const groups = groupExcerptsByTitle(excerpts);

  groups.forEach(group => {
    const section = document.createElement("section");
    section.className = "poem-group";

    const header = document.createElement("div");
    header.className = "poem-group__header";
    header.innerHTML = `
      <h3>${escapeHtml(group.title)}</h3>
      <span class="badge badge--muted">${group.excerpts.length} excerpt${group.excerpts.length === 1 ? "" : "s"}</span>
    `;
    section.appendChild(header);

    group.excerpts.forEach((excerpt, index) => {
      section.appendChild(buildExcerptCard(excerpt, `${slugify(group.title)}-${index}`));
    });

    container.appendChild(section);
  });

  if (options.canShowMore) {
    const controls = document.createElement("div");
    controls.className = "excerpt-batch-controls";
    const nextLabel = (options.batchSize || REVIEW_SINGLE_BATCH_SIZE) === 1
      ? "Show next excerpt"
      : `Show next ${options.batchSize || REVIEW_SINGLE_BATCH_SIZE}`;
    controls.innerHTML = `
      <p class="hint excerpt-batch-controls__hint">Showing ${visibleCount} of ${totalMatching} matching excerpts.</p>
      <button type="button" class="button button--secondary excerpt-batch-controls__button">
        ${nextLabel}
      </button>
    `;
    controls.querySelector("button")?.addEventListener("click", options.onShowMore);
    container.appendChild(controls);
  }
}

function renderCurrentExcerpts() {
  renderExcerpts(applyReviewFilter(currentExcerpts));
}

function renderWeirdCurrentExcerpts() {
  const filtered = applyExtraReviewFilter(currentWeirdExcerpts);
  const totalMatching = filtered.length;
  const visibleExcerpts = orderWeirdExcerptsForRender(filtered).slice(0, weirdVisibleCount);
  renderExcerptCollection(visibleExcerpts, elements.weirdExcerptList, elements.weirdExcerptCountBadge, getWeirdEmptyStateMessage(), {
    totalMatching,
    visibleCount: visibleExcerpts.length,
    batchSize: EXTRA_REVIEW_BATCH_SIZE,
    canShowMore: totalMatching > visibleExcerpts.length,
    onShowMore: () => {
      weirdVisibleCount += EXTRA_REVIEW_BATCH_SIZE;
      renderWeirdCurrentExcerpts();
    }
  });
}

function orderReviewExcerptsForRender(excerpts) {
  if (!reviewPinnedRowOrder.length) {
    return excerpts;
  }

  const pinnedSet = new Set(reviewPinnedRowOrder.map(Number));
  const excerptByRow = new Map(excerpts.map(excerpt => [Number(excerpt.sourceRow), excerpt]));
  const ordered = [];

  reviewPinnedRowOrder.forEach(sourceRow => {
    const excerpt = excerptByRow.get(Number(sourceRow));
    if (excerpt) {
      ordered.push(excerpt);
    }
  });

  excerpts.forEach(excerpt => {
    if (!pinnedSet.has(Number(excerpt.sourceRow))) {
      ordered.push(excerpt);
    }
  });

  return ordered;
}

function orderWeirdExcerptsForRender(excerpts) {
  if (!weirdPinnedRowOrder.length) {
    return excerpts;
  }

  const pinnedSet = new Set(weirdPinnedRowOrder.map(Number));
  const excerptByRow = new Map(excerpts.map(excerpt => [Number(excerpt.sourceRow), excerpt]));
  const ordered = [];

  weirdPinnedRowOrder.forEach(sourceRow => {
    const excerpt = excerptByRow.get(Number(sourceRow));
    if (excerpt) {
      ordered.push(excerpt);
    }
  });

  excerpts.forEach(excerpt => {
    if (!pinnedSet.has(Number(excerpt.sourceRow))) {
      ordered.push(excerpt);
    }
  });

  return ordered;
}

function getSelectedReviewFilter() {
  return elements.reviewFilter?.value || "all";
}

function applyReviewFilter(excerpts) {
  const mode = getSelectedReviewFilter();
  if (mode === "current_titles") {
    const reviewQueueIncludeSet = getReviewQueueIncludeSet();
    if (!reviewQueueIncludeSet.size) {
      return excerpts;
    }
    return excerpts.filter(excerpt => reviewQueueIncludeSet.has(normalizeBookKey(excerpt.bookTitle)));
  }
  if (mode === "new_only") {
    return excerpts.filter(excerpt => !hasLibraryExcerptMatch(excerpt));
  }
  if (mode === "exact_library") {
    return excerpts.filter(excerpt => getLibraryMatchType(excerpt) === "exact");
  }
  if (mode === "possible_library") {
    return excerpts.filter(excerpt => {
      const matchType = getLibraryMatchType(excerpt);
      return Boolean(matchType && matchType !== "exact");
    });
  }
  if (mode === "likely_correction") {
    return excerpts.filter(excerpt => isLikelyCorrectionExcerpt(excerpt));
  }
  return excerpts;
}

function getSelectedExtraReviewFilter() {
  return elements.weirdReviewFilter?.value || "all_extra";
}

function applyExtraReviewFilter(excerpts) {
  const mode = getSelectedExtraReviewFilter();
  if (mode === "pdf_only") {
    return excerpts.filter(excerpt => isPdfOnlyExcerpt(excerpt));
  }
  if (mode === "missing_catalog") {
    return excerpts.filter(excerpt => !hasCatalogStatus(excerpt));
  }
  if (mode === "likely_correction") {
    return excerpts.filter(excerpt => isLikelyCorrectionExcerpt(excerpt));
  }
  if (mode === "possible_library") {
    return excerpts.filter(excerpt => {
      const match = getLibraryMatch(excerpt);
      return Boolean(match && match.matchType && match.matchType !== "exact");
    });
  }
  if (mode === "linebreak_diff") {
    return excerpts.filter(excerpt => isLineBreakDifferenceExcerpt(excerpt));
  }
  return excerpts.filter(excerpt => isExtraReviewExcerpt(excerpt));
}

function isLikelyCorrectionExcerpt(excerpt) {
  const validation =
    currentValidationByRecordId.get(excerpt.recordId || String(excerpt.sourceRow)) || null;

  return Boolean(validation && isLikelyCorrectionStatus(validation.status));
}

function isPdfOnlyExcerpt(excerpt) {
  const validation =
    currentValidationByRecordId.get(excerpt.recordId || String(excerpt.sourceRow)) || null;
  return isPdfOnlyCatalogValidation(validation, excerpt);
}

function hasLibraryExcerptMatch(excerpt) {
  const validation =
    currentValidationByRecordId.get(excerpt.recordId || String(excerpt.sourceRow)) || null;

  return Boolean(validation && validation.libraryExcerptMatch);
}

function getLibraryMatch(excerpt) {
  const validation =
    currentValidationByRecordId.get(excerpt.recordId || String(excerpt.sourceRow)) || null;

  return validation?.libraryExcerptMatch || null;
}

function getLibraryMatchType(excerpt) {
  return getLibraryMatch(excerpt)?.matchType || "";
}

function hasCatalogStatus(excerpt) {
  const validation =
    currentValidationByRecordId.get(excerpt.recordId || String(excerpt.sourceRow)) || null;
  return Boolean(validation && validation.status);
}

function resolveExcerptDisplayAuthor(excerpt, validation = null) {
  const effectiveValidation =
    validation || currentValidationByRecordId.get(excerpt.recordId || String(excerpt.sourceRow)) || null;
  const libraryMatch = effectiveValidation?.libraryExcerptMatch || null;
  return excerpt.author
    || effectiveValidation?.bookCanonicalAuthor
    || effectiveValidation?.globalExcerptMatch?.author
    || libraryMatch?.author
    || "";
}

function isLikelyCorrectionStatus(status) {
  return [
    "author_mismatch",
    "title_mismatch",
    "poem_title_match_only",
    "excerpt_not_found_in_book",
    "book_not_found",
    "epub_not_present"
  ].includes(status);
}

function isPdfOnlyCatalogValidation(validation, record) {
  const direct = String(record?.bookPrimarySourceFormat || "").toLowerCase();
  if (direct) {
    return direct === "pdf";
  }
  return Boolean(
    validation &&
    String(validation.bookPrimarySourceFormat || "").toLowerCase() === "pdf"
  );
}

function isGoodContentExcerpt(excerpt) {
  const validation =
    currentValidationByRecordId.get(excerpt.recordId || String(excerpt.sourceRow)) || null;

  return Boolean(validation && validation.status === "catalog_match");
}

function isStrictExactLibraryMatch(excerpt) {
  const match = getLibraryMatch(excerpt);
  return Boolean(
    match &&
    match.matchType === "exact" &&
    match.formattingMatch &&
    match.lineBreaksMatch
  );
}

function isExtraReviewExcerpt(excerpt) {
  const validation =
    currentValidationByRecordId.get(excerpt.recordId || String(excerpt.sourceRow)) || null;
  const libraryMatch = getLibraryMatch(excerpt);
  const displayAuthor = resolveExcerptDisplayAuthor(excerpt, validation);

  if (!displayAuthor) {
    return true;
  }

  if (!validation || !validation.status) {
    return true;
  }

  if (isPdfOnlyCatalogValidation(validation, excerpt)) {
    return true;
  }

  if (isLikelyCorrectionStatus(validation.status)) {
    return true;
  }

  if (!libraryMatch) {
    return false;
  }

  if (libraryMatch.matchType !== "exact") {
    return true;
  }

  return !(libraryMatch.formattingMatch && libraryMatch.lineBreaksMatch);
}

function isLineBreakDifferenceExcerpt(excerpt) {
  const match = getLibraryMatch(excerpt);
  return Boolean(match && match.matchType === "exact" && match.lineBreaksMatch === false);
}

function isStandardReviewExcerpt(excerpt) {
  return !isExtraReviewExcerpt(excerpt);
}

function isExtraReviewRecord(record, providedValidation = null) {
  const validation = providedValidation || record.catalogValidation || null;
  const libraryMatch = validation?.libraryExcerptMatch || null;

  if (!validation || !validation.status) {
    return true;
  }

  if (isPdfOnlyCatalogValidation(validation, record)) {
    return true;
  }

  if (isLikelyCorrectionStatus(validation.status)) {
    return true;
  }

  if (!libraryMatch) {
    return false;
  }

  if (libraryMatch.matchType !== "exact") {
    return true;
  }

  return !(libraryMatch.formattingMatch && libraryMatch.lineBreaksMatch);
}

function getEmptyStateMessage() {
  if (getSelectedReviewFilter() === "current_titles") {
    return "No pending excerpts in this book are currently in the 2026 titles lane.";
  }
  if (getSelectedReviewFilter() === "new_only") {
    return "No currently loaded excerpts appear to be net new to the excerpt library.";
  }
  if (getSelectedReviewFilter() === "exact_library") {
    return "No currently loaded excerpts are exact matches to the excerpt library.";
  }
  if (getSelectedReviewFilter() === "possible_library") {
    return "No currently loaded excerpts are possible matches to the excerpt library.";
  }
  if (getSelectedReviewFilter() === "likely_correction") {
    return "No pending excerpts in this book are currently flagged as likely needing correction.";
  }
  return "No pending excerpts remain for this book.";
}

function getWeirdEmptyStateMessage() {
  if (getSelectedExtraReviewFilter() === "pdf_only") {
    return "No pending excerpts in this book are currently coming from PDF-backed catalog books.";
  }
  if (getSelectedExtraReviewFilter() === "missing_catalog") {
    return "No pending excerpts in this book are currently missing catalog status.";
  }
  if (getSelectedExtraReviewFilter() === "likely_correction") {
    return "No pending excerpts in this book are currently flagged as likely needing correction.";
  }
  if (getSelectedExtraReviewFilter() === "possible_library") {
    return "No pending excerpts in this book are currently flagged as possible library matches.";
  }
  if (getSelectedExtraReviewFilter() === "linebreak_diff") {
    return "No pending excerpts in this book are currently flagged as text matches with line-break differences.";
  }
  return "No pending excerpts in this book are currently flagged for extra review.";
}

function groupExcerptsByTitle(excerpts) {
  const groups = new Map();

  excerpts.forEach(excerpt => {
    const title = (excerpt.title || "Untitled poem").trim();
    const key = title.toLowerCase().replace(/\s+/g, " ").trim();
    if (!groups.has(key)) {
      groups.set(key, {
        title,
        excerpts: []
      });
    }
    groups.get(key).excerpts.push(excerpt);
  });

  return Array.from(groups.values());
}

function buildExcerptCard(excerpt, uniqueKey) {
  const card = document.createElement("article");
  card.className = "excerpt-card";
  card.dataset.sourceRow = excerpt.sourceRow;
  card.dataset.recordId = excerpt.recordId || "";

  const validation =
    currentValidationByRecordId.get(excerpt.recordId || String(excerpt.sourceRow)) || null;
  const pullBadge =
    excerpt.exactPullCount > 1
      ? `<span class="badge badge--signal">Pulled ${excerpt.exactPullCount}x</span>`
      : "";
  const overlapBadge =
    excerpt.duplicateGroupId
      ? `<span class="badge badge--signal">Group ${escapeHtml(excerpt.duplicateGroupId)}</span>`
      : "";
  const libraryMatch = validation?.libraryExcerptMatch || null;
  const displayAuthor = resolveExcerptDisplayAuthor(excerpt, validation);
  const libraryBadgeLabel = getLibraryBadgeLabel(libraryMatch);
  const libraryBadge = libraryMatch
    ? `<span class="badge badge--warn">${escapeHtml(libraryBadgeLabel)}</span>`
    : "";
  const libraryRecommendation = libraryMatch?.matchType === "exact"
    ? `<span class="badge badge--muted">Suggested: Reject</span>`
    : "";

  const validationMarkup = buildValidationMarkup(validation, excerpt);
  const reviewDecision = getEffectiveReviewDecision(excerpt);
  const useForQi = Boolean(excerpt.useForQi ?? excerpt.useForGraphicsQi);
  const hasOverrides = Boolean(
    normalizeCorrectionNote(excerpt.correctedAuthor) ||
    normalizeCorrectionNote(excerpt.correctedTitle) ||
    normalizeCorrectionNote(excerpt.correctedBookTitle) ||
    normalizeCorrectionNote(excerpt.correctedExcerpt)
  );

  const decisionBadge = reviewDecision === "accept"
    ? `<span class="badge badge--signal">Accepted excerpt</span>`
    : reviewDecision === "reject"
      ? `<span class="badge badge--muted">Rejected excerpt</span>`
      : reviewDecision === "needs_correction"
        ? `<span class="badge badge--warn">Needs correction</span>`
        : "";
  const qiBadge = useForQi
    ? `<span class="badge badge--signal">Use for QI</span>`
    : "";
  const overrideBadge = hasOverrides
    ? `<span class="badge badge--warn">Source override</span>`
    : "";
  const qcBadge = excerpt.quoteCreatedQc === "Y"
    ? `<span class="badge badge--signal">Made + QCed</span>`
    : "";
  const currentDecision = reviewDecision || "";
  const wordCount = resolveWordCount(excerpt);
  const wordCountBadgeClass = getWordCountBadgeClass(wordCount);
  const wordCountTooltip = '10-25 words is often the sweet spot for readability. Longer excerpts should still be considered and can work as EXC, INT or QI in some cases.';
  card.innerHTML = `
      <div class="excerpt-card__meta">
        <span class="badge badge--muted">Row ${excerpt.sourceRow}</span>
        <span class="badge badge--muted">ID ${escapeHtml(excerpt.recordId || "none")}</span>
        <span class="badge badge--muted">Book ${escapeHtml(excerpt.bookTitle || "(blank)")}</span>
        <span class="badge ${wordCountBadgeClass}" title="${escapeHtml(wordCountTooltip)}">Words ${wordCount}</span>
        ${pullBadge}
        ${overlapBadge}
        ${libraryBadge}
        ${libraryRecommendation}
        ${decisionBadge}
        ${qiBadge}
        ${overrideBadge}
        ${qcBadge}
        <span class="excerpt-card__title">${escapeHtml(excerpt.title || "Untitled poem")}</span>
      <span class="excerpt-card__author">${escapeHtml(displayAuthor || "Unknown author")}</span>
    </div>
    ${validationMarkup}
    <blockquote class="excerpt-card__quote">${escapeHtml(excerpt.excerptText)}</blockquote>
    <p class="hint excerpt-card__hint">Accept now means send to the quote-image queue. Reject and Needs correction do not.</p>
    <div class="decision-group">
      <label><input type="radio" name="approval-${uniqueKey}" value="accept" ${currentDecision === "accept" ? "checked" : ""}> Accept</label>
      <label><input type="radio" name="approval-${uniqueKey}" value="reject" ${currentDecision === "reject" ? "checked" : ""}> Reject</label>
      <label><input type="radio" name="approval-${uniqueKey}" value="needs_correction" ${currentDecision === "needs_correction" ? "checked" : ""}> Needs correction</label>
      <label><input type="radio" name="approval-${uniqueKey}" value="" ${currentDecision === "" ? "checked" : ""}> No decision</label>
    </div>
    <div class="correction-block ${currentDecision === "needs_correction" ? "correction-block--active" : ""}">
      <details class="correction-source">
        <summary>Original extracted values</summary>
        <div class="correction-source__grid">
          <div><strong>Author</strong><span>${escapeHtml(excerpt.rawAuthor || excerpt.author || "")}</span></div>
          <div><strong>Poem title</strong><span>${escapeHtml(excerpt.rawTitle || excerpt.title || "")}</span></div>
          <div><strong>Book</strong><span>${escapeHtml(excerpt.rawBookTitle || excerpt.bookTitle || "")}</span></div>
        </div>
        <pre class="correction-source__excerpt">${escapeHtml(excerpt.rawExcerptText || excerpt.excerptText || "")}</pre>
      </details>
      <label class="field">
        <span>Correction note</span>
        <textarea class="correction-note" rows="3" placeholder="Describe what is wrong and how it should be corrected.">${escapeHtml(excerpt.correctionNote || "")}</textarea>
      </label>
      <div class="correction-grid">
        <label class="field">
          <span>Reassign author</span>
          <input class="corrected-author" type="text" value="${escapeHtml(excerpt.correctedAuthor || excerpt.author || "")}" placeholder="Correct author" />
        </label>
        <label class="field">
          <span>Reassign poem title</span>
          <input class="corrected-title" type="text" value="${escapeHtml(excerpt.correctedTitle || excerpt.title || "")}" placeholder="Correct poem title" />
        </label>
        <label class="field">
          <span>Reassign book</span>
          <input class="corrected-book-title" type="text" value="${escapeHtml(excerpt.correctedBookTitle || excerpt.bookTitle || "")}" placeholder="Correct book title" />
        </label>
      </div>
      <label class="field">
        <span>Correct excerpt text</span>
        <textarea class="corrected-excerpt" rows="5" placeholder="Correct or replace the excerpt text.">${escapeHtml(excerpt.correctedExcerpt || excerpt.excerptText || "")}</textarea>
      </label>
      <p class="hint correction-block__hint">These edits save as durable source overrides so the original extracted values remain preserved for reference.</p>
    </div>
  `;

  card.dataset.currentUseForInt = Boolean(excerpt.useForInt ?? excerpt.useForPhotos) ? "1" : "0";

  card.querySelectorAll(`input[name="approval-${uniqueKey}"]`).forEach(input => {
    input.addEventListener("change", () => {
      const isCorrection = input.value === "needs_correction" && input.checked;
      const block = card.querySelector(".correction-block");
      if (block) {
        block.classList.toggle("correction-block--active", isCorrection);
      }
    });
  });

  if (wordCount === 0) {
    const quoteText = card.querySelector(".excerpt-card__quote")?.textContent || "";
    const resolvedFromRenderedText = countWordsFromText(quoteText);
    if (resolvedFromRenderedText > 0) {
      const badge = card.querySelector(".badge--count");
      if (badge) {
        badge.textContent = `Words ${resolvedFromRenderedText}`;
        badge.className = `badge ${getWordCountBadgeClass(resolvedFromRenderedText)}`;
      }
    }
  }

  card.querySelectorAll(".catalog-poem-link").forEach(link => {
    link.addEventListener("click", event => {
      event.preventDefault();
      const href = link.dataset.popupUrl;
      if (!href) {
        return;
      }
      openComparisonWindow(href, "weaverCatalogPoem");
    });
  });
  card.querySelectorAll(".library-excerpt-link").forEach(link => {
    link.addEventListener("click", event => {
      event.preventDefault();
      const href = link.dataset.popupUrl;
      if (!href) {
        return;
      }
      openComparisonWindow(href, "weaverLibraryExcerpt");
    });
  });

  return card;
}

function buildGraphicsCard(record) {
  const card = document.createElement("article");
  card.className = "excerpt-card";
  card.dataset.recordId = record.recordId || "";
  card.dataset.sheetRow = String(record.sheetRow || "");
  card.dataset.storageTarget = record.storageTarget || "sheet_cleanup";
  card.dataset.graphicsRequestId = record.graphicsRequestId || "";
  card.dataset.pigCompletionId = record.pigCompletionId || "";
  card.dataset.author = record.author || "";
  card.dataset.poemTitle = record.poemTitle || "";
  card.dataset.bookTitle = record.bookTitle || "";
  card.dataset.quoteText = record.quoteText || "";
  card.dataset.currentQcDecision = normalizeGraphicsQcDecisionClient(record.graphicsQcDecision || "");
  card.dataset.currentQcNote = record.graphicsQcNote || "";
  card.dataset.poetryPleaseStatus = record.poetryPleaseStatus || "";
  card.dataset.poetryPleaseUpdatedAt = record.poetryPleaseUpdatedAt || "";
  card.dataset.poetryPleaseNote = record.poetryPleaseNote || "";

  const approved = normalizeApprovalForCompare(record.approved) === "Y";
  const created = normalizeApprovalForCompare(record.created) === "Y";
  const workflowStatus = (record.workflowStatus || "").trim();
  const notes = (record.notes || "").trim();
  const wordCount = countWordsFromText(record.quoteText || "");
  const assetMatch = currentGraphicsAssetMatches.get(record.recordId || String(record.sheetRow || ""));
  const displayAuthor = record.author || assetMatch?.author || "";
  const assetLinkUrl = record.assetLinkUrl || assetMatch?.linkUrl || assetMatch?.lowResLinkUrl || assetMatch?.folderLink || "";
  const assetLinkLabel = assetMatch?.fileName || (assetMatch?.folderLink ? "View Drive folder" : "View graphic on Drive");
  const assetPreviewUrl = record.assetPreviewUrl || buildGraphicsPreviewUrl(assetMatch);
  const isQcMode = getSelectedGraphicsMode() === "cleanup";
  const isMismatchMode = getSelectedGraphicsMode() === "mismatch";
  const isHandoffMode = getSelectedGraphicsMode() === "handoff";
  const currentQcDecision = normalizeGraphicsQcDecisionClient(record.graphicsQcDecision || "");
  const currentQcNote = record.graphicsQcNote || "";
  const qcUpdatedAt = record.graphicsQcUpdatedAt || "";
  const poetryPleaseStatus = record.poetryPleaseStatus || (currentQcDecision === "approve" ? "PENDING" : "");
  const poetryPleaseUpdatedAt = record.poetryPleaseUpdatedAt || "";
  const poetryPleaseNote = record.poetryPleaseNote || "";
  const parsedQcNote = parseGraphicsQcStructuredNote(currentQcNote);
  const qcRejectReason = getGraphicsQcRejectReasonFromLegacyDecision(record.graphicsQcDecision || "", currentQcNote);
  const qcMetadataIssue = parsedQcNote.metadataIssue;
  const qcAestheticIssue = parsedQcNote.aestheticIssue;
  const qcDetailText = parsedQcNote.details || currentQcNote || "";
  const defaultQcNote = GRAPHICS_QC_DEFAULT_NOTES[qcRejectReason] || "";

  card.innerHTML = `
    <div class="excerpt-card__meta">
      <span class="badge badge--muted">Sheet row ${escapeHtml(String(record.sheetRow || ""))}</span>
      <span class="badge badge--muted">Record ${escapeHtml(record.recordId || "none")}</span>
      <span class="badge badge--muted">Book ${escapeHtml(record.bookTitle || "(blank)")}</span>
      <span class="badge ${getWordCountBadgeClass(wordCount)}">Words ${wordCount}</span>
      ${workflowStatus ? `<span class="badge badge--warn">${escapeHtml(workflowStatus)}</span>` : ""}
      ${approved ? '<span class="badge badge--signal">Approved</span>' : '<span class="badge badge--muted">Not approved</span>'}
      ${created ? '<span class="badge badge--signal">Created</span>' : '<span class="badge badge--muted">Not created</span>'}
      ${isHandoffMode && poetryPleaseStatus ? `<span class="badge ${poetryPleaseStatus === "HANDED_OFF" ? "badge--signal" : poetryPleaseStatus === "FAILED" ? "badge--warn" : "badge--muted"}">${escapeHtml(poetryPleaseStatus.replace(/_/g, " "))}</span>` : ""}
      <span class="excerpt-card__title">${escapeHtml(record.poemTitle || "Untitled poem")}</span>
      <span class="excerpt-card__author">${escapeHtml(displayAuthor || "Unknown author")}</span>
    </div>
    <blockquote class="excerpt-card__quote">${escapeHtml(record.quoteText || "")}</blockquote>
    ${(isQcMode || isHandoffMode || isMismatchMode) && assetPreviewUrl ? `
      <figure class="graphics-preview">
        <img class="graphics-preview__image" src="${escapeAttribute(assetPreviewUrl)}" alt="Existing graphic preview for ${escapeAttribute(record.poemTitle || "this excerpt")}" loading="lazy" />
      </figure>
    ` : ""}
    <div class="correction-source__grid">
      <div><strong>Book</strong><span>${escapeHtml(record.bookTitle || "")}</span></div>
      <div><strong>Approved?</strong><span>${escapeHtml(record.approved || "")}</span></div>
      <div><strong>Created?</strong><span>${escapeHtml(record.created || "")}</span></div>
      <div><strong>Workflow</strong><span>${escapeHtml(workflowStatus || "—")}</span></div>
      <div><strong>Record ID</strong><span>${escapeHtml(record.recordId || "—")}</span></div>
      <div><strong>Request ID</strong><span>${escapeHtml(record.graphicsRequestId || "—")}</span></div>
      <div><strong>Sheet Row</strong><span>${escapeHtml(String(record.sheetRow || "—"))}</span></div>
      ${isHandoffMode ? `<div><strong>Handoff</strong><span>${escapeHtml(poetryPleaseStatus || "Pending")}</span></div>` : ""}
    </div>
    ${assetLinkUrl ? `<p class="hint"><a href="${escapeAttribute(assetLinkUrl)}" target="_blank" rel="noopener noreferrer">${escapeHtml(assetLinkLabel)}</a>${assetMatch?.matchType === "cleanup_override" ? " · matched from cleanup override" : ""}${card.dataset.storageTarget === "pig_sheet" ? " · returned from P.I.G." : ""}</p>` : ""}
    ${isMismatchMode ? `<p class="hint">Mismatch reason: ${escapeHtml(record.rejectReason || "mismatched_graphic")}${record.graphicsQcUpdatedAt ? ` · ${escapeHtml(record.graphicsQcUpdatedAt)}` : ""}</p>` : ""}
    ${isHandoffMode && (poetryPleaseUpdatedAt || poetryPleaseNote) ? `<p class="hint">Poetry Please: ${escapeHtml(poetryPleaseStatus || "Pending")}${poetryPleaseUpdatedAt ? ` · ${escapeHtml(poetryPleaseUpdatedAt)}` : ""}${poetryPleaseNote ? ` · ${escapeHtml(poetryPleaseNote)}` : ""}</p>` : ""}
    ${isQcMode ? `
      <fieldset class="graphics-qc-controls">
        <legend>QC decision</legend>
        <label><input type="radio" name="graphics-qc-${escapeAttribute(record.recordId || String(record.sheetRow || ""))}" value="approve" ${currentQcDecision === "approve" ? "checked" : ""}> Approve</label>
        <label><input type="radio" name="graphics-qc-${escapeAttribute(record.recordId || String(record.sheetRow || ""))}" value="reject" ${currentQcDecision === "reject" ? "checked" : ""}> Reject</label>
      </fieldset>
      <div class="graphics-qc-details" ${currentQcDecision === "reject" ? "" : "hidden"}>
        <label class="field">
          <span>Reject reason</span>
          <select class="graphics-qc-reject-reason">
            ${renderGraphicsQcOptions(GRAPHICS_QC_REJECT_REASON_OPTIONS, qcRejectReason)}
          </select>
        </label>
        <label class="field">
          <span>Metadata issue</span>
          <select class="graphics-qc-metadata" ${qcRejectReason === "correct_and_recreate" ? "" : "disabled"}>
            ${renderGraphicsQcOptions(GRAPHICS_QC_METADATA_OPTIONS, qcMetadataIssue)}
          </select>
        </label>
        <label class="field">
          <span>Aesthetic concern</span>
          <select class="graphics-qc-aesthetic" ${qcRejectReason === "correct_and_recreate" ? "" : "disabled"}>
            ${renderGraphicsQcOptions(GRAPHICS_QC_AESTHETIC_OPTIONS, qcAestheticIssue)}
          </select>
        </label>
      </div>
      <label class="field">
        <span>QC note</span>
        <textarea class="graphics-qc-note" rows="3" placeholder="Required for mismatched graphics. Required for correct-and-recreate metadata fixes and aesthetic 'other'." data-default-note="${escapeAttribute(defaultQcNote)}">${escapeHtml(qcDetailText || defaultQcNote)}</textarea>
      </label>
      ${qcUpdatedAt ? `<p class="hint">Last QC update: ${escapeHtml(qcUpdatedAt)}</p>` : ""}
    ` : ""}
    ${notes ? `<p class="hint">${escapeHtml(notes)}</p>` : ""}
  `;

  if (isQcMode) {
    const textarea = card.querySelector(".graphics-qc-note");
    const detailsBlock = card.querySelector(".graphics-qc-details");
    const rejectReasonSelect = card.querySelector(".graphics-qc-reject-reason");
    const metadataSelect = card.querySelector(".graphics-qc-metadata");
    const aestheticSelect = card.querySelector(".graphics-qc-aesthetic");
    const radios = Array.from(card.querySelectorAll(`input[name="graphics-qc-${CSS.escape(record.recordId || String(record.sheetRow || ""))}"]`));
    const syncQcFields = decision => {
      if (detailsBlock) {
        detailsBlock.hidden = decision !== "reject";
      }
      const rejectReason = rejectReasonSelect?.value || "";
      const allowReworkFields = decision === "reject" && rejectReason === "correct_and_recreate";
      if (metadataSelect) {
        metadataSelect.disabled = !allowReworkFields;
        if (!allowReworkFields) {
          metadataSelect.value = "";
        }
      }
      if (aestheticSelect) {
        aestheticSelect.disabled = !allowReworkFields;
        if (!allowReworkFields) {
          aestheticSelect.value = "";
        }
      }
    };
    const syncDefaultNote = () => {
      if (!textarea) return;
      const currentDecision = card.querySelector(`input[name="graphics-qc-${CSS.escape(record.recordId || String(record.sheetRow || ""))}"]:checked`)?.value || "";
      const currentRejectReason = rejectReasonSelect?.value || "";
      const nextDefault = currentDecision === "reject"
        ? (GRAPHICS_QC_DEFAULT_NOTES[currentRejectReason] || "")
        : "";
      const previousDefault = textarea.dataset.defaultNote || "";
      const currentValue = textarea.value.trim();
      if (!currentValue || currentValue === previousDefault) {
        textarea.value = nextDefault;
      }
      textarea.dataset.defaultNote = nextDefault;
    };
    radios.forEach(radio => {
      radio.addEventListener("change", () => {
        if (!textarea || !radio.checked) return;
        syncQcFields(radio.value);
        syncDefaultNote();
      });
    });
    rejectReasonSelect?.addEventListener("change", () => {
      syncQcFields(card.querySelector(`input[name="graphics-qc-${CSS.escape(record.recordId || String(record.sheetRow || ""))}"]:checked`)?.value || "");
      syncDefaultNote();
    });
    syncQcFields(currentQcDecision);
    syncDefaultNote();
  }

  return card;
}

function collectGraphicsQcUpdates() {
  if (!elements.graphicsList) return [];

  return Array.from(elements.graphicsList.querySelectorAll(".excerpt-card")).map(card => {
    const recordId = card.dataset.recordId || "";
    const qcDecision = card.querySelector('input[type="radio"]:checked')?.value || "";
    const rejectReason = card.querySelector(".graphics-qc-reject-reason")?.value || "";
    const metadataIssue = card.querySelector(".graphics-qc-metadata")?.value || "";
    const aestheticIssue = card.querySelector(".graphics-qc-aesthetic")?.value || "";
    const qcDetailText = card.querySelector(".graphics-qc-note")?.value?.trim() || "";
    const qcNote = buildGraphicsQcNotePayload(qcDecision, rejectReason, metadataIssue, aestheticIssue, qcDetailText);
    return {
      recordId,
      sheetRow: card.dataset.sheetRow || "",
      storageTarget: card.dataset.storageTarget || "",
      graphicsRequestId: card.dataset.graphicsRequestId || "",
      pigCompletionId: card.dataset.pigCompletionId || "",
      author: card.dataset.author || "",
      poemTitle: card.dataset.poemTitle || "",
      bookTitle: card.dataset.bookTitle || "",
      quoteText: card.dataset.quoteText || "",
      rejectReason,
      metadataIssue,
      aestheticIssue,
      qcDecision,
      qcNote
    };
  });
}

function filterChangedGraphicsQcUpdates(updates) {
  return updates.filter(update => {
    const card = elements.graphicsList?.querySelector(`.excerpt-card[data-record-id="${CSS.escape(update.recordId)}"]`);
    if (!card) return false;
    return (
      (update.qcDecision || "") !== (card.dataset.currentQcDecision || "") ||
      (update.qcNote || "") !== (card.dataset.currentQcNote || "")
    );
  });
}

async function submitGraphicsQc() {
  if (getSelectedGraphicsMode() !== "cleanup") {
    setStatus("Graphics QC decisions only apply in Graphics QC mode.");
    return;
  }

  const updates = filterChangedGraphicsQcUpdates(collectGraphicsQcUpdates());
  if (!updates.length) {
    setStatus("No changed QC decisions to submit.");
    return;
  }

  const noteRequired = updates.find(update => (
    (
      update.qcDecision === "reject" &&
      !update.rejectReason
    ) ||
    (
      update.qcDecision === "reject" &&
      (
        (update.rejectReason === "mismatched_graphic" && !update.qcNote) ||
        (
          update.rejectReason === "correct_and_recreate" &&
          (
            (!update.metadataIssue && !update.aestheticIssue) ||
            (update.metadataIssue && !update.qcNote) ||
            (update.aestheticIssue === "other" && !update.qcNote)
          )
        )
      )
    )
  ));
  if (noteRequired) {
    setStatus("Add the missing QC detail. Rejects need a reason. Mismatched graphics need a note. Correct and recreate needs at least one issue selected, plus a note for metadata fixes or aesthetic 'other'.");
    return;
  }

  try {
    setSubmitState(true, `Saving ${updates.length}...`);
    setStatus(`Saving ${updates.length} QC decisions...`);

    const response = await fetch("/api/save-graphics-qc", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        updates
      })
    });
    const result = await response.json();
    if (!response.ok || !result.ok) {
      throw new Error(result.error || `/api/save-graphics-qc returned ${response.status}`);
    }

    await loadGraphicsRecords();
    const handoff = result.poetryPlease || {};
    const handoffMessage = handoff.ok && !handoff.skipped
      ? ` Sent ${Number(handoff.createdCount || 0) + Number(handoff.updatedCount || 0)} approved graphic${(Number(handoff.createdCount || 0) + Number(handoff.updatedCount || 0)) === 1 ? "" : "s"} to Poetry Please.`
      : (!handoff.ok && handoff.error ? ` Poetry Please handoff needs attention: ${handoff.error}` : "");
    setStatus(`Saved ${updates.length} QC decisions.${handoffMessage}`);
  } catch (error) {
    setStatus(`Graphics QC save failed: ${error.message}`);
  } finally {
    setSubmitState(false, "Submit Decisions");
  }
}

async function exportGraphicsSheet() {
  const mode = getSelectedGraphicsMode();
  if (mode !== "queue") {
    setStatus("Google Sheet export is only available in Graphics Creation Queue mode.");
    return;
  }

  const bookKey = elements.graphicsBookSelect?.value || "";
  if (!bookKey) {
    setStatus("Choose a graphics book title first.");
    return;
  }

  if (!currentGraphicsRecords.length) {
    setStatus("Load queue rows before exporting.");
    return;
  }

  const summary = graphicsBookSummaryByKey.get(bookKey);
  const bookTitle = summary?.title || bookKey;

  try {
    if (elements.exportGraphicsSheet) {
      elements.exportGraphicsSheet.disabled = true;
      elements.exportGraphicsSheet.textContent = "Exporting...";
    }

    if (!intakeOptionsLoaded) {
      await loadGatheringOptions();
    }

    setStatus(`Creating Google Sheet export for "${bookTitle}"...`);
    const result = await createGoogleSheetFromRows({
      title: buildGraphicsExportTitle(bookTitle),
      headers: buildGraphicsExportHeaders(),
      rows: buildGraphicsExportRows(currentGraphicsRecords)
    });

    if (result.spreadsheetUrl) {
      const opened = window.open(result.spreadsheetUrl, "_blank", "noopener,noreferrer");
      if (!opened) {
        showSheetLinkModal(result.spreadsheetUrl);
      }
    }

    setStatus(
      `Created Google Sheet export for "${bookTitle}" (${currentGraphicsRecords.length} rows).`,
      {
        spreadsheetUrl: result.spreadsheetUrl,
        spreadsheetId: result.spreadsheetId
      }
    );
  } catch (error) {
    setStatus(`Graphics export failed: ${error.message}`);
  } finally {
    if (elements.exportGraphicsSheet) {
      elements.exportGraphicsSheet.disabled = false;
      elements.exportGraphicsSheet.textContent = "Export Google Sheet";
    }
  }
}

function buildValidationMarkup(validation, excerpt) {
  if (!validation) {
    return `<p class="validation validation--pending">Catalog validation unavailable for this excerpt.</p>`;
  }

  const canonical = [
    validation.bookCanonicalTitle,
    validation.bookCanonicalAuthor
  ].filter(Boolean).join(" - ");
  const displayAuthor = resolveExcerptDisplayAuthor(excerpt, validation);
  const libraryMarkup = buildLibraryMatchMarkup(validation.libraryExcerptMatch);
  const poemLink = buildCatalogPoemLink(validation, excerpt);
  const catalogFormattingMarkup = buildCatalogFormattingMarkup(validation);

  if (validation.status === "catalog_match") {
    return `<p class="validation validation--good">Catalog match: ${escapeHtml(canonical || "confirmed")}. ${poemLink}</p>${catalogFormattingMarkup}${libraryMarkup}`;
  }

  if (validation.status === "author_mismatch") {
    if (
      cleanSheetWhitespace(displayAuthor).toLowerCase()
      && cleanSheetWhitespace(validation.bookCanonicalAuthor).toLowerCase()
      && cleanSheetWhitespace(displayAuthor).toLowerCase() === cleanSheetWhitespace(validation.bookCanonicalAuthor).toLowerCase()
    ) {
      return `<p class="validation validation--good">Catalog author confirmed: ${escapeHtml(validation.bookCanonicalAuthor)}. ${poemLink}</p>${catalogFormattingMarkup}${libraryMarkup}`;
    }
    return `<p class="validation validation--warn">Author mismatch. Catalog author: ${escapeHtml(validation.bookCanonicalAuthor || "different author")}. ${poemLink}</p>${catalogFormattingMarkup}${libraryMarkup}`;
  }

  if (validation.status === "title_mismatch") {
    return `<p class="validation validation--warn">Excerpt matches the catalog, but the poem title appears wrong. Catalog match: ${escapeHtml(validation.matchedPoemTitle || "different title")}. ${poemLink}</p>${catalogFormattingMarkup}${libraryMarkup}`;
  }

  if (validation.status === "poem_title_match_only") {
    return `<p class="validation validation--warn">Poem title matches this book, but the excerpt text did not match the catalog text. ${poemLink}</p>${libraryMarkup}`;
  }

  if (validation.status === "epub_not_present") {
    return `<p class="validation validation--warn">This title appears to be intentionally absent from EPUB/catalog coverage, not simply mismatched.</p>${libraryMarkup}`;
  }

  if (validation.status === "catalog_unavailable") {
    return `<p class="validation validation--warn">${escapeHtml(getCatalogUnavailableMessage(validation))}</p>${libraryMarkup}`;
  }

  if (validation.status === "excerpt_not_found_in_book" && validation.globalExcerptMatch) {
    return `<p class="validation validation--warn">This book has usable catalog coverage, but this poem title does not match a poem in that book, and this excerpt text was not found there either. Closest catalog hit: ${escapeHtml(validation.globalExcerptMatch.book_title)} / ${escapeHtml(validation.globalExcerptMatch.poem_title)} by ${escapeHtml(validation.globalExcerptMatch.author)}. ${poemLink}</p>${libraryMarkup}`;
  }

  if (validation.status === "excerpt_not_found_in_book") {
    return `<p class="validation validation--warn">This book has usable catalog coverage, but this poem title does not match a poem in that book, and this excerpt text was not found there either.</p>${libraryMarkup}`;
  }

  if (validation.status === "book_not_found") {
    return `<p class="validation validation--warn">This book is not in the catalog yet.</p>${libraryMarkup}`;
  }

  return `<p class="validation validation--warn">Catalog check: ${escapeHtml(validation.status)}.</p>${libraryMarkup}`;
}

function getCatalogUnavailableMessage(validation) {
  const effectiveStatus = cleanSheetWhitespace(validation?.bookEffectiveStatus || "").toLowerCase();
  if (effectiveStatus === "set_aside") {
    return "This title is tracked in the catalog system, but no usable catalog poem text is available yet.";
  }
  if (effectiveStatus === "source_present_extraction_failed") {
    return "This title has a source file, but catalog poem extraction failed, so poem lookup is unavailable right now.";
  }
  if (effectiveStatus === "source_present_not_in_catalog") {
    return "This title has a source file, but it has not been turned into catalog poem data yet.";
  }
  if (effectiveStatus === "broken_source") {
    return "This title is tracked in the catalog system, but its source file is currently broken or unusable.";
  }
  return "This title is tracked in the catalog system, but no usable catalog poem text is available for lookup right now.";
}

function buildCatalogFormattingMarkup(validation) {
  if (!validation || validation.catalogFormattingMatch !== false) {
    return "";
  }

  return `<p class="validation validation--pending">Catalog text matches, but the formatting or line breaks differ from the catalog poem.</p>`;
}

function buildCatalogPoemLink(validation, excerpt) {
  const status = validation?.status || "";
  const linkableStatuses = new Set([
    "catalog_match",
    "author_mismatch",
    "title_mismatch",
    "poem_title_match_only"
  ]);

  if (!linkableStatuses.has(status) && !(status === "excerpt_not_found_in_book" && validation.globalExcerptMatch)) {
    return "";
  }

  let bookTitle = validation.bookCanonicalTitle || excerpt.bookTitle || "";
  let poemTitle = validation.matchedPoemTitle || "";

  if (status === "poem_title_match_only") {
    poemTitle = excerpt.title || "";
  }

  if (status === "excerpt_not_found_in_book" && validation.globalExcerptMatch) {
    bookTitle = validation.globalExcerptMatch.book_title || bookTitle;
    poemTitle = validation.globalExcerptMatch.poem_title || poemTitle;
  }

  if (!bookTitle || !poemTitle) {
    return "";
  }

  const url = new URL("/catalog-poem", window.location.origin);
  url.searchParams.set("bookTitle", bookTitle);
  url.searchParams.set("poemTitle", poemTitle);
  url.searchParams.set("excerptText", excerpt.excerptText || "");
  return `<button type="button" class="catalog-poem-link button-link" data-popup-url="${escapeHtml(url.toString())}">View catalog poem</button>`;
}

function buildLibraryMatchMarkup(match) {
  if (!match) {
    return "";
  }

  const label = getLibraryMatchSentence(match);
  const statusLabel = getLibraryProductionStatusLabel(match.libraryStatus);
  const meta = [
    match.bookTitle,
    match.poemTitle,
    match.author
  ].filter(Boolean).join(" / ");
  const excerptLink = buildLibraryExcerptLink(match);
  const statusMarkup = statusLabel
    ? ` <span class="library-status-note">${escapeHtml(statusLabel)}</span>`
    : "";
  return `<p class="validation validation--warn">${escapeHtml(label)} ${escapeHtml(meta || "Existing source row")}${match.sourceRow ? `, row ${escapeHtml(String(match.sourceRow))}` : ""}.${statusMarkup} ${excerptLink}</p>`;
}

function getLibraryBadgeLabel(match) {
  if (!match) {
    return "";
  }
  if (match.matchType === "exact") {
    return match.lineBreaksMatch === false
      ? "Text match, line breaks differ"
      : "Exact library match";
  }
  return "Possible library match";
}

function getLibraryMatchSentence(match) {
  if (!match) {
    return "";
  }

  if (match.matchType === "exact") {
    if (match.lineBreaksMatch === false) {
      return "Same excerpt text exists in the excerpt library, but the line breaks differ.";
    }
    return "Exact text and formatting match already exists in the excerpt library.";
  }

  if (match.matchType === "substring") {
    return `Possible existing excerpt in library (${match.matchType}, score ${match.score}).`;
  }

  return `Possible near-duplicate in library (score ${match.score}).`;
}

function getLibraryProductionStatusLabel(status) {
  if (!status) {
    return "";
  }

  if (status.made) {
    return "Graphic appears to have already been made.";
  }

  if (status.approvedForQi) {
    return "Approved for quote image, but not confirmed made.";
  }

  if (status.hasQiAsset) {
    return "QI asset exists, but no approval or made status was found.";
  }

  return "In excerpt library only; no quote-image status found.";
}

function buildLibraryExcerptLink(match) {
  if (!match || !match.sourceRow) {
    return "";
  }

  const url = new URL("/library-excerpt", window.location.origin);
  url.searchParams.set("sourceRow", String(match.sourceRow));
  return `<button type="button" class="library-excerpt-link button-link" data-popup-url="${escapeHtml(url.toString())}">View library excerpt</button>`;
}

function slugify(text) {
  return (text || "untitled")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");
}

function escapeHtml(text) {
  return (text || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function escapeAttribute(text) {
  return escapeHtml(text).replaceAll('"', "&quot;");
}

function collectUpdates() {
  return collectUpdatesFromContainer(elements.excerptList);
}

function collectCorrectionUpdates() {
  return collectUpdatesFromContainer(elements.correctionList);
}

function getAutoCorrectionProposal(excerpt) {
  const validation =
    currentValidationByRecordId.get(excerpt.recordId || String(excerpt.sourceRow)) || null;
  if (!validation) {
    return null;
  }

  const proposal = {
    correctedAuthor: "",
    correctedTitle: "",
    correctedBookTitle: "",
    correctedExcerpt: "",
    reasonParts: []
  };

  const canonicalAuthor = cleanSheetWhitespace(validation.bookCanonicalAuthor || "");
  const canonicalTitle = cleanSheetWhitespace(validation.matchedPoemTitle || "");
  const canonicalBookTitle = cleanSheetWhitespace(validation.bookCanonicalTitle || "");
  const libraryMatch = validation.libraryExcerptMatch || null;
  const exactLibraryMatch = libraryMatch?.matchType === "exact" ? libraryMatch : null;
  const libraryAuthor = cleanSheetWhitespace(exactLibraryMatch?.author || "");
  const libraryTitle = cleanSheetWhitespace(exactLibraryMatch?.poemTitle || "");
  const libraryBookTitle = cleanSheetWhitespace(exactLibraryMatch?.bookTitle || "");
  const currentAuthor = cleanSheetWhitespace(excerpt.author || excerpt.rawAuthor || "");
  const currentTitle = cleanSheetWhitespace(excerpt.title || excerpt.rawTitle || "");
  const currentBookTitle = cleanSheetWhitespace(excerpt.bookTitle || excerpt.rawBookTitle || "");

  if (validation.status === "author_mismatch" && canonicalAuthor && canonicalAuthor !== currentAuthor) {
    proposal.correctedAuthor = canonicalAuthor;
    proposal.reasonParts.push("author");
  }

  if (validation.status === "title_mismatch" && canonicalTitle && canonicalTitle !== currentTitle) {
    proposal.correctedTitle = canonicalTitle;
    proposal.reasonParts.push("poem title");
  }

  if ((validation.status === "author_mismatch" || validation.status === "title_mismatch" || validation.status === "catalog_match") && canonicalBookTitle && canonicalBookTitle !== currentBookTitle) {
    proposal.correctedBookTitle = canonicalBookTitle;
    proposal.reasonParts.push("book title");
  }

  if (!proposal.correctedAuthor && exactLibraryMatch && libraryAuthor && libraryAuthor !== currentAuthor) {
    proposal.correctedAuthor = libraryAuthor;
    proposal.reasonParts.push("author");
  }

  if (!proposal.correctedTitle && exactLibraryMatch && libraryTitle && libraryTitle !== currentTitle) {
    proposal.correctedTitle = libraryTitle;
    proposal.reasonParts.push("poem title");
  }

  if (!proposal.correctedBookTitle && exactLibraryMatch && libraryBookTitle && libraryBookTitle !== currentBookTitle) {
    proposal.correctedBookTitle = libraryBookTitle;
    proposal.reasonParts.push("book title");
  }

  if (!proposal.reasonParts.length) {
    return null;
  }

  return proposal;
}

function applyAutoCorrectionsToLoadedQueue() {
  if (!currentCorrectionExcerpts.length || !elements.correctionList) {
    setStatus("Load correction records before applying auto-fixes.");
    return;
  }

  let appliedCount = 0;

  currentCorrectionExcerpts.forEach(excerpt => {
    const proposal = getAutoCorrectionProposal(excerpt);
    if (!proposal) return;

    const card = elements.correctionList.querySelector(`.excerpt-card[data-source-row="${CSS.escape(String(excerpt.sourceRow))}"]`);
    if (!card) return;

    const correctedAuthor = card.querySelector(".corrected-author");
    const correctedTitle = card.querySelector(".corrected-title");
    const correctedBookTitle = card.querySelector(".corrected-book-title");
    const noDecisionRadio = card.querySelector(`input[type="radio"][value=""]`);

    if (proposal.correctedAuthor && correctedAuthor) correctedAuthor.value = proposal.correctedAuthor;
    if (proposal.correctedTitle && correctedTitle) correctedTitle.value = proposal.correctedTitle;
    if (proposal.correctedBookTitle && correctedBookTitle) correctedBookTitle.value = proposal.correctedBookTitle;
    if (noDecisionRadio) {
      noDecisionRadio.checked = true;
      noDecisionRadio.dispatchEvent(new Event("change", { bubbles: true }));
    }

    appliedCount += 1;
  });

  if (!appliedCount) {
    setStatus("No loaded correction rows had a safe automatic metadata fix.");
    return;
  }

  setStatus(`Applied ${appliedCount} safe auto-fix${appliedCount === 1 ? "" : "es"}. Save Corrections to move them back into review.`);
}

function collectWeirdUpdates() {
  return collectUpdatesFromContainer(elements.weirdExcerptList);
}

function collectUpdatesFromContainer(container) {
  return Array.from(container.querySelectorAll(".excerpt-card")).map(card => {
    const reviewDecision = card.querySelector('input[type="radio"]:checked')?.value || "";
    const sourceRow = Number(card.dataset.sourceRow);
    const validation =
      currentValidationByRecordId.get(card.dataset.recordId || String(sourceRow)) || null;
    const libraryMatch = validation?.libraryExcerptMatch || null;
    const duplicateGroupId =
      reviewDecision === "reject" &&
      libraryMatch?.matchType === "exact" &&
      libraryMatch?.lineBreaksMatch === false &&
      libraryMatch?.sourceRow
        ? `library-variant:${libraryMatch.sourceRow}`
        : "";

    return {
      sourceRow,
      recordId: card.dataset.recordId || "",
      reviewDecision,
      correctionNote: card.querySelector(".correction-note")?.value || "",
      correctedAuthor: card.querySelector(".corrected-author")?.value || "",
      correctedTitle: card.querySelector(".corrected-title")?.value || "",
      correctedBookTitle: card.querySelector(".corrected-book-title")?.value || "",
      correctedExcerpt: card.querySelector(".corrected-excerpt")?.value || "",
      duplicateGroupId,
      useForQi: reviewDecision === "accept",
      useForInt: card.dataset.currentUseForInt === "1"
    };
  });
}

function filterChangedUpdates(updates, excerpts) {
  const bySourceRow = new Map(
    excerpts.map(excerpt => [Number(excerpt.sourceRow), excerpt])
  );

  return updates.filter(update => {
    const current = bySourceRow.get(Number(update.sourceRow));
    if (!current) return true;

    return (
      normalizeDecision(update.reviewDecision) !== getEffectiveReviewDecision(current) ||
      normalizeCorrectionNote(update.correctionNote) !== normalizeCorrectionNote(current.correctionNote) ||
      normalizeCorrectionNote(update.correctedAuthor) !== normalizeCorrectionNote(current.correctedAuthor) ||
      normalizeCorrectionNote(update.correctedTitle) !== normalizeCorrectionNote(current.correctedTitle) ||
      normalizeCorrectionNote(update.correctedBookTitle) !== normalizeCorrectionNote(current.correctedBookTitle) ||
      normalizeCorrectionNote(update.correctedExcerpt) !== normalizeCorrectionNote(current.correctedExcerpt) ||
      Boolean(update.useForQi) !== Boolean(current.useForQi ?? current.useForGraphicsQi)
    );
  });
}

function normalizeCorrectionNote(value) {
  return (value || "").toString().trim();
}

function normalizeDecision(value) {
  return (value || "").toString().trim().toLowerCase();
}

function normalizeApprovalForCompare(value) {
  const normalized = (value || "").toString().trim().toLowerCase();
  if (normalized === "accept" || normalized === "y") return "Y";
  if (normalized === "reject" || normalized === "n") return "N";
  return "";
}

function sleep(ms) {
  return new Promise(resolve => {
    window.setTimeout(resolve, ms);
  });
}

function compareUpdatesToExcerpts(updates, excerpts) {
  const bySourceRow = new Map(
    excerpts.map(excerpt => [Number(excerpt.sourceRow), excerpt])
  );

  let matched = 0;

  updates.forEach(update => {
    const saved = bySourceRow.get(Number(update.sourceRow));
    if (!saved) return;

    const decisionMatches =
      normalizeDecision(update.reviewDecision) === getEffectiveReviewDecision(saved);
    const authorMatches =
      normalizeCorrectionNote(update.correctedAuthor) === normalizeCorrectionNote(saved.correctedAuthor);
    const titleMatches =
      normalizeCorrectionNote(update.correctedTitle) === normalizeCorrectionNote(saved.correctedTitle);
    const bookMatches =
      normalizeCorrectionNote(update.correctedBookTitle) === normalizeCorrectionNote(saved.correctedBookTitle);
    const excerptMatches =
      normalizeCorrectionNote(update.correctedExcerpt) === normalizeCorrectionNote(saved.correctedExcerpt);
    const qiMatches =
      Boolean(update.useForQi) === Boolean(saved.useForQi ?? saved.useForGraphicsQi);

    if (decisionMatches && authorMatches && titleMatches && bookMatches && excerptMatches && qiMatches) {
      matched += 1;
    }
  });

  return {
    matched,
    total: updates.length
  };
}

function countRemainingUpdatedRows(updates, excerpts) {
  const remaining = new Set(
    excerpts.map(excerpt => Number(excerpt.sourceRow))
  );

  return updates.filter(update => remaining.has(Number(update.sourceRow))).length;
}

function countWordsFromText(text) {
  const cleaned = (text || "")
    .toString()
    .replace(/[\u2013\u2014]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!cleaned) return 0;

  const matches = cleaned.match(/[A-Za-z0-9']+/g);
  return matches ? matches.length : 0;
}

function resolveWordCount(excerpt) {
  const backendWordCount = Number(excerpt.wordCount);
  if (Number.isFinite(backendWordCount) && backendWordCount > 0) {
    return backendWordCount;
  }

  const candidates = [
    excerpt.excerptText,
    excerpt.excerpt,
    excerpt.text,
    excerpt.rawExcerpt
  ];

  for (const candidate of candidates) {
    const count = countWordsFromText(candidate);
    if (count > 0) {
      return count;
    }
  }

  return 0;
}

function getWordCountBadgeClass(wordCount) {
  if (wordCount <= 0) return "badge--count";
  if (wordCount <= 25) return "badge--count";
  if (wordCount <= 35) return "badge--count badge--count-caution";
  if (wordCount <= 50) return "badge--count badge--count-warning";
  return "badge--count badge--count-danger";
}

async function submitReview() {
  const pinnedSourceRows = Array.from(
    elements.excerptList.querySelectorAll(".excerpt-card")
  ).map(card => Number(card.dataset.sourceRow)).filter(Boolean);
  const bookKey = elements.bookSelect.value;

  return submitExcerptSet({
    sourceExcerpts: currentExcerpts,
    collectUpdates: collectUpdates,
    bookKey,
    reloadRequest: () => requestMergedBookRecords(currentExcerpts),
    afterSaveOptimistic: changedUpdates => {
      const savedSourceRows = new Set(changedUpdates.map(update => Number(update.sourceRow)));
      currentExcerpts = currentExcerpts.filter(excerpt => !savedSourceRows.has(Number(excerpt.sourceRow)));
      reviewPinnedRowOrder = [];
      renderCurrentExcerpts();
    },
    afterReload: async (refreshed, changedUpdates) => {
      const savedSourceRows = new Set(changedUpdates.map(update => Number(update.sourceRow)));
      const records = Array.isArray(refreshed.records) ? refreshed.records : [];
      applyPendingBookData(records, { preserveSelection: true });
      currentExcerpts = getPendingRecordsForBookKey(bookKey, records)
        .filter(excerpt => !savedSourceRows.has(Number(excerpt.sourceRow)));
      const remainingSourceRows = new Set(
        currentExcerpts.map(excerpt => Number(excerpt.sourceRow))
      );
      reviewPinnedRowOrder = pinnedSourceRows.filter(sourceRow => remainingSourceRows.has(sourceRow));
      await loadCatalogValidation(currentExcerpts);
      renderCurrentExcerpts();
      refreshBookCountsInBackground();
    },
    countExcerpts: refreshed => applyReviewFilter(
      getPendingRecordsForBookKey(bookKey, Array.isArray(refreshed.records) ? refreshed.records : [])
    ),
    emptyMessage: "Load excerpts before submitting.",
    idleLabel: "Submit Decisions",
    progressLabel: "Submitting"
  });
}

async function submitWeirdReview() {
  const pinnedSourceRows = Array.from(
    elements.weirdExcerptList.querySelectorAll(".excerpt-card")
  ).map(card => Number(card.dataset.sourceRow)).filter(Boolean);
  const bookKey = elements.weirdBookSelect.value;

  return submitExcerptSet({
    sourceExcerpts: currentWeirdExcerpts,
    collectUpdates: collectWeirdUpdates,
    bookKey,
    reloadRequest: () => requestMergedBookRecords(currentWeirdExcerpts),
    afterSaveOptimistic: changedUpdates => {
      const savedSourceRows = new Set(changedUpdates.map(update => Number(update.sourceRow)));
      currentWeirdExcerpts = currentWeirdExcerpts.filter(excerpt => !savedSourceRows.has(Number(excerpt.sourceRow)));
      weirdPinnedRowOrder = [];
      renderWeirdCurrentExcerpts();
    },
    afterReload: async (refreshed, changedUpdates) => {
      const savedSourceRows = new Set(changedUpdates.map(update => Number(update.sourceRow)));
      const records = Array.isArray(refreshed.records) ? refreshed.records : [];
      applyPendingBookData(records, { preserveSelection: true });
      currentWeirdExcerpts = getPendingRecordsForBookKey(bookKey, records)
        .filter(excerpt => !savedSourceRows.has(Number(excerpt.sourceRow)));
      const remainingSourceRows = new Set(
        applyExtraReviewFilter(currentWeirdExcerpts).map(excerpt => Number(excerpt.sourceRow))
      );
      weirdPinnedRowOrder = pinnedSourceRows.filter(sourceRow => remainingSourceRows.has(sourceRow));
      await loadCatalogValidation(currentWeirdExcerpts);
      renderWeirdCurrentExcerpts();
      refreshBookCountsInBackground();
    },
    countExcerpts: refreshed => applyExtraReviewFilter(
      getPendingRecordsForBookKey(bookKey, Array.isArray(refreshed.records) ? refreshed.records : [])
    ),
    emptyMessage: "Load extra-review excerpts before submitting.",
    idleLabel: "Submit Decisions",
    progressLabel: "Submitting"
  });
}

async function submitCorrections() {
  return submitExcerptSet({
    sourceExcerpts: currentCorrectionExcerpts,
    collectUpdates: collectCorrectionUpdates,
    bookKey: elements.correctionBookSelect.value,
    reloadRequest: () => requestReviewApi("/api/review/corrections", { bookTitle: elements.correctionBookSelect.value }),
    afterSaveOptimistic: null,
    afterReload: async refreshed => {
      currentCorrectionExcerpts = refreshed.excerpts;
      await loadCatalogValidation(refreshed.excerpts);
      renderCorrectionExcerpts(refreshed.excerpts);
      loadCorrectionBooks().catch(() => {});
    },
    countExcerpts: refreshed => refreshed.excerpts,
    emptyMessage: "Load correction records before saving.",
    idleLabel: "Save Corrections",
    progressLabel: "Saving"
  });
}

async function submitExcerptSet({
  sourceExcerpts,
  collectUpdates,
  bookKey,
  reloadRequest,
  afterSaveOptimistic,
  afterReload,
  countExcerpts,
  emptyMessage,
  idleLabel,
  progressLabel
}) {
  if (!sourceExcerpts.length) {
    setStatus(emptyMessage);
    return;
  }

  const updates = collectUpdates();
  const changedUpdates = filterChangedUpdates(updates, sourceExcerpts);

  if (!changedUpdates.length) {
    setStatus("No changed decisions to submit.");
    return;
  }

  const correctionWithoutNote = changedUpdates.find(update => (
    update.reviewDecision === "needs_correction" && !normalizeCorrectionNote(update.correctionNote)
  ));
  if (correctionWithoutNote) {
    setStatus("Add a correction note before submitting a 'Needs correction' decision.");
    return;
  }

  try {
    setSubmitState(true, `Saving ${changedUpdates.length}...`);
    setStatus(`${progressLabel} ${changedUpdates.length} changed records...`);

    let usedFallback = false;
    try {
      await requestBatchSave(changedUpdates);
    } catch (batchError) {
      usedFallback = true;
      setStatus(
        `Batch save failed, falling back to row-by-row saves for ${changedUpdates.length} records...`,
        { error: batchError.message }
      );
      await saveReviewsSequentially(changedUpdates);
    }

    if (typeof afterSaveOptimistic === "function") {
      afterSaveOptimistic(changedUpdates);
    }

    setStatus(`Submitted ${changedUpdates.length} changed decisions. Verifying saved values...`);

    const refreshed = await reloadRequest();
    if (!refreshed.ok) {
      throw new Error(refreshed.error || "Verification reload failed.");
    }

    await afterReload(refreshed, changedUpdates);
    const countPool = typeof countExcerpts === "function" ? countExcerpts(refreshed) : refreshed.excerpts;
    const remainingUpdatedRows = countRemainingUpdatedRows(changedUpdates, countPool);
    const droppedFromQueue = changedUpdates.length - remainingUpdatedRows;

    if (remainingUpdatedRows === 0) {
      setStatus(
        usedFallback
          ? `Saved ${changedUpdates.length} changed decisions with fallback mode. Those excerpts dropped out of the pending queue.`
          : `Saved ${changedUpdates.length} changed decisions. Those excerpts dropped out of the pending queue.`
      );
    } else {
      setStatus(
        usedFallback
          ? `Saved ${changedUpdates.length} changed decisions with fallback mode. ${droppedFromQueue} dropped out of queue and ${remainingUpdatedRows} are still showing.`
          : `Saved ${changedUpdates.length} changed decisions. ${droppedFromQueue} dropped out of queue and ${remainingUpdatedRows} are still showing.`,
        (Array.isArray(refreshed.excerpts) ? refreshed.excerpts : Array.isArray(refreshed.records) ? refreshed.records : []).slice(0, 5)
      );
    }
  } catch (error) {
    setStatus(`Submit failed: ${error.message}`);
  } finally {
    setSubmitState(false, idleLabel);
  }
}

elements.loadBooks.addEventListener("click", loadBooks);
elements.loadExcerpts.addEventListener("click", loadExcerpts);
elements.loadWeirdExcerpts?.addEventListener("click", loadWeirdExcerpts);
elements.submitReview.addEventListener("click", submitReview);
elements.submitWeirdReview?.addEventListener("click", submitWeirdReview);
elements.loadCorrectionBooks?.addEventListener("click", loadCorrectionBooks);
elements.loadCorrections?.addEventListener("click", loadCorrections);
elements.autoApplyCorrections?.addEventListener("click", applyAutoCorrectionsToLoadedQueue);
elements.submitCorrections?.addEventListener("click", submitCorrections);
elements.showGatheringModule?.addEventListener("click", () => {
  setActiveModule("gathering");
  if (!intakeOptionsLoaded) {
    loadGatheringOptions().catch(error => {
      setStatus(`Excerpt gathering option load failed: ${error.message}`);
    });
  }
});
elements.gatheringTabBook?.addEventListener("click", () => setGatheringMode("book"));
elements.gatheringTabVideo?.addEventListener("click", () => setGatheringMode("video"));
elements.gatheringTabFix?.addEventListener("click", () => setGatheringMode("fix"));
elements.showReviewModule?.addEventListener("click", () => setActiveModule("review"));
elements.showWeirdModule?.addEventListener("click", () => {
  setActiveModule("weird");
  if (elements.weirdBookSelect?.options.length <= 1) {
    loadBooks();
  }
});
elements.showCorrectionsModule?.addEventListener("click", () => {
  setActiveModule("corrections");
  if (elements.correctionBookSelect?.options.length <= 1) {
    loadCorrectionBooks();
  }
});
elements.showGraphicsModule?.addEventListener("click", () => {
  setActiveModule("graphics");
  if (elements.graphicsBookSelect?.options.length <= 1) {
    loadGraphicsBooks();
  }
});
elements.loadGraphicsBooks?.addEventListener("click", loadGraphicsBooks);
elements.loadGraphicsRecords?.addEventListener("click", loadGraphicsRecords);
elements.previewGraphicsFolderImport?.addEventListener("click", previewGraphicsFolderImport);
elements.applyGraphicsFolderImport?.addEventListener("click", applyGraphicsFolderImport);
elements.graphicsFolderImportResults?.addEventListener("change", event => {
  if (
    (event.target instanceof HTMLInputElement && event.target.type === "radio") ||
    (event.target instanceof HTMLSelectElement && event.target.classList.contains("folder-import-manual-select"))
  ) {
    refreshGraphicsFolderImportApplyState();
  }
});
elements.exportGraphicsSheet?.addEventListener("click", exportGraphicsSheet);
elements.submitGraphicsQc?.addEventListener("click", submitGraphicsQc);
elements.loadGatheringOptions?.addEventListener("click", () => {
  loadGatheringOptions({ force: true }).catch(error => {
    setStatus(`Excerpt gathering option load failed: ${error.message}`);
  });
});
elements.submitGathering?.addEventListener("click", submitGathering);
elements.gatheringMode?.addEventListener("change", updateGatheringModeUi);
elements.gatheringBookBook?.addEventListener("change", () => {
  handleGatheringBookSelectionChange().catch(error => {
    setStatus(`Book metadata load failed: ${error.message}`);
  });
});
elements.gatheringBookBook?.addEventListener("blur", () => {
  const enteredBook = elements.gatheringBookBook?.value || "";
  const catalogBook = currentIntakeCatalogBooksByKey.get(normalizeBookKey(enteredBook));
  if (catalogBook && elements.gatheringBookBook) {
    elements.gatheringBookBook.value = catalogBook.title;
  }
  handleGatheringBookSelectionChange({ preserveTitle: true }).catch(error => {
    setStatus(`Book metadata load failed: ${error.message}`);
  });
});
elements.gatheringBookTitle?.addEventListener("input", updateGatheringCatalogPreviewState);
elements.gatheringBookTitle?.addEventListener("change", updateGatheringCatalogPreviewState);
elements.gatheringPrevCatalogPoem?.addEventListener("click", () => navigateGatheringCatalogPoem(-1));
elements.gatheringNextCatalogPoem?.addEventListener("click", () => navigateGatheringCatalogPoem(1));
elements.gatheringBookAuthor?.addEventListener("input", () => {
  if (!elements.gatheringBookBook?.value.trim()) {
    setGatheringBookManualMode("");
  }
});
elements.gatheringViewCatalogPoem?.addEventListener("click", openGatheringCatalogPoem);
elements.gatheringBookQuote?.addEventListener("input", updateGatheringQuoteMeta);
elements.gatheringVideoQuote?.addEventListener("input", updateGatheringQuoteMeta);
[
    elements.gatheringBookQuote,
    elements.gatheringBookNotes,
    elements.gatheringBookItalics,
    elements.gatheringBookReaction,
    elements.gatheringVideoQuote,
  elements.gatheringFixIncorrect,
  elements.gatheringFixCorrect
].forEach(field => field?.addEventListener("keydown", handleGatheringQuickSubmit));
if (elements.reviewFilter) {
  elements.reviewFilter.addEventListener("change", () => {
    refreshReviewBookSelect(true);
    reviewVisibleCount = getReviewBatchSize();
    reviewPinnedRowOrder = [];
    renderCurrentExcerpts();
  });
}
if (elements.reviewDisplayMode) {
  elements.reviewDisplayMode.addEventListener("change", () => {
    reviewVisibleCount = getReviewBatchSize();
    reviewPinnedRowOrder = [];
    renderCurrentExcerpts();
  });
}
if (elements.weirdReviewFilter) {
  elements.weirdReviewFilter.addEventListener("change", () => {
    weirdVisibleCount = EXTRA_REVIEW_BATCH_SIZE;
    weirdPinnedRowOrder = [];
    renderWeirdCurrentExcerpts();
  });
}
if (elements.graphicsMode) {
  elements.graphicsMode.addEventListener("change", () => {
    if (elements.graphicsFilter) {
      elements.graphicsFilter.disabled = getSelectedGraphicsMode() !== "queue";
    }
    refreshGraphicsFolderImportVisibility();
    currentGraphicsRecords = [];
    currentGraphicsAssetMatches = new Map();
    renderGraphicsFolderImportPreview(null);
    renderGraphicsRecords([]);
    if (elements.graphicsBookSelect) {
      elements.graphicsBookSelect.innerHTML = '<option value="">Load graphics books first</option>';
    }
    loadGraphicsBooks({ preserveSelection: false });
  });
}
if (elements.graphicsFilter) {
  elements.graphicsFilter.disabled = getSelectedGraphicsMode() !== "queue";
  elements.graphicsFilter.addEventListener("change", () => {
    refreshGraphicsBookSelect(true);
  });
}

async function initializeApp() {
  await loadRuntimeConfig();
  applyRuntimeMode();
  updateGatheringModeUi();
  refreshGraphicsFolderImportVisibility();
  renderGraphicsFolderImportPreview(null);
  setActiveModule("review");
  setStatus(`Ready${runtimeConfig.appVersion ? ` (${runtimeConfig.appVersion})` : ""}. Loading books...`);
  loadBooks();
}

initializeApp();
