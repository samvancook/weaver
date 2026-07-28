let runtimeConfig = window.WEAVER_CONFIG || {};
let currentGraphicsView = "main";

const elements = {
  showGatheringModule: document.getElementById("show-gathering-module"),
  showReviewModule: document.getElementById("show-review-module"),
  showWeirdModule: document.getElementById("show-weird-module"),
  showCorrectionsModule: document.getElementById("show-corrections-module"),
  showGraphicsModule: document.getElementById("show-graphics-module"),
  showGraphicsOps: document.getElementById("show-graphics-ops"),
  graphicsModuleTitle: document.getElementById("graphics-module-title"),
  graphicsOpsPanel: document.getElementById("graphics-ops-panel"),
  graphicsBookPicker: document.getElementById("graphics-book-picker"),
  graphicsModeField: document.getElementById("graphics-mode-field"),
  graphicsFilterField: document.getElementById("graphics-filter-field"),
  graphicsReleaseCatalogField: document.getElementById("graphics-release-catalog-field"),
  graphicsMainHint: document.getElementById("graphics-main-hint"),
  graphicsQueueRowsHeader: document.getElementById("graphics-queue-rows-header"),
  graphicsQueueToolbarTop: document.getElementById("graphics-queue-toolbar-top"),
  graphicsQueueToolbarBottom: document.getElementById("graphics-queue-toolbar-bottom"),
  gatheringTabBook: document.getElementById("gathering-tab-book"),
  gatheringTabVideo: document.getElementById("gathering-tab-video"),
  gatheringTabFix: document.getElementById("gathering-tab-fix"),
  gatheringTabBatch: document.getElementById("gathering-tab-batch"),
  gatheringGuidanceLead: document.getElementById("gathering-guidance-lead"),
  gatheringGuidanceList: document.getElementById("gathering-guidance-list"),
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
  reviewReleaseCatalog: document.getElementById("review-release-catalog"),
  reviewDisplayMode: document.getElementById("review-display-mode"),
  weirdReviewFilter: document.getElementById("weird-review-filter"),
  loadExcerpts: document.getElementById("load-excerpts"),
  loadWeirdExcerpts: document.getElementById("load-weird-excerpts"),
  submitReview: document.getElementById("submit-review"),
  submitReviewBottom: document.getElementById("submit-review-bottom"),
  submitWeirdReview: document.getElementById("submit-weird-review"),
  submitWeirdReviewBottom: document.getElementById("submit-weird-review-bottom"),
  excerptList: document.getElementById("excerpt-list"),
  weirdExcerptList: document.getElementById("weird-excerpt-list"),
  loadCorrectionBooks: document.getElementById("load-correction-books"),
  correctionBookSelect: document.getElementById("correction-book-select"),
  loadCorrections: document.getElementById("load-corrections"),
  submitCorrections: document.getElementById("submit-corrections"),
  submitCorrectionsBottom: document.getElementById("submit-corrections-bottom"),
  autoApplyCorrections: document.getElementById("auto-apply-corrections"),
  correctionList: document.getElementById("correction-list"),
  graphicsMode: document.getElementById("graphics-mode"),
  graphicsFilter: document.getElementById("graphics-filter"),
  graphicsReleaseCatalog: document.getElementById("graphics-release-catalog"),
  graphicsBookSelect: document.getElementById("graphics-book-select"),
  loadGraphicsBooks: document.getElementById("load-graphics-books"),
  loadGraphicsRecords: document.getElementById("load-graphics-records"),
  graphicsFolderImportPanel: document.getElementById("graphics-folder-import-panel"),
  graphicsFolderUrl: document.getElementById("graphics-folder-url"),
  previewGraphicsFolderImport: document.getElementById("preview-graphics-folder-import"),
  applyGraphicsFolderImport: document.getElementById("apply-graphics-folder-import"),
  graphicsFolderImportResults: document.getElementById("graphics-folder-import-results"),
  refreshExcerptHandoffs: document.getElementById("refresh-excerpt-handoffs"),
  retryFailedExcerptHandoffs: document.getElementById("retry-failed-excerpt-handoffs"),
  excerptHandoffSummary: document.getElementById("excerpt-handoff-summary"),
  exportGraphicsSheet: document.getElementById("export-graphics-sheet"),
  submitGraphicsQc: document.getElementById("submit-graphics-qc"),
  submitGraphicsQcBottom: document.getElementById("submit-graphics-qc-bottom"),
  graphicsList: document.getElementById("graphics-list"),
  gatheringMode: document.getElementById("gathering-mode"),
  gatheringEmail: document.getElementById("gathering-email"),
  gatheringEmailWarning: document.getElementById("gathering-email-warning"),
  gatheringBookFields: document.getElementById("gathering-book-fields"),
  gatheringBookAuthor: document.getElementById("gathering-book-author"),
  gatheringBookTitle: document.getElementById("gathering-book-title"),
  gatheringBookBook: document.getElementById("gathering-book-book"),
  gatheringBookQuote: document.getElementById("gathering-book-quote"),
  gatheringBookBold: document.getElementById("gathering-book-bold"),
  gatheringBookItalicize: document.getElementById("gathering-book-italicize"),
  gatheringBookQuoteMeta: document.getElementById("gathering-book-quote-meta"),
  gatheringBookSourceHint: document.getElementById("gathering-book-source-hint"),
  gatheringPrevCatalogPoem: document.getElementById("gathering-prev-catalog-poem"),
  gatheringNextCatalogPoem: document.getElementById("gathering-next-catalog-poem"),
  gatheringViewCatalogPoem: document.getElementById("gathering-view-catalog-poem"),
  gatheringBookNotes: document.getElementById("gathering-book-notes"),
  gatheringBookReaction: document.getElementById("gathering-book-reaction"),
  gatheringVideoFields: document.getElementById("gathering-video-fields"),
  gatheringVideoAuthor: document.getElementById("gathering-video-author"),
  gatheringVideoTitle: document.getElementById("gathering-video-title"),
  gatheringVideoBook: document.getElementById("gathering-video-book"),
  gatheringVideoEvent: document.getElementById("gathering-video-event"),
  gatheringVideoQuote: document.getElementById("gathering-video-quote"),
  gatheringVideoQuoteMeta: document.getElementById("gathering-video-quote-meta"),
  gatheringVideoPlaylistUrl: document.getElementById("gathering-video-playlist-url"),
  gatheringVideoLoadPlaylist: document.getElementById("gathering-video-load-playlist"),
  gatheringVideoPrevItem: document.getElementById("gathering-video-prev-item"),
  gatheringVideoNextItem: document.getElementById("gathering-video-next-item"),
  gatheringVideoOpenItem: document.getElementById("gathering-video-open-item"),
  gatheringVideoPlaylistStatus: document.getElementById("gathering-video-playlist-status"),
  gatheringFixFields: document.getElementById("gathering-fix-fields"),
  gatheringBatchFields: document.getElementById("gathering-batch-fields"),
  gatheringFixPart: document.getElementById("gathering-fix-part"),
  gatheringFixAuthor: document.getElementById("gathering-fix-author"),
  gatheringFixIncorrect: document.getElementById("gathering-fix-incorrect"),
  gatheringFixCorrect: document.getElementById("gathering-fix-correct"),
  gatheringBatchCatalog: document.getElementById("gathering-batch-catalog"),
  gatheringBatchBook: document.getElementById("gathering-batch-book"),
  gatheringBatchShortener: document.getElementById("gathering-batch-shortener"),
  gatheringBatchContentType: document.getElementById("gathering-batch-content-type"),
  gatheringBatchDefaultNotes: document.getElementById("gathering-batch-default-notes"),
  gatheringBatchSource: document.getElementById("gathering-batch-source"),
  gatheringBatchBold: document.getElementById("gathering-batch-bold"),
  gatheringBatchItalicize: document.getElementById("gathering-batch-italicize"),
  gatheringBatchPreview: document.getElementById("gathering-batch-preview"),
  gatheringAuthorOptions: document.getElementById("gathering-author-options"),
  gatheringBookOptions: document.getElementById("gathering-book-options"),
  gatheringBookPoemOptions: document.getElementById("gathering-book-poem-options"),
  loadGatheringOptions: document.getElementById("load-gathering-options"),
  submitGathering: document.getElementById("submit-gathering"),
  previewGatheringBatch: document.getElementById("preview-gathering-batch"),
  submitGatheringBatch: document.getElementById("submit-gathering-batch"),
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
  graphicsCountBadge: document.getElementById("graphics-count-badge"),
  excerptHandoffCountBadge: document.getElementById("excerpt-handoff-count-badge")
};

let currentExcerpts = [];
let currentWeirdExcerpts = [];
let currentCorrectionExcerpts = [];
let currentPendingRecords = [];
let currentGraphicsRecords = [];
let currentGraphicsBookSummaries = [];
let currentGraphicsAssetMatches = new Map();
let currentGraphicsCoverage = null;
let currentGraphicsFolderImportPreview = null;
let graphicsBooksLoadSequence = 0;
let graphicsRecordsLoadSequence = 0;
let currentExcerptHandoffRecords = [];
let isSaving = false;
let currentValidationByRecordId = new Map();
let currentModule = "review";
let intakeOptionsLoaded = false;
let currentIntakeCatalogBooks = [];
let currentIntakeCatalogBooksByKey = new Map();
let currentIntakeLegacyBooks = [];
let currentIntakePublishingBooks = [];
let currentIntakePublishingBooksByKey = new Map();
let currentContributorAccess = null;
let currentGatheringBookPoems = [];
let currentGatheringCatalogBook = null;
let currentGatheringBatchRows = [];
let currentGatheringVideoPlaylist = null;
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
let gatheringCatalogPopup = null;
const contributorShellParams = new URLSearchParams(window.location.search);

const REVIEW_SINGLE_BATCH_SIZE = 1;
const REVIEW_MULTI_BATCH_SIZE = 25;
const EXTRA_REVIEW_BATCH_SIZE = 1;
const REVIEW_VIDEOS_BOOK_KEY = "__video_excerpts__";
const GOOGLE_SHEETS_SCOPES = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file";
const INTAKE_MODE_LABELS = {
  book: "Add a quote from a book",
  video: "Add a quote from a video"
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
  final_reject: "Final reject: this graphic should not move forward.",
  replace: "Manual replacement approved in Weaver."
};

function normalizeGraphicsQcDecisionClient(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "approve") return "approve";
  if (normalized === "replace" || normalized === "replace_here") return "replace";
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
  [elements.submitGraphicsQc, elements.submitGraphicsQcBottom].forEach(button => {
    if (button) {
      button.disabled = isBusy;
    }
  });
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
  [elements.submitGraphicsQc, elements.submitGraphicsQcBottom].forEach(button => {
    if (button) {
      button.textContent = label || (isBusy ? "Saving..." : "Save QC Decisions");
    }
  });
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
  return getIntakeBookBaseTitle(text)
    .replace(/[’‘]/g, "'")
    .replace(/[“”]/g, "\"")
    .replace(/[–—]/g, "-")
    .toLowerCase();
}

function cleanSheetWhitespace(text) {
  return (text || "").toString().replace(/\s+/g, " ").trim();
}

function getContributorAccessForEmail(email) {
  const cleanedEmail = String(email || "").trim().toLowerCase();
  const invites = Array.isArray(runtimeConfig.contributorInvites) ? runtimeConfig.contributorInvites : [];
  if (!cleanedEmail) return null;
  return invites.find(invite => String(invite?.email || "").trim().toLowerCase() === cleanedEmail) || null;
}

function isContributorShellRequested() {
  return contributorShellParams.get("contributor") === "1";
}

function getContributorShellPrefillEmail() {
  return cleanSheetWhitespace(contributorShellParams.get("invite") || contributorShellParams.get("email") || "");
}

function setElementForcedHidden(element, shouldHide) {
  if (!element) return;
  element.toggleAttribute("hidden", shouldHide);
  element.style.display = shouldHide ? "none" : "";
  element.setAttribute("aria-hidden", shouldHide ? "true" : "false");
}

function getAllowedContributorBooks(access) {
  return new Set(
    (Array.isArray(access?.allowedBooks) ? access.allowedBooks : [])
      .map(normalizeBookKey)
      .filter(Boolean)
  );
}

function filterContributorBooks(books = [], access = null) {
  const allowedBookKeys = getAllowedContributorBooks(access);
  if (!allowedBookKeys.size) return books.slice();
  return books.filter(book => allowedBookKeys.has(normalizeBookKey(typeof book === "string" ? book : book?.title)));
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
  const isVideoIntake = cleanSheetWhitespace(row[2]).toLowerCase() === "add a quote from a video";
  const rawVideoAuthor = isVideoIntake ? (row[9] || "").toString() : "";
  const rawVideoTitle = isVideoIntake ? (row[11] || "").toString() : "";
  const rawVideoExcerpt = isVideoIntake ? (row[12] || "").toString() : "";
  const rawVideoBookTitle = isVideoIntake ? cleanSheetWhitespace(row[13] || "") : "";
  const excerptText = (row[config.excerpt - 1] || rawVideoExcerpt).toString();
  const cleanedExcerptText = cleanSheetWhitespace(excerptText);
  const excluded = isSheetYes(row[config.exclude - 1]);
  const reviewDecision = getSheetExcerptReviewDecision(row);
  const bookTitle = cleanSheetWhitespace(row[config.bookTitle - 1]) || rawVideoBookTitle;

  if (!bookTitle || !cleanedExcerptText || excluded || !isPendingSheetReview(reviewDecision)) {
    return null;
  }

  return {
    sourceRow: SHEET_SOURCE_CONFIG.startRow + index,
    recordId: (row[config.recordId - 1] || "").toString(),
    intakeMode: isVideoIntake ? "video" : "book",
    intakeLabel: cleanSheetWhitespace(row[2]),
    author: (row[config.author - 1] || rawVideoAuthor).toString(),
    title: (row[config.title - 1] || rawVideoTitle).toString(),
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

function getReleaseCatalogOptions() {
  return Array.isArray(runtimeConfig.releaseCatalogOptions)
    ? runtimeConfig.releaseCatalogOptions.filter(Boolean)
    : [];
}

function getReleaseCatalogByBookKey() {
  const raw = runtimeConfig.releaseCatalogByTitle || {};
  return raw && typeof raw === "object" ? raw : {};
}

function getSelectedReviewReleaseCatalog() {
  return elements.reviewReleaseCatalog?.value || "__all__";
}

function getSelectedGraphicsReleaseCatalog() {
  return elements.graphicsReleaseCatalog?.value || "__all__";
}

function matchesSelectedReleaseCatalog(bookTitle, selectedCatalog = "__all__") {
  if (!selectedCatalog || selectedCatalog === "__all__") {
    return true;
  }
  const catalogByBookKey = getReleaseCatalogByBookKey();
  const currentCatalog = catalogByBookKey[normalizeBookKey(bookTitle)] || "";
  return currentCatalog === selectedCatalog;
}

function populateReleaseCatalogSelect(select, selectedValue = "__all__") {
  if (!select) return;
  const options = getReleaseCatalogOptions();
  select.innerHTML = "";
  const allOption = document.createElement("option");
  allOption.value = "__all__";
  allOption.textContent = "All release catalogs";
  select.appendChild(allOption);
  options.forEach(optionValue => {
    const option = document.createElement("option");
    option.value = optionValue;
    option.textContent = optionValue;
    select.appendChild(option);
  });
  select.value = options.includes(selectedValue) ? selectedValue : "__all__";
}

function syncReleaseCatalogFilterUi() {
  if (elements.reviewReleaseCatalog) {
    elements.reviewReleaseCatalog.disabled = getSelectedReviewFilter() !== "current_titles";
  }
  if (elements.graphicsReleaseCatalog) {
    elements.graphicsReleaseCatalog.disabled = getSelectedGraphicsMode() !== "queue" || getSelectedGraphicsFilter() !== "current_titles";
  }
}

function getVisibleReviewBookSummaries() {
  if (getSelectedReviewFilter() === "videos") {
    const videoCount = getPendingVideoRecords().length;
    return videoCount
      ? [{ key: REVIEW_VIDEOS_BOOK_KEY, title: "Video excerpts", standardCount: videoCount }]
      : [];
  }
  const reviewQueueIncludeSet = getReviewQueueIncludeSet();
  if (getSelectedReviewFilter() !== "current_titles" || !reviewQueueIncludeSet.size) {
    return currentReviewBookSummaries;
  }
  const selectedCatalog = getSelectedReviewReleaseCatalog();
  return currentReviewBookSummaries.filter(book => (
    reviewQueueIncludeSet.has(book.key)
    && matchesSelectedReleaseCatalog(book.title, selectedCatalog)
  ));
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
  if (bookKey === REVIEW_VIDEOS_BOOK_KEY) {
    return getPendingVideoRecords(records);
  }
  return records.filter(record => normalizeBookKey(record.bookTitle) === bookKey);
}

function isVideoReviewRecord(record) {
  const intakeMode = cleanSheetWhitespace(record?.intakeMode).toLowerCase();
  if (intakeMode === "video") {
    return true;
  }
  return cleanSheetWhitespace(record?.intakeLabel).toLowerCase() === INTAKE_MODE_LABELS.video.toLowerCase();
}

function getPendingVideoRecords(records = currentPendingRecords) {
  return records.filter(isVideoReviewRecord);
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

function refreshCorrectionBooksInBackground() {
  requestReviewApi("/api/review/correction-books")
    .then(data => {
      if (!data.ok || !elements.correctionBookSelect) return;
      const previousSelection = elements.correctionBookSelect.value || "";
      elements.correctionBookSelect.innerHTML = "";
      const placeholder = document.createElement("option");
      placeholder.value = "";
      placeholder.textContent = "Choose a correction book";
      elements.correctionBookSelect.appendChild(placeholder);
      (Array.isArray(data.books) ? data.books : []).forEach(book => {
        const option = document.createElement("option");
        option.value = book.title;
        option.textContent = `${book.title} (${book.count})`;
        elements.correctionBookSelect.appendChild(option);
      });
      if ((Array.isArray(data.books) ? data.books : []).some(book => book.title === previousSelection)) {
        elements.correctionBookSelect.value = previousSelection;
      }
      if (elements.correctionBookCountBadge) {
        elements.correctionBookCountBadge.textContent = `${(Array.isArray(data.books) ? data.books : []).length} Books`;
      }
    })
    .catch(() => {
      // Background refresh is best-effort only.
    });
}

function refreshGraphicsBooksInBackground() {
  const mode = getSelectedGraphicsMode();
  requestReviewApi("/api/review/graphics-books", { mode })
    .then(data => {
      if (!data.ok) return;
      currentGraphicsBookSummaries = Array.isArray(data.books) ? data.books : [];
      graphicsBookSummaryByKey = new Map(
        currentGraphicsBookSummaries.map(book => [normalizeBookKey(book.title), book])
      );
      refreshGraphicsBookSelect(true);
      if (elements.graphicsBookCountBadge) {
        elements.graphicsBookCountBadge.textContent = `${getVisibleGraphicsBookSummaries().length} Books`;
      }
    })
    .catch(() => {
      // Background refresh is best-effort only.
    });
}

function refreshExcerptHandoffsInBackground() {
  requestReviewApi("/api/excerpts/handoffs", { status: "queued,failed,sent" })
    .then(data => {
      if (!data.ok) return;
      renderExcerptHandoffSummary(Array.isArray(data.records) ? data.records : []);
    })
    .catch(() => {
      // Background refresh is best-effort only.
    });
}

function refreshToolSummariesInBackground({
  review = false,
  corrections = false,
  graphics = false,
  handoffs = false,
} = {}) {
  if (review) refreshBookCountsInBackground();
  if (corrections) refreshCorrectionBooksInBackground();
  if (graphics) refreshGraphicsBooksInBackground();
  if (handoffs) refreshExcerptHandoffsInBackground();
}

function applyPendingBookData(records, { preserveSelection = false } = {}) {
  const previousWeirdSelection = preserveSelection ? elements.weirdBookSelect?.value || "" : "";

  currentPendingRecords = records;
  const allBookSummaries = summarizePendingBooks(records);
  const releaseCatalogByTitle = { ...(runtimeConfig.releaseCatalogByTitle || {}) };
  const releaseCatalogOptions = new Set(getReleaseCatalogOptions());
  allBookSummaries.forEach(summary => {
    if (!summary.releaseCatalog) return;
    releaseCatalogByTitle[normalizeBookKey(summary.title)] = summary.releaseCatalog;
    releaseCatalogOptions.add(summary.releaseCatalog);
  });
  runtimeConfig.releaseCatalogByTitle = releaseCatalogByTitle;
  runtimeConfig.releaseCatalogOptions = Array.from(releaseCatalogOptions);
  populateReleaseCatalogSelect(elements.reviewReleaseCatalog, getSelectedReviewReleaseCatalog());
  populateReleaseCatalogSelect(elements.graphicsReleaseCatalog, getSelectedGraphicsReleaseCatalog());
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

function applyContributorAccessMode() {
  currentContributorAccess = getContributorAccessForEmail(elements.gatheringEmail?.value || "");
  const isContributorMode = !!currentContributorAccess;
  const isContributorShell = isContributorShellRequested();
  const hasRequestedPlaylist = !!getRequestedVideoPlaylistUrl();
  const hideIrrelevantUi = isContributorMode || isContributorShell;

  setElementForcedHidden(elements.showReviewModule, hideIrrelevantUi);
  setElementForcedHidden(elements.showWeirdModule, hideIrrelevantUi);
  setElementForcedHidden(elements.showCorrectionsModule, hideIrrelevantUi);
  setElementForcedHidden(elements.showGraphicsModule, hideIrrelevantUi);
  setElementForcedHidden(elements.gatheringTabBook, hideIrrelevantUi && hasRequestedPlaylist);
  setElementForcedHidden(elements.gatheringTabVideo, hideIrrelevantUi && !hasRequestedPlaylist);
  setElementForcedHidden(elements.gatheringTabFix, hideIrrelevantUi);
  setElementForcedHidden(elements.gatheringTabBatch, hideIrrelevantUi);

  if (hideIrrelevantUi) {
    setGatheringMode(hasRequestedPlaylist ? "video" : "book");
  }

  if (isContributorMode) {
    if (elements.gatheringBookBook) {
      elements.gatheringBookBook.disabled = true;
      elements.gatheringBookBook.value = currentContributorAccess.allowedBooks?.[0] || "";
    }
    if (elements.gatheringBookSourceHint) {
      elements.gatheringBookSourceHint.textContent = `Contributor mode: this invite is locked to ${currentContributorAccess.allowedBooks.join(", ")}.`;
    }
  } else if (elements.gatheringBookBook) {
    elements.gatheringBookBook.disabled = false;
  }
}

function getSelectedGatheringMode() {
  return elements.gatheringMode?.value || "book";
}

function getRequestedVideoPlaylistUrl() {
  return contributorShellParams.get("playlist")
    || contributorShellParams.get("playlistUrl")
    || contributorShellParams.get("curationForm")
    || "";
}

const GATHERING_GUIDANCE_BY_MODE = {
  book: {
    lead: "<strong>Working rule of thumb:</strong> copy and paste from the manuscript, capture only the strongest lines, and keep moving past maybes.",
    items: [
      "Prefer the best excerpts, not the most excerpts.",
      "Stay under roughly 3 excerpts per poem and 30 per book.",
      "Aim for plenty of excerpts in the 10–25 word range.",
      "You can read poems inside Weaver, so you do not need to open a separate PDF just to check the poem text.",
      "Use <kbd>Cmd/Ctrl</kbd> + <kbd>Enter</kbd> to submit quickly from a text box."
    ]
  },
  video: {
    lead: "This form is used to gather excerpts from video and add them into Button's excerpts database.",
    items: [
      "<strong>Quote Formatting:</strong> Do not use line breaks.",
      "<strong>For Interns and Employees New to This Task:</strong> please read the following carefully and ask any questions before beginning the task.",
      "Please add quotes into this tool directly, one at a time as you review the assigned video.",
      "Do not create a separate document to stage from.",
      "If possible copy from the transcript on the video.",
      "Please select the very best quotes or lines that stand out to you the most. Maybes = NOs.",
      "Please limit your responses to no more than 3 quotes per poem.",
      "Please do your best to ensure that at least 1/3 of your quotes fall roughly between 10 and 25 words. More than 1/3 in this range is AOK."
    ]
  },
  fix: {
    lead: "Use this lane to correct specific fields in existing quote records without creating a fresh intake row.",
    items: [
      "Paste the incorrect text exactly as it appears now.",
      "Paste the corrected text exactly as you want it to read.",
      "Use the smallest correction that solves the problem.",
      "If the issue is broader than one field, leave a clear correction note."
    ]
  },
  batch: {
    lead: "Use batch paste to preview a set of excerpts or poems before importing them into Weaver in one pass.",
    items: [
      "Set the release catalog, book title, shortener, and content type before previewing.",
      "Paste either poem blocks or rows copied from a spreadsheet.",
      "Review the preview carefully before importing.",
      "Anything missing an author or poem text should be cleaned up before import."
    ]
  }
};

function updateGatheringGuidance(mode) {
  if (!elements.gatheringGuidanceLead || !elements.gatheringGuidanceList) return;
  const guidance = GATHERING_GUIDANCE_BY_MODE[mode] || GATHERING_GUIDANCE_BY_MODE.book;
  elements.gatheringGuidanceLead.innerHTML = guidance.lead;
  elements.gatheringGuidanceList.innerHTML = guidance.items
    .map(item => `<li>${item}</li>`)
    .join("");
}

function updateGatheringModeUi() {
  const mode = getSelectedGatheringMode();
  elements.gatheringBookFields?.toggleAttribute("hidden", mode !== "book");
  elements.gatheringVideoFields?.toggleAttribute("hidden", mode !== "video");
  elements.gatheringFixFields?.toggleAttribute("hidden", mode !== "fix");
  elements.gatheringBatchFields?.toggleAttribute("hidden", mode !== "batch");
  elements.gatheringTabBook?.classList.toggle("gathering-tab--active", mode === "book");
  elements.gatheringTabVideo?.classList.toggle("gathering-tab--active", mode === "video");
  elements.gatheringTabBatch?.classList.toggle("gathering-tab--active", mode === "batch");
  elements.gatheringTabBook?.setAttribute("aria-pressed", mode === "book" ? "true" : "false");
  elements.gatheringTabVideo?.setAttribute("aria-pressed", mode === "video" ? "true" : "false");
  elements.gatheringTabBatch?.setAttribute("aria-pressed", mode === "batch" ? "true" : "false");
  elements.submitGathering?.toggleAttribute("hidden", mode === "batch");
  elements.previewGatheringBatch?.toggleAttribute("hidden", mode !== "batch");
  elements.submitGatheringBatch?.toggleAttribute("hidden", mode !== "batch");
  if (elements.gatheringModeBadge) {
    elements.gatheringModeBadge.textContent = (
      mode === "book" ? "Book Excerpts" :
      mode === "video" ? "Video Excerpts" :
      "Batch Paste"
    );
  }
  updateGatheringGuidance(mode);
  updateGatheringVideoPlaylistUi();
  updateGatheringQuoteMeta();
}

function setGatheringVideoPlaylistStatus(message = "") {
  if (!elements.gatheringVideoPlaylistStatus) return;
  elements.gatheringVideoPlaylistStatus.textContent = message || "Load a curation form to prefill the event name, author, and poem title, then advance through the playlist as you submit.";
}

function updateGatheringVideoPlaylistUi() {
  const playlist = currentGatheringVideoPlaylist;
  const hasPlaylist = !!(playlist && Array.isArray(playlist.items) && playlist.items.length);
  const currentItem = hasPlaylist ? playlist.items[playlist.index] : null;
  if (elements.gatheringVideoPrevItem) {
    elements.gatheringVideoPrevItem.disabled = !hasPlaylist || playlist.index <= 0;
  }
  if (elements.gatheringVideoNextItem) {
    elements.gatheringVideoNextItem.disabled = !hasPlaylist || playlist.index >= playlist.items.length - 1;
  }
  if (elements.gatheringVideoOpenItem) {
    elements.gatheringVideoOpenItem.disabled = !cleanSheetWhitespace(currentItem?.videoUrl);
  }
  if (hasPlaylist) {
    const label = `${playlist.eventName || "Playlist"}: item ${playlist.index + 1} of ${playlist.items.length}${currentItem?.author ? ` · ${currentItem.author}` : ""}${currentItem?.poemTitle ? ` · ${currentItem.poemTitle}` : ""}`;
    setGatheringVideoPlaylistStatus(label);
  } else {
    setGatheringVideoPlaylistStatus("");
  }
}

function applyGatheringVideoPlaylistItem(item, { preserveQuote = false } = {}) {
  if (!item) return;
  if (elements.gatheringVideoEvent) {
    elements.gatheringVideoEvent.value = item.eventName || "";
  }
  if (elements.gatheringVideoAuthor) {
    elements.gatheringVideoAuthor.value = item.author || "";
  }
  if (elements.gatheringVideoTitle) {
    elements.gatheringVideoTitle.value = item.poemTitle || "";
  }
  if (!preserveQuote && elements.gatheringVideoQuote) {
    elements.gatheringVideoQuote.value = "";
  }
  updateGatheringVideoPlaylistUi();
  updateGatheringQuoteMeta();
}

function normalizeVideoPlaylistItem(rawItem, eventName) {
  return {
    author: cleanSheetWhitespace(rawItem?.author),
    poemTitle: cleanSheetWhitespace(rawItem?.poemTitle),
    videoUrl: String(rawItem?.videoUrl || "").trim(),
    eventName: cleanSheetWhitespace(rawItem?.eventName || eventName),
    formResponseUrl: String(rawItem?.formResponseUrl || "").trim(),
    scoreEntryId: cleanSheetWhitespace(rawItem?.scoreEntryId),
    notesEntryId: cleanSheetWhitespace(rawItem?.notesEntryId)
  };
}

function getCurrentGatheringVideoPlaylistItem() {
  return currentGatheringVideoPlaylist?.items?.[currentGatheringVideoPlaylist.index] || null;
}

async function loadGatheringVideoPlaylist({ auto = false } = {}) {
  try {
    const formUrl = elements.gatheringVideoPlaylistUrl?.value.trim() || "";
    if (!formUrl) {
      throw new Error("Add a curation form URL first.");
    }
    if (!auto) {
      setStatus("Loading video playlist...");
    }
    const result = await requestReviewApi("/api/intake/video-playlist", { formUrl });
    const items = Array.isArray(result.items)
      ? result.items.map(item => normalizeVideoPlaylistItem(item, result.eventName)).filter(item => item.author || item.poemTitle)
      : [];
    if (!items.length) {
      throw new Error("No playlist items were found in that form.");
    }
    currentGatheringVideoPlaylist = {
      formUrl,
      eventName: cleanSheetWhitespace(result.eventName),
      formTitle: cleanSheetWhitespace(result.formTitle),
      items,
      index: 0
    };
    applyGatheringVideoPlaylistItem(items[0]);
    if (getSelectedGatheringMode() !== "video") {
      setGatheringMode("video");
    }
    setStatus(`Loaded ${items.length} playlist items from "${currentGatheringVideoPlaylist.eventName || currentGatheringVideoPlaylist.formTitle || "the curation form"}".`);
  } catch (error) {
    currentGatheringVideoPlaylist = null;
    updateGatheringVideoPlaylistUi();
    setStatus(`Video playlist load failed: ${error.message}`);
  }
}

function navigateGatheringVideoPlaylist(direction) {
  if (!currentGatheringVideoPlaylist?.items?.length) return;
  const nextIndex = currentGatheringVideoPlaylist.index + direction;
  if (nextIndex < 0 || nextIndex >= currentGatheringVideoPlaylist.items.length) {
    return;
  }
  currentGatheringVideoPlaylist.index = nextIndex;
  applyGatheringVideoPlaylistItem(currentGatheringVideoPlaylist.items[nextIndex]);
}

function openGatheringVideoPlaylistItem() {
  const item = currentGatheringVideoPlaylist?.items?.[currentGatheringVideoPlaylist.index];
  const videoUrl = item?.videoUrl || "";
  if (!videoUrl) return;
  window.open(videoUrl, "_blank", "noopener,noreferrer");
}

function setGatheringMode(mode) {
  if (!elements.gatheringMode) return;
  elements.gatheringMode.value = mode;
  updateGatheringModeUi();
}

function setActiveModule(moduleName) {
  if ((currentContributorAccess || isContributorShellRequested()) && moduleName !== "gathering") {
    moduleName = "gathering";
  }
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
  elements.showGraphicsModule?.classList.toggle("hero-pill--active", currentModule === "graphics" && currentGraphicsView === "main");
  elements.showGraphicsOps?.classList.toggle("hero-ops-button--active", currentModule === "graphics" && currentGraphicsView === "ops");
}

async function requestReviewApi(path, params = {}) {
  const url = new URL(path, window.location.origin);
  Object.entries(params).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== "") {
      url.searchParams.set(key, String(value));
    }
  });

  const response = await fetch(url, {
    cache: "no-store",
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
  currentIntakeLegacyBooks = filterContributorBooks(currentIntakeLegacyBooks, currentContributorAccess);
  currentIntakeCatalogBooks = filterContributorBooks(currentIntakeCatalogBooks, currentContributorAccess);
  currentIntakeCatalogBooksByKey = new Map(
    currentIntakeCatalogBooks.map(book => [normalizeBookKey(book.title), book])
  );
  currentIntakePublishingBooks = Array.isArray(data.publishingBooks) ? data.publishingBooks : [];
  currentIntakePublishingBooks = filterContributorBooks(currentIntakePublishingBooks, currentContributorAccess);
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
  applyContributorAccessMode();
  const publishingStatusNote = data.publishingBooksError
    ? " Publishing-order titles are temporarily unavailable until that sheet is shared with Weaver."
    : "";
  setStatus(`Loaded ${data.authors?.length || 0} legacy author suggestions, ${currentIntakeLegacyBooks.length} legacy book suggestions, ${currentIntakePublishingBooks.length} publishing-order books, and ${currentIntakeCatalogBooks.length} catalog-backed books for intake.${publishingStatusNote}`);
}

async function loadGatheringPoemsForBook(bookTitle, { preserveTitle = false } = {}) {
  const cleanedBookTitle = (bookTitle || "").trim();
  if (!cleanedBookTitle || !elements.gatheringBookTitle) {
    populateDatalist(elements.gatheringBookPoemOptions, []);
    currentGatheringBookPoems = [];
    currentGatheringCatalogBook = null;
    if (elements.gatheringBookSourceHint) {
      elements.gatheringBookSourceHint.textContent = "Select a book to load poem titles from a catalog-backed source, or keep going manually when no catalog source is available.";
    }
    updateGatheringCatalogPreviewState();
    return;
  }

  const previousValue = preserveTitle ? elements.gatheringBookTitle.value.trim() : "";
  const data = await requestReviewApi("/api/intake/catalog/poems", { bookTitle: cleanedBookTitle });
  const poems = Array.isArray(data.poems) ? data.poems : [];
  currentGatheringCatalogBook = {
    title: data.bookTitle || cleanedBookTitle,
    author: data.author || "",
    primarySourceFormat: data.primarySourceFormat || ""
  };
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
    elements.gatheringBookSourceHint.textContent = `${data.primarySourceFormat || "Catalog"} source connected. ${poems.length} poem titles loaded for ${data.bookTitle || cleanedBookTitle}. You can read the poem in Weaver and also type a manual poem title if needed.`;
  }
  updateGatheringCatalogPreviewState();
}

function updateGatheringCatalogPreviewState() {
  const bookTitle = elements.gatheringBookBook?.value.trim() || "";
  const poemTitle = elements.gatheringBookTitle?.value.trim() || "";
  const hasCatalogBook = !!currentGatheringCatalogBook || !!currentIntakeCatalogBooksByKey.get(normalizeBookKey(bookTitle));
  const needsPoemSelection = hasCatalogBook && bookTitle && !poemTitle;
  const currentIndex = currentGatheringBookPoems.indexOf(poemTitle);
  const hasIndexedPoem = currentIndex >= 0;

  if (elements.gatheringViewCatalogPoem) {
    elements.gatheringViewCatalogPoem.disabled = !(hasCatalogBook && bookTitle);
    elements.gatheringViewCatalogPoem.textContent = needsPoemSelection ? "Select a poem first" : (hasCatalogBook ? "Read the Poem" : "Poem View Unavailable");
    elements.gatheringViewCatalogPoem.title = needsPoemSelection ? "Choose a poem title before reading the poem." : "";
  }
  if (elements.gatheringBookTitle) {
    elements.gatheringBookTitle.classList.toggle("field-error", needsPoemSelection);
    elements.gatheringBookTitle.title = needsPoemSelection ? "Select a poem title before reading the poem." : "";
  }
  if (elements.gatheringBookSourceHint && hasCatalogBook && bookTitle) {
    elements.gatheringBookSourceHint.classList.toggle("gathering-source-hint--warning", needsPoemSelection);
    elements.gatheringBookSourceHint.textContent = needsPoemSelection
      ? "Select a poem title next, then use Read the Poem to open the catalog context."
      : `${currentGatheringCatalogBook?.primarySourceFormat || "Catalog"} source connected. ${currentGatheringBookPoems.length} poem titles loaded for ${currentGatheringCatalogBook?.title || bookTitle}. You can read the poem in Weaver and also type a manual poem title if needed.`;
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
  currentGatheringCatalogBook = null;
  if (elements.gatheringBookTitle) {
    elements.gatheringBookTitle.placeholder = "Type poem title manually";
  }
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
      elements.gatheringBookSourceHint.textContent = `${cleanedBookTitle} is coming from the publishing-order sheet (${catalogLabel}) and does not have a readable catalog source in Weaver yet. Author is prefilled when available; enter the poem title manually.`;
    } else if (cleanedBookTitle && isLegacySuggested) {
      elements.gatheringBookSourceHint.textContent = `${cleanedBookTitle} is available from the older intake suggestions, but not from a readable catalog-backed source yet. Enter the poem title and author manually.`;
    } else if (cleanedBookTitle) {
      elements.gatheringBookSourceHint.textContent = `${cleanedBookTitle} is not catalog-backed in Weaver yet. You can still enter the author, poem title, and quote manually.`;
    } else {
      elements.gatheringBookSourceHint.textContent = "Select a book to load poem titles from a catalog-backed source, or keep going manually when no catalog source is available.";
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
  if (gatheringCatalogPopup && !gatheringCatalogPopup.closed) {
    openGatheringCatalogPoem();
  }
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
  const lookupTitle = catalogBook?.title || bookTitle;
  if (elements.gatheringBookTitle) {
    elements.gatheringBookTitle.placeholder = "Start typing to filter catalog poem titles";
  }
  try {
    await loadGatheringPoemsForBook(lookupTitle, { preserveTitle });
    if (currentGatheringBookPoems.length) {
      return;
    }
  } catch (_error) {
    // Fall through to manual mode if the live catalog lookup fails.
  }
  if (!preserveTitle && elements.gatheringBookTitle) elements.gatheringBookTitle.value = "";
  setGatheringBookManualMode(bookTitle);
}

function resetGatheringForm() {
  [
    elements.gatheringBookAuthor,
    elements.gatheringBookBook,
    elements.gatheringBookQuote,
    elements.gatheringBookNotes,
    elements.gatheringVideoAuthor,
    elements.gatheringVideoTitle,
    elements.gatheringVideoBook,
    elements.gatheringVideoEvent,
    elements.gatheringVideoScore,
    elements.gatheringVideoScoreNotes,
    elements.gatheringVideoQuote,
    elements.gatheringVideoPlaylistUrl,
    elements.gatheringFixPart,
    elements.gatheringFixAuthor,
    elements.gatheringFixIncorrect,
    elements.gatheringFixCorrect,
    elements.gatheringBatchCatalog,
    elements.gatheringBatchBook,
    elements.gatheringBatchShortener,
    elements.gatheringBatchDefaultNotes,
    elements.gatheringBatchSource
  ].forEach(field => {
    if (field) field.value = "";
  });
  currentGatheringBatchRows = [];
  if (elements.gatheringBatchPreview) {
    elements.gatheringBatchPreview.innerHTML = "";
  }
  populateDatalist(elements.gatheringBookPoemOptions, []);
  setGatheringBookManualMode("");
  updateGatheringCatalogPreviewState();
  updateGatheringQuoteMeta();
}

function resetGatheringAfterSubmit(mode) {
  if (mode === "book") {
    if (elements.gatheringBookQuote) elements.gatheringBookQuote.value = "";
    if (elements.gatheringBookNotes) elements.gatheringBookNotes.value = "";
    if (elements.gatheringBookReaction) elements.gatheringBookReaction.value = "";
    updateGatheringCatalogPreviewState();
    updateGatheringQuoteMeta();
    elements.gatheringBookQuote?.focus();
    return;
  }

  if (mode === "video") {
    if (elements.gatheringVideoQuote) elements.gatheringVideoQuote.value = "";
    if (elements.gatheringVideoScore) elements.gatheringVideoScore.value = "";
    if (elements.gatheringVideoScoreNotes) elements.gatheringVideoScoreNotes.value = "";
    if (currentGatheringVideoPlaylist?.items?.length) {
      if (currentGatheringVideoPlaylist.index < currentGatheringVideoPlaylist.items.length - 1) {
        currentGatheringVideoPlaylist.index += 1;
        applyGatheringVideoPlaylistItem(
          currentGatheringVideoPlaylist.items[currentGatheringVideoPlaylist.index],
          { preserveQuote: false }
        );
      } else {
        updateGatheringVideoPlaylistUi();
      }
    }
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

  if (mode === "batch") {
    if (elements.gatheringBatchSource) elements.gatheringBatchSource.value = "";
    if (elements.gatheringBatchDefaultNotes) elements.gatheringBatchDefaultNotes.value = "";
    currentGatheringBatchRows = [];
    if (elements.gatheringBatchPreview) {
      elements.gatheringBatchPreview.innerHTML = "";
    }
    elements.gatheringBatchSource?.focus();
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

function setGatheringEmailWarning(message = "") {
  if (!elements.gatheringEmail || !elements.gatheringEmailWarning) return;
  const hasMessage = !!String(message || "").trim();
  elements.gatheringEmailWarning.hidden = !hasMessage;
  elements.gatheringEmailWarning.textContent = hasMessage ? String(message).trim() : "";
  elements.gatheringEmail.classList.toggle("field-error", hasMessage);
}

function handleGatheringQuickSubmit(event) {
  if ((event.metaKey || event.ctrlKey) && event.key === "Enter") {
    event.preventDefault();
    if (getSelectedGatheringMode() === "batch") {
      submitGatheringBatch();
      return;
    }
    submitGathering();
  }
}

function wrapTextareaSelectionWithItalics(textarea) {
  if (!(textarea instanceof HTMLTextAreaElement)) return;
  const start = textarea.selectionStart ?? 0;
  const end = textarea.selectionEnd ?? 0;
  const value = textarea.value || "";
  const selected = value.slice(start, end);
  const replacement = `*${selected || "italic text"}*`;
  textarea.setRangeText(replacement, start, end, "end");
  if (!selected) {
    textarea.setSelectionRange(start + 1, start + replacement.length - 1);
  }
  textarea.focus();
  updateGatheringQuoteMeta();
}

function wrapTextareaSelectionWithBold(textarea) {
  if (!(textarea instanceof HTMLTextAreaElement)) return;
  const start = textarea.selectionStart ?? 0;
  const end = textarea.selectionEnd ?? 0;
  const value = textarea.value || "";
  const selected = value.slice(start, end);
  const replacement = `**${selected || "bold text"}**`;
  textarea.setRangeText(replacement, start, end, "end");
  if (!selected) {
    textarea.setSelectionRange(start + 2, start + replacement.length - 2);
  }
  textarea.focus();
  updateGatheringQuoteMeta();
}

function splitGatheringBatchBlocks(text) {
  const normalized = String(text || "").replace(/\r\n/g, "\n").trim();
  if (!normalized) return [];
  return normalized
    .split(/\n\s*---+\s*\n|\n{2,}/)
    .map(block => block.trim())
    .filter(Boolean);
}

function parseDelimitedGatheringRows(text) {
  const source = String(text || "").replace(/\r\n/g, "\n");
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    const nextChar = source[index + 1];

    if (char === "\"") {
      if (inQuotes && nextChar === "\"") {
        cell += "\"";
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (!inQuotes && char === "\t") {
      row.push(cell);
      cell = "";
      continue;
    }

    if (!inQuotes && char === "\n") {
      row.push(cell);
      if (row.some(value => String(value || "").trim())) {
        rows.push(row);
      }
      row = [];
      cell = "";
      continue;
    }

    cell += char;
  }

  if (cell.length || row.length) {
    row.push(cell);
    if (row.some(value => String(value || "").trim())) {
      rows.push(row);
    }
  }

  return rows;
}

function normalizeBatchHeader(value) {
  return cleanSheetWhitespace(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function buildGatheringBatchNotes(defaults, igHandle) {
  return [
    defaults.catalog ? `Release Catalog: ${defaults.catalog}` : "",
    defaults.bookShortener ? `Book Shortener: ${defaults.bookShortener}` : "",
    defaults.contentType ? `Content Type: ${defaults.contentType}` : "",
    igHandle ? `Social media handle: ${igHandle}` : "",
    defaults.defaultNotes || ""
  ].filter(Boolean).join("\n\n");
}

function deriveGatheringBatchTitle(submission) {
  const lines = String(submission || "")
    .replace(/\r\n/g, "\n")
    .split("\n")
    .map(line => line.trim())
    .filter(Boolean);
  return (lines[0] || "").slice(0, 120);
}

function parseGatheringBatchSpreadsheetRows(text, defaults) {
  const rows = parseDelimitedGatheringRows(text);
  if (!rows.length) {
    return null;
  }

  const headers = rows[0].map(normalizeBatchHeader);
  const submissionIndex = headers.findIndex(header => header.includes("submission"));
  const firstNameIndex = headers.findIndex(header => header === "first name");
  const lastNameIndex = headers.findIndex(header => header === "last name");
  const handleIndex = headers.findIndex(header => header.includes("instagram handle"));

  let dataRows = rows.slice(1);
  let firstNameColumn = firstNameIndex;
  let lastNameColumn = lastNameIndex;
  let handleColumn = handleIndex;
  let submissionColumn = submissionIndex;

  if (submissionIndex === -1 || firstNameIndex === -1) {
    const firstDataRow = rows[0] || [];
    const looksLikeRawFormRow = firstDataRow.length >= 10;
    if (!looksLikeRawFormRow) {
      return null;
    }
    dataRows = rows;
    handleColumn = 1;
    firstNameColumn = 3;
    lastNameColumn = 4;
    submissionColumn = 9;
  }

  return dataRows.map((cells, index) => {
    const firstName = String(cells[firstNameColumn] || "").trim();
    const lastName = String(cells[lastNameColumn] || "").trim();
    const author = [firstName, lastName].filter(Boolean).join(" ").trim();
    const igHandle = String(cells[handleColumn] || "").trim();
    const quote = String(cells[submissionColumn] || "").trim();
    const title = deriveGatheringBatchTitle(quote);

    return {
      key: `sheet-batch-${index + 1}`,
      index: index + 1,
      title,
      author,
      quote,
      igHandle,
      bookTitle: defaults.bookTitle,
      notes: buildGatheringBatchNotes(defaults, igHandle),
      error: !author || !quote ? "Needs an author and poem text before import." : ""
    };
  }).filter(row => row.author || row.quote || row.igHandle);
}

function parseGatheringBatchBlock(block, index, defaults) {
  const lines = String(block || "")
    .replace(/\r\n/g, "\n")
    .split("\n")
    .map(line => line.trim())
    .filter(Boolean);
  const meta = { title: "", author: "", igHandle: "" };
  const bodyLines = [];

  lines.forEach(line => {
    const titleMatch = line.match(/^title\s*:\s*(.+)$/i);
    const authorMatch = line.match(/^author\s*:\s*(.+)$/i);
    const igMatch = line.match(/^(ig|instagram|handle)\s*:\s*(.+)$/i);
    if (titleMatch) {
      meta.title = titleMatch[1].trim();
      return;
    }
    if (authorMatch) {
      meta.author = authorMatch[1].trim();
      return;
    }
    if (igMatch) {
      meta.igHandle = igMatch[2].trim();
      return;
    }
    bodyLines.push(line);
  });

  if (!meta.igHandle && bodyLines.length && /^@/.test(bodyLines[bodyLines.length - 1])) {
    meta.igHandle = bodyLines.pop().trim();
  }

  if (!meta.author && bodyLines.length > 1) {
    const maybeAuthor = bodyLines[bodyLines.length - 1];
    if (maybeAuthor.length <= 80 && !/[.?!,;:]$/.test(maybeAuthor)) {
      meta.author = bodyLines.pop().trim();
    }
  }

  if (!meta.title && bodyLines.length) {
    meta.title = bodyLines[0].trim().slice(0, 120);
  }

  const quote = bodyLines.join("\n").trim();

  return {
    key: `batch-${index + 1}`,
    index: index + 1,
    title: meta.title,
    author: meta.author,
    quote,
    igHandle: meta.igHandle,
    bookTitle: defaults.bookTitle,
    notes: buildGatheringBatchNotes(defaults, meta.igHandle),
    error: !meta.author || !quote ? "Needs an author and poem text before import." : ""
  };
}

function renderGatheringBatchPreview(rows) {
  if (!elements.gatheringBatchPreview) return;
  if (!rows.length) {
    elements.gatheringBatchPreview.innerHTML = "";
    return;
  }

  elements.gatheringBatchPreview.innerHTML = rows.map(row => `
    <article class="excerpt-card">
      <div class="excerpt-card__meta">
        <span class="pill pill--muted">Item ${row.index}</span>
        ${row.title ? `<span class="pill pill--muted">${escapeHtml(row.title)}</span>` : ""}
        ${row.author ? `<span class="pill pill--muted">${escapeHtml(row.author)}</span>` : ""}
        ${row.igHandle ? `<span class="pill pill--muted">${escapeHtml(row.igHandle)}</span>` : ""}
      </div>
      ${row.error ? `<p class="validation validation--warn">${escapeHtml(row.error)}</p>` : `<p class="validation validation--good">Ready to import into ${escapeHtml(row.bookTitle || "the selected book")}.</p>`}
      <blockquote class="excerpt-card__quote">${escapeHtml(row.quote || "(No poem text parsed yet.)")}</blockquote>
    </article>
  `).join("");
}

function buildGatheringBatchRows() {
  const catalog = elements.gatheringBatchCatalog?.value.trim() || "";
  const bookTitle = elements.gatheringBatchBook?.value.trim() || "";
  const bookShortener = elements.gatheringBatchShortener?.value.trim() || "";
  const contentType = elements.gatheringBatchContentType?.value || "EXC";
  const source = elements.gatheringBatchSource?.value || "";
  const defaultNotes = elements.gatheringBatchDefaultNotes?.value.trim() || "";
  const defaults = { catalog, bookTitle, bookShortener, contentType, defaultNotes };

  if (!bookTitle || !source.trim()) {
    throw new Error("Batch paste needs a book title and pasted content.");
  }
  if (!catalog || !bookShortener) {
    throw new Error("Batch paste needs both a release catalog and a shortener.");
  }

  const spreadsheetRows = parseGatheringBatchSpreadsheetRows(source, defaults);
  const rows = (spreadsheetRows && spreadsheetRows.length)
    ? spreadsheetRows
    : splitGatheringBatchBlocks(source).map((block, index) => (
      parseGatheringBatchBlock(block, index, defaults)
    ));

  if (!rows.length) {
    throw new Error("No poem blocks were found in the pasted content.");
  }

  return rows;
}

function previewGatheringBatch() {
  try {
    currentGatheringBatchRows = buildGatheringBatchRows();
    renderGatheringBatchPreview(currentGatheringBatchRows);
    const validCount = currentGatheringBatchRows.filter(row => !row.error).length;
    const invalidCount = currentGatheringBatchRows.length - validCount;
    setStatus(
      invalidCount
        ? `Previewed ${currentGatheringBatchRows.length} items. ${validCount} are ready and ${invalidCount} need cleanup.`
        : `Previewed ${currentGatheringBatchRows.length} items. Everything is ready to import.`
    );
  } catch (error) {
    currentGatheringBatchRows = [];
    renderGatheringBatchPreview([]);
    setStatus(`Batch preview failed: ${error.message}`);
  }
}

async function submitGatheringBatch() {
  try {
    const email = elements.gatheringEmail?.value.trim() || "";
    if (!email) {
      setGatheringEmailWarning("Add your email before importing a batch.");
      elements.gatheringEmail?.focus();
      throw new Error("Add an email before importing a batch.");
    }
    setGatheringEmailWarning("");

    if (!currentGatheringBatchRows.length) {
      currentGatheringBatchRows = buildGatheringBatchRows();
      renderGatheringBatchPreview(currentGatheringBatchRows);
    }

    const invalidRow = currentGatheringBatchRows.find(row => row.error);
    if (invalidRow) {
      throw new Error(`Item ${invalidRow.index} needs cleanup before import.`);
    }

    setSubmitState(true, "Importing...");
    let savedCount = 0;
    const batchBookTitle = elements.gatheringBatchBook?.value.trim() || "the selected book";
    for (const row of currentGatheringBatchRows) {
      await postReviewApi("/api/intake/submit", {
        mode: "book",
        email,
        author: row.author,
        title: row.title,
        quote: row.quote,
        bookTitle: row.bookTitle,
        notes: row.notes
      });
      savedCount += 1;
      setStatus(`Imported ${savedCount} of ${currentGatheringBatchRows.length} pasted excerpts...`);
    }

    resetGatheringAfterSubmit("batch");
    setStatus(`Imported ${savedCount} pasted excerpts into "${batchBookTitle}".`);
  } catch (error) {
    setStatus(`Batch import failed: ${error.message}`);
  } finally {
    setSubmitState(false, "Submit Excerpt");
  }
}

function buildGatheringPayload() {
  const mode = getSelectedGatheringMode();
  const email = elements.gatheringEmail?.value.trim() || "";
  if (!email) {
    setGatheringEmailWarning("Add your email before submitting.");
    elements.gatheringEmail?.focus();
    throw new Error("Add an email before submitting an excerpt gathering row.");
  }
  setGatheringEmailWarning("");

  if (mode === "book") {
    const author = elements.gatheringBookAuthor?.value.trim() || "";
    const title = elements.gatheringBookTitle?.value.trim() || "";
    const quote = elements.gatheringBookQuote?.value.trim() || "";
    const bookTitle = elements.gatheringBookBook?.value.trim() || "";
    const notes = elements.gatheringBookNotes?.value.trim() || "";
    const reaction = elements.gatheringBookReaction?.value.trim() || "";
    if (!author || !title || !quote) {
      throw new Error("Book intake needs an author, poem title, and quote.");
    }
    const combinedNotes = [
      notes,
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
    if (payload.mode === "book") {
      const shouldContinue = await confirmGatheringBookQuoteAgainstCatalog(payload);
      if (!shouldContinue) {
        setStatus("Excerpt submission canceled so you can review the poem text.");
        return;
      }
    }
    setStatus(`Submitting ${INTAKE_MODE_LABELS[payload.mode] || "excerpt gathering"} row...`);
    const result = await postReviewApi("/api/intake/submit", payload);
    resetGatheringAfterSubmit(payload.mode);
    setGatheringEmailWarning("");
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
    setStatus("Choose a catalog-backed book and poem before opening catalog context.");
    elements.gatheringBookTitle?.focus();
    updateGatheringCatalogPreviewState();
    return;
  }
  const url = new URL("/catalog-poem", window.location.origin);
  url.searchParams.set("bookTitle", bookTitle);
  url.searchParams.set("poemTitle", poemTitle);
  if (excerptText) {
    url.searchParams.set("excerptText", excerptText);
  }
  gatheringCatalogPopup = openComparisonWindow(url.toString(), "weaverGatheringCatalogPoem");
}

async function confirmGatheringBookQuoteAgainstCatalog(payload) {
  const normalizedBookKey = normalizeBookKey(payload.bookTitle || "");
  const hasCatalogBook = !!currentGatheringCatalogBook || !!currentIntakeCatalogBooksByKey.get(normalizedBookKey);
  if (!hasCatalogBook || !payload.bookTitle || !payload.title || !payload.quote) {
    return true;
  }

  try {
    const response = await fetch("/api/catalog/validate", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        records: [{
          recordId: "gathering-preview",
          sourceRow: 0,
          author: payload.author,
          title: payload.title,
          bookTitle: payload.bookTitle,
          excerptText: payload.quote
        }]
      })
    });
    const data = await response.json();
    if (!data.ok) return true;
    const validation = Array.isArray(data.results) ? data.results[0] : null;
    if (!validation) return true;

    const safeStatuses = new Set([
      "catalog_match",
      "formatting_only",
      "author_match_only"
    ]);
    if (safeStatuses.has(validation.status)) {
      return true;
    }

    return window.confirm("This text doesn't appear to be in the poem. Are you sure?\n\nYes = submit anyway\nNo = go back");
  } catch (_error) {
    return true;
  }
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
    return popup;
  } else {
    window.open(href, "_blank", "noopener,noreferrer");
    return null;
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
  const fallbackUrl = assetMatch.linkUrl || assetMatch.lowResLinkUrl || "";
  return normalizeGraphicsPreviewUrl(directUrl, fallbackUrl);
}

function normalizeGraphicsPreviewUrl(url, fallbackUrl = "") {
  const directUrl = String(url || "").trim();
  const fallbackDirectUrl = String(fallbackUrl || "").trim();
  if (!directUrl && !fallbackDirectUrl) return "";

  const driveFileId = extractGoogleDriveFileId(directUrl) || extractGoogleDriveFileId(fallbackDirectUrl);
  if (driveFileId) {
    return `/api/drive-image?fileId=${encodeURIComponent(driveFileId)}`;
  }

  return directUrl || fallbackDirectUrl;
}

function buildReviewSavePayload(update) {
  return {
    sourceRow: update.sourceRow,
    recordId: update.recordId,
    author: update.author,
    bookTitle: update.bookTitle,
    poemTitle: update.title,
    excerptText: update.excerptText,
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
    author: update.author,
    bookTitle: update.bookTitle,
    poemTitle: update.title,
    excerptText: update.excerptText,
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
        releaseCatalog: cleanSheetWhitespace(record.releaseCatalog || ""),
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
    summary.releaseCatalog = summary.releaseCatalog || cleanSheetWhitespace(record.releaseCatalog || "");
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
  return elements.graphicsMode?.value || "qc";
}

function getSelectedGraphicsFilter() {
  return elements.graphicsFilter?.value || "current_titles";
}

function getGraphicsModeLabel(mode = getSelectedGraphicsMode()) {
  if (mode === "qc") return "graphics QC";
  if (mode === "cleanup") return "under the hood";
  if (mode === "mismatch") return "mismatch pairing";
  if (mode === "handoff") return "Poetry Please handoff";
  if (mode === "handoff_retry") return "retry failed handoffs";
  if (mode === "rework") return "send back to P.I.G.";
  if (mode === "coverage_needs") return "coverage needs";
  if (mode === "coverage") return "coverage";
  if (mode === "queue") return "needs graphics";
  return "graphics creation";
}

function syncGraphicsModeOptionsForView() {
  if (!elements.graphicsMode) return;
  const isOpsView = currentGraphicsView === "ops";
  let reworkOption = elements.graphicsMode.querySelector('option[value="rework"]');
  if (!reworkOption) {
    reworkOption = document.createElement("option");
    reworkOption.value = "rework";
    reworkOption.hidden = true;
    elements.graphicsMode.appendChild(reworkOption);
  }
  let handoffRetryOption = elements.graphicsMode.querySelector('option[value="handoff_retry"]');
  if (!handoffRetryOption) {
    handoffRetryOption = document.createElement("option");
    handoffRetryOption.value = "handoff_retry";
    handoffRetryOption.hidden = true;
    elements.graphicsMode.appendChild(handoffRetryOption);
  }
  const queueOption = elements.graphicsMode.querySelector('option[value="queue"]');
  const qcOption = elements.graphicsMode.querySelector('option[value="qc"]');
  const mismatchOption = elements.graphicsMode.querySelector('option[value="mismatch"]');
  const handoffOption = elements.graphicsMode.querySelector('option[value="handoff"]');
  const coverageNeedsOption = elements.graphicsMode.querySelector('option[value="coverage_needs"]');
  const coverageOption = elements.graphicsMode.querySelector('option[value="coverage"]');

  if (queueOption) {
    queueOption.hidden = !isOpsView;
    queueOption.textContent = "Needs graphics";
  }
  if (handoffOption) {
    handoffOption.hidden = !isOpsView;
    handoffOption.textContent = "Poetry Please handoff";
  }
  if (handoffRetryOption) {
    handoffRetryOption.hidden = !isOpsView;
    handoffRetryOption.textContent = "Retry failed handoffs";
  }
  if (reworkOption) {
    reworkOption.hidden = !isOpsView;
    reworkOption.textContent = "Send back to P.I.G.";
  }
  if (coverageNeedsOption) {
    coverageNeedsOption.hidden = !isOpsView;
    coverageNeedsOption.textContent = "Coverage needs";
  }
  if (coverageOption) {
    coverageOption.hidden = !isOpsView;
    coverageOption.textContent = "Coverage";
  }
  if (qcOption) qcOption.hidden = isOpsView;
  if (mismatchOption) mismatchOption.hidden = isOpsView;

  if (isOpsView && !["queue", "handoff", "handoff_retry", "coverage_needs", "coverage", "rework"].includes(elements.graphicsMode.value)) {
    elements.graphicsMode.value = "queue";
  }
  if (!isOpsView && !["qc", "mismatch"].includes(elements.graphicsMode.value)) {
    elements.graphicsMode.value = "qc";
  }
}

function updateGraphicsModuleTitle() {
  if (!elements.graphicsModuleTitle) return;
  const isOpsView = currentGraphicsView === "ops";
  const isOpsHandoff = isOpsView && getSelectedGraphicsMode() === "handoff";
  elements.graphicsModuleTitle.textContent = isOpsView
    ? "Under the Hood"
    : "Graphics QC";
  setElementForcedHidden(elements.graphicsOpsPanel, !isOpsHandoff);
  setElementForcedHidden(elements.graphicsModeField, !isOpsView);
  setElementForcedHidden(elements.graphicsFilterField, isOpsView);
  setElementForcedHidden(elements.graphicsReleaseCatalogField, isOpsView);
  setElementForcedHidden(elements.graphicsMainHint, isOpsView);
  setElementForcedHidden(elements.graphicsFolderImportPanel, isOpsView);
  setElementForcedHidden(elements.graphicsBookPicker, isOpsHandoff);
  setElementForcedHidden(elements.graphicsQueueRowsHeader, isOpsHandoff);
  setElementForcedHidden(elements.graphicsQueueToolbarTop, isOpsHandoff);
  setElementForcedHidden(elements.graphicsList, isOpsHandoff);
  setElementForcedHidden(elements.graphicsQueueToolbarBottom, isOpsHandoff);
}

function getVisibleGraphicsBookSummaries() {
  if (getSelectedGraphicsMode() !== "queue" || getSelectedGraphicsFilter() !== "current_titles") {
    return currentGraphicsBookSummaries;
  }

  const reviewQueueIncludeSet = getReviewQueueIncludeSet();
  if (!reviewQueueIncludeSet.size) {
    return currentGraphicsBookSummaries;
  }
  const selectedCatalog = getSelectedGraphicsReleaseCatalog();
  return currentGraphicsBookSummaries.filter(book => (
    reviewQueueIncludeSet.has(normalizeBookKey(book.title))
    && matchesSelectedReleaseCatalog(book.title, selectedCatalog)
  ));
}

function isGraphicsQcSweepSelection(value = elements.graphicsBookSelect?.value || "") {
  return value === "__qc_sweep__";
}

function refreshGraphicsBookSelect(preserveSelection = true) {
  const previousSelection = preserveSelection ? elements.graphicsBookSelect?.value || "" : "";
  let visibleBooks = getVisibleGraphicsBookSummaries();

  if (getSelectedGraphicsMode() === "qc") {
    const totalCount = visibleBooks.reduce((sum, book) => sum + Number(book.count || 0), 0);
    visibleBooks = [
      { title: "QC Sweep", count: totalCount, key: "__qc_sweep__" },
      ...visibleBooks
    ];
  }

  populateBookSelect(
    elements.graphicsBookSelect,
    visibleBooks.map(book => ({ ...book, key: book.key || normalizeBookKey(book.title) })),
    previousSelection,
    book => book.key === "__qc_sweep__" ? `${book.title} (${book.count} remaining)` : `${book.title} (${book.count})`
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

function wait(ms) {
  return new Promise(resolve => window.setTimeout(resolve, ms));
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
  const loadSequence = ++graphicsBooksLoadSequence;

  try {
    setStatus(`Loading ${getGraphicsModeLabel(mode)} books...`);
    let data;
    try {
      data = mode === "coverage_needs"
        ? await requestReviewApi("/graphics-handoff/books", { filter: "coverage_needs" })
        : await requestReviewApi("/api/review/graphics-books", { mode });
    } catch (primaryError) {
      if (!runtimeConfig.sheetReadFallbackEnabled) {
        throw primaryError;
      }
      setStatus(`Primary graphics backend is unavailable. Loading ${getGraphicsModeLabel(mode)} books directly from Google Sheets...`, {
        error: primaryError.message
      });
      data = await loadGraphicsBooksFromSheetsFallback(mode);
    }

    if (loadSequence !== graphicsBooksLoadSequence || mode !== getSelectedGraphicsMode()) {
      return;
    }

    currentGraphicsBookSummaries = Array.isArray(data.books)
      ? data.books.map(book => ({
        ...book,
        key: book.key || book.bookKey || normalizeBookKey(book.title || book.bookTitle || ""),
        title: book.title || book.bookTitle || "",
        count: Number(book.count ?? book.actionableCount ?? book.remainingActionableNeeded ?? 0)
      }))
      : [];
    graphicsBookSummaryByKey = new Map(
      currentGraphicsBookSummaries.map(book => [book.key || normalizeBookKey(book.title), book])
    );
    refreshGraphicsBookSelect(preserveSelection);

    setStatus(
      `Loaded ${getVisibleGraphicsBookSummaries().length} ${getGraphicsModeLabel(mode)} books. Backend ${data.version || "unknown"}.`
    );
  } catch (error) {
    if (loadSequence === graphicsBooksLoadSequence) {
      setStatus(`Graphics book load failed: ${error.message}`);
    }
  }
}

function renderExcerptHandoffSummary(records) {
  const handoffRecords = Array.isArray(records) ? records : [];
  currentExcerptHandoffRecords = handoffRecords;
  const queuedCount = handoffRecords.filter(record => (record.handoffStatus || "").toLowerCase() === "queued").length;
  const failedCount = handoffRecords.filter(record => (record.handoffStatus || "").toLowerCase() === "failed").length;
  const sentCount = handoffRecords.filter(record => (record.handoffStatus || "").toLowerCase() === "sent").length;

  if (elements.excerptHandoffCountBadge) {
    elements.excerptHandoffCountBadge.textContent = `${handoffRecords.length} Record${handoffRecords.length === 1 ? "" : "s"}`;
  }

  if (!elements.excerptHandoffSummary) {
    return;
  }

  if (!handoffRecords.length) {
    elements.excerptHandoffSummary.innerHTML = `<p class="empty-state">No queued or failed EXC handoffs need attention right now.</p>`;
    return;
  }

  const failedMarkup = handoffRecords
    .filter(record => (record.handoffStatus || "").toLowerCase() === "failed")
    .slice(0, 5)
    .map(record => `
      <li>
        <strong>${escapeHtml(record.poemTitle || "Untitled poem")}</strong>
        <span> · ${escapeHtml(record.bookTitle || "(blank book)")}</span>
        <span> · ${escapeHtml(record.errorMessage || "Unknown error")}</span>
      </li>
    `)
    .join("");

  elements.excerptHandoffSummary.innerHTML = `
    <div class="excerpt-card">
      <div class="excerpt-card__meta">
        <span class="badge badge--muted">Queued ${queuedCount}</span>
        <span class="badge badge--warn">Failed ${failedCount}</span>
        <span class="badge badge--signal">Sent ${sentCount}</span>
      </div>
      ${failedMarkup ? `<ul>${failedMarkup}</ul>` : `<p class="hint">No failed EXC handoffs in the current snapshot.</p>`}
    </div>
  `;
}

async function loadExcerptHandoffs() {
  try {
    const data = await requestReviewApi("/api/excerpts/handoffs", { status: "queued,failed,sent" });
    const records = Array.isArray(data.records) ? data.records : [];
    renderExcerptHandoffSummary(records);
  } catch (error) {
    setStatus(`EXC handoff load failed: ${error.message}`);
  }
}

async function retryFailedExcerptHandoffs() {
  try {
    setStatus("Retrying failed EXC handoffs...");
    const response = await fetch("/api/excerpts/handoffs/retry", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        statuses: ["failed"]
      })
    });
    const result = await response.json();
    if (!response.ok || !result.ok) {
      throw new Error(result.error || `/api/excerpts/handoffs/retry returned ${response.status}`);
    }
    await loadExcerptHandoffs();
    const poetryPlease = result.poetryPlease || {};
    const suffix = poetryPlease.ok ? "" : ` Poetry Please needs attention: ${poetryPlease.error || poetryPlease.reason || "handoff_failed"}.`;
    setStatus(`Retried ${Number(result.retriedCount || 0)} failed EXC handoff${Number(result.retriedCount || 0) === 1 ? "" : "s"}.${suffix}`);
  } catch (error) {
    setStatus(`EXC handoff retry failed: ${error.message}`);
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
  const bookTitle = isGraphicsQcSweepSelection(bookKey) ? "__qc_sweep__" : (summary?.title || bookKey);
  const statusLabel = isGraphicsQcSweepSelection(bookKey) ? "QC Sweep" : bookTitle;
  const loadSequence = ++graphicsRecordsLoadSequence;

  try {
    setStatus(`Loading ${getGraphicsModeLabel(mode)} rows for "${statusLabel}"...`);
    let data;
    try {
      if (mode === "coverage_needs") {
        const queueData = await requestReviewApi("/graphics-handoff/queue", { filter: "coverage_needs", limit: 500 });
        const records = Array.isArray(queueData.records)
          ? queueData.records.filter(record => normalizeBookKey(record.bookTitle || "") === bookKey)
          : [];
        data = { ...queueData, records };
      } else {
        data = await requestReviewApi("/api/review/graphics-records", { mode, bookTitle });
      }
    } catch (primaryError) {
      if (!runtimeConfig.sheetReadFallbackEnabled) {
        throw primaryError;
      }
      setStatus(`Primary graphics backend is unavailable. Loading ${getGraphicsModeLabel(mode)} rows directly from Google Sheets...`, {
        error: primaryError.message
      });
      data = await loadGraphicsRecordsFromSheetsFallback(bookTitle, mode);
    }

    if (
      loadSequence !== graphicsRecordsLoadSequence
      || mode !== getSelectedGraphicsMode()
      || bookKey !== (elements.graphicsBookSelect?.value || "")
    ) {
      return;
    }

    currentGraphicsRecords = Array.isArray(data.records) ? data.records : [];
    currentGraphicsCoverage = data.coverage || null;
    currentGraphicsAssetMatches = mode === "coverage_needs"
      ? new Map()
      : await loadGraphicsAssetMatches(currentGraphicsRecords);
    if (
      loadSequence !== graphicsRecordsLoadSequence
      || mode !== getSelectedGraphicsMode()
      || bookKey !== (elements.graphicsBookSelect?.value || "")
    ) {
      return;
    }
    renderGraphicsRecords(currentGraphicsRecords, currentGraphicsCoverage);
    setStatus(
      `Loaded ${currentGraphicsRecords.length} ${getGraphicsModeLabel(mode)} rows for "${statusLabel}".`,
      currentGraphicsRecords.slice(0, 5)
    );
  } catch (error) {
    if (loadSequence === graphicsRecordsLoadSequence) {
      setStatus(`Graphics record load failed: ${error.message}`);
    }
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

function updateCorrectionBookOptionCount(bookTitle, delta) {
  if (!elements.correctionBookSelect || !bookTitle || !delta) return;
  const option = Array.from(elements.correctionBookSelect.options).find(entry => entry.value === bookTitle);
  if (!option) return;
  const match = option.textContent.match(/\((\d+)\)$/);
  if (!match) return;
  const nextCount = Math.max(0, Number(match[1]) + delta);
  option.textContent = `${bookTitle} (${nextCount})`;
}

function renderGraphicsRecords(records) {
  refreshGraphicsFolderImportVisibility();
  const coverage = arguments[1] || null;
  if (elements.graphicsCountBadge) {
    elements.graphicsCountBadge.textContent = getSelectedGraphicsMode() === "coverage"
      ? `${Number(coverage?.summary?.totalExcerpts || 0)} Excerpts`
      : `${records.length} Rows`;
  }
  [elements.submitGraphicsQc, elements.submitGraphicsQcBottom].forEach(button => {
    if (button) {
      button.hidden = !["qc", "cleanup"].includes(getSelectedGraphicsMode());
    }
  });
  if (elements.exportGraphicsSheet) {
    elements.exportGraphicsSheet.hidden = !runtimeConfig.graphicsExportEnabled || getSelectedGraphicsMode() !== "queue";
  }

  if (!elements.graphicsList) {
    return;
  }

  elements.graphicsList.innerHTML = "";

  if (getSelectedGraphicsMode() === "coverage" && coverage) {
    elements.graphicsList.appendChild(buildCoverageCard(coverage));
    return;
  }

  if (getSelectedGraphicsMode() === "rework") {
    elements.graphicsList.appendChild(buildGraphicsReworkPanel(records));
    if (!records.length) {
      return;
    }
  }

  if (getSelectedGraphicsMode() === "handoff_retry") {
    elements.graphicsList.appendChild(buildGraphicsHandoffRetryPanel(records));
    if (!records.length) {
      return;
    }
  }

  if (!records.length) {
    const empty = document.createElement("p");
    empty.className = "empty-state";
    empty.textContent = getSelectedGraphicsMode() === "qc"
      ? "No QC rows remain for this book."
      : getSelectedGraphicsMode() === "cleanup"
      ? "No oddities remain for this book."
      : getSelectedGraphicsMode() === "mismatch"
        ? "No mismatched graphics remain for this book."
      : getSelectedGraphicsMode() === "rework"
        ? "No approved P.I.G. graphics are available for manual rework in this book."
      : getSelectedGraphicsMode() === "handoff_retry"
        ? "No failed Poetry Please handoffs remain for this book."
      : getSelectedGraphicsMode() === "coverage_needs"
        ? "No coverage needs remain for this book."
      : "No graphics creation rows remain for this book.";
    elements.graphicsList.appendChild(empty);
    return;
  }

  records.forEach(record => {
    elements.graphicsList.appendChild(
      getSelectedGraphicsMode() === "rework"
        ? buildGraphicsReworkCard(record)
        : getSelectedGraphicsMode() === "handoff_retry"
          ? buildGraphicsHandoffRetryCard(record)
          : buildGraphicsCard(record)
    );
  });
}

function buildGraphicsReworkPanel(records) {
  const panel = document.createElement("article");
  panel.className = "excerpt-card";
  panel.innerHTML = `
    <div class="excerpt-card__meta">
      <span class="badge badge--muted">${records.length} eligible graphic${records.length === 1 ? "" : "s"}</span>
      <span class="badge badge--warn">Original approvals stay in history</span>
    </div>
    <label class="field">
      <span>Rework note</span>
      <textarea id="graphics-rework-note" rows="3" placeholder="What needs to change before this goes back to P.I.G.?"></textarea>
    </label>
    <p class="hint">Select one or more approved P.I.G. graphics below, then send them back as fresh rework requests.</p>
    <button id="submit-graphics-rework" class="button">Send selected back to P.I.G.</button>
  `;
  panel.querySelector("#submit-graphics-rework")?.addEventListener("click", submitGraphicsReworkRequests);
  return panel;
}

function buildGraphicsReworkCard(record) {
  const card = document.createElement("article");
  card.className = "excerpt-card";
  card.dataset.recordId = record.recordId || "";
  card.dataset.pigCompletionId = record.pigCompletionId || "";
  card.dataset.graphicsRequestId = record.graphicsRequestId || "";
  card.dataset.author = record.author || "";
  card.dataset.poemTitle = record.poemTitle || "";
  card.dataset.bookTitle = record.bookTitle || "";
  card.dataset.quoteText = record.quoteText || record.text || "";
  card.innerHTML = `
    <div class="excerpt-card__meta">
      <label><input class="graphics-rework-checkbox" type="checkbox"> Select</label>
      <span class="badge badge--muted">Completion ${escapeHtml(record.pigCompletionId || "none")}</span>
      <span class="badge badge--muted">Book ${escapeHtml(record.bookTitle || "(blank)")}</span>
      <span class="excerpt-card__title">${escapeHtml(record.poemTitle || "Untitled poem")}</span>
      <span class="excerpt-card__author">${escapeHtml(record.author || "Unknown author")}</span>
    </div>
    <blockquote class="excerpt-card__quote">${escapeHtml(record.quoteText || "")}</blockquote>
    ${record.assetLinkUrl ? `<p class="hint"><a href="${escapeAttribute(record.assetLinkUrl)}" target="_blank" rel="noopener noreferrer">View current graphic</a></p>` : ""}
  `;
  return card;
}

async function submitGraphicsReworkRequests() {
  const note = document.getElementById("graphics-rework-note")?.value?.trim() || "";
  const selectedRecords = Array.from(elements.graphicsList?.querySelectorAll(".excerpt-card") || [])
    .filter(card => card.querySelector(".graphics-rework-checkbox:checked"))
    .map(card => ({
      recordId: card.dataset.recordId || "",
      pigCompletionId: card.dataset.pigCompletionId || "",
      graphicsRequestId: card.dataset.graphicsRequestId || "",
      author: card.dataset.author || "",
      poemTitle: card.dataset.poemTitle || "",
      bookTitle: card.dataset.bookTitle || "",
      quoteText: card.dataset.quoteText || ""
    }));

  if (!selectedRecords.length) {
    setStatus("Select at least one approved graphic to send back to P.I.G.");
    return;
  }
  if (!note) {
    setStatus("Add a rework note before sending approved graphics back to P.I.G.");
    return;
  }

  try {
    setStatus(`Sending ${selectedRecords.length} approved graphic${selectedRecords.length === 1 ? "" : "s"} back to P.I.G....`);
    const response = await fetch("/api/graphics/rework-request", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ records: selectedRecords, note })
    });
    const result = await response.json();
    if (!response.ok || !result.ok) {
      throw new Error(result.error || `/api/graphics/rework-request returned ${response.status}`);
    }
    await loadGraphicsBooks({ preserveSelection: false });
    if (elements.graphicsBookSelect?.value) {
      await loadGraphicsRecords();
    } else {
      currentGraphicsRecords = [];
      currentGraphicsCoverage = null;
      currentGraphicsAssetMatches = new Map();
      renderGraphicsRecords([], null);
    }
    setStatus(`Sent ${Number(result.createdCount || 0)} approved graphic${Number(result.createdCount || 0) === 1 ? "" : "s"} back to P.I.G. for rework.`);
  } catch (error) {
    setStatus(`Graphics rework request failed: ${error.message}`);
  }
}

function buildGraphicsHandoffRetryPanel(records) {
  const panel = document.createElement("article");
  panel.className = "excerpt-card";
  panel.innerHTML = `
    <div class="excerpt-card__meta">
      <span class="badge badge--warn">${records.length} failed handoff${records.length === 1 ? "" : "s"}</span>
      <span class="badge badge--muted">Retry sends these approved graphics to Poetry Please again</span>
    </div>
    <p class="hint">Select one or more failed records below, then retry their Poetry Please handoff.</p>
    <button id="submit-graphics-handoff-retry" class="button">Retry selected handoffs</button>
  `;
  panel.querySelector("#submit-graphics-handoff-retry")?.addEventListener("click", submitGraphicsHandoffRetries);
  return panel;
}

function buildGraphicsHandoffRetryCard(record) {
  const card = document.createElement("article");
  card.className = "excerpt-card";
  card.dataset.recordId = record.recordId || "";
  card.dataset.pigCompletionId = record.pigCompletionId || "";
  card.dataset.graphicsRequestId = record.graphicsRequestId || "";
  card.dataset.sheetRow = String(record.sheetRow || "");
  card.dataset.author = record.author || "";
  card.dataset.poemTitle = record.poemTitle || "";
  card.dataset.bookTitle = record.bookTitle || "";
  card.dataset.quoteText = record.quoteText || "";
  card.dataset.assetLinkUrl = record.assetLinkUrl || "";
  card.dataset.assetPreviewUrl = record.assetPreviewUrl || "";
  card.dataset.completedAt = record.completedAt || "";
  card.dataset.graphicsQcUpdatedAt = record.graphicsQcUpdatedAt || "";
  card.dataset.graphicsQcDecision = record.graphicsQcDecision || "";
  card.dataset.graphicsQcNote = record.graphicsQcNote || "";
  card.dataset.notes = record.notes || "";
  card.dataset.sourceTool = record.sourceTool || "";
  card.innerHTML = `
    <div class="excerpt-card__meta">
      <label><input class="graphics-handoff-retry-checkbox" type="checkbox"> Select</label>
      <span class="badge badge--warn">${escapeHtml(record.poetryPleaseStatus || "FAILED")}</span>
      <span class="badge badge--muted">Book ${escapeHtml(record.bookTitle || "(blank)")}</span>
      <span class="excerpt-card__title">${escapeHtml(record.poemTitle || "Untitled poem")}</span>
      <span class="excerpt-card__author">${escapeHtml(record.author || "Unknown author")}</span>
    </div>
    <blockquote class="excerpt-card__quote">${escapeHtml(record.quoteText || "")}</blockquote>
    ${record.assetLinkUrl ? `<p class="hint"><a href="${escapeAttribute(record.assetLinkUrl)}" target="_blank" rel="noopener noreferrer">View approved graphic</a></p>` : ""}
    ${record.poetryPleaseNote ? `<p class="hint">Last failure: ${escapeHtml(record.poetryPleaseNote)}</p>` : ""}
  `;
  return card;
}

async function submitGraphicsHandoffRetries() {
  const selectedRecords = Array.from(elements.graphicsList?.querySelectorAll(".excerpt-card") || [])
    .filter(card => card.querySelector(".graphics-handoff-retry-checkbox:checked"))
    .map(card => ({
      recordId: card.dataset.recordId || "",
      pigCompletionId: card.dataset.pigCompletionId || "",
      graphicsRequestId: card.dataset.graphicsRequestId || "",
      sheetRow: card.dataset.sheetRow || "",
      author: card.dataset.author || "",
      poemTitle: card.dataset.poemTitle || "",
      bookTitle: card.dataset.bookTitle || "",
      quoteText: card.dataset.quoteText || "",
      assetLinkUrl: card.dataset.assetLinkUrl || "",
      assetPreviewUrl: card.dataset.assetPreviewUrl || "",
      completedAt: card.dataset.completedAt || "",
      graphicsQcUpdatedAt: card.dataset.graphicsQcUpdatedAt || "",
      graphicsQcDecision: card.dataset.graphicsQcDecision || "",
      graphicsQcNote: card.dataset.graphicsQcNote || "",
      notes: card.dataset.notes || "",
      sourceTool: card.dataset.sourceTool || ""
    }));

  if (!selectedRecords.length) {
    setStatus("Select at least one failed graphics handoff to retry.");
    return;
  }

  try {
    setStatus(`Retrying ${selectedRecords.length} failed graphics handoff${selectedRecords.length === 1 ? "" : "s"}...`);
    const response = await fetch("/api/graphics/handoffs/retry", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ records: selectedRecords })
    });
    const result = await response.json();
    if (!response.ok || !result.ok) {
      throw new Error(result.error || `/api/graphics/handoffs/retry returned ${response.status}`);
    }
    const handoff = result.poetryPlease || {};
    const suffix = handoff.ok ? "" : ` Poetry Please still needs attention: ${handoff.error || handoff.reason || "handoff_failed"}.`;
    await loadGraphicsBooks({ preserveSelection: false });
    if (elements.graphicsBookSelect?.value) {
      await loadGraphicsRecords();
    } else {
      currentGraphicsRecords = [];
      currentGraphicsCoverage = null;
      currentGraphicsAssetMatches = new Map();
      renderGraphicsRecords([], null);
    }
    setStatus(`Retried ${Number(result.retriedCount || 0)} failed graphics handoff${Number(result.retriedCount || 0) === 1 ? "" : "s"}.${suffix}`);
  } catch (error) {
    setStatus(`Graphics handoff retry failed: ${error.message}`);
  }
}

function buildCoverageCard(coverage) {
  const card = document.createElement("article");
  card.className = "excerpt-card";
  const summary = coverage.summary || {};
  const topPoems = Array.isArray(coverage.topPoems) ? coverage.topPoems : [];
  const missingPoems = Array.isArray(coverage.missingPoems) ? coverage.missingPoems : [];
  const contributors = Array.isArray(coverage.contributors) ? coverage.contributors : [];

  card.innerHTML = `
    <div class="excerpt-card__meta">
      <span class="badge badge--muted">Book ${escapeHtml(coverage.bookTitle || "(blank)")}</span>
      <span class="badge badge--signal">${escapeHtml(String(summary.totalExcerpts || 0))} excerpts</span>
      <span class="badge badge--muted">${escapeHtml(String(summary.poemsWithExcerpts || 0))} poems covered</span>
      <span class="badge badge--warn">${escapeHtml(String(summary.poemsWithoutExcerpts || 0))} poems missing</span>
      <span class="badge badge--muted">${escapeHtml(String(summary.contributors || 0))} contributors</span>
    </div>
    <div class="correction-source__grid">
      <div><strong>Catalog poems</strong><span>${escapeHtml(String(summary.totalCatalogPoems || 0))}</span></div>
      <div><strong>Covered poems</strong><span>${escapeHtml(String(summary.poemsWithExcerpts || 0))}</span></div>
      <div><strong>Missing poems</strong><span>${escapeHtml(String(summary.poemsWithoutExcerpts || 0))}</span></div>
      <div><strong>Total pulls</strong><span>${escapeHtml(String(summary.totalExcerpts || 0))}</span></div>
    </div>
    <div class="hint"><strong>Most-covered poems</strong><br>${topPoems.length ? topPoems.map(poem => `${escapeHtml(poem.poemTitle)} · ${escapeHtml(String(poem.excerptCount))} pulls · ${escapeHtml(String(poem.contributorCount))} people`).join("<br>") : "No excerpt pulls yet."}</div>
    <div class="hint"><strong>Contributors</strong><br>${contributors.length ? contributors.map(person => `${escapeHtml(person.email)} · ${escapeHtml(String(person.excerptCount))} pulls · ${escapeHtml(String(person.uniquePoems))} poems`).join("<br>") : "No contributor activity yet."}</div>
    <div class="hint"><strong>Poems still missing</strong><br>${missingPoems.length ? missingPoems.map(title => escapeHtml(title)).join("<br>") : "No missing catalog poems in this snapshot."}</div>
  `;
  return card;
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

    collapseExactQueueDuplicates(group.excerpts).forEach((excerpt, index) => {
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

function applyOptimisticGraphicsBookCountUpdate(removedRecords) {
  if (!Array.isArray(removedRecords) || !removedRecords.length) return;
  const removedByBookKey = new Map();
  removedRecords.forEach(record => {
    const key = normalizeBookKey(record.bookTitle);
    if (!key) return;
    removedByBookKey.set(key, (removedByBookKey.get(key) || 0) + 1);
  });

  currentGraphicsBookSummaries = currentGraphicsBookSummaries
    .map(book => {
      const key = normalizeBookKey(book.title);
      const removedCount = removedByBookKey.get(key) || 0;
      if (!removedCount) return book;
      return {
        ...book,
        count: Math.max(0, Number(book.count || 0) - removedCount)
      };
    })
    .filter(book => Number(book.count || 0) > 0);

  graphicsBookSummaryByKey = new Map(
    currentGraphicsBookSummaries.map(book => [normalizeBookKey(book.title), book])
  );
  refreshGraphicsBookSelect(true);
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
    const selectedCatalog = getSelectedReviewReleaseCatalog();
    return excerpts.filter(excerpt => (
      reviewQueueIncludeSet.has(normalizeBookKey(excerpt.bookTitle))
      && matchesSelectedReleaseCatalog(excerpt.bookTitle, selectedCatalog)
    ));
  }
  if (mode === "videos") {
    return excerpts.filter(isVideoReviewRecord);
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
    return "No pending excerpts in this book are currently in the release-catalog lane.";
  }
  if (getSelectedReviewFilter() === "videos") {
    return "No pending video excerpts are currently loaded.";
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

function normalizeExactQueueDuplicateText(text) {
  return (text || "").toString().replace(/\r\n/g, "\n").trim();
}

function collapseExactQueueDuplicates(excerpts) {
  const byText = new Map();
  const collapsed = [];

  excerpts.forEach(excerpt => {
    const key = normalizeExactQueueDuplicateText(excerpt.excerptText || excerpt.rawExcerptText || "");
    if (!key) {
      collapsed.push(excerpt);
      return;
    }
    if (!byText.has(key)) {
      byText.set(key, excerpt);
      collapsed.push(excerpt);
      return;
    }
    const canonical = byText.get(key);
    canonical.queueDuplicateRows = [
      ...(Array.isArray(canonical.queueDuplicateRows) ? canonical.queueDuplicateRows : []),
      {
        sourceRow: Number(excerpt.sourceRow),
        recordId: excerpt.recordId || "",
        note: `Auto-rejected by Weaver as exact queue duplicate of row ${canonical.sourceRow}.`
      }
    ];
  });

  return collapsed;
}

function buildExcerptCard(excerpt, uniqueKey) {
  const card = document.createElement("article");
  card.className = "excerpt-card";
  card.dataset.sourceRow = excerpt.sourceRow;
  card.dataset.recordId = excerpt.recordId || "";
  card.dataset.queueDuplicateRows = JSON.stringify(excerpt.queueDuplicateRows || []);

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
  const queueDuplicateBadge =
    Array.isArray(excerpt.queueDuplicateRows) && excerpt.queueDuplicateRows.length
      ? `<span class="badge badge--warn">Exact queue duplicate x${escapeHtml(String(excerpt.queueDuplicateRows.length + 1))}</span>`
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

  const decisionBadge = reviewDecision === "accept" || reviewDecision === "accept_skip_graphic"
    ? `<span class="badge badge--signal">${reviewDecision === "accept_skip_graphic" ? "Accepted excerpt, no graphic" : "Accepted excerpt"}</span>`
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
        ${queueDuplicateBadge}
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
    ${queueDuplicateBadge ? `<p class="hint">This card represents ${escapeHtml(String(excerpt.queueDuplicateRows.length + 1))} exact queue duplicates. Accepting it will keep row ${escapeHtml(String(excerpt.sourceRow))} and auto-reject the duplicate row${excerpt.queueDuplicateRows.length === 1 ? "" : "s"} with an explicit Weaver duplicate note.</p>` : ""}
    <blockquote class="excerpt-card__quote">${escapeHtml(excerpt.excerptText)}</blockquote>
    <p class="hint excerpt-card__hint">Accept sends this to the quote-image queue. Accept excerpt, skip graphic approves the excerpt for Poetry Please without sending it to P.I.G.</p>
    <div class="decision-group">
      <label><input type="radio" name="approval-${uniqueKey}" value="accept" ${currentDecision === "accept" ? "checked" : ""}> Accept</label>
      <label><input type="radio" name="approval-${uniqueKey}" value="accept_skip_graphic" ${currentDecision === "accept_skip_graphic" ? "checked" : ""}> Accept excerpt, skip graphic</label>
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
  card.dataset.requestKind = record.requestKind || "";
  card.dataset.author = record.author || "";
  card.dataset.poemTitle = record.poemTitle || "";
  card.dataset.bookTitle = record.bookTitle || "";
  card.dataset.quoteText = record.quoteText || "";
  card.dataset.assetLinkUrl = record.assetLinkUrl || record.previousAssetUrl || "";
  card.dataset.assetPreviewUrl = record.assetPreviewUrl || record.previousAssetPreviewUrl || "";
  card.dataset.contentType = record.contentType || record.imageType || "";
  card.dataset.completedAt = record.completedAt || "";
  card.dataset.sourceTool = record.sourceTool || "P.I.G.";
  card.dataset.currentQcDecision = normalizeGraphicsQcDecisionClient(record.graphicsQcDecision || "");
  card.dataset.currentQcNote = record.graphicsQcNote || "";
  card.dataset.currentReplacementAssetUrl = record.assetLinkUrl || "";
  card.dataset.poetryPleaseStatus = record.poetryPleaseStatus || "";
  card.dataset.poetryPleaseUpdatedAt = record.poetryPleaseUpdatedAt || "";
  card.dataset.poetryPleaseNote = record.poetryPleaseNote || "";

  const approved = normalizeApprovalForCompare(record.approved) === "Y";
  const created = normalizeApprovalForCompare(record.created) === "Y";
  const workflowStatus = (record.workflowStatus || "").trim();
  const notes = (record.notes || "").trim();
  const quoteText = record.quoteText || record.text || "";
  const wordCount = countWordsFromText(quoteText);
  const assetMatch = currentGraphicsAssetMatches.get(record.recordId || String(record.sheetRow || ""));
  const displayAuthor = record.author || assetMatch?.author || "";
  const assetFolderUrl = record.assetFolderUrl || "";
  const assetLinkUrl = record.assetLinkUrl || record.previousAssetUrl || assetMatch?.linkUrl || assetMatch?.lowResLinkUrl || assetMatch?.folderLink || "";
  const assetLinkLabel = assetFolderUrl && !record.assetPreviewUrl
    ? "Open returned folder"
    : (assetMatch?.fileName || (assetMatch?.folderLink ? "View Drive folder" : "View graphic on Drive"));
  const rawAssetPreviewUrl = record.assetPreviewUrl || record.previousAssetPreviewUrl || "";
  const assetPreviewUrl = normalizeGraphicsPreviewUrl(rawAssetPreviewUrl || assetLinkUrl, assetLinkUrl || rawAssetPreviewUrl) || buildGraphicsPreviewUrl(assetMatch);
  const isQcMode = ["qc", "cleanup"].includes(getSelectedGraphicsMode());
  const isMismatchMode = getSelectedGraphicsMode() === "mismatch";
  const isHandoffMode = getSelectedGraphicsMode() === "handoff";
  const isCoverageNeedsMode = getSelectedGraphicsMode() === "coverage_needs";
  const isReworkCompletion = record.requestKind === "rework_completion" || record.isReworkCompletion === true;
  const approvedCount = Number(record.approvedCount || 0);
  const poetryPleaseQiCount = record.poetryPleaseQiCount ?? "";
  const weaverCompletedQiCount = record.weaverCompletedQiCount ?? "";
  const pendingQcCount = Number(record.pendingQcCount || 0);
  const inProgressCount = Number(record.inProgressCount || 0);
  const targetCount = Number(record.targetCount || 25);
  const remainingActionableNeeded = Number(record.remainingActionableNeeded || 0);
  const priorityTier = record.priorityTier || "";
  const priorityScore = record.priorityScore ?? "";
  const excerptRating = record.excerptRating ?? "";
  const poemRating = record.poemRating ?? "";
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
  const defaultQcNote = currentQcDecision === "replace"
    ? (GRAPHICS_QC_DEFAULT_NOTES.replace || "")
    : (GRAPHICS_QC_DEFAULT_NOTES[qcRejectReason] || "");

  card.innerHTML = `
    <div class="excerpt-card__meta">
      <span class="badge badge--muted">Sheet row ${escapeHtml(String(record.sheetRow || ""))}</span>
      <span class="badge badge--muted">Record ${escapeHtml(record.recordId || "none")}</span>
      <span class="badge badge--muted">Book ${escapeHtml(record.bookTitle || "(blank)")}</span>
      <span class="badge ${getWordCountBadgeClass(wordCount)}">Words ${wordCount}</span>
      ${workflowStatus ? `<span class="badge badge--warn">${escapeHtml(workflowStatus)}</span>` : ""}
      ${isReworkCompletion ? '<span class="badge badge--warn">Rework completion</span>' : ""}
      ${approved ? '<span class="badge badge--signal">Approved</span>' : '<span class="badge badge--muted">Not approved</span>'}
      ${created ? '<span class="badge badge--signal">Created</span>' : '<span class="badge badge--muted">Not created</span>'}
      ${isHandoffMode && poetryPleaseStatus ? `<span class="badge ${poetryPleaseStatus === "HANDED_OFF" ? "badge--signal" : poetryPleaseStatus === "FAILED" ? "badge--warn" : "badge--muted"}">${escapeHtml(poetryPleaseStatus.replace(/_/g, " "))}</span>` : ""}
      ${isCoverageNeedsMode ? `<span class="badge badge--signal">${escapeHtml(record.statusLabel || "Needs coverage")}</span>` : ""}
      ${isCoverageNeedsMode ? `<span class="badge badge--muted">${approvedCount} approved + ${pendingQcCount} pending / ${targetCount}</span>` : ""}
      ${isCoverageNeedsMode ? `<span class="badge badge--warn">${remainingActionableNeeded} still needed</span>` : ""}
      <span class="excerpt-card__title">${escapeHtml(record.poemTitle || "Untitled poem")}</span>
      <span class="excerpt-card__author">${escapeHtml(displayAuthor || "Unknown author")}</span>
    </div>
    <blockquote class="excerpt-card__quote">${escapeHtml(quoteText)}</blockquote>
    ${(isQcMode || isHandoffMode || isMismatchMode) && assetPreviewUrl ? `
      <figure class="graphics-preview">
        <img class="graphics-preview__image" src="${escapeAttribute(assetPreviewUrl)}" alt="Existing graphic preview for ${escapeAttribute(record.poemTitle || "this excerpt")}" loading="lazy" />
      </figure>
    ` : ""}
    ${isQcMode && assetFolderUrl && !assetPreviewUrl ? `<p class="hint">P.I.G. returned a Drive folder for this rework, not a directly previewable image file.</p>` : ""}
    ${isQcMode && isReworkCompletion ? `<p class="hint">This graphic is a returned rework from P.I.G. It is back in Weaver for a fresh QC decision.</p>` : ""}
    <div class="correction-source__grid">
      <div><strong>Book</strong><span>${escapeHtml(record.bookTitle || "")}</span></div>
      <div><strong>Approved?</strong><span>${escapeHtml(record.approved || "")}</span></div>
      <div><strong>Created?</strong><span>${escapeHtml(record.created || "")}</span></div>
      <div><strong>Workflow</strong><span>${escapeHtml(workflowStatus || "—")}</span></div>
      ${isReworkCompletion ? `<div><strong>Request kind</strong><span>Rework completion</span></div>` : ""}
      <div><strong>Record ID</strong><span>${escapeHtml(record.recordId || "—")}</span></div>
      <div><strong>Request ID</strong><span>${escapeHtml(record.graphicsRequestId || "—")}</span></div>
      <div><strong>Sheet Row</strong><span>${escapeHtml(String(record.sheetRow || "—"))}</span></div>
      ${isHandoffMode ? `<div><strong>Handoff</strong><span>${escapeHtml(poetryPleaseStatus || "Pending")}</span></div>` : ""}
      ${isCoverageNeedsMode ? `<div><strong>Coverage</strong><span>${approvedCount} approved + ${pendingQcCount} pending + ${inProgressCount} in progress / ${targetCount}</span></div>` : ""}
      ${isCoverageNeedsMode ? `<div><strong>Poetry Please QI</strong><span>${escapeHtml(String(poetryPleaseQiCount === null || poetryPleaseQiCount === "" ? "fallback unavailable" : poetryPleaseQiCount))} imported / approved</span></div>` : ""}
      ${isCoverageNeedsMode ? `<div><strong>Weaver completed QI</strong><span>${escapeHtml(String(weaverCompletedQiCount === null || weaverCompletedQiCount === "" ? "—" : weaverCompletedQiCount))} diagnostic only</span></div>` : ""}
      ${isCoverageNeedsMode ? `<div><strong>Still needed</strong><span>${remainingActionableNeeded} actionable</span></div>` : ""}
      ${isCoverageNeedsMode ? `<div><strong>Priority</strong><span>Tier ${escapeHtml(String(priorityTier || "—"))} · Score ${escapeHtml(String(priorityScore || "—"))}</span></div>` : ""}
      ${isCoverageNeedsMode ? `<div><strong>Ratings</strong><span>Excerpt ${escapeHtml(String(excerptRating || "—"))} · Poem ${escapeHtml(String(poemRating || "—"))}</span></div>` : ""}
      ${isCoverageNeedsMode && (record.reworkReason || record.rejectReason || record.qcNote) ? `<div><strong>Why present</strong><span>${escapeHtml(record.reworkReason || record.rejectReason || record.qcNote || "")}</span></div>` : ""}
    </div>
    ${assetLinkUrl ? `<p class="hint"><a href="${escapeAttribute(assetLinkUrl)}" target="_blank" rel="noopener noreferrer">${escapeHtml(assetLinkLabel)}</a>${assetMatch?.matchType === "cleanup_override" ? " · matched from cleanup override" : ""}${card.dataset.storageTarget === "pig_sheet" ? " · returned from P.I.G." : ""}</p>` : ""}
    ${isMismatchMode ? `<p class="hint">Mismatch reason: ${escapeHtml(record.rejectReason || "mismatched_graphic")}${record.graphicsQcUpdatedAt ? ` · ${escapeHtml(record.graphicsQcUpdatedAt)}` : ""}</p>` : ""}
    ${isHandoffMode && (poetryPleaseUpdatedAt || poetryPleaseNote) ? `<p class="hint">Poetry Please: ${escapeHtml(poetryPleaseStatus || "Pending")}${poetryPleaseUpdatedAt ? ` · ${escapeHtml(poetryPleaseUpdatedAt)}` : ""}${poetryPleaseNote ? ` · ${escapeHtml(poetryPleaseNote)}` : ""}</p>` : ""}
    ${isQcMode ? `
      <fieldset class="graphics-qc-controls">
        <legend>QC decision</legend>
        <label><input type="radio" name="graphics-qc-${escapeAttribute(record.recordId || String(record.sheetRow || ""))}" value="approve" ${currentQcDecision === "approve" ? "checked" : ""}> Approve</label>
        <label><input type="radio" name="graphics-qc-${escapeAttribute(record.recordId || String(record.sheetRow || ""))}" value="reject" ${currentQcDecision === "reject" ? "checked" : ""}> Reject</label>
        <label><input type="radio" name="graphics-qc-${escapeAttribute(record.recordId || String(record.sheetRow || ""))}" value="replace" ${currentQcDecision === "replace" ? "checked" : ""}> Replace here</label>
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
      <div class="graphics-qc-replacement" ${currentQcDecision === "replace" ? "" : "hidden"}>
        <label class="field">
          <span>Replacement graphic Drive link</span>
          <input class="graphics-qc-replacement-url" type="url" value="" placeholder="https://drive.google.com/file/d/...">
        </label>
        <p class="hint">Use a Google Drive image file link or a Drive folder link. If you use a folder, Weaver will choose the best-matching image for this card and send it forward to Poetry Please.</p>
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
    const replacementBlock = card.querySelector(".graphics-qc-replacement");
    const rejectReasonSelect = card.querySelector(".graphics-qc-reject-reason");
    const metadataSelect = card.querySelector(".graphics-qc-metadata");
    const aestheticSelect = card.querySelector(".graphics-qc-aesthetic");
    const radios = Array.from(card.querySelectorAll(`input[name="graphics-qc-${CSS.escape(record.recordId || String(record.sheetRow || ""))}"]`));
    const syncQcFields = decision => {
      if (detailsBlock) {
        detailsBlock.hidden = decision !== "reject";
      }
      if (replacementBlock) {
        replacementBlock.hidden = decision !== "replace";
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
      const nextDefault = currentDecision === "replace"
        ? (GRAPHICS_QC_DEFAULT_NOTES.replace || "")
        : currentDecision === "reject"
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
    const replacementAssetUrl = card.querySelector(".graphics-qc-replacement-url")?.value?.trim() || "";
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
      assetLinkUrl: card.dataset.assetLinkUrl || "",
      assetUrl: card.dataset.assetLinkUrl || "",
      assetPreviewUrl: card.dataset.assetPreviewUrl || "",
      contentType: card.dataset.contentType || "",
      imageType: card.dataset.contentType || "",
      completedAt: card.dataset.completedAt || "",
      sourceTool: card.dataset.sourceTool || "P.I.G.",
      rejectReason,
      metadataIssue,
      aestheticIssue,
      replacementAssetUrl,
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
      (update.qcNote || "") !== (card.dataset.currentQcNote || "") ||
      (
        (update.qcDecision || "") === "replace" &&
        (update.replacementAssetUrl || "") !== (card.dataset.currentReplacementAssetUrl || "")
      )
    );
  });
}

async function submitGraphicsQc() {
  if (!["qc", "cleanup"].includes(getSelectedGraphicsMode())) {
    setStatus("Graphics QC decisions only apply in Graphics QC or Oddities mode.");
    return;
  }

  const updates = filterChangedGraphicsQcUpdates(collectGraphicsQcUpdates());
  if (!updates.length) {
    setStatus("No changed QC decisions to submit.");
    return;
  }

  const noteRequired = updates.find(update => (
    (
      update.qcDecision === "replace" &&
      !update.replacementAssetUrl
    ) ||
    (
      update.qcDecision === "replace" &&
      update.storageTarget !== "pig_sheet"
    ) ||
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
    setStatus("Add the missing QC detail. Replace here needs a Google Drive image or folder link and currently only works for returned P.I.G. graphics. Rejects need a reason. Mismatched graphics need a note. Correct and recreate needs at least one issue selected, plus a note for metadata fixes or aesthetic 'other'.");
    return;
  }

  try {
    setSubmitState(true, `Saving ${updates.length}...`);
    setStatus(`Saving ${updates.length} QC decisions...`);
    const savedRecordIds = new Set(updates.map(update => update.recordId));
    const removedGraphicsRecords = currentGraphicsRecords.filter(record => savedRecordIds.has(record.recordId));

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

    const qcSweepMode = isGraphicsQcSweepSelection();
    currentGraphicsRecords = currentGraphicsRecords.filter(record => !savedRecordIds.has(record.recordId));
    renderGraphicsRecords(currentGraphicsRecords);
    applyOptimisticGraphicsBookCountUpdate(removedGraphicsRecords);
    await loadGraphicsBooks();
    if (elements.graphicsBookSelect?.value) {
      await loadGraphicsRecords();
      if (qcSweepMode) {
        await wait(1200);
        await loadGraphicsBooks();
        if (elements.graphicsBookSelect?.value) {
          await loadGraphicsRecords();
        }
      }
    } else {
      currentGraphicsRecords = [];
      currentGraphicsAssetMatches = new Map();
      renderGraphicsRecords([]);
    }
    const handoff = result.poetryPlease || {};
    const handoffMessage = handoff.ok && !handoff.skipped
      ? ` Sent ${Number(handoff.createdCount || 0) + Number(handoff.updatedCount || 0)} approved graphic${(Number(handoff.createdCount || 0) + Number(handoff.updatedCount || 0)) === 1 ? "" : "s"} to Poetry Please.`
      : (!handoff.ok && handoff.error ? ` Poetry Please handoff needs attention: ${handoff.error}` : "");
    const sweepMessage = qcSweepMode ? " Loaded the next QC sweep item." : "";
    setStatus(`Saved ${updates.length} QC decisions.${handoffMessage}${sweepMessage}`);
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
  const locationLabel = `${escapeHtml(meta || "Existing source row")}${match.sourceRow ? `, row ${escapeHtml(String(match.sourceRow))}` : ""}.`;
  const matchMarkup = `<p class="validation validation--library-match"><strong>Library match:</strong> ${escapeHtml(label)} ${locationLabel} ${excerptLink}</p>`;
  const statusMarkup = statusLabel
    ? `<p class="validation validation--library-status"><strong>Production status:</strong> ${escapeHtml(statusLabel)}</p>`
    : "";
  return `${matchMarkup}${statusMarkup}`;
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
  return Array.from(container.querySelectorAll(".excerpt-card")).flatMap(card => {
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

    const update = {
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

    const updates = [update];
    const duplicateRows = JSON.parse(card.dataset.queueDuplicateRows || "[]");
    if (reviewDecision === "accept" || reviewDecision === "accept_skip_graphic") {
      duplicateRows.forEach(duplicate => {
        updates.push({
          sourceRow: Number(duplicate.sourceRow),
          recordId: duplicate.recordId || "",
          reviewDecision: "reject",
          correctionNote: duplicate.note || `Auto-rejected by Weaver as exact queue duplicate of row ${sourceRow}.`,
          correctedAuthor: "",
          correctedTitle: "",
          correctedBookTitle: "",
          correctedExcerpt: "",
          duplicateGroupId: `queue-exact:${sourceRow}`,
          useForQi: false,
          useForInt: false
        });
      });
    }

    return updates;
  });
}

function filterChangedUpdates(updates, excerpts) {
  const bySourceRow = new Map(
    excerpts.map(excerpt => [Number(excerpt.sourceRow), excerpt])
  );

  return updates.filter(update => {
    const current = bySourceRow.get(Number(update.sourceRow));
    if (!current) return true;

    const currentExplicitDecision = normalizeDecision(current.excerptReviewDecision);
    const currentEffectiveDecision = getEffectiveReviewDecision(current);
    const updateDecision = normalizeDecision(update.reviewDecision);

    // Legacy/backfilled rows can carry approved-for-QI state without an explicit
    // review decision. In batch mode, "No decision" should leave those rows in
    // place rather than clearing the implicit approval flag underneath them.
    if (!updateDecision && !currentExplicitDecision) {
      return (
        normalizeCorrectionNote(update.correctionNote) !== normalizeCorrectionNote(current.correctionNote) ||
        normalizeCorrectionNote(update.correctedAuthor) !== normalizeCorrectionNote(current.correctedAuthor) ||
        normalizeCorrectionNote(update.correctedTitle) !== normalizeCorrectionNote(current.correctedTitle) ||
        normalizeCorrectionNote(update.correctedBookTitle) !== normalizeCorrectionNote(current.correctedBookTitle) ||
        normalizeCorrectionNote(update.correctedExcerpt) !== normalizeCorrectionNote(current.correctedExcerpt)
      );
    }

    return (
      updateDecision !== currentEffectiveDecision ||
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
  const normalized = (value || "").toString().trim().toLowerCase();
  if (normalized === "accept_skip_graphics") return "accept_skip_graphic";
  return normalized;
}

function normalizeApprovalForCompare(value) {
  const normalized = (value || "").toString().trim().toLowerCase();
  if (normalized === "accept" || normalized === "accept_skip_graphic" || normalized === "accept_skip_graphics" || normalized === "y") return "Y";
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
      currentPendingRecords = currentPendingRecords.filter(record => !savedSourceRows.has(Number(record.sourceRow)));
      applyPendingBookData(currentPendingRecords, { preserveSelection: true });
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
    afterSaveOptimistic: changedUpdates => {
      const savedSourceRows = new Set(changedUpdates.map(update => Number(update.sourceRow)));
      currentCorrectionExcerpts = currentCorrectionExcerpts.filter(excerpt => !savedSourceRows.has(Number(excerpt.sourceRow)));
      renderCorrectionExcerpts(currentCorrectionExcerpts);
      updateCorrectionBookOptionCount(elements.correctionBookSelect.value, -savedSourceRows.size);
    },
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
elements.submitReviewBottom?.addEventListener("click", submitReview);
elements.submitWeirdReview?.addEventListener("click", submitWeirdReview);
elements.submitWeirdReviewBottom?.addEventListener("click", submitWeirdReview);
elements.loadCorrectionBooks?.addEventListener("click", loadCorrectionBooks);
elements.loadCorrections?.addEventListener("click", loadCorrections);
elements.autoApplyCorrections?.addEventListener("click", applyAutoCorrectionsToLoadedQueue);
elements.submitCorrections?.addEventListener("click", submitCorrections);
elements.submitCorrectionsBottom?.addEventListener("click", submitCorrections);
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
elements.gatheringTabBatch?.addEventListener("click", () => setGatheringMode("batch"));
elements.gatheringBookBold?.addEventListener("click", () => wrapTextareaSelectionWithBold(elements.gatheringBookQuote));
elements.gatheringBookItalicize?.addEventListener("click", () => wrapTextareaSelectionWithItalics(elements.gatheringBookQuote));
elements.gatheringBatchBold?.addEventListener("click", () => wrapTextareaSelectionWithBold(elements.gatheringBatchSource));
elements.gatheringBatchItalicize?.addEventListener("click", () => wrapTextareaSelectionWithItalics(elements.gatheringBatchSource));
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
  currentGraphicsView = "main";
  syncGraphicsModeOptionsForView();
  setActiveModule("graphics");
  updateGraphicsModuleTitle();
  syncReleaseCatalogFilterUi();
  loadExcerptHandoffs();
  if (elements.graphicsBookSelect?.options.length <= 1) {
    loadGraphicsBooks();
  }
});
elements.showGraphicsOps?.addEventListener("click", () => {
  currentGraphicsView = "ops";
  syncGraphicsModeOptionsForView();
  setActiveModule("graphics");
  updateGraphicsModuleTitle();
  syncReleaseCatalogFilterUi();
  loadExcerptHandoffs();
  loadGraphicsBooks({ preserveSelection: false });
});
elements.loadGraphicsBooks?.addEventListener("click", loadGraphicsBooks);
elements.loadGraphicsRecords?.addEventListener("click", loadGraphicsRecords);
elements.previewGraphicsFolderImport?.addEventListener("click", previewGraphicsFolderImport);
elements.applyGraphicsFolderImport?.addEventListener("click", applyGraphicsFolderImport);
elements.refreshExcerptHandoffs?.addEventListener("click", loadExcerptHandoffs);
elements.retryFailedExcerptHandoffs?.addEventListener("click", retryFailedExcerptHandoffs);
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
elements.submitGraphicsQcBottom?.addEventListener("click", submitGraphicsQc);
elements.loadGatheringOptions?.addEventListener("click", () => {
  loadGatheringOptions({ force: true }).catch(error => {
    setStatus(`Excerpt gathering option load failed: ${error.message}`);
  });
});
elements.submitGathering?.addEventListener("click", submitGathering);
elements.previewGatheringBatch?.addEventListener("click", previewGatheringBatch);
elements.submitGatheringBatch?.addEventListener("click", submitGatheringBatch);
elements.gatheringEmail?.addEventListener("input", () => {
  if (elements.gatheringEmail?.value.trim()) {
    setGatheringEmailWarning("");
  }
  applyContributorAccessMode();
  loadGatheringOptions({ force: true }).catch(error => {
    setStatus(`Excerpt gathering option load failed: ${error.message}`);
  });
  if (currentContributorAccess || isContributorShellRequested()) {
    handleGatheringBookSelectionChange({ preserveTitle: true }).catch(error => {
      setStatus(`Book metadata load failed: ${error.message}`);
    });
    setActiveModule("gathering");
  }
});
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
elements.gatheringVideoLoadPlaylist?.addEventListener("click", () => loadGatheringVideoPlaylist());
elements.gatheringVideoPrevItem?.addEventListener("click", () => navigateGatheringVideoPlaylist(-1));
elements.gatheringVideoNextItem?.addEventListener("click", () => navigateGatheringVideoPlaylist(1));
elements.gatheringVideoOpenItem?.addEventListener("click", openGatheringVideoPlaylistItem);
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
    syncReleaseCatalogFilterUi();
    refreshReviewBookSelect(true);
    reviewVisibleCount = getReviewBatchSize();
    reviewPinnedRowOrder = [];
    if (getSelectedReviewFilter() === "videos" && elements.bookSelect?.querySelector(`option[value="${REVIEW_VIDEOS_BOOK_KEY}"]`)) {
      elements.bookSelect.value = REVIEW_VIDEOS_BOOK_KEY;
      loadExcerpts();
      return;
    }
    renderCurrentExcerpts();
  });
}
if (elements.reviewReleaseCatalog) {
  elements.reviewReleaseCatalog.addEventListener("change", () => {
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
    if (getSelectedGraphicsMode() === "fix") {
      setActiveModule("gathering");
      setGatheringMode("fix");
      if (!intakeOptionsLoaded) {
        loadGatheringOptions().catch(error => {
          setStatus(`Excerpt gathering option load failed: ${error.message}`);
        });
      }
      return;
    }
    if (elements.graphicsFilter) {
      elements.graphicsFilter.disabled = getSelectedGraphicsMode() !== "queue";
    }
    updateGraphicsModuleTitle();
    syncReleaseCatalogFilterUi();
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
    syncReleaseCatalogFilterUi();
    refreshGraphicsBookSelect(true);
  });
}
if (elements.graphicsReleaseCatalog) {
  elements.graphicsReleaseCatalog.addEventListener("change", () => {
    refreshGraphicsBookSelect(true);
  });
}

async function initializeApp() {
  await loadRuntimeConfig();
  if (elements.gatheringEmail && !elements.gatheringEmail.value.trim()) {
    elements.gatheringEmail.value = getContributorShellPrefillEmail();
  }
  if (elements.gatheringVideoPlaylistUrl && !elements.gatheringVideoPlaylistUrl.value.trim()) {
    elements.gatheringVideoPlaylistUrl.value = getRequestedVideoPlaylistUrl();
  }
  populateReleaseCatalogSelect(elements.reviewReleaseCatalog);
  populateReleaseCatalogSelect(elements.graphicsReleaseCatalog);
  syncGraphicsModeOptionsForView();
  syncReleaseCatalogFilterUi();
  updateGraphicsModuleTitle();
  applyRuntimeMode();
  applyContributorAccessMode();
  updateGatheringModeUi();
  refreshGraphicsFolderImportVisibility();
  renderGraphicsFolderImportPreview(null);
  const contributorShell = currentContributorAccess || isContributorShellRequested();
  setActiveModule(contributorShell ? "gathering" : "review");
  setStatus(`Ready${runtimeConfig.appVersion ? ` (${runtimeConfig.appVersion})` : ""}. Loading books...`);
  if (contributorShell) {
    loadGatheringOptions().catch(error => {
      setStatus(`Excerpt gathering option load failed: ${error.message}`);
    });
  } else {
    loadBooks();
  }
  if (getRequestedVideoPlaylistUrl()) {
    setGatheringMode("video");
    loadGatheringVideoPlaylist({ auto: true });
  }
}

initializeApp();
