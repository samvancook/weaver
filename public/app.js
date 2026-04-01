const storageKey = "weaver.appsScriptUrl";
const runtimeConfig = window.WEAVER_CONFIG || {};

const elements = {
  connectionPanelShell: document.getElementById("connection-panel-shell"),
  apiBaseUrl: document.getElementById("api-base-url"),
  saveApiUrl: document.getElementById("save-api-url"),
  showReviewModule: document.getElementById("show-review-module"),
  showWeirdModule: document.getElementById("show-weird-module"),
  showCorrectionsModule: document.getElementById("show-corrections-module"),
  reviewModule: document.getElementById("review-module"),
  reviewQueuePanel: document.getElementById("review-queue-panel"),
  weirdModule: document.getElementById("weird-module"),
  weirdQueuePanel: document.getElementById("weird-queue-panel"),
  correctionsModule: document.getElementById("corrections-module"),
  loadBooks: document.getElementById("load-books"),
  bookSelect: document.getElementById("book-select"),
  weirdBookSelect: document.getElementById("weird-book-select"),
  reviewFilter: document.getElementById("review-filter"),
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
  correctionList: document.getElementById("correction-list"),
  statusOutput: document.getElementById("status-output"),
  appModeBadge: document.getElementById("app-mode-badge"),
  bookCountBadge: document.getElementById("book-count-badge"),
  excerptCountBadge: document.getElementById("excerpt-count-badge"),
  weirdBookCountBadge: document.getElementById("weird-book-count-badge"),
  weirdExcerptCountBadge: document.getElementById("weird-excerpt-count-badge"),
  correctionBookCountBadge: document.getElementById("correction-book-count-badge"),
  correctionExcerptCountBadge: document.getElementById("correction-excerpt-count-badge")
};

let currentExcerpts = [];
let currentWeirdExcerpts = [];
let currentCorrectionExcerpts = [];
let isSaving = false;
let currentValidationByRecordId = new Map();
let currentModule = "review";
let reviewVisibleCount = 25;
let reviewPinnedRowOrder = [];
let weirdVisibleCount = 25;
let weirdPinnedRowOrder = [];
let currentReviewBookSummaries = [];
let currentWeirdBookSummaries = [];

const REVIEW_BATCH_SIZE = 25;

function setStatus(message, details) {
  elements.statusOutput.textContent = details
    ? `${message}\n\n${JSON.stringify(details, null, 2)}`
    : message;
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
  elements.submitReview.textContent = label || (isBusy ? "Saving..." : "Submit Decisions");
  if (elements.submitWeirdReview) {
    elements.submitWeirdReview.textContent = label || (isBusy ? "Saving..." : "Submit Decisions");
  }
  if (elements.submitCorrections) {
    elements.submitCorrections.textContent = label || (isBusy ? "Saving..." : "Save Corrections");
  }
}

function getApiBaseUrl() {
  return elements.apiBaseUrl.value.trim();
}

function saveApiBaseUrl() {
  if (runtimeConfig.lockApiBaseUrl) {
    setStatus("This hosted Weaver environment is pinned to the production Apps Script backend.");
    return;
  }
  const value = getApiBaseUrl();
  localStorage.setItem(storageKey, value);
  setStatus("Saved Apps Script URL.");
  if (value) {
    loadBooks();
  }
}

function restoreApiBaseUrl() {
  if (runtimeConfig.apiBaseUrl) {
    elements.apiBaseUrl.value = runtimeConfig.apiBaseUrl;
    elements.apiBaseUrl.readOnly = Boolean(runtimeConfig.lockApiBaseUrl);
    if (elements.saveApiUrl) {
      elements.saveApiUrl.disabled = Boolean(runtimeConfig.lockApiBaseUrl);
    }
    return;
  }
  const saved = localStorage.getItem(storageKey) || "";
  elements.apiBaseUrl.value = saved;
}

function applyRuntimeMode() {
  const isLocked = Boolean(runtimeConfig.lockApiBaseUrl);
  if (elements.appModeBadge) {
    elements.appModeBadge.textContent = isLocked ? "Hosted" : "Dev";
  }
  if (elements.connectionPanelShell) {
    elements.connectionPanelShell.hidden = isLocked;
  }
}

function setActiveModule(moduleName) {
  currentModule = ["review", "weird", "corrections"].includes(moduleName) ? moduleName : "review";
  elements.reviewModule?.classList.toggle("module-panel--active", currentModule === "review");
  elements.reviewQueuePanel?.classList.toggle("module-panel--active", currentModule === "review");
  elements.weirdModule?.classList.toggle("module-panel--active", currentModule === "weird");
  elements.weirdQueuePanel?.classList.toggle("module-panel--active", currentModule === "weird");
  elements.correctionsModule?.classList.toggle("module-panel--active", currentModule === "corrections");
  elements.showReviewModule?.classList.toggle("hero-pill--active", currentModule === "review");
  elements.showWeirdModule?.classList.toggle("hero-pill--active", currentModule === "weird");
  elements.showCorrectionsModule?.classList.toggle("hero-pill--active", currentModule === "corrections");
}

function buildJsonpUrl(action, extraParams = {}) {
  const baseUrl = getApiBaseUrl();
  if (!baseUrl) {
    throw new Error("Add the Apps Script Web App URL first.");
  }

  const callbackName = `weaverJsonp_${Date.now()}_${Math.floor(Math.random() * 10000)}`;
  const url = new URL(baseUrl);
  url.searchParams.set("action", action);
  url.searchParams.set("callback", callbackName);

  Object.entries(extraParams).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== "") {
      url.searchParams.set(key, value);
    }
  });

  return { url: url.toString(), callbackName };
}

function requestJsonp(action, params = {}) {
  return new Promise((resolve, reject) => {
    let script;

    try {
      const { url, callbackName } = buildJsonpUrl(action, params);

      window[callbackName] = data => {
        delete window[callbackName];
        if (script && script.parentNode) {
          script.parentNode.removeChild(script);
        }
        resolve(data);
      };

      script = document.createElement("script");
      script.src = url;
      script.onerror = () => {
        delete window[callbackName];
        if (script && script.parentNode) {
          script.parentNode.removeChild(script);
        }
        reject(new Error(`JSONP request failed for ${action}.`));
      };
      document.body.appendChild(script);
    } catch (error) {
      reject(error);
    }
  });
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
    useForInt: update.useForInt ? "1" : "0"
  };
}

async function requestBatchSave(updates) {
  const response = await fetch("/api/save-reviews", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      apiBaseUrl: getApiBaseUrl(),
      updates: updates.map(buildReviewSavePayload)
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
    const response = await requestJsonp("saveReview", buildReviewSavePayload(update));

    if (!response.ok) {
      throw new Error(response.error || `Save failed for row ${update.sourceRow}.`);
    }
  }
}

async function loadBooks(options = {}) {
  const preserveSelection = options.preserveSelection !== false;
  const previousSelection = preserveSelection ? elements.bookSelect.value : "";
  const previousWeirdSelection = preserveSelection ? elements.weirdBookSelect?.value || "" : "";

  try {
    setStatus("Loading book titles...");
    const data = await requestJsonp("pendingRecords");
    if (!data.ok) {
      throw new Error(data.error || "Book load failed.");
    }

    const records = Array.isArray(data.records) ? data.records : [];
    const allBookSummaries = summarizePendingBooks(records);
    currentReviewBookSummaries = allBookSummaries.filter(book => book.totalCount > 0);
    currentWeirdBookSummaries = allBookSummaries.filter(book => book.needsCheckingCount > 0);

    populateBookSelect(elements.bookSelect, currentReviewBookSummaries, previousSelection, book => `${book.title} (${book.totalCount})`);
    populateBookSelect(elements.weirdBookSelect, currentWeirdBookSummaries, previousWeirdSelection, book => `${book.title} (${book.needsCheckingCount})`);

    elements.bookCountBadge.textContent = `${currentReviewBookSummaries.length} Books`;
    if (elements.weirdBookCountBadge) {
      elements.weirdBookCountBadge.textContent = `${currentWeirdBookSummaries.length} Books`;
    }
    setStatus(
      `Loaded ${allBookSummaries.length} books. Backend ${data.version || "unknown"}.`,
      allBookSummaries.slice(0, 10)
    );
  } catch (error) {
    setStatus(`Book load failed: ${error.message}`);
  }
}

function summarizePendingBooks(records) {
  const byTitle = new Map();

  records.forEach(record => {
    const title = (record.bookTitle || "").trim();
    if (!title) return;
    const key = record.recordId || String(record.sourceRow);
    if (record.catalogValidation?.status) {
      currentValidationByRecordId.set(key, record.catalogValidation);
    }

    if (!byTitle.has(title)) {
      byTitle.set(title, {
        title,
        totalCount: 0,
        goodCount: 0,
        needsCheckingCount: 0
      });
    }

    const summary = byTitle.get(title);
    summary.totalCount += 1;
    if (isExtraReviewRecord(record)) {
      summary.needsCheckingCount += 1;
    } else {
      summary.goodCount += 1;
    }
  });

  return Array.from(byTitle.values()).sort((left, right) => left.title.localeCompare(right.title));
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
    option.value = book.title;
    option.textContent = labelBuilder(book);
    select.appendChild(option);
  });

  if (previousSelection) {
    const hasPreviousSelection = books.some(book => book.title === previousSelection);
    if (hasPreviousSelection) {
      select.value = previousSelection;
    }
  }
}

function isGoodValidation(validation) {
  return Boolean(validation && validation.status === "catalog_match");
}

async function loadExcerpts() {
  const bookTitle = elements.bookSelect.value;
  if (!bookTitle) {
    setStatus("Choose a book title first.");
    return;
  }

  try {
    setStatus(`Loading excerpts for "${bookTitle}"...`);
    const data = await requestJsonp("excerpts", { bookTitle });
    if (!data.ok) {
      throw new Error(data.error || "Excerpt load failed.");
    }

    currentExcerpts = data.excerpts;
    reviewVisibleCount = REVIEW_BATCH_SIZE;
    reviewPinnedRowOrder = [];
    await loadCatalogValidation(data.excerpts);
    renderCurrentExcerpts();
    setStatus(`Loaded ${data.excerpts.length} excerpts for "${bookTitle}". Backend ${data.version || "unknown"}.`);
  } catch (error) {
    setStatus(`Excerpt load failed: ${error.message}`);
  }
}

async function loadWeirdExcerpts() {
  const bookTitle = elements.weirdBookSelect?.value;
  if (!bookTitle) {
    setStatus("Choose a book title first.");
    return;
  }

  try {
    setStatus(`Loading extra-review excerpts for "${bookTitle}"...`);
    const data = await requestJsonp("excerpts", { bookTitle });
    if (!data.ok) {
      throw new Error(data.error || "Excerpt load failed.");
    }

    currentWeirdExcerpts = data.excerpts;
    weirdVisibleCount = REVIEW_BATCH_SIZE;
    weirdPinnedRowOrder = [];
    await loadCatalogValidation(data.excerpts);
    renderWeirdCurrentExcerpts();
    setStatus(`Loaded ${applyNeedsCheckingFilter(data.excerpts).length} extra-review excerpts for "${bookTitle}". Backend ${data.version || "unknown"}.`);
  } catch (error) {
    setStatus(`Extra review load failed: ${error.message}`);
  }
}

async function loadCorrectionBooks() {
  const previousSelection = elements.correctionBookSelect?.value || "";

  try {
    setStatus("Loading correction books...");
    const data = await requestJsonp("correctionBooks");
    if (!data.ok) {
      throw new Error(data.error || "Correction book load failed.");
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
    const data = await requestJsonp("corrections", { bookTitle });
    if (!data.ok) {
      throw new Error(data.error || "Correction load failed.");
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
    batchSize: REVIEW_BATCH_SIZE,
    canShowMore: totalMatching > visibleExcerpts.length,
    onShowMore: () => {
      reviewVisibleCount += REVIEW_BATCH_SIZE;
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
    controls.innerHTML = `
      <p class="hint excerpt-batch-controls__hint">Showing ${visibleCount} of ${totalMatching} matching excerpts.</p>
      <button type="button" class="button button--secondary excerpt-batch-controls__button">
        Show next ${options.batchSize || REVIEW_BATCH_SIZE}
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
  const filtered = applyNeedsCheckingFilter(currentWeirdExcerpts);
  const totalMatching = filtered.length;
  const visibleExcerpts = orderWeirdExcerptsForRender(filtered).slice(0, weirdVisibleCount);
  renderExcerptCollection(visibleExcerpts, elements.weirdExcerptList, elements.weirdExcerptCountBadge, getWeirdEmptyStateMessage(), {
    totalMatching,
    visibleCount: visibleExcerpts.length,
    batchSize: REVIEW_BATCH_SIZE,
    canShowMore: totalMatching > visibleExcerpts.length,
    onShowMore: () => {
      weirdVisibleCount += REVIEW_BATCH_SIZE;
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

function applyNeedsCheckingFilter(excerpts) {
  return excerpts.filter(excerpt => isExtraReviewExcerpt(excerpt));
}

function isLikelyCorrectionExcerpt(excerpt) {
  const validation =
    currentValidationByRecordId.get(excerpt.recordId || String(excerpt.sourceRow)) || null;

  return Boolean(validation && isLikelyCorrectionStatus(validation.status));
}

function hasLibraryExcerptMatch(excerpt) {
  const validation =
    currentValidationByRecordId.get(excerpt.recordId || String(excerpt.sourceRow)) || null;

  return Boolean(validation && validation.libraryExcerptMatch);
}

function getLibraryMatchType(excerpt) {
  const validation =
    currentValidationByRecordId.get(excerpt.recordId || String(excerpt.sourceRow)) || null;

  return validation?.libraryExcerptMatch?.matchType || "";
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

function isGoodContentExcerpt(excerpt) {
  const validation =
    currentValidationByRecordId.get(excerpt.recordId || String(excerpt.sourceRow)) || null;

  return Boolean(validation && validation.status === "catalog_match");
}

function isStrictExactLibraryMatch(excerpt) {
  const validation =
    currentValidationByRecordId.get(excerpt.recordId || String(excerpt.sourceRow)) || null;
  const match = validation?.libraryExcerptMatch || null;
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
  const libraryMatch = validation?.libraryExcerptMatch || null;

  if (!validation || !validation.status) {
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

function isStandardReviewExcerpt(excerpt) {
  return !isExtraReviewExcerpt(excerpt);
}

function isExtraReviewRecord(record) {
  const validation = record.catalogValidation || null;
  const libraryMatch = validation?.libraryExcerptMatch || null;

  if (!validation || !validation.status) {
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
  const libraryBadgeLabel = getLibraryBadgeLabel(libraryMatch);
  const libraryBadge = libraryMatch
    ? `<span class="badge badge--warn">${escapeHtml(libraryBadgeLabel)}</span>`
    : "";
  const libraryRecommendation = libraryMatch?.matchType === "exact"
    ? `<span class="badge badge--muted">Suggested: Reject</span>`
    : "";

  const validationMarkup = buildValidationMarkup(validation, excerpt);
  const reviewDecision = normalizeDecision(excerpt.excerptReviewDecision);
  const useForQi = Boolean(excerpt.useForQi ?? excerpt.useForGraphicsQi);
  const useForInt = Boolean(excerpt.useForInt ?? excerpt.useForPhotos);
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
      <span class="excerpt-card__author">${escapeHtml(excerpt.author || "Unknown author")}</span>
    </div>
    ${validationMarkup}
    <blockquote class="excerpt-card__quote">${escapeHtml(excerpt.excerptText)}</blockquote>
    <p class="hint excerpt-card__hint">Accept/Reject/Needs correction controls excerpt review. Use for QI preserves the existing quote-image tool linkage. Made + QCed is downstream status only.</p>
    <div class="decision-group">
      <label><input type="radio" name="approval-${uniqueKey}" value="accept" ${currentDecision === "accept" ? "checked" : ""}> Accept</label>
      <label><input type="radio" name="approval-${uniqueKey}" value="reject" ${currentDecision === "reject" ? "checked" : ""}> Reject</label>
      <label><input type="radio" name="approval-${uniqueKey}" value="needs_correction" ${currentDecision === "needs_correction" ? "checked" : ""}> Needs correction</label>
      <label><input type="radio" name="approval-${uniqueKey}" value="" ${currentDecision === "" ? "checked" : ""}> No decision</label>
    </div>
    <div class="correction-block ${currentDecision === "needs_correction" ? "correction-block--active" : ""}">
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
    <div class="decision-flags">
      <label><input type="checkbox" class="graphics-qi" ${useForQi ? "checked" : ""}> Use for QI</label>
      <label><input type="checkbox" class="use-for-int" ${useForInt ? "checked" : ""}> Use for INT</label>
    </div>
  `;

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

function buildValidationMarkup(validation, excerpt) {
  if (!validation) {
    return `<p class="validation validation--pending">Catalog validation unavailable for this excerpt.</p>`;
  }

  const canonical = [
    validation.bookCanonicalTitle,
    validation.bookCanonicalAuthor
  ].filter(Boolean).join(" - ");
  const libraryMarkup = buildLibraryMatchMarkup(validation.libraryExcerptMatch);
  const poemLink = buildCatalogPoemLink(validation, excerpt);

  if (validation.status === "catalog_match") {
    return `<p class="validation validation--good">Catalog match: ${escapeHtml(canonical || "confirmed")}. ${poemLink}</p>${libraryMarkup}`;
  }

  if (validation.status === "author_mismatch") {
    return `<p class="validation validation--warn">Author mismatch. Catalog says ${escapeHtml(canonical || "different author")}. ${poemLink}</p>${libraryMarkup}`;
  }

  if (validation.status === "title_mismatch") {
    return `<p class="validation validation--warn">Excerpt matches the catalog, but the poem title appears wrong. Catalog match: ${escapeHtml(validation.matchedPoemTitle || "different title")}. ${poemLink}</p>${libraryMarkup}`;
  }

  if (validation.status === "poem_title_match_only") {
    return `<p class="validation validation--warn">Poem title matches this book, but the excerpt text did not match the catalog text. ${poemLink}</p>${libraryMarkup}`;
  }

  if (validation.status === "epub_not_present") {
    return `<p class="validation validation--warn">This title appears to be intentionally absent from EPUB/catalog coverage, not simply mismatched.</p>${libraryMarkup}`;
  }

  if (validation.status === "excerpt_not_found_in_book" && validation.globalExcerptMatch) {
    return `<p class="validation validation--warn">Excerpt not found in ${escapeHtml(validation.bookCanonicalTitle || "this book")}. Closest catalog hit: ${escapeHtml(validation.globalExcerptMatch.book_title)} / ${escapeHtml(validation.globalExcerptMatch.poem_title)} by ${escapeHtml(validation.globalExcerptMatch.author)}. ${poemLink}</p>${libraryMarkup}`;
  }

  if (validation.status === "book_not_found") {
    return `<p class="validation validation--warn">Book not found in catalog.</p>${libraryMarkup}`;
  }

  return `<p class="validation validation--warn">Catalog check: ${escapeHtml(validation.status)}.</p>${libraryMarkup}`;
}

function buildCatalogPoemLink(validation, excerpt) {
  let bookTitle = validation.bookCanonicalTitle || excerpt.bookTitle || "";
  let poemTitle = validation.matchedPoemTitle || excerpt.title || "";

  if (validation.status === "excerpt_not_found_in_book" && validation.globalExcerptMatch) {
    bookTitle = validation.globalExcerptMatch.book_title || bookTitle;
    poemTitle = validation.globalExcerptMatch.poem_title || poemTitle;
  }

  if (!bookTitle || !poemTitle) {
    return "";
  }

  const url = new URL("/catalog-poem", window.location.origin);
  url.searchParams.set("bookTitle", bookTitle);
  url.searchParams.set("poemTitle", poemTitle);
  return `<button type="button" class="catalog-poem-link button-link" data-popup-url="${escapeHtml(url.toString())}">View catalog poem</button>`;
}

function buildLibraryMatchMarkup(match) {
  if (!match) {
    return "";
  }

  const label = getLibraryMatchSentence(match);
  const meta = [
    match.bookTitle,
    match.poemTitle,
    match.author
  ].filter(Boolean).join(" / ");
  const excerptLink = buildLibraryExcerptLink(match);
  return `<p class="validation validation--warn">${escapeHtml(label)} ${escapeHtml(meta || "Existing source row")}${match.sourceRow ? `, row ${escapeHtml(String(match.sourceRow))}` : ""}. ${excerptLink}</p>`;
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

function collectUpdates() {
  return collectUpdatesFromContainer(elements.excerptList);
}

function collectCorrectionUpdates() {
  return collectUpdatesFromContainer(elements.correctionList);
}

function collectWeirdUpdates() {
  return collectUpdatesFromContainer(elements.weirdExcerptList);
}

function collectUpdatesFromContainer(container) {
  return Array.from(container.querySelectorAll(".excerpt-card")).map(card => {
    const reviewDecision = card.querySelector('input[type="radio"]:checked')?.value || "";
    return {
      sourceRow: Number(card.dataset.sourceRow),
      recordId: card.dataset.recordId || "",
      reviewDecision,
      correctionNote: card.querySelector(".correction-note")?.value || "",
      correctedAuthor: card.querySelector(".corrected-author")?.value || "",
      correctedTitle: card.querySelector(".corrected-title")?.value || "",
      correctedBookTitle: card.querySelector(".corrected-book-title")?.value || "",
      correctedExcerpt: card.querySelector(".corrected-excerpt")?.value || "",
      useForQi: card.querySelector(".graphics-qi")?.checked || false,
      useForInt: card.querySelector(".use-for-int")?.checked || false
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
      normalizeDecision(update.reviewDecision) !== normalizeDecision(current.excerptReviewDecision) ||
      normalizeCorrectionNote(update.correctionNote) !== normalizeCorrectionNote(current.correctionNote) ||
      normalizeCorrectionNote(update.correctedAuthor) !== normalizeCorrectionNote(current.correctedAuthor) ||
      normalizeCorrectionNote(update.correctedTitle) !== normalizeCorrectionNote(current.correctedTitle) ||
      normalizeCorrectionNote(update.correctedBookTitle) !== normalizeCorrectionNote(current.correctedBookTitle) ||
      normalizeCorrectionNote(update.correctedExcerpt) !== normalizeCorrectionNote(current.correctedExcerpt) ||
      Boolean(update.useForQi) !== Boolean(current.useForQi ?? current.useForGraphicsQi) ||
      Boolean(update.useForInt) !== Boolean(current.useForInt ?? current.useForPhotos)
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
      normalizeDecision(update.reviewDecision) === normalizeDecision(saved.excerptReviewDecision);
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
    const intMatches =
      Boolean(update.useForInt) === Boolean(saved.useForInt ?? saved.useForPhotos);

    if (decisionMatches && authorMatches && titleMatches && bookMatches && excerptMatches && qiMatches && intMatches) {
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

  return submitExcerptSet({
    sourceExcerpts: currentExcerpts,
    collectUpdates: collectUpdates,
    bookTitle: elements.bookSelect.value,
    reloadAction: "excerpts",
    afterReload: async refreshed => {
      currentExcerpts = refreshed.excerpts;
      const remainingSourceRows = new Set(
        refreshed.excerpts.map(excerpt => Number(excerpt.sourceRow))
      );
      reviewPinnedRowOrder = pinnedSourceRows.filter(sourceRow => remainingSourceRows.has(sourceRow));
      await loadCatalogValidation(refreshed.excerpts);
      renderCurrentExcerpts();
      loadBooks({ preserveSelection: true }).catch(() => {});
    },
    countExcerpts: refreshed => applyReviewFilter(refreshed.excerpts),
    emptyMessage: "Load excerpts before submitting.",
    idleLabel: "Submit Decisions",
    progressLabel: "Submitting"
  });
}

async function submitWeirdReview() {
  const pinnedSourceRows = Array.from(
    elements.weirdExcerptList.querySelectorAll(".excerpt-card")
  ).map(card => Number(card.dataset.sourceRow)).filter(Boolean);

  return submitExcerptSet({
    sourceExcerpts: currentWeirdExcerpts,
    collectUpdates: collectWeirdUpdates,
    bookTitle: elements.weirdBookSelect.value,
    reloadAction: "excerpts",
    afterReload: async refreshed => {
      currentWeirdExcerpts = refreshed.excerpts;
      const remainingSourceRows = new Set(
        applyNeedsCheckingFilter(refreshed.excerpts).map(excerpt => Number(excerpt.sourceRow))
      );
      weirdPinnedRowOrder = pinnedSourceRows.filter(sourceRow => remainingSourceRows.has(sourceRow));
      await loadCatalogValidation(refreshed.excerpts);
      renderWeirdCurrentExcerpts();
      loadBooks({ preserveSelection: true }).catch(() => {});
    },
    countExcerpts: refreshed => applyNeedsCheckingFilter(refreshed.excerpts),
    emptyMessage: "Load needs-checking excerpts before submitting.",
    idleLabel: "Submit Decisions",
    progressLabel: "Submitting"
  });
}

async function submitCorrections() {
  return submitExcerptSet({
    sourceExcerpts: currentCorrectionExcerpts,
    collectUpdates: collectCorrectionUpdates,
    bookTitle: elements.correctionBookSelect.value,
    reloadAction: "corrections",
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
  bookTitle,
  reloadAction,
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

  const apiBaseUrl = getApiBaseUrl();
  if (!apiBaseUrl) {
    setStatus("Add the Apps Script Web App URL first.");
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

    setStatus(`Submitted ${changedUpdates.length} changed decisions. Verifying saved values...`);
    await sleep(750);

    const refreshed = await requestJsonp(reloadAction, { bookTitle });
    if (!refreshed.ok) {
      throw new Error(refreshed.error || "Verification reload failed.");
    }

    await afterReload(refreshed);
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
        refreshed.excerpts.slice(0, 5)
      );
    }
  } catch (error) {
    setStatus(`Submit failed: ${error.message}`);
  } finally {
    setSubmitState(false, idleLabel);
  }
}

elements.saveApiUrl.addEventListener("click", saveApiBaseUrl);
elements.loadBooks.addEventListener("click", loadBooks);
elements.loadExcerpts.addEventListener("click", loadExcerpts);
elements.loadWeirdExcerpts?.addEventListener("click", loadWeirdExcerpts);
elements.submitReview.addEventListener("click", submitReview);
elements.submitWeirdReview?.addEventListener("click", submitWeirdReview);
elements.loadCorrectionBooks?.addEventListener("click", loadCorrectionBooks);
elements.loadCorrections?.addEventListener("click", loadCorrections);
elements.submitCorrections?.addEventListener("click", submitCorrections);
elements.showReviewModule?.addEventListener("click", () => setActiveModule("review"));
elements.showWeirdModule?.addEventListener("click", () => {
  setActiveModule("weird");
  if (getApiBaseUrl() && elements.weirdBookSelect?.options.length <= 1) {
    loadBooks();
  }
});
elements.showCorrectionsModule?.addEventListener("click", () => {
  setActiveModule("corrections");
  if (getApiBaseUrl() && elements.correctionBookSelect?.options.length <= 1) {
    loadCorrectionBooks();
  }
});
if (elements.reviewFilter) {
  elements.reviewFilter.addEventListener("change", () => {
    reviewVisibleCount = REVIEW_BATCH_SIZE;
    reviewPinnedRowOrder = [];
    renderCurrentExcerpts();
  });
}

restoreApiBaseUrl();
applyRuntimeMode();
setActiveModule("review");
if (getApiBaseUrl()) {
  setStatus(`Ready${runtimeConfig.appVersion ? ` (${runtimeConfig.appVersion})` : ""}. Loading books...`);
  loadBooks();
} else {
  setStatus("Ready. Save the Apps Script URL, then books will load automatically.");
}
