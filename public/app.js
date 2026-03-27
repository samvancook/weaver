const storageKey = "weaver.appsScriptUrl";

const elements = {
  apiBaseUrl: document.getElementById("api-base-url"),
  saveApiUrl: document.getElementById("save-api-url"),
  loadBooks: document.getElementById("load-books"),
  bookSelect: document.getElementById("book-select"),
  reviewFilter: document.getElementById("review-filter"),
  scanCorrections: document.getElementById("scan-corrections"),
  loadExcerpts: document.getElementById("load-excerpts"),
  submitReview: document.getElementById("submit-review"),
  excerptList: document.getElementById("excerpt-list"),
  statusOutput: document.getElementById("status-output"),
  bookCountBadge: document.getElementById("book-count-badge"),
  excerptCountBadge: document.getElementById("excerpt-count-badge"),
  correctionScanStatus: document.getElementById("correction-scan-status")
};

let currentExcerpts = [];
let isSaving = false;
let currentValidationByRecordId = new Map();
let correctionCountsByBook = new Map();
let pendingCountsByBook = new Map();
let correctionScanState = "idle";
let loadedBooks = [];

function setStatus(message, details) {
  elements.statusOutput.textContent = details
    ? `${message}\n\n${JSON.stringify(details, null, 2)}`
    : message;
}

function setSubmitState(isBusy, label) {
  isSaving = isBusy;
  elements.submitReview.disabled = isBusy;
  elements.submitReview.textContent = label || (isBusy ? "Saving..." : "Submit Decisions");
}

function updateCorrectionScanUi() {
  const mode = getSelectedReviewFilter();
  const button = elements.scanCorrections;
  const status = elements.correctionScanStatus;
  if (!button || !status) return;

  button.disabled = mode !== "likely_correction" || !loadedBooks.length || correctionScanState === "scanning";
  button.textContent = correctionScanState === "scanning"
    ? "Scanning..."
    : correctionScanState === "ready"
      ? "Rescan Correction Signals"
      : "Scan Correction Signals";

  if (mode !== "likely_correction") {
    status.textContent = "Switch Queue Mode to 'Likely need correction' to scan correction signals.";
  } else if (correctionScanState === "scanning") {
    status.textContent = "Scanning correction signals across loaded books...";
  } else if (correctionScanState === "ready") {
    status.textContent = "Correction scan ready.";
  } else if (correctionScanState === "failed") {
    status.textContent = "Correction scan failed. You can try again.";
  } else {
    status.textContent = "Correction scan not run yet.";
  }
}

function getApiBaseUrl() {
  return elements.apiBaseUrl.value.trim();
}

function saveApiBaseUrl() {
  const value = getApiBaseUrl();
  localStorage.setItem(storageKey, value);
  setStatus("Saved Apps Script URL.");
  if (value) {
    loadBooks();
  }
}

function restoreApiBaseUrl() {
  const saved = localStorage.getItem(storageKey) || "";
  elements.apiBaseUrl.value = saved;
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

async function loadBooks() {
  try {
    setStatus("Loading book titles...");
    const data = await requestJsonp("books");
    if (!data.ok) {
      throw new Error(data.error || "Book load failed.");
    }

    elements.bookSelect.innerHTML = "";
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Choose a book";
    elements.bookSelect.appendChild(placeholder);

    pendingCountsByBook = new Map();
    correctionCountsByBook = new Map();
    correctionScanState = "idle";
    loadedBooks = data.books.slice();

    data.books.forEach(book => {
      pendingCountsByBook.set(book.title, Number(book.count) || 0);
      const option = document.createElement("option");
      option.value = book.title;
      option.textContent = buildBookOptionLabel(book.title);
      elements.bookSelect.appendChild(option);
    });

    elements.bookCountBadge.textContent = `${data.books.length} Books`;
    setStatus(`Loaded ${data.books.length} books. Backend ${data.version || "unknown"}.`, data.books.slice(0, 10));
    refreshBookOptionLabels();
    updateCorrectionScanUi();
  } catch (error) {
    setStatus(`Book load failed: ${error.message}`);
  }
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
    await loadCatalogValidation(data.excerpts);
    renderCurrentExcerpts();
    setStatus(`Loaded ${data.excerpts.length} excerpts for "${bookTitle}". Backend ${data.version || "unknown"}.`);
  } catch (error) {
    setStatus(`Excerpt load failed: ${error.message}`);
  }
}

async function loadCatalogValidation(excerpts) {
  currentValidationByRecordId = new Map();
  if (!excerpts.length) return;

  try {
    const response = await fetch("/api/catalog/validate", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        records: excerpts.map(excerpt => ({
          sourceRow: excerpt.sourceRow,
          recordId: excerpt.recordId,
          author: excerpt.author,
          title: excerpt.title,
          bookTitle: excerpt.bookTitle,
          excerptText: excerpt.excerptText
        }))
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
  elements.excerptList.innerHTML = "";
  elements.excerptCountBadge.textContent = `${excerpts.length} Excerpts`;

  if (!excerpts.length) {
    const empty = document.createElement("p");
    empty.className = "empty-state";
    empty.textContent = getEmptyStateMessage();
    elements.excerptList.appendChild(empty);
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

    elements.excerptList.appendChild(section);
  });
}

function renderCurrentExcerpts() {
  renderExcerpts(applyReviewFilter(currentExcerpts));
}

function buildBookOptionLabel(title) {
  const pendingCount = pendingCountsByBook.get(title) || 0;
  const hasCorrectionCount = correctionCountsByBook.has(title);
  const correctionCount = correctionCountsByBook.get(title) || 0;

  if (getSelectedReviewFilter() === "likely_correction") {
    if (correctionScanState !== "ready") {
      return `${title} (${pendingCount})`;
    }
    return `${title} (${hasCorrectionCount ? correctionCount : 0})`;
  }

  if (correctionScanState === "ready" && correctionCount > 0) {
    return `${title} (${pendingCount}) • ${correctionCount} likely correction${correctionCount === 1 ? "" : "s"}`;
  }

  return `${title} (${pendingCount})`;
}

function refreshBookOptionLabels() {
  Array.from(elements.bookSelect.options).forEach(option => {
    if (!option.value) return;
    option.textContent = buildBookOptionLabel(option.value);
  });
}

async function scanLikelyCorrectionsByBook(books = []) {
  if (!Array.isArray(books) || !books.length) {
    correctionCountsByBook = new Map();
    refreshBookOptionLabels();
    updateCorrectionScanUi();
    return;
  }

  const nextCounts = new Map();

  for (const book of books) {
    const excerptData = await requestJsonp("excerpts", { bookTitle: book.title });
    if (!excerptData.ok || !Array.isArray(excerptData.excerpts) || !excerptData.excerpts.length) {
      continue;
    }

    const response = await fetch("/api/catalog/validate", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        records: excerptData.excerpts.map(excerpt => ({
          sourceRow: excerpt.sourceRow,
          recordId: excerpt.recordId,
          author: excerpt.author,
          title: excerpt.title,
          bookTitle: excerpt.bookTitle,
          excerptText: excerpt.excerptText
        }))
      })
    });
    const validationData = await response.json();
    if (!validationData.ok || !Array.isArray(validationData.results)) {
      continue;
    }

    const correctionCount = validationData.results.filter(result =>
      isLikelyCorrectionStatus(result.status)
    ).length;

    if (correctionCount > 0) {
      nextCounts.set(book.title, correctionCount);
    }
  }

  correctionCountsByBook = nextCounts;
  refreshBookOptionLabels();
  updateCorrectionScanUi();
}

async function runCorrectionScan() {
  if (!loadedBooks.length || correctionScanState === "scanning") return;

  try {
    correctionScanState = "scanning";
    correctionCountsByBook = new Map();
    refreshBookOptionLabels();
    updateCorrectionScanUi();
    await scanLikelyCorrectionsByBook(loadedBooks);
    correctionScanState = "ready";
    refreshBookOptionLabels();
    updateCorrectionScanUi();
  } catch (_error) {
    correctionScanState = "failed";
    refreshBookOptionLabels();
    updateCorrectionScanUi();
  }
}

function getSelectedReviewFilter() {
  return elements.reviewFilter?.value || "all";
}

function applyReviewFilter(excerpts) {
  if (getSelectedReviewFilter() !== "likely_correction") {
    return excerpts;
  }

  return excerpts.filter(excerpt => isLikelyCorrectionExcerpt(excerpt));
}

function isLikelyCorrectionExcerpt(excerpt) {
  const validation =
    currentValidationByRecordId.get(excerpt.recordId || String(excerpt.sourceRow)) || null;

  return Boolean(validation && isLikelyCorrectionStatus(validation.status));
}

function isLikelyCorrectionStatus(status) {
  return [
    "author_mismatch",
    "poem_title_match_only",
    "excerpt_not_found_in_book",
    "book_not_found"
  ].includes(status);
}

function getEmptyStateMessage() {
  if (getSelectedReviewFilter() === "likely_correction") {
    return "No pending excerpts in this book are currently flagged as likely needing correction.";
  }

  return "No pending excerpts remain for this book.";
}

function groupExcerptsByTitle(excerpts) {
  const groups = new Map();

  excerpts.forEach(excerpt => {
    const key = excerpt.title || "Untitled poem";
    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key).push(excerpt);
  });

  return Array.from(groups.entries()).map(([title, groupExcerpts]) => ({
    title,
    excerpts: groupExcerpts
  }));
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

  const validationMarkup = buildValidationMarkup(validation);

  const currentDecision = excerpt.statusIndicator === "NEEDS_CORRECTION"
    ? "needs_correction"
    : excerpt.approved === "Y"
      ? "accept"
      : excerpt.approved === "N"
        ? "reject"
        : "";
  const wordCount = resolveWordCount(excerpt);
  card.innerHTML = `
      <div class="excerpt-card__meta">
        <span class="badge badge--muted">Row ${excerpt.sourceRow}</span>
        <span class="badge badge--muted">ID ${escapeHtml(excerpt.recordId || "none")}</span>
        <span class="badge badge--muted">Book ${escapeHtml(excerpt.bookTitle || "(blank)")}</span>
        <span class="badge badge--count">Words ${wordCount}</span>
        ${pullBadge}
        ${overlapBadge}
        <span class="excerpt-card__title">${escapeHtml(excerpt.title || "Untitled poem")}</span>
      <span class="excerpt-card__author">${escapeHtml(excerpt.author || "Unknown author")}</span>
    </div>
    ${validationMarkup}
    <blockquote class="excerpt-card__quote">${escapeHtml(excerpt.excerptText)}</blockquote>
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
    </div>
    <div class="decision-flags">
      <label><input type="checkbox" class="graphics-qi" ${excerpt.useForGraphicsQi ? "checked" : ""}> Use for Graphics QI</label>
      <label><input type="checkbox" class="photos" ${excerpt.useForPhotos ? "checked" : ""}> Use for Photos</label>
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
      }
    }
  }

  return card;
}

function buildValidationMarkup(validation) {
  if (!validation) return "";

  const canonical = [
    validation.bookCanonicalTitle,
    validation.bookCanonicalAuthor
  ].filter(Boolean).join(" - ");

  if (validation.status === "catalog_match") {
    return `<p class="validation validation--good">Catalog match: ${escapeHtml(canonical || "confirmed")}.</p>`;
  }

  if (validation.status === "author_mismatch") {
    return `<p class="validation validation--warn">Author mismatch. Catalog says ${escapeHtml(canonical || "different author")}.</p>`;
  }

  if (validation.status === "poem_title_match_only") {
    return `<p class="validation validation--warn">Poem title matches this book, but the excerpt text did not match the catalog text.</p>`;
  }

  if (validation.status === "excerpt_not_found_in_book" && validation.globalExcerptMatch) {
    return `<p class="validation validation--warn">Excerpt not found in ${escapeHtml(validation.bookCanonicalTitle || "this book")}. Closest catalog hit: ${escapeHtml(validation.globalExcerptMatch.book_title)} / ${escapeHtml(validation.globalExcerptMatch.poem_title)} by ${escapeHtml(validation.globalExcerptMatch.author)}.</p>`;
  }

  if (validation.status === "book_not_found") {
    return `<p class="validation validation--warn">Book not found in catalog.</p>`;
  }

  return `<p class="validation validation--warn">Catalog check: ${escapeHtml(validation.status)}.</p>`;
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
  return Array.from(elements.excerptList.querySelectorAll(".excerpt-card")).map(card => {
    const approval = card.querySelector('input[type="radio"]:checked')?.value || "";
    return {
      sourceRow: Number(card.dataset.sourceRow),
      recordId: card.dataset.recordId || "",
      approval,
      correctionNote: card.querySelector(".correction-note")?.value || "",
      useForGraphicsQi: card.querySelector(".graphics-qi")?.checked || false,
      useForPhotos: card.querySelector(".photos")?.checked || false
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
      normalizeApprovalForCompare(update.approval) !== normalizeApprovalForCompare(current.approved) ||
      normalizeDecision(update.approval) !== normalizeDecision(current.statusIndicator === "NEEDS_CORRECTION" ? "needs_correction" : current.approved === "Y" ? "accept" : current.approved === "N" ? "reject" : "") ||
      normalizeCorrectionNote(update.correctionNote) !== normalizeCorrectionNote(current.correctionNote) ||
      Boolean(update.useForGraphicsQi) !== Boolean(current.useForGraphicsQi) ||
      Boolean(update.useForPhotos) !== Boolean(current.useForPhotos)
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

    const approvalMatches =
      normalizeApprovalForCompare(update.approval) ===
      normalizeApprovalForCompare(saved.approved);
    const graphicsMatches =
      Boolean(update.useForGraphicsQi) === Boolean(saved.useForGraphicsQi);
    const photosMatches =
      Boolean(update.useForPhotos) === Boolean(saved.useForPhotos);

    if (approvalMatches && graphicsMatches && photosMatches) {
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

async function submitReview() {
  if (!currentExcerpts.length) {
    setStatus("Load excerpts before submitting.");
    return;
  }

  const apiBaseUrl = getApiBaseUrl();
  if (!apiBaseUrl) {
    setStatus("Add the Apps Script Web App URL first.");
    return;
  }

  const updates = collectUpdates();
  const changedUpdates = filterChangedUpdates(updates, currentExcerpts);
  const bookTitle = elements.bookSelect.value;

  if (!changedUpdates.length) {
    setStatus("No changed decisions to submit.");
    return;
  }

  const correctionWithoutNote = changedUpdates.find(update => (
    update.approval === "needs_correction" && !normalizeCorrectionNote(update.correctionNote)
  ));
  if (correctionWithoutNote) {
    setStatus("Add a correction note before submitting a 'Needs correction' decision.");
    return;
  }

  try {
    setSubmitState(true, `Saving ${changedUpdates.length}...`);
    setStatus(`Submitting ${changedUpdates.length} changed review decisions...`);

    for (const update of changedUpdates) {
      const response = await requestJsonp("saveReview", {
        sourceRow: update.sourceRow,
        recordId: update.recordId,
        approval: update.approval,
        correctionNote: update.correctionNote,
        graphicsQi: update.useForGraphicsQi ? "1" : "0",
        photos: update.useForPhotos ? "1" : "0"
      });

      if (!response.ok) {
        throw new Error(response.error || `Save failed for row ${update.sourceRow}.`);
      }
    }

    setStatus(`Submitted ${changedUpdates.length} changed decisions. Verifying saved values...`);
    await sleep(750);

    const refreshed = await requestJsonp("excerpts", { bookTitle });
    if (!refreshed.ok) {
      throw new Error(refreshed.error || "Verification reload failed.");
    }

    currentExcerpts = refreshed.excerpts;
    await loadCatalogValidation(refreshed.excerpts);
    renderCurrentExcerpts();
    await loadBooks();

    const remainingUpdatedRows = countRemainingUpdatedRows(changedUpdates, refreshed.excerpts);
    const droppedFromQueue = changedUpdates.length - remainingUpdatedRows;

    if (remainingUpdatedRows === 0) {
      setStatus(
        `Saved ${changedUpdates.length} changed decisions. Those excerpts dropped out of the pending queue.`
      );
    } else {
      setStatus(
        `Saved ${changedUpdates.length} changed decisions. ${droppedFromQueue} dropped out of queue and ${remainingUpdatedRows} are still showing.`,
        refreshed.excerpts.slice(0, 5)
      );
    }
  } catch (error) {
    setStatus(`Submit failed: ${error.message}`);
  } finally {
    setSubmitState(false, "Submit Decisions");
  }
}

elements.saveApiUrl.addEventListener("click", saveApiBaseUrl);
elements.loadBooks.addEventListener("click", loadBooks);
elements.loadExcerpts.addEventListener("click", loadExcerpts);
elements.submitReview.addEventListener("click", submitReview);
if (elements.reviewFilter) {
  elements.reviewFilter.addEventListener("change", () => {
    refreshBookOptionLabels();
    renderCurrentExcerpts();
    updateCorrectionScanUi();
  });
}
if (elements.scanCorrections) {
  elements.scanCorrections.addEventListener("click", runCorrectionScan);
}

restoreApiBaseUrl();
if (getApiBaseUrl()) {
  setStatus("Ready. Loading books...");
  loadBooks();
} else {
  setStatus("Ready. Save the Apps Script URL, then books will load automatically.");
}
updateCorrectionScanUi();
