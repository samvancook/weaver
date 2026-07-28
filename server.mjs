import http from "node:http";
import { spawn } from "node:child_process";
import { createHash, randomUUID } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const publicDir = path.join(__dirname, "public");
const port = Number(process.env.PORT || 8080);
const appVersion =
  process.env.K_REVISION ||
  process.env.WEAVER_APP_VERSION ||
  "dev-local";
const spreadsheetId =
  process.env.WEAVER_SPREADSHEET_ID ||
  "1yTCRQKAavimDEJka1-Ice4xlJ1mCm8hq0-G1PTQkTLM";
const sourceSheetName =
  process.env.WEAVER_SOURCE_SHEET_NAME ||
  "Excerpt Tool 1.20";
const graphicsReviewSheetName =
  process.env.WEAVER_GRAPHICS_REVIEW_SHEET_NAME ||
  "New - Quote Creation Tool Database";
const graphicsQueueSheetName =
  process.env.WEAVER_GRAPHICS_QUEUE_SHEET_NAME ||
  "Queue - Needs Graphics";
const graphicsCleanupSheetName =
  process.env.WEAVER_GRAPHICS_CLEANUP_SHEET_NAME ||
  "Cleanup - Created Graphics";
const pigCompletedGraphicsSheetName =
  process.env.WEAVER_PIG_COMPLETIONS_SHEET_NAME ||
  "PIG - Completed Graphics";
const publishingOrderSpreadsheetId =
  process.env.WEAVER_PUBLISHING_ORDER_SPREADSHEET_ID ||
  "1eG0llCHGRukTtxi7SS-VUoSez7g4MfECqejKcyLyrR0";
const publishingOrderSheetName =
  process.env.WEAVER_PUBLISHING_ORDER_SHEET_NAME ||
  "Publishing Order & Shorthand";
const defaultGoogleOAuthClientId =
  process.env.WEAVER_GOOGLE_OAUTH_CLIENT_ID ||
  "912447899335-a7uuvddqb8g1llm7bt3ov5rao0ie13qf.apps.googleusercontent.com";
const poetryPleaseApiUrl =
  process.env.POETRY_PLEASE_API_URL ||
  "https://poetryplease.org/api";
const poetryPleaseApiKey =
  String(process.env.POETRY_PLEASE_API_KEY || "").trim();
const weaverPublicBaseUrl =
  cleanSheetWhitespace(process.env.WEAVER_PUBLIC_BASE_URL) ||
  "https://weaver.buttonpoetry.com";
const fallbackReviewQueueIncludeTitles = [
  "A Choir of Honest Killers",
  "all the ugly bits",
  "Coin Laundry at Midnight",
  "DON’T BE AFRAID TO BE BAD: A Big Book of Button Poetry Writing Prompts",
  "Flee",
  "Living at Baggage Claim",
  "My Dear Cult Leader",
  "Pink Elephant",
  "Roads",
  "Stunt Water: The Work of Buddy Wakefield",
  "Tooth Gaps in the Archives",
  "good luck in the real world",
  "without the frills"
];

const contentTypes = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".svg": "image/svg+xml"
};
const SHEET_SOURCE_CONFIG = {
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
    validationPrimarySourceFormat: 61,
    excerptPoetryPleaseStatus: 62,
    excerptPoetryPleaseUpdatedAt: 63,
    excerptPoetryPleaseNote: 64
  }
};
const PIG_COMPLETION_COLUMNS = {
  completionId: 1,
  requestId: 2,
  author: 3,
  poemTitle: 4,
  bookTitle: 5,
  quoteText: 6,
  assetUrl: 7,
  assetPreviewUrl: 8,
  sourceRecordId: 9,
  sourceSheetRow: 10,
  productionNotes: 11,
  completedAt: 12,
  qcDecision: 13,
  qcNote: 14,
  qcUpdatedAt: 15,
  sourceTool: 16,
  poetryPleaseStatus: 17,
  poetryPleaseUpdatedAt: 18,
  poetryPleaseNote: 19
};
const QUEUE_CACHE_TTL_MS = 5000;
const queueSnapshotCache = new Map();
const SOURCE_SHEET_CACHE_KEY = "source-sheet-values";

async function getCachedQueueSnapshot(key, loader, ttlMs = QUEUE_CACHE_TTL_MS) {
  const now = Date.now();
  const current = queueSnapshotCache.get(key);
  if (current && current.expiresAt > now) {
    return current.value;
  }
  if (current?.promise) {
    return current.promise;
  }

  queueSnapshotCache.set(key, {
    value: current?.value,
    expiresAt: current?.expiresAt || 0,
    promise: null
  });
  const entry = queueSnapshotCache.get(key);
  entry.promise = Promise.resolve()
    .then(loader)
    .then(value => {
      if (queueSnapshotCache.get(key) === entry) {
        queueSnapshotCache.set(key, {
          value,
          expiresAt: Date.now() + ttlMs,
          promise: null
        });
      }
      return value;
    })
    .catch(error => {
      if (queueSnapshotCache.get(key) === entry) {
        queueSnapshotCache.delete(key);
      }
      throw error;
    });
  return entry.promise;
}

function invalidateQueueSnapshots() {
  queueSnapshotCache.clear();
}

async function getSourceSheetValuesCached() {
  await ensureExcerptPoetryPleaseColumnsServer();
  const range = `'${sourceSheetName.replace(/'/g, "''")}'!A2:BL`;
  return getCachedQueueSnapshot(SOURCE_SHEET_CACHE_KEY, () => fetchSheetValuesServer(range));
}
const PIG_COMPLETION_HEADERS = [[
  "completion_id",
  "request_id",
  "author",
  "poem_title",
  "book_title",
  "quote_text",
  "asset_url",
  "asset_preview_url",
  "source_record_id",
  "source_sheet_row",
  "production_notes",
  "completed_at",
  "graphic_qc_decision",
  "graphic_qc_note",
  "graphic_qc_updated_at",
  "source_tool",
  "poetry_please_handoff_status",
  "poetry_please_handoff_updated_at",
  "poetry_please_handoff_note"
]];
const GRAPHICS_QC_REJECT_REASON_LABELS = new Map([
  ["mismatched_graphic", "Mismatched graphic"],
  ["correct_and_recreate", "Correct and recreate"],
  ["final_reject", "Final reject"]
]);

function cleanSheetWhitespace(text) {
  return String(text || "").replace(/\s+/g, " ").trim();
}

function slugToken(text) {
  return cleanSheetWhitespace(text)
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/['’]/g, "")
    .replace(/[^A-Za-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .toUpperCase();
}

function normalizeExcerptTransferText(text) {
  return String(text || "").replace(/\r\n?/g, "\n").trim();
}

function canonicalFullPoemSourceRecordId(record) {
  const bookShortener = cleanSheetWhitespace(record?.bookShortener);
  const poemTitle = cleanSheetWhitespace(record?.poemTitle || record?.title);
  if (!bookShortener || !poemTitle) return "";
  return `${slugToken(bookShortener)}-FP-${slugToken(poemTitle)}`;
}

function normalizeExcerptContentType(value) {
  const cleaned = cleanSheetWhitespace(value).toUpperCase();
  if (cleaned === "FP" || cleaned === "FULL POEM") return "FP";
  if (
    cleaned === "FPI"
    || cleaned === "FULL POEM IMAGE"
    || cleaned === "FULL POEM IMAGE-BACKED"
    || cleaned === "FULL POEM IMAGE BACKED"
  ) return "FPI";
  return "EXC";
}

function normalizePoetryPleaseContentType(value, fallback = "EXC") {
  const cleaned = cleanSheetWhitespace(value).toUpperCase();
  if (cleaned === "QI" || cleaned === "QUOTE IMAGE") return "QI";
  if (cleaned === "FP" || cleaned === "FULL POEM") return "FP";
  if (
    cleaned === "FPI"
    || cleaned === "FULL POEM IMAGE"
    || cleaned === "FULL POEM IMAGE-BACKED"
    || cleaned === "FULL POEM IMAGE BACKED"
  ) return "FPI";
  return fallback;
}

function inferPoetryPleaseContentType(record, fallback = "QI") {
  const explicitType = normalizePoetryPleaseContentType(record?.contentType || record?.imageType, fallback);
  const ids = [
    record?.graphicsRequestId,
    record?.sourceRequestId,
    record?.sourceRecordId,
    record?.recordId
  ].map(cleanSheetWhitespace);
  if (explicitType === "QI" && ids.some(id => /(^|[-:])FP[-:]/i.test(id))) {
    return "FPI";
  }
  return explicitType;
}

function parseIntakeMetadataFromNotes(text) {
  const noteText = String(text || "");
  const result = {
    releaseCatalog: "",
    bookShortener: "",
    contentType: "",
    socialMediaHandle: "",
    sourceEvent: ""
  };
  noteText.split(/\r?\n/).forEach(line => {
    const match = line.match(/^\s*([^:]+):\s*(.+?)\s*$/);
    if (!match) return;
    const label = cleanSheetWhitespace(match[1]).toLowerCase();
    const value = cleanSheetWhitespace(match[2]);
    if (!value) return;
    if (label === "release catalog" || label === "catalog") {
      result.releaseCatalog = value;
    } else if (label === "book shortener" || label === "shortener") {
      result.bookShortener = value;
    } else if (label === "content type" || label === "type") {
      result.contentType = normalizeExcerptContentType(value);
    } else if (label === "project/event" || label === "source event" || label === "event") {
      result.sourceEvent = value;
    } else if (
      label === "instagram handle"
      || label === "ig handle"
      || label === "social media handle"
      || label === "instagram"
      || label === "handle"
    ) {
      result.socialMediaHandle = value;
    }
  });
  return result;
}

function decodeUnicodeEscapes(text) {
  return String(text || "").replace(/\\u([0-9a-fA-F]{4})/g, (_match, hex) => {
    try {
      return String.fromCharCode(parseInt(hex, 16));
    } catch {
      return "";
    }
  });
}

function decodeHtmlEntities(text) {
  return String(text || "")
    .replace(/&amp;/g, "&")
    .replace(/&#39;/g, "'")
    .replace(/&quot;/g, "\"")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">");
}

function decodeJsString(text) {
  return decodeHtmlEntities(
    decodeUnicodeEscapes(String(text || ""))
      .replace(/\\"/g, "\"")
      .replace(/\\\\/g, "\\")
  );
}

function stripHtmlTags(text) {
  return String(text || "").replace(/<[^>]+>/g, " ");
}

function normalizeVideoPlaylistLabel(text) {
  return decodeHtmlEntities(decodeUnicodeEscapes(String(text || "")))
    .replace(/\s+/g, " ")
    .trim();
}

function derivePlaylistEventName(formTitle) {
  const cleaned = normalizeVideoPlaylistLabel(formTitle);
  return cleaned
    .replace(/^Curation Form\s*-?\s*/i, "")
    .replace(/^Vertical Versions\s*-?\s*/i, "")
    .trim();
}

function parseVideoQuestionLabel(labelText) {
  const cleaned = normalizeVideoPlaylistLabel(labelText);
  if (!cleaned.includes("[Vertical Version]")) return null;
  const versionIndex = cleaned.indexOf("[Vertical Version]");
  const labelStart = versionIndex >= 0 ? versionIndex + "[Vertical Version]".length : 0;
  const movIndex = cleaned.indexOf(".mov");
  if (movIndex < 0) return null;
  const coreLabel = normalizeVideoPlaylistLabel(cleaned.slice(labelStart, movIndex));
  const separatorIndex = coreLabel.indexOf(" - ");
  if (separatorIndex < 0) return null;
  const author = normalizeVideoPlaylistLabel(coreLabel.slice(0, separatorIndex));
  const poemTitle = normalizeVideoPlaylistLabel(coreLabel.slice(separatorIndex + 3));
  const videoUrlMatch = cleaned.match(/https:\/\/drive\.google\.com\/file\/d\/[A-Za-z0-9_-]+\/view\?usp=drivesdk/i);
  const videoUrl = String(videoUrlMatch?.[0] || "").trim();
  if (!author || !poemTitle || !videoUrl) return null;
  return { author, poemTitle, videoUrl };
}

function extractGoogleFormResponseUrl(html) {
  const formIdMatch = String(html || "").match(/"e\/1FAIpQL[^"]+"/);
  if (!formIdMatch?.[0]) return "";
  const formId = formIdMatch[0].slice(1, -1);
  return `https://docs.google.com/forms/d/${formId}/formResponse`;
}

function parseVideoPlaylistItemsFromFormHtml(html, sourceUrl) {
  const decodedHtml = decodeUnicodeEscapes(String(html || ""));
  const titleMatch = String(html || "").match(/<title>([^<]+)<\/title>/i);
  const formTitle = normalizeVideoPlaylistLabel(titleMatch?.[1] || "");
  const eventName = derivePlaylistEventName(formTitle);
  const formResponseUrl = extractGoogleFormResponseUrl(html);
  const items = [];
  const seen = new Set();
  const questionRecords = [];
  const normalizedQuestionSource = decodeHtmlEntities(decodedHtml);
  const questionPattern = /\[(\d+),"((?:[^"\\]|\\.)*)",(?:null|"((?:[^"\\]|\\.)*)"),([01]),\[\[(\d+)(?:,[^\]]*)?\]\]/g;
  let questionMatch;
  while ((questionMatch = questionPattern.exec(normalizedQuestionSource))) {
    questionRecords.push({
      questionId: cleanSheetWhitespace(questionMatch[1]),
      label: normalizeVideoPlaylistLabel(stripHtmlTags(decodeJsString(questionMatch[2] || ""))),
      required: questionMatch[4] === "0",
      entryId: cleanSheetWhitespace(questionMatch[5])
    });
  }

  const byCompositeKey = new Map();
  for (let index = 0; index < questionRecords.length; index += 1) {
    const record = questionRecords[index];
    if (!record.label.includes("[Vertical Version]")) continue;
    const noteRecord = questionRecords[index + 1];
    const parsed = parseVideoQuestionLabel(record.label);
    if (!parsed) continue;
    const compositeKey = `${parsed.author}||${parsed.poemTitle}||${parsed.videoUrl}`;
    byCompositeKey.set(compositeKey, {
      scoreEntryId: record.entryId,
      notesEntryId: noteRecord?.label?.startsWith("Additional Notes - ") ? noteRecord.entryId : "",
      formResponseUrl
    });
  }
  const spanPattern = /<span class="M7eMe">([\s\S]*?)<\/span>/g;
  let match;
  while ((match = spanPattern.exec(decodedHtml))) {
    const spanText = normalizeVideoPlaylistLabel(stripHtmlTags(match[1] || ""));
    if (!spanText.includes("[Vertical Version]") || !spanText.includes("drive.google.com/file/d/")) {
      continue;
    }
    const versionIndex = spanText.indexOf("[Vertical Version]");
    const labelStart = versionIndex >= 0 ? versionIndex + "[Vertical Version]".length : 0;
    const movIndex = spanText.indexOf(".mov");
    if (movIndex < 0) continue;
    const label = normalizeVideoPlaylistLabel(spanText.slice(labelStart, movIndex));
    const videoUrlMatch = spanText.match(/https:\/\/drive\.google\.com\/file\/d\/[A-Za-z0-9_-]+\/view\?usp=drivesdk/i);
    const videoUrl = String(videoUrlMatch?.[0] || "").trim();
    if (!label || !videoUrl) continue;
    const separatorIndex = label.indexOf(" - ");
    if (separatorIndex < 0) continue;
    const author = normalizeVideoPlaylistLabel(label.slice(0, separatorIndex));
    const poemTitle = normalizeVideoPlaylistLabel(label.slice(separatorIndex + 3));
    if (!author || !poemTitle) continue;
    const responseMeta = byCompositeKey.get(`${author}||${poemTitle}||${videoUrl}`) || {};
    const key = `${author}||${poemTitle}||${videoUrl}`;
    if (seen.has(key)) continue;
    seen.add(key);
    items.push({
      author,
      poemTitle,
      videoUrl,
      eventName,
      formResponseUrl: responseMeta.formResponseUrl || formResponseUrl,
      scoreEntryId: responseMeta.scoreEntryId || "",
      notesEntryId: responseMeta.notesEntryId || ""
    });
  }
  return {
    ok: true,
    sourceUrl,
    formTitle,
    eventName,
    formResponseUrl,
    count: items.length,
    items
  };
}

function validateVideoPlaylistScore(scoreText) {
  const raw = String(scoreText || "").trim();
  if (!raw) {
    throw new Error("Curation score is required.");
  }
  const score = Number(raw);
  if (!Number.isFinite(score) || score < 0 || score > 10) {
    throw new Error("Curation score must be a number between 0 and 10.");
  }
  return raw;
}

async function submitVideoPlaylistScore(payload) {
  const formResponseUrl = String(payload?.formResponseUrl || "").trim();
  const scoreEntryId = cleanSheetWhitespace(payload?.scoreEntryId);
  const notesEntryId = cleanSheetWhitespace(payload?.notesEntryId);
  const score = validateVideoPlaylistScore(payload?.score);
  const notes = String(payload?.notes || "").trim();
  if (!formResponseUrl || !scoreEntryId) {
    throw new Error("Missing playlist scoring metadata.");
  }
  const formData = new URLSearchParams();
  formData.set(`entry.${scoreEntryId}`, score);
  if (notesEntryId && notes) {
    formData.set(`entry.${notesEntryId}`, notes);
  }
  formData.set("fvv", "1");
  formData.set("pageHistory", "0");
  const response = await fetch(formResponseUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
      "User-Agent": "Weaver video playlist scorer"
    },
    body: formData.toString(),
    redirect: "manual"
  });
  if (!(response.status >= 200 && response.status < 400)) {
    throw new Error(`Form score submit failed with ${response.status}.`);
  }
  return {
    ok: true,
    formResponseUrl,
    scoreEntryId,
    notesEntryId,
    score,
    notes
  };
}

async function loadVideoPlaylistFromFormUrl(formUrl) {
  const sourceUrl = String(formUrl || "").trim();
  if (!sourceUrl) {
    throw new Error("formUrl is required.");
  }
  let parsedUrl;
  try {
    parsedUrl = new URL(sourceUrl);
  } catch {
    throw new Error("Invalid form URL.");
  }
  const allowedHost = /(^|\.)forms\.gle$|(^|\.)docs\.google\.com$/i.test(parsedUrl.hostname);
  if (!allowedHost) {
    throw new Error("Only Google Form URLs are supported.");
  }
  const response = await fetch(sourceUrl, {
    redirect: "follow",
    headers: {
      "User-Agent": "Weaver video playlist loader"
    }
  });
  if (!response.ok) {
    throw new Error(`Form fetch failed with ${response.status}.`);
  }
  const html = await response.text();
  const result = parseVideoPlaylistItemsFromFormHtml(html, sourceUrl);
  if (!result.items.length) {
    throw new Error("No playable video entries were found in that form.");
  }
  return result;
}

function getBookBaseTitle(text) {
  const cleanedTitle = cleanSheetWhitespace(text);
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

function normalizeBookKey(text) {
  return getBookBaseTitle(text).toLowerCase();
}

const CONTRIBUTOR_INVITES = [
  {
    email: "samvancook@gmail.com",
    role: "contributor",
    allowedBooks: ["without the frills"]
  },
  {
    email: "gigibellag@gmail.com",
    role: "contributor",
    allowedBooks: ["without the frills"]
  }
];

function getContributorInviteByEmail(email) {
  const cleanedEmail = cleanSheetWhitespace(email).toLowerCase();
  if (!cleanedEmail) return null;
  return CONTRIBUTOR_INVITES.find(invite => invite.email.toLowerCase() === cleanedEmail) || null;
}

function assertContributorIntakeAccess(payload = {}) {
  const invite = getContributorInviteByEmail(payload.email);
  if (!invite) {
    return;
  }

  const mode = cleanSheetWhitespace(payload.mode).toLowerCase();
  if (mode !== "book") {
    throw new Error("This contributor invite is limited to book excerpt gathering only.");
  }

  const allowedBooks = Array.isArray(invite.allowedBooks) ? invite.allowedBooks : [];
  const allowedBookKeys = new Set(allowedBooks.map(normalizeBookKey).filter(Boolean));
  const submittedBookKey = normalizeBookKey(payload.bookTitle);
  if (!submittedBookKey || !allowedBookKeys.has(submittedBookKey)) {
    throw new Error(`This contributor invite is limited to: ${allowedBooks.join(", ")}.`);
  }
}

function buildReviewQueueIncludeSetFromBooks(books = []) {
  return new Set(
    books
      .filter(book => cleanSheetWhitespace(book?.releaseCatalog) && cleanSheetWhitespace(book?.bookShortener))
      .map(book => normalizeBookKey(book.title))
      .filter(Boolean)
  );
}

async function getReleaseCatalogQueueBooks() {
  const books = await getPublishingOrderBooks();
  return books.filter(book => cleanSheetWhitespace(book?.releaseCatalog) && cleanSheetWhitespace(book?.bookShortener));
}

async function getReviewQueueIncludeSet() {
  try {
    return buildReviewQueueIncludeSetFromBooks(await getReleaseCatalogQueueBooks());
  } catch {
    return new Set(fallbackReviewQueueIncludeTitles.map(normalizeBookKey).filter(Boolean));
  }
}

async function getReviewQueueIncludeTitles() {
  try {
    return (await getReleaseCatalogQueueBooks()).map(book => book.title);
  } catch {
    return fallbackReviewQueueIncludeTitles;
  }
}

async function getReleaseCatalogMetadata() {
  try {
    const books = await getReleaseCatalogQueueBooks();
    const sortReleaseCatalogOptions = (left, right) => {
      const parseCatalog = value => {
        const cleaned = cleanSheetWhitespace(value);
        const match = cleaned.match(/^(Spring|Fall)\s+(\d{4})$/i);
        if (match) {
          return {
            raw: cleaned,
            recognized: true,
            seasonRank: match[1].toLowerCase() === "fall" ? 2 : 1,
            year: Number(match[2])
          };
        }
        return {
          raw: cleaned,
          recognized: false,
          seasonRank: 0,
          year: -Infinity
        };
      };

      const a = parseCatalog(left);
      const b = parseCatalog(right);
      if (a.recognized && b.recognized) {
        if (a.year !== b.year) return b.year - a.year;
        if (a.seasonRank !== b.seasonRank) return b.seasonRank - a.seasonRank;
        return a.raw.localeCompare(b.raw);
      }
      if (a.recognized) return -1;
      if (b.recognized) return 1;
      return a.raw.localeCompare(b.raw);
    };

    const releaseCatalogOptions = Array.from(new Set(
      books.map(book => cleanSheetWhitespace(book.releaseCatalog)).filter(Boolean)
    )).sort(sortReleaseCatalogOptions);
    const releaseCatalogByTitle = {};
    books.forEach(book => {
      const key = normalizeBookKey(book.title);
      if (!key) return;
      releaseCatalogByTitle[key] = cleanSheetWhitespace(book.releaseCatalog);
    });
    return { releaseCatalogOptions, releaseCatalogByTitle };
  } catch {
    return { releaseCatalogOptions: [], releaseCatalogByTitle: {} };
  }
}

function resolvePublishingBookMeta(bookTitle = "", canonicalBookTitle = "") {
  const rawTitles = [bookTitle, canonicalBookTitle].filter(Boolean);
  const keys = new Set();
  rawTitles.forEach(title => {
    const cleaned = cleanSheetWhitespace(title);
    const normalized = normalizeBookKey(cleaned);
    if (normalized) {
      keys.add(normalized);
    }
    if (cleaned.includes(":")) {
      const baseTitle = cleanSheetWhitespace(cleaned.split(/\s*:\s*/, 1)[0] || "");
      const baseKey = normalizeBookKey(baseTitle);
      if (baseKey) {
        keys.add(baseKey);
      }
    }
  });
  for (const key of keys) {
    const match = publishingBooksCache?.get(key);
    if (match) {
      return match;
    }
  }
  if (keys.has(normalizeBookKey("Tooth Gaps in the Archives"))) {
    return {
      title: "Tooth Gaps in the Archives",
      releaseCatalog: "Fall 2026",
      bookShortener: "TGIT"
    };
  }
  return null;
}

function canonicalizeKnownAuthorName(value, { bookTitle = "" } = {}) {
  const cleaned = cleanSheetWhitespace(value)
    .replace(/[’‘]/g, "'")
    .replace(/[“”]/g, "\"");
  const normalizedAuthor = cleaned
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[.,()]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
  const normalizedBook = normalizeBookKey(bookTitle);

  if (normalizedBook.startsWith("poetry by chance")) {
    return "Taylor Mali";
  }
  if (normalizedBook === "tooth gaps in the archives" && !cleaned) {
    return "Francis Dylan Waguespack";
  }
  if (normalizedBook === "tooth gaps in the archives" && normalizedAuthor.startsWith("francis")) {
    return "Francis Dylan Waguespack";
  }

  return cleaned;
}

function buildWeaverGraphicsRequestId({ recordId = "", author = "", poemTitle = "", bookTitle = "", quoteText = "" } = {}) {
  const stable = cleanSheetWhitespace(recordId);
  if (stable && !/^(?:weaver:)?row-\d+$/i.test(stable)) {
    return stable.startsWith("weaver:") ? stable : `weaver:${stable}`;
  }

  const excerptHash = createHash("sha256")
    .update([
      cleanSheetWhitespace(bookTitle).toLowerCase(),
      cleanSheetWhitespace(author).toLowerCase(),
      cleanSheetWhitespace(poemTitle).toLowerCase(),
      cleanSheetWhitespace(quoteText).toLowerCase()
    ].join("\n"))
    .digest("hex")
    .slice(0, 32);
  return `weaver:qi:${excerptHash}`;
}

function buildPigCompletionId(completion = {}) {
  const explicit = cleanSheetWhitespace(completion.completionId || completion.id);
  if (explicit) return explicit;

  const requestId = cleanSheetWhitespace(completion.requestId);
  if (requestId) {
    return `pig-${requestId}`;
  }

  return `pig-${createHash("sha1")
    .update(JSON.stringify({
      author: cleanSheetWhitespace(completion.author).toLowerCase(),
      poemTitle: cleanSheetWhitespace(completion.poemTitle).toLowerCase(),
      bookTitle: cleanSheetWhitespace(completion.bookTitle).toLowerCase(),
      quoteText: String(completion.quoteText || "").replace(/\r\n/g, "\n").trim(),
      assetUrl: cleanSheetWhitespace(completion.assetUrl || completion.assetLinkUrl || completion.driveUrl),
      completedAt: cleanSheetWhitespace(completion.completedAt)
    }))
    .digest("hex")
    .slice(0, 16)}`;
}

function buildGraphicsLookupKey({ author = "", poemTitle = "", bookTitle = "", quoteText = "" } = {}) {
  return [
    cleanSheetWhitespace(author).toLowerCase(),
    cleanSheetWhitespace(poemTitle).toLowerCase(),
    cleanSheetWhitespace(bookTitle).toLowerCase(),
    String(quoteText || "").replace(/\r\n/g, "\n").trim()
  ].join("||");
}

function normalizeDriveMatchText(value) {
  return String(value || "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/&/g, " and ")
    .replace(/[’‘]/g, "'")
    .replace(/[“”]/g, "\"")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function stripFileExtension(name) {
  return String(name || "").replace(/\.[a-z0-9]{2,5}$/i, "").trim();
}

function buildDriveFolderImportCompletionId(fileId) {
  const cleaned = cleanSheetWhitespace(fileId);
  return cleaned ? `drive-folder-${cleaned}` : "";
}

function extractGoogleDriveFolderId(input) {
  const raw = cleanSheetWhitespace(input);
  if (!raw) {
    return "";
  }

  const directFolderMatch = raw.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (directFolderMatch) {
    return directFolderMatch[1];
  }

  const idQueryMatch = raw.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (idQueryMatch) {
    return idQueryMatch[1];
  }

  if (/^[a-zA-Z0-9_-]{20,}$/.test(raw)) {
    return raw;
  }

  return "";
}

function extractGoogleDriveFileId(input) {
  const raw = cleanSheetWhitespace(input).replace(/[.,;:!?]+$/, "");
  if (!raw) {
    return "";
  }

  const filePathMatch = raw.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (filePathMatch) {
    return filePathMatch[1];
  }

  const ucIdMatch = raw.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (ucIdMatch && !raw.includes("/folders/")) {
    return ucIdMatch[1];
  }

  return "";
}

function buildDriveImportBookVariants(bookTitle) {
  const variants = new Set();
  const cleaned = cleanSheetWhitespace(bookTitle);
  const normalized = normalizeDriveMatchText(cleaned);
  if (normalized) {
    variants.add(normalized);
  }
  const beforeColon = cleaned.split(":")[0] || "";
  const normalizedBeforeColon = normalizeDriveMatchText(beforeColon);
  if (normalizedBeforeColon) {
    variants.add(normalizedBeforeColon);
  }
  return Array.from(variants).filter(Boolean);
}

function buildDriveImportAuthorVariants(author) {
  const variants = new Set();
  const cleaned = cleanSheetWhitespace(author);
  const normalized = normalizeDriveMatchText(cleaned);
  if (normalized) {
    variants.add(normalized);
  }
  const parts = normalized.split(" ").filter(Boolean);
  if (parts.length) {
    variants.add(parts[parts.length - 1]);
  }
  return Array.from(variants).filter(Boolean);
}

function buildDriveImportShortenerVariants(shortener) {
  const cleaned = cleanSheetWhitespace(shortener);
  if (!cleaned) {
    return [];
  }
  const baseVariants = [
    normalizeDriveMatchText(cleaned),
    normalizeDriveMatchText(cleaned.replace(/[^a-zA-Z0-9]/g, ""))
  ].filter(Boolean);
  const allVariants = new Set(baseVariants);
  baseVariants.forEach(variant => {
    if (variant.length < 4) return;
    for (let index = 0; index < variant.length; index += 1) {
      const oneCharDropped = variant.slice(0, index) + variant.slice(index + 1);
      if (oneCharDropped.length >= 3) {
        allVariants.add(oneCharDropped);
      }
    }
  });
  return Array.from(allVariants);
}

async function loadDriveImportBookMetadata() {
  let publishingBooks = [];
  let catalogBooks = [];

  try {
    publishingBooks = await getPublishingOrderBooks();
  } catch (_error) {
    publishingBooks = [];
  }

  try {
    const catalogResult = await runIntakeCatalogLookup({ action: "books" });
    catalogBooks = Array.isArray(catalogResult?.books) ? catalogResult.books : [];
  } catch (_error) {
    catalogBooks = [];
  }

  const shortenerByBookKey = new Map();
  const sourceCatalogByBookKey = new Map();

  publishingBooks.forEach(book => {
    const key = normalizeBookKey(book.title);
    if (!key) return;
    if (book.bookShortener) {
      shortenerByBookKey.set(key, cleanSheetWhitespace(book.bookShortener));
    }
    sourceCatalogByBookKey.set(key, {
      releaseCatalog: cleanSheetWhitespace(book.releaseCatalog),
      bookShortener: cleanSheetWhitespace(book.bookShortener)
    });
  });

  catalogBooks.forEach(book => {
    const key = normalizeBookKey(book.title);
    if (!key) return;
    if (!shortenerByBookKey.has(key) && book.bookShortener) {
      shortenerByBookKey.set(key, cleanSheetWhitespace(book.bookShortener));
    }
    if (!sourceCatalogByBookKey.has(key)) {
      sourceCatalogByBookKey.set(key, {
        releaseCatalog: "",
        bookShortener: cleanSheetWhitespace(book.bookShortener)
      });
    }
  });

  return { shortenerByBookKey, sourceCatalogByBookKey };
}

function getAuthorLastNameLabelServer(author) {
  const cleaned = cleanSheetWhitespace(author);
  if (!cleaned) return "";
  const parts = cleaned.split(/\s+/).filter(Boolean);
  const lastName = parts.length ? parts[parts.length - 1] : cleaned;
  return lastName.replace(/[^A-Za-z0-9'-]/g, "").toUpperCase();
}

function buildGraphicsImportLabel(match, shortenerByBookKey = null) {
  const authorLastName = getAuthorLastNameLabelServer(match.author || "") || "UNKNOWN";
  const bookShortener = cleanSheetWhitespace(
    shortenerByBookKey?.get(normalizeBookKey(match.bookTitle)) || ""
  ) || "BOOK";
  const poemTitle = cleanSheetWhitespace(match.poemTitle || "").toUpperCase() || "UNTITLED";
  return `${authorLastName} - ${bookShortener} - QI - ${poemTitle}`;
}

function buildAuthorBookCardinality(records) {
  const booksByAuthor = new Map();
  records.forEach(record => {
    const authorKey = normalizeDriveMatchText(record.author);
    const bookKey = normalizeBookKey(record.bookTitle);
    if (!authorKey || !bookKey) return;
    if (!booksByAuthor.has(authorKey)) {
      booksByAuthor.set(authorKey, new Set());
    }
    booksByAuthor.get(authorKey).add(bookKey);
  });
  return booksByAuthor;
}

function scoreDriveFolderFileAgainstRequest(fileName, record, context = {}) {
  const normalizedFileName = normalizeDriveMatchText(stripFileExtension(fileName));
  const poemVariants = buildDriveImportBookVariants(record.poemTitle);
  const bookVariants = buildDriveImportBookVariants(record.bookTitle);
  const authorVariants = buildDriveImportAuthorVariants(record.author);
  const shortenerVariants = buildDriveImportShortenerVariants(
    context.shortenerByBookKey?.get(normalizeBookKey(record.bookTitle)) || ""
  );
  const authorBooks = context.booksByAuthor?.get(normalizeDriveMatchText(record.author));
  const reasons = [];

  let poemScore = 0;
  let hasBook = false;
  let hasAuthor = false;

  for (const variant of poemVariants) {
    if (!variant || variant.length < 3) continue;
    if (normalizedFileName === variant) {
      poemScore = Math.max(poemScore, 70);
      reasons.push("exact poem title");
      break;
    }
    if (normalizedFileName.includes(variant)) {
      poemScore = Math.max(poemScore, 45);
      reasons.push("poem title in filename");
    }
  }

  if (!poemScore) {
    return {
      score: 0,
      normalizedFileName,
      reasons: []
    };
  }

  for (const variant of bookVariants) {
    if (!variant || variant.length < 4) continue;
    if (normalizedFileName.includes(variant)) {
      hasBook = true;
      reasons.push("book title in filename");
      break;
    }
  }

  if (!hasBook) {
    for (const variant of shortenerVariants) {
      if (!variant || variant.length < 2) continue;
      if (normalizedFileName.includes(variant)) {
        hasBook = true;
        reasons.push("book shortener in filename");
        break;
      }
    }
  }

  for (const variant of authorVariants) {
    if (!variant || variant.length < 4) continue;
    if (normalizedFileName.includes(variant)) {
      hasAuthor = true;
      reasons.push("author in filename");
      break;
    }
  }

  if (!hasBook && hasAuthor && authorBooks?.size === 1) {
    hasBook = true;
    reasons.push("author maps to one open book");
  }

  const score = poemScore + (hasBook ? 20 : 0) + (hasAuthor ? 10 : 0);
  return {
    score,
    normalizedFileName,
    hasBook,
    hasAuthor,
    reasons: Array.from(new Set(reasons))
  };
}

function buildDrivePreviewUrl(fileId) {
  const cleaned = cleanSheetWhitespace(fileId);
  return cleaned ? `https://drive.google.com/uc?export=view&id=${encodeURIComponent(cleaned)}` : "";
}

function isSheetYes(value) {
  return cleanSheetWhitespace(value).toUpperCase() === "Y";
}

function parseSheetInteger(value) {
  const parsed = parseInt(value, 10);
  return Number.isNaN(parsed) ? 0 : parsed;
}

function toA1Column(columnNumber) {
  let current = Number(columnNumber);
  let label = "";
  while (current > 0) {
    const remainder = (current - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    current = Math.floor((current - 1) / 26);
  }
  return label;
}

function isTruthyParam(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return normalized === "1" || normalized === "true" || normalized === "y" || normalized === "yes";
}

function normalizeDecision(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "accept") return "accept";
  if (normalized === "accept_skip_graphic" || normalized === "accept_skip_graphics") return "accept_skip_graphic";
  if (normalized === "reject") return "reject";
  if (normalized === "needs_correction") return "needs_correction";
  return "";
}

function normalizeReviewDecisionValue(value) {
  const normalized = normalizeDecision(value);
  if (normalized === "accept") return "ACCEPT";
  if (normalized === "accept_skip_graphic") return "ACCEPT_SKIP_GRAPHIC";
  if (normalized === "reject") return "REJECT";
  if (normalized === "needs_correction") return "NEEDS_CORRECTION";
  return "";
}

function isAcceptedExcerptReviewDecision(value) {
  const normalized = normalizeReviewDecisionValue(value);
  return normalized === "ACCEPT" || normalized === "ACCEPT_SKIP_GRAPHIC";
}

function getBookTitleDisplayScore(title) {
  const text = String(title || "").trim().replace(/\s+/g, " ");
  if (!text) return -Infinity;

  let score = 0;
  if (text === String(title || "")) score += 2;
  if (/[a-z]/.test(text)) score += 3;
  if (/^[A-Z0-9\s&'"?!:;.,()\/|-]+$/.test(text) && !/[a-z]/.test(text)) score -= 2;
  score -= Math.max(0, text.length - text.trim().length);
  return score;
}

function choosePreferredBookTitle(currentTitle, candidateTitle) {
  if (!currentTitle) return candidateTitle;
  if (!candidateTitle) return currentTitle;

  const currentScore = getBookTitleDisplayScore(currentTitle);
  const candidateScore = getBookTitleDisplayScore(candidateTitle);
  if (candidateScore > currentScore) return candidateTitle;
  if (candidateScore < currentScore) return currentTitle;

  return candidateTitle.length < currentTitle.length ? candidateTitle : currentTitle;
}

function mergeBookSummariesByTitle(books = []) {
  const byKey = new Map();

  books.forEach(book => {
    const title = String(book?.title || "").trim().replace(/\s+/g, " ");
    const key = normalizeBookKey(title);
    if (!key) return;

    if (!byKey.has(key)) {
      byKey.set(key, {
        key,
        title,
        count: 0,
        variants: new Set()
      });
    }

    const summary = byKey.get(key);
    summary.title = choosePreferredBookTitle(summary.title, title);
    summary.count += Number(book?.count || 0);
    if (title) {
      summary.variants.add(title);
    }
  });

  return Array.from(byKey.values())
    .map(summary => ({
      title: summary.title,
      count: summary.count,
      variants: Array.from(summary.variants).sort((left, right) => left.localeCompare(right))
    }))
    .sort((left, right) => left.title.localeCompare(right.title));
}

function mergeRecordCollections(recordSets = []) {
  const byKey = new Map();

  recordSets.flat().forEach(record => {
    const key = record?.recordId || String(record?.sheetRow || record?.sourceRow || "");
    if (!key) return;
    if (!byKey.has(key)) {
      byKey.set(key, record);
    }
  });

  return Array.from(byKey.values());
}

function collapsePendingGraphicsQcRecords(records = []) {
  const groups = new Map();

  records.forEach(record => {
    const key = cleanSheetWhitespace(record?.graphicsRequestId)
      || cleanSheetWhitespace(record?.recordId)
      || String(record?.sheetRow || "");
    if (!key) return;
    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key).push(record);
  });

  return Array.from(groups.values()).flatMap(group => {
    if (group.some(record => hasResolvedGraphicsQcDecision(record))) {
      return [];
    }

    return [group.slice().sort((left, right) => {
      const leftTime = Date.parse(left.completedAt || left.graphicsQcUpdatedAt || "") || 0;
      const rightTime = Date.parse(right.completedAt || right.graphicsQcUpdatedAt || "") || 0;
      if (leftTime !== rightTime) {
        return rightTime - leftTime;
      }
      return Number(right.sheetRow || 0) - Number(left.sheetRow || 0);
    })[0]];
  });
}

let canonicalGraphicsBookAuthorMapPromise = null;

async function getCanonicalGraphicsBookAuthorMap() {
  if (!canonicalGraphicsBookAuthorMapPromise) {
    canonicalGraphicsBookAuthorMapPromise = (async () => {
      const byBookKey = new Map();
      const register = (bookTitle, author) => {
        const key = normalizeBookKey(bookTitle);
        const cleanedAuthor = cleanSheetWhitespace(author);
        if (!key || !cleanedAuthor || byBookKey.has(key)) return;
        byBookKey.set(key, cleanedAuthor);
      };

      const catalogResult = await runIntakeCatalogLookup({ action: "books" }).catch(() => ({ books: [] }));
      const catalogBooks = Array.isArray(catalogResult?.books) ? catalogResult.books : [];
      catalogBooks.forEach(book => register(book.title, book.author));

      const publishingBooks = await getPublishingOrderBooks().catch(() => []);
      publishingBooks.forEach(book => register(book.title, book.author));

      return byBookKey;
    })();
  }

  return canonicalGraphicsBookAuthorMapPromise;
}

function resolveGraphicsAuthor(rawAuthor, bookTitle, canonicalBookAuthorMap = null) {
  const explicitAuthor = canonicalizeKnownAuthorName(rawAuthor || "", { bookTitle });
  if (explicitAuthor) return explicitAuthor;
  const fallbackAuthor = canonicalBookAuthorMap?.get(normalizeBookKey(bookTitle)) || "";
  return canonicalizeKnownAuthorName(fallbackAuthor, { bookTitle });
}

function buildGraphicsRequestRecordFromQueueRow(row, index, canonicalBookAuthorMap = null) {
  const bookTitle = cleanSheetWhitespace(row[2]);
  const author = resolveGraphicsAuthor(row[0] || "", bookTitle, canonicalBookAuthorMap);
  const poemTitle = String(row[1] || "");
  const quoteText = String(row[3] || "");
  const notes = String(row[4] || "");
  const noteMeta = parseIntakeMetadataFromNotes(notes);
  const approved = cleanSheetWhitespace(row[5]);
  const created = cleanSheetWhitespace(row[6]);
  const workflowStatus = cleanSheetWhitespace(row[7]);
  const recordId = String(row[8] || "");

  if (
    !bookTitle ||
    !cleanSheetWhitespace(quoteText) ||
    workflowStatus.toLowerCase() === "duplicate_suppressed"
  ) {
    return null;
  }

  return {
    queueSheetRow: index + 2,
    bookTitle,
    poemTitle,
    author,
    quoteText,
    notes,
    socialMediaHandle: noteMeta.socialMediaHandle || "",
    instagramHandle: noteMeta.socialMediaHandle || "",
    igHandle: noteMeta.socialMediaHandle || "",
    approved,
    created,
    workflowStatus,
    recordId,
    graphicsRequestId: buildWeaverGraphicsRequestId({
      recordId,
      author,
      poemTitle,
      bookTitle,
      quoteText
    }),
    source: "weaver_graphics_queue"
  };
}

function buildGraphicsHandoffLedgerRequest(record) {
  const socialMediaHandle = cleanSheetWhitespace(
    record?.socialMediaHandle || record?.instagramHandle || record?.igHandle
  );
  const source = cleanSheetWhitespace(record?.source);
  const isRework = source === "weaver_qc_rework" || cleanSheetWhitespace(record?.queueView) === "rework";
  const contentType = normalizePoetryPleaseContentType(
    record?.contentType || record?.imageType,
    "QI"
  );
  const reworkReason = cleanSheetWhitespace(record?.reworkReason || record?.rejectReason || record?.rejectedReason);
  const requestedChanges = String(record?.requestedChanges || record?.qcNote || record?.graphicsQcNote || "");
  const bookKey = cleanSheetWhitespace(record?.bookKey) || normalizeBookKey(record?.bookTitle);
  return {
    graphicsRequestId: cleanSheetWhitespace(record?.graphicsRequestId),
    sourceSystem: isRework ? "weaver_qc_rework" : "weaver",
    sourceStatus: cleanSheetWhitespace(record?.requestStatus || (isRework ? "rework_requested" : "open")),
    contentType,
    imageType: contentType,
    bookKey,
    pigProjectId: cleanSheetWhitespace(record?.pigProjectId),
    editableProjectFileId: cleanSheetWhitespace(record?.editableProjectFileId || record?.projectFileId),
    editableProjectUrl: cleanSheetWhitespace(record?.editableProjectUrl),
    originalGraphicsRequestId: cleanSheetWhitespace(record?.originalGraphicsRequestId),
    revisionOf: cleanSheetWhitespace(record?.revisionOf),
    version: record?.version || "",
    reworkReason,
    metadataIssue: cleanSheetWhitespace(record?.metadataIssue),
    aestheticIssue: cleanSheetWhitespace(record?.aestheticIssue),
    qcNote: String(record?.qcNote || record?.graphicsQcNote || ""),
    requestedChanges,
    sourcePayload: {
      graphicsRequestId: cleanSheetWhitespace(record?.graphicsRequestId),
      originalGraphicsRequestId: cleanSheetWhitespace(record?.originalGraphicsRequestId),
      sourceCompletionId: cleanSheetWhitespace(record?.sourceCompletionId),
      revisionOf: cleanSheetWhitespace(record?.revisionOf),
      queueSheetRow: record?.queueSheetRow || "",
      recordId: cleanSheetWhitespace(record?.recordId),
      author: String(record?.author || ""),
      poemTitle: String(record?.poemTitle || ""),
      bookTitle: String(record?.bookTitle || ""),
      bookKey,
      quoteText: String(record?.quoteText || ""),
      contentType,
      imageType: contentType,
      pigProjectId: cleanSheetWhitespace(record?.pigProjectId),
      editableProjectFileId: cleanSheetWhitespace(record?.editableProjectFileId || record?.projectFileId),
      editableProjectUrl: cleanSheetWhitespace(record?.editableProjectUrl),
      version: record?.version || "",
      reworkReason,
      metadataIssue: cleanSheetWhitespace(record?.metadataIssue),
      aestheticIssue: cleanSheetWhitespace(record?.aestheticIssue),
      qcNote: String(record?.qcNote || record?.graphicsQcNote || ""),
      requestedChanges,
      previousAssetUrl: cleanSheetWhitespace(record?.previousAssetUrl || record?.assetUrl),
      previousAssetPreviewUrl: cleanSheetWhitespace(record?.previousAssetPreviewUrl || record?.assetPreviewUrl),
      socialMediaHandle,
      instagramHandle: socialMediaHandle,
      igHandle: socialMediaHandle
    }
  };
}

function sendJson(res, statusCode, body) {
  res.writeHead(statusCode, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(body, null, 2));
}

function getRequestCorrelationId(req) {
  const headerValue = cleanSheetWhitespace(req?.headers?.["x-request-id"]);
  return headerValue || `weaver-${randomUUID()}`;
}

function escapeHtml(text) {
  return String(text || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function getLibraryStatusLabel(status) {
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

function readRequestBody(req) {
  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", chunk => {
      body += chunk.toString();
    });
    req.on("end", () => resolve(body));
    req.on("error", reject);
  });
}

async function getServiceAccountAccessToken() {
  if (!process.env.K_SERVICE) {
    throw new Error("Service-account Sheets read is only available from the deployed Cloud Run service.");
  }

  const response = await fetch("http://metadata.google.internal/computeMetadata/v1/instance/service-accounts/default/token", {
    headers: {
      "Metadata-Flavor": "Google"
    }
  });

  if (!response.ok) {
    throw new Error(`Metadata token request failed with status ${response.status}.`);
  }

  const data = await response.json();
  if (!data?.access_token) {
    throw new Error("Metadata token response did not include an access token.");
  }

  return data.access_token;
}

async function fetchSheetValuesServer(
  range,
  { valueRenderOption = "FORMATTED_VALUE", targetSpreadsheetId = spreadsheetId } = {}
) {
  const token = await getServiceAccountAccessToken();
  const url = new URL(`https://sheets.googleapis.com/v4/spreadsheets/${targetSpreadsheetId}/values/${encodeURIComponent(range)}`);
  url.searchParams.set("valueRenderOption", valueRenderOption);

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });

  const text = await response.text();
  if (!response.ok) {
    throw new Error(`Sheets read failed for ${range}: ${response.status} ${text.slice(0, 240)}`);
  }

  const data = JSON.parse(text || "{}");
  return Array.isArray(data.values) ? data.values : [];
}

async function batchUpdateSheetValuesServer(data) {
  const token = await getServiceAccountAccessToken();
  const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      valueInputOption: "RAW",
      data
    })
  });

  const text = await response.text();
  if (!response.ok) {
    throw new Error(`Sheets batch update failed: ${response.status} ${text.slice(0, 240)}`);
  }

  return JSON.parse(text || "{}");
}

async function appendSheetValuesServer(range, values) {
  const token = await getServiceAccountAccessToken();
  const url = new URL(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}:append`);
  url.searchParams.set("valueInputOption", "RAW");
  url.searchParams.set("insertDataOption", "INSERT_ROWS");

  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ values })
  });

  const text = await response.text();
  if (!response.ok) {
    throw new Error(`Sheets append failed: ${response.status} ${text.slice(0, 240)}`);
  }

  return JSON.parse(text || "{}");
}

async function fetchSpreadsheetMetadataServer() {
  const token = await getServiceAccountAccessToken();
  const url = new URL(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`);
  url.searchParams.set("fields", "sheets(properties(sheetId,title,gridProperties(columnCount)))");

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });

  const text = await response.text();
  if (!response.ok) {
    throw new Error(`Sheets metadata read failed: ${response.status} ${text.slice(0, 240)}`);
  }

  return JSON.parse(text || "{}");
}

async function sheetBatchUpdateServer(requests) {
  const token = await getServiceAccountAccessToken();
  const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ requests })
  });

  const text = await response.text();
  if (!response.ok) {
    throw new Error(`Sheets structural update failed: ${response.status} ${text.slice(0, 240)}`);
  }

  return JSON.parse(text || "{}");
}

async function ensureCleanupQcColumnsServer() {
  const metadata = await fetchSpreadsheetMetadataServer();
  const targetSheet = (metadata.sheets || []).find(sheet => {
    return cleanSheetWhitespace(sheet?.properties?.title) === cleanSheetWhitespace(graphicsCleanupSheetName);
  });

  if (!targetSheet?.properties?.sheetId) {
    throw new Error(`Cleanup sheet "${graphicsCleanupSheetName}" not found.`);
  }

  const sheetId = targetSheet.properties.sheetId;
  const currentColumnCount = Number(targetSheet.properties.gridProperties?.columnCount || 0);
  const requests = [];

  if (currentColumnCount < 12) {
    requests.push({
      appendDimension: {
        sheetId,
        dimension: "COLUMNS",
        length: 12 - currentColumnCount
      }
    });
  }

  if (requests.length) {
    await sheetBatchUpdateServer(requests);
  }

  await batchUpdateSheetValuesServer([
    {
      range: `'${graphicsCleanupSheetName.replace(/'/g, "''")}'!J1:L1`,
      values: [["graphic_qc_decision", "graphic_qc_note", "graphic_qc_updated_at"]]
    }
  ]);
}

async function ensurePigCompletedGraphicsSheetServer() {
  const metadata = await fetchSpreadsheetMetadataServer();
  let targetSheet = (metadata.sheets || []).find(sheet => {
    return cleanSheetWhitespace(sheet?.properties?.title) === cleanSheetWhitespace(pigCompletedGraphicsSheetName);
  });

  if (!targetSheet) {
    const created = await sheetBatchUpdateServer([
      {
        addSheet: {
          properties: {
            title: pigCompletedGraphicsSheetName,
            gridProperties: {
              rowCount: 1000,
            columnCount: 19
            }
          }
        }
      }
    ]);
    targetSheet = created?.replies?.[0]?.addSheet || null;
  }

  const sheetId = Number(targetSheet?.properties?.sheetId || targetSheet?.properties?.sheetProperties?.sheetId || 0);
  if (!sheetId) {
    throw new Error(`P.I.G. sheet "${pigCompletedGraphicsSheetName}" could not be created or found.`);
  }

  const currentColumnCount = Number(
    targetSheet?.properties?.gridProperties?.columnCount ||
    targetSheet?.properties?.sheetProperties?.gridProperties?.columnCount ||
    0
  );
  if (currentColumnCount < 19) {
    await sheetBatchUpdateServer([
      {
        appendDimension: {
          sheetId,
          dimension: "COLUMNS",
          length: 19 - currentColumnCount
        }
      }
    ]);
  }

  await batchUpdateSheetValuesServer([
    {
      range: `'${pigCompletedGraphicsSheetName.replace(/'/g, "''")}'!A1:S1`,
      values: PIG_COMPLETION_HEADERS
    }
  ]);
}

async function ensureExcerptPoetryPleaseColumnsServer() {
  const metadata = await fetchSpreadsheetMetadataServer();
  const targetSheet = (metadata.sheets || []).find(sheet => {
    return cleanSheetWhitespace(sheet?.properties?.title) === cleanSheetWhitespace(sourceSheetName);
  });

  if (!targetSheet?.properties?.sheetId) {
    throw new Error(`Source sheet "${sourceSheetName}" not found.`);
  }

  const sheetId = targetSheet.properties.sheetId;
  const currentColumnCount = Number(targetSheet.properties.gridProperties?.columnCount || 0);
  if (currentColumnCount < SHEET_SOURCE_CONFIG.columnMap.excerptPoetryPleaseNote) {
    await sheetBatchUpdateServer([
      {
        appendDimension: {
          sheetId,
          dimension: "COLUMNS",
          length: SHEET_SOURCE_CONFIG.columnMap.excerptPoetryPleaseNote - currentColumnCount
        }
      }
    ]);
  }

  await batchUpdateSheetValuesServer([
    {
      range: `'${sourceSheetName.replace(/'/g, "''")}'!BJ1:BL1`,
      values: [["excerpt_poetry_please_status", "excerpt_poetry_please_updated_at", "excerpt_poetry_please_note"]]
    }
  ]);
}

function isDriveFolderAssetUrl(value) {
  const text = cleanSheetWhitespace(value);
  return /drive\.google\.com\/drive\/folders\//i.test(text) || /\/folders\//i.test(text);
}

function assertPigCompletionFileAsset(completion) {
  const assetUrl = cleanSheetWhitespace(completion.assetUrl || completion.assetLinkUrl || completion.driveUrl);
  const assetPreviewUrl = cleanSheetWhitespace(completion.assetPreviewUrl || completion.previewUrl || completion.thumbnailUrl || assetUrl);
  if (!assetUrl) {
    throw new Error('P.I.G. completion must include a row-level assetUrl.');
  }
  if (!assetPreviewUrl) {
    throw new Error('P.I.G. completion must include a row-level assetPreviewUrl.');
  }
  if (isDriveFolderAssetUrl(assetUrl) || isDriveFolderAssetUrl(assetPreviewUrl)) {
    throw new Error('P.I.G. completion assetUrl/assetPreviewUrl must point to the specific image file, not a Drive folder.');
  }
}

function buildPigCompletionRowValues(completion, existingRow = []) {
  assertPigCompletionFileAsset(completion);
  const completionId = buildPigCompletionId(completion);
  const requestId = cleanSheetWhitespace(completion.requestId || completion.graphicsRequestId);
  const sourceRecordId = cleanSheetWhitespace(completion.sourceRecordId || completion.recordId);
  const sourceSheetRow = parseInt(completion.sourceSheetRow || completion.sheetRow, 10) || "";
  const assetUrl = cleanSheetWhitespace(completion.assetUrl || completion.assetLinkUrl || completion.driveUrl);
  const assetPreviewUrl = cleanSheetWhitespace(completion.assetPreviewUrl || completion.previewUrl || completion.thumbnailUrl || assetUrl);
  const completedAt = cleanSheetWhitespace(completion.completedAt || new Date().toISOString());
  const sourceTool = cleanSheetWhitespace(completion.sourceTool || "P.I.G.");
  const contentType = normalizePoetryPleaseContentType(completion.contentType || completion.imageType || "", "");
  let productionNotes = String(completion.productionNotes || completion.notes || "");
  if (contentType && !/^\s*(content type|image type)\s*:/im.test(productionNotes)) {
    productionNotes = `${productionNotes}${productionNotes.trim() ? "\n" : ""}Content Type: ${contentType}`;
  }

  return [
    completionId,
    requestId,
    String(completion.author || ""),
    String(completion.poemTitle || ""),
    String(completion.bookTitle || ""),
    String(completion.quoteText || ""),
    assetUrl,
    assetPreviewUrl,
    sourceRecordId,
    sourceSheetRow,
    productionNotes,
    completedAt,
    existingRow[PIG_COMPLETION_COLUMNS.qcDecision - 1] || "",
    existingRow[PIG_COMPLETION_COLUMNS.qcNote - 1] || "",
    existingRow[PIG_COMPLETION_COLUMNS.qcUpdatedAt - 1] || "",
    sourceTool,
    existingRow[PIG_COMPLETION_COLUMNS.poetryPleaseStatus - 1] || "",
    existingRow[PIG_COMPLETION_COLUMNS.poetryPleaseUpdatedAt - 1] || "",
    existingRow[PIG_COMPLETION_COLUMNS.poetryPleaseNote - 1] || ""
  ];
}

function buildPigQcRecordFromSheetRow(row, index, canonicalBookAuthorMap = null) {
  const completionId = cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.completionId - 1]);
  if (!completionId) return null;

  const requestId = cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.requestId - 1]);
  const isReworkCompletion = requestId.toLowerCase().startsWith("rework:");
  const poemTitle = String(row[PIG_COMPLETION_COLUMNS.poemTitle - 1] || "");
  const bookTitle = String(row[PIG_COMPLETION_COLUMNS.bookTitle - 1] || "");
  const author = resolveGraphicsAuthor(row[PIG_COMPLETION_COLUMNS.author - 1] || "", bookTitle, canonicalBookAuthorMap);
  const quoteText = String(row[PIG_COMPLETION_COLUMNS.quoteText - 1] || "");
  const assetUrl = cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.assetUrl - 1]);
  const assetPreviewUrl = cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.assetPreviewUrl - 1]) || assetUrl;
  const sourceRecordId = cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.sourceRecordId - 1]);
  const sourceSheetRow = parseInt(row[PIG_COMPLETION_COLUMNS.sourceSheetRow - 1], 10) || "";
  const productionNotes = String(row[PIG_COMPLETION_COLUMNS.productionNotes - 1] || "");
  const noteMeta = parseIntakeMetadataFromNotes(productionNotes);
  const completedAt = cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.completedAt - 1]);
  const sourceTool = cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.sourceTool - 1]) || "P.I.G.";
  const qcDecision = cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.qcDecision - 1]);
  const qcNote = String(row[PIG_COMPLETION_COLUMNS.qcNote - 1] || "");
  const qcUpdatedAt = cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.qcUpdatedAt - 1]);
  const poetryPleaseStatus = cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.poetryPleaseStatus - 1]);
  const poetryPleaseUpdatedAt = cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.poetryPleaseUpdatedAt - 1]);
  const poetryPleaseNote = String(row[PIG_COMPLETION_COLUMNS.poetryPleaseNote - 1] || "");

  return {
    storageTarget: "pig_sheet",
    pigCompletionId: completionId,
    pigRequestId: requestId,
    isReworkCompletion,
    requestKind: isReworkCompletion ? "rework_completion" : "completion",
    sheetRow: index + 2,
    author,
    poemTitle,
    bookTitle,
    quoteText,
    notes: productionNotes ? `${sourceTool}: ${productionNotes}` : `Returned by ${sourceTool}`,
    approved: "Y",
    created: "Y",
    workflowStatus: isReworkCompletion ? "Rework returned by P.I.G." : "Returned by P.I.G.",
    recordId: `pig:${completionId}`,
    graphicsRequestId: requestId || buildWeaverGraphicsRequestId({ recordId: sourceRecordId, author, poemTitle, bookTitle, quoteText }),
    assetLinkUrl: assetUrl,
    assetPreviewUrl,
    sourceRecordId,
    sourceSheetRow,
    completedAt,
    graphicsQcDecision: qcDecision,
    graphicsQcNote: qcNote,
    graphicsQcUpdatedAt: qcUpdatedAt,
    rejectReason: "",
    metadataIssue: "",
    aestheticIssue: "",
    sourceTool,
    contentType: normalizePoetryPleaseContentType(noteMeta.contentType, "QI"),
    imageType: normalizePoetryPleaseContentType(noteMeta.contentType, "QI"),
    poetryPleaseStatus,
    poetryPleaseUpdatedAt,
    poetryPleaseNote
  };
}

function overlayRuntimeGraphicsState(record, runtimeState = null) {
  if (!record || !runtimeState) return record;
  const completionId = cleanSheetWhitespace(record.pigCompletionId);
  if (!completionId) return record;

  const qc = runtimeState.qc_reviews?.[completionId] || null;
  const handoff = runtimeState.handoffs?.[completionId] || null;
  const parsedNote = parseGraphicsQcStructuredNoteServer(qc?.note ?? record.graphicsQcNote);

  return {
    ...record,
    graphicsQcDecision: cleanSheetWhitespace(qc?.decision) || record.graphicsQcDecision,
    graphicsQcNote: qc?.note ?? record.graphicsQcNote,
    graphicsQcUpdatedAt: cleanSheetWhitespace(qc?.reviewed_at) || record.graphicsQcUpdatedAt,
    rejectReason: parsedNote.rejectReason || record.rejectReason || "",
    metadataIssue: cleanSheetWhitespace(qc?.metadata_issue) || parsedNote.metadataIssue || record.metadataIssue || "",
    aestheticIssue: cleanSheetWhitespace(qc?.aesthetic_issue) || parsedNote.aestheticIssue || record.aestheticIssue || "",
    poetryPleaseStatus: cleanSheetWhitespace(handoff?.handoff_status) || record.poetryPleaseStatus,
    poetryPleaseUpdatedAt: cleanSheetWhitespace(handoff?.handed_off_at) || record.poetryPleaseUpdatedAt,
    poetryPleaseNote: cleanSheetWhitespace(handoff?.poetry_please_item_id)
      ? `item=${cleanSheetWhitespace(handoff.poetry_please_item_id)}`
      : record.poetryPleaseNote
  };
}

async function readPigCompletedGraphicsRows() {
  await ensurePigCompletedGraphicsSheetServer();
  const range = `'${pigCompletedGraphicsSheetName.replace(/'/g, "''")}'!A2:S`;
  return fetchSheetValuesServer(range);
}

async function getRuntimeGraphicsState(completionIds = []) {
  const ids = Array.isArray(completionIds)
    ? completionIds.map(value => cleanSheetWhitespace(value)).filter(Boolean)
    : [];
  if (!ids.length) {
    return { qc_reviews: {}, handoffs: {} };
  }
  try {
    const result = await syncWeaverRuntimeDb("get_graphics_state", { completionIds: ids });
    return {
      qc_reviews: result?.qc_reviews || {},
      handoffs: result?.handoffs || {}
    };
  } catch {
    return { qc_reviews: {}, handoffs: {} };
  }
}

async function getRuntimeGraphicsHandoffState(graphicsRequestIds = []) {
  const ids = Array.isArray(graphicsRequestIds)
    ? graphicsRequestIds.map(value => cleanSheetWhitespace(value)).filter(Boolean)
    : [];
  if (!ids.length) {
    return {};
  }
  try {
    const result = await syncWeaverRuntimeDb("get_handoff_requests", { graphicsRequestIds: ids });
    return Object.fromEntries(
      (Array.isArray(result?.records) ? result.records : [])
        .map(record => [cleanSheetWhitespace(record.graphicsRequestId), record])
        .filter(([requestId]) => requestId)
    );
  } catch {
    return {};
  }
}

function overlayRuntimeHandoffState(record, handoffStateByRequestId = null) {
  if (!record || !handoffStateByRequestId) return record;
  const requestId = cleanSheetWhitespace(record.graphicsRequestId);
  if (!requestId) return record;
  const handoff = handoffStateByRequestId[requestId];
  if (!handoff) return record;

  const qcStatus = cleanSheetWhitespace(handoff.qcStatus).toLowerCase();
  const qcPayload = handoff.qcPayload && typeof handoff.qcPayload === "object" ? handoff.qcPayload : {};
  const pigPayload = handoff.pigPayload && typeof handoff.pigPayload === "object" ? handoff.pigPayload : {};

  return {
    ...record,
    contentType: normalizePoetryPleaseContentType(handoff.contentType || handoff.imageType || record.contentType || record.imageType, record.contentType || "QI"),
    imageType: normalizePoetryPleaseContentType(handoff.imageType || handoff.contentType || record.imageType || record.contentType, record.imageType || "QI"),
    graphicsQcDecision: (
      qcStatus === "approved"
        ? "APPROVE"
        : (qcStatus === "rejected" || qcStatus === "needs_revision" ? "REJECT" : cleanSheetWhitespace(record.graphicsQcDecision))
    ) || record.graphicsQcDecision,
    graphicsQcNote: String(qcPayload.qcNote || qcPayload.note || record.graphicsQcNote || ""),
    rejectReason: cleanSheetWhitespace(qcPayload.rejectReason || qcPayload.qcRejectReason || record.rejectReason),
    metadataIssue: cleanSheetWhitespace(qcPayload.metadataIssue || record.metadataIssue),
    aestheticIssue: cleanSheetWhitespace(qcPayload.aestheticIssue || record.aestheticIssue),
    assetLinkUrl: cleanSheetWhitespace(handoff.assetUrl) || record.assetLinkUrl,
    assetPreviewUrl: cleanSheetWhitespace(handoff.assetPreviewUrl) || record.assetPreviewUrl,
    pigProjectId: cleanSheetWhitespace(
      handoff.pigProjectId || pigPayload.pigProjectId || record.pigProjectId,
    ),
    editableProjectFileId: cleanSheetWhitespace(
      handoff.editableProjectFileId
        || handoff.projectFileId
        || pigPayload.editableProjectFileId
        || pigPayload.projectFileId
        || record.editableProjectFileId,
    ),
    editableProjectUrl: cleanSheetWhitespace(
      handoff.editableProjectUrl || pigPayload.editableProjectUrl || record.editableProjectUrl,
    ),
    ledgerHandoffStatus: cleanSheetWhitespace(handoff.handoffStatus),
    ledgerPigStatus: cleanSheetWhitespace(handoff.pigStatus),
    ledgerQcStatus: cleanSheetWhitespace(handoff.qcStatus)
  };
}

async function fetchDriveJson(url) {
  const token = await getServiceAccountAccessToken();
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });
  const text = await response.text();
  if (!response.ok) {
    throw new Error(`Drive API request failed: ${response.status} ${text.slice(0, 240)}`);
  }
  return JSON.parse(text || "{}");
}

async function updateDriveFileName(fileId, fileName) {
  const token = await getServiceAccountAccessToken();
  const response = await fetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?supportsAllDrives=true`, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json; charset=utf-8"
    },
    body: JSON.stringify({ name: fileName })
  });
  const text = await response.text();
  if (!response.ok) {
    throw new Error(`Drive rename failed: ${response.status} ${text.slice(0, 240)}`);
  }
  return JSON.parse(text || "{}");
}

async function fetchDriveFileResponse(fileId) {
  const token = await getServiceAccountAccessToken();
  const url = new URL(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}`);
  url.searchParams.set("alt", "media");
  url.searchParams.set("supportsAllDrives", "true");
  return fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });
}

async function fetchDriveThumbnailResponse(fileId) {
  const url = new URL("https://drive.google.com/thumbnail");
  url.searchParams.set("id", fileId);
  url.searchParams.set("sz", "w1600");
  return fetch(url);
}

async function getDriveFolderMetadata(folderId) {
  const url = new URL(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(folderId)}`);
  url.searchParams.set("fields", "id,name,webViewLink");
  url.searchParams.set("supportsAllDrives", "true");
  return fetchDriveJson(url);
}

async function getDriveFileMetadata(fileId) {
  const url = new URL(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}`);
  url.searchParams.set("fields", "id,name,mimeType,thumbnailLink,webViewLink,createdTime");
  url.searchParams.set("supportsAllDrives", "true");
  return fetchDriveJson(url);
}

async function listDriveFolderImageFiles(folderId) {
  let pageToken = "";
  const files = [];

  do {
    const url = new URL("https://www.googleapis.com/drive/v3/files");
    url.searchParams.set("q", `'${folderId}' in parents and trashed=false and mimeType contains 'image/'`);
    url.searchParams.set("fields", "nextPageToken,files(id,name,mimeType,thumbnailLink,webViewLink,createdTime)");
    url.searchParams.set("orderBy", "name_natural");
    url.searchParams.set("pageSize", "1000");
    url.searchParams.set("supportsAllDrives", "true");
    url.searchParams.set("includeItemsFromAllDrives", "true");
    if (pageToken) {
      url.searchParams.set("pageToken", pageToken);
    }

    const data = await fetchDriveJson(url);
    files.push(...(Array.isArray(data.files) ? data.files : []));
    pageToken = cleanSheetWhitespace(data.nextPageToken);
  } while (pageToken);

  return files;
}

function chooseBestDriveFolderImageFile(files = []) {
  if (!Array.isArray(files) || files.length === 0) {
    return null;
  }

  return files
    .slice()
    .sort((left, right) => {
      const leftTime = Date.parse(left?.createdTime || "") || 0;
      const rightTime = Date.parse(right?.createdTime || "") || 0;
      if (leftTime !== rightTime) {
        return rightTime - leftTime;
      }
      return String(right?.name || "").localeCompare(String(left?.name || ""));
    })[0] || null;
}

function buildAuthorMatchTokens(author) {
  const withoutParens = String(author || "").replace(/\([^)]*\)/g, " ");
  return Array.from(new Set(
    normalizeDriveMatchText(withoutParens)
      .split(/\s+/)
      .filter(token => token.length >= 3)
  ));
}

function chooseMatchingDriveFolderImageFile(files = [], record = null) {
  if (!Array.isArray(files) || files.length === 0 || !record) {
    return null;
  }

  const titleNeedle = normalizeDriveMatchText(record.poemTitle || "");
  const authorTokens = buildAuthorMatchTokens(record.author || "");
  const scored = files.map(file => {
    const haystack = normalizeDriveMatchText(file?.name || "");
    let score = 0;

    if (titleNeedle && haystack.includes(titleNeedle)) {
      score += 200;
    }

    authorTokens.forEach(token => {
      if (haystack.includes(token)) {
        score += 25;
      }
    });

    if (String(file?.mimeType || "").startsWith("image/")) {
      score += 5;
    }

    return { file, score };
  }).sort((left, right) => {
    if (left.score !== right.score) {
      return right.score - left.score;
    }
    const leftTime = Date.parse(left.file?.createdTime || "") || 0;
    const rightTime = Date.parse(right.file?.createdTime || "") || 0;
    return rightTime - leftTime;
  });

  return scored[0]?.score > 0 ? scored[0].file : null;
}

async function hydrateGraphicsFolderAssets(records = []) {
  const folderIds = Array.from(new Set(records
    .map(record => extractGoogleDriveFolderId(record?.assetLinkUrl || record?.assetPreviewUrl || ""))
    .filter(Boolean)));

  if (folderIds.length === 0) {
    return records;
  }

  const folderFilesById = new Map();
  await Promise.all(folderIds.map(async folderId => {
    try {
      folderFilesById.set(folderId, await listDriveFolderImageFiles(folderId));
    } catch (_error) {
      folderFilesById.set(folderId, []);
    }
  }));

  const repairedWrites = [];
  const hydratedRecords = records.map(record => {
    const folderUrl = cleanSheetWhitespace(record?.assetLinkUrl || record?.assetPreviewUrl || "");
    const folderId = extractGoogleDriveFolderId(folderUrl);
    if (!folderId) {
      return record;
    }

    const chosenFile = chooseMatchingDriveFolderImageFile(folderFilesById.get(folderId) || [], record)
      || chooseBestDriveFolderImageFile(folderFilesById.get(folderId) || []);
    if (!chosenFile) {
      return {
        ...record,
        assetFolderUrl: folderUrl,
        assetLinkUrl: folderUrl,
        assetPreviewUrl: ""
      };
    }

    const repairedAssetUrl = cleanSheetWhitespace(chosenFile.webViewLink) || `https://drive.google.com/file/d/${encodeURIComponent(chosenFile.id)}/view`;
    const repairedPreviewUrl = cleanSheetWhitespace(chosenFile.thumbnailLink) || buildDrivePreviewUrl(chosenFile.id);
    const sheetRow = parseInt(record?.sheetRow, 10) || 0;
    if (sheetRow > 1) {
      repairedWrites.push({
        range: `'${pigCompletedGraphicsSheetName.replace(/'/g, "''")}'!G${sheetRow}:H${sheetRow}`,
        values: [[repairedAssetUrl, repairedPreviewUrl]]
      });
    }

    return {
      ...record,
      assetFolderUrl: folderUrl,
      assetLinkUrl: repairedAssetUrl,
      assetPreviewUrl: repairedPreviewUrl
    };
  });

  if (repairedWrites.length > 0) {
    await batchUpdateSheetValuesServer(repairedWrites);
  }

  return hydratedRecords;
}

function chooseDriveFolderImportMatch(files, openRequests, metadata = {}, completedRecords = []) {
  const usedRequestIds = new Set();
  const matches = [];
  const unmatched = [];
  const duplicates = [];
  const booksByAuthor = buildAuthorBookCardinality(openRequests);
  const completedBooksByAuthor = buildAuthorBookCardinality(completedRecords);
  const completedByDriveFileId = new Map();

  completedRecords.forEach(record => {
    const assetFileId = extractGoogleDriveFileId(record.assetLinkUrl || "");
    if (!assetFileId || completedByDriveFileId.has(assetFileId)) return;
    completedByDriveFileId.set(assetFileId, record);
  });

  files.forEach(file => {
    const exactCompletedRecord = cleanSheetWhitespace(file.id)
      ? completedByDriveFileId.get(cleanSheetWhitespace(file.id)) || null
      : null;
    if (exactCompletedRecord) {
      duplicates.push({
        fileId: file.id,
        fileName: file.name,
        mimeType: file.mimeType || "",
        assetUrl: cleanSheetWhitespace(file.webViewLink) || `https://drive.google.com/file/d/${encodeURIComponent(file.id)}/view`,
        assetPreviewUrl: cleanSheetWhitespace(file.thumbnailLink) || buildDrivePreviewUrl(file.id),
        graphicsRequestId: exactCompletedRecord.graphicsRequestId,
        queueSheetRow: exactCompletedRecord.queueSheetRow || "",
        recordId: exactCompletedRecord.recordId || "",
        author: exactCompletedRecord.author,
        poemTitle: exactCompletedRecord.poemTitle,
        bookTitle: exactCompletedRecord.bookTitle,
        quoteText: exactCompletedRecord.quoteText || "",
        reason: "This exact Drive file is already returned to Weaver.",
        existingAssetUrl: exactCompletedRecord.assetLinkUrl || "",
        existingAssetPreviewUrl: exactCompletedRecord.assetPreviewUrl || ""
      });
      return;
    }

    const scoredCandidates = openRequests
      .map(record => ({
        record,
        ...scoreDriveFolderFileAgainstRequest(file.name, record, {
          shortenerByBookKey: metadata.shortenerByBookKey,
          booksByAuthor
        })
      }))
      .filter(candidate => candidate.score > 0)
      .sort((left, right) => right.score - left.score);

    const byRequestId = new Map();
    scoredCandidates.forEach(candidate => {
      const requestId = cleanSheetWhitespace(candidate.record.graphicsRequestId);
      if (!requestId) return;
      const existing = byRequestId.get(requestId);
      if (!existing || candidate.score > existing.score) {
        byRequestId.set(requestId, candidate);
      }
    });

    const candidates = Array.from(byRequestId.values()).sort((left, right) => right.score - left.score);

    if (!candidates.length) {
      const completedCandidates = completedRecords
        .map(record => ({
          record,
          ...scoreDriveFolderFileAgainstRequest(file.name, record, {
            shortenerByBookKey: metadata.shortenerByBookKey,
            booksByAuthor: completedBooksByAuthor
          })
        }))
        .filter(candidate => candidate.score > 0)
        .sort((left, right) => right.score - left.score);
      const bestCompleted = completedCandidates[0] || null;
      const runnerUpCompleted = completedCandidates[1] || null;
      const clearlyBestCompleted = !runnerUpCompleted || (bestCompleted.score - runnerUpCompleted.score >= 15);
      if (bestCompleted && bestCompleted.score >= 45 && clearlyBestCompleted) {
        duplicates.push({
          fileId: file.id,
          fileName: file.name,
          mimeType: file.mimeType || "",
          assetUrl: cleanSheetWhitespace(file.webViewLink) || `https://drive.google.com/file/d/${encodeURIComponent(file.id)}/view`,
          assetPreviewUrl: cleanSheetWhitespace(file.thumbnailLink) || buildDrivePreviewUrl(file.id),
          graphicsRequestId: bestCompleted.record.graphicsRequestId,
          queueSheetRow: bestCompleted.record.queueSheetRow || "",
          recordId: bestCompleted.record.recordId || "",
          author: bestCompleted.record.author,
          poemTitle: bestCompleted.record.poemTitle,
          bookTitle: bestCompleted.record.bookTitle,
          quoteText: bestCompleted.record.quoteText || "",
          reason: "Already returned to Weaver under another matching graphic file.",
          existingAssetUrl: bestCompleted.record.assetLinkUrl || "",
          existingAssetPreviewUrl: bestCompleted.record.assetPreviewUrl || ""
        });
        return;
      }
      unmatched.push({
        fileId: file.id,
        fileName: file.name,
        mimeType: file.mimeType || "",
        assetUrl: cleanSheetWhitespace(file.webViewLink) || `https://drive.google.com/file/d/${encodeURIComponent(file.id)}/view`,
        assetPreviewUrl: cleanSheetWhitespace(file.thumbnailLink) || buildDrivePreviewUrl(file.id),
        reason: "No open queue row matched the filename safely."
      });
      return;
    }

    const best = candidates[0];
    const runnerUp = candidates[1] || null;
    const uniquePoemHit = candidates.filter(candidate => candidate.score >= 45).length === 1;
    const safeEnough =
      best.score >= 65 ||
      (best.score >= 45 && uniquePoemHit && (best.hasBook || best.hasAuthor || !runnerUp));
    const clearlyBest = !runnerUp || (best.score - runnerUp.score >= 15);

    if (!safeEnough || !clearlyBest) {
      unmatched.push({
        fileId: file.id,
        fileName: file.name,
        mimeType: file.mimeType || "",
        assetUrl: cleanSheetWhitespace(file.webViewLink) || `https://drive.google.com/file/d/${encodeURIComponent(file.id)}/view`,
        assetPreviewUrl: cleanSheetWhitespace(file.thumbnailLink) || buildDrivePreviewUrl(file.id),
        reason: "Filename matched too ambiguously to mark as made safely.",
        topScore: best.score,
        candidates: candidates.slice(0, 6).map(candidate => ({
          graphicsRequestId: candidate.record.graphicsRequestId,
          queueSheetRow: candidate.record.queueSheetRow,
          recordId: candidate.record.recordId,
          author: candidate.record.author,
          poemTitle: candidate.record.poemTitle,
          bookTitle: candidate.record.bookTitle,
          quoteText: candidate.record.quoteText,
          score: candidate.score,
          reasons: candidate.reasons
        }))
      });
      return;
    }

    if (usedRequestIds.has(best.record.graphicsRequestId)) {
      unmatched.push({
        fileId: file.id,
        fileName: file.name,
        reason: "Another file in this folder already matched the same request."
      });
      return;
    }

    usedRequestIds.add(best.record.graphicsRequestId);
    matches.push({
      fileId: file.id,
      fileName: file.name,
      mimeType: file.mimeType || "",
      assetUrl: cleanSheetWhitespace(file.webViewLink) || `https://drive.google.com/file/d/${encodeURIComponent(file.id)}/view`,
      assetPreviewUrl: cleanSheetWhitespace(file.thumbnailLink) || buildDrivePreviewUrl(file.id),
      graphicsRequestId: best.record.graphicsRequestId,
      queueSheetRow: best.record.queueSheetRow,
      recordId: best.record.recordId,
      author: best.record.author,
      poemTitle: best.record.poemTitle,
      bookTitle: best.record.bookTitle,
      quoteText: best.record.quoteText,
      reason: best.reasons.join(", ")
    });
  });

  const matchedRequestIds = new Set(matches.map(match => cleanSheetWhitespace(match.graphicsRequestId)).filter(Boolean));
  const unmatchedRequests = openRequests
    .filter(record => !matchedRequestIds.has(cleanSheetWhitespace(record.graphicsRequestId)))
    .map(record => ({
      graphicsRequestId: record.graphicsRequestId,
      queueSheetRow: record.queueSheetRow,
      recordId: record.recordId,
      author: record.author,
      poemTitle: record.poemTitle,
      bookTitle: record.bookTitle,
      quoteText: record.quoteText
    }));

  return { matches, unmatched, unmatchedRequests };
}

async function previewDriveFolderImport(folderUrl) {
  const folderId = extractGoogleDriveFolderId(folderUrl);
  const fileId = folderId ? "" : extractGoogleDriveFileId(folderUrl);
  if (!folderId && !fileId) {
    return { ok: false, error: "That does not look like a valid Google Drive folder or file link." };
  }

  let sourceType = "folder";
  let sourceId = folderId;
  let sourceName = "";
  let files = [];

  if (folderId) {
    const folder = await getDriveFolderMetadata(folderId);
    sourceName = cleanSheetWhitespace(folder?.name) || "";
    files = await listDriveFolderImageFiles(folderId);
  } else {
    sourceType = "file";
    sourceId = fileId;
    const file = await getDriveFileMetadata(fileId);
    if (!String(file?.mimeType || "").startsWith("image/")) {
      return {
        ok: false,
        error: "That Drive file is not an image, so Weaver cannot import it into Graphics QC."
      };
    }
    sourceName = cleanSheetWhitespace(file?.name) || "";
    files = [file];
  }

  const [openRequests, completionRows, metadata] = await Promise.all([
    getPigGraphicsRequests("all"),
    readPigCompletedGraphicsRows(),
    loadDriveImportBookMetadata()
  ]);
  const completedRecords = completionRows
    .map((row, index) => buildPigQcRecordFromSheetRow(row, index))
    .filter(Boolean);
  const { matches, unmatched, unmatchedRequests, duplicates } = chooseDriveFolderImportMatch(files, openRequests, metadata, completedRecords);

  return {
    ok: true,
    version: `${appVersion}-service-account`,
    sourceType,
    folderId: folderId || "",
    fileId: fileId || "",
    folderName: sourceName,
    folderUrl: cleanSheetWhitespace(folderUrl),
    imageCount: files.length,
    matches,
    duplicates,
    unmatched,
    unmatchedRequests
  };
}

async function applyDriveFolderImport(folderUrl, manualSelections = {}, options = {}) {
  const preview = await previewDriveFolderImport(folderUrl);
  if (!preview.ok) {
    return preview;
  }

  const manualMatches = [];
  const manualSelectionMap = manualSelections && typeof manualSelections === "object" ? manualSelections : {};
  const expectedSafeMatchCount = Math.max(0, parseInt(options.expectedSafeMatchCount, 10) || 0);
  const attemptedCount = expectedSafeMatchCount + Object.keys(manualSelectionMap).filter(key => cleanSheetWhitespace(manualSelectionMap[key])).length;
  const usedRequestIds = new Set((preview.matches || []).map(match => cleanSheetWhitespace(match.graphicsRequestId)).filter(Boolean));
  const unmatchedRequestById = new Map(
    (preview.unmatchedRequests || [])
      .map(request => [cleanSheetWhitespace(request.graphicsRequestId), request])
      .filter(([requestId]) => requestId)
  );
  (preview.unmatched || []).forEach(item => {
    const selectedRequestId = cleanSheetWhitespace(manualSelectionMap[item.fileId]);
    if (!selectedRequestId || usedRequestIds.has(selectedRequestId)) return;
    let selected = null;
    if (Array.isArray(item.candidates)) {
      selected = item.candidates.find(candidate => cleanSheetWhitespace(candidate.graphicsRequestId) === selectedRequestId) || null;
    }
    if (!selected) {
      selected = unmatchedRequestById.get(selectedRequestId) || null;
    }
    if (!selected) return;
    usedRequestIds.add(selectedRequestId);
    manualMatches.push({
      fileId: item.fileId,
      fileName: item.fileName,
      mimeType: item.mimeType || "",
      assetUrl: item.assetUrl || "",
      assetPreviewUrl: item.assetPreviewUrl || "",
      graphicsRequestId: selected.graphicsRequestId,
      queueSheetRow: selected.queueSheetRow,
      recordId: selected.recordId || "",
      author: selected.author,
      poemTitle: selected.poemTitle,
      bookTitle: selected.bookTitle,
      quoteText: selected.quoteText || "",
      reason: "manual selection"
    });
  });

  const metadata = await loadDriveImportBookMetadata().catch(() => ({
    shortenerByBookKey: new Map()
  }));
  const allMatches = [...(preview.matches || []), ...manualMatches];

  if (!allMatches.length) {
    return {
      ok: false,
      error: "No safe or manually confirmed Drive-link matches were found to import."
    };
  }

  const completions = [];
  for (const match of allMatches) {
    const canonicalLabel = buildGraphicsImportLabel(match, metadata.shortenerByBookKey);
    const canonicalFileName = `${canonicalLabel}.png`;
    let finalFileName = cleanSheetWhitespace(match.fileName) || canonicalFileName;

    if (cleanSheetWhitespace(match.fileId) && finalFileName !== canonicalFileName) {
      try {
        await updateDriveFileName(match.fileId, canonicalFileName);
        finalFileName = canonicalFileName;
      } catch {
        finalFileName = cleanSheetWhitespace(match.fileName) || canonicalFileName;
      }
    }

    completions.push({
      completionId: buildDriveFolderImportCompletionId(match.fileId),
      requestId: match.graphicsRequestId,
      author: match.author,
      poemTitle: match.poemTitle,
      bookTitle: match.bookTitle,
      quoteText: match.quoteText,
      assetUrl: match.assetUrl,
      assetPreviewUrl: match.assetPreviewUrl,
      sourceRecordId: match.recordId,
      sourceSheetRow: match.queueSheetRow,
      productionNotes: `Imported from Drive folder "${preview.folderName || preview.folderId}" as "${finalFileName}"`,
      completedAt: new Date().toISOString(),
      sourceTool: "Drive folder import"
    });
  }

  const result = await upsertPigCompletedGraphics(completions);
  return {
    ...result,
    attemptedCount,
    importedCount: completions.length,
    skippedCount: Math.max(0, attemptedCount - completions.length),
    manualMatchCount: manualMatches.length,
    folderId: preview.folderId,
    folderName: preview.folderName
  };
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

function buildCatalogValidationPayload(row) {
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

function buildPendingRecordFromSheetRow(row, index, canonicalBookAuthorMap = null) {
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
  const noteMeta = parseIntakeMetadataFromNotes(row[8] || "");
  const bookMeta = resolvePublishingBookMeta(bookTitle, bookTitle);

  if (!bookTitle || !cleanedExcerptText || excluded || !isPendingSheetReview(reviewDecision)) {
    return null;
  }

  return {
    sourceRow: SHEET_SOURCE_CONFIG.startRow + index,
    recordId: (row[config.recordId - 1] || "").toString(),
    intakeMode: isVideoIntake ? "video" : "book",
    intakeLabel: cleanSheetWhitespace(row[2]),
    author: resolveGraphicsAuthor(row[config.author - 1] || rawVideoAuthor, bookTitle, canonicalBookAuthorMap),
    title: (row[config.title - 1] || rawVideoTitle).toString(),
    bookTitle,
    excerptText,
    releaseCatalog: noteMeta.releaseCatalog || cleanSheetWhitespace(bookMeta?.releaseCatalog),
    bookShortener: noteMeta.bookShortener || cleanSheetWhitespace(bookMeta?.bookShortener),
    contentType: noteMeta.contentType || "EXC",
    socialMediaHandle: noteMeta.socialMediaHandle || "",
    bookPrimarySourceFormat: cleanSheetWhitespace(row[config.validationPrimarySourceFormat - 1]),
    catalogValidation: buildCatalogValidationPayload(row)
  };
}

function buildApprovedExcerptExportRecordFromSheetRow(row, index, canonicalBookAuthorMap = null) {
  const config = SHEET_SOURCE_CONFIG.columnMap;
  const isVideoIntake = cleanSheetWhitespace(row[2]).toLowerCase() === "add a quote from a video";
  const rawVideoAuthor = isVideoIntake ? (row[9] || "").toString() : "";
  const rawVideoTitle = isVideoIntake ? (row[11] || "").toString() : "";
  const rawVideoExcerpt = isVideoIntake ? (row[12] || "").toString() : "";
  const rawVideoBookTitle = isVideoIntake ? cleanSheetWhitespace(row[13] || "") : "";
  const excerptText = (isVideoIntake ? rawVideoExcerpt : row[config.excerpt - 1] || "").toString();
  const cleanedExcerptText = cleanSheetWhitespace(excerptText);
  const excluded = isSheetYes(row[config.exclude - 1]);
  const reviewDecision = getSheetExcerptReviewDecision(row);
  const explicitDecision = cleanSheetWhitespace(row[config.excerptReviewDecision - 1]).toUpperCase();
  const bookTitle = isVideoIntake
    ? rawVideoBookTitle
    : cleanSheetWhitespace(row[config.bookTitle - 1]);
  const approvedForUse = (row[config.approved - 1] || "").toString().trim().toUpperCase() === "Y";
  const approvedForExcerpt = isAcceptedExcerptReviewDecision(explicitDecision) || approvedForUse;
  const noteMeta = parseIntakeMetadataFromNotes(row[8] || "");

  if ((!bookTitle && !isVideoIntake) || !cleanedExcerptText || excluded || !approvedForExcerpt) {
    return null;
  }
  if (reviewDecision === "REJECT" || reviewDecision === "NEEDS_CORRECTION") {
    return null;
  }

  const sourceRow = SHEET_SOURCE_CONFIG.startRow + index;
  const sourceRecordId = cleanSheetWhitespace(row[config.recordId - 1]) || `weaver:row-${sourceRow}`;
  const author = resolveGraphicsAuthor(
    isVideoIntake ? rawVideoAuthor : row[config.author - 1] || "",
    bookTitle,
    canonicalBookAuthorMap
  );
  const poemTitle = (isVideoIntake ? rawVideoTitle : row[config.title - 1] || "").toString();
  const sourceEvent = cleanSheetWhitespace(row[17]) || noteMeta.sourceEvent || null;
  const validationStatus = cleanSheetWhitespace(row[config.validationStatus - 1]);
  const canonicalAuthor = cleanSheetWhitespace(row[config.validationCanonicalAuthor - 1]) || author;
  const canonicalPoemTitle = cleanSheetWhitespace(row[config.validationMatchedPoemTitle - 1]) || poemTitle;
  const canonicalBookTitle = cleanSheetWhitespace(row[config.validationCanonicalBook - 1]) || bookTitle;
  const duplicateGroupId = (row[config.duplicateGroupId - 1] || "").toString();
  const createdTimestamp = cleanSheetWhitespace(row[0]);
  const normalizedTimestamp = parseChicagoSheetTimestampToIso(createdTimestamp);
  const lineCount = excerptText
    ? excerptText.split(/\r?\n/).map(line => line.trim()).filter(Boolean).length
    : 0;
  const correctionApplied = Boolean(
    cleanSheetWhitespace(row[config.correctedAuthor - 1])
    || cleanSheetWhitespace(row[config.correctedTitle - 1])
    || cleanSheetWhitespace(row[config.correctedBookTitle - 1])
    || cleanSheetWhitespace(row[config.correctedExcerpt - 1])
  );
  const bookMeta = resolvePublishingBookMeta(bookTitle, canonicalBookTitle);
  const poetryPleaseStatus = cleanSheetWhitespace(row[config.excerptPoetryPleaseStatus - 1]);
  const poetryPleaseUpdatedAt = cleanSheetWhitespace(row[config.excerptPoetryPleaseUpdatedAt - 1]);
  const poetryPleaseNote = String(row[config.excerptPoetryPleaseNote - 1] || "");

  return {
    sourceKind: "weaver",
    sourceRecordId,
    sourceRow,
    sourceApprovedAt: normalizedTimestamp,
    sourceUpdatedAt: normalizedTimestamp,
    sourceEvent,
    sourceEventLabel: sourceEvent,
    author,
    poemTitle,
    bookTitle,
    excerptText,
    approval: {
      reviewDecision: isAcceptedExcerptReviewDecision(explicitDecision) || approvedForUse ? "approve" : explicitDecision.toLowerCase(),
      approvedForUse: approvedForExcerpt,
      approvedForQuoteImage: approvedForUse,
      approvedForGraphics: approvedForUse
    },
    status: {
      excluded: false,
      needsCorrection: false,
      correctionApplied,
      validationStatus,
      duplicateGroupId
    },
    canonical: {
      canonicalAuthor,
      canonicalPoemTitle,
      canonicalBookTitle,
      catalogMatchId: "",
      libraryMatchId: ""
    },
    metadata: {
      wordCount: cleanedExcerptText ? cleanedExcerptText.split(/\s+/).length : 0,
      lineCount,
      updatedAt: normalizedTimestamp,
      sourceEvent,
      sourceEventLabel: sourceEvent
    },
    contentType: noteMeta.contentType || "EXC",
    socialMediaHandle: noteMeta.socialMediaHandle || "",
    instagramHandle: noteMeta.socialMediaHandle || "",
    igHandle: noteMeta.socialMediaHandle || "",
    bookShortener: noteMeta.bookShortener || bookMeta?.bookShortener || "",
    bookLink: "",
    releaseCatalog: noteMeta.releaseCatalog || bookMeta?.releaseCatalog || "",
    driveLink: "",
    sourceUrl: "",
    poetryPleaseStatus,
    poetryPleaseUpdatedAt,
    poetryPleaseNote,
    sourcePayload: {}
  };
}

function buildCorrectionRecordFromSheetRow(row, index, canonicalBookAuthorMap = null) {
  const config = SHEET_SOURCE_CONFIG.columnMap;
  const rawExcerptText = (row[config.excerpt - 1] || "").toString();
  const correctedExcerptText = (row[config.correctedExcerpt - 1] || "").toString();
  const excerptText = correctedExcerptText || rawExcerptText;
  const cleanedExcerptText = cleanSheetWhitespace(excerptText);
  const reviewDecision = getSheetExcerptReviewDecision(row);
  const rawBookTitle = cleanSheetWhitespace(row[config.bookTitle - 1]);
  const correctedBookTitle = (row[config.correctedBookTitle - 1] || "").toString();
  const bookTitle = cleanSheetWhitespace(correctedBookTitle || rawBookTitle);

  if (!bookTitle || !cleanedExcerptText || reviewDecision !== "NEEDS_CORRECTION" || isSheetYes(row[config.exclude - 1])) {
    return null;
  }

  const rawAuthor = resolveGraphicsAuthor(row[config.author - 1] || "", rawBookTitle, canonicalBookAuthorMap);
  const rawTitle = (row[config.title - 1] || "").toString();
  const correctedAuthor = (row[config.correctedAuthor - 1] || "").toString();
  const correctedTitle = (row[config.correctedTitle - 1] || "").toString();

  return {
    sourceRow: SHEET_SOURCE_CONFIG.startRow + index,
    recordId: (row[config.recordId - 1] || "").toString(),
    author: correctedAuthor || rawAuthor,
    title: correctedTitle || rawTitle,
    bookTitle,
    excerptText,
    wordCount: cleanedExcerptText ? cleanedExcerptText.split(/\s+/).length : 0,
    approved: (row[config.approved - 1] || "").toString(),
    statusIndicator: cleanSheetWhitespace(row[config.statusIndicator - 1]),
    quoteCreatedQc: (row[config.quoteCreatedQc - 1] || "").toString(),
    correctionNote: (row[config.correctionNote - 1] || "").toString(),
    excludeRaw: (row[config.exclude - 1] || "").toString(),
    duplicateGroupId: (row[config.duplicateGroupId - 1] || "").toString(),
    exactPullCount: parseSheetInteger(row[config.exactPullCount - 1]),
    pending: false,
    excerptReviewDecision: reviewDecision,
    useForQi: (row[config.approved - 1] || "").toString() === "Y",
    useForInt: isSheetYes(row[config.useForInt - 1]),
    useForGraphicsQi: (row[config.approved - 1] || "").toString() === "Y",
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
    catalogValidation: buildCatalogValidationPayload(row)
  };
}

function summarizeBooks(records) {
  const byKey = new Map();

  records.forEach(record => {
    const title = cleanSheetWhitespace(record.bookTitle);
    const key = normalizeBookKey(title);
    if (!key) return;
    if (!byKey.has(key)) {
      byKey.set(key, { title, count: 0 });
    }
    const summary = byKey.get(key);
    summary.title = choosePreferredBookTitle(summary.title, title);
    summary.count += 1;
  });

  return Array.from(byKey.entries())
    .map(([key, summary]) => ({ key, title: summary.title, count: summary.count }))
    .sort((left, right) => left.title.localeCompare(right.title));
}

function formatChicagoTimestamp(date = new Date()) {
  const formatter = new Intl.DateTimeFormat("en-US", {
    timeZone: "America/Chicago",
    year: "numeric",
    month: "numeric",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
    second: "2-digit",
    hour12: false
  });

  const parts = Object.fromEntries(
    formatter.formatToParts(date)
      .filter(part => part.type !== "literal")
      .map(part => [part.type, part.value])
  );

  return `${parts.month}/${parts.day}/${parts.year} ${parts.hour}:${parts.minute}:${parts.second}`;
}

function formatDateInChicagoParts(date) {
  const formatter = new Intl.DateTimeFormat("en-US", {
    timeZone: "America/Chicago",
    year: "numeric",
    month: "numeric",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
    second: "2-digit",
    hour12: false
  });

  return Object.fromEntries(
    formatter.formatToParts(date)
      .filter(part => part.type !== "literal")
      .map(part => [part.type, part.value])
  );
}

function parseChicagoSheetTimestampToIso(value) {
  const text = cleanSheetWhitespace(value);
  if (!text) return "";

  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})$/);
  if (!match) return text;

  const [, month, day, year, hour, minute, second] = match;
  const mm = month.padStart(2, "0");
  const dd = day.padStart(2, "0");
  const hh = hour.padStart(2, "0");
  const offsets = ["-06:00", "-05:00"];

  for (const offset of offsets) {
    const candidate = new Date(`${year}-${mm}-${dd}T${hh}:${minute}:${second}${offset}`);
    if (Number.isNaN(candidate.getTime())) continue;
    const parts = formatDateInChicagoParts(candidate);
    if (
      parts.year === year
      && parts.month === String(Number(month))
      && parts.day === String(Number(day))
      && parts.hour === String(Number(hour))
      && parts.minute === minute
      && parts.second === second
    ) {
      return candidate.toISOString();
    }
  }

  return text;
}

function toUniqueSortedValues(values) {
  return Array.from(new Set(values.map(value => cleanSheetWhitespace(value)).filter(Boolean)))
    .sort((left, right) => left.localeCompare(right));
}

function normalizeAuthorSuggestionKey(value) {
  return cleanSheetWhitespace(value)
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[’‘]/g, "'")
    .replace(/[“”]/g, "\"")
    .replace(/\s*&\s*/g, " and ")
    .replace(/,\s*age\s+\d+\b/gi, "")
    .replace(/\b(age)\s+\d+\b/gi, "")
    .replace(/[.,()]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function getAuthorSuggestionScore(value, sourcePriority = 1) {
  const text = cleanSheetWhitespace(value);
  if (!text) return -Infinity;

  let score = sourcePriority * 100;
  if (/[a-z]/.test(text) && /[A-Z]/.test(text)) score += 8;
  if (!/[a-z]/.test(text) && /[A-Z]/.test(text)) score -= 6;
  if (/^[A-Z][a-z]+(?:[\s'".-][A-Za-z][A-Za-z'".-]*)*$/.test(text)) score += 6;
  if (/, age \d+/i.test(text) || /\bage \d+\b/i.test(text)) score -= 8;
  if (/["“”]/.test(text)) score -= 6;
  if (/\d/.test(text)) score -= 4;
  if (/^[—–-]/.test(text)) score -= 20;
  if (/\b(curation batch|read by|written by)\b/i.test(text)) score -= 20;
  if (/&/.test(text)) score -= 2;
  if (/\band\b/i.test(text)) score += 1;
  if (/[()]/.test(text)) score -= 2;
  if (text.length > 40) score -= 2;
  return score;
}

function cleanAuthorSuggestion(value, { bookTitle = "" } = {}) {
  let cleaned = cleanSheetWhitespace(value)
    .replace(/^[—–-]+\s*/, "")
    .replace(/[’‘]/g, "'")
    .replace(/[“”]/g, "\"")
    .replace(/,\s*age\s+\d+\b/gi, "")
    .replace(/\b(age)\s+\d+\b/gi, "")
    .trim();

  const normalizedBook = normalizeBookKey(bookTitle)
    .replace(/\s*:\s*/g, ": ")
    .replace(/\s+/g, " ")
    .trim();
  if (normalizedBook.startsWith("poetry by chance")) {
    return "Taylor Mali";
  }
  if (normalizedBook === "tooth gaps in the archives" && normalizeAuthorSuggestionKey(cleaned).startsWith("francis")) {
    return "Francis Dylan Waguespack";
  }

  const authorAliases = new Map([
    ["read by taylor mali written bya 17 y/o named kohana", "Taylor Mali"]
  ]);
  const blockedAuthors = new Set([
    "addrianna allred addy lugo",
    "adrianna allred addy lugo"
  ]);

  const lower = cleaned.toLowerCase();
  if (!cleaned) return "";
  if (["n/a", "na", "other"].includes(lower)) return "";
  if (lower.startsWith("tab before")) return "";
  if (["authors multiple", "group piece", "poetryteacher"].includes(lower)) return "";
  if (lower.startsWith("group ")) return "";
  if (lower.startsWith("group piece")) return "";
  if (/\bcuration batch\b/i.test(cleaned)) return "";
  if (/\bread by\b/i.test(cleaned) || /\bwritten by\b/i.test(cleaned)) return "";
  if (cleaned.startsWith("\"") && cleaned.split(/\s+/).length > 4) return "";
  if (/[.!?].*[.!?]/.test(cleaned)) return "";
  if (cleaned.length > 80) return "";
  if (blockedAuthors.has(normalizeAuthorSuggestionKey(cleaned))) return "";

  const aliased = authorAliases.get(normalizeAuthorSuggestionKey(cleaned));
  if (aliased) return aliased;

  return cleaned;
}

function buildCleanAuthorSuggestions({ sourceRows = [], catalogBooks = [], publishingBooks = [] } = {}) {
  const byKey = new Map();
  const canonicalAuthorByBookKey = new Map();

  function registerCanonicalBookAuthor(bookTitle, author) {
    const normalizedTitle = normalizeBookKey(bookTitle);
    const cleanedAuthor = cleanAuthorSuggestion(author, { bookTitle });
    if (!normalizedTitle || !cleanedAuthor) return;
    canonicalAuthorByBookKey.set(normalizedTitle, cleanedAuthor);
  }

  publishingBooks.forEach(book => registerCanonicalBookAuthor(book.title, book.author));
  catalogBooks.forEach(book => registerCanonicalBookAuthor(book.title, book.author));

  function upsert(author, sourcePriority, bookTitle = "") {
    const canonicalForBook = canonicalAuthorByBookKey.get(normalizeBookKey(bookTitle));
    const cleaned = cleanAuthorSuggestion(canonicalForBook || author, { bookTitle });
    const key = normalizeAuthorSuggestionKey(cleaned);
    if (!key) return;

    const nextScore = getAuthorSuggestionScore(cleaned, sourcePriority);
    const current = byKey.get(key);
    if (!current || nextScore > current.score) {
      byKey.set(key, {
        value: cleaned,
        score: nextScore
      });
    }
  }

  sourceRows.forEach(row => upsert(row.author, 1, row.bookTitle));
  publishingBooks.forEach(book => upsert(book.author, 3, book.title));
  catalogBooks.forEach(book => upsert(book.author, 2, book.title));

  return Array.from(byKey.values())
    .map(entry => entry.value)
    .sort((left, right) => left.localeCompare(right));
}

async function getPublishingOrderBooks() {
  const range = `'${publishingOrderSheetName.replace(/'/g, "''")}'!A3:F`;
  const values = await fetchSheetValuesServer(range, {
    targetSpreadsheetId: publishingOrderSpreadsheetId
  });

  const books = values
    .map(row => {
      const lastName = cleanSheetWhitespace(row[0]);
      const firstName = cleanSheetWhitespace(row[1]);
      const releaseCatalog = cleanSheetWhitespace(row[2]);
      const bookShortener = cleanSheetWhitespace(row[3]);
      const bookTitle = cleanSheetWhitespace(row[4]);
      const pubDate = cleanSheetWhitespace(row[5]);
      const author = [firstName, lastName].filter(Boolean).join(" ").trim();
      if (!bookTitle) return null;
      return {
        title: bookTitle,
        author,
        bookShortener,
        releaseCatalog,
        pubDate
      };
    })
    .filter(Boolean);

  books.sort((left, right) => {
    const leftDate = left.pubDate || "";
    const rightDate = right.pubDate || "";
    if (leftDate && rightDate && leftDate !== rightDate) {
      return leftDate.localeCompare(rightDate);
    }
    return left.title.localeCompare(right.title);
  });

  return books;
}

function runIntakeCatalogLookup(payload) {
  return new Promise((resolve, reject) => {
    const child = spawn("python3", [path.join(__dirname, "intake_catalog_lookup.py")], {
      env: {
        ...process.env
      }
    });
    let stdout = "";
    let stderr = "";

    child.stdout.on("data", chunk => {
      stdout += chunk.toString();
    });
    child.stderr.on("data", chunk => {
      stderr += chunk.toString();
    });
    child.on("error", reject);
    child.on("close", code => {
      if (code !== 0) {
        reject(new Error(stderr || `intake_catalog_lookup.py exited with code ${code}`));
        return;
      }
      try {
        resolve(JSON.parse(stdout));
      } catch (error) {
        reject(error);
      }
    });

    child.stdin.write(JSON.stringify(payload || {}));
    child.stdin.end();
  });
}

async function getExcerptGatheringOptionsFromSheets() {
  const range = `'${sourceSheetName.replace(/'/g, "''")}'!Z2:AC`;
  const values = await fetchSheetValuesServer(range);
  const sourceRows = [];
  const books = [];

  values.forEach(row => {
    sourceRows.push({
      author: row[0] || "",
      bookTitle: row[3] || ""
    });
    books.push(row[3] || "");
  });

  const catalogResult = await runIntakeCatalogLookup({ action: "books" });
  const catalogBooks = Array.isArray(catalogResult?.books) ? catalogResult.books : [];
  let publishingBooks = [];
  let publishingBooksError = "";
  try {
    publishingBooks = await getPublishingOrderBooks();
  } catch (error) {
    publishingBooksError = error.message;
  }

  return {
    ok: true,
    version: `${appVersion}-service-account`,
    authors: buildCleanAuthorSuggestions({
      sourceRows,
      catalogBooks,
      publishingBooks
    }),
    books: toUniqueSortedValues(books),
    catalogBooks,
    publishingBooks,
    publishingBooksError,
    fixPartOptions: [
      "The Quote",
      "The Poem Name",
      "The Author Name",
      "The Book Title",
      "Other"
    ]
  };
}

async function getExcerptGatheringPoemsForBook(bookTitle) {
  const result = await runIntakeCatalogLookup({
    action: "poems",
    bookTitle
  });
  return {
    ok: true,
    version: `${appVersion}-service-account`,
    ...result
  };
}

function buildExcerptGatheringAppendRow(payload = {}) {
  const mode = cleanSheetWhitespace(payload.mode).toLowerCase();
  const email = cleanSheetWhitespace(payload.email);
  if (!email) {
    throw new Error("Email is required.");
  }

  const row = new Array(19).fill("");
  row[0] = formatChicagoTimestamp();
  row[1] = email;

  if (mode === "book") {
    const author = cleanSheetWhitespace(payload.author);
    const title = cleanSheetWhitespace(payload.title);
    const quote = String(payload.quote || "").trim();
    if (!author || !title || !quote) {
      throw new Error("Book intake requires author, poem title, and quote.");
    }
    row[2] = "Add a quote from a book";
    row[3] = author;
    row[5] = title;
    row[6] = quote;
    row[7] = cleanSheetWhitespace(payload.bookTitle);
    row[8] = String(payload.notes || "").trim();
    return row;
  }

  if (mode === "video") {
    const author = cleanSheetWhitespace(payload.author);
    const quote = String(payload.quote || "").trim();
    if (!author || !quote) {
      throw new Error("Video intake requires author and quote.");
    }
    row[2] = "Add a quote from a video";
    row[9] = author;
    row[11] = cleanSheetWhitespace(payload.title);
    row[12] = quote;
    row[13] = cleanSheetWhitespace(payload.bookTitle);
    row[17] = cleanSheetWhitespace(payload.eventName);
    return row;
  }

  throw new Error("Unsupported excerpt gathering mode.");
}

async function appendExcerptGatheringRow(payload = {}) {
  const rowValues = buildExcerptGatheringAppendRow(payload);
  const response = await appendSheetValuesServer(`'${sourceSheetName.replace(/'/g, "''")}'!A:S`, [rowValues]);
  const updatedRange = cleanSheetWhitespace(response?.updates?.updatedRange || "");
  const rowMatch = updatedRange.match(/![A-Z]+(\d+):/);
  const rowNumber = rowMatch ? Number(rowMatch[1]) : 0;
  const mode = cleanSheetWhitespace(payload.mode).toLowerCase();

  if (rowNumber && (mode === "book" || mode === "video")) {
    const config = SHEET_SOURCE_CONFIG.columnMap;
    const writes = [
      {
        range: `'${sourceSheetName.replace(/'/g, "''")}'!${toA1Column(config.author)}${rowNumber}`,
        values: [[cleanSheetWhitespace(payload.author)]]
      },
      {
        range: `'${sourceSheetName.replace(/'/g, "''")}'!${toA1Column(config.title)}${rowNumber}`,
        values: [[cleanSheetWhitespace(payload.title)]]
      },
      {
        range: `'${sourceSheetName.replace(/'/g, "''")}'!${toA1Column(config.excerpt)}${rowNumber}`,
        values: [[String(payload.quote || "").trim()]]
      },
      {
        range: `'${sourceSheetName.replace(/'/g, "''")}'!${toA1Column(config.bookTitle)}${rowNumber}`,
        values: [[cleanSheetWhitespace(payload.bookTitle)]]
      }
    ];
    const sourceRecordId = cleanSheetWhitespace(payload.recordId);
    if (sourceRecordId) {
      writes.unshift({
        range: `'${sourceSheetName.replace(/'/g, "''")}'!${toA1Column(config.recordId)}${rowNumber}`,
        values: [[sourceRecordId]]
      });
    }
    await batchUpdateSheetValuesServer(writes);
  }

  return {
    ok: true,
    version: `${appVersion}-service-account`,
    rowNumber,
    intakeMode: mode,
    updatedRange
  };
}

async function getPendingRecordsFromSheets() {
  return getCachedQueueSnapshot("pending-records", async () => {
    const values = await getSourceSheetValuesCached();
    const canonicalBookAuthorMap = await getCanonicalGraphicsBookAuthorMap().catch(() => new Map());
    const records = values
      .map((row, index) => buildPendingRecordFromSheetRow(row, index, canonicalBookAuthorMap))
      .filter(Boolean);

    return {
      ok: true,
      version: `${appVersion}-service-account`,
      records
    };
  });
}

async function getPendingExcerptsForBookFromSheets(bookTitle) {
  const requestedKey = normalizeBookKey(bookTitle);
  const values = await getSourceSheetValuesCached();
  const canonicalBookAuthorMap = await getCanonicalGraphicsBookAuthorMap().catch(() => new Map());
  const excerpts = values
    .map((row, index) => buildPendingRecordFromSheetRow(row, index, canonicalBookAuthorMap))
    .filter(record => record && normalizeBookKey(record.bookTitle) === requestedKey);

  const preferredBookTitle = excerpts.reduce(
    (current, record) => choosePreferredBookTitle(current, record.bookTitle),
    cleanSheetWhitespace(bookTitle)
  );

  return {
    ok: true,
    version: `${appVersion}-service-account`,
    bookTitle: preferredBookTitle,
    excerpts
  };
}

let publishingBooksCache = null;

async function getPublishingBooksByKey() {
  if (publishingBooksCache) return publishingBooksCache;
  try {
    const books = await getPublishingOrderBooks();
    publishingBooksCache = new Map(
      books.map(book => [normalizeBookKey(book.title), book]).filter(([key]) => key)
    );
  } catch {
    publishingBooksCache = new Map();
  }
  return publishingBooksCache;
}

async function getApprovedExcerptsExportFromSheets({ since = "" } = {}) {
  await getPublishingBooksByKey();
  const sinceValue = cleanSheetWhitespace(since);
  const values = await getSourceSheetValuesCached();
  const canonicalBookAuthorMap = await getCanonicalGraphicsBookAuthorMap().catch(() => new Map());
  const records = values
    .map((row, index) => buildApprovedExcerptExportRecordFromSheetRow(row, index, canonicalBookAuthorMap))
    .filter(Boolean)
    .filter(record => !sinceValue || cleanSheetWhitespace(record.metadata?.updatedAt) >= sinceValue);

  return {
    ok: true,
    version: `${appVersion}-service-account`,
    sourceImportId: `weaver-approved-sync-${new Date().toISOString()}`,
    records
  };
}

async function getCorrectionBooksFromSheets() {
  const values = await getSourceSheetValuesCached();
  const canonicalBookAuthorMap = await getCanonicalGraphicsBookAuthorMap().catch(() => new Map());
  const records = values
    .map((row, index) => buildCorrectionRecordFromSheetRow(row, index, canonicalBookAuthorMap))
    .filter(Boolean);

  return {
    ok: true,
    version: `${appVersion}-service-account`,
    books: summarizeBooks(records)
  };
}

async function getCorrectionsForBookFromSheets(bookTitle) {
  const requestedKey = normalizeBookKey(bookTitle);
  const range = `'${sourceSheetName.replace(/'/g, "''")}'!A2:BK`;
  const values = await fetchSheetValuesServer(range);
  const canonicalBookAuthorMap = await getCanonicalGraphicsBookAuthorMap().catch(() => new Map());
  const excerpts = values
    .map((row, index) => buildCorrectionRecordFromSheetRow(row, index, canonicalBookAuthorMap))
    .filter(record => record && normalizeBookKey(record.bookTitle) === requestedKey);

  const preferredBookTitle = excerpts.reduce(
    (current, record) => choosePreferredBookTitle(current, record.bookTitle),
    cleanSheetWhitespace(bookTitle)
  );

  return {
    ok: true,
    version: `${appVersion}-service-account`,
    bookTitle: preferredBookTitle,
    excerpts
  };
}

function getGraphicsSheetName(mode) {
  return cleanSheetWhitespace(mode).toLowerCase() === "cleanup"
    ? graphicsCleanupSheetName
    : graphicsQueueSheetName;
}

async function getGraphicsQcStateMapFromSheets() {
  await ensureCleanupQcColumnsServer();
  const range = `'${graphicsCleanupSheetName.replace(/'/g, "''")}'!J2:L`;
  const values = await fetchSheetValuesServer(range);
  const state = new Map();

  values.forEach((row, index) => {
    state.set(String(index + 2), {
      decision: cleanSheetWhitespace(row[0]),
      note: (row[1] || "").toString(),
      updatedAt: cleanSheetWhitespace(row[2])
    });
  });

  return state;
}

async function getPigCompletedGraphicsBooks() {
  const rows = await readPigCompletedGraphicsRows();
  const canonicalBookAuthorMap = await getCanonicalGraphicsBookAuthorMap();
  const runtimeState = await getRuntimeGraphicsState(rows.map(row => row[PIG_COMPLETION_COLUMNS.completionId - 1]));
  const handoffState = await getRuntimeGraphicsHandoffState(
    rows.map(row => cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.requestId - 1]))
  );
  const records = rows
    .map((row, index) => overlayRuntimeHandoffState(
      overlayRuntimeGraphicsState(buildPigQcRecordFromSheetRow(row, index, canonicalBookAuthorMap), runtimeState),
      handoffState
    ))
    .filter(Boolean);

  return summarizeBooks(records);
}

async function getPigCompletedGraphicsRecordsForBook(bookTitle) {
  const requestedKey = normalizeBookKey(bookTitle);
  const rows = await readPigCompletedGraphicsRows();
  const canonicalBookAuthorMap = await getCanonicalGraphicsBookAuthorMap();
  const runtimeState = await getRuntimeGraphicsState(rows.map(row => row[PIG_COMPLETION_COLUMNS.completionId - 1]));
  const handoffState = await getRuntimeGraphicsHandoffState(
    rows.map(row => cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.requestId - 1]))
  );
  return rows
    .map((row, index) => overlayRuntimeHandoffState(
      overlayRuntimeGraphicsState(buildPigQcRecordFromSheetRow(row, index, canonicalBookAuthorMap), runtimeState),
      handoffState
    ))
    .filter(record => record && normalizeBookKey(record.bookTitle) === requestedKey);
}

function buildPoetryPleaseGraphicRecord(record) {
  const qcDecision = cleanSheetWhitespace(record.graphicsQcDecision).toUpperCase();
  const assetUrl = cleanSheetWhitespace(record.assetLinkUrl);
  if (qcDecision !== "APPROVE" || !assetUrl) {
    return null;
  }
  const bookMeta = resolvePublishingBookMeta(record.bookTitle, record.bookTitle);
  const bookShortener = cleanSheetWhitespace(record.bookShortener) || cleanSheetWhitespace(bookMeta?.bookShortener);
  const releaseCatalog = cleanSheetWhitespace(record.releaseCatalog) || cleanSheetWhitespace(bookMeta?.releaseCatalog);
  const sourceTool = cleanSheetWhitespace(record.sourceTool || "P.I.G.");
  const driveFileId = extractGoogleDriveFileId(assetUrl || record.assetPreviewUrl || "");
  const proxyAssetUrl = driveFileId
    ? `${weaverPublicBaseUrl.replace(/\/$/, "")}/api/drive-image?fileId=${encodeURIComponent(driveFileId)}`
    : "";
  const handoffAssetUrl = sourceTool === "Weaver replacement" && proxyAssetUrl
    ? proxyAssetUrl
    : assetUrl;
  const handoffPreviewUrl = sourceTool === "Weaver replacement" && proxyAssetUrl
    ? proxyAssetUrl
    : cleanSheetWhitespace(record.assetPreviewUrl);
  const contentType = inferPoetryPleaseContentType(record, "");
  if (!contentType) {
    throw new Error(`Approved graphic ${cleanSheetWhitespace(record.graphicsRequestId) || "without an ID"} is missing contentType`);
  }

  if (contentType === "FPI") {
    const fpiSourceRecordId = bookShortener && cleanSheetWhitespace(record.poemTitle || record.title)
      ? `${bookShortener}-FPI-${slugToken(record.poemTitle || record.title)}`
      : cleanSheetWhitespace(record.sourceRecordId || record.recordId) || `pig:${cleanSheetWhitespace(record.pigCompletionId)}`;
    return {
      contentType: "FPI",
      sourceSystem: "weaver",
      sourceRecordId: fpiSourceRecordId,
      bookShortener,
      author: String(record.author || ""),
      book: String(record.bookTitle || ""),
      title: String(record.poemTitle || ""),
      releaseCatalog,
      driveLink: handoffAssetUrl,
      imageUrl: handoffPreviewUrl,
      pageNumber: cleanSheetWhitespace(record.pageNumber),
      ocrText: String(record.ocrText || record.quoteText || ""),
      reviewStatus: cleanSheetWhitespace(record.reviewStatus || "needs_ocr"),
      sourceCompletionId: cleanSheetWhitespace(record.pigCompletionId),
      sourceRequestId: cleanSheetWhitespace(record.graphicsRequestId),
      completedAt: cleanSheetWhitespace(record.completedAt),
      qcApprovedAt: cleanSheetWhitespace(record.graphicsQcUpdatedAt),
      qcNote: String(record.graphicsQcNote || ""),
      productionNotes: String(record.notes || ""),
      sourceTool
    };
  }

  return {
    source: "weaver_qc_approved_graphic",
    imageType: "QI",
    sourceCompletionId: cleanSheetWhitespace(record.pigCompletionId),
    sourceRequestId: cleanSheetWhitespace(record.graphicsRequestId),
    title: String(record.poemTitle || ""),
    author: String(record.author || ""),
    book: String(record.bookTitle || ""),
    excerpt: String(record.quoteText || ""),
    driveLink: handoffAssetUrl,
    previewUrl: handoffPreviewUrl,
    completedAt: cleanSheetWhitespace(record.completedAt),
    qcApprovedAt: cleanSheetWhitespace(record.graphicsQcUpdatedAt),
    qcNote: String(record.graphicsQcNote || ""),
    productionNotes: String(record.notes || ""),
    sourceTool,
    bookLink: "",
    bookShortener,
    releaseCatalog
  };
}

async function getPoetryPleaseApprovedGraphics(bookTitle = "") {
  const requestedKey = normalizeBookKey(bookTitle);
  const rows = await readPigCompletedGraphicsRows();
  const canonicalBookAuthorMap = await getCanonicalGraphicsBookAuthorMap();
  const runtimeState = await getRuntimeGraphicsState(rows.map(row => row[PIG_COMPLETION_COLUMNS.completionId - 1]));
  const handoffState = await getRuntimeGraphicsHandoffState(
    rows.map(row => cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.requestId - 1]))
  );
  return rows
    .map((row, index) => overlayRuntimeHandoffState(
      overlayRuntimeGraphicsState(buildPigQcRecordFromSheetRow(row, index, canonicalBookAuthorMap), runtimeState),
      handoffState
    ))
    .filter(Boolean)
    .map(buildPoetryPleaseGraphicRecord)
    .filter(Boolean)
    .filter(record => !requestedKey || normalizeBookKey(record.book) === requestedKey);
}

async function getPoetryPleaseApprovedGraphicBooks() {
  const records = await getPoetryPleaseApprovedGraphics();
  return summarizeBooks(records.map(record => ({ bookTitle: record.book })));
}

async function getManualGraphicsReworkCandidates(bookTitle = "") {
  const requestedKey = normalizeBookKey(bookTitle);
  const rows = await readPigCompletedGraphicsRows();
  const canonicalBookAuthorMap = await getCanonicalGraphicsBookAuthorMap();
  const runtimeState = await getRuntimeGraphicsState(rows.map(row => row[PIG_COMPLETION_COLUMNS.completionId - 1]));
  const handoffState = await getRuntimeGraphicsHandoffState(
    rows.map(row => cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.requestId - 1]))
  );
  let existingManualReworkCompletionIds = new Set();
  try {
    const runtimeQueue = await syncWeaverRuntimeDb("get_handoff_queue", { limit: 5000 });
    existingManualReworkCompletionIds = new Set(
      (Array.isArray(runtimeQueue?.records) ? runtimeQueue.records : [])
        .filter(record => cleanSheetWhitespace(record?.sourceSystem) === "weaver_manual_rework")
        .map(record => cleanSheetWhitespace(
          record?.sourcePayload?.sourceCompletionId
          || record?.sourceCompletionId
        ))
        .filter(Boolean)
    );
  } catch {
    existingManualReworkCompletionIds = new Set();
  }
  return rows
    .map((row, index) => overlayRuntimeHandoffState(
      overlayRuntimeGraphicsState(buildPigQcRecordFromSheetRow(row, index, canonicalBookAuthorMap), runtimeState),
      handoffState
    ))
    .filter(Boolean)
    .filter(record => normalizeGraphicsQcDecision(record.graphicsQcDecision) === "APPROVE")
    .filter(record => cleanSheetWhitespace(record.sourceTool) === "P.I.G.")
    .filter(record => !existingManualReworkCompletionIds.has(cleanSheetWhitespace(record.pigCompletionId)))
    .filter(record => !requestedKey || normalizeBookKey(record.bookTitle) === requestedKey);
}

async function createManualGraphicsReworkRequests(records = [], note = "") {
  const selected = Array.isArray(records) ? records.filter(Boolean) : [];
  const trimmedNote = String(note || "").trim();
  if (!selected.length) {
    return { ok: false, error: "No approved graphics were selected." };
  }
  if (!trimmedNote) {
    return { ok: false, error: "A rework note is required." };
  }

  const now = new Date().toISOString();
  const requests = selected.map((record, index) => {
    const originalCompletionId = cleanSheetWhitespace(record.pigCompletionId || record.sourceCompletionId);
    const requestId = `manual-rework:${originalCompletionId}:${Date.now()}-${index + 1}`;
    const contentType = inferPoetryPleaseContentType(record, "");
    if (!contentType) {
      throw new Error(`Approved graphic ${cleanSheetWhitespace(record.graphicsRequestId) || "without an ID"} is missing contentType`);
    }
    return {
      graphicsRequestId: requestId,
      sourceSystem: "weaver_manual_rework",
      sourceStatus: "manual_rework",
      contentType,
      imageType: contentType,
      sourceCompletionId: originalCompletionId,
      revisionOf: originalCompletionId,
      originalGraphicsRequestId: cleanSheetWhitespace(record.graphicsRequestId),
      version: record.version || 2,
      pigProjectId: cleanSheetWhitespace(record.pigProjectId),
      editableProjectFileId: cleanSheetWhitespace(record.editableProjectFileId || record.projectFileId),
      editableProjectUrl: cleanSheetWhitespace(record.editableProjectUrl),
      reworkReason: "correct_and_recreate",
      metadataIssue: cleanSheetWhitespace(record.metadataIssue),
      aestheticIssue: cleanSheetWhitespace(record.aestheticIssue),
      qcNote: trimmedNote,
      requestedChanges: trimmedNote,
      handoffStatus: "requested",
      pigStatus: "not_started",
      qcStatus: "needs_revision",
      sourcePayload: {
        queueSheetRow: record.sourceSheetRow || record.sheetRow || "",
        bookTitle: record.bookTitle || "",
        poemTitle: record.poemTitle || "",
        author: record.author || "",
        quoteText: record.quoteText || "",
        contentType,
        imageType: contentType,
        notes: `Manual rework requested from approved graphic. ${trimmedNote}`,
        approved: "Y",
        created: "",
        workflowStatus: "Manual rework requested",
        recordId: requestId,
        graphicsRequestId: requestId,
        rejectReason: "correct_and_recreate",
        metadataIssue: "",
        aestheticIssue: "",
        qcNote: trimmedNote,
        source: "weaver_manual_rework",
        sourceRequestId: cleanSheetWhitespace(record.graphicsRequestId),
        sourceCompletionId: originalCompletionId,
        revisionOf: originalCompletionId,
        originalGraphicsRequestId: cleanSheetWhitespace(record.graphicsRequestId),
        version: record.version || 2,
        pigProjectId: cleanSheetWhitespace(record.pigProjectId),
        editableProjectFileId: cleanSheetWhitespace(record.editableProjectFileId || record.projectFileId),
        editableProjectUrl: cleanSheetWhitespace(record.editableProjectUrl),
        reworkReason: "correct_and_recreate",
        previousAssetUrl: cleanSheetWhitespace(record.assetLinkUrl || record.assetUrl),
        previousAssetPreviewUrl: cleanSheetWhitespace(record.assetPreviewUrl),
        requestedAt: now
      }
    };
  });

  const runtime = await syncWeaverRuntimeDb("upsert_handoff_requests", { requests });
  invalidateQueueSnapshots();
  return {
    ok: true,
    version: `${appVersion}-service-account`,
    createdCount: requests.length,
    requests,
    runtime
  };
}

async function getPoetryPleaseHandoffRecords(bookTitle = "") {
  const requestedKey = normalizeBookKey(bookTitle);
  const rows = await readPigCompletedGraphicsRows();
  const canonicalBookAuthorMap = await getCanonicalGraphicsBookAuthorMap();
  const runtimeState = await getRuntimeGraphicsState(rows.map(row => row[PIG_COMPLETION_COLUMNS.completionId - 1]));
  const handoffState = await getRuntimeGraphicsHandoffState(
    rows.map(row => cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.requestId - 1]))
  );
  return rows
    .map((row, index) => overlayRuntimeHandoffState(
      overlayRuntimeGraphicsState(buildPigQcRecordFromSheetRow(row, index, canonicalBookAuthorMap), runtimeState),
      handoffState
    ))
    .filter(Boolean)
    .filter(record => normalizeGraphicsQcDecision(record.graphicsQcDecision) === "APPROVE")
    .filter(record => !requestedKey || normalizeBookKey(record.bookTitle) === requestedKey);
}

async function getFailedPoetryPleaseHandoffRecords(bookTitle = "") {
  const records = await getPoetryPleaseHandoffRecords(bookTitle);
  return records.filter(record => cleanSheetWhitespace(record.poetryPleaseStatus).toUpperCase() === "FAILED");
}

async function handoffApprovedGraphicsToPoetryPlease(records = []) {
  const approvedRecords = Array.isArray(records) ? records.filter(Boolean) : [];
  if (!approvedRecords.length) {
    return { ok: true, skipped: true, reason: "no_records" };
  }
  if (!poetryPleaseApiKey) {
    return { ok: false, skipped: true, reason: "missing_poetry_please_api_key" };
  }

  const byContentType = new Map();
  approvedRecords.forEach(record => {
    const contentType = inferPoetryPleaseContentType(record, "");
    if (!contentType) {
      throw new Error(`Approved graphic ${cleanSheetWhitespace(record.graphicsRequestId) || "without an ID"} is missing contentType`);
    }
    if (!byContentType.has(contentType)) byContentType.set(contentType, []);
    byContentType.get(contentType).push(record);
  });

  const combinedResults = [];
  let createdCount = 0;
  let updatedCount = 0;
  let errorCount = 0;

  for (const [contentType, groupedRecords] of byContentType.entries()) {
    const requestBody = contentType === "FPI"
      ? {
          records: Array.from(new Map(groupedRecords.map(record => ({
            contentType: "FPI",
            sourceSystem: cleanSheetWhitespace(record.sourceSystem || "weaver"),
            sourceRecordId: cleanSheetWhitespace(record.sourceRecordId),
            bookShortener: cleanSheetWhitespace(record.bookShortener),
            author: String(record.author || ""),
            book: String(record.book || record.bookTitle || ""),
            title: String(record.title || record.poemTitle || ""),
            releaseCatalog: cleanSheetWhitespace(record.releaseCatalog),
            driveLink: cleanSheetWhitespace(record.driveLink),
            imageUrl: cleanSheetWhitespace(record.imageUrl || record.previewUrl),
            pageNumber: String(record.pageNumber || ""),
            ocrText: String(record.ocrText || ""),
            reviewStatus: cleanSheetWhitespace(record.reviewStatus || "needs_ocr")
          })).map(record => [record.sourceRecordId, record])).values())
        }
      : {
          imageType: "QI",
          payload: {
            records: groupedRecords
          }
        };
    console.log("[poetry-please/weaverImport] request", {
      contentType,
      recordCount: groupedRecords.length,
      payload: requestBody
    });

    const response = await fetch(`${poetryPleaseApiUrl.replace(/\/$/, "")}/internal/weaverImport`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": poetryPleaseApiKey
      },
      body: JSON.stringify(requestBody)
    });

    const result = await response.json().catch(() => ({}));
    console.log("[poetry-please/weaverImport] response", {
      contentType,
      status: response.status,
      ok: response.ok,
      result
    });
    if (!response.ok || !result.ok) {
      const message = result.error || `Poetry Please ${contentType} handoff failed with ${response.status}`;
      const error = new Error(message);
      error.responseStatus = response.status;
      error.responseBody = result;
      error.requestBody = requestBody;
      throw error;
    }

    const resultErrorCount = Number(result?.errorCount || 0);
    const failedResults = Array.isArray(result?.results)
      ? result.results.filter(item => item && item.ok === false)
      : [];
    if (resultErrorCount > 0 || failedResults.length > 0) {
      const message = failedResults
        .map(item => cleanSheetWhitespace(item?.error || item?.message))
        .filter(Boolean)
        .join(" | ")
        || result.error
        || `Poetry Please ${contentType} handoff reported ${Math.max(resultErrorCount, failedResults.length)} error${Math.max(resultErrorCount, failedResults.length) === 1 ? "" : "s"}`;
      const error = new Error(message);
      error.responseStatus = response.status;
      error.responseBody = result;
      error.requestBody = requestBody;
      throw error;
    }

    createdCount += Number(result?.createdCount || 0);
    updatedCount += Number(result?.updatedCount || 0);
    errorCount += resultErrorCount;
    combinedResults.push({ contentType, count: groupedRecords.length, request: requestBody, response: result });
  }

  return {
    ok: true,
    createdCount,
    updatedCount,
    errorCount,
    results: combinedResults
  };
}

async function retryFailedGraphicsHandoffs(records = []) {
  const failedRecords = Array.isArray(records) ? records.filter(Boolean) : [];
  if (!failedRecords.length) {
    return { ok: true, skipped: true, reason: "no_records", retriedCount: 0 };
  }

  const approvedRecords = failedRecords
    .map(buildPoetryPleaseGraphicRecord)
    .filter(Boolean);
  if (!approvedRecords.length) {
    return { ok: false, error: "No retryable approved graphics were selected." };
  }

  let poetryPlease;
  try {
    poetryPlease = await handoffApprovedGraphicsToPoetryPlease(approvedRecords);
  } catch (error) {
    poetryPlease = {
      ok: false,
      error: error.message,
      responseStatus: error.responseStatus || 0,
      responseBody: error.responseBody || null,
      requestBody: error.requestBody || null
    };
  }

  const handoffUpdatedAt = new Date().toISOString();
  const handoffStatus = poetryPlease.ok ? "HANDED_OFF" : "FAILED";
  const handoffNote = poetryPlease.ok
    ? `retry created=${Number(poetryPlease.createdCount || 0)} updated=${Number(poetryPlease.updatedCount || 0)} errors=${Number(poetryPlease.errorCount || 0)}`
    : `retry failed: ${String(poetryPlease.error || poetryPlease.reason || "handoff_failed")}`;

  const handoffRequests = failedRecords.flatMap(record => {
    const rowNumber = Number(record.sheetRow || 0);
    if (!rowNumber) return [];
    return [
      {
        range: `'${pigCompletedGraphicsSheetName.replace(/'/g, "''")}'!Q${rowNumber}`,
        values: [[handoffStatus]]
      },
      {
        range: `'${pigCompletedGraphicsSheetName.replace(/'/g, "''")}'!R${rowNumber}`,
        values: [[handoffUpdatedAt]]
      },
      {
        range: `'${pigCompletedGraphicsSheetName.replace(/'/g, "''")}'!S${rowNumber}`,
        values: [[handoffNote]]
      }
    ];
  });
  if (handoffRequests.length) {
    await batchUpdateSheetValuesServer(handoffRequests);
  }

  try {
    await syncWeaverRuntimeDb("insert_poetry_please_handoffs", {
      handoffs: approvedRecords.map(record => ({
        graphicsCompletionId: record.sourceCompletionId,
        handoffStatus,
        handedOffAt: handoffUpdatedAt,
        handoffMode: "retry",
        payload: {
          qcApprovedAt: record.qcApprovedAt,
          sourceRequestId: record.sourceRequestId,
          createdCount: Number(poetryPlease.createdCount || 0),
          updatedCount: Number(poetryPlease.updatedCount || 0),
          errorCount: Number(poetryPlease.errorCount || 0),
          results: poetryPlease.results || [],
          responseStatus: poetryPlease.responseStatus || 0,
          responseBody: poetryPlease.responseBody || null,
          requestBody: poetryPlease.requestBody || null,
          note: handoffNote
        }
      }))
    });
  } catch {
    // Keep the sheet as source of truth if DB handoff shadow write fails.
  }

  return {
    ok: true,
    retriedCount: approvedRecords.length,
    poetryPlease
  };
}

function buildPoetryPleaseExcerptRecord(record) {
  if (!record) return null;
  const excerpt = normalizeExcerptTransferText(record.excerpt || record.quoteText || "");
  const recordId = cleanSheetWhitespace(record.recordId);
  const bookShortener = cleanSheetWhitespace(record.bookShortener);
  const releaseCatalog = cleanSheetWhitespace(record.releaseCatalog);
  const contentType = normalizeExcerptContentType(record.contentType || "EXC");
  const intakeSourceRecordId = cleanSheetWhitespace(record.sourceRecordId);
  const canonicalSourceRecordId = contentType === "FP"
    ? canonicalFullPoemSourceRecordId(record)
    : intakeSourceRecordId;
  if (!recordId || !cleanSheetWhitespace(excerpt)) {
    return null;
  }
  if (!bookShortener || !releaseCatalog || (contentType === "FP" && !canonicalSourceRecordId)) {
    return {
      invalid: true,
      recordId,
      error: `Missing required publishing metadata: ${[
        !bookShortener ? "bookShortener" : "",
        !releaseCatalog ? "releaseCatalog" : "",
        contentType === "FP" && !canonicalSourceRecordId ? "poemTitle" : ""
      ].filter(Boolean).join(", ")}`
    };
  }

  return {
    contentType,
    recordId,
    sourceSystem: cleanSheetWhitespace(record.sourceSystem || "weaver"),
    sourceRecordId: canonicalSourceRecordId,
    sourceIntakeRecordId: intakeSourceRecordId,
    author: String(record.author || ""),
    bookTitle: String(record.bookTitle || ""),
    poemTitle: String(record.poemTitle || ""),
    excerpt,
    approvedAt: cleanSheetWhitespace(record.approvedAt),
    updatedAt: cleanSheetWhitespace(record.updatedAt),
    bookShortener,
    bookLink: String(record.bookLink || ""),
    releaseCatalog,
    socialMediaHandle: cleanSheetWhitespace(record.socialMediaHandle || record.instagramHandle || record.igHandle),
    instagramHandle: cleanSheetWhitespace(record.socialMediaHandle || record.instagramHandle || record.igHandle),
    igHandle: cleanSheetWhitespace(record.socialMediaHandle || record.instagramHandle || record.igHandle),
    driveLink: String(record.driveLink || ""),
    sourceUrl: String(record.sourceUrl || ""),
    pageNumber: String(record.pageNumber || "")
  };
}

async function handoffApprovedExcerptsToPoetryPlease(records = []) {
  const approvedRecords = Array.isArray(records) ? records.map(buildPoetryPleaseExcerptRecord).filter(Boolean) : [];
  if (!approvedRecords.length) {
    return { ok: true, skipped: true, reason: "no_records", results: [] };
  }
  if (!poetryPleaseApiKey) {
    return { ok: false, skipped: true, reason: "missing_poetry_please_api_key", results: [] };
  }

  const results = [];
  let createdCount = 0;
  let updatedCount = 0;
  let errorCount = 0;
  for (const record of approvedRecords) {
    if (record.invalid) {
      errorCount += 1;
      results.push({
        recordId: record.recordId,
        ok: false,
        error: record.error,
        response: { ok: false, error: record.error }
      });
      continue;
    }
    try {
      const response = await fetch(`${poetryPleaseApiUrl.replace(/\/$/, "")}/internal/weaverImport`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": poetryPleaseApiKey
        },
        body: JSON.stringify({
          contentType: record.contentType,
          payload: record
        })
      });

      const result = await response.json().catch(() => ({}));
      if (!response.ok || !result.ok) {
        errorCount += 1;
        results.push({
          recordId: record.recordId,
          ok: false,
          error: result.error || `Poetry Please EXC handoff failed with ${response.status}`,
          response: result
        });
        continue;
      }

      const itemId = cleanSheetWhitespace(
        result?.results?.[0]?.id ||
        result?.id ||
        result?.recordId
      );
      createdCount += Number(result?.createdCount || 0);
      updatedCount += Number(result?.updatedCount || 0);
      results.push({
        recordId: record.recordId,
        ok: true,
        poetryPleaseItemId: itemId,
        response: result
      });
    } catch (error) {
      errorCount += 1;
      results.push({
        recordId: record.recordId,
        ok: false,
        error: error.message,
        response: {}
      });
    }
  }

  return {
    ok: errorCount === 0,
    createdCount,
    updatedCount,
    errorCount,
    results
  };
}

function normalizeExcerptHandoffStatus(value) {
  return cleanSheetWhitespace(value).toLowerCase();
}

async function updateExcerptHandoffStatuses(records, poetryPlease) {
  const handoffRecords = Array.isArray(records) ? records : [];
  if (!handoffRecords.length) {
    return { ok: true, savedCount: 0, records: [] };
  }

  const now = new Date().toISOString();
  const poetryPleaseResults = new Map(
    Array.isArray(poetryPlease?.results)
      ? poetryPlease.results.map(result => [cleanSheetWhitespace(result.recordId), result])
      : []
  );
  const statusUpdates = handoffRecords.map(record => ({
    ...record,
    handoffStatus: poetryPleaseResults.get(cleanSheetWhitespace(record.recordId))?.ok === false ? "failed" : "sent",
    handedOffAt: poetryPleaseResults.get(cleanSheetWhitespace(record.recordId))?.ok === false
      ? String(record.handedOffAt || "")
      : now,
    poetryPleaseItemId: poetryPleaseResults.get(cleanSheetWhitespace(record.recordId))?.ok === false
      ? String(record.poetryPleaseItemId || "")
      : cleanSheetWhitespace(poetryPleaseResults.get(cleanSheetWhitespace(record.recordId))?.poetryPleaseItemId),
    errorMessage: poetryPleaseResults.get(cleanSheetWhitespace(record.recordId))?.ok === false
      ? String(
          poetryPleaseResults.get(cleanSheetWhitespace(record.recordId))?.error ||
          poetryPlease.error ||
          poetryPlease.reason ||
          "handoff_failed"
        )
      : "",
    updatedAt: now
  }));
  const sheetWrites = statusUpdates
    .flatMap(record => {
      const sourceRow = parseInt(record?.payload?.sourceRow || 0, 10);
      if (!sourceRow || sourceRow < SHEET_SOURCE_CONFIG.startRow) {
        return [];
      }
      const noteValue = cleanSheetWhitespace(record.poetryPleaseItemId)
        ? `item=${cleanSheetWhitespace(record.poetryPleaseItemId)}`
        : String(record.errorMessage || "");
      return [
        {
          range: `'${sourceSheetName.replace(/'/g, "''")}'!${toA1Column(SHEET_SOURCE_CONFIG.columnMap.excerptPoetryPleaseStatus)}${sourceRow}`,
          values: [[String(record.handoffStatus || "")]]
        },
        {
          range: `'${sourceSheetName.replace(/'/g, "''")}'!${toA1Column(SHEET_SOURCE_CONFIG.columnMap.excerptPoetryPleaseUpdatedAt)}${sourceRow}`,
          values: [[String(record.handedOffAt || record.updatedAt || "")]]
        },
        {
          range: `'${sourceSheetName.replace(/'/g, "''")}'!${toA1Column(SHEET_SOURCE_CONFIG.columnMap.excerptPoetryPleaseNote)}${sourceRow}`,
          values: [[noteValue]]
        }
      ];
    });
  if (sheetWrites.length) {
    await batchUpdateSheetValuesServer(sheetWrites);
    invalidateQueueSnapshots();
  }
  const statusResult = await syncWeaverRuntimeDb("upsert_excerpt_handoffs", { handoffs: statusUpdates });
  return {
    ok: true,
    savedCount: Array.isArray(statusResult?.records) ? statusResult.records.length : handoffRecords.length,
    records: Array.isArray(statusResult?.records) ? statusResult.records : handoffRecords
  };
}

async function getExcerptHandoffRecords({ recordIds = [], statuses = [] } = {}) {
  const exportResult = await getApprovedExcerptsExportFromSheets();
  const approvedRecords = Array.isArray(exportResult?.records) ? exportResult.records : [];
  const approvedHandoffs = approvedRecords
    .map(buildExcerptHandoffFromApprovedExportRecord)
    .filter(Boolean);
  const approvedById = new Map(
    approvedHandoffs.map(record => [cleanSheetWhitespace(record.recordId), record])
  );
  const result = await syncWeaverRuntimeDb("get_excerpt_handoffs", { recordIds: approvedHandoffs.map(record => record.recordId) });
  const runtimeRecords = Array.isArray(result?.records) ? result.records : [];
  const runtimeById = new Map(
    runtimeRecords.map(record => [cleanSheetWhitespace(record.recordId), record])
  );
  const records = approvedHandoffs.map(record => {
    const runtimeRecord = runtimeById.get(cleanSheetWhitespace(record.recordId));
    return runtimeRecord ? { ...record, ...runtimeRecord } : record;
  });
  const requestedIds = new Set(
    (Array.isArray(recordIds) ? recordIds : [])
      .map(value => cleanSheetWhitespace(value))
      .filter(Boolean)
  );
  const scopedRecords = requestedIds.size
    ? records.filter(record => requestedIds.has(cleanSheetWhitespace(record.recordId)))
    : records;
  const normalizedStatuses = (Array.isArray(statuses) ? statuses : [])
    .map(normalizeExcerptHandoffStatus)
    .filter(Boolean);
  const filtered = normalizedStatuses.length
    ? scopedRecords.filter(record => normalizedStatuses.includes(normalizeExcerptHandoffStatus(record.handoffStatus)))
    : scopedRecords;
  return {
    ok: true,
    count: filtered.length,
    records: filtered
  };
}

async function retryExcerptHandoffs({ recordIds = [], statuses = ["queued", "failed"] } = {}) {
  const handoffResult = await getExcerptHandoffRecords({ recordIds, statuses });
  const records = Array.isArray(handoffResult.records) ? handoffResult.records : [];
  if (!records.length) {
    return { ok: true, retriedCount: 0, skipped: true, reason: "no_records", records: [] };
  }

  let poetryPlease = { ok: true, skipped: true, reason: "no_records" };
  try {
    poetryPlease = await handoffApprovedExcerptsToPoetryPlease(records);
  } catch (error) {
    poetryPlease = { ok: false, error: error.message };
  }

  const statusResult = await updateExcerptHandoffStatuses(records, poetryPlease);
  return {
    ok: Boolean(statusResult.ok),
    retriedCount: records.length,
    records: statusResult.records,
    poetryPlease
  };
}

async function backfillApprovedExcerptHandoffs({ since = "", sourceRecordIds = [] } = {}) {
  const exportResult = await getApprovedExcerptsExportFromSheets({ since });
  const approvedRecords = Array.isArray(exportResult?.records) ? exportResult.records : [];
  const requestedIds = new Set(
    (Array.isArray(sourceRecordIds) ? sourceRecordIds : [])
      .map(value => cleanSheetWhitespace(value))
      .filter(Boolean)
  );
  const candidates = requestedIds.size
    ? approvedRecords.filter(record => requestedIds.has(cleanSheetWhitespace(record?.sourceRecordId)))
    : approvedRecords;
  const targetRecordIds = candidates
    .map(record => {
      const sourceRecordId = cleanSheetWhitespace(record?.sourceRecordId);
      if (!sourceRecordId) return "";
      const contentType = normalizeExcerptContentType(record?.contentType || "EXC");
      return `weaver-${contentType.toLowerCase()}-${sourceRecordId}`;
    })
    .filter(Boolean);
  const existingResult = await getExcerptHandoffRecords({ recordIds: targetRecordIds });
  const existingIds = new Set(
    (Array.isArray(existingResult?.records) ? existingResult.records : [])
      .map(record => cleanSheetWhitespace(record?.recordId))
      .filter(Boolean)
  );
  const missing = candidates.filter(
    record => {
      const sourceRecordId = cleanSheetWhitespace(record?.sourceRecordId);
      const contentType = normalizeExcerptContentType(record?.contentType || "EXC");
      return !existingIds.has(`weaver-${contentType.toLowerCase()}-${sourceRecordId}`);
    }
  );
  const handoffs = missing
    .map(buildExcerptHandoffFromApprovedExportRecord)
    .filter(Boolean);

  if (!handoffs.length) {
    return {
      ok: true,
      savedCount: 0,
      skipped: true,
      reason: "no_missing_records",
      records: []
    };
  }

  const upsertResult = await syncWeaverRuntimeDb("upsert_excerpt_handoffs", { handoffs });
  const records = Array.isArray(upsertResult?.records) ? upsertResult.records : handoffs;
  let poetryPlease = { ok: true, skipped: true, reason: "no_records" };
  try {
    poetryPlease = await handoffApprovedExcerptsToPoetryPlease(records);
  } catch (error) {
    poetryPlease = { ok: false, error: error.message };
  }
  const statusResult = await updateExcerptHandoffStatuses(records, poetryPlease);
  return {
    ok: Boolean(statusResult.ok),
    savedCount: records.length,
    records: statusResult.records,
    poetryPlease
  };
}

function hasResolvedGraphicsQcDecision(record) {
  return !!normalizeGraphicsQcDecision(record?.graphicsQcDecision);
}

function buildCleanupSheetGraphicsRecords(values = [], qcState = new Map(), canonicalBookAuthorMap = null) {
  const records = [];

  values.forEach((row, index) => {
    const currentBookTitle = cleanSheetWhitespace(row[2]);
    if (!currentBookTitle) return;

    const recordId = (row[8] || "").toString();
    const qc = qcState.get(String(index + 2)) || {};
    const notes = (row[4] || "").toString();
    const noteMeta = parseIntakeMetadataFromNotes(notes);
    records.push({
      sheetRow: index + 2,
      storageTarget: "cleanup_sheet",
      author: resolveGraphicsAuthor(row[0] || "", currentBookTitle, canonicalBookAuthorMap),
      poemTitle: (row[1] || "").toString(),
      bookTitle: currentBookTitle,
      quoteText: (row[3] || "").toString(),
      notes,
      socialMediaHandle: noteMeta.socialMediaHandle || "",
      instagramHandle: noteMeta.socialMediaHandle || "",
      igHandle: noteMeta.socialMediaHandle || "",
      approved: cleanSheetWhitespace(row[5]),
      created: cleanSheetWhitespace(row[6]),
      workflowStatus: cleanSheetWhitespace(row[7]),
      recordId,
      graphicsRequestId: buildWeaverGraphicsRequestId({
        recordId,
        author: resolveGraphicsAuthor(row[0] || "", currentBookTitle, canonicalBookAuthorMap),
        poemTitle: (row[1] || "").toString(),
        bookTitle: currentBookTitle,
        quoteText: (row[3] || "").toString()
      }),
      graphicsQcDecision: cleanSheetWhitespace(qc.decision),
      graphicsQcNote: qc.note || "",
      graphicsQcUpdatedAt: cleanSheetWhitespace(qc.updatedAt)
    });
  });

  return records;
}

async function getPendingGraphicsQcRecords({ includeCleanup = true } = {}) {
  return getCachedQueueSnapshot(`graphics-qc-pending:${includeCleanup ? "all" : "pig-only"}`, async () => {
    const qcReadBackend = cleanSheetWhitespace(process.env.WEAVER_GRAPHICS_QC_READ_BACKEND).toLowerCase();
    const cleanupReadBackend = cleanSheetWhitespace(process.env.WEAVER_GRAPHICS_CLEANUP_READ_BACKEND).toLowerCase();
    if (qcReadBackend === "firestore" && (!includeCleanup || cleanupReadBackend === "firestore")) {
      const result = await syncWeaverRuntimeDb("get_pending_graphics_qc", { includeCleanup });
      if (!result?.ok || !Array.isArray(result.records)) {
        throw new Error(result?.error || "Firestore Graphics QC queue read failed");
      }
      return result.records;
    }
    const range = `'${graphicsCleanupSheetName.replace(/'/g, "''")}'!A2:I`;
    const [pigRows, canonicalBookAuthorMap, cleanupState] = await Promise.all([
      readPigCompletedGraphicsRows(),
      getCanonicalGraphicsBookAuthorMap(),
      includeCleanup
        ? Promise.all([fetchSheetValuesServer(range), getGraphicsQcStateMapFromSheets()])
        : Promise.resolve([[], new Map()])
    ]);
    const completionIds = pigRows.map(row => row[PIG_COMPLETION_COLUMNS.completionId - 1]);
    const requestIds = pigRows.map(row => cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.requestId - 1]));
    const [runtimeState, handoffState] = await Promise.all([
      getRuntimeGraphicsState(completionIds),
      getRuntimeGraphicsHandoffState(requestIds)
    ]);
    const [values, qcState] = cleanupState;

    const cleanupRecords = includeCleanup
      ? buildCleanupSheetGraphicsRecords(values, qcState, canonicalBookAuthorMap)
      : [];
    const pendingCleanupRecords = cleanupRecords.filter(record => !hasResolvedGraphicsQcDecision(record));
    if (
      pendingCleanupRecords.length
      && cleanSheetWhitespace(process.env.WEAVER_GRAPHICS_CLEANUP_MIGRATE_ON_READ).toLowerCase() === "true"
    ) {
      const migration = await syncWeaverRuntimeDb("upsert_graphics_qc_queue_cards", {
        cards: pendingCleanupRecords
      });
      if (!migration?.ok || Number(migration.written || 0) !== pendingCleanupRecords.length) {
        throw new Error(migration?.error || "Cleanup Graphics QC Firestore migration was incomplete");
      }
    }
    const pigRecords = await hydrateGraphicsFolderAssets(pigRows
      .map((row, index) => overlayRuntimeHandoffState(
        overlayRuntimeGraphicsState(buildPigQcRecordFromSheetRow(row, index, canonicalBookAuthorMap), runtimeState),
        handoffState
      ))
      .filter(Boolean));

    return collapsePendingGraphicsQcRecords(
      mergeRecordCollections([cleanupRecords, pigRecords])
    ).filter(record => !hasResolvedGraphicsQcDecision(record));
  }, 30000);
}

function buildGraphicsReworkRequestId(completion) {
  const completionId = cleanSheetWhitespace(completion?.pigCompletionId);
  if (!completionId) return "";
  return `rework:${completionId}`;
}

function buildGraphicsReworkRequestRecord(completion) {
  const rejectReason = getGraphicsQcRejectReason(completion);
  if (normalizeGraphicsQcDecision(completion?.graphicsQcDecision) !== "REJECT" || rejectReason !== "correct_and_recreate") {
    return null;
  }

  const parsedNote = parseGraphicsQcStructuredNoteServer(completion.graphicsQcNote || "");
  const metadataIssue = completion.metadataIssue || parsedNote.metadataIssue;
  const aestheticIssue = completion.aestheticIssue || parsedNote.aestheticIssue;
  const reworkNotes = [
    "QC requested rework.",
    metadataIssue ? `Metadata issue: ${metadataIssue}` : "",
    aestheticIssue ? `Aesthetic issue: ${aestheticIssue}` : "",
    parsedNote.details ? `Details: ${parsedNote.details}` : ""
  ].filter(Boolean).join(" ");
  const originalGraphicsRequestId = cleanSheetWhitespace(completion.graphicsRequestId);
  const revisionOf = cleanSheetWhitespace(completion.pigCompletionId);
  const sourceRecordId = cleanSheetWhitespace(completion.sourceRecordId || completion.recordId);
  const contentType = inferPoetryPleaseContentType(completion, "");
  if (!contentType) {
    throw new Error(`Rejected completion ${revisionOf || "without an ID"} is missing contentType`);
  }
  const contentId = cleanSheetWhitespace(completion.contentId || completion.imageId || sourceRecordId || `pig:${revisionOf}`);
  const requestedChanges = parsedNote.details || completion.graphicsQcNote || reworkNotes;

  return {
    queueSheetRow: completion.sourceSheetRow || completion.sheetRow || "",
    sourceSheetRow: completion.sourceSheetRow || completion.sheetRow || "",
    bookTitle: completion.bookTitle,
    poemTitle: completion.poemTitle,
    author: completion.author,
    quoteText: completion.quoteText,
    text: completion.quoteText,
    notes: reworkNotes,
    approved: "Y",
    created: "",
    workflowStatus: "QC requested rework",
    recordId: `rework:${cleanSheetWhitespace(completion.pigCompletionId)}`,
    graphicsRequestId: buildGraphicsReworkRequestId(completion),
    originalGraphicsRequestId,
    revisionOf,
    version: 2,
    pigProjectId: cleanSheetWhitespace(completion.pigProjectId),
    editableProjectFileId: cleanSheetWhitespace(completion.editableProjectFileId || completion.projectFileId),
    editableProjectUrl: cleanSheetWhitespace(completion.editableProjectUrl),
    sourceRecordId,
    contentId,
    imageId: contentId,
    contentType,
    imageType: contentType,
    queueView: "rework",
    isActionable: true,
    statusLabel: "rework requested",
    nextAction: "rework",
    reworkReason: rejectReason,
    rejectReason,
    rejectedReason: rejectReason,
    metadataIssue,
    aestheticIssue,
    qcNote: completion.graphicsQcNote || "",
    requestedChanges,
    previousAssetUrl: cleanSheetWhitespace(completion.assetLinkUrl),
    previousAssetPreviewUrl: cleanSheetWhitespace(completion.assetPreviewUrl),
    assetUrl: cleanSheetWhitespace(completion.assetLinkUrl),
    assetPreviewUrl: cleanSheetWhitespace(completion.assetPreviewUrl),
    handoffStatus: "rework_requested",
    pigStatus: "needs_rework",
    qcStatus: "rejected",
    source: "weaver_qc_rework",
    sourceRequestId: cleanSheetWhitespace(completion.graphicsRequestId),
    sourceCompletionId: cleanSheetWhitespace(completion.pigCompletionId)
  };
}

function buildGraphicsMismatchRecord(completion) {
  const rejectReason = getGraphicsQcRejectReason(completion);
  if (normalizeGraphicsQcDecision(completion?.graphicsQcDecision) !== "REJECT" || rejectReason !== "mismatched_graphic") {
    return null;
  }

  const parsedNote = parseGraphicsQcStructuredNoteServer(completion.graphicsQcNote || "");

  return {
    ...completion,
    workflowStatus: "QC flagged mismatch",
    notes: parsedNote.details ? `Mismatch note: ${parsedNote.details}` : (completion.notes || ""),
    rejectReason,
    metadataIssue: completion.metadataIssue || parsedNote.metadataIssue || "",
    aestheticIssue: completion.aestheticIssue || parsedNote.aestheticIssue || "",
    source: "weaver_qc_mismatch",
    sourceRequestId: cleanSheetWhitespace(completion.graphicsRequestId),
    sourceCompletionId: cleanSheetWhitespace(completion.pigCompletionId)
  };
}

async function getPigReworkRequests(filterMode = "all") {
  const rows = await readPigCompletedGraphicsRows();
  const canonicalBookAuthorMap = await getCanonicalGraphicsBookAuthorMap();
  const runtimeState = await getRuntimeGraphicsState(rows.map(row => row[PIG_COMPLETION_COLUMNS.completionId - 1]));
  const handoffState = await getRuntimeGraphicsHandoffState(
    rows.map(row => cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.requestId - 1]))
  );
  let records = rows
    .map((row, index) => overlayRuntimeHandoffState(
      overlayRuntimeGraphicsState(buildPigQcRecordFromSheetRow(row, index, canonicalBookAuthorMap), runtimeState),
      handoffState
    ))
    .filter(Boolean)
    .map(buildGraphicsReworkRequestRecord)
    .filter(Boolean);

  if (cleanSheetWhitespace(filterMode).toLowerCase() === "current_titles") {
    const allowed = await getReviewQueueIncludeSet();
    records = records.filter(record => allowed.has(normalizeBookKey(record.bookTitle)));
  }

  return records;
}

async function getPigMismatchRecords(filterMode = "all") {
  return getCachedQueueSnapshot(`graphics-mismatch:${cleanSheetWhitespace(filterMode).toLowerCase() || "all"}`, async () => {
    const rows = await readPigCompletedGraphicsRows();
    const canonicalBookAuthorMap = await getCanonicalGraphicsBookAuthorMap();
    const runtimeState = await getRuntimeGraphicsState(rows.map(row => row[PIG_COMPLETION_COLUMNS.completionId - 1]));
    const handoffState = await getRuntimeGraphicsHandoffState(
      rows.map(row => cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.requestId - 1]))
    );
    let records = rows
      .map((row, index) => overlayRuntimeHandoffState(
        overlayRuntimeGraphicsState(buildPigQcRecordFromSheetRow(row, index, canonicalBookAuthorMap), runtimeState),
        handoffState
      ))
      .filter(Boolean)
      .map(buildGraphicsMismatchRecord)
      .filter(Boolean);

    if (cleanSheetWhitespace(filterMode).toLowerCase() === "current_titles") {
      const allowed = await getReviewQueueIncludeSet();
      records = records.filter(record => allowed.has(normalizeBookKey(record.bookTitle)));
    }

    return records;
  });
}

function buildCompletedRequestLookup(rows = []) {
  const requestIds = new Set();
  const sourceSheetRows = new Set();
  const sourceRecordIds = new Set();
  rows.forEach((row, index) => {
    const completion = buildPigQcRecordFromSheetRow(row, index);
    if (!completion) return;
    const requestId = cleanSheetWhitespace(completion.graphicsRequestId);
    if (requestId) {
      requestIds.add(requestId);
    }
    const pigRequestId = cleanSheetWhitespace(completion.pigRequestId);
    if (pigRequestId) {
      requestIds.add(pigRequestId);
    }
    if (completion.sourceSheetRow) {
      sourceSheetRows.add(String(completion.sourceSheetRow));
      requestIds.add(`weaver:row-${completion.sourceSheetRow}`);
    }
    const sourceRecordId = cleanSheetWhitespace(completion.sourceRecordId);
    if (sourceRecordId) {
      sourceRecordIds.add(sourceRecordId);
    }
    const sourceRowRequestId = buildWeaverGraphicsRequestId({
      recordId: sourceRecordId,
      author: completion.author,
      poemTitle: completion.poemTitle,
      bookTitle: completion.bookTitle,
      quoteText: completion.quoteText
    });
    if (sourceRowRequestId) {
      requestIds.add(sourceRowRequestId);
    }
  });
  return { requestIds, sourceSheetRows, sourceRecordIds };
}

function dedupeGraphicsQueueRecords(records = []) {
  const byKey = new Map();

  records.forEach(record => {
    const key = [
      normalizeBookKey(record?.bookTitle),
      cleanSheetWhitespace(record?.author).toLowerCase(),
      cleanSheetWhitespace(record?.poemTitle).toLowerCase(),
      cleanSheetWhitespace(record?.quoteText)
    ].join("::");
    if (!key.replace(/:+/g, "")) return;
    const existing = byKey.get(key);
    if (!existing) {
      byKey.set(key, record);
      return;
    }
    const existingRow = Number(existing.queueSheetRow || 0) || Number.MAX_SAFE_INTEGER;
    const nextRow = Number(record.queueSheetRow || 0) || Number.MAX_SAFE_INTEGER;
    if (nextRow < existingRow) {
      byKey.set(key, record);
    }
  });

  return Array.from(byKey.values());
}

async function getPigGraphicsRequests(filterMode = "all", { includeCompleted = false } = {}) {
  const cacheKey = `graphics-requests:${cleanSheetWhitespace(filterMode).toLowerCase() || "all"}:${includeCompleted ? "with-completed" : "open-only"}`;
  return getCachedQueueSnapshot(cacheKey, async () => {
    const range = `'${graphicsQueueSheetName.replace(/'/g, "''")}'!A2:I`;
    const [rows, completionRows, reworkRequests, canonicalBookAuthorMap] = await Promise.all([
      fetchSheetValuesServer(range),
      readPigCompletedGraphicsRows(),
      getPigReworkRequests(filterMode),
      getCanonicalGraphicsBookAuthorMap()
    ]);
    const completedLookup = buildCompletedRequestLookup(completionRows);
    let records = rows
      .map((row, index) => buildGraphicsRequestRecordFromQueueRow(row, index, canonicalBookAuthorMap))
      .filter(Boolean)
      .filter(record => {
        if (includeCompleted) return true;
        const requestId = cleanSheetWhitespace(record.graphicsRequestId);
        const sourceRow = record.queueSheetRow ? String(record.queueSheetRow) : "";
        const recordId = cleanSheetWhitespace(record.recordId);
        return requestId
          && !completedLookup.requestIds.has(requestId)
          && (!sourceRow || !completedLookup.sourceSheetRows.has(sourceRow))
          && (!recordId || !completedLookup.sourceRecordIds.has(recordId));
      });

    const queueLedgerRequests = records
      .map(buildGraphicsHandoffLedgerRequest)
      .filter(request => cleanSheetWhitespace(request.graphicsRequestId));
    if (queueLedgerRequests.length) {
      await syncWeaverRuntimeDb("upsert_handoff_requests", { requests: queueLedgerRequests }).catch(() => null);
      const ledgerByRequestId = await getRuntimeGraphicsHandoffState(
        queueLedgerRequests.map(request => request.graphicsRequestId)
      );
      records = records.filter(record => {
        const ledger = ledgerByRequestId[cleanSheetWhitespace(record.graphicsRequestId)];
        if (!ledger) return true;
        const handoffStatus = cleanSheetWhitespace(ledger.handoffStatus).toLowerCase();
        const pigStatus = cleanSheetWhitespace(ledger.pigStatus).toLowerCase();
        const qcStatus = cleanSheetWhitespace(ledger.qcStatus).toLowerCase();
        return (
          (handoffStatus === "requested" || handoffStatus === "claimed" || handoffStatus === "rejected")
          && !["generated", "exported", "uploaded", "failed"].includes(pigStatus)
          && (qcStatus === "not_sent" || qcStatus === "needs_revision")
        );
      });
    }

    if (cleanSheetWhitespace(filterMode).toLowerCase() === "current_titles") {
      const allowed = await getReviewQueueIncludeSet();
      records = records.filter(record => allowed.has(normalizeBookKey(record.bookTitle)));
    }

    return [...dedupeGraphicsQueueRecords(records), ...reworkRequests];
  });
}

async function getPigGraphicsRequestBooks(filterMode = "all") {
  const records = await getPigGraphicsRequests(filterMode);
  return summarizeBooks(records);
}

function summarizeGraphicsHandoffBooks(records = []) {
  const byKey = new Map();

  records.forEach(record => {
    const bookTitle = cleanSheetWhitespace(record.bookTitle);
    const bookKey = normalizeBookKey(bookTitle);
    if (!bookKey) return;
    if (!byKey.has(bookKey)) {
      byKey.set(bookKey, {
        bookTitle,
        bookKey,
        count: 0,
        actionableCount: 0,
        reworkCount: 0
      });
    }

    const summary = byKey.get(bookKey);
    summary.bookTitle = choosePreferredBookTitle(summary.bookTitle, bookTitle);
    summary.count += 1;
    summary.actionableCount += 1;
    if (
      cleanSheetWhitespace(record.queueView).toLowerCase() === "rework"
      || cleanSheetWhitespace(record.sourceSystem || record.source).toLowerCase() === "weaver_qc_rework"
    ) {
      summary.reworkCount += 1;
    }
  });

  return Array.from(byKey.values())
    .map(summary => {
      const nextAction = summary.reworkCount > 0 && summary.reworkCount === summary.actionableCount
        ? "rework"
        : "generate";
      const statusLabel = nextAction === "rework"
        ? `${summary.actionableCount} rework`
        : `${summary.actionableCount} open`;
      return {
        ...summary,
        nextAction,
        statusLabel
      };
    })
    .sort((left, right) => left.bookTitle.localeCompare(right.bookTitle));
}

const COVERAGE_NEEDS_DEFAULT_TARGET_COUNT = 10;
const COVERAGE_NEEDS_CONTENT_TYPES = ["INT", "FPI"];
const COVERAGE_NEEDS_CACHE_TTL_MS = 60_000;
let poetryPleaseCoverageCache = null;
const COVERAGE_NEEDS_EXCLUDED_BOOK_KEYS = new Set([
  "short form contest may 2026",
  "smoke test ledger 1778781334"
]);

function isExcludedCoverageNeedsBook(bookTitle = "") {
  const key = normalizeBookKey(bookTitle);
  return Boolean(key && (
    COVERAGE_NEEDS_EXCLUDED_BOOK_KEYS.has(key)
      || key.startsWith("smoke test ledger ")
  ));
}

function parseOptionalRating(value) {
  if (value === null || value === undefined || value === "") return null;
  const rating = Number(value);
  return Number.isFinite(rating) ? rating : null;
}

function getCoveragePriority(record = {}, bookSummary = {}) {
  const excerptRating = parseOptionalRating(record.excerptRating);
  const poemRating = parseOptionalRating(record.poemRating);
  if (isCoverageReworkRecord(record)) {
    return {
      priorityTier: 0,
      priorityScore: 100000 + Number(bookSummary.remainingActionableNeeded || 0)
    };
  }
  if (excerptRating !== null && poemRating !== null) {
    return {
      priorityTier: 1,
      priorityScore: excerptRating + poemRating
    };
  }
  if (excerptRating !== null) {
    return {
      priorityTier: 2,
      priorityScore: excerptRating
    };
  }
  return {
    priorityTier: 3,
    priorityScore: 0
  };
}

function isCoverageReworkRecord(record = {}) {
  return cleanSheetWhitespace(record.queueView).toLowerCase() === "rework"
    || cleanSheetWhitespace(record.nextAction).toLowerCase() === "rework"
    || cleanSheetWhitespace(record.source || record.sourceSystem).toLowerCase() === "weaver_qc_rework";
}

function getCoverageRequestId(record = {}) {
  return cleanSheetWhitespace(record.graphicsRequestId)
    || buildWeaverGraphicsRequestId({
      recordId: record.recordId,
      author: record.author,
      poemTitle: record.poemTitle,
      bookTitle: record.bookTitle,
      quoteText: record.quoteText || record.text
    });
}

function buildCoverageQueueRecord(record = {}, bookSummary = {}) {
  const priority = getCoveragePriority(record, bookSummary);
  const isRework = isCoverageReworkRecord(record);
  return {
    queueView: "coverage_needs",
    isActionable: true,
    nextAction: isRework ? "rework" : "generate",
    statusLabel: isRework ? "Needs rework" : "Needs coverage",
    bookTitle: cleanSheetWhitespace(record.bookTitle),
    graphicsRequestId: getCoverageRequestId(record),
    excerptId: cleanSheetWhitespace(record.recordId || record.sourceRecordId),
    poemId: cleanSheetWhitespace(record.poemId),
    quoteText: String(record.quoteText || record.text || ""),
    text: String(record.quoteText || record.text || ""),
    author: String(record.author || ""),
    poemTitle: String(record.poemTitle || ""),
    bookKey: bookSummary.bookKey || normalizeBookKey(record.bookTitle),
    approvedCount: bookSummary.approvedCount || 0,
    intCount: bookSummary.intCount ?? null,
    fpiCount: bookSummary.fpiCount ?? null,
    combinedCount: bookSummary.combinedCount ?? null,
    coverageCountSource: bookSummary.coverageCountSource || "weaver_fallback",
    coverageCountsAuthoritative: Boolean(bookSummary.coverageCountsAuthoritative),
    coverageFallbackReason: bookSummary.coverageFallbackReason || "",
    weaverCompletedQiCount: bookSummary.weaverCompletedQiCount || 0,
    pendingQcCount: bookSummary.pendingQcCount || 0,
    inProgressCount: bookSummary.inProgressCount || 0,
    reworkCount: bookSummary.reworkCount || 0,
    target: bookSummary.targetCount || COVERAGE_NEEDS_DEFAULT_TARGET_COUNT,
    targetCount: bookSummary.targetCount || COVERAGE_NEEDS_DEFAULT_TARGET_COUNT,
    remaining: bookSummary.remainingApprovedNeeded || 0,
    complete: Boolean(bookSummary.complete),
    remainingActionableNeeded: bookSummary.remainingActionableNeeded || 0,
    remainingGenerationNeeded: bookSummary.remainingGenerationNeeded || 0,
    actionableReworkCount: bookSummary.actionableReworkCount || 0,
    acceptableContentTypes: [...COVERAGE_NEEDS_CONTENT_TYPES],
    coverageLabel: "INT / FPI coverage",
    queueLabel: "INT / FPI coverage",
    priorityTier: priority.priorityTier,
    priorityScore: priority.priorityScore,
    excerptRating: parseOptionalRating(record.excerptRating),
    poemRating: parseOptionalRating(record.poemRating),
    coverageStatus: isRework ? "rework" : "candidate",
    countsTowardCoverage: false,
    reworkReason: cleanSheetWhitespace(record.reworkReason || record.rejectReason),
    revisionOf: cleanSheetWhitespace(record.revisionOf),
    previousAssetUrl: cleanSheetWhitespace(record.previousAssetUrl || record.assetUrl),
    previousAssetPreviewUrl: cleanSheetWhitespace(record.previousAssetPreviewUrl || record.assetPreviewUrl),
    queueSheetRow: record.queueSheetRow || record.sourceSheetRow || "",
    sourceSheetRow: record.sourceSheetRow || record.queueSheetRow || "",
    recordId: cleanSheetWhitespace(record.recordId),
    source: cleanSheetWhitespace(record.source || "weaver_graphics_queue")
  };
}

function isApprovedCoverageGraphic(record = {}) {
  return normalizeGraphicsQcDecision(record.graphicsQcDecision) === "APPROVE";
}

function isPendingCoverageQcGraphic(record = {}) {
  return !normalizeGraphicsQcDecision(record.graphicsQcDecision);
}

function isCoverageInProgressLedgerState(record = {}) {
  const pigStatus = cleanSheetWhitespace(record.pigStatus || record.ledgerPigStatus).toLowerCase();
  const handoffStatus = cleanSheetWhitespace(record.handoffStatus || record.ledgerHandoffStatus).toLowerCase();
  return ["claimed", "generated", "exported", "uploaded"].includes(pigStatus)
    || ["claimed"].includes(handoffStatus);
}

function normalizeCoverageBookKey(value) {
  return cleanSheetWhitespace(value).toLowerCase();
}

function getCoverageRecordBookKey(record = {}) {
  return normalizeCoverageBookKey(record.bookKey) || normalizeBookKey(record.bookTitle);
}

async function fetchPoetryPleaseIntFpiCoverageCounts() {
  if (poetryPleaseCoverageCache?.expiresAt > Date.now()) {
    return {
      ...poetryPleaseCoverageCache.value,
      source: "poetry_please_cache",
      cached: true
    };
  }
  if (!poetryPleaseApiKey) {
    return {
      countsByBookKey: new Map(),
      ok: false,
      lane: "INT_FPI",
      defaultTarget: COVERAGE_NEEDS_DEFAULT_TARGET_COUNT,
      error: "missing_poetry_please_api_key"
    };
  }
  try {
    const url = new URL(`${poetryPleaseApiUrl.replace(/\/$/, "")}/internal/coverageCounts`);
    url.searchParams.set("lane", "INT_FPI");
    const response = await fetch(url, {
      headers: {
        Accept: "application/json",
        "x-api-key": poetryPleaseApiKey
      },
      signal: AbortSignal.timeout(12000)
    });
    const result = await response.json().catch(() => ({}));
    if (!response.ok || !result.ok) {
      throw new Error(result.error || `coverageCounts failed with ${response.status}`);
    }
    const countsByBookKey = new Map();
    (Array.isArray(result.counts) ? result.counts : []).forEach(entry => {
      const bookKey = normalizeCoverageBookKey(entry.bookKey);
      if (!bookKey) return;
      const intCount = Math.max(0, Number(entry.intCount || 0));
      const fpiCount = Math.max(0, Number(entry.fpiCount || 0));
      const combinedCount = Math.max(0, Number(entry.combinedCount ?? (intCount + fpiCount)));
      const target = Math.max(0, Number(entry.target ?? result.target ?? COVERAGE_NEEDS_DEFAULT_TARGET_COUNT));
      countsByBookKey.set(bookKey, {
        bookKey: cleanSheetWhitespace(entry.bookKey),
        bookTitle: cleanSheetWhitespace(entry.bookTitle),
        intCount,
        fpiCount,
        combinedCount,
        target,
        remaining: Math.max(0, Number(entry.remaining ?? (target - combinedCount))),
        complete: entry.complete === true || combinedCount >= target
      });
    });
    const coverage = {
      countsByBookKey,
      ok: true,
      source: "poetry_please",
      lane: "INT_FPI",
      defaultTarget: Math.max(0, Number(result.target ?? COVERAGE_NEEDS_DEFAULT_TARGET_COUNT)),
      entryCount: countsByBookKey.size,
      snapshotMeta: result.snapshotMeta || null
    };
    poetryPleaseCoverageCache = {
      expiresAt: Date.now() + COVERAGE_NEEDS_CACHE_TTL_MS,
      value: coverage
    };
    return coverage;
  } catch (error) {
    console.warn("[coverage_needs] Poetry Please INT/FPI count fallback:", error.message);
    return {
      countsByBookKey: new Map(),
      ok: false,
      lane: "INT_FPI",
      defaultTarget: COVERAGE_NEEDS_DEFAULT_TARGET_COUNT,
      error: error.message
    };
  }
}

function buildCoverageBookSummaries({ openRecords = [], pendingQcRecords = [], completedRecords = [], reworkRecords = [], ledgerByRequestId = {}, poetryPleaseCoverage = {} } = {}) {
  const byKey = new Map();
  const ensure = record => {
    const title = cleanSheetWhitespace(record?.bookTitle || record);
    const key = typeof record === "string" ? normalizeBookKey(title) : getCoverageRecordBookKey(record);
    if (!key) return null;
    if (isExcludedCoverageNeedsBook(title)) return null;
    if (!byKey.has(key)) {
      byKey.set(key, {
        bookTitle: title,
        bookKey: key,
        targetCount: poetryPleaseCoverage.defaultTarget || COVERAGE_NEEDS_DEFAULT_TARGET_COUNT,
        approvedCount: 0,
        weaverCompletedQiCount: 0,
        pendingQcCount: 0,
        inProgressCount: 0,
        reworkCount: 0,
        candidateCount: 0
      });
    }
    const summary = byKey.get(key);
    summary.bookTitle = choosePreferredBookTitle(summary.bookTitle, title);
    return summary;
  };

  const completedRequestIds = new Set();
  completedRecords.forEach(record => {
    if (!isApprovedCoverageGraphic(record)) return;
    const requestId = getCoverageRequestId(record);
    if (completedRequestIds.has(requestId)) return;
    completedRequestIds.add(requestId);
    const summary = ensure(record);
    if (summary) summary.weaverCompletedQiCount += 1;
  });

  const activeByRequestId = new Map();
  const setActiveState = (record, state, priority) => {
    const requestId = getCoverageRequestId(record);
    const existing = activeByRequestId.get(requestId);
    if (!existing || priority > existing.priority) {
      activeByRequestId.set(requestId, { record, state, priority });
    }
  };
  openRecords.forEach(record => {
    const ledger = ledgerByRequestId[getCoverageRequestId(record)] || {};
    setActiveState(record, isCoverageInProgressLedgerState(ledger) ? "in_progress" : "candidate", 1);
  });
  pendingQcRecords.forEach(record => {
    if (isPendingCoverageQcGraphic(record)) setActiveState(record, "pending_qc", 2);
  });
  reworkRecords.forEach(record => setActiveState(record, "rework", 3));
  activeByRequestId.forEach(({ record, state }) => {
    const summary = ensure(record);
    if (!summary) return;
    if (state === "pending_qc") summary.pendingQcCount += 1;
    else if (state === "in_progress") summary.inProgressCount += 1;
    else if (state === "rework") summary.reworkCount += 1;
    else summary.candidateCount += 1;
  });

  return Array.from(byKey.values()).map(summary => {
    const coverageEntry = poetryPleaseCoverage.countsByBookKey?.get(summary.bookKey);
    const endpointSucceeded = Boolean(poetryPleaseCoverage.ok);
    const hasCanonicalMatch = Boolean(coverageEntry);
    const coverageCountsAuthoritative = endpointSucceeded && hasCanonicalMatch;
    summary.targetCount = coverageEntry?.target || poetryPleaseCoverage.defaultTarget || COVERAGE_NEEDS_DEFAULT_TARGET_COUNT;
    summary.intCount = coverageCountsAuthoritative ? coverageEntry.intCount : null;
    summary.fpiCount = coverageCountsAuthoritative ? coverageEntry.fpiCount : null;
    summary.combinedCount = coverageCountsAuthoritative ? coverageEntry.combinedCount : null;
    summary.approvedCount = coverageCountsAuthoritative
      ? summary.combinedCount
      : summary.weaverCompletedQiCount;
    const remainingApprovedNeeded = coverageCountsAuthoritative
      ? coverageEntry.remaining
      : Math.max(0, summary.targetCount - summary.approvedCount);
    const remainingActionableNeeded = coverageCountsAuthoritative
      ? Math.max(
        0,
        remainingApprovedNeeded - summary.pendingQcCount - summary.inProgressCount
      )
      : 0;
    const actionableReworkCount = Math.min(summary.reworkCount, remainingActionableNeeded);
    const remainingGenerationNeeded = Math.max(0, remainingActionableNeeded - actionableReworkCount);
    const nextAction = actionableReworkCount > 0
      ? "rework"
      : (remainingGenerationNeeded > 0 ? "generate" : "hold");
    return {
      bookTitle: summary.bookTitle,
      bookKey: summary.bookKey,
      target: summary.targetCount,
      targetCount: summary.targetCount,
      approvedCount: summary.approvedCount,
      intCount: summary.intCount,
      fpiCount: summary.fpiCount,
      combinedCount: summary.combinedCount,
      complete: coverageCountsAuthoritative ? Boolean(coverageEntry.complete) : remainingApprovedNeeded === 0,
      acceptableContentTypes: [...COVERAGE_NEEDS_CONTENT_TYPES],
      coverageLabel: "INT / FPI coverage",
      coverageCountSource: coverageCountsAuthoritative ? "poetry_please" : "weaver_completed_int_only_fallback",
      coverageCountsAuthoritative,
      coverageFallbackReason: coverageCountsAuthoritative
        ? ""
        : (endpointSucceeded ? "canonical_book_key_not_found" : (poetryPleaseCoverage.error || "coverage_endpoint_failed")),
      weaverCompletedQiCount: summary.weaverCompletedQiCount,
      pendingQcCount: summary.pendingQcCount,
      inProgressCount: summary.inProgressCount,
      reworkCount: summary.reworkCount,
      actionableReworkCount,
      remaining: remainingApprovedNeeded,
      remainingApprovedNeeded,
      remainingActionableNeeded,
      remainingGenerationNeeded,
      statusLabel: `${summary.approvedCount} approved + ${summary.pendingQcCount} pending / ${summary.targetCount}`,
      nextAction
    };
  }).filter(summary => summary.remainingApprovedNeeded > 0);
}

async function getAllRuntimeHandoffQueueRecords(filter = "all") {
  const records = [];
  let cursor = 0;
  for (let page = 0; page < 10; page += 1) {
    const result = await syncWeaverRuntimeDb("get_handoff_queue", {
      filter,
      limit: 500,
      cursor
    });
    const pageRecords = Array.isArray(result?.records) ? result.records : [];
    records.push(...pageRecords);
    const nextCursor = Number(result?.nextCursor);
    if (!Number.isFinite(nextCursor) || nextCursor <= cursor || pageRecords.length === 0) break;
    cursor = nextCursor;
  }
  return records;
}

async function getCoverageNeedsView(filterMode = "coverage_needs") {
  const [handoffRecords, pendingQcResult, poetryPleaseCoverage] = await Promise.all([
    getAllRuntimeHandoffQueueRecords("all"),
    syncWeaverRuntimeDb("get_pending_graphics_qc", { includeCleanup: false }),
    fetchPoetryPleaseIntFpiCoverageCounts()
  ]);
  const pendingQcRecords = Array.isArray(pendingQcResult?.records) ? pendingQcResult.records : [];
  const reworkRecords = handoffRecords.filter(isCoverageReworkRecord);
  const openRecords = handoffRecords.filter(record => !isCoverageReworkRecord(record));
  const ledgerByRequestId = Object.fromEntries(
    handoffRecords.map(record => [getCoverageRequestId(record), record])
  );
  const books = buildCoverageBookSummaries({
    openRecords,
    pendingQcRecords,
    completedRecords: [],
    reworkRecords,
    ledgerByRequestId,
    poetryPleaseCoverage
  }).sort((left, right) => {
    if (left.nextAction === "rework" && right.nextAction !== "rework") return -1;
    if (left.nextAction !== "rework" && right.nextAction === "rework") return 1;
    return right.remainingActionableNeeded - left.remainingActionableNeeded
      || left.bookTitle.localeCompare(right.bookTitle);
  });
  const bookByKey = Object.fromEntries(books.map(book => [book.bookKey, book]));
  const freshCandidates = openRecords
    .filter(record => {
      const summary = bookByKey[getCoverageRecordBookKey(record)];
      if (!summary || summary.remainingGenerationNeeded <= 0) return false;
      const ledger = ledgerByRequestId[getCoverageRequestId(record)] || {};
      return !isCoverageInProgressLedgerState(ledger);
    })
    .map(record => buildCoverageQueueRecord(record, bookByKey[getCoverageRecordBookKey(record)]))
    .sort((left, right) => left.priorityTier - right.priorityTier
      || right.priorityScore - left.priorityScore
      || Number(left.queueSheetRow || left.sourceSheetRow || 0) - Number(right.queueSheetRow || right.sourceSheetRow || 0));
  const freshSlotsByBookKey = new Map(books.map(book => [book.bookKey, book.remainingGenerationNeeded]));
  const freshRecords = freshCandidates.filter(record => {
    const slots = Number(freshSlotsByBookKey.get(record.bookKey) || 0);
    if (slots <= 0) return false;
    freshSlotsByBookKey.set(record.bookKey, slots - 1);
    return true;
  });
  const reworkSlotsByBookKey = new Map(books.map(book => [book.bookKey, book.actionableReworkCount]));
  const reworkQueueRecords = reworkRecords
    .filter(record => bookByKey[getCoverageRecordBookKey(record)])
    .map(record => buildCoverageQueueRecord(record, bookByKey[getCoverageRecordBookKey(record)]))
    .sort((left, right) => left.priorityTier - right.priorityTier
      || right.priorityScore - left.priorityScore
      || Number(left.queueSheetRow || left.sourceSheetRow || 0) - Number(right.queueSheetRow || right.sourceSheetRow || 0))
    .filter(record => {
      const slots = Number(reworkSlotsByBookKey.get(record.bookKey) || 0);
      if (slots <= 0) return false;
      reworkSlotsByBookKey.set(record.bookKey, slots - 1);
      return true;
    });
  const records = [...reworkQueueRecords, ...freshRecords]
    .sort((left, right) => left.priorityTier - right.priorityTier
      || right.priorityScore - left.priorityScore
      || right.remainingActionableNeeded - left.remainingActionableNeeded
      || Number(left.queueSheetRow || left.sourceSheetRow || 0) - Number(right.queueSheetRow || right.sourceSheetRow || 0)
      || cleanSheetWhitespace(left.graphicsRequestId).localeCompare(cleanSheetWhitespace(right.graphicsRequestId)));

  return {
    ok: true,
    filter: "coverage_needs",
    poetryPleaseCoverage: {
      ok: Boolean(poetryPleaseCoverage.ok),
      source: poetryPleaseCoverage.source || "weaver_fallback",
      lane: "INT_FPI",
      label: "INT / FPI coverage",
      target: poetryPleaseCoverage.defaultTarget || COVERAGE_NEEDS_DEFAULT_TARGET_COUNT,
      entryCount: poetryPleaseCoverage.entryCount || 0,
      cached: Boolean(poetryPleaseCoverage.cached),
      fallbackActive: !poetryPleaseCoverage.ok || books.some(book => !book.coverageCountsAuthoritative),
      fallbackBookCount: books.filter(book => !book.coverageCountsAuthoritative).length,
      error: poetryPleaseCoverage.error || "",
      snapshotMeta: poetryPleaseCoverage.snapshotMeta || null
    },
    books,
    records
  };
}

function summarizePigCompletionsByRequest(rows = []) {
  const byRequestId = new Map();

  rows.forEach((row, index) => {
    const completion = buildPigQcRecordFromSheetRow(row, index);
    if (!completion) return;
    const requestId = cleanSheetWhitespace(completion.graphicsRequestId);
    if (!requestId) return;

    if (!byRequestId.has(requestId)) {
      byRequestId.set(requestId, {
        count: 0,
        latestCompletedAt: "",
        latestAssetUrl: "",
        latestAssetPreviewUrl: "",
        latestCompletionId: ""
      });
    }

    const summary = byRequestId.get(requestId);
    summary.count += 1;
    const completedAt = cleanSheetWhitespace(completion.completedAt);
    if (!summary.latestCompletedAt || (completedAt && completedAt > summary.latestCompletedAt)) {
      summary.latestCompletedAt = completedAt;
      summary.latestAssetUrl = cleanSheetWhitespace(completion.assetLinkUrl);
      summary.latestAssetPreviewUrl = cleanSheetWhitespace(completion.assetPreviewUrl);
      summary.latestCompletionId = cleanSheetWhitespace(completion.pigCompletionId);
    }
  });

  return byRequestId;
}

function buildPigGraphicsRequestSummary(group) {
  const excerpts = group.records.map(record => ({
    queueSheetRow: record.queueSheetRow,
    recordId: record.recordId,
    poemTitle: record.poemTitle,
    author: record.author,
    quoteText: record.quoteText,
    notes: record.notes,
    rejectReason: record.rejectReason || "",
    metadataIssue: record.metadataIssue || "",
    aestheticIssue: record.aestheticIssue || "",
    qcNote: record.qcNote || ""
  }));

  return {
    graphicsRequestId: group.graphicsRequestId,
    bookTitle: group.bookTitle,
    poemTitle: group.poemTitle,
    author: group.author,
    excerptCount: group.records.length,
    queueRowCount: group.records.length,
    queueSheetRows: group.records.map(record => record.queueSheetRow).filter(Boolean),
    requestStatus: group.requestStatus,
    completionCount: group.completionCount,
    latestCompletedAt: group.latestCompletedAt,
    assetUrl: group.assetUrl,
    assetPreviewUrl: group.assetPreviewUrl,
    latestCompletionId: group.latestCompletionId,
    rejectReason: group.rejectReason || "",
    metadataIssue: group.metadataIssue || "",
    aestheticIssue: group.aestheticIssue || "",
    qcNote: group.qcNote || "",
    sourceRequestId: group.sourceRequestId || "",
    sourceCompletionId: group.sourceCompletionId || "",
    source: "weaver_graphics_queue",
    excerpts
  };
}

async function getPigGraphicsRequestSummaries(filterMode = "all", bookTitle = "") {
  const [records, completionRows] = await Promise.all([
    getPigGraphicsRequests(filterMode),
    readPigCompletedGraphicsRows()
  ]);
  const completionSummaryByRequestId = summarizePigCompletionsByRequest(completionRows);
  const requestedKey = normalizeBookKey(bookTitle);
  const grouped = new Map();

  records.forEach(record => {
    if (requestedKey && normalizeBookKey(record.bookTitle) !== requestedKey) {
      return;
    }

    const requestId = cleanSheetWhitespace(record.graphicsRequestId);
    if (!requestId) return;

    if (!grouped.has(requestId)) {
      grouped.set(requestId, {
        graphicsRequestId: requestId,
        bookTitle: record.bookTitle,
        poemTitle: record.poemTitle,
        author: record.author,
        rejectReason: record.rejectReason || "",
        metadataIssue: record.metadataIssue || "",
        aestheticIssue: record.aestheticIssue || "",
        qcNote: record.qcNote || "",
        sourceRequestId: record.sourceRequestId || "",
        sourceCompletionId: record.sourceCompletionId || "",
        records: []
      });
    }

    const group = grouped.get(requestId);
    group.bookTitle = choosePreferredBookTitle(group.bookTitle, record.bookTitle);
    group.rejectReason = group.rejectReason || record.rejectReason || "";
    group.metadataIssue = group.metadataIssue || record.metadataIssue || "";
    group.aestheticIssue = group.aestheticIssue || record.aestheticIssue || "";
    group.qcNote = group.qcNote || record.qcNote || "";
    group.sourceRequestId = group.sourceRequestId || record.sourceRequestId || "";
    group.sourceCompletionId = group.sourceCompletionId || record.sourceCompletionId || "";
    group.records.push(record);
  });

  return Array.from(grouped.values())
    .map(group => {
      const completion = completionSummaryByRequestId.get(group.graphicsRequestId);
      return buildPigGraphicsRequestSummary({
        ...group,
        requestStatus: group.records.some(record => cleanSheetWhitespace(record.source).toLowerCase() === "weaver_qc_rework")
          ? "rework_requested"
          : (completion ? "completed_returned" : "open"),
        completionCount: completion?.count || 0,
        latestCompletedAt: completion?.latestCompletedAt || "",
        assetUrl: completion?.latestAssetUrl || "",
        assetPreviewUrl: completion?.latestAssetPreviewUrl || "",
        latestCompletionId: completion?.latestCompletionId || ""
      });
    })
    .sort((left, right) => left.bookTitle.localeCompare(right.bookTitle) || left.poemTitle.localeCompare(right.poemTitle));
}

async function getGraphicsBooksFromSheets(mode) {
  const resolvedMode = cleanSheetWhitespace(mode).toLowerCase() || "queue";
  let books = [];

  if (resolvedMode === "qc") {
    const pendingRecords = await getPendingGraphicsQcRecords({ includeCleanup: false });
    books = summarizeBooks(pendingRecords.map(record => ({ bookTitle: record.bookTitle })));
  } else if (resolvedMode === "cleanup") {
    const pendingRecords = await getPendingGraphicsQcRecords({ includeCleanup: true });
    books = summarizeBooks(pendingRecords.map(record => ({ bookTitle: record.bookTitle })));
  } else if (resolvedMode === "mismatch") {
    const mismatchRecords = await getPigMismatchRecords("all");
    books = summarizeBooks(mismatchRecords.map(record => ({ bookTitle: record.bookTitle })));
  } else if (resolvedMode === "handoff") {
    const handoffRecords = await getPoetryPleaseHandoffRecords();
    books = summarizeBooks(handoffRecords.map(record => ({ bookTitle: record.bookTitle })));
  } else if (resolvedMode === "handoff_retry") {
    const failedHandoffRecords = await getFailedPoetryPleaseHandoffRecords();
    books = summarizeBooks(failedHandoffRecords.map(record => ({ bookTitle: record.bookTitle })));
  } else if (resolvedMode === "rework") {
    const reworkRecords = await getManualGraphicsReworkCandidates();
    books = summarizeBooks(reworkRecords.map(record => ({ bookTitle: record.bookTitle })));
  } else if (resolvedMode === "coverage") {
    const rows = await getSourceSheetValuesCached();
    books = summarizeBooks(rows
      .map(row => buildCoverageRecordFromSourceRow(row))
      .filter(Boolean)
      .map(record => ({ bookTitle: record.bookTitle })));
  } else {
    const openRequests = await getPigGraphicsRequests("all");
    books = summarizeBooks(openRequests.map(record => ({ bookTitle: record.bookTitle })));
  }

  return {
    ok: true,
    version: `${appVersion}-service-account`,
    mode: resolvedMode,
    books
  };
}

function buildCoverageRecordFromSourceRow(row) {
  const config = SHEET_SOURCE_CONFIG.columnMap;
  const bookTitle = cleanSheetWhitespace(
    row[config.correctedBookTitle - 1]
    || row[config.validationCanonicalBook - 1]
    || row[config.bookTitle - 1]
    || row[7]
    || row[13]
  );
  const excerptText = String(
    row[config.correctedExcerpt - 1]
    || row[config.excerpt - 1]
    || row[6]
    || row[12]
    || ""
  ).trim();
  if (!bookTitle || !excerptText || isSheetYes(row[config.exclude - 1])) {
    return null;
  }

  return {
    timestamp: cleanSheetWhitespace(row[0]),
    email: cleanSheetWhitespace(row[1]).toLowerCase(),
    author: resolveGraphicsAuthor(
      row[config.correctedAuthor - 1]
      || row[config.validationCanonicalAuthor - 1]
      || row[config.author - 1]
      || row[3]
      || row[9]
      || "",
      bookTitle
    ),
    poemTitle: cleanSheetWhitespace(
      row[config.correctedTitle - 1]
      || row[config.validationMatchedPoemTitle - 1]
      || row[config.title - 1]
      || row[5]
      || row[11]
    ),
    bookTitle,
    excerptText
  };
}

async function getBookCoverageFromSheets(bookTitle) {
  const requestedKey = normalizeBookKey(bookTitle);
  const rows = (await getSourceSheetValuesCached())
    .map(row => buildCoverageRecordFromSourceRow(row))
    .filter(record => record && normalizeBookKey(record.bookTitle) === requestedKey);

  const preferredBookTitle = rows.reduce(
    (current, record) => choosePreferredBookTitle(current, record.bookTitle),
    cleanSheetWhitespace(bookTitle)
  );

  let catalogPoems = [];
  try {
    const catalogResult = await getExcerptGatheringPoemsForBook(preferredBookTitle);
    catalogPoems = Array.isArray(catalogResult?.poems) ? catalogResult.poems : [];
  } catch (_error) {
    catalogPoems = [];
  }

  const poemStats = new Map();
  const contributorStats = new Map();

  rows.forEach(record => {
    const poemKey = normalizeBookKey(record.poemTitle || "__untitled__");
    if (!poemStats.has(poemKey)) {
      poemStats.set(poemKey, {
        poemTitle: record.poemTitle || "Untitled poem",
        excerptCount: 0,
        totalWords: 0,
        contributors: new Set()
      });
    }
    const poem = poemStats.get(poemKey);
    poem.excerptCount += 1;
    poem.totalWords += String(record.excerptText || "").trim().split(/\s+/).filter(Boolean).length;
    if (record.email) poem.contributors.add(record.email);

    const contributorKey = record.email || "__unknown__";
    if (!contributorStats.has(contributorKey)) {
      contributorStats.set(contributorKey, {
        email: record.email || "Unknown",
        excerptCount: 0,
        poems: new Set(),
        totalWords: 0,
        firstTimestamp: "",
        lastTimestamp: ""
      });
    }
    const contributor = contributorStats.get(contributorKey);
    contributor.excerptCount += 1;
    contributor.totalWords += String(record.excerptText || "").trim().split(/\s+/).filter(Boolean).length;
    contributor.poems.add(record.poemTitle || "Untitled poem");
    if (record.timestamp && (!contributor.firstTimestamp || record.timestamp < contributor.firstTimestamp)) {
      contributor.firstTimestamp = record.timestamp;
    }
    if (record.timestamp && (!contributor.lastTimestamp || record.timestamp > contributor.lastTimestamp)) {
      contributor.lastTimestamp = record.timestamp;
    }
  });

  const catalogTitles = catalogPoems
    .map(poem => cleanSheetWhitespace(poem.title))
    .filter(Boolean);
  const coveredCatalogKeys = new Set(Array.from(poemStats.values()).map(poem => normalizeBookKey(poem.poemTitle)));
  const missingPoems = catalogTitles.filter(title => !coveredCatalogKeys.has(normalizeBookKey(title)));

  return {
    ok: true,
    version: `${appVersion}-service-account`,
    bookTitle: preferredBookTitle,
    summary: {
      totalExcerpts: rows.length,
      totalCatalogPoems: catalogTitles.length,
      poemsWithExcerpts: poemStats.size,
      poemsWithoutExcerpts: missingPoems.length,
      contributors: contributorStats.size
    },
    topPoems: Array.from(poemStats.values())
      .map(poem => ({
        poemTitle: poem.poemTitle,
        excerptCount: poem.excerptCount,
        contributorCount: poem.contributors.size,
        averageWords: poem.excerptCount ? Math.round(poem.totalWords / poem.excerptCount) : 0
      }))
      .sort((left, right) => right.excerptCount - left.excerptCount || left.poemTitle.localeCompare(right.poemTitle))
      .slice(0, 12),
    missingPoems: missingPoems.slice(0, 30),
    contributors: Array.from(contributorStats.values())
      .map(person => ({
        email: person.email,
        excerptCount: person.excerptCount,
        uniquePoems: person.poems.size,
        averageWords: person.excerptCount ? Math.round(person.totalWords / person.excerptCount) : 0,
        firstTimestamp: person.firstTimestamp,
        lastTimestamp: person.lastTimestamp
      }))
      .sort((left, right) => right.excerptCount - left.excerptCount || left.email.localeCompare(right.email))
  };
}

async function getGraphicsRecordsForBookFromSheets(bookTitle, mode) {
  const qcSweepBatchSize = 5;
  const requestedKey = normalizeBookKey(bookTitle);
  const resolvedMode = cleanSheetWhitespace(mode).toLowerCase() || "queue";
  const isQcSweep = ["qc", "cleanup"].includes(resolvedMode) && cleanSheetWhitespace(bookTitle) === "__qc_sweep__";
  let preferredBookTitle = cleanSheetWhitespace(bookTitle);
  let records = [];

  if (resolvedMode === "coverage") {
    const coverage = await getBookCoverageFromSheets(bookTitle);
    return {
      ok: true,
      version: coverage.version,
      mode: resolvedMode,
      bookTitle: coverage.bookTitle,
      records: [],
      coverage
    };
  }

  if (resolvedMode === "qc") {
    records = await getPendingGraphicsQcRecords({ includeCleanup: false });
    if (isQcSweep) {
      records = records
        .slice()
        .sort((left, right) => {
          const leftTime = Date.parse(left.completedAt || left.graphicsQcUpdatedAt || "") || 0;
          const rightTime = Date.parse(right.completedAt || right.graphicsQcUpdatedAt || "") || 0;
          if (leftTime !== rightTime) {
            return leftTime - rightTime;
          }
          return Number(left.sheetRow || 0) - Number(right.sheetRow || 0);
        })
        .slice(0, qcSweepBatchSize);
      preferredBookTitle = "QC Sweep";
    } else {
      records = records.filter(record => normalizeBookKey(record.bookTitle) === requestedKey);
    }
  } else if (resolvedMode === "cleanup") {
    records = await getPendingGraphicsQcRecords({ includeCleanup: true });
    if (isQcSweep) {
      records = records
        .slice()
        .sort((left, right) => {
          const leftTime = Date.parse(left.completedAt || left.graphicsQcUpdatedAt || "") || 0;
          const rightTime = Date.parse(right.completedAt || right.graphicsQcUpdatedAt || "") || 0;
          if (leftTime !== rightTime) {
            return leftTime - rightTime;
          }
          return Number(left.sheetRow || 0) - Number(right.sheetRow || 0);
        })
        .slice(0, qcSweepBatchSize);
      preferredBookTitle = "QC Sweep";
    } else {
      records = records.filter(record => normalizeBookKey(record.bookTitle) === requestedKey);
    }
  } else if (resolvedMode === "mismatch") {
    records = (await getPigMismatchRecords("all"))
      .filter(record => normalizeBookKey(record.bookTitle) === requestedKey);
  } else if (resolvedMode === "handoff") {
    records = await getPoetryPleaseHandoffRecords(bookTitle);
  } else if (resolvedMode === "handoff_retry") {
    records = await getFailedPoetryPleaseHandoffRecords(bookTitle);
  } else if (resolvedMode === "rework") {
    records = await getManualGraphicsReworkCandidates(bookTitle);
  } else {
    records = (await getPigGraphicsRequests("all"))
      .filter(record => normalizeBookKey(record.bookTitle) === requestedKey);
  }

  preferredBookTitle = records.reduce(
    (current, record) => choosePreferredBookTitle(current, record.bookTitle),
    preferredBookTitle
  );

  return {
    ok: true,
    version: `${appVersion}-service-account`,
    mode: resolvedMode,
    bookTitle: preferredBookTitle,
    records
  };
}

async function upsertPigCompletedGraphics(completions) {
  if (!Array.isArray(completions) || !completions.length) {
    return { ok: false, error: "No completed graphics provided." };
  }

  const rows = await readPigCompletedGraphicsRows();
  const existingById = new Map();
  rows.forEach((row, index) => {
    const completionId = cleanSheetWhitespace(row[PIG_COMPLETION_COLUMNS.completionId - 1]);
    if (!completionId) return;
    existingById.set(completionId, { row, rowNumber: index + 2 });
  });

  const updates = [];
  const appendRows = [];

  completions.forEach(completion => {
    const completionId = buildPigCompletionId(completion);
    const existing = existingById.get(completionId);
    const rowValues = buildPigCompletionRowValues({ ...completion, completionId }, existing?.row || []);
    if (existing) {
      updates.push({
        range: `'${pigCompletedGraphicsSheetName.replace(/'/g, "''")}'!A${existing.rowNumber}:S${existing.rowNumber}`,
        values: [rowValues]
      });
      return;
    }
    appendRows.push(rowValues);
  });

  let runtimeDb = { ok: false, skipped: true };
  try {
    runtimeDb = await syncWeaverRuntimeDb("upsert_completions", { completions });
  } catch (error) {
    runtimeDb = { ok: false, error: error.message };
  }
  if (!runtimeDb?.ok) {
    throw new Error(runtimeDb?.error || "Runtime DB ledger write failed.");
  }

  const sheetSync = { ok: true, wroteUpdates: updates.length, wroteAppends: appendRows.length };
  try {
    if (updates.length) {
      await batchUpdateSheetValuesServer(updates);
    }
    if (appendRows.length) {
      await appendSheetValuesServer(`'${pigCompletedGraphicsSheetName.replace(/'/g, "''")}'!A:S`, appendRows);
    }
  } catch (error) {
    sheetSync.ok = false;
    sheetSync.error = error.message;
  }

  return {
    ok: true,
    version: `${appVersion}-service-account`,
    savedCount: completions.length,
    runtimeDb,
    sheetSync
  };
}

function buildReviewWriteRanges(update) {
  const sourceRow = parseInt(update?.sourceRow, 10);
  if (!sourceRow || sourceRow < SHEET_SOURCE_CONFIG.startRow) {
    return [];
  }

  const reviewDecision = normalizeReviewDecisionValue(update.reviewDecision || update.approval);
  if (!reviewDecision) {
    return [];
  }
  const needsCorrection = reviewDecision === "NEEDS_CORRECTION";
  const useForQi = isTruthyParam(update.useForQi || update.graphicsQi);
  const useForInt = isTruthyParam(update.useForInt || update.photos);
  const correctedAuthor = String(update.correctedAuthor || "").trim();
  const correctedTitle = String(update.correctedTitle || "").trim();
  const correctedBookTitle = String(update.correctedBookTitle || "").trim();
  const correctedExcerpt = String(update.correctedExcerpt || "");
  const correctionNote = needsCorrection ? String(update.correctionNote || "").trim() : "";
  const hasDuplicateGroupId = Object.prototype.hasOwnProperty.call(update || {}, "duplicateGroupId");
  const duplicateGroupId = hasDuplicateGroupId ? String(update.duplicateGroupId || "").trim() : "";

  const writes = [
    { column: SHEET_SOURCE_CONFIG.columnMap.approved, value: useForQi ? "Y" : "N" },
    { column: SHEET_SOURCE_CONFIG.columnMap.statusIndicator, value: needsCorrection ? "NEEDS_CORRECTION" : "" },
    { column: SHEET_SOURCE_CONFIG.columnMap.correctionNote, value: correctionNote },
    { column: SHEET_SOURCE_CONFIG.columnMap.excerptReviewDecision, value: reviewDecision },
    { column: SHEET_SOURCE_CONFIG.columnMap.useForInt, value: useForInt ? "Y" : "" },
    { column: SHEET_SOURCE_CONFIG.columnMap.correctedAuthor, value: correctedAuthor },
    { column: SHEET_SOURCE_CONFIG.columnMap.correctedTitle, value: correctedTitle },
    { column: SHEET_SOURCE_CONFIG.columnMap.correctedBookTitle, value: correctedBookTitle },
    { column: SHEET_SOURCE_CONFIG.columnMap.correctedExcerpt, value: correctedExcerpt }
  ];

  if (hasDuplicateGroupId) {
    writes.push({
      column: SHEET_SOURCE_CONFIG.columnMap.duplicateGroupId,
      value: duplicateGroupId
    });
  }

  return writes.map(write => ({
    range: `'${sourceSheetName.replace(/'/g, "''")}'!${toA1Column(write.column)}${sourceRow}`,
    values: [[write.value]]
  }));
}

async function saveReviewsToSheets(updates) {
  const validUpdates = (Array.isArray(updates) ? updates : []).filter(update => buildReviewWriteRanges(update).length);
  const requests = validUpdates.flatMap(buildReviewWriteRanges);
  if (!requests.length) {
    return { ok: false, error: "No updates provided." };
  }

  await batchUpdateSheetValuesServer(requests);
  const excerptHandoffs = await syncAcceptedExcerptHandoffs(validUpdates);
  return {
    ok: true,
    version: `${appVersion}-service-account`,
    savedCount: validUpdates.length,
    excerptHandoffs
  };
}

async function saveSingleReviewToSheets(update) {
  const requests = buildReviewWriteRanges(update);
  if (!requests.length) {
    return { ok: false, error: "Missing source row target." };
  }

  await batchUpdateSheetValuesServer(requests);
  const excerptHandoffs = await syncAcceptedExcerptHandoffs([update]);
  return {
    ok: true,
    version: `${appVersion}-service-account`,
    sourceRow: parseInt(update.sourceRow, 10) || 0,
    recordId: String(update.recordId || ""),
    excerptHandoffs
  };
}

async function enrichAcceptedExcerptUpdate(update) {
  const normalized = update && typeof update === "object" ? { ...update } : {};
  const sourceRow = parseInt(normalized?.sourceRow, 10);
  if (!sourceRow || sourceRow < SHEET_SOURCE_CONFIG.startRow) {
    return normalized;
  }
  if (
    cleanSheetWhitespace(normalized.excerptText) &&
    cleanSheetWhitespace(normalized.bookTitle) &&
    cleanSheetWhitespace(normalized.bookShortener) &&
    cleanSheetWhitespace(normalized.releaseCatalog)
  ) {
    return normalized;
  }

  const range = `'${sourceSheetName.replace(/'/g, "''")}'!A${sourceRow}:BK${sourceRow}`;
  const values = await fetchSheetValuesServer(range);
  const row = Array.isArray(values) ? values[0] : null;
  if (!row) {
    return normalized;
  }
  const noteMeta = parseIntakeMetadataFromNotes(row[8] || "");
  const canonicalBookAuthorMap = await getCanonicalGraphicsBookAuthorMap().catch(() => new Map());
  const record = buildPendingRecordFromSheetRow(row, sourceRow - SHEET_SOURCE_CONFIG.startRow, canonicalBookAuthorMap) || {
    sourceRow,
    recordId: (row[SHEET_SOURCE_CONFIG.columnMap.recordId - 1] || "").toString(),
    author: resolveGraphicsAuthor(row[SHEET_SOURCE_CONFIG.columnMap.author - 1] || "", cleanSheetWhitespace(row[SHEET_SOURCE_CONFIG.columnMap.bookTitle - 1] || ""), canonicalBookAuthorMap),
    title: (row[SHEET_SOURCE_CONFIG.columnMap.title - 1] || "").toString(),
    bookTitle: cleanSheetWhitespace(row[SHEET_SOURCE_CONFIG.columnMap.bookTitle - 1] || ""),
    excerptText: (row[SHEET_SOURCE_CONFIG.columnMap.excerpt - 1] || "").toString()
  };

  normalized.recordId = cleanSheetWhitespace(normalized.recordId) || record.recordId || "";
  normalized.author = cleanSheetWhitespace(normalized.author) || record.author || "";
  normalized.title = cleanSheetWhitespace(normalized.title) || record.title || "";
  normalized.poemTitle = cleanSheetWhitespace(normalized.poemTitle) || record.title || "";
  normalized.bookTitle = cleanSheetWhitespace(normalized.bookTitle) || record.bookTitle || "";
  normalized.excerptText = normalized.excerptText || record.excerptText || "";
  normalized.contentType = normalizeExcerptContentType(normalized.contentType || noteMeta.contentType || record.contentType || "EXC");
  const bookMeta = resolvePublishingBookMeta(normalized.bookTitle, normalized.bookTitle);
  normalized.bookShortener = cleanSheetWhitespace(normalized.bookShortener) || noteMeta.bookShortener || cleanSheetWhitespace(record.bookShortener) || cleanSheetWhitespace(bookMeta?.bookShortener);
  normalized.releaseCatalog = cleanSheetWhitespace(normalized.releaseCatalog) || noteMeta.releaseCatalog || cleanSheetWhitespace(record.releaseCatalog) || cleanSheetWhitespace(bookMeta?.releaseCatalog);
  normalized.socialMediaHandle = cleanSheetWhitespace(normalized.socialMediaHandle || normalized.instagramHandle || normalized.igHandle)
    || noteMeta.socialMediaHandle
    || cleanSheetWhitespace(record.socialMediaHandle)
    || "";
  return normalized;
}

function buildAcceptedExcerptHandoff(update) {
  const reviewDecision = normalizeReviewDecisionValue(update?.reviewDecision || update?.approval);
  if (reviewDecision !== "ACCEPT" && reviewDecision !== "ACCEPT_SKIP_GRAPHIC") {
    return null;
  }

  const sourceRow = parseInt(update?.sourceRow, 10);
  const sourceRecordId = String(update?.recordId || "").trim() || (sourceRow ? `weaver:row-${sourceRow}` : "");
  const excerptText = String(update?.excerptText || "").trim();
  if (!sourceRecordId || !excerptText) {
    return null;
  }

  const now = new Date().toISOString();
  const contentType = normalizeExcerptContentType(update?.contentType || "EXC");
  return {
    recordId: `weaver-${contentType.toLowerCase()}-${sourceRecordId}`,
    contentType,
    sourceSystem: "weaver",
    sourceRecordId,
    author: String(update?.author || "").trim(),
    bookTitle: String(update?.bookTitle || "").trim(),
    poemTitle: String(update?.poemTitle || update?.title || "").trim(),
    excerpt: excerptText,
    handoffStatus: "queued",
    handoffMode: "auto",
    approvedAt: now,
    updatedAt: now,
    bookShortener: cleanSheetWhitespace(update?.bookShortener),
    bookLink: cleanSheetWhitespace(update?.bookLink),
    releaseCatalog: cleanSheetWhitespace(update?.releaseCatalog),
    socialMediaHandle: cleanSheetWhitespace(update?.socialMediaHandle || update?.instagramHandle || update?.igHandle),
    instagramHandle: cleanSheetWhitespace(update?.socialMediaHandle || update?.instagramHandle || update?.igHandle),
    igHandle: cleanSheetWhitespace(update?.socialMediaHandle || update?.instagramHandle || update?.igHandle),
    driveLink: cleanSheetWhitespace(update?.driveLink),
    sourceUrl: cleanSheetWhitespace(update?.sourceUrl),
    pageNumber: cleanSheetWhitespace(update?.pageNumber),
    payload: {
      sourceRow: sourceRow || 0,
      sourceRecordId,
      reviewDecision: normalizeDecision(update?.reviewDecision || update?.approval) || "accept"
    }
  };
}

function buildExcerptHandoffFromApprovedExportRecord(record) {
  const sourceRecordId = cleanSheetWhitespace(record?.sourceRecordId);
  const excerptText = normalizeExcerptTransferText(record?.excerptText);
  if (!sourceRecordId || !cleanSheetWhitespace(excerptText)) {
    return null;
  }

  const approvedAt =
    cleanSheetWhitespace(record?.sourceApprovedAt) ||
    cleanSheetWhitespace(record?.sourceUpdatedAt) ||
    new Date().toISOString();
  const updatedAt =
    cleanSheetWhitespace(record?.sourceUpdatedAt) ||
    approvedAt;
  const sheetHandoffStatus = normalizeExcerptHandoffStatus(record?.poetryPleaseStatus);
  const sheetHandoffNote = String(record?.poetryPleaseNote || "");
  const itemIdMatch = sheetHandoffNote.match(/item=([^\s]+)/);

  return {
    recordId: `weaver-${normalizeExcerptContentType(record?.contentType).toLowerCase()}-${sourceRecordId}`,
    contentType: normalizeExcerptContentType(record?.contentType),
    sourceSystem: "weaver",
    sourceRecordId,
    author: cleanSheetWhitespace(record?.author),
    bookTitle: cleanSheetWhitespace(record?.bookTitle),
    poemTitle: cleanSheetWhitespace(record?.poemTitle),
    excerpt: excerptText,
    handoffStatus: sheetHandoffStatus || "queued",
    handoffMode: "backfill",
    handedOffAt: cleanSheetWhitespace(record?.poetryPleaseUpdatedAt),
    poetryPleaseItemId: itemIdMatch ? cleanSheetWhitespace(itemIdMatch[1]) : "",
    errorMessage: sheetHandoffStatus === "failed" ? sheetHandoffNote : "",
    approvedAt,
    updatedAt,
    bookShortener: cleanSheetWhitespace(record?.bookShortener),
    bookLink: cleanSheetWhitespace(record?.bookLink),
    releaseCatalog: cleanSheetWhitespace(record?.releaseCatalog),
    socialMediaHandle: cleanSheetWhitespace(record?.socialMediaHandle || record?.instagramHandle || record?.igHandle),
    instagramHandle: cleanSheetWhitespace(record?.socialMediaHandle || record?.instagramHandle || record?.igHandle),
    igHandle: cleanSheetWhitespace(record?.socialMediaHandle || record?.instagramHandle || record?.igHandle),
    driveLink: cleanSheetWhitespace(record?.driveLink),
    sourceUrl: cleanSheetWhitespace(record?.sourceUrl),
    pageNumber: cleanSheetWhitespace(record?.pageNumber),
    payload: {
      sourceRow: parseInt(record?.sourceRow, 10) || 0,
      sourceRecordId,
      reviewDecision: "accept"
    }
  };
}

async function syncAcceptedExcerptHandoffs(updates) {
  const normalizedUpdates = await Promise.all((Array.isArray(updates) ? updates : []).map(enrichAcceptedExcerptUpdate));
  const handoffs = normalizedUpdates
    .map(buildAcceptedExcerptHandoff)
    .filter(Boolean);
  if (!handoffs.length) {
    return { ok: true, savedCount: 0, records: [] };
  }

  try {
    const result = await syncWeaverRuntimeDb("upsert_excerpt_handoffs", { handoffs });
    const records = Array.isArray(result?.records) ? result.records : [];
    let poetryPlease = { ok: true, skipped: true, reason: "no_records" };
    try {
      poetryPlease = await handoffApprovedExcerptsToPoetryPlease(records);
    } catch (error) {
      poetryPlease = { ok: false, error: error.message };
    }

    if (records.length) {
      const statusResult = await updateExcerptHandoffStatuses(records, poetryPlease);
      return {
        ok: Boolean(result?.ok),
        savedCount: statusResult.savedCount,
        records: statusResult.records,
        poetryPlease
      };
    }

    return {
      ok: Boolean(result?.ok),
      savedCount: records.length || handoffs.length,
      records,
      poetryPlease
    };
  } catch (error) {
    return {
      ok: false,
      error: error.message,
      savedCount: 0,
      records: []
    };
  }
}

function normalizeGraphicsQcDecision(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "approve") return "APPROVE";
  if (normalized === "replace" || normalized === "replace_here") return "APPROVE";
  if (normalized === "mismatch" || normalized === "mismatched_graphic") return "REJECT";
  if (normalized === "correct" || normalized === "recreate" || normalized === "correct_and_recreate") {
    return "REJECT";
  }
  if (normalized === "reject") return "REJECT";
  return "";
}

async function buildReplacementGraphicCompletion(update) {
  const replacementAssetUrl = cleanSheetWhitespace(update?.replacementAssetUrl);
  const folderId = extractGoogleDriveFolderId(replacementAssetUrl);
  const fileId = extractGoogleDriveFileId(replacementAssetUrl);
  if (!folderId && !fileId) {
    throw new Error("Replacement graphic must be a Google Drive image file or folder link.");
  }

  let file = null;
  if (folderId) {
    const files = await listDriveFolderImageFiles(folderId);
    file = chooseMatchingDriveFolderImageFile(files, update) || chooseBestDriveFolderImageFile(files);
    if (!file) {
      throw new Error("No image files were found in that Google Drive folder.");
    }
  } else {
    try {
      file = await getDriveFileMetadata(fileId);
      if (!String(file?.mimeType || "").startsWith("image/")) {
        throw new Error("Replacement graphic link must point to an image file in Google Drive.");
      }
    } catch (_error) {
      file = {
        id: fileId,
        name: "",
        mimeType: "image/*",
        webViewLink: replacementAssetUrl,
        thumbnailLink: ""
      };
    }
  }

  const resolvedFileId = cleanSheetWhitespace(file?.id);
  const assetUrl = cleanSheetWhitespace(file?.webViewLink) || `https://drive.google.com/file/d/${encodeURIComponent(resolvedFileId)}/view`;
  const assetPreviewUrl = cleanSheetWhitespace(file?.thumbnailLink) || `https://drive.google.com/thumbnail?id=${encodeURIComponent(resolvedFileId)}&sz=w1600`;
  const fileName = cleanSheetWhitespace(file?.name);

  return {
    completionId: cleanSheetWhitespace(update?.pigCompletionId),
    requestId: cleanSheetWhitespace(update?.graphicsRequestId),
    author: String(update?.author || ""),
    poemTitle: String(update?.poemTitle || ""),
    bookTitle: String(update?.bookTitle || ""),
    quoteText: String(update?.quoteText || ""),
    sourceRecordId: cleanSheetWhitespace(update?.recordId).replace(/^pig:/i, ""),
    sourceSheetRow: parseInt(update?.sheetRow, 10) || "",
    assetUrl,
    assetPreviewUrl,
    productionNotes: fileName
      ? `Replacement graphic approved in Weaver (${fileName}).`
      : "Replacement graphic approved in Weaver.",
    completedAt: new Date().toISOString(),
    sourceTool: "Weaver replacement"
  };
}

function normalizeGraphicsQcRejectReason(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "mismatch" || normalized === "mismatched_graphic") return "mismatched_graphic";
  if (normalized === "correct" || normalized === "recreate" || normalized === "correct_and_recreate") {
    return "correct_and_recreate";
  }
  if (normalized === "reject" || normalized === "final_reject") return "final_reject";
  return "";
}

function parseGraphicsQcStructuredNoteServer(note) {
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
      const match = Array.from(GRAPHICS_QC_REJECT_REASON_LABELS.entries()).find(([, value]) => value === label);
      if (match) {
        parsed.rejectReason = match[0];
        return;
      }
    }
    if (line.startsWith("Metadata issue: ")) {
      parsed.metadataIssue = line.slice("Metadata issue: ".length).trim();
      return;
    }
    if (line.startsWith("Aesthetic issue: ")) {
      parsed.aestheticIssue = line.slice("Aesthetic issue: ".length).trim();
      return;
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

function getGraphicsQcRejectReason(record = {}) {
  const parsed = parseGraphicsQcStructuredNoteServer(record.graphicsQcNote || "");
  if (parsed.rejectReason) {
    return parsed.rejectReason;
  }
  return normalizeGraphicsQcRejectReason(record.rejectReason || record.graphicsQcDecision || "");
}

async function saveGraphicsQcToSheets(updates) {
  if (!Array.isArray(updates) || !updates.length) {
    return { ok: false, error: "No graphics QC updates provided." };
  }

  const storageTargets = new Set(
    updates.map(update => cleanSheetWhitespace(update?.storageTarget).toLowerCase())
  );
  if (storageTargets.has("sheet_cleanup") || storageTargets.has("cleanup_sheet")) {
    await ensureCleanupQcColumnsServer();
  }
  if (storageTargets.has("pig_sheet")) {
    await ensurePigCompletedGraphicsSheetServer();
  }

  const cleanupRequests = [];
  const pigRequests = [];
  const approvedPigSheetRows = [];
  const approvedFirestoreRecords = [];
  const replacementCompletions = [];
  const matchedUpdates = [];
  let firestoreMatchedCount = 0;
  let savedCount = 0;

  for (const update of updates) {
    const decision = normalizeGraphicsQcDecision(update?.qcDecision);
    const note = String(update?.qcNote || "").trim();
    const updatedAt = decision ? new Date().toISOString() : "";
    const rowNumber = parseInt(update?.sheetRow, 10) || 0;
    const storageTarget = cleanSheetWhitespace(update?.storageTarget).toLowerCase();
    const completionId = cleanSheetWhitespace(update?.pigCompletionId);

    if (storageTarget === "firestore") {
      if (!completionId || !decision) continue;
      const normalizedUpdate = {
        ...update,
        graphicsQcDecision: decision,
        graphicsQcNote: note,
        graphicsQcUpdatedAt: updatedAt
      };
      matchedUpdates.push(normalizedUpdate);
      firestoreMatchedCount++;
      if (decision === "APPROVE") {
        approvedFirestoreRecords.push(normalizedUpdate);
      }
      savedCount++;
      continue;
    }

    if (!rowNumber) continue;

    const replacementDecision = String(update?.qcDecision || "").trim().toLowerCase();
    if (replacementDecision === "replace") {
      if (cleanSheetWhitespace(update?.storageTarget).toLowerCase() !== "pig_sheet") {
        throw new Error("Replacement graphics currently only work for returned P.I.G. graphics.");
      }
      if (!cleanSheetWhitespace(update?.pigCompletionId)) {
        throw new Error("Replacement graphics need an existing P.I.G. completion id.");
      }
      replacementCompletions.push(await buildReplacementGraphicCompletion(update));
    }

    if (storageTarget === "pig_sheet") {
      if (decision === "APPROVE") {
        approvedPigSheetRows.push(rowNumber);
      }
      pigRequests.push(
        {
          range: `'${pigCompletedGraphicsSheetName.replace(/'/g, "''")}'!M${rowNumber}`,
          values: [[decision]]
        },
        {
          range: `'${pigCompletedGraphicsSheetName.replace(/'/g, "''")}'!N${rowNumber}`,
          values: [[note]]
        },
        {
          range: `'${pigCompletedGraphicsSheetName.replace(/'/g, "''")}'!O${rowNumber}`,
          values: [[updatedAt]]
        }
      );
    } else {
      cleanupRequests.push(
        {
          range: `'${graphicsCleanupSheetName.replace(/'/g, "''")}'!J${rowNumber}`,
          values: [[decision]]
        },
        {
          range: `'${graphicsCleanupSheetName.replace(/'/g, "''")}'!K${rowNumber}`,
          values: [[note]]
        },
        {
          range: `'${graphicsCleanupSheetName.replace(/'/g, "''")}'!L${rowNumber}`,
          values: [[updatedAt]]
        }
      );
    }
    matchedUpdates.push(update);
    savedCount++;
  }

  if (!matchedUpdates.length) {
    return { ok: false, error: "No matching graphics QC rows found." };
  }

  if (replacementCompletions.length) {
    await upsertPigCompletedGraphics(replacementCompletions);
  }

  let runtimeDb = { ok: false, skipped: true };
  if (matchedUpdates.length) {
    try {
      runtimeDb = await syncWeaverRuntimeDb("insert_qc_reviews", { reviews: matchedUpdates });
    } catch (error) {
      runtimeDb = { ok: false, error: error.message };
    }
    if (!runtimeDb?.ok) {
      throw new Error(runtimeDb?.error || "Runtime DB QC ledger write failed.");
    }
    if (firestoreMatchedCount && Number(runtimeDb?.written || 0) < firestoreMatchedCount) {
      throw new Error("Firestore QC ledger did not save every submitted review.");
    }
  }
  const sheetSync = { ok: true, cleanupWrites: cleanupRequests.length, pigWrites: pigRequests.length };
  try {
    if (cleanupRequests.length) {
      await batchUpdateSheetValuesServer(cleanupRequests);
    }
    if (pigRequests.length) {
      await batchUpdateSheetValuesServer(pigRequests);
    }
  } catch (error) {
    sheetSync.ok = false;
    sheetSync.error = error.message;
  }

  let poetryPlease = { ok: true, skipped: true, reason: "no_new_approvals" };
  if ((approvedPigSheetRows.length || approvedFirestoreRecords.length) && sheetSync.ok) {
    const approvedRowSet = new Set(approvedPigSheetRows.map(value => String(value)));
    const approvedSheetRecords = approvedPigSheetRows.length ? (await readPigCompletedGraphicsRows())
      .map((row, index) => buildPigQcRecordFromSheetRow(row, index))
      .filter(record => record && approvedRowSet.has(String(record.sheetRow)))
      .filter(record => normalizeGraphicsQcDecision(record.graphicsQcDecision) === "APPROVE")
      .map(buildPoetryPleaseGraphicRecord)
      .filter(Boolean) : [];
    const approvedRecords = [
      ...approvedSheetRecords,
      ...approvedFirestoreRecords.map(buildPoetryPleaseGraphicRecord).filter(Boolean)
    ];

    try {
      poetryPlease = await handoffApprovedGraphicsToPoetryPlease(approvedRecords);
    } catch (error) {
      poetryPlease = {
        ok: false,
        error: error.message,
        responseStatus: error.responseStatus || 0,
        responseBody: error.responseBody || null,
        requestBody: error.requestBody || null
      };
    }

    const handoffUpdatedAt = new Date().toISOString();
    const handoffStatus = poetryPlease.ok ? "HANDED_OFF" : "FAILED";
    const handoffNote = poetryPlease.ok
      ? `created=${Number(poetryPlease.createdCount || 0)} updated=${Number(poetryPlease.updatedCount || 0)} errors=${Number(poetryPlease.errorCount || 0)}`
      : String(poetryPlease.error || poetryPlease.reason || "handoff_failed");

    const handoffRequests = approvedPigSheetRows.flatMap(rowNumber => ([
      {
        range: `'${pigCompletedGraphicsSheetName.replace(/'/g, "''")}'!Q${rowNumber}`,
        values: [[handoffStatus]]
      },
      {
        range: `'${pigCompletedGraphicsSheetName.replace(/'/g, "''")}'!R${rowNumber}`,
        values: [[handoffUpdatedAt]]
      },
      {
        range: `'${pigCompletedGraphicsSheetName.replace(/'/g, "''")}'!S${rowNumber}`,
        values: [[handoffNote]]
      }
    ]));
    if (handoffRequests.length) {
      await batchUpdateSheetValuesServer(handoffRequests);
    }

    try {
      await syncWeaverRuntimeDb("insert_poetry_please_handoffs", {
        handoffs: approvedRecords.map(record => ({
          graphicsCompletionId: record.sourceCompletionId,
          handoffStatus,
          handedOffAt: handoffUpdatedAt,
          handoffMode: "auto",
          payload: {
            qcApprovedAt: record.qcApprovedAt,
            sourceRequestId: record.sourceRequestId,
            createdCount: Number(poetryPlease.createdCount || 0),
            updatedCount: Number(poetryPlease.updatedCount || 0),
            errorCount: Number(poetryPlease.errorCount || 0),
            results: poetryPlease.results || [],
            responseStatus: poetryPlease.responseStatus || 0,
            responseBody: poetryPlease.responseBody || null,
            requestBody: poetryPlease.requestBody || null,
            note: handoffNote
          }
        }))
      });
    } catch {
      // Leave the sheet as source of truth if DB handoff shadow write fails.
    }
  }

  return {
    ok: true,
    version: `${appVersion}-service-account`,
    savedCount,
    runtimeDb,
    poetryPlease,
    sheetSync
  };
}

function runCatalogValidation(records) {
  return new Promise((resolve, reject) => {
    const child = spawn("python3", [path.join(__dirname, "catalog_validate.py")], {
      env: {
        ...process.env
      }
    });
    let stdout = "";
    let stderr = "";

    child.stdout.on("data", chunk => {
      stdout += chunk.toString();
    });
    child.stderr.on("data", chunk => {
      stderr += chunk.toString();
    });
    child.on("error", reject);
    child.on("close", code => {
      if (code !== 0) {
        reject(new Error(stderr || `catalog_validate.py exited with code ${code}`));
        return;
      }
      try {
        resolve(JSON.parse(stdout));
      } catch (error) {
        reject(error);
      }
    });

    child.stdin.write(JSON.stringify({ records }));
    child.stdin.end();
  });
}

function runCatalogPoemLookup(bookTitle, poemTitle, excerptText) {
  return new Promise((resolve, reject) => {
    const child = spawn("python3", [path.join(__dirname, "catalog_poem_text.py")], {
      env: {
        ...process.env
      }
    });
    let stdout = "";
    let stderr = "";

    child.stdout.on("data", chunk => {
      stdout += chunk.toString();
    });
    child.stderr.on("data", chunk => {
      stderr += chunk.toString();
    });
    child.on("error", reject);
    child.on("close", code => {
      if (code !== 0) {
        reject(new Error(stderr || `catalog_poem_text.py exited with code ${code}`));
        return;
      }
      try {
        resolve(JSON.parse(stdout));
      } catch (error) {
        reject(error);
      }
    });

    child.stdin.write(JSON.stringify({ bookTitle, poemTitle, excerptText }));
    child.stdin.end();
  });
}

function runLibraryExcerptLookup(sourceRow) {
  return new Promise((resolve, reject) => {
    const child = spawn("python3", [path.join(__dirname, "excerpt_library_text.py")], {
      env: {
        ...process.env
      }
    });
    let stdout = "";
    let stderr = "";

    child.stdout.on("data", chunk => {
      stdout += chunk.toString();
    });
    child.stderr.on("data", chunk => {
      stderr += chunk.toString();
    });
    child.on("error", reject);
    child.on("close", code => {
      if (code !== 0) {
        reject(new Error(stderr || `excerpt_library_text.py exited with code ${code}`));
        return;
      }
      try {
        resolve(JSON.parse(stdout));
      } catch (error) {
        reject(error);
      }
    });

    child.stdin.write(JSON.stringify({ sourceRow }));
    child.stdin.end();
  });
}

function runGraphicsAssetLookup(records) {
  return new Promise((resolve, reject) => {
    const child = spawn("python3", [path.join(__dirname, "graphics_qi_lookup.py")], {
      env: {
        ...process.env
      }
    });
    let stdout = "";
    let stderr = "";

    child.stdout.on("data", chunk => {
      stdout += chunk.toString();
    });
    child.stderr.on("data", chunk => {
      stderr += chunk.toString();
    });
    child.on("error", reject);
    child.on("close", code => {
      if (code !== 0) {
        reject(new Error(stderr || `graphics_qi_lookup.py exited with code ${code}`));
        return;
      }
      try {
        resolve(JSON.parse(stdout));
      } catch (error) {
        reject(error);
      }
    });

    child.stdin.write(JSON.stringify({ records }));
    child.stdin.end();
  });
}

function syncWeaverRuntimeDb(action, payload) {
  return new Promise((resolve, reject) => {
    const child = spawn("python3", [path.join(__dirname, "weaver_runtime_sync.py")], {
      env: {
        ...process.env
      }
    });
    let stdout = "";
    let stderr = "";

    child.stdout.on("data", chunk => {
      stdout += chunk.toString();
    });
    child.stderr.on("data", chunk => {
      stderr += chunk.toString();
    });
    child.on("error", reject);
    child.on("close", code => {
      if (code !== 0 && !stdout.trim()) {
        reject(new Error(stderr || `weaver_runtime_sync.py exited with code ${code}`));
        return;
      }
      try {
        resolve(JSON.parse(stdout || "{}"));
      } catch (error) {
        reject(error);
      }
    });

    child.stdin.write(JSON.stringify({
      action,
      ...payload
    }));
    child.stdin.end();
  });
}

async function serveFile(res, filePath) {
  try {
    const data = await readFile(filePath);
    const ext = path.extname(filePath).toLowerCase();
    const headers = {
      "Content-Type": contentTypes[ext] || "application/octet-stream"
    };

    if (ext === ".html") {
      headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0";
    } else if (ext === ".js" || ext === ".css") {
      headers["Cache-Control"] = "public, max-age=31536000, immutable";
    }

    res.writeHead(200, headers);
    res.end(data);
  } catch {
    res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Not found");
  }
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url || "/", `http://${req.headers.host}`);

  if (url.pathname === "/health") {
    return sendJson(res, 200, { ok: true, service: "weaver-web" });
  }

  if (url.pathname === "/health/config") {
    return sendJson(res, 200, {
      ok: true,
      service: "weaver-web",
      revision: process.env.K_REVISION || "local",
      authorities: {
        excerptContent: "excerpt_database",
        graphicsLifecycle: "firestore",
        sourceIntake: "weaver_and_google_sheets",
        downstreamContent: "poetry_please"
      },
      firestore: {
        backend: process.env.WEAVER_LEDGER_BACKEND || "sqlite_local",
        projectId: process.env.WEAVER_FIRESTORE_PROJECT_ID || "",
        databaseId: process.env.WEAVER_FIRESTORE_DATABASE_ID || ""
      },
      graphicsQc: {
        activeReadBackend: cleanSheetWhitespace(process.env.WEAVER_GRAPHICS_QC_READ_BACKEND).toLowerCase() === "firestore"
          ? "firestore_queue_cards"
          : "sheets_firestore_composite",
        firestoreReadModel: "structured_query_paginated",
        cleanupReadBackend: cleanSheetWhitespace(process.env.WEAVER_GRAPHICS_CLEANUP_READ_BACKEND).toLowerCase() === "firestore"
          ? "firestore_queue_cards"
          : "sheets_drive_composite"
      },
      routes: {
        approvedExcerptExport: "POST /api/excerpts/approved/export",
        graphicsRequestUpsert: "POST /graphics-handoff/requests",
        graphicsQueue: "GET /graphics-handoff/queue",
        graphicsCompletion: "POST /api/pig/completed-graphics",
        graphicsQcDecision: "POST /api/save-graphics-qc",
        graphicsHandoffRetry: "POST /api/graphics/handoffs/retry"
      }
    });
  }

  if (url.pathname === "/api/bootstrap") {
    const reviewQueueIncludeTitles = await getReviewQueueIncludeTitles();
    const { releaseCatalogOptions, releaseCatalogByTitle } = await getReleaseCatalogMetadata();
    return sendJson(res, 200, {
      appName: "Weaver",
      appVersion,
      googleOAuthClientId: defaultGoogleOAuthClientId,
      graphicsExportEnabled: true,
      sheetReadFallbackEnabled: false,
      spreadsheetId,
      sourceSheetName,
      reviewQueueIncludeTitles,
      releaseCatalogOptions,
      releaseCatalogByTitle,
      reviewApiMode: "cloud-run-sheet-proxy",
      contributorInvites: CONTRIBUTOR_INVITES.map(invite => ({
        email: invite.email,
        role: invite.role,
        allowedBooks: Array.isArray(invite.allowedBooks) ? invite.allowedBooks.slice() : []
      })),
      sections: [
        { id: "gathering", label: "Excerpt gathering" },
        { id: "review", label: "Review queue" },
        { id: "corrections", label: "Needs correction" },
        { id: "graphics", label: "Graphics creation queue" }
      ]
    });
  }

  if (url.pathname === "/api/catalog/validate" && req.method === "POST") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const records = Array.isArray(parsed.records) ? parsed.records : [];
      const result = await runCatalogValidation(records);
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/intake/options" && req.method === "GET") {
    try {
      const result = await getExcerptGatheringOptionsFromSheets();
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/intake/catalog/poems" && req.method === "GET") {
    try {
      const result = await getExcerptGatheringPoemsForBook(url.searchParams.get("bookTitle") || "");
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/intake/video-playlist" && req.method === "GET") {
    try {
      const result = await loadVideoPlaylistFromFormUrl(url.searchParams.get("formUrl") || "");
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/intake/submit" && req.method === "POST") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      assertContributorIntakeAccess(parsed);
      const result = await appendExcerptGatheringRow(parsed);
      invalidateQueueSnapshots();
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/review/pending-records" && req.method === "GET") {
    try {
      const result = await getPendingRecordsFromSheets();
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/excerpts/approved" && req.method === "GET") {
    try {
      const result = await getApprovedExcerptsExportFromSheets({
        since: url.searchParams.get("since") || ""
      });
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/excerpts/approved/export" && req.method === "POST") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const result = await getApprovedExcerptsExportFromSheets({
        since: parsed.since || ""
      });
      if (cleanSheetWhitespace(parsed.sourceImportId)) {
        result.sourceImportId = cleanSheetWhitespace(parsed.sourceImportId);
      }
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/excerpts/handoffs" && req.method === "GET") {
    try {
      const statuses = (url.searchParams.get("status") || "")
        .split(",")
        .map(value => value.trim())
        .filter(Boolean);
      const recordIds = (url.searchParams.get("recordIds") || "")
        .split(",")
        .map(value => value.trim())
        .filter(Boolean);
      const result = await getExcerptHandoffRecords({ recordIds, statuses });
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/excerpts/handoffs/retry" && req.method === "POST") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const recordIds = Array.isArray(parsed.recordIds) ? parsed.recordIds : [];
      const statuses = Array.isArray(parsed.statuses) ? parsed.statuses : ["queued", "failed"];
      const result = await retryExcerptHandoffs({ recordIds, statuses });
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/excerpts/handoffs/backfill-approved" && req.method === "POST") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const since = cleanSheetWhitespace(parsed.since);
      const sourceRecordIds = Array.isArray(parsed.sourceRecordIds) ? parsed.sourceRecordIds : [];
      const result = await backfillApprovedExcerptHandoffs({ since, sourceRecordIds });
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/review/excerpts" && req.method === "GET") {
    try {
      const result = await getPendingExcerptsForBookFromSheets(url.searchParams.get("bookTitle") || "");
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/review/correction-books" && req.method === "GET") {
    try {
      const result = await getCorrectionBooksFromSheets();
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/review/corrections" && req.method === "GET") {
    try {
      const result = await getCorrectionsForBookFromSheets(url.searchParams.get("bookTitle") || "");
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/review/graphics-books" && req.method === "GET") {
    try {
      const result = await getGraphicsBooksFromSheets(url.searchParams.get("mode") || "");
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/review/graphics-records" && req.method === "GET") {
    try {
      const result = await getGraphicsRecordsForBookFromSheets(
        url.searchParams.get("bookTitle") || "",
        url.searchParams.get("mode") || ""
      );
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/pig/graphics-request-books" && req.method === "GET") {
    try {
      const filterMode = url.searchParams.get("filter") || "all";
      const books = await getPigGraphicsRequestBooks(filterMode);
      return sendJson(res, 200, {
        ok: true,
        version: `${appVersion}-service-account`,
        filter: cleanSheetWhitespace(filterMode).toLowerCase() || "all",
        books
      });
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/pig/graphics-requests" && req.method === "GET") {
    try {
      const filterMode = url.searchParams.get("filter") || "all";
      const requestedBookTitle = url.searchParams.get("bookTitle") || "";
      const requests = await getPigGraphicsRequestSummaries(filterMode, requestedBookTitle);
      return sendJson(res, 200, {
        ok: true,
        version: `${appVersion}-service-account`,
        filter: cleanSheetWhitespace(filterMode).toLowerCase() || "all",
        requests
      });
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/pig/graphics-request-records" && req.method === "GET") {
    try {
      const filterMode = url.searchParams.get("filter") || "all";
      const requestedBookTitle = url.searchParams.get("bookTitle") || "";
      let requests = await getPigGraphicsRequests(filterMode);
      if (cleanSheetWhitespace(requestedBookTitle)) {
        const requestedKey = normalizeBookKey(requestedBookTitle);
        requests = requests.filter(record => normalizeBookKey(record.bookTitle) === requestedKey);
      }
      return sendJson(res, 200, {
        ok: true,
        version: `${appVersion}-service-account`,
        filter: cleanSheetWhitespace(filterMode).toLowerCase() || "all",
        requests
      });
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/graphics-handoff/requests" && req.method === "POST") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const requests = Array.isArray(parsed.requests)
        ? parsed.requests
        : (parsed.graphicsRequestId || parsed.request ? [parsed.request || parsed] : []);
      const result = await syncWeaverRuntimeDb("upsert_handoff_requests", { requests });
      console.log("[graphics-handoff] upsert", {
        count: Number(result.count || 0),
        requestIds: (result.records || []).map(record => record.graphicsRequestId)
      });
      return sendJson(res, result.ok ? 200 : 400, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/graphics-handoff/queue" && req.method === "GET") {
    const requestId = getRequestCorrelationId(req);
    const startedAt = performance.now();
    const limit = parseInt(url.searchParams.get("limit") || "100", 10) || 100;
    const cursor = Math.max(0, parseInt(url.searchParams.get("cursor") || "0", 10) || 0);
      const filterMode = url.searchParams.get("filter") || "all";
      const normalizedFilterMode = cleanSheetWhitespace(filterMode).toLowerCase() || "all";
      try {
      const refreshSource = ["1", "true", "yes"].includes(cleanSheetWhitespace(url.searchParams.get("refresh")).toLowerCase());
      if (!refreshSource && normalizedFilterMode !== "coverage_needs") {
        const ledgerStartedAt = performance.now();
        const ledgerResult = await syncWeaverRuntimeDb("get_handoff_queue", {
          limit,
          cursor,
          filter: normalizedFilterMode
        });
        const dbElapsedMs = performance.now() - ledgerStartedAt;
        const records = Array.isArray(ledgerResult?.records) ? ledgerResult.records : [];
        const result = {
          ok: Boolean(ledgerResult?.ok),
          filter: normalizedFilterMode,
          source: "firestore",
          nextCursor: cleanSheetWhitespace(ledgerResult?.nextCursor),
          records
        };
        const serializeStartedAt = performance.now();
        const responseBody = JSON.stringify(result, null, 2);
        const serializeElapsedMs = performance.now() - serializeStartedAt;
        const totalElapsedMs = performance.now() - startedAt;
        console.log("[graphics-handoff/queue]", {
          requestId,
          url: req.url,
          filter: normalizedFilterMode,
          limit,
          cursor,
          source: "firestore",
          statusCode: result.ok ? 200 : 400,
          recordCount: records.length,
          handlerElapsedMs: Number(totalElapsedMs.toFixed(1)),
          dbElapsedMs: Number(dbElapsedMs.toFixed(1)),
          serializationElapsedMs: Number(serializeElapsedMs.toFixed(1))
        });
        res.writeHead(result.ok ? 200 : 400, {
          "Content-Type": "application/json; charset=utf-8",
          "X-Request-Id": requestId,
          "Server-Timing": [
            "queue_build;dur=0.0",
            `db;dur=${dbElapsedMs.toFixed(1)}`,
            `serialize;dur=${serializeElapsedMs.toFixed(1)}`,
            `total;dur=${totalElapsedMs.toFixed(1)}`
          ].join(", "),
          "X-Weaver-Timings": `queue_build=0.0;db=${dbElapsedMs.toFixed(1)};serialize=${serializeElapsedMs.toFixed(1)};total=${totalElapsedMs.toFixed(1)}`
        });
        res.end(responseBody);
        return;
      }
        const queueBuildStartedAt = performance.now();
      const coverageNeeds = normalizedFilterMode === "coverage_needs"
        ? await getCoverageNeedsView(filterMode)
        : null;
      const actionableRecords = coverageNeeds
        ? coverageNeeds.records
        : (normalizedFilterMode === "rework"
          ? await getPigReworkRequests(filterMode)
          : await getPigGraphicsRequests(filterMode));
      const queueBuildElapsedMs = performance.now() - queueBuildStartedAt;
      const selectedRecords = actionableRecords.slice(0, limit);
      const queueLedgerRequests = selectedRecords
        .map(buildGraphicsHandoffLedgerRequest)
        .filter(request => cleanSheetWhitespace(request.graphicsRequestId));
      const ledgerSyncStartedAt = performance.now();
      let records = [];
      if (coverageNeeds) {
        records = selectedRecords.map(record => ({
          ...record,
          handoffLedger: true
        }));
      } else if (normalizedFilterMode === "rework") {
        records = selectedRecords.map(record => ({
          ...record,
          handoffLedger: true
        }));
      } else if (queueLedgerRequests.length) {
        await syncWeaverRuntimeDb("upsert_handoff_requests", { requests: queueLedgerRequests }).catch(() => null);
        const ledgerResult = await syncWeaverRuntimeDb("get_handoff_requests", {
          graphicsRequestIds: queueLedgerRequests.map(request => request.graphicsRequestId)
        }).catch(() => ({ ok: false, records: [] }));
        const ledgerByRequestId = Object.fromEntries(
          (Array.isArray(ledgerResult?.records) ? ledgerResult.records : [])
            .map(record => [cleanSheetWhitespace(record.graphicsRequestId), record])
            .filter(([graphicsRequestId]) => graphicsRequestId)
        );
        records = queueLedgerRequests.map(request => {
          const existing = ledgerByRequestId[request.graphicsRequestId];
          if (existing) {
            return {
              ...existing,
              handoffLedger: true
            };
          }
          return {
            graphicsRequestId: request.graphicsRequestId,
            sourceSystem: request.sourceSystem,
            sourceStatus: request.sourceStatus,
            sourcePayload: request.sourcePayload,
            handoffStatus: "requested",
            pigStatus: "not_started",
            qcStatus: "not_sent",
            handoffLedger: true
          };
        });
      }
      const dbElapsedMs = performance.now() - ledgerSyncStartedAt;
      const result = {
        ok: true,
        filter: normalizedFilterMode,
        poetryPleaseCoverage: coverageNeeds?.poetryPleaseCoverage || null,
        records
      };
      const serializeStartedAt = performance.now();
      const responseBody = JSON.stringify(result, null, 2);
      const serializeElapsedMs = performance.now() - serializeStartedAt;
      const totalElapsedMs = performance.now() - startedAt;
      const recordCount = Array.isArray(result?.records) ? result.records.length : 0;
      console.log("[graphics-handoff/queue]", {
        requestId,
        url: req.url,
        filter: normalizedFilterMode,
        limit,
        statusCode: result.ok ? 200 : 400,
        recordCount,
        handlerElapsedMs: Number(totalElapsedMs.toFixed(1)),
        queueBuildElapsedMs: Number(queueBuildElapsedMs.toFixed(1)),
        dbElapsedMs: Number(dbElapsedMs.toFixed(1)),
        serializationElapsedMs: Number(serializeElapsedMs.toFixed(1))
      });
      res.writeHead(result.ok ? 200 : 400, {
        "Content-Type": "application/json; charset=utf-8",
        "X-Request-Id": requestId,
        "Server-Timing": [
          `queue_build;dur=${queueBuildElapsedMs.toFixed(1)}`,
          `db;dur=${dbElapsedMs.toFixed(1)}`,
          `serialize;dur=${serializeElapsedMs.toFixed(1)}`,
          `total;dur=${totalElapsedMs.toFixed(1)}`
        ].join(", "),
        "X-Weaver-Timings": `queue_build=${queueBuildElapsedMs.toFixed(1)};db=${dbElapsedMs.toFixed(1)};serialize=${serializeElapsedMs.toFixed(1)};total=${totalElapsedMs.toFixed(1)}`
      });
      res.end(responseBody);
      return;
    } catch (error) {
      const totalElapsedMs = performance.now() - startedAt;
      console.error("[graphics-handoff/queue]", {
        requestId,
        url: req.url,
        filter: normalizedFilterMode,
        limit,
        handlerElapsedMs: Number(totalElapsedMs.toFixed(1)),
        error: error.message
      });
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/graphics-handoff/books" && req.method === "GET") {
    const requestId = getRequestCorrelationId(req);
    const startedAt = performance.now();
      const filterMode = url.searchParams.get("filter") || "all";
      const normalizedFilterMode = cleanSheetWhitespace(filterMode).toLowerCase() || "all";
      try {
      const refreshSource = ["1", "true", "yes"].includes(cleanSheetWhitespace(url.searchParams.get("refresh")).toLowerCase());
      if (!refreshSource && normalizedFilterMode !== "coverage_needs") {
        const ledgerStartedAt = performance.now();
        const ledgerResult = await syncWeaverRuntimeDb("get_handoff_queue", {
          limit: 500,
          cursor: 0,
          filter: normalizedFilterMode
        });
        const dbElapsedMs = performance.now() - ledgerStartedAt;
        const records = Array.isArray(ledgerResult?.records) ? ledgerResult.records : [];
        const books = summarizeGraphicsHandoffBooks(records);
        const result = {
          ok: Boolean(ledgerResult?.ok),
          filter: normalizedFilterMode,
          source: "firestore",
          poetryPleaseCoverage: null,
          books
        };
        const serializeStartedAt = performance.now();
        const responseBody = JSON.stringify(result, null, 2);
        const serializeElapsedMs = performance.now() - serializeStartedAt;
        const totalElapsedMs = performance.now() - startedAt;
        console.log("[graphics-handoff/books]", {
          requestId,
          url: req.url,
          filter: result.filter,
          source: "firestore",
          statusCode: result.ok ? 200 : 400,
          bookCount: books.length,
          handlerElapsedMs: Number(totalElapsedMs.toFixed(1)),
          dbElapsedMs: Number(dbElapsedMs.toFixed(1)),
          serializationElapsedMs: Number(serializeElapsedMs.toFixed(1))
        });
        res.writeHead(result.ok ? 200 : 400, {
          "Content-Type": "application/json; charset=utf-8",
          "X-Request-Id": requestId,
          "Server-Timing": [
            "queue_build;dur=0.0",
            `db;dur=${dbElapsedMs.toFixed(1)}`,
            `serialize;dur=${serializeElapsedMs.toFixed(1)}`,
            `total;dur=${totalElapsedMs.toFixed(1)}`
          ].join(", "),
          "X-Weaver-Timings": `queue_build=0.0;db=${dbElapsedMs.toFixed(1)};serialize=${serializeElapsedMs.toFixed(1)};total=${totalElapsedMs.toFixed(1)}`
        });
        res.end(responseBody);
        return;
      }
        const queueBuildStartedAt = performance.now();
      const coverageNeeds = normalizedFilterMode === "coverage_needs"
        ? await getCoverageNeedsView(filterMode)
        : null;
      const actionableRecords = coverageNeeds
        ? coverageNeeds.records
        : (normalizedFilterMode === "rework"
          ? await getPigReworkRequests(filterMode)
          : await getPigGraphicsRequests(filterMode));
      const queueBuildElapsedMs = performance.now() - queueBuildStartedAt;
      const dbStartedAt = performance.now();
      const books = coverageNeeds
        ? coverageNeeds.books
        : summarizeGraphicsHandoffBooks(actionableRecords);
      const dbElapsedMs = performance.now() - dbStartedAt;
      const result = {
        ok: true,
        filter: normalizedFilterMode,
        poetryPleaseCoverage: coverageNeeds?.poetryPleaseCoverage || null,
        books
      };
      const serializeStartedAt = performance.now();
      const responseBody = JSON.stringify(result, null, 2);
      const serializeElapsedMs = performance.now() - serializeStartedAt;
      const totalElapsedMs = performance.now() - startedAt;
      console.log("[graphics-handoff/books]", {
        requestId,
        url: req.url,
        filter: result.filter,
        statusCode: 200,
        bookCount: books.length,
        handlerElapsedMs: Number(totalElapsedMs.toFixed(1)),
        queueBuildElapsedMs: Number(queueBuildElapsedMs.toFixed(1)),
        dbElapsedMs: Number(dbElapsedMs.toFixed(1)),
        serializationElapsedMs: Number(serializeElapsedMs.toFixed(1))
      });
      res.writeHead(200, {
        "Content-Type": "application/json; charset=utf-8",
        "X-Request-Id": requestId,
        "Server-Timing": [
          `queue_build;dur=${queueBuildElapsedMs.toFixed(1)}`,
          `db;dur=${dbElapsedMs.toFixed(1)}`,
          `serialize;dur=${serializeElapsedMs.toFixed(1)}`,
          `total;dur=${totalElapsedMs.toFixed(1)}`
        ].join(", "),
        "X-Weaver-Timings": `queue_build=${queueBuildElapsedMs.toFixed(1)};db=${dbElapsedMs.toFixed(1)};serialize=${serializeElapsedMs.toFixed(1)};total=${totalElapsedMs.toFixed(1)}`
      });
      res.end(responseBody);
      return;
    } catch (error) {
      const totalElapsedMs = performance.now() - startedAt;
      console.error("[graphics-handoff/books]", {
        requestId,
        url: req.url,
        filter: cleanSheetWhitespace(filterMode).toLowerCase() || "all",
        handlerElapsedMs: Number(totalElapsedMs.toFixed(1)),
        error: error.message
      });
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/graphics-handoff/status" && req.method === "GET") {
    try {
      const result = await syncWeaverRuntimeDb("get_handoff_queue", { limit: 25 });
      return sendJson(res, result.ok ? 200 : 400, {
        ok: result.ok,
        version: `${appVersion}-service-account`,
        openCount: Array.isArray(result.records) ? result.records.length : 0,
        sample: result.records || []
      });
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  const handoffMatch = url.pathname.match(/^\/graphics-handoff\/([^/]+)(?:\/(claim))?$/);
  if (handoffMatch && req.method === "GET" && !handoffMatch[2]) {
    try {
      const result = await syncWeaverRuntimeDb("get_handoff_request", {
        graphicsRequestId: decodeURIComponent(handoffMatch[1])
      });
      return sendJson(res, result.ok ? 200 : 404, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (handoffMatch && req.method === "POST" && handoffMatch[2] === "claim") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const graphicsRequestId = decodeURIComponent(handoffMatch[1]);
      const result = await syncWeaverRuntimeDb("claim_handoff_request", {
        graphicsRequestId,
        claimedBy: parsed.claimedBy || parsed.worker || "P.I.G."
      });
      console.log("[graphics-handoff] claim", { graphicsRequestId, ok: result.ok });
      return sendJson(res, result.ok ? 200 : 400, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (handoffMatch && req.method === "PATCH" && !handoffMatch[2]) {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const graphicsRequestId = decodeURIComponent(handoffMatch[1]);
      const result = await syncWeaverRuntimeDb("patch_handoff_request", {
        graphicsRequestId,
        update: parsed.update && typeof parsed.update === "object" ? parsed.update : parsed
      });
      console.log("[graphics-handoff] update", {
        graphicsRequestId,
        handoffStatus: result.record?.handoffStatus,
        pigStatus: result.record?.pigStatus,
        qcStatus: result.record?.qcStatus
      });
      return sendJson(res, result.ok ? 200 : 400, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/poetry-please/qc-approved-graphics/books" && req.method === "GET") {
    try {
      const books = await getPoetryPleaseApprovedGraphicBooks();
      return sendJson(res, 200, {
        ok: true,
        version: `${appVersion}-service-account`,
        books
      });
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/poetry-please/qc-approved-graphics" && req.method === "GET") {
    try {
      const bookTitle = url.searchParams.get("bookTitle") || "";
      const records = await getPoetryPleaseApprovedGraphics(bookTitle);
      return sendJson(res, 200, {
        ok: true,
        version: `${appVersion}-service-account`,
        records
      });
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/graphics/rework-request" && req.method === "POST") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const records = Array.isArray(parsed.records) ? parsed.records : [];
      const note = String(parsed.note || "");
      const result = await createManualGraphicsReworkRequests(records, note);
      return sendJson(res, result.ok ? 200 : 400, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/graphics/handoffs/retry" && req.method === "POST") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const records = Array.isArray(parsed?.records) ? parsed.records : [];
      const result = await retryFailedGraphicsHandoffs(records);
      return sendJson(res, 200, {
        ok: true,
        version: `${appVersion}-service-account`,
        ...result
      });
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/save-reviews" && req.method === "POST") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const updates = Array.isArray(parsed.updates) ? parsed.updates : [];

      const result = await saveReviewsToSheets(updates);
      if (result?.ok) invalidateQueueSnapshots();
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/save-review-single" && req.method === "POST") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const update = parsed.update && typeof parsed.update === "object" ? parsed.update : {};

      const result = await saveSingleReviewToSheets(update);
      if (result?.ok) invalidateQueueSnapshots();
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/save-graphics-qc" && req.method === "POST") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const updates = Array.isArray(parsed.updates) ? parsed.updates : [];

      const result = await saveGraphicsQcToSheets(updates);
      if (result?.ok) invalidateQueueSnapshots();
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/pig/completed-graphics" && req.method === "POST") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const completions = Array.isArray(parsed.completions)
        ? parsed.completions
        : (parsed.completion && typeof parsed.completion === "object" ? [parsed.completion] : []);

      const result = await upsertPigCompletedGraphics(completions);
      if (result?.ok) invalidateQueueSnapshots();
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/graphics/links" && req.method === "POST") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const records = Array.isArray(parsed.records) ? parsed.records : [];
      const result = await runGraphicsAssetLookup(records);
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/drive-image" && req.method === "GET") {
    try {
      const fileId = cleanSheetWhitespace(url.searchParams.get("fileId"));
      if (!fileId) {
        return sendJson(res, 400, {
          ok: false,
          error: "Missing fileId"
        });
      }

      let response = await fetchDriveFileResponse(fileId);
      if (!response.ok) {
        response = await fetchDriveThumbnailResponse(fileId);
      }

      const body = await response.arrayBuffer();
      if (!response.ok) {
        const text = Buffer.from(body).toString("utf8");
        return sendJson(res, response.status, {
          ok: false,
          error: `Drive image request failed: ${response.status} ${text.slice(0, 240)}`
        });
      }

      res.writeHead(200, {
        "Content-Type": response.headers.get("content-type") || "application/octet-stream",
        "Cache-Control": "public, max-age=3600"
      });
      res.end(Buffer.from(body));
      return;
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/graphics/folder-import/preview" && req.method === "POST") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const folderUrl = cleanSheetWhitespace(parsed.folderUrl);
      const result = await previewDriveFolderImport(folderUrl);
      return sendJson(res, result.ok ? 200 : 400, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/api/graphics/folder-import/apply" && req.method === "POST") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const folderUrl = cleanSheetWhitespace(parsed.folderUrl);
      const result = await applyDriveFolderImport(folderUrl);
      return sendJson(res, result.ok ? 200 : 400, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/catalog-poem" && req.method === "GET") {
    try {
      const bookTitle = url.searchParams.get("bookTitle") || "";
      const poemTitle = url.searchParams.get("poemTitle") || "";
      const excerptText = url.searchParams.get("excerptText") || "";
      const result = await runCatalogPoemLookup(bookTitle, poemTitle, excerptText);
      const html = result.ok
        ? `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>${escapeHtml(result.poemTitle)} · Weaver Catalog Context</title>
  <style>
    body { margin: 0; font-family: Georgia, "Times New Roman", serif; background: #f7f2ea; color: #1d1a17; }
    main { max-width: 820px; margin: 0 auto; padding: 32px 24px 56px; }
    .meta { color: #6b6259; font-size: 14px; text-transform: uppercase; letter-spacing: .08em; margin-bottom: 12px; }
    h1 { margin: 0 0 8px; font-size: clamp(30px, 5vw, 46px); line-height: 1.05; }
    h2 { margin: 0 0 22px; font-size: 22px; color: #b84f2d; font-weight: 600; }
    pre { white-space: pre-wrap; word-break: break-word; background: rgba(255,252,247,.92); border: 1px solid rgba(29,26,23,.1); border-radius: 24px; padding: 24px; font: 18px/1.7 Georgia, "Times New Roman", serif; box-shadow: 0 18px 40px rgba(62,39,27,.08); }
  </style>
</head>
<body>
  <main>
    <div class="meta">Catalog Poem Context · ${escapeHtml(result.wordCount || "")} words</div>
    <h1>${escapeHtml(result.poemTitle)}</h1>
    <h2>${escapeHtml(result.author)} · ${escapeHtml(result.bookTitle)}</h2>
    <pre>${escapeHtml(result.text)}</pre>
  </main>
</body>
</html>`
        : `<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8" /><meta name="viewport" content="width=device-width, initial-scale=1.0" /><title>Weaver Catalog Context</title></head>
<body style="font-family: Georgia, 'Times New Roman', serif; background:#f7f2ea; color:#1d1a17; padding:32px;">
  <h1 style="margin-top:0;">Catalog poem unavailable</h1>
  <p>${escapeHtml(result.error || "Unable to load poem text.")}</p>
</body>
</html>`;
      res.writeHead(result.ok ? 200 : 404, { "Content-Type": "text/html; charset=utf-8" });
      res.end(html);
      return;
    } catch (error) {
      res.writeHead(500, { "Content-Type": "text/html; charset=utf-8" });
      res.end(`<html><body style="font-family: sans-serif; padding: 32px;"><h1>Catalog poem lookup failed</h1><p>${escapeHtml(error.message)}</p></body></html>`);
      return;
    }
  }

  if (url.pathname === "/library-excerpt" && req.method === "GET") {
    try {
      const sourceRow = Number(url.searchParams.get("sourceRow") || 0);
      const result = await runLibraryExcerptLookup(sourceRow);
      const libraryStatusLabel = getLibraryStatusLabel(result.libraryStatus);
      const html = result.ok
        ? `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Excerpt Library Match · ${escapeHtml(result.poemTitle || "Weaver")}</title>
  <style>
    body { margin: 0; font-family: Georgia, "Times New Roman", serif; background: #f7f2ea; color: #1d1a17; }
    main { max-width: 820px; margin: 0 auto; padding: 32px 24px 56px; }
    .meta { color: #6b6259; font-size: 14px; text-transform: uppercase; letter-spacing: .08em; margin-bottom: 12px; }
    .status { margin: 0 0 16px; color: #6b6259; font-size: 16px; }
    h1 { margin: 0 0 8px; font-size: clamp(30px, 5vw, 46px); line-height: 1.05; }
    h2 { margin: 0 0 22px; font-size: 22px; color: #b84f2d; font-weight: 600; }
    pre { white-space: pre-wrap; word-break: break-word; background: rgba(255,252,247,.92); border: 1px solid rgba(29,26,23,.1); border-radius: 24px; padding: 24px; font: 18px/1.7 Georgia, "Times New Roman", serif; box-shadow: 0 18px 40px rgba(62,39,27,.08); }
  </style>
</head>
<body>
  <main>
    <div class="meta">Excerpt Library Match · Row ${escapeHtml(result.sourceRow)}${result.wordCount ? ` · ${escapeHtml(result.wordCount)} words` : ""}</div>
    <h1>${escapeHtml(result.poemTitle || "Untitled")}</h1>
    <h2>${escapeHtml(result.author || "Unknown author")}${result.bookTitle ? ` · ${escapeHtml(result.bookTitle)}` : ""}</h2>
    ${libraryStatusLabel ? `<p class="status">${escapeHtml(libraryStatusLabel)}</p>` : ""}
    <pre>${escapeHtml(result.text)}</pre>
  </main>
</body>
</html>`
        : `<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8" /><meta name="viewport" content="width=device-width, initial-scale=1.0" /><title>Excerpt Library Match</title></head>
<body style="font-family: Georgia, 'Times New Roman', serif; background:#f7f2ea; color:#1d1a17; padding:32px;">
  <h1 style="margin-top:0;">Library excerpt unavailable</h1>
  <p>${escapeHtml(result.error || "Unable to load excerpt text.")}</p>
</body>
</html>`;
      res.writeHead(result.ok ? 200 : 404, { "Content-Type": "text/html; charset=utf-8" });
      res.end(html);
      return;
    } catch (error) {
      res.writeHead(500, { "Content-Type": "text/html; charset=utf-8" });
      res.end(`<html><body style="font-family: sans-serif; padding: 32px;"><h1>Library excerpt lookup failed</h1><p>${escapeHtml(error.message)}</p></body></html>`);
      return;
    }
  }

  const requestedPath = url.pathname === "/" ? "/index.html" : url.pathname;
  const safePath = path.normalize(requestedPath).replace(/^(\.\.[/\\])+/, "");
  const filePath = path.join(publicDir, safePath);

  if (!filePath.startsWith(publicDir)) {
    res.writeHead(403, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Forbidden");
    return;
  }

  return serveFile(res, filePath);
});

if (process.env.K_SERVICE) {
  const requiredFirestoreConfig = {
    WEAVER_LEDGER_BACKEND: "firestore",
    WEAVER_FIRESTORE_PROJECT_ID: "button-weaver-internal",
    WEAVER_FIRESTORE_DATABASE_ID: "weaverledger"
  };
  const invalidFirestoreConfig = Object.entries(requiredFirestoreConfig)
    .filter(([name, expected]) => String(process.env[name] || "").trim() !== expected)
    .map(([name]) => name);
  if (invalidFirestoreConfig.length) {
    throw new Error(`Invalid production Firestore configuration: ${invalidFirestoreConfig.join(", ")}`);
  }
}

server.listen(port, () => {
  console.log(`Weaver server running on port ${port}`);
});
