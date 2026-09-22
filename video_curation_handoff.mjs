function clean(value) {
  return String(value ?? "").trim();
}

function driveFileId(value) {
  const url = clean(value);
  return url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/)?.[1]
    || url.match(/[?&]id=([a-zA-Z0-9_-]+)/)?.[1]
    || "";
}

export function buildWeaverVideoImport(candidate, gate) {
  if (clean(gate?.decision) !== "ready_for_poetry_please") {
    throw new Error("Only a ready-for-Poetry Please video can be handed off.");
  }
  if (candidate?.publicationRestricted || gate?.publicationRestricted || clean(candidate?.releaseStatus)) {
    throw new Error("A publication-restricted source cannot be handed off to Poetry Please.");
  }
  const sourceFileId = clean(candidate?.sourceFileId);
  const sourceRecordId = `weaver:video:${sourceFileId}`;
  const sourceEvent = clean(gate?.sourceEvent) || clean(candidate?.sourceEvent);
  const sourceEventLabel = clean(gate?.sourceEventLabel) || clean(candidate?.sourceEventLabel);
  const finalAssetUrl = clean(gate?.publishableAssetUrl);
  const sourceVideoUrl = clean(gate?.sourceVideoUrl) || clean(candidate?.sourceVideoUrl);
  const missing = Object.entries({ sourceFileId, sourceEvent, sourceEventLabel, finalAssetUrl })
    .filter(([, value]) => !value)
    .map(([key]) => key);
  if (missing.length) {
    throw new Error(`Video handoff is missing required metadata: ${missing.join(", ")}.`);
  }
  let finalUrl;
  try {
    finalUrl = new URL(finalAssetUrl);
  } catch {
    throw new Error("The final publishable asset URL is invalid.");
  }
  if (finalUrl.protocol !== "https:") {
    throw new Error("The final publishable asset URL must use HTTPS.");
  }
  if (finalAssetUrl === sourceVideoUrl || driveFileId(finalAssetUrl) === sourceFileId) {
    throw new Error("The final publishable asset cannot be the raw Weaver source video.");
  }

  return {
    contentType: "VV",
    sourceSystem: "weaver",
    sourceRecordId,
    candidateId: clean(candidate.candidateId),
    prioritySetId: clean(candidate.prioritySetId),
    sourceFileId,
    sourceDriveFileId: sourceFileId,
    sourceVideoUrl,
    finalAssetUrl,
    author: clean(gate.author) || clean(candidate.author),
    title: clean(gate.poemTitle) || clean(candidate.poemTitle),
    book: clean(gate.bookTitle) || clean(candidate.bookTitle),
    releaseCatalog: clean(gate.releaseCatalog) || clean(candidate.releaseCatalog),
    eventReleaseCatalog: clean(gate.eventReleaseCatalog) || clean(candidate.eventReleaseCatalog),
    sourceEvent,
    sourceEventLabel,
    gateId: clean(gate.gateId),
    releaseStatus: clean(gate.releaseStatus) || clean(candidate.releaseStatus),
    publicationRestricted: false,
    selectedExcerptRecordIds: Array.isArray(gate.selectedExcerptRecordIds) ? gate.selectedExcerptRecordIds : [],
    baseScore: gate.baseScore ?? candidate.baseScore,
    excerptBonus: gate.excerptBonus ?? candidate.excerptBonus,
    candidateScore: gate.candidateScore ?? candidate.candidateScore,
    reviewCount: Number(candidate.reviewCount || 0),
    excerptCount: Array.isArray(candidate.excerptRecordIds) ? candidate.excerptRecordIds.length : 0,
    decision: "ready_for_poetry_please",
    decidedBy: clean(gate.decidedBy),
    decidedAt: clean(gate.decidedAt)
  };
}

export function parsePoetryPleaseVideoImport(response, body, sourceRecordId, apiUrl) {
  const item = (Array.isArray(body?.results) ? body.results : [])
    .find(result => clean(result?.sourceRecordId) === sourceRecordId) || {};
  const status = clean(item.status).toLowerCase();
  const canonicalVideoId = clean(item.canonicalVideoId);
  const canonicalVideoUrl = clean(item.canonicalVideoUrl);
  const ok = response.ok && body?.ok === true
    && ["created", "updated", "duplicate"].includes(status)
    && Boolean(canonicalVideoId && canonicalVideoUrl);
  return {
    ok,
    status: ok ? "sent_to_poetry_please" : (status || "failed"),
    canonicalVideoId,
    canonicalVideoUrl: canonicalVideoUrl ? new URL(canonicalVideoUrl, apiUrl).toString() : "",
    finalAssetUrl: clean(item.finalAssetUrl),
    error: ok ? "" : clean(item.error || body?.error || `Poetry Please returned ${response.status}`)
  };
}
