import { createHash } from "node:crypto";

function clean(value) {
  return String(value ?? "").trim();
}

export function driveFileId(value) {
  const url = clean(value);
  return url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/)?.[1]
    || url.match(/[?&]id=([a-zA-Z0-9_-]+)/)?.[1]
    || "";
}

export function classifyDriveFolderPath(folders) {
  const matches = folders.flatMap(folder => {
    const name = clean(folder.name);
    const id = clean(folder.id);
    return [
      /(?:^|[\s([/-])(?:FV|Finished Video)(?:$|[\s)\]/:-])/i.test(name) ? { id, type: "FV" } : null,
      /(?:^|[\s([/-])(?:OM|Original Media)(?:$|[\s)\]/:-])/i.test(name) ? { id, type: "OM" } : null
    ].filter(Boolean);
  });
  const types = new Set(matches.map(folder => folder.type));
  return {
    sourceMediaType: types.size === 1 ? [...types][0] : "unknown",
    sourceMediaTypeFolderId: types.size === 1 ? matches[0].id : ""
  };
}

export function resolveCurrentVideoFileId(record, files) {
  const sourceFileId = clean(record?.sourceFileId);
  if (files.some(file => clean(file.id) === sourceFileId)) return sourceFileId;
  const sourceFileName = clean(record?.sourceFileName);
  if (!sourceFileName) return "";
  const matches = files.filter(file => clean(file.name) === sourceFileName);
  if (matches.length === 1) return clean(matches[0].id);
  if (matches.length > 1 || !/\(Cam Audio\)/i.test(sourceFileName)) return "";
  const withoutCamAudio = name => clean(name).replace(/\s*\(Cam Audio\)/ig, "");
  const movedLabelMatches = files.filter(file => (
    /\(Cam Audio\)/i.test(clean(file.name))
    && withoutCamAudio(file.name) === withoutCamAudio(sourceFileName)
  ));
  return movedLabelMatches.length === 1 ? clean(movedLabelMatches[0].id) : "";
}

export function reconcileVideoReviews(files, records) {
  const byReviewerAndFile = new Map();
  for (const record of records) {
    const currentFileId = resolveCurrentVideoFileId(record, files);
    const reviewerEmail = clean(record.reviewerEmail).toLowerCase();
    if (!currentFileId || !reviewerEmail) continue;
    const key = `${currentFileId}|${reviewerEmail}`;
    const previous = byReviewerAndFile.get(key);
    const isDirect = clean(record.sourceFileId) === currentFileId;
    const previousIsDirect = previous?.originalSourceFileIds?.includes(currentFileId);
    const preferCurrent = !previous || (isDirect && !previousIsDirect)
      || (isDirect === previousIsDirect && clean(record.updatedAt) > clean(previous.updatedAt));
    const selected = preferCurrent ? record : previous;
    byReviewerAndFile.set(key, {
      ...selected,
      sourceFileId: currentFileId,
      originalSourceFileIds: [...new Set([
        ...(previous?.originalSourceFileIds || []),
        clean(record.sourceFileId)
      ].filter(Boolean))],
      excerptRecordIds: [...new Set([
        ...(previous?.excerptRecordIds || []),
        ...(record.excerptRecordIds || [])
      ])]
    });
  }
  return [...byReviewerAndFile.values()];
}

// Retain stored source IDs separately from the dashboard's filename-based
// aggregation. A count is not proof of canonical recording identity.
export function buildVideoReviewEvidence(files, records, currentFileId) {
  return records.filter(record => clean(record.rating)
    && resolveCurrentVideoFileId(record, files) === clean(currentFileId))
    .map(record => ({
      reviewId: clean(record.reviewId), sourceFileId: clean(record.sourceFileId),
      sourceFileName: clean(record.sourceFileName),
      mapping: clean(record.sourceFileId) === clean(currentFileId) ? "direct_file_id" : "filename_reconciliation",
      reviewerEmail: clean(record.reviewerEmail), rating: clean(record.rating),
      excerptRecordIds: [...new Set((record.excerptRecordIds || []).map(clean).filter(Boolean))],
      updatedAt: clean(record.updatedAt)
    }));
}

export function buildReviewerProgressExport(reviewerEmail, setResults) {
  const email = clean(reviewerEmail).toLowerCase();
  const sets = setResults.map(({ set, files, reviews }) => {
    const matchingReviews = reconcileVideoReviews(files, reviews.filter(review => (
      clean(review.prioritySetId) === set.id && clean(review.reviewerEmail).toLowerCase() === email
    )));
    const reviewsByFileId = new Map(matchingReviews.map(review => [clean(review.sourceFileId), review]));
    const videos = files.map(file => {
      const fileId = clean(file.id);
      const review = reviewsByFileId.get(fileId);
      return {
        sourceFileId: fileId,
        sourceFileName: clean(file.name),
        videoUrl: clean(file.webViewLink) || `https://drive.google.com/file/d/${encodeURIComponent(fileId)}/view`,
        reviewStatus: clean(review?.rating) ? "reviewed" : "not_reviewed",
        rating: clean(review?.rating) || null,
        excerptCount: new Set((review?.excerptRecordIds || []).map(clean).filter(Boolean)).size,
        reviewedAt: clean(review?.updatedAt) || null
      };
    });
    const reviewedCount = videos.filter(video => video.reviewStatus === "reviewed").length;
    return {
      prioritySetId: set.id,
      label: set.label,
      eventName: set.eventName || set.label,
      totalVideos: videos.length,
      reviewedCount,
      remainingCount: videos.length - reviewedCount,
      videos
    };
  });
  return {
    ok: true,
    generatedAt: new Date().toISOString(),
    reviewerEmail: email,
    totals: {
      prioritySets: sets.length,
      videos: sets.reduce((total, set) => total + set.totalVideos, 0),
      reviewed: sets.reduce((total, set) => total + set.reviewedCount, 0),
      remaining: sets.reduce((total, set) => total + set.remainingCount, 0)
    },
    sets
  };
}

export function poetryPleaseVideoIdForSourceFile(sourceFileId, gates = []) {
  const fileId = clean(sourceFileId);
  if (!fileId) return "";
  const ids = new Set(gates.filter(gate => (
    clean(gate?.sourceReviewFileId) === fileId || clean(gate?.sourceFileId) === fileId
  )).map(gate => clean(gate?.poetryPleaseHandoff?.canonicalVideoId)).filter(Boolean));
  return ids.size === 1 ? [...ids][0] : "";
}

export function assertExistingCanonicalReceipt(receipt, sourceRecordId, canonicalVideoId, sourceEvent) {
  if (receipt?.ok !== true || receipt.received !== true || receipt.ambiguousVideoMatch
    || clean(receipt.sourceRecordId) !== clean(sourceRecordId)
    || clean(receipt.canonicalVideoId) !== clean(canonicalVideoId)
    || clean(receipt.event?.sourceEvent) !== clean(sourceEvent)) {
    throw new Error("The existing Poetry Please video identity could not be verified; no update was sent.");
  }
}

export function assertVideoImportPreflight(receipt, sourceRecordId) {
  if (receipt?.ok !== true || clean(receipt.sourceRecordId) !== sourceRecordId
    || receipt.received !== false || receipt.ambiguousVideoMatch
    || clean(receipt.canonicalVideoId) || receipt.ledger?.count !== 0) {
    throw new Error("Poetry Please already has this source or its receipt needs reconciliation; no video was sent.");
  }
}

export function assertVideoRetryPreflight(receipt, record, gate) {
  const prior = gate?.poetryPleaseHandoff || {};
  const latest = receipt?.ledger?.latest;
  const canonical = `WEAVER-VV-${createHash("sha256").update(record.sourceRecordId).digest("hex").slice(0, 24).toUpperCase()}`;
  if (prior.status !== "failed" || prior.ok === true || clean(prior.canonicalVideoId)
    || clean(prior.sourceRecordId) !== record.sourceRecordId
    || clean(gate.videoRecordId) !== record.sourceRecordId
    || receipt?.ok !== true || receipt.received !== false || receipt.ambiguousVideoMatch
    || clean(receipt.sourceRecordId) !== record.sourceRecordId || clean(receipt.canonicalVideoId)
    || receipt.expectedCanonicalVideoId !== canonical || receipt.ledger?.count !== 1
    || latest?.status !== "failed" || !clean(latest.importLedgerId) || clean(latest.canonicalVideoId)
    || !/(?:timeout|timed?\s*out|deadline)/i.test(clean(latest.error))
    || latest.receivedReviewCount !== 0 || latest.receivedSelectedExcerptIdCount !== 0
    || receipt.review?.reviewCount !== 0 || receipt.review?.selectedExcerptIdCount !== 0) {
    throw new Error("This failed send is not a verified empty timeout attempt; no retry was sent.");
  }
  assertVideoClaimMatchesRecord(gate, record);
  return { importLedgerId: latest.importLedgerId, sourceRecordId: record.sourceRecordId,
    expectedHandoff: prior, receiptCheckedAtMs: Date.now() };
}

export function buildVideoContextUpdate(candidate, receipt) {
  const sourceRecordId = clean(candidate.videoRecordId);
  const canonicalVideoId = clean(candidate.existingCanonicalVideoId);
  if (!canonicalVideoId || !candidate.sourceReviewFileId || !candidate.gate
    || candidate.gate.decision !== "ready_for_poetry_please"
    || candidate.publicationRestricted) throw new Error("An existing, ready canonical video is required.");
  assertExistingCanonicalReceipt(receipt, sourceRecordId, canonicalVideoId, candidate.sourceEvent);
  const evidence = candidate.reviewEvidence || [];
  const reviews = (candidate.ratings || []).flatMap(review => {
    const origins = evidence.filter(row => row.reviewId === review.reviewId);
    if (origins.length !== 1 || !clean(review.reviewId)) return [];
    const origin = origins[0];
    // Only the exact canonical review source or the current FV, not a third
    // source included solely through filename reconciliation.
    if (![clean(candidate.sourceFileId), clean(candidate.sourceReviewFileId)].includes(clean(origin.sourceFileId))) return [];
    return [{ ...review, sourceRecordId: `weaver:video:${origin.sourceFileId}`,
      excerptRecordIds: [...new Set((origin.excerptRecordIds || []).map(clean).filter(Boolean))] }];
  });
  if (!reviews.length || new Set(reviews.map(row => row.reviewId)).size !== reviews.length) {
    throw new Error("Exact stored review identities are required for context delivery.");
  }
  const selectedExcerptRecordIds = [...new Set((candidate.gate.selectedExcerptRecordIds || []).map(clean).filter(Boolean))];
  return { canonicalVideoId, sourceRecordId, sourceEvent: clean(candidate.sourceEvent), reviews, selectedExcerptRecordIds };
}

export async function syncVideoReviewContext(candidate, { readReceipt, postContext }) {
  const before = await readReceipt(candidate.videoRecordId);
  const record = buildVideoContextUpdate(candidate, before);
  const context = await postContext(record);
  if (context?.ok !== true || context.canonicalVideoId !== record.canonicalVideoId
    || context.sourceRecordId !== record.sourceRecordId
    || context.receivedReviewCount !== record.reviews.length
    || context.receivedSelectedExcerptIdCount !== record.selectedExcerptRecordIds.length) {
    throw new Error("Poetry Please did not verify all submitted context; inspect its receipt before retrying.");
  }
  const after = await readReceipt(record.sourceRecordId);
  assertExistingCanonicalReceipt(after, record.sourceRecordId, record.canonicalVideoId, record.sourceEvent);
  for (const field of ["driveLink", "videoUrl", "storageMode", "reviewOnly"]) {
    if (JSON.stringify(before.media?.[field]) !== JSON.stringify(after.media?.[field])) {
      throw new Error("Primary media changed during context reconciliation; inspect this video.");
    }
  }
  if (after.review?.reviewCount !== context.storedReviewCount
    || after.review?.selectedExcerptIdCount !== context.storedSelectedExcerptIdCount) {
    throw new Error("The saved Poetry Please context receipt does not match its write result.");
  }
  return { ok: true, context, submittedReviewIds: record.reviews.map(row => row.reviewId),
    selectedExcerptRecordIds: record.selectedExcerptRecordIds,
    receiver: summarizeVideoReceiverState(candidate, after) };
}

export function summarizeVideoReceiverState(candidate, receipt, versions = null) {
  if (receipt?.ok !== true || receipt.ambiguousVideoMatch
    || clean(receipt.sourceRecordId) !== clean(candidate.videoRecordId)) {
    throw new Error("Poetry Please receipt does not uniquely match this reviewed source.");
  }
  if (receipt.received === false) {
    if (clean(receipt.canonicalVideoId)) throw new Error("Poetry Please absence receipt contains a canonical video.");
    return { received: false, sourceRecordId: candidate.videoRecordId, ledger: receipt.ledger || null };
  }
  const canonicalVideoId = clean(receipt.canonicalVideoId);
  if (receipt.received !== true || !canonicalVideoId
    || (candidate.existingCanonicalVideoId && canonicalVideoId !== candidate.existingCanonicalVideoId)
    || clean(receipt.event?.sourceEvent) !== clean(candidate.sourceEvent)) {
    throw new Error("Poetry Please canonical video or event does not match this candidate.");
  }
  if (versions && (versions.ok !== true || versions.schemaVersion !== "1" || !Array.isArray(versions.renditions)
    || clean(versions.canonicalVideoId) !== canonicalVideoId
    || clean(versions.sourceRecordId) !== clean(candidate.videoRecordId))) {
    throw new Error("Poetry Please versions belong to a different canonical video or source.");
  }
  return {
    received: true, sourceRecordId: candidate.videoRecordId, canonicalVideoId,
    event: receipt.event, media: receipt.media, review: receipt.review,
    currentFinalFileId: driveFileId(candidate.gate?.publishableAssetUrl || candidate.sourceVideoUrl),
    receiverFinalFileId: driveFileId(receipt.media?.driveLink),
    expectedReviewCount: candidate.reviewCount,
    expectedExcerptCount: Array.isArray(candidate.gate?.selectedExcerptRecordIds)
      ? new Set(candidate.gate.selectedExcerptRecordIds.map(clean).filter(Boolean)).size
      : candidate.excerptRecordIds?.length || 0,
    gatheredExcerptCount: candidate.excerptRecordIds?.length || 0,
    renditions: Array.isArray(versions?.renditions) ? versions.renditions : [],
  };
}

export function assertVideoImportReadback(receipt, record, result) {
  let driveHost = "";
  try { driveHost = new URL(receipt?.media?.driveLink).hostname; } catch {}
  if (receipt?.ok !== true || receipt.received !== true || receipt.ambiguousVideoMatch
    || clean(receipt.sourceRecordId) !== record.sourceRecordId
    || clean(receipt.canonicalVideoId) !== result.canonicalVideoId
    || clean(receipt.event?.sourceEvent) !== record.sourceEvent
    || clean(receipt.event?.sourceEventLabel) !== record.sourceEventLabel
    || driveHost !== "drive.google.com"
    || driveFileId(receipt.media?.driveLink) !== driveFileId(record.finalAssetUrl)
    || !driveFileId(record.finalAssetUrl)
    || receipt.media?.reviewOnly !== true
    || receipt.review?.reviewCount !== record.reviews.length
    || receipt.review?.selectedExcerptIdCount !== record.selectedExcerptRecordIds.length) {
    throw new Error("Poetry Please import response was not confirmed by its current receipt; reconcile this attempt before another send.");
  }
}

export function assertVideoClaimMatchesRecord(claimedGate, record) {
  const excerpts = [...new Set((claimedGate?.selectedExcerptRecordIds || []).map(clean).filter(Boolean))].sort();
  if (clean(claimedGate?.decision) !== "ready_for_poetry_please"
    || clean(claimedGate?.videoRecordId) !== record.sourceRecordId
    || clean(claimedGate?.sourceEvent) !== record.sourceEvent
    || clean(claimedGate?.sourceEventLabel) !== record.sourceEventLabel
    || clean(claimedGate?.publishableAssetUrl) !== record.finalAssetUrl
    || JSON.stringify(excerpts) !== JSON.stringify([...record.selectedExcerptRecordIds].sort())) {
    throw new Error("The saved video gate changed during send preparation; this attempt is held for reconciliation.");
  }
}

export function listKnownVideoAssets(currentFile, historicalSourceIds = [], gates = [], sourceReviewFileId = "") {
  const currentId = clean(currentFile?.id);
  const root = clean(sourceReviewFileId) || currentId;
  const assets = new Map();
  const add = (fileId, role, url = "") => {
    const id = clean(fileId);
    if (!id) return;
    const prior = assets.get(id) || { fileId: id, url: `https://drive.google.com/file/d/${encodeURIComponent(id)}/view`, roles: [] };
    if (url.startsWith("https://") && driveFileId(url) === id) prior.url = url;
    if (!prior.roles.includes(role)) prior.roles.push(role);
    assets.set(id, prior);
  };
  add(currentId, "current_candidate", clean(currentFile?.webViewLink));
  for (const id of historicalSourceIds) add(id, "review_history");
  for (const gate of gates) {
    if (clean(gate.sourceReviewFileId) !== root && clean(gate.sourceFileId) !== root) continue;
    add(gate.sourceFileId, "saved_gate");
    add(driveFileId(gate.publishableAssetUrl), "saved_final", clean(gate.publishableAssetUrl));
  }
  return [...assets.values()];
}

export function resolveVideoReviewSourceFileId(sourceFileId, historicalSourceIds = [], gates = [], currentGate = {}) {
  const current = clean(sourceFileId);
  const historical = new Set(historicalSourceIds.map(clean).filter(id => id && id !== current));
  const receipts = new Map();
  const failedRoots = new Set();
  for (const gate of gates) {
    if (clean(currentGate.prioritySetId) && clean(gate.prioritySetId) !== clean(currentGate.prioritySetId)) continue;
    if (!historical.has(clean(gate.sourceFileId))) continue;
    const receipt = gate.poetryPleaseHandoff || {};
    const root = clean(gate.sourceReviewFileId) || clean(gate.sourceFileId);
    // A failed legacy attempt still records the identity that was submitted.
    // Require its explicit review-source mapping, matching event and saved
    // sourceRecordId; a title or failed receipt alone is not an identity anchor.
    if (clean(currentGate.prioritySetId) && historical.has(root)
      && root !== clean(gate.sourceFileId) && clean(gate.videoRecordId) === `weaver:video:${root}`
      && clean(receipt.status) === "failed" && receipt.ok === false && clean(receipt.error)
      && clean(receipt.sourceRecordId) === `weaver:video:${root}` && !clean(receipt.canonicalVideoId)) {
      failedRoots.add(root);
    }
    if (!historical.has(root) || clean(receipt.status) !== "sent_to_poetry_please"
      || clean(receipt.sourceRecordId) !== `weaver:video:${root}` || !clean(receipt.canonicalVideoId)) continue;
    const ids = receipts.get(root) || new Set();
    ids.add(clean(receipt.canonicalVideoId));
    receipts.set(root, ids);
  }
  // Prefer an actual earlier receiver receipt over choosing between the
  // reviewed 720p file and its intermediate FV file by name or array order.
  if (receipts.size) {
    if (receipts.size !== 1) return "";
    const [root, ids] = [...receipts][0];
    if ([...failedRoots].some(failedRoot => failedRoot !== root)) return "";
    return ids.size === 1 ? root : "";
  }
  if (failedRoots.size) return failedRoots.size === 1 ? [...failedRoots][0] : "";
  if (historical.size === 1) return [...historical][0];
  return clean(currentGate.sourceReviewFileId);
}

export function buildWeaverVideoImport(candidate, gate) {
  if (clean(gate?.decision) !== "ready_for_poetry_please") {
    throw new Error("Only a ready-for-Poetry Please video can be handed off.");
  }
  if (candidate?.publicationRestricted || gate?.publicationRestricted || clean(candidate?.releaseStatus)) {
    throw new Error("A publication-restricted source cannot be handed off to Poetry Please.");
  }
  if (clean(gate?.publishableAssetMediaType) !== "FV") {
    throw new Error("The final asset must be verified as Finished Video (FV) before handoff.");
  }
  const reviewSourceFileId = clean(candidate?.sourceReviewFileId) || clean(gate?.sourceReviewFileId);
  const sourceFileId = reviewSourceFileId || clean(candidate?.sourceFileId);
  const sourceRecordId = `weaver:video:${sourceFileId}`;
  const sourceEvent = clean(gate?.sourceEvent) || clean(candidate?.sourceEvent);
  const sourceEventLabel = clean(gate?.sourceEventLabel) || clean(candidate?.sourceEventLabel);
  const releaseCatalog = sourceEvent === "BPL Charm City 2026"
    ? "BPL Charm City 2026"
    : clean(gate?.releaseCatalog) || clean(candidate?.releaseCatalog);
  const finalAssetUrl = clean(gate?.publishableAssetUrl);
  const declaredSourceVideoUrl = reviewSourceFileId
    ? `https://drive.google.com/file/d/${encodeURIComponent(sourceFileId)}/view`
    : (clean(gate?.sourceVideoUrl) || clean(candidate?.sourceVideoUrl));
  // A video already in the FV folder can itself be the publishable asset.
  // In that case, sourceFileId carries its provenance; sourceVideoUrl would
  // incorrectly declare the same finished file as raw footage to the receiver.
  const sourceVideoUrl = clean(candidate?.sourceMediaType) === "FV"
    && driveFileId(finalAssetUrl) === sourceFileId
    ? "" : declaredSourceVideoUrl;
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
  if (clean(candidate?.sourceMediaType) !== "FV"
    && (finalAssetUrl === sourceVideoUrl || driveFileId(finalAssetUrl) === sourceFileId)) {
    throw new Error("The final publishable asset cannot be the raw Weaver source video.");
  }
  if (requiresManualVideoIdentityReview(candidate, finalAssetUrl, gate)) {
    throw new Error("This finished video needs a canonical identity review before handoff.");
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
    releaseCatalog,
    eventReleaseCatalog: clean(gate.eventReleaseCatalog) || clean(candidate.eventReleaseCatalog),
    sourceEvent,
    sourceEventLabel,
    gateId: clean(gate.gateId),
    releaseStatus: clean(gate.releaseStatus) || clean(candidate.releaseStatus),
    publicationRestricted: false,
    selectedExcerptRecordIds: [...new Set((Array.isArray(gate.selectedExcerptRecordIds)
      ? gate.selectedExcerptRecordIds : (candidate.excerptRecordIds || [])).map(clean).filter(Boolean))],
    reviews: Array.isArray(candidate.ratings) ? candidate.ratings : [],
    baseScore: candidate.baseScore ?? gate.baseScore,
    excerptBonus: 0,
    candidateScore: candidate.baseScore ?? gate.baseScore,
    reviewCount: Number(candidate.reviewCount || 0),
    excerptCount: Array.isArray(candidate.excerptRecordIds) ? candidate.excerptRecordIds.length : 0,
    decision: "ready_for_poetry_please",
    decidedBy: clean(gate.decidedBy),
    decidedAt: clean(gate.decidedAt)
  };
}

export function requiresManualVideoIdentityReview(candidate, finalAssetUrl, gate = {}) {
  return clean(candidate?.sourceMediaType) === "FV"
    && driveFileId(finalAssetUrl) === clean(candidate?.sourceFileId)
    && !clean(candidate?.sourceReviewFileId)
    && !clean(gate?.sourceReviewFileId);
}

export function prepareVideoIdentityReview(candidate, payload, actor) {
  const gate = candidate?.gate;
  if (!gate || gate.decision !== "ready_for_poetry_please" || candidate.publicationRestricted
    || candidate.existingCanonicalVideoId || clean(gate.poetryPleaseHandoff?.status)) {
    throw new Error("Only an approved, unsent video with no receiver identity can be reviewed here.");
  }
  const sourceReviewFileId = clean(payload.sourceReviewFileId);
  if (clean(gate.sourceReviewFileId) === sourceReviewFileId && gate.identityReview?.status === "verified") {
    return { alreadyVerified: true };
  }
  if (!requiresManualVideoIdentityReview(candidate, gate.publishableAssetUrl, gate)) {
    throw new Error("This gate does not need an unresolved finished-video identity review.");
  }
  const reviewSourceFileIds = [...new Set((candidate.reviewEvidence || []).map(r => clean(r.sourceFileId)).filter(Boolean))];
  if (!sourceReviewFileId || !reviewSourceFileIds.includes(sourceReviewFileId)) {
    throw new Error("Choose a source file with actual stored reviews for this candidate.");
  }
  let evidence;
  try { evidence = new URL(clean(payload.evidenceUrl)); } catch { throw new Error("A Drive or Google Docs evidence link is required."); }
  if (evidence.protocol !== "https:" || !["drive.google.com", "docs.google.com"].includes(evidence.hostname)
    || !/^\/(?:file|document)\/d\/[A-Za-z0-9_-]+(?:\/|$)/.test(evidence.pathname)
    || clean(payload.note).length < 40 || !clean(actor)) {
    throw new Error("Record a Drive/Docs evidence link and a note explaining the verified recording relationship.");
  }
  return { sourceReviewFileId, reviewSourceFileIds, reviewedBy: clean(actor),
    evidenceUrl: evidence.href, note: clean(payload.note) };
}

export function parsePoetryPleaseVideoImport(response, body, sourceRecordId, apiUrl, expected = {}) {
  const item = (Array.isArray(body?.results) ? body.results : [])
    .find(result => clean(result?.sourceRecordId) === sourceRecordId) || {};
  const warnings = [...new Set([
    ...(Array.isArray(body?.warnings) ? body.warnings : []),
    ...(Array.isArray(item.warnings) ? item.warnings : [])
  ].map(clean).filter(Boolean))];
  const status = clean(item.status).toLowerCase();
  const canonicalVideoId = clean(item.canonicalVideoId);
  const canonicalVideoUrl = clean(item.canonicalVideoUrl);
  const reviewsReceived = Number(item.receivedReviewCount) === Number(expected.reviewCount || 0);
  const excerptLinksReceived = Number(item.receivedSelectedExcerptIdCount) === Number(expected.selectedExcerptIdCount || 0);
  const ok = response.ok && body?.ok === true && warnings.length === 0
    && ["created", "updated", "duplicate"].includes(status)
    && Boolean(canonicalVideoId && canonicalVideoUrl)
    && reviewsReceived && excerptLinksReceived;
  return {
    ok,
    status: ok ? "sent_to_poetry_please"
      : (["blocked", "needs_reconciliation"].includes(status) ? status : "failed"),
    canonicalVideoId,
    canonicalVideoUrl: canonicalVideoUrl ? new URL(canonicalVideoUrl, apiUrl).toString() : "",
    finalAssetUrl: clean(item.finalAssetUrl),
    receivedReviewCount: Number(item.receivedReviewCount || 0),
    receivedSelectedExcerptIdCount: Number(item.receivedSelectedExcerptIdCount || 0),
    ...(item.storedReviewCount != null ? { storedReviewCount: Number(item.storedReviewCount) } : {}),
    ...(item.storedSelectedExcerptIdCount != null ? { storedSelectedExcerptIdCount: Number(item.storedSelectedExcerptIdCount) } : {}),
    ...(item.durableCopyStatus != null ? {
      durableCopyStatus: clean(item.durableCopyStatus),
      durableCopyQueued: item.durableCopyQueued === true,
      durableCopyError: clean(item.durableCopyError),
      storageMode: clean(item.storageMode),
      mediaStatus: clean(item.mediaStatus),
      reviewOnly: item.reviewOnly === true,
    } : {}),
    warnings,
    error: ok ? "" : clean(item.error || body?.error || (
      warnings.length ? `Poetry Please import warnings: ${warnings.join(", ")}`
        : !reviewsReceived || !excerptLinksReceived
          ? "Poetry Please did not confirm receipt of all reviews and excerpt links"
          : `Poetry Please returned ${response.status}`
    ))
  };
}

export function parsePoetryPleaseVideoReceipt(receipt, gate) {
  const previous = gate?.poetryPleaseHandoff || {};
  if (receipt?.ok !== true || receipt.received !== true || receipt.ambiguousVideoMatch
    || clean(receipt.sourceRecordId) !== clean(previous.sourceRecordId)
    || clean(receipt.canonicalVideoId) !== clean(previous.canonicalVideoId)
    || !clean(previous.canonicalVideoId)
    || clean(receipt.event?.sourceEvent) !== clean(gate.sourceEvent)
    || (clean(gate.sourceEventLabel) && clean(receipt.event?.sourceEventLabel) !== clean(gate.sourceEventLabel))) {
    throw new Error("Poetry Please receipt identity or event does not match this gate.");
  }
  const reviews = Number(receipt.review?.reviewCount);
  const excerpts = Number(receipt.review?.selectedExcerptIdCount);
  if (!Number.isInteger(reviews) || !Number.isInteger(excerpts)
    || reviews < Number(previous.receivedReviewCount || 0)
    || excerpts < Math.max(Number(previous.receivedSelectedExcerptIdCount || 0), (gate.selectedExcerptRecordIds || []).length)) {
    throw new Error("Poetry Please receipt is missing previously received review context.");
  }
  const media = receipt.media || {};
  const bytes = Number(media.durableByteCount || 0);
  if (clean(media.durableCopyStatus) === "verified") {
    let hosted;
    try { hosted = new URL(media.videoUrl); } catch {}
    if (clean(media.storageMode) !== "durable" || !Number.isSafeInteger(bytes) || bytes <= 0
      || !clean(media.durableChecksum) || hosted?.protocol !== "https:"
      || !["firebasestorage.googleapis.com", "storage.googleapis.com"].includes(hosted?.hostname)) {
      throw new Error("Poetry Please verified copy receipt is incomplete.");
    }
  }
  return {
    sourceRecordId: previous.sourceRecordId, canonicalVideoId: previous.canonicalVideoId,
    durableCopyStatus: clean(media.durableCopyStatus), durableCopyError: clean(media.durableCopyError),
    durableByteCount: bytes, durableChecksum: clean(media.durableChecksum),
    storageMode: clean(media.storageMode), mediaStatus: clean(media.mediaStatus),
    reviewOnly: media.reviewOnly === true,
    storedReviewCount: reviews, storedSelectedExcerptIdCount: excerpts,
    ...(clean(media.durableCopyStatus) === "verified" ? { finalAssetUrl: media.videoUrl } : {}),
  };
}
