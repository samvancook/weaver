const clean = value => String(value ?? '').trim();
const httpsUrl = value => {
  try { const url = new URL(value); return url.protocol === 'https:' ? url.toString() : ''; }
  catch { return ''; }
};

export function footageProvenanceForFile(connections, fileId) {
  return connections.filter(c => c.sourceFileId === fileId).map(c => ({
    inventoryEventId:c.inventoryEventId,inventoryEventName:c.inventoryEventName,
    sourceSpreadsheetId:c.sourceSpreadsheetId,sourceRow:c.sourceRow,
    familyId:c.familyId,originalMediaFileId:c.workflow?.originalMediaFileId || null,
    finishedVideoFileId:c.workflow?.finishedVideoFileId || null,
  }));
}

// Read-only projection of exact event/file identities. A connection is not a
// review, a PP receipt is not publication, and neither proves calendar placement.
export function buildFootageProgress(record, reviews = [], gates = [], checkedAt = new Date().toISOString()) {
  const byFileId = {};
  for (const fileId of new Set((record.accepted || []).map(t => t.sourceFileId))) {
    const people = new Map();
    for (const review of reviews) {
      if (review.prioritySetId !== record.prioritySetId || review.sourceFileId !== fileId || !clean(review.rating)) continue;
      const person = clean(review.reviewerEmail).toLowerCase();
      if (!person) continue;
      const old = people.get(person);
      if (!old || clean(review.updatedAt) > clean(old.updatedAt)) people.set(person, review);
    }
    const matchingGates = gates.filter(g => g.prioritySetId === record.prioritySetId && g.sourceFileId === fileId)
      .sort((a,b) => clean(b.updatedAt || b.decidedAt).localeCompare(clean(a.updatedAt || a.decidedAt)));
    const gate = matchingGates[0] || {};
    const handoff = gate.poetryPleaseHandoff || {};
    const root = clean(gate.sourceReviewFileId) || fileId;
    const received = handoff.status === 'sent_to_poetry_please' && handoff.ok === true
      && handoff.sourceRecordId === `weaver:video:${root}` && Boolean(clean(handoff.canonicalVideoId))
      && Boolean(httpsUrl(handoff.canonicalVideoUrl));
    const decision = clean(gate.decision);
    const excerptIds = new Set([...people.values()].flatMap(r => r.excerptRecordIds || [])
      .concat(gate.selectedExcerptRecordIds || []).map(clean).filter(Boolean));
    const state = received ? 'received_by_poetry_please'
      : handoff.status === 'sending' ? 'handoff_in_progress'
      : handoff.status ? 'handoff_needs_attention'
      : decision === 'send_to_editing' ? 'editing_requested'
      : decision === 'hold' ? 'on_hold' : decision === 'reject' ? 'rejected'
      : decision === 'ready_for_poetry_please' ? 'ready_for_handoff'
      : people.size ? 'curation_in_progress' : 'awaiting_curation';
    byFileId[fileId] = {
      state, reviewCount: people.size, linkedExcerptCount: excerptIds.size, decision,
      poetryPlease: {
        status: received ? 'received' : handoff.status === 'sent_to_poetry_please' ? 'receipt_identity_unverified' : clean(handoff.status) || 'not_sent',
        canonicalVideoId: received ? clean(handoff.canonicalVideoId) : null,
        canonicalVideoUrl: received ? httpsUrl(handoff.canonicalVideoUrl) : '',
        durableCopyStatus: received ? clean(handoff.durableCopyStatus) || 'not_checked' : 'not_checked',
        error: clean(handoff.error), lastReceiptAt: clean(handoff.updatedAt),
      },
      calendarStatus: 'not_checked', checkedAt,
    };
  }
  return {status:'checked', checkedAt, byFileId};
}
