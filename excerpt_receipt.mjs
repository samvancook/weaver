import { createHash } from 'node:crypto';

const clean = value => String(value ?? '').trim();
export const excerptFingerprint = value => createHash('sha256')
  .update(clean(value).normalize('NFKC').replace(/\s+/gu, ' ').toLowerCase()).digest('hex');

// A bounded ledger read can prove a positive acknowledgement. An absent row
// cannot prove that PP has no record, and must never authorize a resend.
export function findExcerptAcknowledgement(record, health, canonicalVideoId = '') {
  if (health?.ok !== true || !Array.isArray(health.ledgers)) throw new Error('Receiver ledger unavailable.');
  const sourceRecordId = clean(record.sourceRecordId);
  if (!sourceRecordId) throw new Error('The excerpt has no source identity.');
  const matches = health.ledgers.flatMap(ledger =>
    ['completed', 'completed_with_errors'].includes(ledger.status)
      ? (ledger.outcomes || []).filter(row => clean(row.sourceRecordId) === sourceRecordId
        && ['created', 'updated', 'duplicate'].includes(row.outcome) && !row.error)
        .map(row => ({ ...row, importLedgerId: ledger.id })) : []);
  const ids = [...new Set(matches.map(row => clean(row.canonicalExcerptId || row.contentId)).filter(Boolean))];
  if (ids.length !== 1) throw new Error(ids.length ? 'Conflicting receiver identities require review.'
    : 'No exact acknowledgement in the checked receiver ledgers. This does not prove absence; do not resend.');
  const acknowledgements = matches.filter(row => clean(row.canonicalExcerptId || row.contentId) === ids[0]);
  if (canonicalVideoId && acknowledgements.some(row => clean(row.canonicalVideoId) !== canonicalVideoId)) {
    throw new Error('The receiver acknowledgement links a different or unconfirmed video.');
  }
  if (record.poetryPleaseItemId && record.poetryPleaseItemId !== ids[0]) throw new Error('The saved receiver identity conflicts.');
  return { sourceRecordId, canonicalExcerptId: ids[0], canonicalVideoId,
    importLedgerIds: [...new Set(acknowledgements.map(row => row.importLedgerId))] };
}

export function verifyExcerptReadback(record, acknowledgement, body) {
  const item = body?.item;
  if (!Array.isArray(item) || body.imageId !== acknowledgement.canonicalExcerptId
    || item[3] !== acknowledgement.canonicalExcerptId || item[7] !== 'EXC') {
    throw new Error('The current receiver excerpt is unavailable or has a different identity.');
  }
  if (!clean(record.excerpt) || excerptFingerprint(item[8]) !== excerptFingerprint(record.excerpt)
    || excerptFingerprint(item[0]) !== excerptFingerprint(record.author)
    || excerptFingerprint(item[1]) !== excerptFingerprint(record.poemTitle)) {
    throw new Error('The current receiver author, title or excerpt differs from the approved source.');
  }
  return { ...acknowledgement, verified: true, excerptFingerprint: excerptFingerprint(record.excerpt),
    checkedAtMs: Date.now(), checkedAt: new Date().toISOString() };
}
