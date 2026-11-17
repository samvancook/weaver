// A grouped performance may display a new representative while its accepted
// handoff remains on the original gate. Refresh that gate, never the UI file.
export async function refreshAcceptedVideoReceipt({candidate, receipt, db, parseReceipt, checkedBy = 'weaver:video-pipeline'}) {
  const gate = candidate.gate;
  const handoff = gate?.poetryPleaseHandoff || {};
  if (handoff.status !== 'sent_to_poetry_please') throw Error('Accepted video gate required for receipt refresh');
  if (!gate.prioritySetId || !gate.sourceFileId || gate.prioritySetId !== candidate.prioritySetId
      || handoff.sourceRecordId !== candidate.videoRecordId) {
    throw Error('Accepted video gate does not match the candidate performance');
  }
  const parsed = parseReceipt(receipt, gate);
  const fields = ['durableCopyStatus','durableCopyError','durableByteCount','durableChecksum','storageMode','mediaStatus','storedReviewCount','storedSelectedExcerptIdCount'];
  if (!fields.some(key => String(parsed[key] ?? '') !== String(handoff[key] ?? ''))) return {saved:false, unchanged:true};
  const result = await db('refresh_video_curation_receipt', {
    prioritySetId:gate.prioritySetId, sourceFileId:gate.sourceFileId,
    attemptId:handoff.attemptId || '', checkedBy, receipt:parsed
  });
  if (!result?.saved) throw Error(`Receipt refresh not saved: ${result?.reason || 'unknown_failure'}`);
  return result;
}
