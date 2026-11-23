// Read-only milestone over the current included performance groups. No score,
// publication, editing, hold, source registration or closure state is changed.
export function videoEventCurationMilestone(items = []) {
  const total = items.length;
  const reviewsComplete = items.filter(item => item.reviewProgress?.reviewComplete === true).length;
  const withExcerpts = items.filter(item => Number(item.reviewProgress?.excerptCount) > 0).length;
  const complete = total > 0 && reviewsComplete === total && withExcerpts === total;
  return {
    curationComplete: complete,
    curationStatus: complete ? 'curation_complete' : 'curation_in_progress',
    curationLabel: complete ? 'Curation complete — reviews and excerpts collected' : 'Curation in progress',
    reviewsCompleteCount: reviewsComplete,
    excerptCoveredCount: withExcerpts,
    missingExcerptCount: total - withExcerpts
  };
}
