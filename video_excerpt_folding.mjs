function exactKey(value) {
  return String(value || "").trim().replace(/\s+/g, " ").toLowerCase();
}

export function findApprovedTextExcerptMatch(videoExcerpt, approvedTextExcerpts) {
  const fields = ["author", "bookTitle", "poemTitle", "excerptText"];
  if (fields.some(field => !exactKey(videoExcerpt?.[field]))) return null;

  const matches = approvedTextExcerpts.filter(record =>
    !record.sourceVideoFileId
    && record.sourceRecordId
    && record.approval?.approvedForUse === true
    && fields.every(field => exactKey(record[field]) === exactKey(videoExcerpt[field]))
  );
  return matches.length === 1 ? matches[0] : null;
}
