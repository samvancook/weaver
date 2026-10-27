import { createHash } from "node:crypto";

export function bookIntakeContentHash(payload = {}) {
  const fields = ["author", "title", "bookTitle", "quote"];
  return createHash("sha256")
    .update(JSON.stringify(fields.map(field => String(payload[field] || "").trim())))
    .digest("hex");
}

export async function claimBookIntakeRecord({ recordId, payload, token, projectId,
  databaseId = "weaverledger", fetchImpl = fetch, now = new Date(), staleAfterMs = 10 * 60 * 1000 }) {
  if (!recordId || !token || !projectId) throw new Error("Book intake claim is missing its identity or service credential.");
  const documentId = createHash("sha256").update(recordId).digest("hex");
  const url = new URL(`https://firestore.googleapis.com/v1/projects/${encodeURIComponent(projectId)}`
    + `/databases/${encodeURIComponent(databaseId)}/documents/bookIntakeClaims/${documentId}`);
  const hash = bookIntakeContentHash(payload);
  const fields = {
    recordId: { stringValue: recordId },
    contentHash: { stringValue: hash },
    claimedAt: { timestampValue: now.toISOString() }
  };
  url.searchParams.set("currentDocument.exists", "false");
  const response = await fetchImpl(url, {
    method: "PATCH",
    signal: AbortSignal.timeout(15000),
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ fields })
  });
  if (response.ok) return { claimed: true };
  if (response.status !== 412) {
    throw new Error(`Book intake claim failed (HTTP ${response.status}); the Sheet was not appended.`);
  }
  url.searchParams.delete("currentDocument.exists");
  const existingResponse = await fetchImpl(url, {
    signal: AbortSignal.timeout(15000),
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!existingResponse.ok) {
    throw new Error("Book intake claim could not be reconciled; the Sheet was not appended.");
  }
  const existing = await existingResponse.json();
  if (existing?.fields?.recordId?.stringValue !== recordId
    || existing?.fields?.contentHash?.stringValue !== hash) {
    throw new Error(`Excerpt record ID ${recordId} already belongs to different content.`);
  }
  const claimedAt = Date.parse(existing?.fields?.claimedAt?.timestampValue || "");
  if (Number.isFinite(claimedAt) && now.getTime() - claimedAt >= staleAfterMs && existing.updateTime) {
    url.searchParams.set("currentDocument.updateTime", existing.updateTime);
    const takeover = await fetchImpl(url, {
      method: "PATCH",
      signal: AbortSignal.timeout(15000),
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ fields })
    });
    if (takeover.ok) return { claimed: true, recoveredStaleClaim: true };
    if (takeover.status !== 412) throw new Error(`Book intake claim recovery failed (HTTP ${takeover.status}).`);
  }
  return { claimed: false };
}
