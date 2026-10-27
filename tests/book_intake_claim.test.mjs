import assert from "node:assert/strict";
import test from "node:test";
import { bookIntakeContentHash, claimBookIntakeRecord } from "../book_intake_claim.mjs";

const payload = { author: "A", title: "P", bookTitle: "B", quote: "line" };
const base = { recordId: "weaver:one", payload, token: "test", projectId: "test" };

test("only one concurrent request claims an intake identity", async () => {
  let saved;
  const fetchImpl = async (_url, options) => {
    if (options.method === "PATCH") {
      if (saved) return { ok: false, status: 412 };
      saved = JSON.parse(options.body);
      return { ok: true, status: 200 };
    }
    return { ok: true, json: async () => saved };
  };
  const [first, second] = await Promise.all([
    claimBookIntakeRecord({ ...base, fetchImpl }),
    claimBookIntakeRecord({ ...base, fetchImpl })
  ]);
  assert.deepEqual([first.claimed, second.claimed].sort(), [false, true]);
  assert.equal(saved.fields.contentHash.stringValue, bookIntakeContentHash(payload));
});

test("an existing identity rejects different excerpt content", async () => {
  const fetchImpl = async (_url, options) => options.method === "PATCH"
    ? { ok: false, status: 412 }
    : { ok: true, json: async () => ({ fields: {
      recordId: { stringValue: base.recordId },
      contentHash: { stringValue: bookIntakeContentHash(payload) }
    } }) };
  await assert.rejects(
    claimBookIntakeRecord({ ...base, payload: { ...payload, quote: "changed" }, fetchImpl }),
    /different content/
  );
});

test("a failed claim prevents an uncertain Sheet append", async () => {
  await assert.rejects(
    claimBookIntakeRecord({ ...base, fetchImpl: async () => ({ ok: false, status: 503 }) }),
    /Sheet was not appended/
  );
});

test("a stale interrupted claim is recovered with a version precondition", async () => {
  const requests = [];
  const fetchImpl = async (url, options) => {
    requests.push({ url: String(url), method: options.method });
    if (!options.method) return { ok: true, json: async () => ({
      updateTime: "2026-10-27T12:00:00Z",
      fields: {
        recordId: { stringValue: base.recordId },
        contentHash: { stringValue: bookIntakeContentHash(payload) },
        claimedAt: { timestampValue: "2026-10-27T12:00:00Z" }
      }
    }) };
    return requests.length === 1 ? { ok: false, status: 412 } : { ok: true, status: 200 };
  };
  const result = await claimBookIntakeRecord({ ...base, fetchImpl,
    now: new Date("2026-10-27T12:11:00Z") });
  assert.deepEqual(result, { claimed: true, recoveredStaleClaim: true });
  assert.match(requests[2].url, /currentDocument.updateTime=/);
});
