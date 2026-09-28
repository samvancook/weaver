import test from "node:test";
import assert from "node:assert/strict";
import { verifyProgressExportCaller } from "../weaver_progress_auth.mjs";

const req = token => ({ headers: { authorization: `Bearer ${token}` } });
const audience = "https://weaver-912447899335.us-central1.run.app";
const email = "weaver-deployer@button-weaver-internal.iam.gserviceaccount.com";
const verifier = payload => ({
  async verifyIdToken({ audience: requestedAudience }) {
    assert.equal(requestedAudience, audience);
    return { getPayload: () => payload };
  }
});

test("existing administrator access tokens retain their authorization path", async () => {
  const principal = { email: "sam@buttonpoetry.com" };
  assert.equal(await verifyProgressExportCaller(req("opaque-access-token"), async () => principal), principal);
});

test("only the exact verified deployer ID token can read the export", async () => {
  const principal = await verifyProgressExportCaller(req("a.b.c"), () => {
    throw new Error("admin path should not run");
  }, verifier({ email, email_verified: true, sub: "123" }));
  assert.equal(principal.email, email);
  await assert.rejects(verifyProgressExportCaller(req("a.b.c"), () => {}, verifier({
    email: "other@button-weaver-internal.iam.gserviceaccount.com", email_verified: true
  })), error => error.statusCode === 403);
  await assert.rejects(verifyProgressExportCaller(req("a.b.c"), () => {}, verifier({
    email, email_verified: false
  })), error => error.statusCode === 403);
});

test("invalid signed tokens do not fall back to administrator access", async () => {
  await assert.rejects(verifyProgressExportCaller(req("a.b.c"), () => {
    throw new Error("admin path should not run");
  }, { async verifyIdToken() { throw new Error("bad signature"); } }), error => error.statusCode === 401);
});
