import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

const html = readFileSync(new URL("../public/index.html", import.meta.url), "utf8");
const app = readFileSync(new URL("../public/app.js", import.meta.url), "utf8");
const server = readFileSync(new URL("../server.mjs", import.meta.url), "utf8");

test("priority video progress and scoreboard remain visible from the gear", () => {
  assert.match(html, /<details[^>]*id="video-progress-management"[^>]*\bopen\b/);
  assert.match(html, /id="video-progress-set-rows"/);
  assert.match(html, /id="video-progress-reviewer-rows"/);
  assert.match(html, /id="video-curation-candidates"/);

  const gearHandler = app.match(/elements\.showGraphicsOps\?\.addEventListener\("click", \(\) => \{([\s\S]*?)\n\}\);/);
  assert.ok(gearHandler, "gear click handler must exist");
  assert.match(gearHandler[1], /currentGraphicsView = "ops"/);
  assert.match(gearHandler[1], /videoProgressManagement\.open = true/);
  assert.match(gearHandler[1], /loadPriorityVideoProgress\(\)/);
  assert.match(gearHandler[1], /loadVideoCurationCandidates\(\)/);
});

test("video scoreboard offers an event filter and ranks within event groups", () => {
  assert.match(html, /<select id="video-curation-event">/);
  assert.match(html, /<option value="">All events<\/option>/);
  assert.match(app, /videoCurationEvent\?\.addEventListener\("change", renderVideoCurationCandidates\)/);
  assert.match(app, /const groups = new Map\(\)/);
  assert.match(app, /candidates\.sort\(\(a, b\) => Number\(b\.candidateScore/);
});

test("scoreboard keeps reviewed videos below threshold and links every excerpt", () => {
  assert.match(server, /if \(!reviews\.length && !existingGate\) return null/);
  assert.match(server, /return candidate;\s*\}\)\.filter\(Boolean\)/);
  assert.match(server, /selectedExcerptRecordIds: candidate\.excerptRecordIds/);
  assert.match(server, /const gate = \{ \.\.\.candidate\.gate, selectedExcerptRecordIds: candidate\.excerptRecordIds \}/);
  assert.doesNotMatch(app, /data-excerpt-id/);
  assert.match(app, /candidate\.ratings/);
  assert.match(app, /candidate\.baseScore/);
});

test("empty imported numeric scores are not displayed or averaged as zero", () => {
  assert.match(server, /rawLegacyScore === "" \? NaN : Number\(rawLegacyScore\)/);
  assert.match(app, /rating\.legacyScore !== "" && Number\.isFinite\(Number\(rating\.legacyScore\)\)/);
});
