import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

const html = readFileSync(new URL("../public/index.html", import.meta.url), "utf8");
const app = readFileSync(new URL("../public/app.js", import.meta.url), "utf8");

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
