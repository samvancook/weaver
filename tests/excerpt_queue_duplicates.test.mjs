import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { runInNewContext } from "node:vm";

const source = readFileSync(new URL("../public/app.js", import.meta.url), "utf8");
function extract(name, next) {
  const start = source.indexOf(`function ${name}(`);
  const end = source.indexOf(`function ${next}(`, start);
  assert.ok(start >= 0 && end > start);
  return source.slice(start, end);
}
const { annotateExactQueueDuplicates } = runInNewContext(
  `${extract("normalizeExactQueueDuplicateText", "buildExcerptCard")}\n({ annotateExactQueueDuplicates })`
);
function row(sourceRow, title, excerptText) {
  return { sourceRow, recordId: `row-${sourceRow}`, title, excerptText, bookTitle: "Event", author: "Poet" };
}

test("identical excerpt records remain independently selectable, including the later approved source", () => {
  const rows = [row(11256, "Poem", "Same excerpt"), row(11273, "Poem", "Same excerpt"), row(11260, "Other", "Same excerpt")];
  const cards = annotateExactQueueDuplicates(rows, { includeTitle: true });
  assert.equal(cards.length, 3);
  assert.equal(cards[1].sourceRow, 11273);
  assert.equal(cards[0].queueDuplicateRows[0].sourceRow, 11273);
  assert.equal(cards[1].queueDuplicateRows[0].sourceRow, 11256);
  assert.equal(cards[2].queueDuplicateRows, undefined);
});

test("annotations preserve author and book boundaries and clear stale peers", () => {
  const rows = [row(20, "Poem", "Same excerpt"), row(21, "Poem", "Same excerpt"), { ...row(22, "Poem", "Same excerpt"), author: "Other" }];
  annotateExactQueueDuplicates(rows, { includeTitle: true });
  annotateExactQueueDuplicates(rows, { includeTitle: true });
  assert.equal(rows[0].queueDuplicateRows.length, 1);
  assert.equal(rows[2].queueDuplicateRows, undefined);
  annotateExactQueueDuplicates([rows[0]], { includeTitle: true });
  assert.equal(rows[0].queueDuplicateRows, undefined);
});

function card(sourceRow, decision, peers = []) {
  return {
    dataset: { sourceRow: String(sourceRow), recordId: `row-${sourceRow}`, currentUseForInt: "1", queueDuplicateRows: JSON.stringify(peers) },
    querySelector: selector => selector === 'input[type="radio"]:checked' ? { value: decision } : null
  };
}
const { collectUpdatesFromContainer, filterChangedUpdates } = runInNewContext(
  `${extract("collectUpdatesFromContainer", "normalizeCorrectionNote")}\n({ collectUpdatesFromContainer, filterChangedUpdates })`,
  { currentValidationByRecordId: new Map(), normalizeDecision: value => value || "", normalizeCorrectionNote: value => String(value || "").trim(), getEffectiveReviewDecision: row => row.excerptReviewDecision || "" }
);

test("caption-only save uses each selected source ID, does not reject peers, and disables both graphic lanes", () => {
  const cards = [card(11260, "accept_skip_graphic"), card(11273, "accept_skip_graphic", [{ sourceRow: 11256, recordId: "row-11256" }]), card(11253, "accept_skip_graphic"), card(11256, "")];
  const updates = collectUpdatesFromContainer({ querySelectorAll: () => cards });
  const changed = filterChangedUpdates(updates, cards.map(c => ({ sourceRow: Number(c.dataset.sourceRow), excerptReviewDecision: "" })));
  assert.deepEqual(Array.from(changed, r => r.sourceRow), [11260, 11273, 11253]);
  assert.equal(changed.length, 3);
  for (const saved of changed) {
    assert.equal(saved.recordId, `row-${saved.sourceRow}`);
    assert.equal(saved.reviewDecision, "accept_skip_graphic");
    assert.equal(saved.useForQi, false);
    assert.equal(saved.useForInt, false);
  }
});

test("ordinary acceptance and rejection never synthesize a decision for an identical record", () => {
  for (const decision of ["accept", "reject"]) {
    const updates = collectUpdatesFromContainer({ querySelectorAll: () => [card(11273, decision, [{ sourceRow: 11256 }])] });
    assert.equal(updates.length, 1);
    assert.equal(updates[0].sourceRow, 11273);
    assert.equal(updates[0].reviewDecision, decision);
  }
});
