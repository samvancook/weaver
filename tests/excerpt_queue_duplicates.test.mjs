import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { runInNewContext } from "node:vm";

const source = readFileSync(new URL("../public/app.js", import.meta.url), "utf8");
const start = source.indexOf("function normalizeExactQueueDuplicateText(");
const end = source.indexOf("function buildExcerptCard(", start);
assert.ok(start >= 0 && end > start);
const { collapseExactQueueDuplicates } = runInNewContext(
  `${source.slice(start, end)}\n({ collapseExactQueueDuplicates })`
);

function row(sourceRow, title, excerptText) {
  return { sourceRow, recordId: `row-${sourceRow}`, title, excerptText };
}

test("exact duplicates beyond the first page stay attached to the canonical card", () => {
  const rows = [
    row(10, "Poem", "Same excerpt"),
    row(11, "Poem", "Same excerpt"),
    row(12, "Other poem", "Same excerpt"),
    row(13, "Poem", "Different excerpt")
  ];

  const cards = collapseExactQueueDuplicates(rows, { includeTitle: true });
  assert.equal(cards.length, 3);
  assert.equal(cards[0].sourceRow, 10);
  assert.equal(cards[0].queueDuplicateRows.length, 1);
  assert.equal(cards[0].queueDuplicateRows[0].sourceRow, 11);
  assert.equal(cards.slice(0, 1)[0].queueDuplicateRows[0].sourceRow, 11);
  assert.equal(cards[1].sourceRow, 12);
});

test("repeated renders rebuild duplicate annotations without stale rows", () => {
  const rows = [row(20, "Poem", "Same excerpt"), row(21, "Poem", "Same excerpt")];

  collapseExactQueueDuplicates(rows, { includeTitle: true });
  collapseExactQueueDuplicates(rows, { includeTitle: true });
  assert.equal(rows[0].queueDuplicateRows.length, 1);

  collapseExactQueueDuplicates([rows[0]], { includeTitle: true });
  assert.equal(rows[0].queueDuplicateRows, undefined);
});

test("the review page is sliced after exact duplicates are collapsed", () => {
  const renderStart = source.indexOf("function renderExcerpts(");
  const renderEnd = source.indexOf("function renderCorrectionExcerpts(", renderStart);
  assert.ok(renderStart >= 0 && renderEnd > renderStart);
  let visibleCards;
  const { renderExcerpts } = runInNewContext(
    `${source.slice(renderStart, renderEnd)}\n({ renderExcerpts })`,
    {
      orderReviewExcerptsForRender: rows => rows,
      collapseExactQueueDuplicates,
      normalizeExactQueueDuplicateText: value => String(value).replace(/\r\n/g, "\n").trim(),
      cleanSheetWhitespace: value => String(value).trim().replace(/\s+/g, " "),
      reviewShowAdditionalPulls: false,
      reviewVisibleCount: 1,
      renderExcerptCollection: rows => { visibleCards = rows; },
      elements: { excerptList: {}, excerptCountBadge: {} },
      getEmptyStateMessage: () => "",
      getReviewBatchSize: () => 1
    }
  );

  renderExcerpts([row(30, "Poem", "Same excerpt"), row(31, "Poem", "Same excerpt")]);
  assert.equal(visibleCards.length, 1);
  assert.equal(visibleCards[0].queueDuplicateRows[0].sourceRow, 31);
});
