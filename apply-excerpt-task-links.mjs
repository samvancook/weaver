import fs from 'node:fs';import path from 'node:path';import assert from 'node:assert/strict';
const root=process.argv[2]||'/app',patches=[
  {
    "file": "public/app.js",
    "before": "let graphicsBookSummaryByKey = new Map();\nlet gatheringCatalogPopup = null;\nconst contributorShellParams = new URLSearchParams(window.location.search);\n\nconst REVIEW_SINGLE_BATCH_SIZE = 1;\nconst REVIEW_MULTI_BATCH_SIZE = 25;\n",
    "after": "let graphicsBookSummaryByKey = new Map();\nlet gatheringCatalogPopup = null;\nconst contributorShellParams = new URLSearchParams(window.location.search);\nlet excerptReviewTask = null;\nconst REVIEW_EXCERPT_TASK_KEY = \"__excerpt_review_task__\";\n\nconst REVIEW_SINGLE_BATCH_SIZE = 1;\nconst REVIEW_MULTI_BATCH_SIZE = 25;\n"
  },
  {
    "file": "public/app.js",
    "before": "  };\n}\n\nfunction getReviewQueueIncludeSet() {\n  const titles = Array.isArray(runtimeConfig.reviewQueueIncludeTitles)\n    ? runtimeConfig.reviewQueueIncludeTitles\n",
    "after": "  };\n}\n\nfunction parseExcerptReviewTask(search) {\n  const params = new URLSearchParams(search);\n  if (!params.has(\"task\")) return null;\n  if (params.getAll(\"task\").length !== 1 || params.get(\"task\") !== \"approve-video-excerpts\") throw new Error(\"Unknown excerpt review assignment. No other queue was opened.\");\n  if (params.getAll(\"videoFiles\").length !== 1 || params.has(\"videoSet\") || params.has(\"view\")) throw new Error(\"This assignment needs one unambiguous video scope.\");\n  const ids = params.get(\"videoFiles\").split(\",\");\n  if (!ids.length || ids.length > 50 || ids.some(id => !/^[A-Za-z0-9_-]{15,100}$/.test(id))) throw new Error(\"Invalid video IDs in this assignment. No other queue was opened.\");\n  return {kind:\"approve-video-excerpts\", videoFileIds:[...new Set(ids)].sort()};\n}\n\nfunction filterExcerptTaskRecords(records, task) {\n  if (!task) return records;\n  const ids = new Set(task.videoFileIds);\n  return records.filter(record => ids.has(record.sourceVideoFileId || extractGoogleDriveFileId(record.sourceVideoUrl)));\n}\n\nfunction buildExcerptReviewTaskUrl(origin, ids) {\n  const files = [...new Set(ids)].sort();\n  const url = new URL(\"/\", origin);\n  url.searchParams.set(\"task\", \"approve-video-excerpts\");\n  url.searchParams.set(\"videoFiles\", files.join(\",\"));\n  parseExcerptReviewTask(url.search);\n  return url.href;\n}\n\nfunction updateExcerptAssignmentUi() {\n  const panel = document.getElementById(\"excerpt-assignment-panel\");\n  const description = document.getElementById(\"excerpt-assignment-status\");\n  if (excerptReviewTask && panel && description) {\n    panel.hidden = false;\n    const pending = filterExcerptTaskRecords(currentPendingRecords, excerptReviewTask);\n    const titles = [...new Set(pending.map(r => `${r.author} \u2014 ${r.title}`))];\n    const progressKnown = excerptReviewTask.videoFileIds.every(id => videoExcerptProgressByFileId.has(id));\n    const approvedVideos = excerptReviewTask.videoFileIds.filter(id => videoExcerptProgressByFileId.get(id)?.approvedCount > 0).length;\n    const progressLabel = progressKnown ? `Approved excerpt present for ${approvedVideos} of ${excerptReviewTask.videoFileIds.length} videos. ` : \"Approval counts unavailable. \";\n    description.textContent = progressLabel + `${excerptReviewTask.videoFileIds.length} linked video(s) \u00b7 ${pending.length} pending excerpts. ${titles.join(\"; \")}${pending.length ? \". For calendar readiness, approve at least one suitable excerpt per video. Use a light typo/quality check; leave weak excerpts unapproved.\" : \"No pending excerpts are available for these exact videos. They may already be reviewed or unavailable; another queue has not been substituted.\"}`;\n  }\n  const copy = document.getElementById(\"copy-excerpt-review-link\");\n  if (copy) copy.hidden = getSelectedReviewFilter() !== \"videos\";\n}\n\nasync function copyExcerptReviewLink() {\n  try {\n    const key = elements.bookSelect.value;\n    const records = key ? getPendingRecordsForBookKey(key) : [];\n    const ids = excerptReviewTask && key === REVIEW_EXCERPT_TASK_KEY ? excerptReviewTask.videoFileIds : records.map(r => r.sourceVideoFileId || extractGoogleDriveFileId(r.sourceVideoUrl)).filter(Boolean);\n    const url = buildExcerptReviewTaskUrl(window.location.origin, ids);\n    const field = document.getElementById(\"excerpt-review-task-url\");\n    field.value = url; field.hidden = false;\n    try { await navigator.clipboard.writeText(url); setStatus(\"Review assignment link copied. It opens this exact video scope and grants no access or approval.\"); }\n    catch { setStatus(\"Copy the review assignment URL shown beside the button.\"); }\n  } catch (error) { setStatus(error.message); }\n}\n\nasync function openExcerptReviewAssignment(task) {\n  excerptReviewTask = task;\n  setActiveModule(\"review\");\n  elements.reviewFilter.value = \"videos\";\n  elements.reviewFilter.disabled = true;\n  elements.reviewDisplayMode.value = \"batch\";\n  updateExcerptAssignmentUi();\n  if (await loadBooks({preserveSelection:false}) !== true) throw new Error(\"This assignment could not load. Refresh Books to retry; no other queue was opened.\");\n  elements.bookSelect.value = REVIEW_EXCERPT_TASK_KEY;\n  await loadVideoExcerptProgress();\n  await loadExcerpts();\n  updateExcerptAssignmentUi();\n}\n\nfunction getReviewQueueIncludeSet() {\n  const titles = Array.isArray(runtimeConfig.reviewQueueIncludeTitles)\n    ? runtimeConfig.reviewQueueIncludeTitles\n"
  },
  {
    "file": "public/app.js",
    "before": "  if (getSelectedReviewFilter() === \"videos\" && elements.bookSelect) {\n    const select = elements.bookSelect;\n    select.replaceChildren(new Option(\"Choose an event or video\", \"\"));\n    let group = null;\n    let previousEvent = \"\";\n    for (const video of visibleBooks) {\n",
    "after": "  if (getSelectedReviewFilter() === \"videos\" && elements.bookSelect) {\n    const select = elements.bookSelect;\n    select.replaceChildren(new Option(\"Choose an event or video\", \"\"));\n    if (excerptReviewTask) select.append(new Option(\"This assignment \u2014 all linked videos\", REVIEW_EXCERPT_TASK_KEY));\n    let group = null;\n    let previousEvent = \"\";\n    for (const video of visibleBooks) {\n"
  },
  {
    "file": "public/app.js",
    "before": "      const progressLabel = progress === undefined ? \"counts unavailable\" : `${progress.pulledCount} pulled (goal 3+), ${progress.approvedCount} approved`;\n      group.append(new Option(`${video.title} \u00b7 ${video.author} (${video.standardCount} pending, ${progressLabel})`, video.key));\n    }\n    if (visibleBooks.some(video => video.key === previousSelection)\n      || [...select.options].some(option => option.value === previousSelection)) select.value = previousSelection;\n  } else {\n    populateBookSelect(elements.bookSelect, visibleBooks, previousSelection, book => `${book.title} (${book.standardCount})`);\n  }\n",
    "after": "      const progressLabel = progress === undefined ? \"counts unavailable\" : `${progress.pulledCount} pulled (goal 3+), ${progress.approvedCount} approved`;\n      group.append(new Option(`${video.title} \u00b7 ${video.author} (${video.standardCount} pending, ${progressLabel})`, video.key));\n    }\n    if (excerptReviewTask && (!previousSelection || previousSelection === REVIEW_EXCERPT_TASK_KEY)) select.value = REVIEW_EXCERPT_TASK_KEY;\n    if (previousSelection && (visibleBooks.some(video => video.key === previousSelection)\n      || [...select.options].some(option => option.value === previousSelection))) select.value = previousSelection;\n  } else {\n    populateBookSelect(elements.bookSelect, visibleBooks, previousSelection, book => `${book.title} (${book.standardCount})`);\n  }\n"
  },
  {
    "file": "public/app.js",
    "before": "    videoExcerptProgressByFileId = new Map();\n  }\n  refreshReviewBookSelect(true);\n}\n\nfunction getBookTitleDisplayScore(title) {\n",
    "after": "    videoExcerptProgressByFileId = new Map();\n  }\n  refreshReviewBookSelect(true);\n  updateExcerptAssignmentUi();\n}\n\nfunction getBookTitleDisplayScore(title) {\n"
  },
  {
    "file": "public/app.js",
    "before": "}\n\nfunction getPendingRecordsForBookKey(bookKey, records = currentPendingRecords) {\n  if (bookKey.startsWith(REVIEW_VIDEO_EVENT_KEY)) {\n    const event = decodeURIComponent(bookKey.slice(REVIEW_VIDEO_EVENT_KEY.length));\n    return getPendingVideoRecords(records).filter(record => (record.sourceEvent || \"Unassigned event\") === event);\n",
    "after": "}\n\nfunction getPendingRecordsForBookKey(bookKey, records = currentPendingRecords) {\n  records = filterExcerptTaskRecords(records, excerptReviewTask);\n  if (bookKey === REVIEW_EXCERPT_TASK_KEY) return getPendingVideoRecords(records);\n  if (bookKey.startsWith(REVIEW_VIDEO_EVENT_KEY)) {\n    const event = decodeURIComponent(bookKey.slice(REVIEW_VIDEO_EVENT_KEY.length));\n    return getPendingVideoRecords(records).filter(record => (record.sourceEvent || \"Unassigned event\") === event);\n"
  },
  {
    "file": "public/app.js",
    "before": "}\n\nfunction getPendingVideoRecords(records = currentPendingRecords) {\n  return records.filter(isVideoReviewRecord);\n}\n\nasync function requestMergedBookRecords(sourceExcerpts) {\n",
    "after": "}\n\nfunction getPendingVideoRecords(records = currentPendingRecords) {\n  return filterExcerptTaskRecords(records, excerptReviewTask).filter(isVideoReviewRecord);\n}\n\nasync function requestMergedBookRecords(sourceExcerpts) {\n"
  },
  {
    "file": "public/app.js",
    "before": "}\n\nfunction applyPendingBookData(records, { preserveSelection = false } = {}) {\n  const previousWeirdSelection = preserveSelection ? elements.weirdBookSelect?.value || \"\" : \"\";\n\n  currentPendingRecords = records;\n",
    "after": "}\n\nfunction applyPendingBookData(records, { preserveSelection = false } = {}) {\n  records = filterExcerptTaskRecords(records, excerptReviewTask);\n  const previousWeirdSelection = preserveSelection ? elements.weirdBookSelect?.value || \"\" : \"\";\n\n  currentPendingRecords = records;\n"
  },
  {
    "file": "public/app.js",
    "before": "    const event = video.sourceEvent || \"Unassigned event\";\n    reviewBookSummaryByKey.set(`${REVIEW_VIDEO_EVENT_KEY}${encodeURIComponent(event)}`, { title: event });\n  }\n  weirdBookSummaryByKey = indexBookSummariesByKey(currentWeirdBookSummaries);\n\n  refreshReviewBookSelect(preserveSelection);\n",
    "after": "    const event = video.sourceEvent || \"Unassigned event\";\n    reviewBookSummaryByKey.set(`${REVIEW_VIDEO_EVENT_KEY}${encodeURIComponent(event)}`, { title: event });\n  }\n  if (excerptReviewTask) reviewBookSummaryByKey.set(REVIEW_EXCERPT_TASK_KEY, {title:\"this video excerpt assignment\"});\n  weirdBookSummaryByKey = indexBookSummariesByKey(currentWeirdBookSummaries);\n\n  refreshReviewBookSelect(preserveSelection);\n"
  },
  {
    "file": "public/app.js",
    "before": "    elements.weirdBookCountBadge.textContent = `${currentWeirdBookSummaries.length} Books`;\n  }\n\n  return allBookSummaries;\n}\n\n",
    "after": "    elements.weirdBookCountBadge.textContent = `${currentWeirdBookSummaries.length} Books`;\n  }\n\n  updateExcerptAssignmentUi();\n  return allBookSummaries;\n}\n\n"
  },
  {
    "file": "public/app.js",
    "before": "        <button type=\"button\" class=\"button button--secondary\" data-action=\"resolve-identity\">Save verified recording identity</button>\n        <p data-identity-state role=\"status\"></p>\n      </details>` : \"\"}\n      <button type=\"button\" class=\"button button--secondary\" data-action=\"inspect-receiver\">Check Poetry Please video and versions</button>\n      <div data-receiver-state aria-live=\"polite\"></div>\n      ${candidate.existingCanonicalVideoId && gate.decision === \"ready_for_poetry_please\" ? '<button type=\"button\" class=\"button button--secondary\" data-action=\"sync-context\">Sync saved reviews and selected excerpts</button><div data-context-state aria-live=\"polite\"></div>' : \"\"}\n",
    "after": "        <button type=\"button\" class=\"button button--secondary\" data-action=\"resolve-identity\">Save verified recording identity</button>\n        <p data-identity-state role=\"status\"></p>\n      </details>` : \"\"}\n      <p><a href=\"${escapeAttribute(buildExcerptReviewTaskUrl(window.location.origin, [candidate.sourceReviewFileId || candidate.sourceFileId]))}\" target=\"_blank\" rel=\"noopener noreferrer\">Review excerpts for this video</a></p>\n      <button type=\"button\" class=\"button button--secondary\" data-action=\"inspect-receiver\">Check Poetry Please video and versions</button>\n      <div data-receiver-state aria-live=\"polite\"></div>\n      ${candidate.existingCanonicalVideoId && gate.decision === \"ready_for_poetry_please\" ? '<button type=\"button\" class=\"button button--secondary\" data-action=\"sync-context\">Sync saved reviews and selected excerpts</button><div data-context-state aria-live=\"polite\"></div>' : \"\"}\n"
  },
  {
    "file": "public/app.js",
    "before": "      .catch(() => {\n        // Keep the queue usable even if validation refresh fails.\n      });\n  } catch (error) {\n    setStatus(`Book load failed: ${error.message}`);\n  }\n}\n\n",
    "after": "      .catch(() => {\n        // Keep the queue usable even if validation refresh fails.\n      });\n    return true;\n  } catch (error) {\n    setStatus(`Book load failed: ${error.message}`);\n    return false;\n  }\n}\n\n"
  },
  {
    "file": "public/app.js",
    "before": "    const bookTitle = summary?.title || bookKey;\n    setStatus(`Loading excerpts for \"${bookTitle}\"...`);\n    currentExcerpts = getPendingRecordsForBookKey(bookKey);\n    if (bookKey.startsWith(REVIEW_VIDEO_EVENT_KEY) && elements.reviewDisplayMode) {\n      elements.reviewDisplayMode.value = \"batch\";\n    }\n    reviewVisibleCount = getReviewBatchSize();\n",
    "after": "    const bookTitle = summary?.title || bookKey;\n    setStatus(`Loading excerpts for \"${bookTitle}\"...`);\n    currentExcerpts = getPendingRecordsForBookKey(bookKey);\n    if ((bookKey.startsWith(REVIEW_VIDEO_EVENT_KEY) || bookKey === REVIEW_EXCERPT_TASK_KEY) && elements.reviewDisplayMode) {\n      elements.reviewDisplayMode.value = \"batch\";\n    }\n    reviewVisibleCount = getReviewBatchSize();\n"
  },
  {
    "file": "public/app.js",
    "before": "  const contributorShell = currentContributorAccess || isContributorShellRequested();\n  setActiveModule(contributorShell ? \"gathering\" : \"review\");",
    "after": "  const contributorShell = currentContributorAccess || isContributorShellRequested();\n  if (contributorShellParams.has(\"task\")) {\n    try {\n      const task = parseExcerptReviewTask(window.location.search);\n      if (contributorShell) throw new Error(\"This contributor account is configured for intake. Sign in with a review-authorized account to approve excerpts.\");\n      await openExcerptReviewAssignment(task);\n    } catch (error) {\n      setActiveModule(\"review\");\n      const panel = document.getElementById(\"excerpt-assignment-panel\"); panel.hidden = false;\n      document.getElementById(\"excerpt-assignment-status\").textContent = error.message;\n      if (!excerptReviewTask) { elements.loadExcerpts.disabled = true; elements.loadBooks.disabled = true; }\n      setStatus(error.message);\n    }\n    return;\n  }\n  setActiveModule(contributorShell ? \"gathering\" : \"review\");"
  },
  {
    "file": "public/app.js",
    "before": "\ndocument.getElementById(\"video-media-resolution-refresh\")?.addEventListener(\"click\",loadVideoMediaResolutionQueue);\ndocument.getElementById(\"video-finished-scope-save\")?.addEventListener(\"click\",saveVideoFinishedScope);\n",
    "after": "\ndocument.getElementById(\"video-media-resolution-refresh\")?.addEventListener(\"click\",loadVideoMediaResolutionQueue);\ndocument.getElementById(\"video-finished-scope-save\")?.addEventListener(\"click\",saveVideoFinishedScope);\n\ndocument.getElementById(\"copy-excerpt-review-link\")?.addEventListener(\"click\", copyExcerptReviewLink);\n"
  },
  {
    "file": "public/index.html",
    "before": "      </section>\n\n      <section class=\"panel panel--wide module-panel module-panel--active\" id=\"review-module\">\n        <div class=\"panel__header\">\n          <h2>Book Selection</h2>\n          <span class=\"badge badge--muted\" id=\"book-count-badge\">0 Books</span>\n",
    "after": "      </section>\n\n      <section class=\"panel panel--wide module-panel module-panel--active\" id=\"review-module\">\n        <section id=\"excerpt-assignment-panel\" hidden aria-label=\"Excerpt review assignment\">\n          <h2>Excerpt review assignment</h2>\n          <p id=\"excerpt-assignment-status\" role=\"status\"></p>\n          <p>Opening this link saves no decisions. Sign in to submit approvals. <a href=\"/\">Leave assignment and open Weaver</a></p>\n        </section>\n        <div class=\"panel__header\">\n          <h2>Book Selection</h2>\n          <span class=\"badge badge--muted\" id=\"book-count-badge\">0 Books</span>\n"
  },
  {
    "file": "public/index.html",
    "before": "          </label>\n          <button id=\"load-books\" class=\"button button--secondary\">Refresh Books</button>\n          <button id=\"load-excerpts\" class=\"button\">Load Excerpts</button>\n        </div>\n        <p class=\"hint\" id=\"book-selection-hint\">\n          `Review queue` is the normal approval lane. Use the separate `Extra review` module for incomplete, conflicting, or otherwise weird rows that deserve their own pass.\n",
    "after": "          </label>\n          <button id=\"load-books\" class=\"button button--secondary\">Refresh Books</button>\n          <button id=\"load-excerpts\" class=\"button\">Load Excerpts</button>\n          <button id=\"copy-excerpt-review-link\" class=\"button button--secondary\" type=\"button\" hidden>Copy review link</button>\n          <input id=\"excerpt-review-task-url\" aria-label=\"Review assignment URL\" type=\"url\" readonly hidden>\n\n        </div>\n        <p class=\"hint\" id=\"book-selection-hint\">\n          `Review queue` is the normal approval lane. Use the separate `Extra review` module for incomplete, conflicting, or otherwise weird rows that deserve their own pass.\n"
  },
  {
    "file": "public/index.html",
    "before": "    </main>\n  </div>\n\n  <script src=\"/app.js?v=20261110-legacy-identity\"></script>\n</body>\n</html>\n",
    "after": "    </main>\n  </div>\n\n  <script src=\"/app.js?v=20261113-excerpt-task-links\"></script>\n</body>\n</html>\n"
  }
];
const files=new Map([...new Set(patches.map(p=>p.file))].map(f=>[f,fs.readFileSync(path.join(root,f),'utf8')])),errors=[];
for(const p of patches){const s=files.get(p.file);if(s.split(p.before).length!==2){errors.push('Changed excerpt-task patch target: '+p.file);continue;}files.set(p.file,s.replace(p.before,p.after));}
assert.equal(errors.length,0,errors.join('\n'));for(const[f,s]of files)fs.writeFileSync(path.join(root,f),s);
console.log('Applied guarded excerpt approval assignment links.');
