function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
    { name: "Copy Rows with Blank Cell in Column Q", functionName: "copyRowsWithBlankQ" },
    { name: "Check for typos", functionName: "checkForTypos" },
    { name: "Check excerpt duplicates", functionName: "checkExcerptDuplicates" },
    { name: "Mark exact duplicate delete candidates", functionName: "markExactDuplicateDeleteCandidates" },
    { name: "Delete next 100 exact duplicates", functionName: "deleteNext100ExactDuplicates" },
    { name: "Delete next 250 exact duplicates", functionName: "deleteNext250ExactDuplicates" },
    { name: "Refresh duplicate review columns", functionName: "refreshDuplicateReviewColumns" },
    { name: "Rebuild pull counts from backup", functionName: "rebuildExactPullCountsFromBackup" }
  ];
  sheet.addMenu("Custom Menu", menuEntries);
}

function copyRowsWithBlankQ() {
  var currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = currentSpreadsheet.getSheetByName("New - Quote Creation Tool Database");

  var destinationSpreadsheet = SpreadsheetApp.openById("16WDniM-dosOy0rDvCMWaLGVGWOs8-CGeY_K3tuB1ui4");
  var destinationSheet = destinationSpreadsheet.getSheetByName("Automatic Merged Import Template");

  var destinationLastRow = destinationSheet.getLastRow();
  if (destinationLastRow > 2) {
    destinationSheet.getRange(3, 1, destinationLastRow - 2, 16).clearContent();
  }

  var sourceData = currentSheet.getDataRange().getValues();
  var filteredRows = [];
  var sourceRowsToUpdate = [];

  for (var i = 1; i < sourceData.length; i++) {
    var valP = sourceData[i][15];
    var valQ = sourceData[i][16];

    if (valP === "Y" && valQ === "") {
      filteredRows.push(sourceData[i]);
      sourceRowsToUpdate.push(i + 1);
    }
  }

  if (filteredRows.length > 0) {
    destinationSheet
      .getRange(3, 1, filteredRows.length, filteredRows[0].length)
      .setValues(filteredRows);

    for (var j = 0; j < sourceRowsToUpdate.length; j++) {
      currentSheet.getRange(sourceRowsToUpdate[j], 17).setValue("RBT");
    }
  } else {
    Logger.log('No rows with "Y" in P and a blank cell in Q found.');
  }
}

const SPANISH_WORDS = [
  "miercoles", "cuchara", "oro", "mexican", "chola", "femme", "poem", "anti", "lola", "muerto", "cielo",
  "de", "la", "el", "y", "con"
];

const WHITELIST = [
  "i", "poemed"
];

function rowHasSpanishWord(text) {
  let words = (text || "").toLowerCase().split(/\s+/);
  return words.some(word => SPANISH_WORDS.includes(word));
}

function checkForTypos() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const colG = 7;
  const colR = 18, colS = 19, colT = 20, colU = 21, colV = 22;

  sheet.getRange(1, colR).setValue("Typos (Y/N)?");
  sheet.getRange(1, colS).setValue("Where?");
  sheet.getRange(1, colU).setValue("Spanish (Y/N)?");
  sheet.getRange(1, colV).setValue("Misspelled Word");

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  for (let i = 0; i < data.length; i++) {
    const tVal = (data[i][colT - 1] || "").toString().trim().toUpperCase();
    const rVal = (data[i][colR - 1] || "").toString().trim().toUpperCase();
    if (tVal !== "Y" || rVal === "Y" || rVal === "N") continue;

    let gVal = data[i][colG - 1];
    let gText = (gVal || "").toString().trim();

    if (rowHasSpanishWord(gText)) {
      sheet.getRange(i + 2, colU, 1, 1).setValue("Y");
      continue;
    } else {
      sheet.getRange(i + 2, colU, 1, 1).setValue("N");
    }

    let result = checkTyposAndReturnWords(gText);
    let typoFound = result.typo;
    let typoWords = result.words;

    sheet.getRange(i + 2, colR, 1, 1).setValue(typoFound ? "Y" : "N");
    sheet.getRange(i + 2, colS, 1, 1).setValue(typoFound ? "Title" : "");
    sheet.getRange(i + 2, colV, 1, 1).setValue(typoWords.join(", "));
    Utilities.sleep(120);
  }
}

function checkTyposAndReturnWords(text) {
  const url = "https://api.languagetool.org/v2/check";
  const payload = {
    text: text,
    language: "en-US"
  };
  const options = {
    method: "post",
    payload: payload,
    muteHttpExceptions: true,
    followRedirects: true
  };
  let words = [];
  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    if (result.matches && result.matches.length) {
      for (let match of result.matches) {
        if (match.rule.issueType === "misspelling") {
          let word = text.substring(match.offset, match.offset + match.length).toLowerCase();
          if (WHITELIST.indexOf(word) === -1 && words.indexOf(word) === -1) {
            words.push(word);
          }
        }
      }
    }
    return { typo: words.length > 0, words: words };
  } catch (e) {
    Logger.log("Spellcheck API error: " + e);
    return { typo: false, words: [] };
  }
}

const EXCERPT_DUPLICATE_CONFIG = {
  sheetName: "New - Quote Creation Tool Database",
  poemTitleColumn: 4,
  excerptColumn: 7,
  startRow: 2,
  outputStartColumn: 23,
  deleteReviewStartColumn: 30,
  exactPullCountColumn: 35,
  recoveredPullCountColumn: 36,
  backupSheetName: "Exact Duplicate Delete Backup",
  debugSheetName: "Exact Pull Count Debug",
  deleteAuditSheetName: "Exact Duplicate Delete Audit",
  nearDuplicateThreshold: 0.72,
  previewLength: 140,
  minTokenLength: 3,
  maxFingerprintTokens: 6,
  maxCandidatesPerRow: 40,
  deleteBatchSize: 100
};

function checkExcerptDuplicates() {
  const config = EXCERPT_DUPLICATE_CONFIG;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = config.sheetName ? spreadsheet.getSheetByName(config.sheetName) : spreadsheet.getActiveSheet();

  if (!sheet) {
    throw new Error('Sheet "' + config.sheetName + '" not found.');
  }

  spreadsheet.toast('Starting duplicate check on "' + sheet.getName() + '"...', "Excerpt Check", 5);

  const lastRow = sheet.getLastRow();
  if (lastRow < config.startRow) {
    spreadsheet.toast("No data rows found to analyze.", "Excerpt Check", 5);
    return;
  }

  const rowCount = lastRow - config.startRow + 1;
  const existingPullCountColumn = config.deleteReviewStartColumn + 5;
  const dataWidth = Math.max(config.poemTitleColumn, config.excerptColumn, existingPullCountColumn);
  const sheetValues = sheet.getRange(config.startRow, 1, rowCount, dataWidth).getValues();
  sheet.getRange(config.startRow, config.outputStartColumn, rowCount, 7).clearContent();

  const headers = [[
    "Excerpt Word Count",
    "Excerpt Character Count",
    "Duplicate Status",
    "Matched Row Number",
    "Overlap Score",
    "Shared Token Count",
    "Matched Excerpt Preview"
  ]];
  sheet.getRange(1, config.outputStartColumn, 1, headers[0].length).setValues(headers);

  const rows = sheetValues.map(function(row, index) {
    const originalText = (row[config.excerptColumn - 1] || "").toString();
    const rawText = cleanWhitespace_(originalText);
    const poemTitle = cleanWhitespace_(row[config.poemTitleColumn - 1]);
    const tokenList = tokenizeExcerpt_(rawText);
    return {
      sheetRow: config.startRow + index,
      poemTitle: poemTitle,
      originalText: originalText,
      rawText: rawText,
      normalizedText: normalizeExcerpt_(rawText),
      tokenList: tokenList,
      tokenSet: tokenSetFromList_(tokenList),
      wordCount: rawText ? rawText.split(/\s+/).length : 0,
      characterCount: rawText.length
    };
  });

  const nonBlankRows = rows.filter(function(row) {
    return !!row.rawText;
  });

  if (!nonBlankRows.length) {
    spreadsheet.toast(
      "No excerpt text found in column " + columnNumberToLetter_(config.excerptColumn) + ".",
      "Excerpt Check",
      7
    );
    return;
  }

  const matches = computeExcerptMatches_(rows, config.nearDuplicateThreshold);

  const output = rows.map(function(row) {
    const match = matches[row.sheetRow];
    return [
      row.wordCount,
      row.characterCount,
      match ? match.matchType : "",
      match ? match.otherRowNumber : "",
      match ? roundScore_(match.score) : "",
      match ? match.sharedTokenCount : "",
      match ? truncateText_(match.otherExcerpt, config.previewLength) : ""
    ];
  });

  sheet.getRange(config.startRow, config.outputStartColumn, output.length, headers[0].length).setValues(output);
  spreadsheet.toast(
    "Duplicate check finished. Reviewed " + nonBlankRows.length +
      " excerpts across " + countDistinctTitles_(nonBlankRows) + " poem titles.",
    "Excerpt Check",
    7
  );
}

function computeExcerptMatches_(rows, threshold) {
  const matches = {};
  const poemGroups = {};

  rows.forEach(function(row) {
    if (!row.rawText) return;
    const poemKey = row.poemTitle || "__NO_POEM_TITLE__";
    if (!poemGroups[poemKey]) {
      poemGroups[poemKey] = [];
    }
    poemGroups[poemKey].push(row);
  });

  Object.keys(poemGroups).forEach(function(poemKey) {
    const poemRows = poemGroups[poemKey];
    const exactGroups = {};
    const fingerprintGroups = {};

    poemRows.forEach(function(row) {
      if (!row.normalizedText) return;
      if (!exactGroups[row.normalizedText]) {
        exactGroups[row.normalizedText] = [];
      }
      exactGroups[row.normalizedText].push(row);

      row.fingerprintKeys = buildFingerprintKeys_(row, EXCERPT_DUPLICATE_CONFIG);
      row.fingerprintKeys.forEach(function(key) {
        if (!fingerprintGroups[key]) {
          fingerprintGroups[key] = [];
        }
        fingerprintGroups[key].push(row);
      });
    });

    Object.keys(exactGroups).forEach(function(key) {
      const group = exactGroups[key];
      if (group.length < 2) return;

      group.forEach(function(row, index) {
        const other = group[index === 0 ? 1 : 0];
        matches[row.sheetRow] = {
          matchType: "exact_duplicate",
          otherRowNumber: other.sheetRow,
          score: 1,
          sharedTokenCount: row.tokenSet.size,
          otherExcerpt: other.rawText
        };
      });
    });

    poemRows.forEach(function(current) {
      if (!current.fingerprintKeys.length) return;

      const candidateMap = {};
      current.fingerprintKeys.forEach(function(key) {
        const candidates = fingerprintGroups[key] || [];
        candidates.forEach(function(other) {
          if (other.sheetRow === current.sheetRow) return;
          if (other.normalizedText === current.normalizedText) return;
          candidateMap[other.sheetRow] = other;
        });
      });

      const candidates = Object.keys(candidateMap)
        .map(function(rowNumber) {
          return candidateMap[rowNumber];
        })
        .sort(function(a, b) {
          return b.tokenSet.size - a.tokenSet.size;
        })
        .slice(0, EXCERPT_DUPLICATE_CONFIG.maxCandidatesPerRow);

      candidates.forEach(function(other) {
        if (current.sheetRow >= other.sheetRow) return;

        const tokenData = tokenOverlap_(current.tokenSet, other.tokenSet);
        if (tokenData.sharedCount < 2) return;

        const phraseScore = phraseOverlapScore_(current.tokenList, other.tokenList);
        const combinedScore = Math.max(tokenData.score, phraseScore);

        if (combinedScore < threshold) return;

        storeBestMatch_(matches, current, other, combinedScore, tokenData.sharedCount);
        storeBestMatch_(matches, other, current, combinedScore, tokenData.sharedCount);
      })
    });
  });

  return matches;
}

function storeBestMatch_(matches, row, other, score, sharedCount) {
  const existing = matches[row.sheetRow];
  if (existing && existing.score >= score) return;

  matches[row.sheetRow] = {
    matchType: "high_overlap",
    otherRowNumber: other.sheetRow,
    score: score,
    sharedTokenCount: sharedCount,
    otherExcerpt: other.rawText
  };
}

function cleanWhitespace_(text) {
  return (text || "").toString().replace(/\s+/g, " ").trim();
}

function normalizeExcerpt_(text) {
  return cleanWhitespace_(text)
    .toLowerCase()
    .replace(/[—–]/g, " ")
    .replace(/[""'“”‘’`.,!?;:()[\]{}]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function tokenizeExcerpt_(text) {
  const normalized = normalizeExcerpt_(text);
  const matches = normalized.match(/[a-z0-9]+/g);
  return matches ? matches : [];
}

function buildFingerprintKeys_(row, config) {
  const filteredTokens = row.tokenList.filter(function(token) {
    return token.length >= config.minTokenLength;
  });

  if (!filteredTokens.length) return [];

  const counts = {};
  filteredTokens.forEach(function(token) {
    counts[token] = (counts[token] || 0) + 1;
  });

  const limitedTokens = Object.keys(counts)
    .sort(function(a, b) {
      if (counts[b] !== counts[a]) return counts[b] - counts[a];
      if (a.length !== b.length) return b.length - a.length;
      return a.localeCompare(b);
    })
    .slice(0, config.maxFingerprintTokens);
  const keys = [];

  for (let i = 0; i < limitedTokens.length; i++) {
    for (let j = i + 1; j < limitedTokens.length; j++) {
      keys.push(limitedTokens[i] + "|" + limitedTokens[j]);
    }
  }

  if (!keys.length && limitedTokens.length === 1) {
    keys.push(limitedTokens[0]);
  }

  return keys;
}

function tokenSetFromList_(tokenList) {
  const tokenMap = {};
  tokenList.forEach(function(token) {
    tokenMap[token] = true;
  });
  return {
    values: Object.keys(tokenMap),
    size: Object.keys(tokenMap).length,
    has: function(token) {
      return !!tokenMap[token];
    }
  };
}

function tokenOverlap_(tokenSetA, tokenSetB) {
  if (!tokenSetA.size || !tokenSetB.size) {
    return { score: 0, sharedCount: 0 };
  }

  let sharedCount = 0;
  tokenSetA.values.forEach(function(token) {
    if (tokenSetB.has(token)) sharedCount++;
  });

  const unionCount = tokenSetA.size + tokenSetB.size - sharedCount;
  return {
    score: unionCount ? sharedCount / unionCount : 0,
    sharedCount: sharedCount
  };
}

function phraseOverlapScore_(tokensA, tokensB) {
  if (!tokensA.length || !tokensB.length) return 0;

  const phrasesA = buildNgrams_(tokensA, 3);
  const phrasesB = buildNgrams_(tokensB, 3);
  const phraseOverlap = tokenOverlap_(phrasesA, phrasesB).score;

  if (phraseOverlap > 0) return phraseOverlap;

  const bigramsA = buildNgrams_(tokensA, 2);
  const bigramsB = buildNgrams_(tokensB, 2);
  return tokenOverlap_(bigramsA, bigramsB).score;
}

function roundScore_(score) {
  return Math.round(score * 1000) / 1000;
}

function truncateText_(text, maxLength) {
  if (!text || text.length <= maxLength) return text;
  return text.slice(0, maxLength - 3) + "...";
}

function buildNgrams_(tokens, size) {
  if (tokens.length < size) {
    return tokenSetFromList_([]);
  }

  const ngrams = [];
  for (let i = 0; i <= tokens.length - size; i++) {
    ngrams.push(tokens.slice(i, i + size).join(" "));
  }
  return tokenSetFromList_(ngrams);
}

function columnNumberToLetter_(column) {
  let letter = "";
  let current = column;

  while (current > 0) {
    const remainder = (current - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    current = Math.floor((current - 1) / 26);
  }

  return letter;
}

function countDistinctTitles_(rows) {
  const titles = {};
  rows.forEach(function(row) {
    titles[row.poemTitle || "__NO_POEM_TITLE__"] = true;
  });
  return Object.keys(titles).length;
}

function markExactDuplicateDeleteCandidates() {
  const config = EXCERPT_DUPLICATE_CONFIG;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = config.sheetName ? spreadsheet.getSheetByName(config.sheetName) : spreadsheet.getActiveSheet();

  if (!sheet) {
    throw new Error('Sheet "' + config.sheetName + '" not found.');
  }

  spreadsheet.toast('Marking exact duplicate delete candidates on "' + sheet.getName() + '"...', "Duplicate Review", 5);

  const lastRow = sheet.getLastRow();
  if (lastRow < config.startRow) {
    spreadsheet.toast("No data rows found to analyze.", "Duplicate Review", 5);
    return;
  }

  const rowCount = lastRow - config.startRow + 1;
  const existingPullCountColumn = config.exactPullCountColumn;
  const dataWidth = Math.max(config.poemTitleColumn, config.excerptColumn, existingPullCountColumn);
  const sheetValues = sheet.getRange(config.startRow, 1, rowCount, dataWidth).getValues();

  const reviewHeaders = [[
    "Delete Candidate?",
    "Keep/Delete Decision",
    "Duplicate Group ID",
    "Keep Row Number",
    "Decision Reason"
  ]];
  sheet.getRange(1, config.deleteReviewStartColumn, 1, reviewHeaders[0].length).setValues(reviewHeaders);
  sheet.getRange(config.startRow, config.deleteReviewStartColumn, rowCount, reviewHeaders[0].length).clearContent();

  const rows = sheetValues.map(function(row, index) {
    const originalText = (row[config.excerptColumn - 1] || "").toString();
    const rawText = cleanWhitespace_(originalText);
    const poemTitle = cleanWhitespace_(row[config.poemTitleColumn - 1]);
    const tokenList = tokenizeExcerpt_(rawText);
    return {
      sheetRow: config.startRow + index,
      poemTitle: poemTitle,
      originalText: originalText,
      rawText: rawText,
      normalizedText: normalizeExcerpt_(rawText),
      tokenList: tokenList,
      tokenSet: tokenSetFromList_(tokenList),
      wordCount: rawText ? rawText.split(/\s+/).length : 0,
      characterCount: rawText.length,
      existingPullCount: parsePullCount_(row[existingPullCountColumn - 1])
    };
  });

  const exactGroups = buildExactDuplicateGroups_(rows);
  const outputMap = {};
  let groupNumber = 0;
  let candidateCount = 0;

  Object.keys(exactGroups).forEach(function(groupKey) {
    const group = exactGroups[groupKey];
    if (group.length < 2) return;

    groupNumber++;
    const keepRow = chooseExactDuplicateWinner_(group);
    const groupId = "ED-" + groupNumber;

    group.forEach(function(row) {
      if (row.sheetRow === keepRow.sheetRow) {
        outputMap[row.sheetRow] = [
          "",
          "KEEP",
          groupId,
          keepRow.sheetRow,
          "Preferred exact duplicate; preserving line breaks first, then shorter text"
        ];
      } else {
        candidateCount++;
        outputMap[row.sheetRow] = [
          "Y",
          "DELETE_CANDIDATE",
          groupId,
          keepRow.sheetRow,
          "Exact duplicate; merged into keep row"
        ];
      }
    });
  });

  const output = rows.map(function(row) {
    if (outputMap[row.sheetRow]) {
      return outputMap[row.sheetRow];
    }
    return ["", "", "", "", ""];
  });

  sheet.getRange(config.startRow, config.deleteReviewStartColumn, output.length, reviewHeaders[0].length).setValues(output);
  spreadsheet.toast(
    "Marked " + candidateCount + " exact duplicate delete candidates across " + groupNumber + " groups.",
    "Duplicate Review",
    7
  );
}

function buildExactDuplicateGroups_(rows) {
  const groups = {};

  rows.forEach(function(row) {
    if (!row.rawText || !row.normalizedText) return;
    const poemKey = row.poemTitle || "__NO_POEM_TITLE__";
    const groupKey = poemKey + "||" + row.normalizedText;
    if (!groups[groupKey]) {
      groups[groupKey] = [];
    }
    groups[groupKey].push(row);
  });

  return groups;
}

function chooseExactDuplicateWinner_(group) {
  return group.slice().sort(function(a, b) {
    const aHasLineBreaks = hasLineBreaks_(a.originalText);
    const bHasLineBreaks = hasLineBreaks_(b.originalText);
    if (aHasLineBreaks !== bHasLineBreaks) return aHasLineBreaks ? -1 : 1;
    if (a.characterCount !== b.characterCount) return a.characterCount - b.characterCount;
    if (a.wordCount !== b.wordCount) return a.wordCount - b.wordCount;
    return a.sheetRow - b.sheetRow;
  })[0];
}

function hasLineBreaks_(text) {
  return /\r|\n/.test(text || "");
}

function parsePullCount_(value) {
  const parsed = parseInt(value, 10);
  return isNaN(parsed) ? 0 : parsed;
}

function deleteNext100ExactDuplicates() {
  deleteMarkedExactDuplicatesBatch_(100);
}

function deleteNext250ExactDuplicates() {
  deleteMarkedExactDuplicatesBatch_(250);
}

function deleteMarkedExactDuplicatesBatch_(batchSize) {
  const config = EXCERPT_DUPLICATE_CONFIG;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = config.sheetName ? spreadsheet.getSheetByName(config.sheetName) : spreadsheet.getActiveSheet();

  if (!sheet) {
    throw new Error('Sheet "' + config.sheetName + '" not found.');
  }

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Delete next batch of exact duplicates?",
    "This will back up and delete up to " + batchSize +
      " rows currently marked DELETE_CANDIDATE in \"" + sheet.getName() + "\".",
    ui.ButtonSet.OK_CANCEL
  );

  if (response !== ui.Button.OK) {
    spreadsheet.toast("Deletion cancelled.", "Duplicate Review", 5);
    return;
  }

  spreadsheet.toast("Preparing exact duplicate deletion...", "Duplicate Review", 5);

  const lastRow = sheet.getLastRow();
  if (lastRow < config.startRow) {
    spreadsheet.toast("No data rows found to analyze.", "Duplicate Review", 5);
    return;
  }

  const rowCount = lastRow - config.startRow + 1;
  const totalColumns = sheet.getLastColumn();
  const allValues = sheet.getRange(config.startRow, 1, rowCount, totalColumns).getValues();
  const reviewOffset = config.deleteReviewStartColumn - 1;
  const rowsToDelete = [];
  const beforeCandidateCount = countDeleteCandidates_(allValues, reviewOffset);
  const beforeExcerptCount = countNonBlankExcerptRows_(allValues, config.excerptColumn - 1);

  allValues.forEach(function(row, index) {
    const deleteCandidate = (row[reviewOffset] || "").toString().trim().toUpperCase();
    const decision = (row[reviewOffset + 1] || "").toString().trim().toUpperCase();
    if (deleteCandidate === "Y" && decision === "DELETE_CANDIDATE") {
      rowsToDelete.push({
        sheetRow: config.startRow + index,
        values: row
      });
    }
  });

  if (!rowsToDelete.length) {
    spreadsheet.toast("No marked exact duplicate rows found to delete.", "Duplicate Review", 7);
    return;
  }

  const batchRowsToDelete = rowsToDelete
    .sort(function(a, b) {
      return a.sheetRow - b.sheetRow;
    })
    .slice(0, batchSize);

  const beforeLastRow = sheet.getLastRow();
  backupRowsBeforeDelete_(spreadsheet, sheet, batchRowsToDelete, totalColumns, config.backupSheetName);
  deleteRowsInBatches_(sheet, batchRowsToDelete.map(function(rowInfo) {
    return rowInfo.sheetRow;
  }));
  SpreadsheetApp.flush();
  const afterLastRow = sheet.getLastRow();
  const actualDeleted = beforeLastRow - afterLastRow;
  const remainingCandidates = rowsToDelete.length - batchRowsToDelete.length;

  const afterValues = sheet.getRange(
    config.startRow,
    1,
    Math.max(sheet.getLastRow() - config.startRow + 1, 0),
    Math.min(sheet.getLastColumn(), Math.max(totalColumns, config.deleteReviewStartColumn + 4))
  ).getValues();
  const afterCandidateCount = countDeleteCandidates_(afterValues, reviewOffset);
  const afterExcerptCount = countNonBlankExcerptRows_(afterValues, config.excerptColumn - 1);

  writeDeleteAuditRow_(
    spreadsheet,
    config.deleteAuditSheetName,
    {
      timestamp: new Date(),
      batchSize: batchSize,
      requestedDeleteCount: batchRowsToDelete.length,
      actualDeleted: actualDeleted,
      beforeLastRow: beforeLastRow,
      afterLastRow: afterLastRow,
      beforeCandidateCount: beforeCandidateCount,
      afterCandidateCount: afterCandidateCount,
      beforeExcerptCount: beforeExcerptCount,
      afterExcerptCount: afterExcerptCount,
      firstDeletedRow: batchRowsToDelete.length ? batchRowsToDelete[0].sheetRow : "",
      lastDeletedRow: batchRowsToDelete.length ? batchRowsToDelete[batchRowsToDelete.length - 1].sheetRow : "",
      sampleTitles: batchRowsToDelete.slice(0, 5).map(function(rowInfo) {
        return cleanWhitespace_(rowInfo.values[config.poemTitleColumn - 1]);
      }).join(" | ")
    }
  );

  spreadsheet.toast(
    "Backed up and deleted " + actualDeleted +
      " rows. Candidates before/after: " + beforeCandidateCount + "/" + afterCandidateCount +
      ". Run \"Refresh duplicate review columns\" next.",
    "Duplicate Review",
    7
  );
}

function refreshDuplicateReviewColumns() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.toast("Refreshing duplicate review columns...", "Duplicate Review", 5);
  checkExcerptDuplicates();
  markExactDuplicateDeleteCandidates();
  rebuildExactPullCountsFromBackup();
  spreadsheet.toast("Duplicate review columns refreshed.", "Duplicate Review", 7);
}

function backupRowsBeforeDelete_(spreadsheet, sourceSheet, rowsToDelete, totalColumns, backupSheetName) {
  let backupSheet = spreadsheet.getSheetByName(backupSheetName);
  if (!backupSheet) {
    backupSheet = spreadsheet.insertSheet(backupSheetName);
  }

  const backupHeader = [
    "Backup Timestamp",
    "Source Sheet",
    "Original Row Number"
  ];
  const sourceHeaders = sourceSheet.getRange(1, 1, 1, totalColumns).getValues()[0];

  if (backupSheet.getLastRow() === 0) {
    backupSheet
      .getRange(1, 1, 1, backupHeader.length + sourceHeaders.length)
      .setValues([backupHeader.concat(sourceHeaders)]);
  }

  const timestamp = new Date();
  const backupValues = rowsToDelete.map(function(rowInfo) {
    return [timestamp, sourceSheet.getName(), rowInfo.sheetRow].concat(rowInfo.values);
  });

  backupSheet
    .getRange(backupSheet.getLastRow() + 1, 1, backupValues.length, backupValues[0].length)
    .setValues(backupValues);
}

function rebuildExactPullCountsFromBackup() {
  const config = EXCERPT_DUPLICATE_CONFIG;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = config.sheetName ? spreadsheet.getSheetByName(config.sheetName) : spreadsheet.getActiveSheet();
  const backupSheet = spreadsheet.getSheetByName(config.backupSheetName);

  if (!mainSheet) {
    throw new Error('Sheet "' + config.sheetName + '" not found.');
  }

  if (!backupSheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Backup sheet "' + config.backupSheetName + '" not found.',
      "Duplicate Review",
      7
    );
    return;
  }

  const mainLastRow = mainSheet.getLastRow();
  if (mainLastRow < config.startRow) {
    spreadsheet.toast("No data rows found on the main sheet.", "Duplicate Review", 7);
    return;
  }

  const mainRowCount = mainLastRow - config.startRow + 1;
  const mainDataWidth = Math.max(config.poemTitleColumn, config.excerptColumn, config.deleteReviewStartColumn + 5);
  const mainValues = mainSheet.getRange(config.startRow, 1, mainRowCount, mainDataWidth).getValues();

  const backupLastRow = backupSheet.getLastRow();
  if (backupLastRow < 2) {
    spreadsheet.toast("Backup sheet is empty.", "Duplicate Review", 7);
    return;
  }

  const backupLastColumn = backupSheet.getLastColumn();
  const backupHeaders = backupSheet.getRange(1, 1, 1, backupLastColumn).getValues()[0];
  const backupValues = backupSheet.getRange(2, 1, backupLastRow - 1, backupLastColumn).getValues();

  let poemTitleIndex = 3 + (config.poemTitleColumn - 1);
  let excerptIndex = 3 + (config.excerptColumn - 1);
  const backupDecisionIndex = config.deleteReviewStartColumn - 1 + 3;
  const backupPullCountIndex = config.deleteReviewStartColumn - 1 + 5 + 3;

  if (poemTitleIndex >= backupLastColumn || excerptIndex >= backupLastColumn) {
    const mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    const poemTitleHeader = mainHeaders[config.poemTitleColumn - 1];
    const excerptHeader = mainHeaders[config.excerptColumn - 1];
    poemTitleIndex = backupHeaders.indexOf(poemTitleHeader);
    excerptIndex = backupHeaders.indexOf(excerptHeader);
  }

  if (poemTitleIndex === -1 || excerptIndex === -1) {
    spreadsheet.toast(
      "Could not locate poem title or excerpt columns in the backup tab.",
      "Duplicate Review",
      7
    );
    return;
  }

  const countsByGroup = {};
  const currentRowsByGroup = {};
  const debugRows = [[
    "Poem Title",
    "Normalized Excerpt",
    "Current Rows",
    "Backup Rows",
    "Recovered Count",
    "Matched?"
  ]];

  mainValues.forEach(function(row) {
    const poemTitle = cleanWhitespace_(row[config.poemTitleColumn - 1]);
    const excerptText = cleanWhitespace_(row[config.excerptColumn - 1]);
    const normalizedText = normalizeExcerpt_(excerptText);
    if (!normalizedText) return;

    const key = buildExactGroupKey_(poemTitle, normalizedText);
    countsByGroup[key] = (countsByGroup[key] || 0) + 1;
    currentRowsByGroup[key] = (currentRowsByGroup[key] || 0) + 1;
  });

  backupValues.forEach(function(row) {
    const poemTitle = cleanWhitespace_(row[poemTitleIndex]);
    const excerptText = cleanWhitespace_(row[excerptIndex]);
    const normalizedText = normalizeExcerpt_(excerptText);
    if (!normalizedText) return;

    const key = buildExactGroupKey_(poemTitle, normalizedText);
    const backupPullCount = backupPullCountIndex < row.length ? parsePullCount_(row[backupPullCountIndex]) : 0;
    const decision = backupDecisionIndex < row.length ? (row[backupDecisionIndex] || "").toString().trim().toUpperCase() : "";

    if (backupPullCount > 1) {
      countsByGroup[key] = Math.max(countsByGroup[key] || 0, backupPullCount);
    } else if (decision === "DELETE_CANDIDATE") {
      countsByGroup[key] = (countsByGroup[key] || 0) + 1;
    } else {
      countsByGroup[key] = (countsByGroup[key] || 0) + 1;
    }
  });

  const pullCountColumn = config.exactPullCountColumn;
  const recoveredPullCountColumn = config.recoveredPullCountColumn;
  const output = mainValues.map(function(row) {
    const poemTitle = cleanWhitespace_(row[config.poemTitleColumn - 1]);
    const excerptText = cleanWhitespace_(row[config.excerptColumn - 1]);
    const normalizedText = normalizeExcerpt_(excerptText);
    if (!normalizedText) return [""];

    const key = buildExactGroupKey_(poemTitle, normalizedText);
    const count = countsByGroup[key] && countsByGroup[key] > 1 ? countsByGroup[key] : "";
    return [count];
  });

  mainSheet.getRange(1, pullCountColumn).setValue("Exact Pull Count");
  mainSheet.getRange(config.startRow, pullCountColumn, output.length, 1).setValues(output);
  mainSheet.getRange(1, recoveredPullCountColumn).setValue("Recovered Exact Pull Count");
  mainSheet.getRange(config.startRow, recoveredPullCountColumn, output.length, 1).setValues(output);

  const writtenCount = output.filter(function(row) {
    return row[0] !== "";
  }).length;
  const backupMatchedGroups = Object.keys(countsByGroup).filter(function(key) {
    return currentRowsByGroup[key] && countsByGroup[key] > 1;
  });
  const recoveredGroups = Object.keys(countsByGroup).filter(function(key) {
    return countsByGroup[key] > 1;
  });

  recoveredGroups.slice(0, 200).forEach(function(key) {
    const parts = key.split("||");
    debugRows.push([
      parts[0],
      parts.slice(1).join("||"),
      currentRowsByGroup[key] || 0,
      Math.max((countsByGroup[key] || 0) - (currentRowsByGroup[key] || 0), 0),
      countsByGroup[key] || 0,
      currentRowsByGroup[key] ? "Y" : "N"
    ]);
  });

  writeDebugSheet_(spreadsheet, config.debugSheetName, debugRows);

  spreadsheet.toast(
    "Rebuilt pull counts for " + writtenCount + " rows. Matched " + backupMatchedGroups.length +
      " current groups. See \"" + config.debugSheetName + "\".",
    "Duplicate Review",
    7
  );
}

function buildExactGroupKey_(poemTitle, normalizedText) {
  return (poemTitle || "__NO_POEM_TITLE__") + "||" + normalizedText;
}

function writeDebugSheet_(spreadsheet, sheetName, rows) {
  let debugSheet = spreadsheet.getSheetByName(sheetName);
  if (!debugSheet) {
    debugSheet = spreadsheet.insertSheet(sheetName);
  } else {
    debugSheet.clearContents();
  }

  debugSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}

function deleteRowsInBatches_(sheet, rowNumbers) {
  const sorted = rowNumbers.slice().sort(function(a, b) {
    return b - a;
  });

  let runStart = null;
  let runLength = 0;

  sorted.forEach(function(rowNumber, index) {
    if (runStart === null) {
      runStart = rowNumber;
      runLength = 1;
    } else if (rowNumber === runStart - runLength) {
      runLength++;
    } else {
      sheet.deleteRows(runStart - runLength + 1, runLength);
      runStart = rowNumber;
      runLength = 1;
    }

    if (index === sorted.length - 1 && runStart !== null) {
      sheet.deleteRows(runStart - runLength + 1, runLength);
    }
  });
}

function countDeleteCandidates_(values, reviewOffset) {
  let count = 0;
  values.forEach(function(row) {
    const deleteCandidate = (row[reviewOffset] || "").toString().trim().toUpperCase();
    const decision = (row[reviewOffset + 1] || "").toString().trim().toUpperCase();
    if (deleteCandidate === "Y" && decision === "DELETE_CANDIDATE") {
      count++;
    }
  });
  return count;
}

function countNonBlankExcerptRows_(values, excerptOffset) {
  let count = 0;
  values.forEach(function(row) {
    if (cleanWhitespace_(row[excerptOffset])) {
      count++;
    }
  });
  return count;
}

function writeDeleteAuditRow_(spreadsheet, sheetName, audit) {
  let auditSheet = spreadsheet.getSheetByName(sheetName);
  if (!auditSheet) {
    auditSheet = spreadsheet.insertSheet(sheetName);
    auditSheet.getRange(1, 1, 1, 12).setValues([[
      "Timestamp",
      "Batch Size",
      "Requested Delete Count",
      "Actual Deleted",
      "Before Last Row",
      "After Last Row",
      "Before Candidate Count",
      "After Candidate Count",
      "Before Excerpt Count",
      "After Excerpt Count",
      "Deleted Row Span",
      "Sample Titles"
    ]]);
  }

  auditSheet.getRange(auditSheet.getLastRow() + 1, 1, 1, 12).setValues([[
    audit.timestamp,
    audit.batchSize,
    audit.requestedDeleteCount,
    audit.actualDeleted,
    audit.beforeLastRow,
    audit.afterLastRow,
    audit.beforeCandidateCount,
    audit.afterCandidateCount,
    audit.beforeExcerptCount,
    audit.afterExcerptCount,
    audit.firstDeletedRow + " - " + audit.lastDeletedRow,
    audit.sampleTitles
  ]]);
}
