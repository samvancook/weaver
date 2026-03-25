function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
    { name: "Copy Rows with Blank Cell in Column Q", functionName: "copyRowsWithBlankQ" },
    { name: "Check for typos", functionName: "checkForTypos" },
    { name: "Check excerpt duplicates", functionName: "checkExcerptDuplicates" },
    { name: "Mark exact duplicate delete candidates", functionName: "markExactDuplicateDeleteCandidates" },
    { name: "Sync exact-duplicate excludes upstream", functionName: "syncExactDuplicateExclusionsToSource" },
    { name: "Refresh duplicate review columns", functionName: "refreshDuplicateReviewColumns" }
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
  sourceSheetName: "Excerpt Tool 1.20",
  poemTitleColumn: 4,
  excerptColumn: 7,
  startRow: 2,
  outputStartColumn: 23,
  deleteReviewStartColumn: 30,
  exactPullCountColumn: 35,
  recoveredPullCountColumn: 36,
  supportStartColumn: 37,
  backupSheetName: "Exact Duplicate Delete Backup",
  debugSheetName: "Exact Pull Count Debug",
  deleteAuditSheetName: "Exact Duplicate Delete Audit",
  previewSheetName: "Quote DB Exclude Preview",
  nearDuplicateThreshold: 0.72,
  previewLength: 140,
  minTokenLength: 3,
  maxFingerprintTokens: 6,
  maxCandidatesPerRow: 40,
  deleteBatchSize: 100
};

const SOURCE_WORKFLOW_COLUMNS = [
  { column: 25, header: "record_id" },
  { column: 26, header: "source_author" },
  { column: 27, header: "source_title" },
  { column: 28, header: "source_excerpt" },
  { column: 29, header: "source_book_title" },
  { column: 30, header: "exclude_from_quote_db" },
  { column: 31, header: "exclude_reason" },
  { column: 32, header: "duplicate_group_id" },
  { column: 33, header: "duplicate_keep_record_id" },
  { column: 34, header: "exact_pull_count" }
];

const SOURCE_REVIEW_STATE_COLUMNS = [
  { column: 35, header: "approved_for_quote" },
  { column: 36, header: "status_indicator" },
  { column: 37, header: "quote_created_qc" },
  { column: 38, header: "added_to_primary_excerpt_db" },
  { column: 39, header: "typos_yn" },
  { column: 40, header: "typo_where" },
  { column: 41, header: "check_for_typos" },
  { column: 42, header: "spanish_yn" },
  { column: 43, header: "misspelled_word" }
];

const REVIEW_SUPPORT_COLUMNS = [
  { column: 37, header: "Source Row" },
  { column: 38, header: "Record ID" },
  { column: 39, header: "Source Exclude?" },
  { column: 40, header: "Source Exclude Reason" },
  { column: 41, header: "Source Exact Pull Count" }
];

function upgradeQuoteWorkflowFoundation() {
  const config = EXCERPT_DUPLICATE_CONFIG;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(config.sourceSheetName);
  const reviewSheet = spreadsheet.getSheetByName(config.sheetName);

  if (!sourceSheet) {
    throw new Error('Source sheet "' + config.sourceSheetName + '" not found.');
  }
  if (!reviewSheet) {
    throw new Error('Review sheet "' + config.sheetName + '" not found.');
  }

  spreadsheet.toast("Upgrading quote workflow foundation...", "Workflow Upgrade", 5);

  setupSourceWorkflowColumns_(sourceSheet);
  setupReviewSupportColumns_(reviewSheet, sourceSheet, config);

  spreadsheet.toast(
    "Workflow foundation upgraded. Source helpers and review support columns are in place.",
    "Workflow Upgrade",
    7
  );
}

function setupSourceWorkflowColumns_(sourceSheet) {
  ensureHeadersAreSafe_(sourceSheet, SOURCE_WORKFLOW_COLUMNS);
  ensureHeadersAreSafe_(sourceSheet, SOURCE_REVIEW_STATE_COLUMNS);

  SOURCE_WORKFLOW_COLUMNS.forEach(function(definition) {
    sourceSheet.getRange(1, definition.column).setValue(definition.header);
  });
  SOURCE_REVIEW_STATE_COLUMNS.forEach(function(definition) {
    sourceSheet.getRange(1, definition.column).setValue(definition.header);
  });

  sourceSheet.getRange(2, 25).setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",TEXT(A2:A,"yyyymmddhhmmss")&"-"&ROW(A2:A)))'
  );
  sourceSheet.getRange(2, 26).setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",IF(D2:D<>"",D2:D,IF(J2:J="other",K2:K,J2:J))))'
  );
  sourceSheet.getRange(2, 27).setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",IF(F2:F<>"",F2:F,L2:L)))'
  );
  sourceSheet.getRange(2, 28).setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",IF(G2:G<>"",G2:G,M2:M)))'
  );
  sourceSheet.getRange(2, 29).setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",IF(H2:H<>"",H2:H,N2:N)))'
  );
}

function setupReviewSupportColumns_(reviewSheet, sourceSheet, config) {
  ensureHeadersAreSafe_(reviewSheet, REVIEW_SUPPORT_COLUMNS);

  REVIEW_SUPPORT_COLUMNS.forEach(function(definition) {
    reviewSheet.getRange(1, definition.column).setValue(definition.header);
  });

  const sourceName = sourceSheet.getName().replace(/'/g, "''");
  reviewSheet.getRange(2, 37).setFormula(
    "=ARRAYFORMULA(IF('" + sourceName + "'!A2:A=\"\",\"\",ROW('" + sourceName + "'!A2:A)))"
  );
  reviewSheet.getRange(2, 38).setFormula(
    "=ARRAYFORMULA('" + sourceName + "'!Y2:Y)"
  );
  reviewSheet.getRange(2, 39).setFormula(
    "=ARRAYFORMULA('" + sourceName + "'!AD2:AD)"
  );
  reviewSheet.getRange(2, 40).setFormula(
    "=ARRAYFORMULA('" + sourceName + "'!AE2:AE)"
  );
  reviewSheet.getRange(2, 41).setFormula(
    "=ARRAYFORMULA('" + sourceName + "'!AH2:AH)"
  );
}

function syncExactDuplicateExclusionsToSource() {
  const config = EXCERPT_DUPLICATE_CONFIG;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const reviewSheet = spreadsheet.getSheetByName(config.sheetName);
  const sourceSheet = spreadsheet.getSheetByName(config.sourceSheetName);

  if (!reviewSheet) {
    throw new Error('Review sheet "' + config.sheetName + '" not found.');
  }
  if (!sourceSheet) {
    throw new Error('Source sheet "' + config.sourceSheetName + '" not found.');
  }

  const reviewLastRow = reviewSheet.getLastRow();
  if (reviewLastRow < config.startRow) {
    spreadsheet.toast("No review rows found.", "Duplicate Review", 7);
    return;
  }

  const sourceRowColumn = config.supportStartColumn;
  const recordIdColumn = config.supportStartColumn + 1;
  const sourceExcludeReasonColumn = config.supportStartColumn + 3;
  const sourceReviewWidth = Math.max(sourceExcludeReasonColumn, config.recoveredPullCountColumn);
  const reviewValues = reviewSheet
    .getRange(config.startRow, 1, reviewLastRow - config.startRow + 1, sourceReviewWidth)
    .getValues();

  const reviewRowMap = {};
  reviewValues.forEach(function(row, index) {
    const sheetRow = config.startRow + index;
    reviewRowMap[sheetRow] = {
      sourceRow: parseInt(row[sourceRowColumn - 1], 10),
      recordId: (row[recordIdColumn - 1] || "").toString().trim()
    };
  });

  const sourceUpdates = {};
  let excludeCount = 0;
  let keepCount = 0;
  let candidateRowsFound = 0;
  let keepRowsFound = 0;
  let missingSourceLinkCount = 0;

  reviewValues.forEach(function(row, index) {
    const reviewRowNumber = config.startRow + index;
    let sourceRowNumber = parseInt(row[sourceRowColumn - 1], 10);

    const deleteCandidate = (row[config.deleteReviewStartColumn - 1] || "").toString().trim().toUpperCase();
    const decision = (row[config.deleteReviewStartColumn] || "").toString().trim().toUpperCase();
    const duplicateGroupId = (row[config.deleteReviewStartColumn + 2] || "").toString().trim();
    const keepRowNumber = parseInt(row[config.deleteReviewStartColumn + 3], 10);
    const exactPullCount = parsePullCount_(row[config.exactPullCountColumn - 1]);
    const keepSourceRecordId = keepRowNumber && reviewRowMap[keepRowNumber]
      ? reviewRowMap[keepRowNumber].recordId
      : "";

    if (decision === "DELETE_CANDIDATE" && deleteCandidate === "Y") {
      candidateRowsFound++;
      if (!sourceRowNumber) {
        sourceRowNumber = reviewRowNumber;
      }
      if (!sourceRowNumber) {
        missingSourceLinkCount++;
        return;
      }
      sourceUpdates[sourceRowNumber] = {
        exclude: "Y",
        reason: "Exact duplicate; merged into keep row",
        groupId: duplicateGroupId,
        keepRecordId: keepSourceRecordId,
        pullCount: exactPullCount || ""
      };
      excludeCount++;
    } else if (decision === "KEEP") {
      keepRowsFound++;
      if (!sourceRowNumber) {
        sourceRowNumber = reviewRowNumber;
      }
      if (!sourceRowNumber) {
        missingSourceLinkCount++;
        return;
      }
      sourceUpdates[sourceRowNumber] = {
        exclude: "",
        reason: "",
        groupId: duplicateGroupId,
        keepRecordId: keepSourceRecordId || (reviewRowMap[reviewRowNumber] ? reviewRowMap[reviewRowNumber].recordId : ""),
        pullCount: exactPullCount || ""
      };
      keepCount++;
    }
  });

  const sourceRows = Object.keys(sourceUpdates)
    .map(function(rowNumber) {
      return parseInt(rowNumber, 10);
    })
    .sort(function(a, b) {
      return a - b;
    });

  sourceRows.forEach(function(sourceRowNumber) {
    const update = sourceUpdates[sourceRowNumber];
    sourceSheet.getRange(sourceRowNumber, 30).setValue(update.exclude);
    sourceSheet.getRange(sourceRowNumber, 31).setValue(update.reason);
    sourceSheet.getRange(sourceRowNumber, 32).setValue(update.groupId);
    sourceSheet.getRange(sourceRowNumber, 33).setValue(update.keepRecordId);
    sourceSheet.getRange(sourceRowNumber, 34).setValue(update.pullCount);
  });

  spreadsheet.toast(
    "Found " + candidateRowsFound + " candidates and " + keepRowsFound +
      " keep rows. Synced " + excludeCount + " excludes and " + keepCount +
      " keeps upstream. Missing links: " + missingSourceLinkCount + ".",
    "Duplicate Review",
    7
  );
}

function applyExcludeFilterToReviewTab() {
  const config = EXCERPT_DUPLICATE_CONFIG;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const reviewSheet = spreadsheet.getSheetByName(config.sheetName);

  if (!reviewSheet) {
    throw new Error('Review sheet "' + config.sheetName + '" not found.');
  }

  const lastRow = reviewSheet.getLastRow();
  const lastColumn = reviewSheet.getLastColumn();
  if (lastRow < 2) return;

  let filter = reviewSheet.getFilter();
  if (!filter) {
    filter = reviewSheet.getRange(1, 1, lastRow, lastColumn).createFilter();
  }

  const criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextDoesNotContain("Y")
    .build();
  filter.setColumnFilterCriteria(config.supportStartColumn + 2, criteria);

  spreadsheet.toast(
    "Applied review-tab filter to hide rows excluded upstream.",
    "Duplicate Review",
    7
  );
}

function buildExcludeAwareReviewPreview() {
  const config = EXCERPT_DUPLICATE_CONFIG;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(config.sourceSheetName);
  const reviewSheet = spreadsheet.getSheetByName(config.sheetName);

  if (!sourceSheet) {
    throw new Error('Source sheet "' + config.sourceSheetName + '" not found.');
  }
  if (!reviewSheet) {
    throw new Error('Review sheet "' + config.sheetName + '" not found.');
  }

  let previewSheet = spreadsheet.getSheetByName(config.previewSheetName);
  if (!previewSheet) {
    previewSheet = spreadsheet.insertSheet(config.previewSheetName);
  } else {
    previewSheet.clearContents();
    previewSheet.clearFormats();
  }

  const reviewHeaders = reviewSheet.getRange(1, 1, 1, REVIEW_SUPPORT_COLUMNS[REVIEW_SUPPORT_COLUMNS.length - 1].column).getValues()[0];
  previewSheet.getRange(1, 1, 1, reviewHeaders.length).setValues([reviewHeaders]);

  const sourceName = sourceSheet.getName().replace(/'/g, "''");
  const reviewName = reviewSheet.getName().replace(/'/g, "''");
  const includeCondition = "'" + sourceName + "'!AB2:AB<>\"\"";
  const excludeCondition = "'" + sourceName + "'!AD2:AD<>\"Y\"";

  const generatedColumnFormulas = {
    A: "=FILTER('" + sourceName + "'!A2:A," + includeCondition + "," + excludeCondition + ")",
    B: "=FILTER('" + sourceName + "'!B2:B," + includeCondition + "," + excludeCondition + ")",
    C: "=FILTER('" + sourceName + "'!Z2:Z," + includeCondition + "," + excludeCondition + ")",
    D: "=FILTER('" + sourceName + "'!AA2:AA," + includeCondition + "," + excludeCondition + ")",
    E: "=FILTER('" + sourceName + "'!C2:C," + includeCondition + "," + excludeCondition + ")",
    F: "=FILTER('" + sourceName + "'!AC2:AC," + includeCondition + "," + excludeCondition + ")",
    G: "=FILTER('" + sourceName + "'!AB2:AB," + includeCondition + "," + excludeCondition + ")",
    I: "=FILTER('" + sourceName + "'!I2:I," + includeCondition + "," + excludeCondition + ")",
    K: "=FILTER('" + sourceName + "'!O2:O," + includeCondition + "," + excludeCondition + ")",
    L: "=FILTER('" + sourceName + "'!P2:P," + includeCondition + "," + excludeCondition + ")",
    AK: "=FILTER(ROW('" + sourceName + "'!A2:A), " + includeCondition + "," + excludeCondition + ")",
    AL: "=FILTER('" + sourceName + "'!Y2:Y," + includeCondition + "," + excludeCondition + ")",
    AM: "=FILTER('" + sourceName + "'!AD2:AD," + includeCondition + "," + excludeCondition + ")",
    AN: "=FILTER('" + sourceName + "'!AE2:AE," + includeCondition + "," + excludeCondition + ")",
    AO: "=FILTER('" + sourceName + "'!AH2:AH," + includeCondition + "," + excludeCondition + ")"
  };

  Object.keys(generatedColumnFormulas).forEach(function(columnLetter) {
    previewSheet.getRange(columnLetter + "2").setFormula(generatedColumnFormulas[columnLetter]);
  });

  previewSheet.getRange("N2").setFormula(
    '=ARRAYFORMULA(IF(G2:G="", "", LEN(G2:G) - LEN(SUBSTITUTE(G2:G, " ", "")) + 1))'
  );

  const lookupColumns = [
    "M", "O", "P", "Q", "R", "S", "T", "U", "V",
    "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE",
    "AF", "AG", "AH", "AI", "AJ"
  ];

  lookupColumns.forEach(function(columnLetter) {
    previewSheet.getRange(columnLetter + "2").setFormula(
      '=ARRAYFORMULA(IF($AL2:$AL="", "", IFNA(XLOOKUP($AL2:$AL, \'' + reviewName +
      '\'!$AL$2:$AL, \'' + reviewName + '\'!' + columnLetter + '$2:' + columnLetter + ', ""), "")))'
    );
  });

  previewSheet.setFrozenRows(1);
  spreadsheet.toast(
    'Built "' + config.previewSheetName + '" using source-side excludes plus record-linked review lookups.',
    "Workflow Upgrade",
    7
  );
}

function promoteExcludeAwareReviewTab() {
  const config = EXCERPT_DUPLICATE_CONFIG;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(config.sourceSheetName);
  const reviewSheet = spreadsheet.getSheetByName(config.sheetName);

  if (!sourceSheet) {
    throw new Error('Source sheet "' + config.sourceSheetName + '" not found.');
  }
  if (!reviewSheet) {
    throw new Error('Review sheet "' + config.sheetName + '" not found.');
  }

  spreadsheet.toast("Promoting exclude-aware review tab...", "Workflow Upgrade", 5);

  setupSourceWorkflowColumns_(sourceSheet);
  setupReviewSupportColumns_(reviewSheet, sourceSheet, config);
  syncReviewStateToSource();
  promoteMainReviewFormulas();

  spreadsheet.toast(
    "Promotion complete. Run \"Refresh duplicate review columns\" next.",
    "Workflow Upgrade",
    7
  );
}

function syncReviewStateToSource() {
  const config = EXCERPT_DUPLICATE_CONFIG;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(config.sourceSheetName);
  const reviewSheet = spreadsheet.getSheetByName(config.sheetName);

  if (!sourceSheet) {
    throw new Error('Source sheet "' + config.sourceSheetName + '" not found.');
  }
  if (!reviewSheet) {
    throw new Error('Review sheet "' + config.sheetName + '" not found.');
  }

  spreadsheet.toast("Syncing editable review state to source...", "Workflow Upgrade", 5);
  setupSourceWorkflowColumns_(sourceSheet);
  setupReviewSupportColumns_(reviewSheet, sourceSheet, config);
  syncCurrentReviewStateToSource_(reviewSheet, sourceSheet, config);
  spreadsheet.toast("Review state synced to source.", "Workflow Upgrade", 7);
}

function promoteMainReviewFormulas() {
  const config = EXCERPT_DUPLICATE_CONFIG;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(config.sourceSheetName);
  const reviewSheet = spreadsheet.getSheetByName(config.sheetName);

  if (!sourceSheet) {
    throw new Error('Source sheet "' + config.sourceSheetName + '" not found.');
  }
  if (!reviewSheet) {
    throw new Error('Review sheet "' + config.sheetName + '" not found.');
  }

  spreadsheet.toast("Replacing main review-tab formulas...", "Workflow Upgrade", 5);
  setupReviewSupportColumns_(reviewSheet, sourceSheet, config);
  replaceReviewTabPopulationFormulas_(reviewSheet, sourceSheet, config);
  spreadsheet.toast(
    "Main review-tab formulas replaced. Run \"Refresh duplicate review columns\" next.",
    "Workflow Upgrade",
    7
  );
}

function syncCurrentReviewStateToSource_(reviewSheet, sourceSheet, config) {
  const reviewLastRow = reviewSheet.getLastRow();
  if (reviewLastRow < config.startRow) return;
  const sourceLastRow = sourceSheet.getLastRow();
  if (sourceLastRow < config.startRow) return;

  const width = Math.max(config.supportStartColumn + 1, 22);
  const reviewValues = reviewSheet
    .getRange(config.startRow, 1, reviewLastRow - config.startRow + 1, width)
    .getValues();
  const rowCount = sourceLastRow - config.startRow + 1;
  const sourceStateRange = sourceSheet.getRange(config.startRow, 35, rowCount, 9);
  const sourceStateValues = sourceStateRange.getValues();

  reviewValues.forEach(function(row, index) {
    const sourceRow = parseInt(row[config.supportStartColumn - 1], 10) || (config.startRow + index);
    const sourceIndex = sourceRow - config.startRow;
    if (sourceIndex < 0 || sourceIndex >= sourceStateValues.length) return;

    sourceStateValues[sourceIndex][0] = row[12];
    sourceStateValues[sourceIndex][1] = row[14];
    sourceStateValues[sourceIndex][2] = row[15];
    sourceStateValues[sourceIndex][3] = row[16];
    sourceStateValues[sourceIndex][4] = row[17];
    sourceStateValues[sourceIndex][5] = row[18];
    sourceStateValues[sourceIndex][6] = row[19];
    sourceStateValues[sourceIndex][7] = row[20];
    sourceStateValues[sourceIndex][8] = row[21];
  });

  sourceStateRange.setValues(sourceStateValues);
}

function replaceReviewTabPopulationFormulas_(reviewSheet, sourceSheet, config) {
  const sourceName = sourceSheet.getName().replace(/'/g, "''");
  const includeCondition = "'" + sourceName + "'!AB2:AB<>\"\"";
  const excludeCondition = "'" + sourceName + "'!AD2:AD<>\"Y\"";
  const lastRow = Math.max(reviewSheet.getMaxRows(), reviewSheet.getLastRow(), 2);

  // Clear old manual review-state values so the new array formulas can spill.
  reviewSheet.getRange(2, 13, lastRow - 1, 10).clearContent();

  const formulas = {
    A: "=FILTER('" + sourceName + "'!A2:A," + includeCondition + "," + excludeCondition + ")",
    B: "=FILTER('" + sourceName + "'!B2:B," + includeCondition + "," + excludeCondition + ")",
    C: "=FILTER('" + sourceName + "'!Z2:Z," + includeCondition + "," + excludeCondition + ")",
    D: "=FILTER('" + sourceName + "'!AA2:AA," + includeCondition + "," + excludeCondition + ")",
    E: "=FILTER('" + sourceName + "'!C2:C," + includeCondition + "," + excludeCondition + ")",
    F: "=FILTER('" + sourceName + "'!AC2:AC," + includeCondition + "," + excludeCondition + ")",
    G: "=FILTER('" + sourceName + "'!AB2:AB," + includeCondition + "," + excludeCondition + ")",
    I: "=FILTER('" + sourceName + "'!I2:I," + includeCondition + "," + excludeCondition + ")",
    K: "=FILTER('" + sourceName + "'!O2:O," + includeCondition + "," + excludeCondition + ")",
    L: "=FILTER('" + sourceName + "'!P2:P," + includeCondition + "," + excludeCondition + ")",
    M: "=FILTER('" + sourceName + "'!AI2:AI," + includeCondition + "," + excludeCondition + ")",
    O: "=FILTER('" + sourceName + "'!AJ2:AJ," + includeCondition + "," + excludeCondition + ")",
    P: "=FILTER('" + sourceName + "'!AK2:AK," + includeCondition + "," + excludeCondition + ")",
    Q: "=FILTER('" + sourceName + "'!AL2:AL," + includeCondition + "," + excludeCondition + ")",
    R: "=FILTER('" + sourceName + "'!AM2:AM," + includeCondition + "," + excludeCondition + ")",
    S: "=FILTER('" + sourceName + "'!AN2:AN," + includeCondition + "," + excludeCondition + ")",
    T: "=FILTER('" + sourceName + "'!AO2:AO," + includeCondition + "," + excludeCondition + ")",
    U: "=FILTER('" + sourceName + "'!AP2:AP," + includeCondition + "," + excludeCondition + ")",
    V: "=FILTER('" + sourceName + "'!AQ2:AQ," + includeCondition + "," + excludeCondition + ")",
    AK: "=FILTER(ROW('" + sourceName + "'!A2:A), " + includeCondition + "," + excludeCondition + ")",
    AL: "=FILTER('" + sourceName + "'!Y2:Y," + includeCondition + "," + excludeCondition + ")",
    AM: "=FILTER('" + sourceName + "'!AD2:AD," + includeCondition + "," + excludeCondition + ")",
    AN: "=FILTER('" + sourceName + "'!AE2:AE," + includeCondition + "," + excludeCondition + ")",
    AO: "=FILTER('" + sourceName + "'!AH2:AH," + includeCondition + "," + excludeCondition + ")"
  };

  Object.keys(formulas).forEach(function(columnLetter) {
    reviewSheet.getRange(columnLetter + "2").setFormula(formulas[columnLetter]);
  });

  reviewSheet.getRange("N2").setFormula(
    '=ARRAYFORMULA(IF(G2:G="", "", LEN(G2:G) - LEN(SUBSTITUTE(G2:G, " ", "")) + 1))'
  );
}

// Legacy migration/recovery helpers remain below for safety, but the active
// review workflow is now:
// 1. Check excerpt duplicates
// 2. Mark exact duplicate delete candidates
// 3. Sync exact-duplicate excludes upstream
// 4. Refresh duplicate review columns

function ensureHeadersAreSafe_(sheet, definitions) {
  definitions.forEach(function(definition) {
    const existingHeader = (sheet.getRange(1, definition.column).getValue() || "").toString().trim();
    if (existingHeader && existingHeader !== definition.header) {
      throw new Error(
        'Refusing to overwrite existing header "' + existingHeader +
        '" in ' + sheet.getName() + "!" + columnNumberToLetter_(definition.column) + "1."
      );
    }
  });
}

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

function createEbookProofingFormV5() {
  var form = FormApp.create('Button Poetry eBook Proofing Form v5');

  form.setDescription(
    'Use this form when proofing converted eBook files. You will be comparing the converted eBook to the final source PDF used for conversion. ' +
    'When helpful, you may also use the print book as a reference. ' +
    'A file name with (k) usually indicates the Kindle version. ' +
    'Proof standard EPUB files in Apple Books or Adobe Digital Editions. ' +
    'Kindle-optimized EPUB files may be reviewed in Apple Books, but should also be checked in the Kindle app when possible. ' +
    'MOBI files should be checked in the Kindle app. ' +
    'Proof ePDF files in Preview or Adobe Acrobat. Do not use Preview as the main proofing environment for reflowable EPUB or Kindle files. ' +
    'On the first proofing round, check the file line-by-line against the PDF. For later rounds, focus on verifying requested corrections. ' +
    'For reflowable versions, set the font using the correct longest eligible poetic line before judging orphaned lines, split stanzas, wraparounds, or form preservation.'
  );
  form.setConfirmationMessage(
    'Thank you for proofing materials for this title. Publishing will be in touch if we have any questions.'
  );
  form.setCollectEmail(false);

  form.addSectionHeaderItem().setTitle('Before You Begin').setHelpText(
    'Open both the converted file and the final source PDF before continuing. ' +
    'On first proofs, compare line-by-line. Unless a question says otherwise, answer based on your review of the full book.'
  );

  addEbookFormTextItem_(form, 'Proofer First Name', true);
  addEbookFormTextItem_(form, 'Proofer Last Name', true);
  addEbookFormTextItem_(form, 'Proofer Email', true);
  addEbookFormTextItem_(form, 'Book Title', true);
  addEbookFormTextItem_(form, 'Author First Name', true);
  addEbookFormTextItem_(form, 'Author Last Name', true);
  addEbookFormMultipleChoiceItem_(
    form,
    'Which version did you proof?',
    ['Standard EPUB', 'Kindle-optimized EPUB (usually marked with (k) in the filename)', 'MOBI', 'ePDF', 'Other'],
    true
  );
  addEbookFormTextItem_(form, 'Exact file name proofed', true);
  addEbookFormCheckboxItem_(
    form,
    'Which device(s) or app(s) did you use for proofing?',
    [
      'iPad / Apple Books (standard EPUB; acceptable for Kindle EPUB review)',
      'Adobe Digital Editions (standard EPUB)',
      'Kindle app (Kindle EPUB / MOBI)',
      'Kindle device (optional, if available)',
      'Preview (ePDF only)',
      'Adobe Acrobat (ePDF only)',
      'Other'
    ],
    true
  );
  addEbookFormParagraphItem_(form, 'If needed, note the proofing environment used and why', false);
  addEbookFormTextItem_(form, 'If other, list device/app used', false);
  addEbookFormMultipleChoiceItem_(form, 'Is this the first proofing round for this file?', ['Yes', 'No'], true);
  addEbookFormMultipleChoiceItem_(
    form,
    'If this is the first proofing round, did you complete a line-by-line comparison against the source PDF?',
    ['Yes', 'No', 'N/A - Not the first proofing round'],
    true
  );
  addEbookFormCheckboxItem_(
    form,
    'Which areas did you review in this pass?',
    ['Library listing', 'Front matter', 'TOC', 'Body content', 'Back matter', 'Correction round only'],
    true
  );

  form.addPageBreakItem().setTitle('REFLOWABLE SETUP - LONGEST LINE').setHelpText(
    'For reflowable versions only, compare against the source PDF when setting type size. Use the longest eligible poetic line before evaluating orphaned lines, split stanzas, wraparounds, and form preservation.'
  );
  addEbookFormMultipleChoiceItem_(
    form,
    'For reflowable versions (standard EPUB, Kindle EPUB, MOBI), did you set the proofing font size using the correct longest eligible poetic line before evaluating layout issues?',
    ['Yes', 'No', 'N/A - ePDF/fixed layout'],
    true
  );
  addEbookFormParagraphItem_(form, 'If applicable, paste or describe the longest line used for sizing', false);
  addEbookFormMultipleChoiceItem_(
    form,
    'Were orphaned lines, split stanzas, wraparounds, and form-preservation checks judged only after longest-line sizing was set?',
    ['Yes', 'No', 'N/A - ePDF/fixed layout'],
    true
  );
  addEbookFormParagraphItem_(form, 'Additional notes for longest-line setup', false);

  form.addPageBreakItem().setTitle('LIBRARY LISTING + COVER').setHelpText(
    'Answer the first three questions before opening the book. Then open the file and answer the in-book cover questions for the version you reviewed.'
  );
  addEbookFormMultipleChoiceItem_(form, 'Before opening the book, is the book title spelled correctly in the library listing?', ['Yes', 'No'], true);
  addEbookFormMultipleChoiceItem_(form, 'Before opening the book, is the author\'s name spelled correctly in the library listing?', ['Yes', 'No'], true);
  addEbookFormMultipleChoiceItem_(form, 'Before opening the book, does the cover image appear correctly in the library listing?', ['Yes', 'No'], true);
  addEbookFormMultipleChoiceItem_(
    form,
    'For all non-Kindle ebook files, after opening the book, is the front cover present and correct?',
    ['Yes', 'No', 'N/A - Kindle EPUB version'],
    true
  );
  addEbookFormMultipleChoiceItem_(
    form,
    'For all non-Kindle ebook files, after opening the book, is the back cover present and correct?',
    ['Yes', 'No', 'N/A - Kindle EPUB version or not applicable'],
    false
  );
  addEbookFormMultipleChoiceItem_(
    form,
    'If this is the Kindle EPUB version, after opening the book, is the cover absent from the reading flow?',
    ['Yes', 'No', 'N/A - Not the Kindle EPUB version'],
    false
  );
  addEbookFormParagraphItem_(form, 'If any library listing or cover issue appears above, note where it occurs', false);

  form.addPageBreakItem().setTitle('ePDF-SPECIFIC CHECKS').setHelpText(
    'Answer this section only for the ePDF version. Review the full PDF when possible. If you answer No to any question, note where the issue appears.'
  );
  addEbookFormMultipleChoiceItem_(
    form,
    'If this is the ePDF version, does the page layout match the intended print layout across the full book?',
    ['Yes', 'No', 'N/A - Not the ePDF version'],
    false
  );
  addEbookFormMultipleChoiceItem_(
    form,
    'If this is the ePDF version, is the text readable at a normal viewing size without unexpected distortion or cropping?',
    ['Yes', 'No', 'N/A - Not the ePDF version'],
    false
  );
  addEbookFormMultipleChoiceItem_(
    form,
    'If this is the ePDF version, does the clickable TOC/PDF navigation work correctly?',
    ['Yes', 'No', 'N/A - No clickable TOC/navigation present', 'N/A - Not the ePDF version'],
    false
  );
  addEbookFormMultipleChoiceItem_(
    form,
    'If this is the ePDF version, do internal and external links work correctly?',
    ['Yes', 'No', 'N/A - No links present', 'N/A - Not the ePDF version'],
    false
  );
  addEbookFormParagraphItem_(form, 'If any ePDF issue appears above, note where it occurs', false);

  form.addPageBreakItem().setTitle('FRONT MATTER').setHelpText(
    'Compare the front matter to the source PDF and use the print book as a reference when helpful. The expected eBook front matter usually includes the title page, copyright page, A Note on Poetry E-Books, and the TOC. Optional front matter may include praise, an author note, and a quote or epigraph page.'
  );
  addEbookFormMultipleChoiceItem_(form, 'Is the copyright page content correct for this title when compared to the source PDF/print version?', ['Yes', 'No'], true);
  addEbookFormCheckboxItem_(
    form,
    'Which of the following appear correct on the copyright page, in reading order?',
    [
      'Title',
      'Category / Poetry',
      'Author name',
      'Credits / logos',
      'All Rights Reserved',
      'Copyright year',
      'Publisher name',
      'Publisher location',
      'Publisher website',
      'Print ISBN',
      'eBook ISBN',
      'Audiobook ISBN',
      'Printing information'
    ],
    false
  );
  addEbookFormMultipleChoiceItem_(form, 'Does the front matter present the expected eBook structure for this title, with required elements present and in the correct order?', ['Yes', 'No'], true);
  addEbookFormCheckboxItem_(
    form,
    'Which optional front-matter elements are present in this title?',
    ['Praise', 'Dedication', 'Epigraph / quote page', 'Note from the Author', 'A Note on Poetry E-Books'],
    false
  );
  addEbookFormMultipleChoiceItem_(
    form,
    'If present, are the optional front-matter elements free of typos and in the correct order?',
    ['Yes', 'No', 'N/A - No optional front-matter elements present'],
    false
  );
  addEbookFormMultipleChoiceItem_(
    form,
    'Are logo black pages at the front and back of the book preserved where they should be in this version?',
    ['Yes', 'No', 'N/A - No logo black pages intended for this title'],
    false
  );
  addEbookFormMultipleChoiceItem_(
    form,
    'If this title uses section black pages, are they preserved where they should be in this version when compared to the source PDF/print version?',
    ['Yes', 'No', 'N/A - This title has no section black pages'],
    false
  );
  addEbookFormMultipleChoiceItem_(
    form,
    'For reflowable versions (standard EPUB, Kindle EPUB, MOBI), have blank black back sides connected to logo or section black pages been removed where appropriate?',
    ['Yes', 'No', 'N/A - ePDF preserves blank black back sides'],
    false
  );
  addEbookFormParagraphItem_(form, 'If any front-matter issue appears above, note where it occurs', false);

  form.addPageBreakItem().setTitle('TOC').setHelpText(
    'Compare the visible TOC to the source PDF. The visible TOC should include poems or main reading entries, sections if the book has sections, and intended navigable back matter. It should not include front matter such as the cover, title page, or copyright page.'
  );
  addEbookFormMultipleChoiceItem_(form, 'Is any expected poem, section, or intended back matter entry missing from the visible TOC when compared to the source PDF?', ['Yes', 'No'], true);
  addEbookFormMultipleChoiceItem_(form, 'Does the linked TOC allow you to jump to every poem or intended entry correctly?', ['Yes', 'No'], true);
  addEbookFormMultipleChoiceItem_(
    form,
    'If this is the Kindle EPUB version, does the book begin at the title page rather than a duplicate cover page?',
    ['Yes', 'No', 'N/A - Not the Kindle EPUB version'],
    false
  );
  addEbookFormMultipleChoiceItem_(form, 'If applicable, do non-TOC internal links work correctly?', ['Yes', 'No', 'N/A - No non-TOC internal links present'], false);
  addEbookFormParagraphItem_(form, 'If any TOC or navigation issue appears above, note where it occurs', false);

  form.addPageBreakItem().setTitle('BODY CONTENT FORMATTING').setHelpText(
    'Compare these items to the source PDF. If you answer No or Yes with an issue to any of the questions below, note the poem/title/location in the notes field at the end of this section.'
  );
  addEbookFormMultipleChoiceItem_(form, 'Does the eBook match the source PDF line-by-line with no missing text?', ['Yes', 'No'], true);
  addEbookFormMultipleChoiceItem_(form, 'Are all poems, sections, and end-matter pages present with no missing content?', ['Yes', 'No'], true);
  addEbookFormMultipleChoiceItem_(form, 'Do prose sections reflow as prose when font size changes, and do poetry sections preserve their intended line breaks?', ['Yes', 'No'], true);
  addEbookFormMultipleChoiceItem_(
    form,
    'For reflowable versions (standard EPUB, Kindle EPUB, MOBI), have all blank pages, including blank black back sides, been removed from the eBook?',
    ['Yes', 'No', 'N/A - ePDF keeps blank pages'],
    true
  );
  addEbookFormMultipleChoiceItem_(form, 'Do poem titles and section titles match the source PDF clearly enough to preserve their intended distinction from the body text?', ['Yes', 'No'], true);
  addEbookFormMultipleChoiceItem_(form, 'Are italics, bold, and other emphasis retained correctly?', ['Yes', 'No'], true);
  addEbookFormMultipleChoiceItem_(form, 'Are strikethroughs, special characters, ornaments, symbols, or other special text treatments retained correctly?', ['Yes', 'No', 'N/A - None present'], false);
  addEbookFormMultipleChoiceItem_(form, 'Are stanza breaks, line breaks, and indents preserved correctly?', ['Yes', 'No'], true);
  addEbookFormMultipleChoiceItem_(form, 'Are lines that should remain in the same stanza kept together correctly?', ['Yes', 'No'], true);
  addEbookFormMultipleChoiceItem_(form, 'If applicable, are unusually formatted or image-based poems handled correctly?', ['Yes', 'No', 'N/A - None present'], false);
  addEbookFormMultipleChoiceItem_(form, 'Are there any orphaned lines that should be flagged?', ['Yes', 'No', 'N/A - ePDF/fixed layout'], true);
  addEbookFormMultipleChoiceItem_(form, 'Are there any split stanzas that should be flagged?', ['Yes', 'No', 'N/A - ePDF/fixed layout'], true);
  addEbookFormMultipleChoiceItem_(form, 'Are there any broken or awkward line wraps that create reading errors when compared to the source PDF?', ['Yes', 'No', 'N/A - ePDF/fixed layout'], true);
  addEbookFormMultipleChoiceItem_(form, 'Are section breaks and chapter starts rendering correctly?', ['Yes', 'No', 'N/A - No sections or chapter starts in this title'], false);
  addEbookFormParagraphItem_(form, 'If any body-content issue appears above, note where it occurs', false);

  form.addPageBreakItem().setTitle('BACK MATTER');
  addEbookFormMultipleChoiceItem_(form, 'Is the acknowledgments page present for the correct book and free of typos?', ['Yes', 'No', 'N/A - Not present'], false);
  addEbookFormMultipleChoiceItem_(form, 'Is the author bio present for the correct author and free of typos?', ['Yes', 'No', 'N/A - Not present'], false);
  addEbookFormMultipleChoiceItem_(form, 'If present, is the author photo correct and clean?', ['Yes', 'No', 'N/A - No photo present'], false);
  addEbookFormMultipleChoiceItem_(form, 'Are author recommendations correct for this book?', ['Yes', 'No', 'N/A - No recommendations present'], false);
  addEbookFormMultipleChoiceItem_(form, 'Are publisher back matter pages such as Other Books or Forthcoming pages correct for this book?', ['Yes', 'No', 'N/A - Not present'], false);
  addEbookFormParagraphItem_(form, 'If any back-matter issue appears above, note where it occurs', false);

  form.addPageBreakItem().setTitle('GLOBAL ISSUES + FINAL NOTES');
  addEbookFormMultipleChoiceItem_(form, 'Did you notice any recurring issue that appears throughout the eBook?', ['Yes', 'No'], true);
  addEbookFormParagraphItem_(form, 'If yes, describe the recurring issue, note, question, or suggested correction', false);
  addEbookFormParagraphItem_(form, 'List any title-specific notes or corrections that should be entered into the spreadsheet', false);
  addEbookFormParagraphItem_(form, 'Final notes or questions for Publishing', false);

  Logger.log('Edit URL: ' + form.getEditUrl());
  Logger.log('Published URL: ' + form.getPublishedUrl());
}

function addEbookFormTextItem_(form, title, required) {
  var item = form.addTextItem().setTitle(title);
  item.setRequired(required);
}

function addEbookFormParagraphItem_(form, title, required) {
  var item = form.addParagraphTextItem().setTitle(title);
  item.setRequired(required);
}

function addEbookFormMultipleChoiceItem_(form, title, choices, required) {
  var item = form.addMultipleChoiceItem().setTitle(title);
  item.setChoiceValues(choices);
  item.setRequired(required);
}

function addEbookFormCheckboxItem_(form, title, choices, required) {
  var item = form.addCheckboxItem().setTitle(title);
  item.setChoiceValues(choices);
  item.setRequired(required);
}
