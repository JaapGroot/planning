function sortPlanningByLocation() {
  sortWorkBlocksWithoutGroups_({
    blockSortCol: 6,   // F
    blockSortCol2: 8,  // H
    detailSortCol: 3   // C
  });
}

function sortPlanningByWorkNumber() {
  sortWorkBlocksWithoutGroups_({
    blockSortCol: 1,   // A = werknummer
    blockSortCol2: 8, // H
    detailSortCol: 3   // C
  });
}

function sortWorkBlocksWithoutGroups_(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = SpreadsheetApp.getActiveSheet();

  const startRow = CONFIG.DATA_START_ROW;
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  const blockSortCol = options.blockSortCol;
  const blockSortCol2 = options.blockSortCol2;
  const detailSortCol = options.detailSortCol;

  if (lastRow < startRow || lastCol < 1) {
    SpreadsheetApp.getUi().alert("Geen sorteerbare data gevonden vanaf rij " + startRow + ".");
    return;
  }

  const numRows = lastRow - startRow + 1;
  const values = sh.getRange(startRow, 1, numRows, lastCol).getDisplayValues();

  const blocks = [];
  let currentBlock = null;
  let previousWorkNumber = null;

  for (let i = 0; i < numRows; i++) {
    const absoluteRow = startRow + i;
    const rowValues = values[i];
    const workNumber = safeCell_(rowValues, 1); // kolom A

    if (!workNumber) continue;

    const isHeader = workNumber !== previousWorkNumber;

    if (isHeader) {
      if (currentBlock) blocks.push(currentBlock);

      currentBlock = {
        workNumber: workNumber,
        headerRow: absoluteRow,
        sortValue: safeCell_(rowValues, blockSortCol),
        sortValue2: blockSortCol2 ? safeCell_(rowValues, blockSortCol2) : "",
        detailRows: []
      };
    } else {
      currentBlock.detailRows.push({
        rowNumber: absoluteRow,
        detailValue: safeCell_(rowValues, detailSortCol)
      });
    }

    previousWorkNumber = workNumber;
  }

  if (currentBlock) blocks.push(currentBlock);

  if (!blocks.length) {
    SpreadsheetApp.getUi().alert("Geen blokken gevonden vanaf rij " + startRow + ".");
    return;
  }

  for (const block of blocks) {
    if (block.detailRows.length < 2) continue;

    block.detailRows.sort((a, b) => {
      return compareDetailValues_(a.detailValue, b.detailValue);
    });
  }

  blocks.sort((a, b) => {
    const firstCompare = String(a.sortValue).localeCompare(
      String(b.sortValue),
      "nl",
      { numeric: true, sensitivity: "base" }
    );

    if (firstCompare !== 0) return firstCompare;

    if (!blockSortCol2) return 0;

    return String(a.sortValue2).localeCompare(
      String(b.sortValue2),
      "nl",
      { numeric: true, sensitivity: "base" }
    );
  });

  let temp = ss.getSheetByName(CONFIG.TEMP_SHEET_NAME);
  if (!temp) {
    temp = ss.insertSheet(CONFIG.TEMP_SHEET_NAME);
    temp.hideSheet();
  } else {
    temp.clear();
    temp.clearFormats();
  }

  if (temp.getMaxRows() < numRows + 50) {
    temp.insertRowsAfter(temp.getMaxRows(), (numRows + 50) - temp.getMaxRows());
  }
  if (temp.getMaxColumns() < lastCol) {
    temp.insertColumnsAfter(temp.getMaxColumns(), lastCol - temp.getMaxColumns());
  }

  let tempRow = 1;

  for (const block of blocks) {
    sh.getRange(block.headerRow, 1, 1, lastCol)
      .copyTo(temp.getRange(tempRow, 1, 1, lastCol), { contentsOnly: false });
    tempRow++;

    for (const detail of block.detailRows) {
      sh.getRange(detail.rowNumber, 1, 1, lastCol)
        .copyTo(temp.getRange(tempRow, 1, 1, lastCol), { contentsOnly: false });
      tempRow++;
    }
  }

  const totalOutputRows = tempRow - 1;

  if (totalOutputRows < 1) {
    SpreadsheetApp.getUi().alert("Er is geen output om terug te schrijven.");
    return;
  }

  if (temp.getMaxRows() < totalOutputRows) {
    temp.insertRowsAfter(temp.getMaxRows(), totalOutputRows - temp.getMaxRows());
  }

  const neededRowsOnTarget = startRow + totalOutputRows - 1;
  if (sh.getMaxRows() < neededRowsOnTarget) {
    sh.insertRowsAfter(sh.getMaxRows(), neededRowsOnTarget - sh.getMaxRows());
  }
  if (sh.getMaxColumns() < lastCol) {
    sh.insertColumnsAfter(sh.getMaxColumns(), lastCol - sh.getMaxColumns());
  }

  sh.getRange(startRow, 1, totalOutputRows, lastCol).clear({ contentsOnly: false });

  temp.getRange(1, 1, totalOutputRows, lastCol)
    .copyTo(sh.getRange(startRow, 1, totalOutputRows, lastCol), { contentsOnly: false });

  temp.clear();
  temp.clearFormats();
}

function safeCell_(rowValues, colIndex1Based) {
  if (!rowValues || !colIndex1Based || colIndex1Based < 1) return "";
  return String(rowValues[colIndex1Based - 1] || "").trim();
}

function compareDetailValues_(a, b) {

  const aParts = splitDetailValue_(a);
  const bParts = splitDetailValue_(b);

  const orderMap = getDetailOrderMap_();

  const suffixCompare = aParts.suffix.localeCompare(
    bParts.suffix,
    "nl",
    { numeric: true, sensitivity: "base" }
  );

  if (suffixCompare !== 0) return suffixCompare;

  const aRank = Object.prototype.hasOwnProperty.call(orderMap, aParts.base)
    ? orderMap[aParts.base]
    : 999;

  const bRank = Object.prototype.hasOwnProperty.call(orderMap, bParts.base)
    ? orderMap[bParts.base]
    : 999;

  if (aRank !== bRank) return aRank - bRank;

  return aParts.base.localeCompare(
    bParts.base,
    "nl",
    { numeric: true, sensitivity: "base" }
  );
}

function getDetailOrderMap_() {
  const map = {};
  CONFIG.DETAIL_ORDER.forEach((item, index) => {
    map[normalizeText_(item)] = index;
  });
  return map;
}

function normalizeText_(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function splitDetailValue_(value) {

  const norm = normalizeText_(value);

  const match = norm.match(/^([^(]+)\s*(?:\((.*)\))?$/);

  return {
    base: match ? match[1].trim() : norm,
    suffix: match && match[2] ? match[2].trim() : ""
  };
}