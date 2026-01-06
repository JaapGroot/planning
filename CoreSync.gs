function startSingleTeamUpdate(teamName) {
  withScriptLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planningSheet = ss.getSheetByName(CONFIG.PLANNING_SHEET);

    const cache = buildTeamFileCache_(ss);
    const count = updateTeamFileDetail_(ss, planningSheet, teamName, cache);

    SpreadsheetApp.getUi().alert(`✅ Update voor ${teamName} voltooid\nWerknummerblokken: ${count}\nVersie: v5.5`);
  });
}

function updateTeamFileDetail_(ss, planningSheet, teamName, fileCache) {
  const lastRow = planningSheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return 0;
  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;

  // Values A..M
  const valuesAM = planningSheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, CONFIG.MASTER_VALUES_COLS).getValues();

  // Backgrounds for mapped info columns (A, H..K, M) + L
  const bgA = planningSheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, 1).getBackgrounds();
  const bgHtoK = planningSheet.getRange(CONFIG.DATA_START_ROW, 8, numRows, 4).getBackgrounds();
  const bgM = planningSheet.getRange(CONFIG.DATA_START_ROW, 13, numRows, 1).getBackgrounds();
  const bgL = planningSheet.getRange(CONFIG.DATA_START_ROW, 12, numRows, 1).getBackgrounds(); // L

  // Text formatting (A, H..K, M) + L
  const fmt = readMappedTextFormats_(planningSheet, CONFIG.DATA_START_ROW, numRows);
  const fmtL = readSingleColTextFormats_(planningSheet, CONFIG.DATA_START_ROW, numRows, 12); // L

  // Week backgrounds from master H..end
  const masterWeekCols = CONFIG.MASTER_TOTAL_COLS - CONFIG.MASTER_WEEK_START_COL + 1;
  const bgWeeksMaster = planningSheet
    .getRange(CONFIG.DATA_START_ROW, CONFIG.MASTER_WEEK_START_COL, numRows, masterWeekCols)
    .getBackgrounds();

  // Team file/sheet
  const teamFile = openOrCreateTeamFileCached_(ss, teamName, fileCache);
  const teamSheet = getOrCreateTeamSheet_(teamFile, teamName);

  // Preserve existing afmelden states
  const afmeldenMap = readExistingAfmeldenMap_(teamSheet);

  // Heavy setup once
  ensureTeamSheetSetupHeavyOnce_(teamFile, teamSheet, planningSheet, teamName);

  // Keep headers styled like master
  refreshAllHeaderStylingEveryRun_(teamSheet, planningSheet);

  // Build output
  const detail = buildDetailOutputMappedTeamOnly_(
    valuesAM, bgA, bgHtoK, bgM, bgL, bgWeeksMaster, fmt, fmtL, teamName, afmeldenMap
  );

  // Write
  writeDetailToTeamSheetMapped_(teamFile, teamSheet, detail);

  // Ensure checkboxes on header rows
  applyAfmeldenCheckboxes_(teamSheet, detail.headerRowIndexes);

  // Log skipped
  logSkippedFast_(detail.skipped, teamFile);

  return detail.blockSizes.length;
}
