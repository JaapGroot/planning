function logSkippedFast_(skipped, file) {
  if (!skipped || !skipped.length) return;

  let logSheet = file.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = file.insertSheet(CONFIG.LOG_SHEET_NAME);
    logSheet.hideSheet();
  }

  logSheet.clearContents();
  logSheet.getRange(1, 1, 1, 4).setValues([["Timestamp", "Team", "Werknummer", "Reason"]]);
  logSheet.getRange(2, 1, skipped.length, 4).setValues(skipped);
}
