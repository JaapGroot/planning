function getOrCreateTeamSheet_(file, teamName) {
  let sheet = file.getSheetByName(teamName);
  if (!sheet) sheet = file.insertSheet(teamName);

  for (const n of CONFIG.TEAM_SHEET_DEFAULTS) {
    const def = file.getSheetByName(n);
    if (def) file.deleteSheet(def);
  }

  while (sheet.getMaxColumns() < CONFIG.TOTAL_COLS) {
    sheet.insertColumnAfter(sheet.getMaxColumns());
  }
  return sheet;
}

function ensureTeamSheetSetupHeavyOnce_(file, sheet, planningSheet, teamName) {
  const props = PropertiesService.getDocumentProperties();
  const key = setupKey_(file, sheet);
  if (props.getProperty(key) === "1") return;

  sheet.getRange(CONFIG.TEAM_NAME_CELL).setValue(teamName);
  sheet.setFrozenRows(CONFIG.HEADER_ROWS);
  sheet.setFrozenColumns(CONFIG.INFO_COLS_END); // freeze A..G

  // Column widths for A..F from master A,H..K,M
  sheet.setColumnWidth(1, planningSheet.getColumnWidth(1));   // A
  sheet.setColumnWidth(2, planningSheet.getColumnWidth(8));   // B
  sheet.setColumnWidth(3, planningSheet.getColumnWidth(9));   // C
  sheet.setColumnWidth(4, planningSheet.getColumnWidth(10));  // D
  sheet.setColumnWidth(5, planningSheet.getColumnWidth(11));  // E
  sheet.setColumnWidth(6, planningSheet.getColumnWidth(13));  // F
  sheet.setColumnWidth(7, 110);                               // G (Afmelden)

  // Week widths H.. follow master H..
  const masterWeekCols = CONFIG.MASTER_TOTAL_COLS - CONFIG.MASTER_WEEK_START_COL + 1;
  for (let i = 0; i < masterWeekCols; i++) {
    sheet.setColumnWidth(CONFIG.WEEK_START_COL + i, planningSheet.getColumnWidth(CONFIG.MASTER_WEEK_START_COL + i));
  }

  // Default collapsed: hide H..M
  sheet.hideColumns(CONFIG.COLLAPSE_START_COL, CONFIG.COLLAPSE_NUM_COLS);

  // Week header VALUES once
  const hdr = planningSheet.getRange(1, CONFIG.MASTER_WEEK_START_COL, 3, masterWeekCols);
  const dest = sheet.getRange(1, CONFIG.WEEK_START_COL, 3, masterWeekCols);
  dest.setValues(hdr.getDisplayValues());

  mergeRepeatingHeaderCells_(sheet, 1, CONFIG.WEEK_START_COL, masterWeekCols);
  mergeRepeatingHeaderCells_(sheet, 2, CONFIG.WEEK_START_COL, masterWeekCols);

  protectSheetOnce_(sheet);

  if (!sheet.getFilter()) {
    sheet.getRange(CONFIG.INFO_HEADER_ROW, 1, 1, CONFIG.INFO_COLS_END).createFilter();
  }

  props.setProperty(key, "1");
}

function mergeRepeatingHeaderCells_(sheet, row, startCol, numCols) {
  const rowVals = sheet.getRange(row, startCol, 1, numCols).getDisplayValues()[0];
  let mergeStart = null;

  for (let i = 0; i < rowVals.length; i++) {
    const val = rowVals[i];
    const next = rowVals[i + 1];

    if (val && mergeStart === null) mergeStart = startCol + i;

    const isEnd = mergeStart !== null && (next || i === rowVals.length - 1);
    if (isEnd) {
      const mergeEnd = startCol + i;
      const width = mergeEnd - mergeStart + 1;
      if (width > 1) sheet.getRange(row, mergeStart, 1, width).mergeAcross();
      mergeStart = null;
    }
  }
}

function protectSheetOnce_(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

  // protect headers
  sheet.getRange("1:3").protect().addEditor(Session.getEffectiveUser());

  // protect info A..F (teams may edit G only)
  sheet.getRange(1, 1, sheet.getMaxRows(), 6)
    .protect()
    .addEditor(Session.getEffectiveUser());
}
