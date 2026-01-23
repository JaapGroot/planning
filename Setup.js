function getOrCreateTeamSheet_(file, teamName) {
  // Template-copy should already have a sheet named teamName.
  let sheet = file.getSheetByName(teamName);

  // Fallback (if user manually renamed)
  if (!sheet) sheet = file.getSheets()[0].setName(teamName);

  // Remove default sheets if present
  for (const n of CONFIG.TEAM_SHEET_DEFAULTS) {
    const def = file.getSheetByName(n);
    if (def) file.deleteSheet(def);
  }

  // Ensure enough columns
  if (sheet.getMaxColumns() < CONFIG.TOTAL_COLS) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), CONFIG.TOTAL_COLS - sheet.getMaxColumns());
  }

  return sheet;
}

function ensureTeamSheetSetupLight_(file, sheet, teamName) {
  const props = PropertiesService.getDocumentProperties();
  const key = setupKey_(file, sheet);
  if (props.getProperty(key) === "1") return;

  // Always keep A1 team name
  sheet.getRange(CONFIG.TEAM_NAME_CELL).setValue(teamName);

  // Keep frozen panes consistent (safe even if already set)
  sheet.setFrozenRows(CONFIG.HEADER_ROWS);
  sheet.setFrozenColumns(CONFIG.INFO_COLS_END); // A..F

  // Hide spacer column G (important so teams don't see it)
  sheet.hideColumns(CONFIG.SPACER_COL_G);

  // Filter on row 3 for A..F if missing
  const filter = sheet.getFilter();
  if (!filter) {
    sheet.getRange(CONFIG.INFO_HEADER_ROW, 1, 1, CONFIG.INFO_COLS_END).createFilter();
  }

  // Protect headers + info cols (optional; keep weeks editable)
  protectSheetLight_(sheet);

  props.setProperty(key, "1");
}

function protectSheetLight_(sheet) {
  // remove existing protections to avoid duplicates
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

  // protect header rows
  sheet.getRange("1:3").protect().addEditor(Session.getEffectiveUser());

  // protect info columns A..F only (weeks remain editable)
  sheet.getRange(1, 1, sheet.getMaxRows(), CONFIG.INFO_COLS_END)
    .protect()
    .addEditor(Session.getEffectiveUser());
}
