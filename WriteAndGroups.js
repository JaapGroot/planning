function writeDetailToTeamSheetMapped_(file, sheet, detail) {
  const { rows, infoBgs, weekBgs, blockSizes, textFormats } = detail;
  const props = PropertiesService.getDocumentProperties();

  const lastKey = lastRowKey_(file, sheet);
  const prevLast = parseInt(props.getProperty(lastKey) || "0", 10);

  const clearRows = Math.max(prevLast, 80);
  const clearRange = sheet.getRange(CONFIG.DATA_START_ROW, 1, clearRows, CONFIG.TOTAL_COLS);
  clearRange.clearContent();
  clearRange.setBackground("#ffffff");
  clearRange.setFontColor("#000000");
  clearRowGroups_(sheet, CONFIG.DATA_START_ROW, clearRows);

  if (!rows.length) {
    props.setProperty(lastKey, "0");
    detail.headerRowIndexes = [];
    return;
  }

  // Ensure enough columns (template normally already has this)
  if (sheet.getMaxColumns() < CONFIG.TOTAL_COLS) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), CONFIG.TOTAL_COLS - sheet.getMaxColumns());
  }

  const fullRange = sheet.getRange(CONFIG.DATA_START_ROW, 1, rows.length, CONFIG.TOTAL_COLS);
  fullRange.setValues(rows);

  // Info A..F (no G)
  const infoRange = sheet.getRange(CONFIG.DATA_START_ROW, 1, rows.length, CONFIG.INFO_COLS_END);
  infoRange.setBackgrounds(infoBgs);

  infoRange.setNumberFormats(textFormats.numFormats);
  infoRange.setFontFamilies(textFormats.fontFamilies);
  infoRange.setFontSizes(textFormats.fontSizes);
  infoRange.setFontWeights(textFormats.fontWeights);
  infoRange.setFontStyles(textFormats.fontStyles);
  infoRange.setFontColors(textFormats.fontColors);
  infoRange.setHorizontalAlignments(textFormats.hAligns);
  infoRange.setVerticalAlignments(textFormats.vAligns);
  infoRange.setWrapStrategies(textFormats.wraps);

  // Week columns H.. (58 cols)
  const weekCols = CONFIG.TOTAL_COLS - CONFIG.WEEK_START_COL + 1; // 65-8+1=58
  const weekRange = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.WEEK_START_COL, rows.length, weekCols);

  const existing = weekRange.getBackgrounds();
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < weekCols; c++) {
      const master = weekBgs[r][c];
      const cur = existing[r][c];
      // keep rule: only overwrite if master has color AND target is white or same color
      if (master !== "#ffffff" && (cur === "#ffffff" || cur === master)) {
        existing[r][c] = master;
      }
    }
  }
  weekRange.setBackgrounds(existing);

  detail.headerRowIndexes = applyWerknummerGroupsExpandedAndReturnHeaderRows_(sheet, blockSizes);

  props.setProperty(lastKey, String(rows.length));
}

function applyWerknummerGroupsExpandedAndReturnHeaderRows_(sheet, blockSizes) {
  const headerRows = [];
  let cursor = CONFIG.DATA_START_ROW;

  for (const b of blockSizes) {
    headerRows.push(cursor);

    const detailStart = cursor + b.headerRowCount;
    const detailRows = b.detailCount;

    if (detailRows > 0) {
      sheet.getRange(detailStart, 1, detailRows, 1).shiftRowGroupDepth(1);
      const g = sheet.getRowGroup(detailStart, 1);
      if (g) g.expand();
    }

    cursor += b.headerRowCount + b.detailCount;
  }

  return headerRows;
}

function clearRowGroups_(sheet, startRow, numRows) {
  const endRow = startRow + Math.max(0, numRows) - 1;
  for (let r = startRow; r <= endRow; r++) {
    let depth = sheet.getRowGroupDepth(r);
    while (depth > 0) {
      sheet.getRange(r, 1, 1, 1).shiftRowGroupDepth(-1);
      depth = sheet.getRowGroupDepth(r);
    }
  }
}
