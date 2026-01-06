function refreshAllHeaderStylingEveryRun_(teamSheet, planningSheet) {
  refreshInfoHeadersEveryRun_(teamSheet, planningSheet);
  refreshWeekHeaderStylingEveryRun_(teamSheet, planningSheet);
}

function refreshInfoHeadersEveryRun_(teamSheet, planningSheet) {
  const hr = CONFIG.MASTER_INFO_HEADER_ROW;

  // Values (A,H..K,M) -> team row 3 A..F
  const hA = planningSheet.getRange(hr, 1, 1, 1).getDisplayValues()[0][0];
  const hHK = planningSheet.getRange(hr, 8, 1, 4).getDisplayValues()[0];
  const hM = planningSheet.getRange(hr, 13, 1, 1).getDisplayValues()[0][0];

  const targetAF = teamSheet.getRange(CONFIG.INFO_HEADER_ROW, 1, 1, 6);
  targetAF.setValues([[hA, hHK[0], hHK[1], hHK[2], hHK[3], hM]]);

  // Backgrounds for A..F
  const bgA = planningSheet.getRange(hr, 1, 1, 1).getBackgrounds()[0][0];
  const bgHK = planningSheet.getRange(hr, 8, 1, 4).getBackgrounds()[0];
  const bgM = planningSheet.getRange(hr, 13, 1, 1).getBackgrounds()[0][0];
  targetAF.setBackgrounds([[bgA, bgHK[0], bgHK[1], bgHK[2], bgHK[3], bgM]]);

  // Text style for A..F
  const headerFmt = readMappedTextFormats_(planningSheet, hr, 1);
  targetAF.setNumberFormats(headerFmt.num);
  targetAF.setFontFamilies(headerFmt.family);
  targetAF.setFontSizes(headerFmt.size);
  targetAF.setFontWeights(headerFmt.weight);
  targetAF.setFontStyles(headerFmt.style);
  targetAF.setFontColors(headerFmt.color);
  targetAF.setHorizontalAlignments(headerFmt.hAlign);
  targetAF.setVerticalAlignments(headerFmt.vAlign);
  targetAF.setWrapStrategies(headerFmt.wrap);

  // Column G header ("Afmelden") — style gelijk aan F
  const gCell = teamSheet.getRange(CONFIG.INFO_HEADER_ROW, 7, 1, 1);
  gCell.setValue("Afmelden");

  const fCell = teamSheet.getRange(CONFIG.INFO_HEADER_ROW, 6, 1, 1);
  gCell.setBackground(fCell.getBackground());
  gCell.setFontColor(fCell.getFontColor());
  gCell.setFontFamily(fCell.getFontFamily());
  gCell.setFontSize(fCell.getFontSize());
  gCell.setFontWeight(fCell.getFontWeight());
  gCell.setFontStyle(fCell.getFontStyle());
  gCell.setHorizontalAlignment(fCell.getHorizontalAlignment());
  gCell.setVerticalAlignment(fCell.getVerticalAlignment());
  gCell.setWrapStrategy(fCell.getWrapStrategy());
}

function refreshWeekHeaderStylingEveryRun_(teamSheet, planningSheet) {
  const masterWeekCols = CONFIG.MASTER_TOTAL_COLS - CONFIG.MASTER_WEEK_START_COL + 1;
  const src = planningSheet.getRange(1, CONFIG.MASTER_WEEK_START_COL, 3, masterWeekCols);
  const dst = teamSheet.getRange(1, CONFIG.WEEK_START_COL, 3, masterWeekCols);

  dst.setBackgrounds(src.getBackgrounds());
  dst.setFontWeights(src.getFontWeights());
  dst.setFontSizes(src.getFontSizes());
  dst.setFontColors(src.getFontColors());
  dst.setFontFamilies(src.getFontFamilies());
  dst.setHorizontalAlignments(src.getHorizontalAlignments());
  dst.setVerticalAlignments(src.getVerticalAlignments());
  dst.setWrapStrategies(src.getWrapStrategies());
}
