function readExistingAfmeldenMap_(teamSheet) {
  const map = {};
  const last = teamSheet.getLastRow();
  if (last < CONFIG.DATA_START_ROW) return map;

  const n = last - CONFIG.DATA_START_ROW + 1;
  const wnVals = teamSheet.getRange(CONFIG.DATA_START_ROW, 1, n, 1).getValues();
  const afVals = teamSheet.getRange(CONFIG.DATA_START_ROW, 7, n, 1).getValues(); // col G

  for (let i = 0; i < n; i++) {
    const wn = wnVals[i][0];
    const af = afVals[i][0];
    if (wn && (af === true || af === false)) map[String(wn)] = af;
  }
  return map;
}

function applyAfmeldenCheckboxes_(sheet, headerRowIndexes) {
  if (!headerRowIndexes || !headerRowIndexes.length) return;
  headerRowIndexes.forEach(r => sheet.getRange(r, 7, 1, 1).insertCheckboxes());
}
