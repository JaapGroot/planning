/**
 * Planning Team Splitter v1.0.4 - Production version
 * ✔️ Lost background colors fixed
 * ✔️ 'Opdracht = Nee' rows removed
 * ✔️ Prevents color-copy bleed
 * ✔️ Master white → clears team colors
 * ✔️ Column A–G formatting preserved
 */

const CONFIG = {
  TOTAL_COLS: 65,
  WEEK_START_COL: 8,
  INFO_COLS_END: 7,
  TEAM_NAME_CELL: "A1",
  INFO_HEADER_ROW: 3,
  DATA_START_ROW: 4,
  LOG_SHEET_NAME: "Log",
  HEADER_ROWS: 3,
  PLANNING_SHEET: "Planning",
  TEAM_SHEET_DEFAULTS: ["Blad1", "Sheet1"],
};

function updateTeamFile(teamName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planningSheet = ss.getSheetByName(CONFIG.PLANNING_SHEET);
  const valuesRange = planningSheet.getRange(
    CONFIG.DATA_START_ROW,
    1,
    planningSheet.getLastRow() - CONFIG.DATA_START_ROW + 1,
    CONFIG.TOTAL_COLS
  );
  const values = valuesRange.getValues();
  const backgrounds = valuesRange.getBackgrounds();

  const werkMap = buildWerknummerMap(values, backgrounds, teamName);
  const teamFile = openOrCreateTeamFile(ss, teamName);
  const teamSheet = prepareTeamSheet(teamFile, planningSheet, teamName);

  cleanObsoleteRows(teamSheet, werkMap);
  writePlanningData(teamSheet, werkMap, planningSheet);
  protectSheet(teamSheet);
  removeExtraRows(teamSheet);
  logSkipped(werkMap.skipped, teamFile, teamName);

  return Object.keys(werkMap.data).length;
}

function buildWerknummerMap(values, backgrounds, teamName) {
  const data = {}, skipped = [];
  let current = null, include = false, firstRow = null, rowIndex = 0;

  for (let r = 0; r < values.length; r++) {
    const wn = values[r][0], opdracht = values[r][5], team = values[r][11];

    if (wn && wn !== current) {
      current = wn;
      rowIndex = 1;
      include = typeof opdracht === "string" && opdracht.trim().toLowerCase() === "ja";
      firstRow = {
        data: values[r].slice(0, CONFIG.INFO_COLS_END),
        colors: backgrounds[r].slice(0, CONFIG.INFO_COLS_END)
      };
      if (!include) skipped.push([new Date(), teamName, wn, "Opdracht != Ja"]);
      continue;
    }

    rowIndex++;
    if (rowIndex < 2 || !include) continue;
    if (typeof team === "string" && team.trim().toLowerCase() === teamName.trim().toLowerCase()) {
      if (!data[current]) {
        data[current] = {
          infoCols: firstRow.data,
          colorCols: firstRow.colors,
          weeks: new Array(CONFIG.TOTAL_COLS - CONFIG.WEEK_START_COL + 1).fill("#ffffff")
        };
      }
      for (let c = 0; c < data[current].weeks.length; c++) {
        const col = backgrounds[r][CONFIG.WEEK_START_COL - 1 + c];
        if (col !== "#ffffff") data[current].weeks[c] = col;
      }
    } else {
      skipped.push([new Date(), teamName, current, "No team match"]);
    }
  }
  return { data, skipped };
}

function openOrCreateTeamFile(ss, teamName) {
  const files = DriveApp.getFilesByName("Planning - " + teamName);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());

  const file = SpreadsheetApp.create("Planning - " + teamName);
  const parent = DriveApp.getFileById(ss.getId()).getParents().next();
  parent.addFile(DriveApp.getFileById(file.getId()));
  DriveApp.getRootFolder().removeFile(DriveApp.getFileById(file.getId()));
  return file;
}

function prepareTeamSheet(file, planningSheet, teamName) {
  let sheet = file.getSheetByName(teamName);
  if (!sheet) sheet = file.insertSheet(teamName);
  CONFIG.TEAM_SHEET_DEFAULTS.forEach(name => {
    const def = file.getSheetByName(name);
    if (def) file.deleteSheet(def);
  });

  while (sheet.getMaxColumns() < CONFIG.TOTAL_COLS) sheet.insertColumnAfter(sheet.getMaxColumns());
  for (let c = 1; c <= CONFIG.TOTAL_COLS; c++) sheet.setColumnWidth(c, planningSheet.getColumnWidth(c));

  sheet.getRange(CONFIG.TEAM_NAME_CELL).setValue(teamName);
  sheet.setFrozenRows(CONFIG.HEADER_ROWS);
  sheet.setFrozenColumns(7);

  sheet.getRange(3, 1, 1, 7).setValues(planningSheet.getRange(4, 1, 1, 7).getDisplayValues());
  sheet.hideColumns(6);
  sheet.hideColumns(8, 6);

  const hdr = planningSheet.getRange(1, CONFIG.WEEK_START_COL, 3, CONFIG.TOTAL_COLS - CONFIG.WEEK_START_COL + 1);
  const dest = sheet.getRange(1, CONFIG.WEEK_START_COL, 3, CONFIG.TOTAL_COLS - CONFIG.WEEK_START_COL + 1);
  dest.setValues(hdr.getDisplayValues());
  dest.setBackgrounds(hdr.getBackgrounds());
  dest.setFontWeights(hdr.getFontWeights());
  dest.setFontSizes(hdr.getFontSizes());
  dest.setFontColors(hdr.getFontColors());
  dest.setHorizontalAlignments(hdr.getHorizontalAlignments());

  for (let row = 1; row <= 2; row++) {
    let mergeStart = null;
    const rowVals = sheet.getRange(row, CONFIG.WEEK_START_COL, 1, CONFIG.TOTAL_COLS - CONFIG.WEEK_START_COL + 1).getDisplayValues()[0];
    for (let c = 0; c <= rowVals.length; c++) {
      const val = rowVals[c], next = rowVals[c + 1];
      if (val && !mergeStart) mergeStart = CONFIG.WEEK_START_COL + c;
      if (mergeStart && (next || c === rowVals.length - 1)) {
        const mergeEnd = CONFIG.WEEK_START_COL + c;
        const numCols = mergeEnd - mergeStart + 1;
        if (numCols > 1) sheet.getRange(row, mergeStart, 1, numCols).mergeAcross();
        mergeStart = null;
      }
    }
  }
  return sheet;
}

function cleanObsoleteRows(sheet, werkMap) {
  if (sheet.getLastRow() <= CONFIG.INFO_HEADER_ROW) return;
  const existing = sheet.getRange(CONFIG.DATA_START_ROW, 1, sheet.getLastRow() - CONFIG.INFO_HEADER_ROW, 1).getValues().flat();
  for (let i = existing.length - 1; i >= 0; i--) {
    const wn = existing[i];
    const isInMaster = !!werkMap.data[wn];
    if (!isInMaster) sheet.deleteRow(i + CONFIG.DATA_START_ROW);
  }
}

function removeExtraRows(sheet) {
  const lastDataRow = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  if (lastDataRow < maxRows) {
    sheet.deleteRows(lastDataRow + 1, maxRows - lastDataRow);
  }
}

function writePlanningData(sheet, werkMap, planningSheet) {
  const entries = Object.keys(werkMap.data);
  const output = [], colors = [];

  entries.forEach(wn => {
    const row = [...werkMap.data[wn].infoCols];
    while (row.length < CONFIG.TOTAL_COLS) row.push("");
    output.push(row);

    const infoColors = werkMap.data[wn].colorCols;
    const weekColors = werkMap.data[wn].weeks;
    const fullRowColors = infoColors.concat(weekColors);
    while (fullRowColors.length < CONFIG.TOTAL_COLS) fullRowColors.push("#ffffff");
    colors.push(fullRowColors);
  });

  if (!output.length) return;
  const range = sheet.getRange(CONFIG.DATA_START_ROW, 1, output.length, CONFIG.TOTAL_COLS);
  range.setValues(output);

  const currentBgs = range.getBackgrounds();
  const fontColors = range.getFontColors();

  for (let r = 0; r < output.length; r++) {
    for (let c = 0; c < CONFIG.TOTAL_COLS; c++) {
      const master = colors[r][c], existing = currentBgs[r][c];
      if (c < CONFIG.WEEK_START_COL - 1) {
        currentBgs[r][c] = master; // Always preserve info colors
      }
      else if (master !== "#ffffff" && (existing === "#ffffff" || existing === master)) {
        currentBgs[r][c] = master;
      }
      fontColors[r][c] = "#ffffff";
    }
  }
  range.setBackgrounds(currentBgs);
  range.setFontColors(fontColors);

  const filter = sheet.getFilter();
  if (filter) filter.remove();
  sheet.getRange(CONFIG.INFO_HEADER_ROW, 1, 1, 7).createFilter();
}

function protectSheet(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
  sheet.getRange("1:3").protect().addEditor(Session.getEffectiveUser());
  sheet.getRange("A:G").protect().addEditor(Session.getEffectiveUser());
}

function logSkipped(skipped, file, team) {
  if (!skipped.length) return;
  let logSheet = file.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = file.insertSheet(CONFIG.LOG_SHEET_NAME);
    logSheet.hideSheet();
  } else {
    logSheet.clear();
  }
  logSheet.getRange(1, 1, 1, 4).setValues([["Timestamp", "Team", "Werknummer", "Reason"]]);
  logSheet.getRange(2, 1, skipped.length, 4).setValues(skipped);
}

function syncAllTeams() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.PLANNING_SHEET);
  const values = sheet.getRange(2, 12, sheet.getLastRow() - 1, 1).getValues().flat(); // Kolom L (team)
  const excluded = ["team", "ja", "nee", "opdracht"];

  const teams = [...new Set(values.filter(t =>
    typeof t === "string" &&
    !excluded.includes(t.trim().toLowerCase())
  ))].sort();

  teams.forEach(team => {
    try {
      const rows = updateTeamFile(team);
      Logger.log(`✅ ${team} bijgewerkt (${rows} regels)`);
    } catch (err) {
      Logger.log(`❌ Fout bij team ${team}: ${err.message}`);
    }
  });
}

function startSingleTeamUpdate(teamName) {
  const count = updateTeamFile(teamName);
  SpreadsheetApp.getUi().alert(`✅ Update voor ${teamName} voltooid\nAantal regels: ${count}\nVersie: v1.0.4`);
}

function showTeamDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planningSheet = ss.getSheetByName(CONFIG.PLANNING_SHEET);
  const teamCol = planningSheet.getRange(2, 12, planningSheet.getLastRow() - 1, 1).getValues(); // kolom L
  const excluded = ["team", "ja", "nee", "opdracht"];

  const teams = [...new Set(teamCol.flat()
    .filter(t => t && typeof t === "string" && !excluded.includes(t.trim().toLowerCase()))
  )].sort();

  if (!teams.length) {
    SpreadsheetApp.getUi().alert("Geen teamnamen gevonden.");
    return;
  }

  const html = `
    <html><body>
    <form onsubmit="google.script.run.startSingleTeamUpdate(this.team.value);google.script.host.close();return false;">
    <label>Selecteer een team:</label><br>
    <select name="team">
      ${teams.map(t => `<option value="${t}">${t}</option>`).join("")}
    </select>
    <br><br><input type="submit" value="Bijwerken">
    </form></body></html>`;

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(300).setHeight(160), "Team bijwerken");
}

