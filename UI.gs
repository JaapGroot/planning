function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Team Planning")
    .addItem("Kolommen inklappen", "collapseTeamColumnsActiveSheet")
    .addItem("Kolommen uitklappen", "expandTeamColumnsActiveSheet")
    .addItem("Toggle kolommen", "toggleTeamColumnsActiveSheet")
    .addSeparator()
    .addItem("Team bijwerken (dropdown)", "showTeamDropdown")
    .addItem("Alle teams (batch)", "syncAllTeamsBatched")
    .addToUi();
}

function collapseTeamColumnsActiveSheet() {
  withScriptLock_(() => {
    const sh = SpreadsheetApp.getActiveSheet();
    sh.hideColumns(CONFIG.COLLAPSE_START_COL, CONFIG.COLLAPSE_NUM_COLS);
  });
}

function expandTeamColumnsActiveSheet() {
  withScriptLock_(() => {
    const sh = SpreadsheetApp.getActiveSheet();
    sh.showColumns(CONFIG.COLLAPSE_START_COL, CONFIG.COLLAPSE_NUM_COLS);
  });
}

function toggleTeamColumnsActiveSheet() {
  withScriptLock_(() => {
    const sh = SpreadsheetApp.getActiveSheet();
    const isHidden = sh.isColumnHiddenByUser(CONFIG.COLLAPSE_START_COL);
    if (isHidden) sh.showColumns(CONFIG.COLLAPSE_START_COL, CONFIG.COLLAPSE_NUM_COLS);
    else sh.hideColumns(CONFIG.COLLAPSE_START_COL, CONFIG.COLLAPSE_NUM_COLS);
  });
}

function showTeamDropdown() {
  withScriptLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planningSheet = ss.getSheetByName(CONFIG.PLANNING_SHEET);

    const teams = getTeamsFromPlanning_(planningSheet);
    if (!teams.length) {
      SpreadsheetApp.getUi().alert("Geen teamnamen gevonden.");
      return;
    }

    const html = `
    <html><body>
      <form onsubmit="google.script.run.startSingleTeamUpdate(this.team.value);google.script.host.close();return false;">
        <label>Selecteer een team:</label><br>
        <select name="team">
          ${teams.map(t => `<option value="${escapeHtml_(t)}">${escapeHtml_(t)}</option>`).join("")}
        </select>
        <br><br><input type="submit" value="Bijwerken">
      </form>
    </body></html>`;

    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(html).setWidth(320).setHeight(190),
      "Team bijwerken"
    );
  });
}
