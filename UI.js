function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Team Planning")
    .addItem("Team bijwerken (dropdown)", "showTeamDropdown")
    .addItem("Alle teams (batch)", "syncAllTeamsBatched")
    .addToUi();

  SpreadsheetApp.getUi()
    .createMenu("Print")
    .addItem("Print 1 werknummer (A3)", "uiPrintSingle")
    .addItem("Batch: print alle locaties (PDF per locatie)", "uiPrintBatch")
    .addItem("Batch hervatten", "uiResumeBatch")
    .addToUi();
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
