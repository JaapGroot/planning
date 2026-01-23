function syncAllTeamsBatched() {
  withScriptLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planningSheet = ss.getSheetByName(CONFIG.PLANNING_SHEET);

    const teams = getTeamsFromPlanning_(planningSheet);
    if (!teams.length) return;

    const props = PropertiesService.getDocumentProperties();
    props.setProperty(CONFIG.PROP_BATCH_TEAMS, JSON.stringify(teams));
    props.setProperty(CONFIG.PROP_BATCH_INDEX, "0");

    cleanupBatchTriggers_();
    runNextBatch_();
  });
}

// Trigger handler
function runNextBatch_() {
  withScriptLock_(() => {
    const props = PropertiesService.getDocumentProperties();
    const teamsJson = props.getProperty(CONFIG.PROP_BATCH_TEAMS);
    if (!teamsJson) {
      cleanupBatchTriggers_();
      return;
    }

    const teams = JSON.parse(teamsJson);
    const index = parseInt(props.getProperty(CONFIG.PROP_BATCH_INDEX) || "0", 10);

    if (index >= teams.length) {
      clearBatchState_();
      cleanupBatchTriggers_();
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planningSheet = ss.getSheetByName(CONFIG.PLANNING_SHEET);
    const cache = buildTeamFileCache_(ss);

    const end = Math.min(index + CONFIG.TEAMS_PER_RUN, teams.length);

    for (let i = index; i < end; i++) {
      const team = teams[i];
      try {
        updateTeamFileDetail_(ss, planningSheet, team, cache);
      } catch (err) {
        Logger.log(`❌ Fout bij team ${team}: ${err && err.message ? err.message : err}`);
      }
    }

    props.setProperty(CONFIG.PROP_BATCH_INDEX, String(end));
    if (end < teams.length) scheduleNextBatchTrigger_();
    else {
      clearBatchState_();
      cleanupBatchTriggers_();
    }
  });
}

function scheduleNextBatchTrigger_() {
  cleanupBatchTriggers_();
  ScriptApp.newTrigger("runNextBatch_")
    .timeBased()
    .after(CONFIG.BATCH_TRIGGER_MINUTES * 60 * 1000)
    .create();
}

function cleanupBatchTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === "runNextBatch_") {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function clearBatchState_() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty(CONFIG.PROP_BATCH_INDEX);
  props.deleteProperty(CONFIG.PROP_BATCH_TEAMS);
}
