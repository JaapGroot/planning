function buildTeamFileCache_(ss) {
  const parent = DriveApp.getFileById(ss.getId()).getParents().next();
  const files = parent.getFiles();

  const cache = {};
  while (files.hasNext()) {
    const f = files.next();
    cache[f.getName()] = f.getId();
  }
  return { parentId: parent.getId(), byName: cache };
}

function openOrCreateTeamFileCached_(ss, teamName, cache) {
  const name = "Planning - " + teamName;

  // Exists -> open
  const existingId = cache.byName[name];
  if (existingId) return SpreadsheetApp.openById(existingId);

  if (!CONFIG.TEAM_TEMPLATE_SPREADSHEET_ID) {
    throw new Error("TEAM_TEMPLATE_SPREADSHEET_ID is niet ingesteld in Config.gs");
  }

  const templateFile = DriveApp.getFileById(CONFIG.TEAM_TEMPLATE_SPREADSHEET_ID);
  const parentFolder = DriveApp.getFolderById(cache.parentId);

  // Copy template (includes its container-bound Apps Script)
  const copiedFile = templateFile.makeCopy(name, parentFolder);
  const teamSS = SpreadsheetApp.openById(copiedFile.getId());

  // Rename template sheet to team name
  const templateSheet = teamSS.getSheetByName(CONFIG.TEAM_TEMPLATE_SHEET_NAME);
  if (!templateSheet) {
    throw new Error(`Template sheet '${CONFIG.TEAM_TEMPLATE_SHEET_NAME}' niet gevonden`);
  }
  templateSheet.setName(teamName);

  cache.byName[name] = copiedFile.getId();
  return teamSS;
}
