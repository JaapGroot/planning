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
  const existingId = cache.byName[name];
  if (existingId) return SpreadsheetApp.openById(existingId);

  const file = SpreadsheetApp.create(name);
  const created = DriveApp.getFileById(file.getId());

  const parent = DriveApp.getFolderById(cache.parentId);
  parent.addFile(created);
  DriveApp.getRootFolder().removeFile(created);

  cache.byName[name] = file.getId();
  return file;
}
