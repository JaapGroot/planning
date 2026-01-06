function getTeamsFromPlanning_(planningSheet) {
  const lastRow = planningSheet.getLastRow();
  if (lastRow < 2) return [];

  const rows = planningSheet.getRange(2, 1, lastRow - 1, 12).getValues(); // A..L
  const excluded = new Set(CONFIG.EXCLUDED_TEAM_WORDS.map(s => s.toLowerCase()));
  const teamsSet = new Set();

  let prevWn = null;
  for (let i = 0; i < rows.length; i++) {
    const wn = rows[i][CONFIG.COL_WERKNUMMER];
    const team = rows[i][CONFIG.COL_TEAM];

    const isHeaderRow = wn && wn !== prevWn;
    if (wn) prevWn = wn;
    if (isHeaderRow) continue;

    if (typeof team === "string") {
      const t = team.trim();
      if (t && !excluded.has(t.toLowerCase())) teamsSet.add(t);
    }
  }
  return [...teamsSet].sort();
}
