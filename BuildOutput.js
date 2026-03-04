function buildDetailOutputMappedTeamOnly_(valuesAM, bgA, bgHtoK, bgM, bgL, bgWeeksMaster, fmt, fmtL, teamName, driveUrls) {
  const teamNorm = normalize_(teamName);

  const rows = [];
  const infoBgs = [];
  const weekBgs = [];
  const blockSizes = [];
  const skipped = [];

  // Formats for A..F
  const numFormats = [];
  const fontFamilies = [];
  const fontSizes = [];
  const fontWeights = [];
  const fontStyles = [];
  const fontColors = [];
  const hAligns = [];
  const vAligns = [];
  const wraps = [];

  let currentWn = null;
  let include = false;
  let header = null;
  let details = [];
  let hadTeamMatch = false;

  function resetBlock_() {
    currentWn = null;
    include = false;
    header = null;
    details = [];
    hadTeamMatch = false;
  }

  function pushOut_(valsAF, bgsAF, weekBgRow, fmtAF, headerDriveUrl) {
    const full = buildTeamRowFull_(valsAF);

    if (headerDriveUrl) {
      const safeUrl = String(headerDriveUrl).replace(/"/g, '""');
      full[CONFIG.TEAM_IDX_DRIVE_LINK_TARGET] = `=HYPERLINK("${safeUrl}";"${CONFIG.DRIVE_LINK_LABEL}")`; // 22 = kolom W
    }

    rows.push(full);
    infoBgs.push(bgsAF);
    weekBgs.push(weekBgRow);

    numFormats.push(fmtAF.num);
    fontFamilies.push(fmtAF.family);
    fontSizes.push(fmtAF.size);
    fontWeights.push(fmtAF.weight);
    fontStyles.push(fmtAF.style);
    fontColors.push(fmtAF.color);
    hAligns.push(fmtAF.hAlign);
    vAligns.push(fmtAF.vAlign);
    wraps.push(fmtAF.wrap);
  }

  function flushBlock_() {
    if (!currentWn) return;

    if (!include) {
      resetBlock_();
      return;
    }

    if (!hadTeamMatch || details.length === 0) {
      skipped.push([new Date(), teamName, currentWn, "No team match"]);
      resetBlock_();
      return;
    }

    pushOut_(
      header.valuesAF,
      header.bgsAF,
      new Array(bgWeeksMaster[0].length).fill("#ffffff"),
      header.fmtAF,
      header.driveUrl
    );

    for (const d of details) pushOut_(d.valuesAF, d.bgsAF, d.weekBg, d.fmtAF);

    blockSizes.push({ headerRowCount: 1, detailCount: details.length });
    resetBlock_();
  }

  for (let r = 0; r < valuesAM.length; r++) {
    const rowVals = valuesAM[r];
    const wn = rowVals[CONFIG.COL_WERKNUMMER];
    const opdracht = rowVals[CONFIG.COL_OPDRACHT];
    const team = rowVals[CONFIG.COL_TEAM];

    const isNewWerknummer = wn && wn !== currentWn;

    if (isNewWerknummer) {
      flushBlock_();

      currentWn = wn;
      include = typeof opdracht === "string" && normalize_(opdracht) === "ja";

      // KOPREGEL: A..F, but E = master L
      const valuesAF = mapMasterToTeamValuesDetail_(rowVals);
      valuesAF[4] = rowVals[CONFIG.MASTER_IDX_L]; // E = L (kopregel)

      const bgsAF = mapMasterToTeamBgsDetail_(r, bgA, bgHtoK, bgM);
      bgsAF[4] = bgL[r][0]; // E bg from L

      const fmtAF = mapMasterToTeamFmtDetail_(r, fmt);
      // override E formatting from L
      fmtAF.num[4] = fmtL.num[r][0];
      fmtAF.family[4] = fmtL.family[r][0];
      fmtAF.size[4] = fmtL.size[r][0];
      fmtAF.weight[4] = fmtL.weight[r][0];
      fmtAF.style[4] = fmtL.style[r][0];
      fmtAF.color[4] = fmtL.color[r][0];
      fmtAF.hAlign[4] = fmtL.hAlign[r][0];
      fmtAF.vAlign[4] = fmtL.vAlign[r][0];
      fmtAF.wrap[4] = fmtL.wrap[r][0];

      const driveUrl = (driveUrls && driveUrls[r]) ? driveUrls[r] : "";
      header = { valuesAF, bgsAF, fmtAF, driveUrl };

      if (!include) skipped.push([new Date(), teamName, wn, "Opdracht != Ja"]);
      continue;
    }

    if (!currentWn || !include) continue;

    const teamMatch = (typeof team === "string" && normalize_(team) === teamNorm);
    if (!teamMatch) continue;

    hadTeamMatch = true;

    const valuesAF = mapMasterToTeamValuesDetail_(rowVals);
    const bgsAF = mapMasterToTeamBgsDetail_(r, bgA, bgHtoK, bgM);
    const fmtAF = mapMasterToTeamFmtDetail_(r, fmt);

    details.push({
      valuesAF,
      bgsAF,
      fmtAF,
      weekBg: bgWeeksMaster[r],
    });
  }

  flushBlock_();

  return {
    rows,
    infoBgs,
    weekBgs,
    blockSizes,
    skipped,
    textFormats: { numFormats, fontFamilies, fontSizes, fontWeights, fontStyles, fontColors, hAligns, vAligns, wraps },
    headerRowIndexes: [],
  };
}

function mapMasterToTeamValuesDetail_(rowValsAM) {
  return CONFIG.MAP_MASTER_IDX_DETAIL.map(idx => rowValsAM[idx]);
}

function mapMasterToTeamBgsDetail_(r, bgA, bgHtoK, bgM) {
  return [bgA[r][0], bgHtoK[r][0], bgHtoK[r][1], bgHtoK[r][2], bgHtoK[r][3], bgM[r][0]];
}

function mapMasterToTeamFmtDetail_(r, fmt) {
  return {
    num: fmt.num[r].slice(),
    family: fmt.family[r].slice(),
    size: fmt.size[r].slice(),
    weight: fmt.weight[r].slice(),
    style: fmt.style[r].slice(),
    color: fmt.color[r].slice(),
    hAlign: fmt.hAlign[r].slice(),
    vAlign: fmt.vAlign[r].slice(),
    wrap: fmt.wrap[r].slice(),
  };
}

/**
 * Build a full row array length 65:
 * A..F = valuesAF
 * G    = "" (spacer)
 * H..  = ""
 */
function buildTeamRowFull_(valuesAF) {
  const row = new Array(CONFIG.TOTAL_COLS).fill("");
  for (let i = 0; i < CONFIG.INFO_COLS_END; i++) row[i] = valuesAF[i] ?? "";
  // row[6] (G) stays blank
  return row;
}
