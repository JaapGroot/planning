function buildDetailOutputMappedTeamOnly_(valuesAM, bgA, bgHtoK, bgM, bgL, bgWeeksMaster, fmt, fmtL, teamName, afmeldenMap) {
  const teamNorm = normalize_(teamName);

  const rows = [];
  const infoBgs = [];
  const weekBgs = [];
  const blockSizes = [];
  const skipped = [];

  // Formats for info range A..G (per output row)
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

  function pushOut_(valuesAFG, bgsAFG, weekBg, fmtAFG) {
    rows.push(buildTeamRowFull_(valuesAFG));
    infoBgs.push(bgsAFG);
    weekBgs.push(weekBg);

    numFormats.push(fmtAFG.num);
    fontFamilies.push(fmtAFG.family);
    fontSizes.push(fmtAFG.size);
    fontWeights.push(fmtAFG.weight);
    fontStyles.push(fmtAFG.style);
    fontColors.push(fmtAFG.color);
    hAligns.push(fmtAFG.hAlign);
    vAligns.push(fmtAFG.vAlign);
    wraps.push(fmtAFG.wrap);
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

    // Kopregel: week leeg/wit, afmelden uit bestaande map
    const af = afmeldenMap[String(currentWn)];
    header.valuesAFG[6] = (af === true || af === false) ? af : ""; // col G

    pushOut_(
      header.valuesAFG,
      header.bgsAFG,
      new Array(bgWeeksMaster[0].length).fill("#ffffff"),
      header.fmtAFG
    );

    for (const d of details) pushOut_(d.valuesAFG, d.bgsAFG, d.weekBg, d.fmtAFG);

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

      // KOPREGEL: A..F but E = master L
      const valuesAF = mapMasterToTeamValuesDetail_(rowVals);
      valuesAF[4] = rowVals[CONFIG.MASTER_IDX_L];
      const valuesAFG = [...valuesAF, ""];

      // backgrounds
      const bgsAF = mapMasterToTeamBgsDetail_(r, bgA, bgHtoK, bgM);
      bgsAF[4] = bgL[r][0];
      const bgsAFG = [...bgsAF, "#ffffff"];

      // formats
      const fmtAF = mapMasterToTeamFmtDetail_(r, fmt);
      fmtAF.num[4] = fmtL.num[r][0];
      fmtAF.family[4] = fmtL.family[r][0];
      fmtAF.size[4] = fmtL.size[r][0];
      fmtAF.weight[4] = fmtL.weight[r][0];
      fmtAF.style[4] = fmtL.style[r][0];
      fmtAF.color[4] = fmtL.color[r][0];
      fmtAF.hAlign[4] = fmtL.hAlign[r][0];
      fmtAF.vAlign[4] = fmtL.vAlign[r][0];
      fmtAF.wrap[4] = fmtL.wrap[r][0];

      const fmtAFG = extendFmtToG_(fmtAF);

      header = { valuesAFG, bgsAFG, fmtAFG };

      if (!include) skipped.push([new Date(), teamName, wn, "Opdracht != Ja"]);
      continue;
    }

    if (!currentWn || !include) continue;

    const teamMatch = (typeof team === "string" && normalize_(team) === teamNorm);
    if (!teamMatch) continue;

    hadTeamMatch = true;

    const vAF = mapMasterToTeamValuesDetail_(rowVals);
    const vAFG = [...vAF, ""];

    const bAF = mapMasterToTeamBgsDetail_(r, bgA, bgHtoK, bgM);
    const bAFG = [...bAF, "#ffffff"];

    const fAF = mapMasterToTeamFmtDetail_(r, fmt);
    const fAFG = extendFmtToG_(fAF);

    details.push({
      valuesAFG: vAFG,
      bgsAFG: bAFG,
      fmtAFG: fAFG,
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
    headerRowIndexes: [], // set during write
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

function extendFmtToG_(fmtAF) {
  const i = 5; // base on F
  return {
    num: [...fmtAF.num, fmtAF.num[i]],
    family: [...fmtAF.family, fmtAF.family[i]],
    size: [...fmtAF.size, fmtAF.size[i]],
    weight: [...fmtAF.weight, fmtAF.weight[i]],
    style: [...fmtAF.style, fmtAF.style[i]],
    color: [...fmtAF.color, fmtAF.color[i]],
    hAlign: [...fmtAF.hAlign, fmtAF.hAlign[i]],
    vAlign: [...fmtAF.vAlign, fmtAF.vAlign[i]],
    wrap: [...fmtAF.wrap, fmtAF.wrap[i]],
  };
}

function buildTeamRowFull_(valuesAFG) {
  const row = valuesAFG.slice(0, CONFIG.INFO_COLS_END); // A..G
  while (row.length < CONFIG.TOTAL_COLS) row.push("");
  return row;
}
