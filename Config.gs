const CONFIG = {
  MASTER_TOTAL_COLS: 65,
  MASTER_WEEK_START_COL: 8,         // H
  MASTER_VALUES_COLS: 13,           // A..M

  // Team sheet: A..F info + G Afmelden + H.. week
  INFO_COLS_END: 7,                 // A..G
  WEEK_START_COL: 8,                // H
  TOTAL_COLS: 7 + (65 - 8 + 1),     // 65

  TEAM_NAME_CELL: "A1",
  INFO_HEADER_ROW: 3,
  MASTER_INFO_HEADER_ROW: 4,
  DATA_START_ROW: 4,
  HEADER_ROWS: 3,
  PLANNING_SHEET: "Planning",
  LOG_SHEET_NAME: "Log",
  TEAM_SHEET_DEFAULTS: ["Blad1", "Sheet1"],

  COL_WERKNUMMER: 0,                // A
  COL_OPDRACHT: 5,                  // F
  COL_TEAM: 11,                     // L

  // detail: A,H,I,J,K,M -> A..F
  MAP_MASTER_IDX_DETAIL: [0, 7, 8, 9, 10, 12],

  // header override: master L -> team E (index 4)
  MASTER_IDX_L: 11,

  EXCLUDED_TEAM_WORDS: ["team", "ja", "nee", "opdracht"],

  // batch
  TEAMS_PER_RUN: 5,
  BATCH_TRIGGER_MINUTES: 2,

  // properties keys
  PROP_BATCH_INDEX: "PTS_BATCH_INDEX",
  PROP_BATCH_TEAMS: "PTS_BATCH_TEAMS",
  PROP_SETUP_PREFIX: "PTS_SETUP_",
  PROP_LASTROW_PREFIX: "PTS_LASTROW_",

  // column collapse (team sheet)
  // we keep Afmelden (G) visible, so collapse H..M
  COLLAPSE_START_COL: 8,            // H
  COLLAPSE_NUM_COLS: 6,             // H..M
};
