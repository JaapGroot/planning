const CONFIG = {
  // Master
  MASTER_TOTAL_COLS: 65,
  MASTER_WEEK_START_COL: 8,          // H
  MASTER_VALUES_COLS: 13,            // A..M

  // Team sheet layout
  INFO_COLS_END: 6,                  // A..F (info)
  WEEK_START_COL: 8,                 // H (weeks)
  TOTAL_COLS: 65,                    // keep same as master to preserve H.. alignment
  SPACER_COL_G: 7,                   // G stays empty & hidden

  TEAM_NAME_CELL: "A1",
  INFO_HEADER_ROW: 3,                // team header row
  MASTER_INFO_HEADER_ROW: 4,         // master header row
  DATA_START_ROW: 4,
  HEADER_ROWS: 3,
  PLANNING_SHEET: "Planning",
  LOG_SHEET_NAME: "Log",

  TEAM_SHEET_DEFAULTS: ["Blad1", "Sheet1"],

  // Master column indexes (0-based)
  COL_WERKNUMMER: 0,                 // A
  COL_OPDRACHT: 5,                   // F
  COL_TEAM: 11,                      // L

  // Mapping detail: Master A,H,I,J,K,M -> Team A..F
  MAP_MASTER_IDX_DETAIL: [0, 7, 8, 9, 10, 12],

  // Header override: master L -> team E (index 4) on header rows only
  MASTER_IDX_L: 11,

  EXCLUDED_TEAM_WORDS: ["team", "ja", "nee", "opdracht"],

  // Batch
  TEAMS_PER_RUN: 5,
  BATCH_TRIGGER_MINUTES: 2,

  // Properties
  PROP_BATCH_INDEX: "PTS_BATCH_INDEX",
  PROP_BATCH_TEAMS: "PTS_BATCH_TEAMS",
  PROP_SETUP_PREFIX: "PTS_SETUP_",
  PROP_LASTROW_PREFIX: "PTS_LASTROW_",

  // Template
  TEAM_TEMPLATE_SPREADSHEET_ID: "1u_UPDYRf4ccVr6XQeiC0RVuKgPabDIcJrn25tvE4aJc",
  TEAM_TEMPLATE_SHEET_NAME: "Teamsheet",

  // Column collapse in team sheet (H..M)
  COLLAPSE_START_COL: 8,             // H
  COLLAPSE_NUM_COLS: 6,              // H..M
};
