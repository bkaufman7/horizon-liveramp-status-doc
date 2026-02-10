/**
 * One-time initializer for your destination Google Sheet UI.
 * - Creates required tabs if missing
 * - Adds headers
 * - Applies clean header styling, freezes header rows, sets column widths
 * - Adds basic data validation placeholders (dropdowns, checkboxes)
 * - Adds conditional formatting for New/Updated/Ready-to-resolve flags
 *
 * Run: initHmiLiveRampUi()
 */
function initHmiLiveRampUi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Tabs to create
  const SHEETS = [
    "Config",
    "Recipients",
    "Raw_Alerts",
    "Working_Alerts",
    "Email_Log",
    "Push_Log",
    "README_Technical",
    "README_UserGuide",
  ];

  // Create any missing sheets
  const sheetsByName = {};
  SHEETS.forEach((name) => {
    sheetsByName[name] = ensureSheet_(ss, name);
  });

  // Apply per-tab layouts
  setupConfig_(sheetsByName["Config"]);
  setupRecipients_(sheetsByName["Recipients"]);
  setupRawAlerts_(sheetsByName["Raw_Alerts"]);
  setupWorkingAlerts_(sheetsByName["Working_Alerts"]);
  setupEmailLog_(sheetsByName["Email_Log"]);
  setupPushLog_(sheetsByName["Push_Log"]);
  setupReadme_(sheetsByName["README_Technical"], "README_Technical");
  setupReadme_(sheetsByName["README_UserGuide"], "README_UserGuide");

  // Optional: reorder sheets to match preferred order
  reorderSheets_(ss, SHEETS);

  // Optional: set active sheet
  ss.setActiveSheet(sheetsByName["Working_Alerts"]);

  SpreadsheetApp.flush();
}

/* --------------------------- Helpers: sheet mgmt --------------------------- */

function ensureSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function reorderSheets_(ss, orderedNames) {
  // Moves sheets in the given order to the front in sequence.
  // Safe if some names are missing (should not happen).
  orderedNames.forEach((name, idx) => {
    const sh = ss.getSheetByName(name);
    if (sh) ss.setActiveSheet(sh).moveActiveSheet(idx + 1);
  });
}

/* --------------------------- Helpers: styling ------------------------------ */

function styleHeaderRow_(sheet, headerRow, lastCol, opts) {
  const o = Object.assign(
    {
      bg: "#111827", // slate-ish
      fg: "#FFFFFF",
      fontSize: 11,
      bold: true,
      wrap: true,
      align: "center",
      vAlign: "middle",
      height: 32,
      border: true,
    },
    opts || {}
  );

  sheet.setFrozenRows(headerRow);
  sheet.setRowHeight(headerRow, o.height);

  const rng = sheet.getRange(headerRow, 1, 1, lastCol);
  rng
    .setBackground(o.bg)
    .setFontColor(o.fg)
    .setFontSize(o.fontSize)
    .setFontWeight(o.bold ? "bold" : "normal")
    .setWrap(o.wrap)
    .setHorizontalAlignment(o.align)
    .setVerticalAlignment(o.vAlign);

  if (o.border) {
    rng.setBorder(true, true, true, true, true, true, "#374151", SpreadsheetApp.BorderStyle.SOLID);
  }
}

function styleSectionTitle_(sheet, cellA1, text) {
  const rng = sheet.getRange(cellA1);
  rng.setValue(text);
  rng
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground("#F3F4F6")
    .setFontColor("#111827")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");
}

function applyBanding_(sheet, startRow, startCol, numRows, numCols) {
  // Remove existing banding
  const bandings = sheet.getBandings();
  bandings.forEach((b) => b.remove());

  const rng = sheet.getRange(startRow, startCol, Math.max(1, numRows), Math.max(1, numCols));
  const banding = rng.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  // Subtle tweaks
  banding.setHeaderRowColor("#111827");
  banding.setFirstRowColor("#FFFFFF");
  banding.setSecondRowColor("#F9FAFB");
}

function setColumnWidths_(sheet, widths) {
  // widths: array of pixel widths per column starting at 1
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

/* --------------------------- Config tab ----------------------------------- */

function setupConfig_(sheet) {
  sheet.clear();

  styleSectionTitle_(sheet, "A1", "HMI LiveRamp Tracker Configuration");
  sheet.getRange("A1:B1").merge();

  const rows = [
    ["LiveRamp_Sheet_URL", "https://docs.google.com/spreadsheets/d/1aNSO95dnIEBL5RfN1169TmPvUcXyt6VuCzeeGr6x58I/edit", "Source spreadsheet URL (LiveRamp-owned)"],
    ["LiveRamp_Tab_Name", "Alerts", "Source tab name in LiveRamp sheet"],
    ["Header_Row_Number", "1", "Header row number (keep as 1 unless needed)"],
    ["Timezone", "America/New_York", "Timezone for timestamps and triggers"],
    ["Daily_Email_Time", "07:30", "HH:MM (24h) time for morning email"],
    ["Auto_Push_Time", "23:00", "HH:MM (24h) time for end-of-day auto push"],
    ["Weekdays_Only", "TRUE", "TRUE to run only Mon-Fri"],
    ["Email_Subject_Prefix", "LiveRamp Alerts Daily Status", "Email subject prefix"],
    ["Ready_To_Be_Resolved_Label", "READY TO BE RESOLVED", "Special template label used for clearing"],
    ["Enable_Writeback_To_LR", "TRUE", "TRUE to enable pushing updates to LiveRamp column F"],
    ["LR_Columns_Fixed", "TRUE", "TRUE if LR columns are fixed as A:G in Alerts"],
  ];

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  // Labels styling
  sheet.getRange("A2:A").setFontWeight("bold").setFontColor("#111827");
  sheet.getRange("C2:C").setFontColor("#6B7280").setWrap(true);

  // Header styling for key/value/notes
  sheet.getRange("A2:C2").setBackground("#E5E7EB").setFontWeight("bold");
  sheet.getRange("A2").setValue("Key");
  sheet.getRange("B2").setValue("Value");
  sheet.getRange("C2").setValue("Notes");

  // Freeze header row 2 (the table header), keep title row 1 frozen as well
  sheet.setFrozenRows(2);

  // Column widths
  setColumnWidths_(sheet, [220, 520, 520]);

  // Data validation for booleans
  const boolRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["TRUE", "FALSE"], true)
    .setAllowInvalid(false)
    .build();

  // Apply to Weekdays_Only, Enable_Writeback_To_LR, LR_Columns_Fixed (rows are dynamic but we know positions)
  // Weekdays_Only row index = 2 + 6 = 8? Actually rows begin at 2 (title row 1, header row 2), data begins row 3.
  // Our data starts at row 3. Let's find by key.
  const dataRange = sheet.getRange(3, 1, rows.length, 2);
  const data = dataRange.getValues();
  for (let r = 0; r < data.length; r++) {
    const key = String(data[r][0] || "");
    if (["Weekdays_Only", "Enable_Writeback_To_LR", "LR_Columns_Fixed"].includes(key)) {
      sheet.getRange(3 + r, 2).setDataValidation(boolRule);
    }
  }

  // Nice borders
  const tableRange = sheet.getRange(2, 1, rows.length + 1, 3);
  tableRange.setBorder(true, true, true, true, true, true, "#D1D5DB", SpreadsheetApp.BorderStyle.SOLID);

  // Wrap notes
  sheet.getRange(2, 3, rows.length + 1, 1).setWrap(true);

  sheet.setTabColor("#111827");
}

/* --------------------------- Recipients tab ------------------------------- */

function setupRecipients_(sheet) {
  sheet.clear();

  styleSectionTitle_(sheet, "A1", "Email Recipients");
  sheet.getRange("A1:C1").merge();

  const headers = ["Name", "Email", "Active"];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, 2, headers.length, { bg: "#0F766E" }); // teal

  // Column widths
  setColumnWidths_(sheet, [220, 360, 90]);

  // Checkboxes for Active
  const activeRange = sheet.getRange(3, 3, 200, 1);
  activeRange.insertCheckboxes();

  // Sample row (optional)
  sheet.getRange(3, 1, 1, 3).setValues([["", "", true]]);

  applyBanding_(sheet, 2, 1, 250, 3);
  sheet.setTabColor("#0F766E");
}

/* --------------------------- Raw_Alerts tab ------------------------------- */

function setupRawAlerts_(sheet) {
  sheet.clear();

  styleSectionTitle_(sheet, "A1", "Raw Alerts (Mirror of LiveRamp Alerts)");
  sheet.getRange("A1:K1").merge();

  // LiveRamp columns A:G + helper columns
  const headers = [
    "Date",
    "Product",
    "Workflow/Audience Name",
    "Issue & LR Action",
    "Request to Horizon Team",
    "Horizon Comment (LR)",
    "Resolved (LR)",
    "_synced_at",
    "_thread_key",
    "_row_hash",
    "_source_row_number",
    "_first_seen_at",
    "_last_seen_at",
  ];

  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, 2, headers.length, { bg: "#374151" });

  // Column widths tuned for readability
  setColumnWidths_(sheet, [
    110, // Date
    140, // Product
    260, // Workflow
    360, // Issue
    320, // Request
    320, // Horizon Comment
    90,  // Resolved
    160, // synced
    160, // thread key
    160, // hash
    140, // source row
    160, // first seen
    160, // last seen
  ]);

  sheet.setFrozenRows(2);
  sheet.getRange(3, 1, 1, headers.length).setFontColor("#6B7280");
  applyBanding_(sheet, 2, 1, 250, headers.length);

  sheet.setTabColor("#374151");
}

/* --------------------------- Working_Alerts tab --------------------------- */

function setupWorkingAlerts_(sheet) {
  sheet.clear();

  styleSectionTitle_(sheet, "A1", "Working Alerts (Horizon Working Layer)");
  sheet.getRange("A1:R1").merge();

  // Base LR columns + local working columns
  const headers = [
    "Date",
    "Product",
    "Workflow/Audience Name",
    "Issue & LR Action",
    "Request to Horizon Team",
    "Horizon Comment (LR)",
    "Resolved (LR)",
    "HMI_Group", // New / Updated / Ongoing / Resolved
    "HMI_Update_Template",
    "HMI_Update_FreeText",
    "HMI_Composed_Update",
    "HMI_Push_Ready",
    "HMI_Last_Pushed_At",
    "HMI_Last_Pushed_Text",
    "HMI_Ready_To_Be_Resolved",
    "HMI_LR_Updated_After_Resolution",
    "HMI_Notes_Internal",
    "_thread_key",
    "_row_hash",
    "_source_row_number",
  ];

  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, 2, headers.length, { bg: "#1D4ED8" }); // blue

  // Column widths
  setColumnWidths_(sheet, [
    110, // Date
    140, // Product
    260, // Workflow
    360, // Issue
    320, // Request
    320, // LR comment
    90,  // Resolved
    110, // group
    210, // template
    260, // free text
    320, // composed
    110, // push ready
    160, // last pushed at
    240, // last pushed text
    160, // ready to be resolved
    190, // LR updated after resolution
    260, // notes
    160, // thread key
    160, // row hash
    140, // source row
  ]);

  sheet.setFrozenRows(2);

  // Wrap long text columns
  sheet.getRange(3, 4, 500, 3).setWrap(true); // Issue, Request, LR Comment
  sheet.getRange(3, 10, 500, 2).setWrap(true); // FreeText, Composed
  sheet.getRange(3, 14, 500, 1).setWrap(true); // Last pushed text
  sheet.getRange(3, 17, 500, 1).setWrap(true); // Notes

  // Data validations
  const groupRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["New", "Updated", "Ongoing", "Resolved"], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(3, 8, 500, 1).setDataValidation(groupRule);

  const templates = [
    "Investigating internally",
    "Waiting on client",
    "Waiting on site team",
    "Waiting on tagging team",
    "Need more info from LiveRamp",
    "Fix deployed, monitoring",
    "READY TO BE RESOLVED",
  ];
  const templateRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(templates, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(3, 9, 500, 1).setDataValidation(templateRule);

  // Checkboxes
  sheet.getRange(3, 12, 500, 1).insertCheckboxes(); // HMI_Push_Ready
  sheet.getRange(3, 15, 500, 1).insertCheckboxes(); // HMI_Ready_To_Be_Resolved
  sheet.getRange(3, 16, 500, 1).insertCheckboxes(); // HMI_LR_Updated_After_Resolution

  // Default formula for composed update (row-wise) in column K (11)
  // K = IF(AND(I not blank, J not blank), I & ", " & J, IF(I not blank, I, J))
  // Apply formula to multiple rows
  const startRow = 3;
  const numRows = 500;
  const formulas = [];
  for (let i = 0; i < numRows; i++) {
    const rowNum = startRow + i;
    formulas.push([`=IF(AND($I${rowNum}<>"",$J${rowNum}<>""),$I${rowNum}&", "&$J${rowNum},IF($I${rowNum}<>"",$I${rowNum},$J${rowNum}))`]);
  }
  sheet.getRange(startRow, 11, numRows, 1).setFormulas(formulas);

  // Apply banding
  applyBanding_(sheet, 2, 1, 600, headers.length);

  // Conditional formatting rules for quick scanning
  applyWorkingAlertsConditionalFormatting_(sheet, headers.length);

  sheet.setTabColor("#1D4ED8");
}

function applyWorkingAlertsConditionalFormatting_(sheet, lastCol) {
  // Columns:
  // HMI_Group = col 8 (H)
  // Template = col 9 (I)
  // Push Ready = col 12 (L)
  // Ready To Be Resolved = col 15 (O)
  // LR Updated After Resolution = col 16 (P)

  const dataRange = sheet.getRange(3, 1, 1000, lastCol);

  const rules = [];

  // New = light blue
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$H3="New"`)
      .setBackground("#DBEAFE")
      .setRanges([dataRange])
      .build()
  );

  // Updated = light yellow
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$H3="Updated"`)
      .setBackground("#FEF9C3")
      .setRanges([dataRange])
      .build()
  );

  // Ready to be resolved checkbox = TRUE -> light green
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$O3=TRUE`)
      .setBackground("#DCFCE7")
      .setRanges([dataRange])
      .build()
  );

  // LR updated after resolution -> light orange
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$P3=TRUE`)
      .setBackground("#FFEDD5")
      .setRanges([dataRange])
      .build()
  );

  // Push ready TRUE -> bold the row slightly by changing font color darker (Sheets API is limited, so do font weight)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$L3=TRUE`)
      .setFontWeight("bold")
      .setRanges([dataRange])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}

/* --------------------------- Email_Log tab -------------------------------- */

function setupEmailLog_(sheet) {
  sheet.clear();

  styleSectionTitle_(sheet, "A1", "Email Log");
  sheet.getRange("A1:H1").merge();

  const headers = [
    "Sent At",
    "Subject",
    "New Count",
    "Updated Count",
    "Ongoing Count",
    "Recipients",
    "Status",
    "Error",
  ];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, 2, headers.length, { bg: "#7C3AED" }); // purple

  setColumnWidths_(sheet, [170, 320, 90, 110, 110, 360, 110, 360]);
  applyBanding_(sheet, 2, 1, 500, headers.length);

  sheet.setTabColor("#7C3AED");
}

/* --------------------------- Push_Log tab --------------------------------- */

function setupPushLog_(sheet) {
  sheet.clear();

  styleSectionTitle_(sheet, "A1", "Push Log (Writebacks to LiveRamp)");
  sheet.getRange("A1:I1").merge();

  const headers = [
    "Pushed At",
    "Mode",
    "Rows Considered",
    "Rows Pushed",
    "Rows Skipped (No Change)",
    "Rows Skipped (Not Ready)",
    "Errors",
    "Notes",
    "Run ID",
  ];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, 2, headers.length, { bg: "#B91C1C" }); // red

  setColumnWidths_(sheet, [170, 90, 130, 110, 160, 160, 220, 320, 140]);
  applyBanding_(sheet, 2, 1, 500, headers.length);

  sheet.setTabColor("#B91C1C");
}

/* --------------------------- README tabs ---------------------------------- */

function setupReadme_(sheet, title) {
  sheet.clear();

  styleSectionTitle_(sheet, "A1", title);
  sheet.getRange("A1:H1").merge();

  sheet.getRange("A2").setValue(
    title === "README_Technical"
      ? "This tab will be generated/updated by the main Apps Script project. It will document the architecture, triggers, and sync/push logic."
      : "This tab will be generated/updated by the main Apps Script project. It will explain the daily workflow for Horizon users, and what LiveRamp updates are expected."
  );

  sheet.getRange("A2:H2").merge();
  sheet.getRange("A2")
    .setWrap(true)
    .setFontColor("#374151")
    .setVerticalAlignment("top");

  setColumnWidths_(sheet, [140, 140, 140, 140, 140, 140, 140, 140]);
  sheet.setFrozenRows(1);
  sheet.setTabColor("#6B7280");
}
