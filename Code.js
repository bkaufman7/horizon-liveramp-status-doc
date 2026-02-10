/**
 * =====================================================================
 * HMI LIVERAMP ALERTS TRACKER - MAIN APPS SCRIPT
 * =====================================================================
 * Version: 2.0
 * Last Updated: 2026-02-10
 * 
 * PURPOSE:
 * - Sync LiveRamp Alerts tab into our internal sheet (Raw_Alerts)
 * - Maintain a working layer (Working_Alerts) with New/Updated/Ongoing/Resolved states
 * - Let Horizon add structured updates via templates + free text
 * - Push updates back to LiveRamp column F (Horizon Comment) with append-only logic
 * - Send daily email with New/Updated/Ongoing status
 * - Auto-push updates at 11 PM ET weekdays
 * 
 * SETUP:
 * 1. Run initHmiLiveRampUi() from init.js to create all tabs
 * 2. Fill Config tab with your values
 * 3. Add Recipients
 * 4. Run setup() to install triggers
 * 5. Test with menu actions
 * 
 * ERROR NOTIFICATIONS:
 * - Errors will be sent to bkaufman@horizonmedia.com
 * =====================================================================
 */

// ==================== GLOBAL CONSTANTS ====================

const SCRIPT_VERSION = "2.0";
const ERROR_EMAIL = "bkaufman@horizonmedia.com";

// LiveRamp column indices (1-based)
const LR_COL = {
  DATE: 1,              // A
  PRODUCT: 2,           // B
  WORKFLOW: 3,          // C
  ISSUE: 4,             // D
  REQUEST: 5,           // E
  HORIZON_COMMENT: 6,   // F (we write here)
  RESOLVED: 7           // G (checkbox, we don't touch)
};

// Working_Alerts column names
const WA_COLS = {
  DATE: 0,
  PRODUCT: 1,
  WORKFLOW: 2,
  ISSUE: 3,
  REQUEST: 4,
  LR_COMMENT: 5,
  LR_RESOLVED: 6,
  GROUP: 7,                    // New/Updated/Ongoing/Resolved
  TEMPLATE: 8,
  FREE_TEXT: 9,
  COMPOSED: 10,
  PUSH_READY: 11,
  LAST_PUSHED_AT: 12,
  LAST_PUSHED_TEXT: 13,
  READY_TO_RESOLVE: 14,
  LR_UPDATED_AFTER_RES: 15,
  NOTES: 16,
  THREAD_KEY: 17,
  ROW_HASH: 18,
  SOURCE_ROW: 19
};

// ==================== ON OPEN MENU ====================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("HMI LiveRamp")
    .addItem("Refresh from LiveRamp", "menuRefresh")
    .addSeparator()
    .addItem("Build Email Preview (no send)", "menuEmailPreview")
    .addItem("Send Email Now", "menuSendEmail")
    .addItem("Refresh + Send Email Now", "menuRefreshAndEmail")
    .addSeparator()
    .addItem("Push Updates to LiveRamp Now", "menuPushUpdates")
    .addItem("Refresh + Push Updates Now", "menuRefreshAndPush")
    .addSeparator()
    .addItem("Refresh Template Dropdowns", "menuRefreshTemplates")
    .addSeparator()
    .addSubMenu(
      ui.createMenu("Admin")
        .addItem("Run Setup (Install Triggers)", "setup")
        .addItem("Run Diagnostics", "diagnostics")
        .addItem("Update README Tabs", "updateReadmeTabs")
    )
    .addToUi();
}

// ==================== MENU HANDLERS ====================

function menuRefresh() {
  try {
    syncFromLiveRamp();
    SpreadsheetApp.getUi().alert("✓ Refresh complete!");
  } catch (err) {
    SpreadsheetApp.getUi().alert("Error: " + err.message);
    sendErrorEmail("menuRefresh", err);
  }
}

function menuEmailPreview() {
  try {
    const html = buildEmailHtml_(true);
    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(900)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Email Preview");
  } catch (err) {
    SpreadsheetApp.getUi().alert("Error: " + err.message);
    sendErrorEmail("menuEmailPreview", err);
  }
}

function menuSendEmail() {
  try {
    sendDailyEmail();
    SpreadsheetApp.getUi().alert("✓ Email sent!");
  } catch (err) {
    SpreadsheetApp.getUi().alert("Error: " + err.message);
    sendErrorEmail("menuSendEmail", err);
  }
}

function menuRefreshAndEmail() {
  try {
    syncFromLiveRamp();
    sendDailyEmail();
    SpreadsheetApp.getUi().alert("✓ Refresh and email complete!");
  } catch (err) {
    SpreadsheetApp.getUi().alert("Error: " + err.message);
    sendErrorEmail("menuRefreshAndEmail", err);
  }
}

function menuPushUpdates() {
  try {
    pushUpdatesToLiveRamp();
    SpreadsheetApp.getUi().alert("✓ Push complete!");
  } catch (err) {
    SpreadsheetApp.getUi().alert("Error: " + err.message);
    sendErrorEmail("menuPushUpdates", err);
  }
}

function menuRefreshAndPush() {
  try {
    syncFromLiveRamp();
    pushUpdatesToLiveRamp();
    SpreadsheetApp.getUi().alert("✓ Refresh and push complete!");
  } catch (err) {
    SpreadsheetApp.getUi().alert("Error: " + err.message);
    sendErrorEmail("menuRefreshAndPush", err);
  }
}

function menuRefreshTemplates() {
  try {
    refreshTemplateDropdowns();
    SpreadsheetApp.getUi().alert("✓ Template dropdowns refreshed!");
  } catch (err) {
    SpreadsheetApp.getUi().alert("Error: " + err.message);
    sendErrorEmail("menuRefreshTemplates", err);
  }
}

// ==================== TRIGGER SETUP ====================

function setup() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    // Delete existing triggers to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => ScriptApp.deleteTrigger(t));
    
    const config = getConfig_();
    const tz = config.Timezone || "America/New_York";
    
    // Daily email trigger
    const emailTime = parseTime_(config.Daily_Email_Time || "07:30");
    ScriptApp.newTrigger("scheduledDailyEmail")
      .timeBased()
      .atHour(emailTime.hour)
      .nearMinute(emailTime.minute)
      .everyDays(1)
      .inTimezone(tz)
      .create();
    
    // Auto push trigger
    const pushTime = parseTime_(config.Auto_Push_Time || "23:00");
    ScriptApp.newTrigger("scheduledAutoPush")
      .timeBased()
      .atHour(pushTime.hour)
      .nearMinute(pushTime.minute)
      .everyDays(1)
      .inTimezone(tz)
      .create();
    
    // Daily sync trigger (runs earlier to ensure data is fresh)
    ScriptApp.newTrigger("scheduledDailySync")
      .timeBased()
      .atHour(emailTime.hour > 0 ? emailTime.hour - 1 : 23)
      .nearMinute(emailTime.minute)
      .everyDays(1)
      .inTimezone(tz)
      .create();
    
    Logger.log("✓ Triggers installed successfully");
    SpreadsheetApp.getUi().alert("✓ Triggers installed successfully!\n\n" +
      "- Daily sync: " + (emailTime.hour > 0 ? emailTime.hour - 1 : 23) + ":" + String(emailTime.minute).padStart(2, '0') + " ET\n" +
      "- Daily email: " + config.Daily_Email_Time + " ET\n" +
      "- Auto push: " + config.Auto_Push_Time + " ET");
  } catch (err) {
    Logger.log("Error in setup: " + err);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function parseTime_(timeStr) {
  const parts = String(timeStr).split(":");
  return {
    hour: parseInt(parts[0]) || 0,
    minute: parseInt(parts[1]) || 0
  };
}

// ==================== SCHEDULED TRIGGER HANDLERS ====================

function scheduledDailySync() {
  try {
    const config = getConfig_();
    if (config.Weekdays_Only === "TRUE" && isWeekend_()) {
      Logger.log("Skipping sync (weekend)");
      return;
    }
    syncFromLiveRamp();
  } catch (err) {
    Logger.log("Error in scheduledDailySync: " + err);
    sendErrorEmail("scheduledDailySync", err);
  }
}

function scheduledDailyEmail() {
  try {
    const config = getConfig_();
    if (config.Weekdays_Only === "TRUE" && isWeekend_()) {
      Logger.log("Skipping email (weekend)");
      return;
    }
    sendDailyEmail();
  } catch (err) {
    Logger.log("Error in scheduledDailyEmail: " + err);
    sendErrorEmail("scheduledDailyEmail", err);
  }
}

function scheduledAutoPush() {
  try {
    const config = getConfig_();
    if (config.Weekdays_Only === "TRUE" && isWeekend_()) {
      Logger.log("Skipping auto push (weekend)");
      return;
    }
    pushUpdatesToLiveRamp();
  } catch (err) {
    Logger.log("Error in scheduledAutoPush: " + err);
    sendErrorEmail("scheduledAutoPush", err);
  }
}

function isWeekend_() {
  const config = getConfig_();
  const now = new Date();
  const tz = config.Timezone || "America/New_York";
  const dayStr = Utilities.formatDate(now, tz, "u"); // 1=Mon, 7=Sun
  const day = parseInt(dayStr);
  return day === 6 || day === 7; // Sat or Sun
}

// ==================== CONFIG HELPER ====================

function getConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Config");
  if (!sheet) throw new Error("Config tab not found");
  
  const data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 2).getValues();
  const config = {};
  data.forEach(row => {
    if (row[0]) config[String(row[0])] = row[1];
  });
  return config;
}

// ==================== SYNC ENGINE ====================

function syncFromLiveRamp() {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(5000)) {
      throw new Error("Another sync is running. Please wait.");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig_();
    
    // Validate and open LiveRamp sheet
    const lrUrl = String(config.LiveRamp_Sheet_URL || "").trim();
    if (!lrUrl) {
      throw new Error("LiveRamp_Sheet_URL is not configured in Config tab");
    }
    if (!lrUrl.startsWith("https://docs.google.com/spreadsheets/")) {
      throw new Error("Invalid LiveRamp_Sheet_URL format. Must be a Google Sheets URL.");
    }
    
    const lrTabName = config.LiveRamp_Tab_Name || "Alerts";
    
    let lrSs, lrSheet;
    try {
      lrSs = SpreadsheetApp.openByUrl(lrUrl);
      lrSheet = lrSs.getSheetByName(lrTabName);
    } catch (e) {
      throw new Error("Cannot access LiveRamp sheet. Check URL and permissions: " + e.message);
    }
    
    if (!lrSheet) throw new Error("LiveRamp tab '" + lrTabName + "' not found");
    
    // Find header row
    const headerRow = parseInt(config.Header_Row_Number) || 1;
    const lrHeaders = lrSheet.getRange(headerRow, 1, 1, 7).getValues()[0];
    
    // Read all data below header
    const lastRow = lrSheet.getLastRow();
    if (lastRow <= headerRow) {
      Logger.log("No data in LiveRamp sheet");
      return;
    }
    
    const dataRange = lrSheet.getRange(headerRow + 1, 1, lastRow - headerRow, 7);
    const lrData = dataRange.getValues();
    
    // Get current timestamp
    const now = new Date();
    const tz = config.Timezone || "America/New_York";
    const syncTime = Utilities.formatDate(now, tz, "MM/dd/yyyy HH:mm:ss");
    
    // Load previous Raw_Alerts for history tracking
    const rawSheet = ss.getSheetByName("Raw_Alerts");
    if (!rawSheet) throw new Error("Raw_Alerts tab not found");
    
    const previousRawData = loadPreviousRawData_(rawSheet);
    
    // Process LiveRamp data
    const rawRows = [];
    lrData.forEach((row, idx) => {
      if (!row[0] && !row[1] && !row[2]) return; // skip empty rows
      
      const sourceRowNum = headerRow + 1 + idx;
      const threadKey = computeThreadKey_(row);
      const rowHash = computeRowHash_(row);
      
      // Check history
      const histKey = threadKey + "|" + sourceRowNum;
      const prev = previousRawData[histKey];
      const firstSeen = prev ? prev.firstSeen : syncTime;
      
      rawRows.push([
        row[0], // Date
        row[1], // Product
        row[2], // Workflow
        row[3], // Issue
        row[4], // Request
        row[5], // Horizon Comment
        row[6], // Resolved
        syncTime,
        threadKey,
        rowHash,
        sourceRowNum,
        firstSeen,
        syncTime // lastSeen
      ]);
    });
    
    // Write to Raw_Alerts
    rawSheet.clear();
    const rawHeaders = [
      "Date", "Product", "Workflow/Audience Name", "Issue & LR Action",
      "Request to Horizon Team", "Horizon Comment (LR)", "Resolved (LR)",
      "_synced_at", "_thread_key", "_row_hash", "_source_row_number",
      "_first_seen_at", "_last_seen_at"
    ];
    rawSheet.getRange(1, 1, 1, rawHeaders.length).setValues([rawHeaders]);
    
    if (rawRows.length > 0) {
      rawSheet.getRange(2, 1, rawRows.length, rawHeaders.length).setValues(rawRows);
    }
    
    // Update Working_Alerts
    updateWorkingAlerts_(ss, rawRows, config);
    
    // Refresh template dropdowns
    refreshTemplateDropdowns();
    
    Logger.log("✓ Sync complete: " + rawRows.length + " rows");
    
  } finally {
    lock.releaseLock();
  }
}

function loadPreviousRawData_(rawSheet) {
  const map = {};
  if (rawSheet.getLastRow() < 2) return map;
  
  const data = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 13).getValues();
  data.forEach(row => {
    const threadKey = row[8];
    const sourceRow = row[10];
    const firstSeen = row[11];
    const key = threadKey + "|" + sourceRow;
    map[key] = { firstSeen: firstSeen };
  });
  
  return map;
}

function computeThreadKey_(row) {
  // Hash of: Product + Workflow + Issue + Request (exclude Date)
  const str = String(row[1] || "") + "|" + String(row[2] || "") + "|" + 
              String(row[3] || "") + "|" + String(row[4] || "");
  return simpleHash_(str);
}

function computeRowHash_(row) {
  // Hash of all columns A:G including Date
  const str = row.slice(0, 7).map(c => String(c || "")).join("|");
  return simpleHash_(str);
}

function simpleHash_(str) {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32bit integer
  }
  return Math.abs(hash).toString(36);
}

// ==================== UPDATE WORKING ALERTS ====================

function updateWorkingAlerts_(ss, rawRows, config) {
  const waSheet = ss.getSheetByName("Working_Alerts");
  if (!waSheet) throw new Error("Working_Alerts tab not found");
  
  // Load existing Working_Alerts data
  const existingWA = loadExistingWorkingAlerts_(waSheet);
  
  // Process each raw row
  const waRows = [];
  const now = new Date();
  const tz = config.Timezone || "America/New_York";
  
  rawRows.forEach(rawRow => {
    const threadKey = rawRow[8];
    const rowHash = rawRow[9];
    const sourceRow = rawRow[10];
    const compositeKey = threadKey + "|" + sourceRow;
    
    const existing = existingWA[compositeKey];
    
    let group = "Ongoing";
    let template = "";
    let freeText = "";
    let composed = "";
    let pushReady = false;
    let lastPushedAt = "";
    let lastPushedText = "";
    let readyToResolve = false;
    let lrUpdatedAfterRes = false;
    let notes = "";
    
    if (existing) {
      // Row exists - check if updated
      if (existing.rowHash !== rowHash) {
        // Hash changed - mark as Updated
        if (existing.readyToResolve) {
          // Was resolved, now LR updated it
          group = "Resolved";
          lrUpdatedAfterRes = true;
        } else {
          group = "Updated";
        }
      } else {
        // No change
        group = existing.group;
      }
      
      // Preserve user fields
      template = existing.template;
      freeText = existing.freeText;
      composed = existing.composed;
      pushReady = existing.pushReady;
      lastPushedAt = existing.lastPushedAt;
      lastPushedText = existing.lastPushedText;
      readyToResolve = existing.readyToResolve;
      lrUpdatedAfterRes = existing.lrUpdatedAfterRes || lrUpdatedAfterRes;
      notes = existing.notes;
    } else {
      // New row
      group = "New";
    }
    
    waRows.push([
      rawRow[0], // Date
      rawRow[1], // Product
      rawRow[2], // Workflow
      rawRow[3], // Issue
      rawRow[4], // Request
      rawRow[5], // LR Comment
      rawRow[6], // LR Resolved
      group,
      template,
      freeText,
      composed,
      pushReady,
      lastPushedAt,
      lastPushedText,
      readyToResolve,
      lrUpdatedAfterRes,
      notes,
      threadKey,
      rowHash,
      sourceRow
    ]);
  });
  
  // Sort by thread_key, then Date, then source_row
  waRows.sort((a, b) => {
    if (a[17] !== b[17]) return a[17].localeCompare(b[17]); // thread_key
    if (a[0] !== b[0]) {
      const dateA = new Date(a[0]);
      const dateB = new Date(b[0]);
      return dateA - dateB;
    }
    return a[19] - b[19]; // source_row
  });
  
  // Write to Working_Alerts
  const waHeaders = [
    "Date", "Product", "Workflow/Audience Name", "Issue & LR Action",
    "Request to Horizon Team", "Horizon Comment (LR)", "Resolved (LR)",
    "HMI_Group", "HMI_Update_Template", "HMI_Update_FreeText",
    "HMI_Composed_Update", "HMI_Push_Ready", "HMI_Last_Pushed_At",
    "HMI_Last_Pushed_Text", "HMI_Ready_To_Be_Resolved",
    "HMI_LR_Updated_After_Resolution", "HMI_Notes_Internal",
    "_thread_key", "_row_hash", "_source_row_number"
  ];
  
  // Clear data area but keep header
  if (waSheet.getLastRow() > 2) {
    waSheet.getRange(3, 1, waSheet.getLastRow() - 2, waSheet.getLastColumn()).clear();
  }
  
  if (waRows.length > 0) {
    waSheet.getRange(3, 1, waRows.length, waHeaders.length).setValues(waRows);
    
    // Re-apply composed update formula for rows without manual overrides
    for (let i = 0; i < waRows.length; i++) {
      const rowNum = 3 + i;
      const template = waRows[i][8];
      const freeText = waRows[i][9];
      
      // Auto-compose if both fields have values or one has value
      let autoComposed = "";
      if (template && freeText) {
        autoComposed = template + ", " + freeText;
      } else if (template) {
        autoComposed = template;
      } else if (freeText) {
        autoComposed = freeText;
      }
      
      if (autoComposed) {
        waSheet.getRange(rowNum, 11).setValue(autoComposed);
      } else {
        // Use formula
        const formula = `=IF(AND($I${rowNum}<>"",$J${rowNum}<>""),$I${rowNum}&", "&$J${rowNum},IF($I${rowNum}<>"",$I${rowNum},$J${rowNum}))`;
        waSheet.getRange(rowNum, 11).setFormula(formula);
      }
    }
    
    // Apply visual formatting
    applyWorkingAlertsFormatting_(waSheet, waRows);
  }
  
  Logger.log("✓ Working_Alerts updated: " + waRows.length + " rows");
}

function loadExistingWorkingAlerts_(waSheet) {
  const map = {};
  if (waSheet.getLastRow() < 3) return map;
  
  const data = waSheet.getRange(3, 1, waSheet.getLastRow() - 2, 20).getValues();
  data.forEach(row => {
    const threadKey = row[17];
    const sourceRow = row[19];
    const compositeKey = threadKey + "|" + sourceRow;
    
    map[compositeKey] = {
      group: row[7],
      template: row[8],
      freeText: row[9],
      composed: row[10],
      pushReady: row[11],
      lastPushedAt: row[12],
      lastPushedText: row[13],
      readyToResolve: row[14],
      lrUpdatedAfterRes: row[15],
      notes: row[16],
      rowHash: row[18]
    };
  });
  
  return map;
}

function applyWorkingAlertsFormatting_(waSheet, waRows) {
  // Apply conditional formatting based on group
  for (let i = 0; i < waRows.length; i++) {
    const rowNum = 3 + i;
    const group = waRows[i][7];
    const readyToResolve = waRows[i][14];
    const lrUpdatedAfterRes = waRows[i][15];
    
    let bgColor = null;
    
    if (readyToResolve) {
      bgColor = "#DCFCE7"; // light green
    } else if (lrUpdatedAfterRes) {
      bgColor = "#FFEDD5"; // light orange
    } else if (group === "New") {
      bgColor = "#DBEAFE"; // light blue
    } else if (group === "Updated") {
      bgColor = "#FEF9C3"; // light yellow
    }
    
    if (bgColor) {
      waSheet.getRange(rowNum, 1, 1, 20).setBackground(bgColor);
    }
  }
}

// ==================== TEMPLATE DROPDOWNS ====================

function refreshTemplateDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templatesSheet = ss.getSheetByName("Templates");
  const waSheet = ss.getSheetByName("Working_Alerts");
  
  if (!templatesSheet) {
    // Create Templates tab if missing
    const sheet = ss.insertSheet("Templates");
    const headers = ["Template Label", "Active", "Category", "Notes"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Seed with default templates
    const templates = [
      ["Investigating internally", "TRUE", "Status", ""],
      ["Waiting on client", "TRUE", "Status", ""],
      ["Waiting on site team", "TRUE", "Status", ""],
      ["Waiting on tagging team", "TRUE", "Status", ""],
      ["Need more info from LiveRamp", "TRUE", "Status", ""],
      ["Fix deployed, monitoring", "TRUE", "Status", ""],
      ["READY TO BE RESOLVED", "TRUE", "Action", "Mark as complete"]
    ];
    sheet.getRange(2, 1, templates.length, 4).setValues(templates);
    
    // Style header
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground("#111827")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold");
    
    // Checkboxes for Active
    sheet.getRange(2, 2, 200, 1).insertCheckboxes();
    
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 4, 200);
    
    return refreshTemplateDropdowns(); // Recursive call after creation
  }
  
  // Load active templates
  const templateData = templatesSheet.getRange(2, 1, Math.max(1, templatesSheet.getLastRow() - 1), 2).getValues();
  const activeTemplates = [];
  
  templateData.forEach(row => {
    const label = String(row[0] || "").trim();
    const active = row[1] === true || row[1] === "TRUE";
    if (label && active) {
      activeTemplates.push(label);
    }
  });
  
  if (activeTemplates.length === 0) {
    Logger.log("Warning: No active templates found");
    return;
  }
  
  // Apply to Working_Alerts column I (HMI_Update_Template)
  if (!waSheet) return;
  
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(activeTemplates, true)
    .setAllowInvalid(true)
    .build();
  
  waSheet.getRange(3, 9, 500, 1).setDataValidation(rule);
  
  Logger.log("✓ Template dropdowns refreshed: " + activeTemplates.length + " templates");
}

// ==================== PUSH UPDATES TO LIVERAMP ====================

function pushUpdatesToLiveRamp() {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(5000)) {
      throw new Error("Another push is running. Please wait.");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig_();
    
    // Check if writeback is enabled
    if (config.Enable_Writeback_To_LR !== "TRUE") {
      Logger.log("Writeback disabled in Config");
      return;
    }
    
    const waSheet = ss.getSheetByName("Working_Alerts");
    if (!waSheet || waSheet.getLastRow() < 3) {
      Logger.log("No data to push");
      return;
    }
    
    // Open LiveRamp sheet
    const lrUrl = config.LiveRamp_Sheet_URL;
    const lrTabName = config.LiveRamp_Tab_Name || "Alerts";
    const lrSs = SpreadsheetApp.openByUrl(lrUrl);
    const lrSheet = lrSs.getSheetByName(lrTabName);
    if (!lrSheet) throw new Error("LiveRamp tab '" + lrTabName + "' not found");
    
    // Load Working_Alerts data
    const waData = waSheet.getRange(3, 1, waSheet.getLastRow() - 2, 20).getValues();
    
    const now = new Date();
    const tz = config.Timezone || "America/New_York";
    const pushTime = Utilities.formatDate(now, tz, "MM/dd/yyyy h:mm a") + " ET";
    
    let rowsConsidered = 0;
    let rowsPushed = 0;
    let rowsSkippedNoChange = 0;
    let rowsSkippedNotReady = 0;
    const errors = [];
    const updates = []; // Track for batch update
    
    waData.forEach((row, idx) => {
      const waRowNum = 3 + idx;
      const composed = String(row[10] || "").trim();
      const pushReady = row[11] === true;
      const lastPushedText = String(row[13] || "").trim();
      const readyToResolve = row[14] === true;
      const sourceRowNum = row[19];
      
      if (!composed) {
        return; // No message to push
      }
      
      rowsConsidered++;
      
      // Check push ready (Option B: READY TO BE RESOLVED doesn't require pushReady checkbox)
      const template = String(row[8] || "").trim();
      const isReadyToResolveTemplate = template === (config.Ready_To_Be_Resolved_Label || "READY TO BE RESOLVED");
      
      if (!pushReady && !isReadyToResolveTemplate) {
        rowsSkippedNotReady++;
        return;
      }
      
      // Check for duplicate
      const normalizedComposed = composed.toLowerCase().replace(/\s+/g, "");
      const normalizedLast = lastPushedText.toLowerCase().replace(/\s+/g, "");
      
      if (normalizedComposed === normalizedLast) {
        rowsSkippedNoChange++;
        return;
      }
      
      // Prepare update
      try {
        // Read existing LR cell F
        const lrCell = lrSheet.getRange(sourceRowNum, LR_COL.HORIZON_COMMENT);
        let existingText = String(lrCell.getValue() || "").trim();
        
        // Append new message
        let newText = existingText;
        if (existingText) {
          newText += "\n\n"; // blank line separator
        }
        newText += pushTime + "\n";
        newText += "Horizon: " + composed;
        
        // Write to LR
        lrCell.setValue(newText);
        
        // Update Working_Alerts with push timestamp and text
        waSheet.getRange(waRowNum, 13).setValue(pushTime); // Last Pushed At
        waSheet.getRange(waRowNum, 14).setValue(composed); // Last Pushed Text
        
        // If readyToResolve, mark the row with light green
        if (readyToResolve || isReadyToResolveTemplate) {
          waSheet.getRange(waRowNum, 1, 1, 20).setBackground("#DCFCE7");
          waSheet.getRange(waRowNum, 8).setValue("Resolved"); // Group
        }
        
        rowsPushed++;
        
      } catch (err) {
        errors.push("Row " + sourceRowNum + ": " + err.message);
      }
    });
    
    // Log to Push_Log
    logPush_(ss, now, tz, rowsConsidered, rowsPushed, rowsSkippedNoChange, rowsSkippedNotReady, errors);
    
    Logger.log("✓ Push complete: " + rowsPushed + " rows pushed");
    
  } finally {
    lock.releaseLock();
  }
}

function logPush_(ss, now, tz, considered, pushed, skippedNoChange, skippedNotReady, errors) {
  const pushLog = ss.getSheetByName("Push_Log");
  if (!pushLog) return;
  
  const timestamp = Utilities.formatDate(now, tz, "MM/dd/yyyy HH:mm:ss");
  const runId = Utilities.getUuid().substring(0, 8);
  
  const errorStr = errors.length > 0 ? errors.join("; ") : "";
  const notes = "Considered: " + considered + ", Pushed: " + pushed;
  
  const row = [
    timestamp,
    "Manual/Scheduled",
    considered,
    pushed,
    skippedNoChange,
    skippedNotReady,
    errorStr,
    notes,
    runId
  ];
  
  pushLog.appendRow(row);
}

// ==================== EMAIL BUILDER ====================

function sendDailyEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig_();
  
  // Get recipients
  const recipients = getActiveRecipients_(ss);
  if (recipients.length === 0) {
    Logger.log("No active recipients");
    return;
  }
  
  // Build email HTML
  const html = buildEmailHtml_(false);
  
  const now = new Date();
  const tz = config.Timezone || "America/New_York";
  const dateStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  const subject = (config.Email_Subject_Prefix || "LiveRamp Alerts Daily Status") + " | " + dateStr;
  
  // Count items
  const counts = getEmailCounts_(ss);
  
  // Send email
  try {
    MailApp.sendEmail({
      to: recipients.join(","),
      subject: subject,
      htmlBody: html
    });
    
    // Log to Email_Log
    logEmail_(ss, now, tz, subject, counts, recipients, "Sent", "");
    
    Logger.log("✓ Email sent to " + recipients.length + " recipients");
    
  } catch (err) {
    logEmail_(ss, now, tz, subject, counts, recipients, "Error", err.message);
    throw err;
  }
}

function getActiveRecipients_(ss) {
  const recipientsSheet = ss.getSheetByName("Recipients");
  if (!recipientsSheet || recipientsSheet.getLastRow() < 3) return [];
  
  const data = recipientsSheet.getRange(3, 1, recipientsSheet.getLastRow() - 2, 3).getValues();
  const emails = [];
  
  data.forEach(row => {
    const email = String(row[1] || "").trim();
    const active = row[2] === true;
    if (email && active) {
      emails.push(email);
    }
  });
  
  return emails;
}

function getEmailCounts_(ss) {
  const waSheet = ss.getSheetByName("Working_Alerts");
  if (!waSheet || waSheet.getLastRow() < 3) {
    return { newCount: 0, updatedCount: 0, ongoingCount: 0 };
  }
  
  const data = waSheet.getRange(3, 8, waSheet.getLastRow() - 2, 1).getValues(); // HMI_Group column
  
  let newCount = 0;
  let updatedCount = 0;
  let ongoingCount = 0;
  
  data.forEach(row => {
    const group = String(row[0] || "");
    if (group === "New") newCount++;
    else if (group === "Updated") updatedCount++;
    else if (group === "Ongoing") ongoingCount++;
  });
  
  return { newCount, updatedCount, ongoingCount };
}

function buildEmailHtml_(preview) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig_();
  const waSheet = ss.getSheetByName("Working_Alerts");
  
  if (!waSheet || waSheet.getLastRow() < 3) {
    return "<html><body><p>No alerts to report.</p></body></html>";
  }
  
  const now = new Date();
  const tz = config.Timezone || "America/New_York";
  const dateStr = Utilities.formatDate(now, tz, "MMMM dd, yyyy");
  
  // Get sheet URLs
  const horizonUrl = ss.getUrl();
  const lrUrl = config.LiveRamp_Sheet_URL;
  
  // Load data
  const data = waSheet.getRange(3, 1, waSheet.getLastRow() - 2, 17).getValues();
  
  // Categorize rows
  const newRows = [];
  const updatedRows = [];
  const ongoingRows = [];
  
  data.forEach(row => {
    const group = String(row[7] || "");
    const readyToResolve = row[14] === true;
    
    // Exclude resolved items from email
    if (group === "Resolved" || readyToResolve) return;
    
    if (group === "New") newRows.push(row);
    else if (group === "Updated") updatedRows.push(row);
    else if (group === "Ongoing") ongoingRows.push(row);
  });
  
  // Build HTML
  let html = `
<!DOCTYPE html>
<html>
<head>
<style>
body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
  font-size: 14px;
  line-height: 1.5;
  color: #111827;
  max-width: 1200px;
  margin: 0 auto;
  padding: 20px;
}
h1 {
  font-size: 24px;
  font-weight: 600;
  margin-bottom: 8px;
  color: #111827;
}
.date {
  font-size: 14px;
  color: #6B7280;
  margin-bottom: 20px;
}
.links {
  margin-bottom: 24px;
  padding: 12px;
  background: #F3F4F6;
  border-radius: 6px;
}
.links a {
  color: #1D4ED8;
  text-decoration: none;
  margin-right: 20px;
  font-weight: 500;
}
.links a:hover {
  text-decoration: underline;
}
.section {
  margin-bottom: 32px;
}
.section-title {
  font-size: 18px;
  font-weight: 600;
  margin-bottom: 12px;
  padding-bottom: 8px;
  border-bottom: 2px solid #E5E7EB;
}
.section-title.new { border-bottom-color: #3B82F6; }
.section-title.updated { border-bottom-color: #F59E0B; }
.section-title.ongoing { border-bottom-color: #6B7280; }
table {
  width: 100%;
  border-collapse: collapse;
  margin-top: 12px;
  border: 1px solid #E5E7EB;
}
th {
  background: #F9FAFB;
  padding: 10px;
  text-align: left;
  font-weight: 600;
  font-size: 13px;
  border: 1px solid #E5E7EB;
  color: #374151;
}
td {
  padding: 10px;
  border: 1px solid #E5E7EB;
  font-size: 13px;
  vertical-align: top;
}
tr.new-row { background: #DBEAFE; }
tr.updated-row { background: #FEF9C3; }
tr.ongoing-row { background: #F9FAFB; }
.no-items {
  color: #6B7280;
  font-style: italic;
  padding: 12px;
}
.footer {
  margin-top: 32px;
  padding-top: 16px;
  border-top: 1px solid #E5E7EB;
  color: #6B7280;
  font-size: 12px;
}
</style>
</head>
<body>
<h1>LiveRamp Alerts Daily Status</h1>
<div class="date">${dateStr}</div>

<div class="links">
<a href="${horizonUrl}" target="_blank">Horizon Internal Sheet</a>
<a href="${lrUrl}" target="_blank">LiveRamp Source Sheet</a>
</div>
`;
  
  // New section
  html += buildEmailSection_("New Alerts", newRows, "new", "new-row");
  
  // Updated section
  html += buildEmailSection_("Updated Alerts", updatedRows, "updated", "updated-row");
  
  // Ongoing section
  html += buildEmailSection_("Ongoing Alerts", ongoingRows, "ongoing", "ongoing-row");
  
  html += `
<div class="footer">
This is an automated report generated by the HMI LiveRamp Alerts Tracker.
<br>For questions or issues, contact <a href="mailto:${ERROR_EMAIL}">${ERROR_EMAIL}</a>
</div>
</body>
</html>
`;
  
  return html;
}

function buildEmailSection_(title, rows, sectionClass, rowClass) {
  let html = `<div class="section">
<div class="section-title ${sectionClass}">${title} (${rows.length})</div>`;
  
  if (rows.length === 0) {
    html += `<div class="no-items">No ${title.toLowerCase()} at this time.</div>`;
  } else {
    html += `
<table>
<thead>
<tr>
<th>Date</th>
<th>Product</th>
<th>Workflow/Audience</th>
<th>Issue & LR Action</th>
<th>Request to Horizon</th>
<th>LR Comment</th>
<th>Horizon Update</th>
</tr>
</thead>
<tbody>`;
    
    rows.forEach(row => {
      const date = formatDateForEmail_(row[0]);
      const product = escapeHtml_(row[1]);
      const workflow = escapeHtml_(row[2]);
      const issue = escapeHtml_(row[3]);
      const request = escapeHtml_(row[4]);
      const lrComment = escapeHtml_(row[5]);
      const template = escapeHtml_(row[8]);
      const freeText = escapeHtml_(row[9]);
      
      let horizonUpdate = "";
      if (template && freeText) {
        horizonUpdate = template + ", " + freeText;
      } else if (template) {
        horizonUpdate = template;
      } else if (freeText) {
        horizonUpdate = freeText;
      }
      horizonUpdate = escapeHtml_(horizonUpdate);
      
      html += `
<tr class="${rowClass}">
<td>${date}</td>
<td>${product}</td>
<td>${workflow}</td>
<td>${issue}</td>
<td>${request}</td>
<td>${lrComment}</td>
<td>${horizonUpdate}</td>
</tr>`;
    });
    
    html += `</tbody></table>`;
  }
  
  html += `</div>`;
  return html;
}

function formatDateForEmail_(dateVal) {
  if (!dateVal) return "";
  try {
    const d = new Date(dateVal);
    return Utilities.formatDate(d, "America/New_York", "MM/dd/yyyy");
  } catch (e) {
    return String(dateVal);
  }
}

function escapeHtml_(text) {
  if (!text) return "";
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;")
    .replace(/\n/g, "<br>");
}

function logEmail_(ss, now, tz, subject, counts, recipients, status, error) {
  const emailLog = ss.getSheetByName("Email_Log");
  if (!emailLog) return;
  
  const timestamp = Utilities.formatDate(now, tz, "MM/dd/yyyy HH:mm:ss");
  const recipientStr = recipients.join(", ");
  
  const row = [
    timestamp,
    subject,
    counts.newCount,
    counts.updatedCount,
    counts.ongoingCount,
    recipientStr,
    status,
    error
  ];
  
  emailLog.appendRow(row);
}

// ==================== ERROR HANDLING ====================

function sendErrorEmail(functionName, error) {
  try {
    const subject = "HMI LiveRamp Script Error: " + functionName;
    const body = "An error occurred in the HMI LiveRamp Alerts Tracker.\n\n" +
                 "Function: " + functionName + "\n" +
                 "Error: " + error.message + "\n\n" +
                 "Stack trace:\n" + error.stack;
    
    MailApp.sendEmail(ERROR_EMAIL, subject, body);
  } catch (e) {
    Logger.log("Failed to send error email: " + e.message);
  }
}

// ==================== DIAGNOSTICS ====================

function diagnostics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig_();
  
  let report = "=== HMI LIVERAMP DIAGNOSTICS ===\n\n";
  
  // Check config
  report += "CONFIG:\n";
  Object.keys(config).forEach(key => {
    report += "  " + key + ": " + config[key] + "\n";
  });
  
  // Check sheets
  report += "\nSHEETS:\n";
  const requiredSheets = ["Config", "Templates", "Recipients", "Raw_Alerts", 
                          "Working_Alerts", "Email_Log", "Push_Log", 
                          "README_Technical", "README_UserGuide"];
  requiredSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    report += "  " + name + ": " + (sheet ? "OK (" + sheet.getLastRow() + " rows)" : "MISSING") + "\n";
  });
  
  // Check triggers
  report += "\nTRIGGERS:\n";
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    report += "  No triggers installed. Run setup().\n";
  } else {
    triggers.forEach(t => {
      report += "  " + t.getHandlerFunction() + " - " + t.getEventType() + "\n";
    });
  }
  
  // Check LiveRamp access
  report += "\nLIVERAMP ACCESS:\n";
  try {
    const lrSs = SpreadsheetApp.openByUrl(config.LiveRamp_Sheet_URL);
    const lrSheet = lrSs.getSheetByName(config.LiveRamp_Tab_Name);
    report += "  LiveRamp sheet: OK (" + lrSheet.getLastRow() + " rows)\n";
  } catch (e) {
    report += "  LiveRamp sheet: ERROR - " + e.message + "\n";
  }
  
  // Check recipients
  report += "\nRECIPIENTS:\n";
  const recipients = getActiveRecipients_(ss);
  report += "  Active recipients: " + recipients.length + "\n";
  recipients.forEach(email => {
    report += "    - " + email + "\n";
  });
  
  Logger.log(report);
  SpreadsheetApp.getUi().alert("Diagnostics complete. Check Logs for full report.");
}

// ==================== README GENERATORS ====================

function updateReadmeTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  updateTechnicalReadme_(ss);
  updateUserGuideReadme_(ss);
  SpreadsheetApp.getUi().alert("✓ README tabs updated!");
}

function updateTechnicalReadme_(ss) {
  const sheet = ss.getSheetByName("README_Technical");
  if (!sheet) return;
  
  sheet.clear();
  
  const content = [
    ["HMI LIVERAMP ALERTS TRACKER - TECHNICAL README"],
    [""],
    ["VERSION", SCRIPT_VERSION],
    ["LAST UPDATED", new Date().toISOString()],
    [""],
    ["=== ARCHITECTURE ==="],
    [""],
    ["This system consists of three main layers:"],
    [""],
    ["1. SOURCE LAYER (LiveRamp-owned Google Sheet)"],
    ["   - Read-only source except columns F and G"],
    ["   - Tab: Alerts"],
    ["   - Columns A:G (Date, Product, Workflow, Issue, Request, Horizon Comment, Resolved)"],
    [""],
    ["2. DATA LAYER (This sheet)"],
    ["   - Raw_Alerts: Exact mirror of LiveRamp + metadata (_thread_key, _row_hash, etc.)"],
    ["   - Working_Alerts: Human working layer with Horizon control fields"],
    [""],
    ["3. CONTROL LAYER (Apps Script)"],
    ["   - Sync engine: LiveRamp -> Raw_Alerts -> Working_Alerts"],
    ["   - Push engine: Working_Alerts -> LiveRamp column F"],
    ["   - Email engine: Daily status email"],
    ["   - Template system: User-editable dropdown templates"],
    [""],
    ["=== KEY CONCEPTS ==="],
    [""],
    ["THREAD KEY"],
    ["- Stable hash of: Product + Workflow + Issue + Request (excludes Date)"],
    ["- Used to track same issue across multiple days"],
    ["- Allows deduping while maintaining visibility of repeats"],
    [""],
    ["ROW HASH"],
    ["- Hash of all columns A:G including Date"],
    ["- Used to detect changes in LiveRamp data"],
    ["- Triggers 'Updated' status when hash changes"],
    [""],
    ["STATUS GROUPS"],
    ["- New: Thread key not seen before"],
    ["- Updated: Thread key exists, but row hash changed"],
    ["- Ongoing: Thread key exists, no changes"],
    ["- Resolved: Marked as 'READY TO BE RESOLVED'"],
    [""],
    ["APPEND-ONLY LOGIC"],
    ["- Never overwrites existing text in LiveRamp column F"],
    ["- Appends with format:"],
    ["  <blank line>"],
    ["  MM/DD/YYYY h:mm AM/PM ET"],
    ["  Horizon: <message>"],
    ["- Skips duplicates (same message text)"],
    [""],
    ["POST-RESOLUTION UPDATES"],
    ["- If item marked resolved and LiveRamp updates it later:"],
    ["  * Stays in Resolved status"],
    ["  * HMI_LR_Updated_After_Resolution flag set to TRUE"],
    ["  * Row highlighted light orange"],
    [""],
    ["=== TRIGGERS ==="],
    [""],
    ["1. Daily Sync (runs 1 hour before email time)"],
    ["   - Syncs LiveRamp -> Raw_Alerts -> Working_Alerts"],
    ["   - Refreshes template dropdowns"],
    [""],
    ["2. Daily Email (configurable, default 7:30 AM ET)"],
    ["   - Sends HTML email with New/Updated/Ongoing tables"],
    ["   - Weekdays only (if configured)"],
    [""],
    ["3. Auto Push (configurable, default 11:00 PM ET)"],
    ["   - Pushes pending updates to LiveRamp column F"],
    ["   - Weekdays only (if configured)"],
    ["   - Safety net for items marked push-ready during the day"],
    [""],
    ["=== FUNCTIONS ==="],
    [""],
    ["Main Functions:"],
    ["- syncFromLiveRamp(): Sync data from LiveRamp"],
    ["- updateWorkingAlerts_(): Process sync data and update working layer"],
    ["- pushUpdatesToLiveRamp(): Push pending updates to LiveRamp"],
    ["- sendDailyEmail(): Build and send daily status email"],
    ["- refreshTemplateDropdowns(): Update dropdown validation from Templates tab"],
    [""],
    ["Setup Functions:"],
    ["- setup(): Install time-based triggers"],
    ["- diagnostics(): Run system health check"],
    ["- updateReadmeTabs(): Regenerate README content"],
    [""],
    ["=== TROUBLESHOOTING ==="],
    [""],
    ["Issue: Sync fails"],
    ["Solution: Check Config tab LiveRamp_Sheet_URL and permissions"],
    [""],
    ["Issue: Push doesn't work"],
    ["Solution: Verify Enable_Writeback_To_LR is TRUE and HMI_Push_Ready is checked"],
    [""],
    ["Issue: Email not sending"],
    ["Solution: Check Recipients tab has active recipients with valid emails"],
    [""],
    ["Issue: Template dropdown empty"],
    ["Solution: Go to Templates tab, ensure Active=TRUE for desired templates"],
    [""],
    ["Issue: Triggers not running"],
    ["Solution: Run setup() from Admin menu to reinstall triggers"],
    [""],
    ["=== ERROR NOTIFICATIONS ==="],
    [""],
    ["All script errors are automatically emailed to: " + ERROR_EMAIL],
    ["Check Email_Log and Push_Log tabs for execution history."]
  ];
  
  sheet.getRange(1, 1, content.length, Math.max(...content.map(r => r.length))).setValues(content);
  sheet.getRange("A1").setFontWeight("bold").setFontSize(14);
  sheet.getRange("A:A").setWrap(true);
  sheet.setColumnWidth(1, 800);
}

function updateUserGuideReadme_(ss) {
  const sheet = ss.getSheetByName("README_UserGuide");
  if (!sheet) return;
  
  sheet.clear();
  
  const content = [
    ["HMI LIVERAMP ALERTS TRACKER - USER GUIDE"],
    [""],
    ["=== DAILY WORKFLOW ==="],
    [""],
    ["1. REVIEW ALERTS"],
    ["   - Open Working_Alerts tab (automatically refreshed daily before email)"],
    ["   - Review color-coded rows:"],
    ["     * Light blue = New alert"],
    ["     * Light yellow = Updated by LiveRamp since last sync"],
    ["     * White = Ongoing (no changes)"],
    ["     * Light green = Ready to be resolved"],
    ["     * Light orange = LR updated after we marked resolved"],
    [""],
    ["2. ADD HORIZON UPDATES"],
    ["   - Column I (HMI_Update_Template): Select from dropdown"],
    ["     * Options come from Templates tab (you can add more)"],
    ["     * Common: 'Investigating internally', 'Waiting on client', etc."],
    ["   - Column J (HMI_Update_FreeText): Add custom details"],
    ["   - Column K (HMI_Composed_Update): Auto-generated preview"],
    ["     * Shows what will be sent to LiveRamp"],
    [""],
    ["3. MARK ITEMS FOR PUSH"],
    ["   - Column L (HMI_Push_Ready): Check the box for items ready to send"],
    ["   - Or use Column O (HMI_Ready_To_Be_Resolved) to mark as complete"],
    [""],
    ["4. PUSH UPDATES"],
    ["   - Manual: Menu > HMI LiveRamp > Push Updates to LiveRamp Now"],
    ["   - Automatic: System pushes at 11:00 PM ET weekdays"],
    ["   - Updates are appended to LiveRamp column F (Horizon Comment)"],
    ["   - Format: Timestamp + 'Horizon: <your message>'"],
    [""],
    ["=== EMAIL NOTIFICATIONS ==="],
    [""],
    ["You will receive a daily email (default 7:30 AM ET) with:"],
    ["- New Alerts table"],
    ["- Updated Alerts table"],
    ["- Ongoing Alerts table"],
    ["- Links to both Horizon sheet and LiveRamp sheet"],
    [""],
    ["To add/remove recipients:"],
    ["- Go to Recipients tab"],
    ["- Add Name, Email, and check Active box"],
    [""],
    ["=== TEMPLATE MANAGEMENT ==="],
    [""],
    ["To add or edit response templates:"],
    ["1. Go to Templates tab"],
    ["2. Add new row with Template Label, set Active=TRUE"],
    ["3. Menu > HMI LiveRamp > Refresh Template Dropdowns"],
    ["4. New template appears in Working_Alerts dropdown"],
    [""],
    ["Special template: 'READY TO BE RESOLVED'"],
    ["- Marks item as complete"],
    ["- Highlights row light green"],
    ["- Moves to Resolved status"],
    ["- Does NOT update LiveRamp column G (checkbox)"],
    [""],
    ["=== LIVERAMP INTERACTION ==="],
    [""],
    ["LiveRamp provides:"],
    ["- Column A: Date"],
    ["- Column B: Product"],
    ["- Column C: Workflow/Audience Name"],
    ["- Column D: Issue & LR Action"],
    ["- Column E: Request to Horizon Team"],
    ["- Column G: Resolved checkbox (they manage)"],
    [""],
    ["Horizon provides (via script):"],
    ["- Column F: Horizon Comment (our updates)"],
    ["  * We append, never overwrite"],
    ["  * Each update includes timestamp"],
    ["  * Format: MM/DD/YYYY h:mm AM/PM ET"],
    ["            Horizon: <message>"],
    [""],
    ["=== MENU ACTIONS ==="],
    [""],
    ["HMI LiveRamp menu:"],
    ["- Refresh from LiveRamp: Sync latest data"],
    ["- Build Email Preview: See email without sending"],
    ["- Send Email Now: Send status email immediately"],
    ["- Refresh + Send Email Now: Sync then email"],
    ["- Push Updates to LiveRamp Now: Send pending updates"],
    ["- Refresh + Push Updates Now: Sync then push"],
    ["- Refresh Template Dropdowns: Update dropdown options"],
    [""],
    ["Admin menu:"],
    ["- Run Setup (Install Triggers): Set up automation"],
    ["- Run Diagnostics: Check system health"],
    ["- Update README Tabs: Regenerate this guide"],
    [""],
    ["=== TIPS ==="],
    [""],
    ["- You can edit Working_Alerts freely (your columns are preserved)"],
    ["- Sync does NOT overwrite your HMI_* columns"],
    ["- If you need to re-push, just uncheck then re-check HMI_Push_Ready"],
    ["- Duplicate messages are automatically skipped"],
    ["- System handles ET timezone and daylight saving automatically"],
    ["- All actions are logged (Email_Log and Push_Log tabs)"],
    [""],
    ["=== SUPPORT ==="],
    [""],
    ["For questions or issues:"],
    ["Contact: " + ERROR_EMAIL],
    [""],
    ["System will automatically email errors to this address."]
  ];
  
  sheet.getRange(1, 1, content.length, Math.max(...content.map(r => r.length))).setValues(content);
  sheet.getRange("A1").setFontWeight("bold").setFontSize(14);
  sheet.getRange("A:A").setWrap(true);
  sheet.setColumnWidth(1, 800);
}