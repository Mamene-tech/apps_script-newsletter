/**
 * Code.js
 * Bulletproof Controller for Google Apps Script Newsletter Studio.
 * 
 * Features:
 * - 4.5-Minute Execution Watchdog: Automatically pauses and chains a 2-minute trigger before the 6-minute GAS cutoff
 * - 24.5-Hour Safety Offset: 30-minute buffer to ensure Google's rolling 24h quota window has fully restored
 * - Active Quota Sniffing: If quota drops <= 5, pauses and auto-reschedules rather than crashing
 * - Row-by-Row "ODESLÁNO" stamping: Real-time progress tracking, zero duplicate sends
 * - Strict Unsubscribe Protection: Permanently skips anyone marked "ODHLÁŠENO"
 * - Drive Assets Library: Lists previously uploaded images for the user-friendly Image Picker
 */

// ==========================================
// CONFIGURATION
// ==========================================
var CONFIG = {
  ADMIN_PASSWORD: "admin", // Password to access the web portal
  ADMIN_EMAIL: "", // Notification email (if empty, uses active account email)
  DEFAULT_SENDER_NAME: "SHM Uhříněves - Kolovraty",
  ASSETS_FOLDER_NAME: "Newsletter Images",
  
  // Timeout Watchdog: Stop at 4.5 minutes (270s) to prevent 6-minute Google crash
  MAX_RUNTIME_MS: 270 * 1000, 
  
  // Quota Safety Threshold: Stop if remaining quota <= 5
  SAFETY_QUOTA_THRESHOLD: 5,
  
  // Group 2 delay: 24 hours + 30 minutes buffer (88,200,000 ms)
  GROUP2_DELAY_MS: (24 * 60 + 30) * 60 * 1000,
  
  // Test simulation delay: 2 minutes (120,000 ms)
  SIMULATION_DELAY_MS: 2 * 60 * 1000
};

// ==========================================
// WEB APP ROUTING (doGet)
// ==========================================
function doGet(e) {
  if (e && e.parameter && e.parameter.action === "unsub") {
    var email = e.parameter.email || "";
    var template = HtmlService.createTemplateFromFile("UnsubPage");
    template.email = email;
    return template.evaluate()
      .setTitle("Odhlášení z odběru")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  var html = HtmlService.createTemplateFromFile("Index").evaluate();
  return html
    .setTitle("Newsletter Studio")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

function verifyPassword(password) {
  return password === CONFIG.ADMIN_PASSWORD;
}

// ==========================================
// SYSTEM INFO & REAL-TIME DASHBOARD
// ==========================================
function getSystemInfo() {
  var ss = getSpreadsheet();
  initializeSheets(ss);

  var g1Sheet = ss.getSheetByName("Group 1");
  var g2Sheet = ss.getSheetByName("Group 2");
  var unSheet = ss.getSheetByName("Unsubscribed");
  var campSheet = ss.getSheetByName("_Active_Campaign");

  var g1Total = Math.max(0, g1Sheet.getLastRow() - 1);
  var g2Total = Math.max(0, g2Sheet.getLastRow() - 1);
  var unTotal = Math.max(0, unSheet.getLastRow() - 1);

  var g1Sent = countStatus(g1Sheet, "ODESLÁNO");
  var g2Sent = countStatus(g2Sheet, "ODESLÁNO");

  var activeCamp = null;
  if (campSheet.getLastRow() >= 2) {
    var data = campSheet.getRange(2, 1, 1, 7).getValues()[0];
    if (data[0]) {
      activeCamp = {
        subject: data[0],
        status: data[1],
        createdTime: data[2],
        scheduledTime: data[3],
        isSimulation: data[4],
        triggerId: data[5]
      };
    }
  }

  var triggers = ScriptApp.getProjectTriggers();
  var pendingTrigger = false;
  var triggerType = "";
  for (var i = 0; i < triggers.length; i++) {
    var handler = triggers[i].getHandlerFunction();
    if (handler === "executeGroup2Trigger" || handler === "executeGroup1WatchdogTrigger" || handler === "executeGroup2WatchdogTrigger") {
      pendingTrigger = true;
      triggerType = handler;
      break;
    }
  }

  return {
    dailyQuotaRemaining: MailApp.getRemainingDailyQuota(),
    userEmail: Session.getEffectiveUser().getEmail(),
    webAppUrl: getWebAppUrl(),
    group1Count: g1Total,
    group1Sent: g1Sent,
    group2Count: g2Total,
    group2Sent: g2Sent,
    unsubscribedCount: unTotal,
    activeCampaign: activeCamp,
    hasPendingGroup2Trigger: pendingTrigger,
    pendingTriggerType: triggerType
  };
}

function countStatus(sheet, keyword) {
  if (!sheet || sheet.getLastRow() <= 1) return 0;
  var rows = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
  var count = 0;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] && rows[i][0].toString().indexOf(keyword) !== -1) {
      count++;
    }
  }
  return count;
}

// ==========================================
// SPREADSHEET INITIALIZATION (Email, Jméno, Stav)
// ==========================================
function getSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Tento skript musí být vložen do Google Tabulky (Rozšíření > Apps Script).");
  }
  return ss;
}

function initializeSheets(ss) {
  if (!ss) ss = getSpreadsheet();

  // 1. Group 1
  var g1 = ss.getSheetByName("Group 1");
  if (!g1) {
    g1 = ss.insertSheet("Group 1");
    g1.appendRow(["Email", "Jméno", "Stav"]);
    g1.getRange("A1:C1").setFontWeight("bold").setBackground("#298DC0").setFontColor("#ffffff");
    g1.setFrozenRows(1);
  } else if (g1.getRange("C1").getValue() === "") {
    g1.getRange("C1").setValue("Stav").setFontWeight("bold").setBackground("#298DC0").setFontColor("#ffffff");
  }

  // 2. Group 2
  var g2 = ss.getSheetByName("Group 2");
  if (!g2) {
    g2 = ss.insertSheet("Group 2");
    g2.appendRow(["Email", "Jméno", "Stav"]);
    g2.getRange("A1:C1").setFontWeight("bold").setBackground("#298DC0").setFontColor("#ffffff");
    g2.setFrozenRows(1);
  } else if (g2.getRange("C1").getValue() === "") {
    g2.getRange("C1").setValue("Stav").setFontWeight("bold").setBackground("#298DC0").setFontColor("#ffffff");
  }

  // 3. Unsubscribed
  var un = ss.getSheetByName("Unsubscribed");
  if (!un) {
    un = ss.insertSheet("Unsubscribed");
    un.appendRow(["Email", "Datum odhlášení"]);
    un.getRange("A1:B1").setFontWeight("bold").setBackground("#e53e3e").setFontColor("#ffffff");
    un.setFrozenRows(1);
  }

  // 4. _Active_Campaign
  var camp = ss.getSheetByName("_Active_Campaign");
  if (!camp) {
    camp = ss.insertSheet("_Active_Campaign");
    camp.appendRow(["Subject", "Status", "Created Time", "Scheduled Time", "Is Simulation", "Trigger ID", "Data JSON"]);
    camp.getRange("A1:G1").setFontWeight("bold").setBackground("#718096").setFontColor("#ffffff");
    camp.setFrozenRows(1);
  }
}

// ==========================================
// UNSUBSCRIBE CONFIRMATION (2-Step & Sheet Updates)
// ==========================================
function confirmUnsubscribe(email) {
  if (!email || email.trim() === "") return { success: false, error: "Prázdný email" };
  email = email.trim().toLowerCase();

  var ss = getSpreadsheet();
  initializeSheets(ss);

  var unSheet = ss.getSheetByName("Unsubscribed");
  var lastRow = unSheet.getLastRow();
  var alreadyInUnsub = false;

  if (lastRow > 1) {
    var existing = unSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < existing.length; i++) {
      if (existing[i][0] && existing[i][0].toString().trim().toLowerCase() === email) {
        alreadyInUnsub = true;
        break;
      }
    }
  }

  if (!alreadyInUnsub) {
    unSheet.appendRow([email, new Date()]);
  }

  markStatusInGroupSheet(ss.getSheetByName("Group 1"), email, "ODHLÁŠENO");
  markStatusInGroupSheet(ss.getSheetByName("Group 2"), email, "ODHLÁŠENO");

  return { success: true, email: email };
}

function markStatusInGroupSheet(sheet, email, statusText) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  var emails = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < emails.length; i++) {
    if (emails[i][0] && emails[i][0].toString().trim().toLowerCase() === email) {
      sheet.getRange(i + 2, 3).setValue(statusText).setFontColor("#e53e3e").setFontWeight("bold");
    }
  }
}

function getUnsubscribedSet() {
  var ss = getSpreadsheet();
  var set = {};

  var unSheet = ss.getSheetByName("Unsubscribed");
  if (unSheet && unSheet.getLastRow() > 1) {
    var unRows = unSheet.getRange(2, 1, unSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < unRows.length; i++) {
      if (unRows[i][0]) set[unRows[i][0].toString().trim().toLowerCase()] = true;
    }
  }

  ["Group 1", "Group 2"].forEach(function(sheetName) {
    var s = ss.getSheetByName(sheetName);
    if (s && s.getLastRow() > 1) {
      var rows = s.getRange(2, 1, s.getLastRow() - 1, 3).getValues();
      for (var j = 0; j < rows.length; j++) {
        var em = rows[j][0] ? rows[j][0].toString().trim().toLowerCase() : "";
        var st = rows[j][2] ? rows[j][2].toString().trim().toUpperCase() : "";
        if (em && (st === "ODHLÁŠENO" || st === "ODHLASENO" || st === "UNSUBSCRIBED")) {
          set[em] = true;
        }
      }
    }
  });

  return set;
}

// ==========================================
// IMAGE UPLOAD & DRIVE ASSETS LIBRARY
// ==========================================
function uploadImage(base64Data, fileName, mimeType) {
  try {
    var folder;
    var folders = DriveApp.getFoldersByName(CONFIG.ASSETS_FOLDER_NAME);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(CONFIG.ASSETS_FOLDER_NAME);
      try {
        folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (shareErr) {
        Logger.log("Folder public sharing warning: " + shareErr);
      }
    }

    // Robust base64 extraction regardless of data: prefix
    var cleanBase64 = base64Data.indexOf(",") !== -1 ? base64Data.split(",")[1] : base64Data;
    cleanBase64 = cleanBase64.replace(/\s/g, ""); // strip whitespace/newlines
    var decoded = Utilities.base64Decode(cleanBase64);
    var blob = Utilities.newBlob(decoded, mimeType || "image/jpeg", fileName || "newsletter_image.jpg");

    var file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log("File public sharing warning: " + shareErr);
    }

    var fileId = file.getId();
    // Use lh3 CDN URL format which is the standard, working direct embedding URL for Google Drive
    var directUrl = "https://lh3.googleusercontent.com/d/" + fileId;

    return { 
      success: true, 
      url: directUrl, 
      fileId: fileId, 
      name: file.getName() 
    };
  } catch (err) {
    Logger.log("uploadImage error: " + err);
    return { success: false, error: err.toString() };
  }
}

/**
 * Returns a list of recently uploaded images from the Google Drive "Newsletter Images" folder
 * so the user can visually choose previously uploaded pictures!
 */
function getDriveAssets() {
  try {
    var images = [];
    var folders = DriveApp.getFoldersByName(CONFIG.ASSETS_FOLDER_NAME);
    if (!folders.hasNext()) {
      return { success: true, images: [] };
    }

    var folder = folders.next();
    var files = folder.getFiles();
    var count = 0;

    while (files.hasNext() && count < 30) {
      var file = files.next();
      var mime = file.getMimeType();
      if (mime.indexOf("image/") !== -1) {
        var fileId = file.getId();
        images.push({
          id: fileId,
          name: file.getName(),
          url: "https://lh3.googleusercontent.com/d/" + fileId,
          thumbnailUrl: "https://lh3.googleusercontent.com/d/" + fileId
        });
        count++;
      }
    }

    return { success: true, images: images };
  } catch (err) {
    return { success: false, error: err.toString(), images: [] };
  }
}

// ==========================================
// TEST EMAIL
// ==========================================
function sendTestEmail(campaignData, testRecipient) {
  try {
    if (!testRecipient || testRecipient.indexOf("@") === -1) {
      throw new Error("Zadejte platnou testovací emailovou adresu.");
    }

    var webAppUrl = getWebAppUrl();
    var htmlBody = EmailTemplate.buildHtml(campaignData, testRecipient, "Test", webAppUrl);
    var senderName = campaignData.senderName || CONFIG.DEFAULT_SENDER_NAME;

    GmailApp.sendEmail(testRecipient, "[TEST] " + (campaignData.subject || "Newsletter"), "", {
      htmlBody: htmlBody,
      name: senderName
    });

    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ==========================================
// CORE SENDING ENGINE (WITH 4.5m WATCHDOG & QUOTA CHECK)
// ==========================================
function processGroupWithWatchdog(groupNumber, isSimulation) {
  var ss = getSpreadsheet();
  var sheetName = "Group " + groupNumber;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return { allFinished: true, sentInThisRun: 0 };

  var campSheet = ss.getSheetByName("_Active_Campaign");
  var campRow = campSheet.getRange(2, 1, 1, 7).getValues()[0];
  var subject = campRow[0];
  var campaignDataJson = campRow[6];
  if (!campaignDataJson) return { allFinished: true, sentInThisRun: 0 };

  var campaignData = JSON.parse(campaignDataJson);
  var unSet = getUnsubscribedSet();
  var webAppUrl = getWebAppUrl();
  var senderName = campaignData.senderName || CONFIG.DEFAULT_SENDER_NAME;

  var startTime = Date.now();
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var sentInThisRun = 0;
  var allFinished = true;

  for (var i = 0; i < rows.length; i++) {
    var email = rows[i][0] ? rows[i][0].toString().trim() : "";
    var name = rows[i][1] ? rows[i][1].toString().trim() : "";
    var status = rows[i][2] ? rows[i][2].toString().trim().toUpperCase() : "";

    if (!email || email.indexOf("@") === -1) continue;

    if (status.indexOf("ODESLÁNO") !== -1 || status === "ODHLÁŠENO" || status === "UNSUBSCRIBED" || unSet[email.toLowerCase()]) {
      continue;
    }

    // 1. Quota Check
    var quotaRemaining = MailApp.getRemainingDailyQuota();
    if (quotaRemaining <= CONFIG.SAFETY_QUOTA_THRESHOLD) {
      Logger.log("Daily quota exhausted (" + quotaRemaining + " left). Scheduling 1-hour check...");
      allFinished = false;
      scheduleTrigger("executeGroup" + groupNumber + "WatchdogTrigger", 60 * 60 * 1000);
      return { allFinished: false, pausedForQuota: true, sentInThisRun: sentInThisRun };
    }

    // 2. Runtime Watchdog (Stop at 4.5 minutes)
    if (Date.now() - startTime > CONFIG.MAX_RUNTIME_MS) {
      Logger.log("Approaching 4.5m execution limit. Chaining continuation trigger in 2 minutes...");
      allFinished = false;
      scheduleTrigger("executeGroup" + groupNumber + "WatchdogTrigger", 2 * 60 * 1000);
      return { allFinished: false, pausedForTimeout: true, sentInThisRun: sentInThisRun };
    }

    // 3. Send Email
    var html = EmailTemplate.buildHtml(campaignData, email, name, webAppUrl);
    GmailApp.sendEmail(email, subject, "", {
      htmlBody: html,
      name: senderName
    });

    var timeStamp = Utilities.formatDate(new Date(), "Europe/Prague", "HH:mm");
    sheet.getRange(i + 2, 3).setValue("ODESLÁNO (" + timeStamp + ")").setFontColor("#2b6cb0").setFontWeight("bold");
    sentInThisRun++;
  }

  return { allFinished: allFinished, sentInThisRun: sentInThisRun };
}

// ==========================================
// LAUNCH CAMPAIGN
// ==========================================
function launchCampaign(campaignData, isSimulation) {
  try {
    var ss = getSpreadsheet();
    initializeSheets(ss);
    clearAllCampaignTriggers();

    resetSentStatusForNewCampaign(ss.getSheetByName("Group 1"));
    resetSentStatusForNewCampaign(ss.getSheetByName("Group 2"));

    var subject = campaignData.subject || "Newsletter";
    var delayMs = isSimulation ? CONFIG.SIMULATION_DELAY_MS : CONFIG.GROUP2_DELAY_MS;
    var scheduledDate = new Date(Date.now() + delayMs);

    var campSheet = ss.getSheetByName("_Active_Campaign");
    campSheet.getRange("A2:G2").clearContent();
    campSheet.getRange(2, 1, 1, 7).setValues([[
      subject,
      "GROUP_1_PROCESSING",
      new Date(),
      scheduledDate,
      isSimulation ? true : false,
      "",
      JSON.stringify(campaignData)
    ]]);

    var result = processGroupWithWatchdog(1, isSimulation);

    if (result.allFinished) {
      campSheet.getRange(2, 2).setValue("GROUP_1_SENT");
      var trigger = scheduleTrigger("executeGroup2Trigger", delayMs);
      campSheet.getRange(2, 6).setValue(trigger.getUniqueId());
    } else {
      campSheet.getRange(2, 2).setValue("GROUP_1_PAUSED");
    }

    return {
      success: true,
      sentGroup1Count: result.sentInThisRun,
      allGroup1Finished: result.allFinished,
      isSimulation: isSimulation,
      scheduledTimeFormatted: Utilities.formatDate(scheduledDate, "Europe/Prague", "dd.MM.yyyy HH:mm:ss")
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function resetSentStatusForNewCampaign(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return;
  var rows = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < rows.length; i++) {
    var val = rows[i][0] ? rows[i][0].toString().trim().toUpperCase() : "";
    if (val.indexOf("ODESLÁNO") !== -1) {
      sheet.getRange(i + 2, 3).clearContent();
    }
  }
}

// ==========================================
// TRIGGER HANDLERS
// ==========================================
function executeGroup1WatchdogTrigger() {
  clearSpecificTriggers("executeGroup1WatchdogTrigger");
  var ss = getSpreadsheet();
  var campSheet = ss.getSheetByName("_Active_Campaign");
  var isSimulation = campSheet.getRange(2, 5).getValue() === true;

  var result = processGroupWithWatchdog(1, isSimulation);
  if (result.allFinished) {
    campSheet.getRange(2, 2).setValue("GROUP_1_SENT");
    var delayMs = isSimulation ? CONFIG.SIMULATION_DELAY_MS : CONFIG.GROUP2_DELAY_MS;
    var scheduledDate = new Date(Date.now() + delayMs);
    campSheet.getRange(2, 4).setValue(scheduledDate);
    var trigger = scheduleTrigger("executeGroup2Trigger", delayMs);
    campSheet.getRange(2, 6).setValue(trigger.getUniqueId());
  }
}

function executeGroup2Trigger() {
  clearSpecificTriggers("executeGroup2Trigger");
  var ss = getSpreadsheet();
  var campSheet = ss.getSheetByName("_Active_Campaign");
  var isSimulation = campSheet.getRange(2, 5).getValue() === true;
  campSheet.getRange(2, 2).setValue("GROUP_2_PROCESSING");

  var result = processGroupWithWatchdog(2, isSimulation);
  if (result.allFinished) {
    finishCampaign(ss, campSheet);
  }
}

function executeGroup2WatchdogTrigger() {
  clearSpecificTriggers("executeGroup2WatchdogTrigger");
  var ss = getSpreadsheet();
  var campSheet = ss.getSheetByName("_Active_Campaign");
  var isSimulation = campSheet.getRange(2, 5).getValue() === true;

  var result = processGroupWithWatchdog(2, isSimulation);
  if (result.allFinished) {
    finishCampaign(ss, campSheet);
  }
}

function finishCampaign(ss, campSheet) {
  campSheet.getRange(2, 2).setValue("COMPLETED");
  campSheet.getRange(2, 4).setValue(new Date());

  var subject = campSheet.getRange(2, 1).getValue();
  var g1Sent = countStatus(ss.getSheetByName("Group 1"), "ODESLÁNO");
  var g2Sent = countStatus(ss.getSheetByName("Group 2"), "ODESLÁNO");
  var adminEmail = CONFIG.ADMIN_EMAIL || Session.getEffectiveUser().getEmail();

  GmailApp.sendEmail(
    adminEmail,
    "🎉 Newsletter kampaň dokončena: " + subject,
    "Dobrý den,\n\nkampaň pro obě skupiny byla úspěšně dokončena.\n\n" +
    "- Odesláno Skupině 1: " + g1Sent + " emailů\n" +
    "- Odesláno Skupině 2: " + g2Sent + " emailů\n\n" +
    "Všechny emaily byly spolehlivě doručeny bez překročení časových limitů.\n\nNewsletter Studio."
  );
}

function scheduleTrigger(functionName, delayMs) {
  clearSpecificTriggers(functionName);
  return ScriptApp.newTrigger(functionName)
    .timeBased()
    .after(delayMs)
    .create();
}

function clearSpecificTriggers(functionName) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function clearAllCampaignTriggers() {
  clearSpecificTriggers("executeGroup1WatchdogTrigger");
  clearSpecificTriggers("executeGroup2Trigger");
  clearSpecificTriggers("executeGroup2WatchdogTrigger");
}

function runGroup2Now() {
  clearAllCampaignTriggers();
  executeGroup2Trigger();
  return { success: true };
}

function cancelScheduledTrigger() {
  clearAllCampaignTriggers();
  var ss = getSpreadsheet();
  var campSheet = ss.getSheetByName("_Active_Campaign");
  if (campSheet && campSheet.getLastRow() >= 2) {
    campSheet.getRange(2, 2).setValue("CANCELLED");
  }
  return { success: true };
}