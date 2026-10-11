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
  
  // Default Group 2 delay: 24 hours + 30 minutes buffer (in hours: 24.5)
  DEFAULT_GROUP2_DELAY_HOURS: 24.5,
  
  // Pause duration when quota is exhausted before re-checking (1 hour)
  QUOTA_PAUSE_CHECK_DELAY_MS: 60 * 60 * 1000,
  
  // Dry run delay between Group 1 and Group 2: 2 minutes
  DRY_RUN_DELAY_MS: 2 * 60 * 1000
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

  var g1Sent = countStatus(g1Sheet, "ODESLÁNO") + countStatus(g1Sheet, "DRY RUN: OK");
  var g2Sent = countStatus(g2Sheet, "ODESLÁNO") + countStatus(g2Sheet, "DRY RUN: OK");

  var activeCamp = null;
  if (campSheet.getLastRow() >= 2) {
    var data = campSheet.getRange(2, 1, 1, 8).getValues()[0];
    if (data[0]) {
      activeCamp = {
        subject: data[0],
        status: data[1],
        createdTime: data[2],
        scheduledTime: data[3],
        isDryRun: data[4] === true || data[4] === "true",
        triggerId: data[5],
        delayHours: data[7] || CONFIG.DEFAULT_GROUP2_DELAY_HOURS
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
    camp.appendRow(["Subject", "Status", "Created Time", "Scheduled Time", "Is Dry Run", "Trigger ID", "Data JSON", "Delay Hours"]);
    camp.getRange("A1:H1").setFontWeight("bold").setBackground("#718096").setFontColor("#ffffff");
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
// CORE SENDING ENGINE (WITH 4.5m WATCHDOG & QUOTA CHECK & DRY RUN)
// ==========================================
function processGroupWithWatchdog(groupNumber, isDryRun, delayHours) {
  var ss = getSpreadsheet();
  var sheetName = "Group " + groupNumber;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return { allFinished: true, sentInThisRun: 0, isDryRun: isDryRun };

  var campSheet = ss.getSheetByName("_Active_Campaign");
  var campRow = campSheet.getRange(2, 1, 1, 8).getValues()[0];
  var subject = campRow[0];
  var campaignDataJson = campRow[6];
  if (!campaignDataJson) return { allFinished: true, sentInThisRun: 0, isDryRun: isDryRun };

  var campaignData = JSON.parse(campaignDataJson);
  var unSet = getUnsubscribedSet();
  var webAppUrl = getWebAppUrl();
  var senderName = campaignData.senderName || CONFIG.DEFAULT_SENDER_NAME;
  var adminEmail = CONFIG.ADMIN_EMAIL || Session.getEffectiveUser().getEmail();

  var startTime = Date.now();
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var sentInThisRun = 0;
  var allFinished = true;
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  for (var i = 0; i < rows.length; i++) {
    var email = rows[i][0] ? rows[i][0].toString().trim() : "";
    var name = rows[i][1] ? rows[i][1].toString().trim() : "";
    var status = rows[i][2] ? rows[i][2].toString().trim().toUpperCase() : "";

    if (!email) continue;

    // Check if already processed or unsubscribed
    if (status.indexOf("ODESLÁNO") !== -1 || status.indexOf("DRY RUN: OK") !== -1 || 
        status === "ODHLÁŠENO" || status === "UNSUBSCRIBED" || unSet[email.toLowerCase()]) {
      continue;
    }

    // Check email syntax validity
    if (!emailRegex.test(email)) {
      if (isDryRun) {
        sheet.getRange(i + 2, 3).setValue("DRY RUN: NEPLATNÝ EMAIL").setFontColor("#c53030").setFontWeight("bold");
      }
      continue;
    }

    // 1. Quota Check (Active sniffing for real sends)
    if (!isDryRun) {
      var quotaRemaining = MailApp.getRemainingDailyQuota();
      if (quotaRemaining <= CONFIG.SAFETY_QUOTA_THRESHOLD) {
        Logger.log("Daily quota exhausted (" + quotaRemaining + " left). Scheduling 1-hour check...");
        allFinished = false;
        
        var pauseUntil = new Date(Date.now() + CONFIG.QUOTA_PAUSE_CHECK_DELAY_MS);
        campSheet.getRange(2, 2).setValue("GROUP_" + groupNumber + "_PAUSED_QUOTA");
        campSheet.getRange(2, 4).setValue(pauseUntil);
        
        var trig = scheduleTrigger("executeGroup" + groupNumber + "WatchdogTrigger", CONFIG.QUOTA_PAUSE_CHECK_DELAY_MS);
        campSheet.getRange(2, 6).setValue(trig.getUniqueId());

        try {
          GmailApp.sendEmail(
            adminEmail,
            "⚠️ Kampaň pozastavena (Kvóta Google): " + subject,
            "Dobrý den,\n\ndenní limit odesílání Google Workspace byl téměř vyčerpán (zbývá: " + quotaRemaining + " emailů).\n\n" +
            "- Skupina: " + groupNumber + "\n" +
            "- Dosud odesláno v této dávce: " + sentInThisRun + "\n" +
            "- Automatické obnovení: Skript se automaticky probudí za 1 hodinu (" + Utilities.formatDate(pauseUntil, "Europe/Prague", "HH:mm") + ") a zkontroluje uvolnění kvóty.\n" +
            "- Žádné emaily se nepošlou duplicitně – skript bude pokračovat přesně tam, kde přestal.\n\nNewsletter Studio."
          );
        } catch (e) {
          Logger.log("Admin alert email failed: " + e);
        }

        return { allFinished: false, pausedForQuota: true, sentInThisRun: sentInThisRun, isDryRun: false };
      }
    }

    // 2. Runtime Watchdog (Stop at 4.5 minutes before Google's 6-minute wall)
    if (Date.now() - startTime > CONFIG.MAX_RUNTIME_MS) {
      Logger.log("Approaching 4.5m limit. Chaining continuation trigger in 2 minutes...");
      allFinished = false;
      
      var resumeTime = new Date(Date.now() + 2 * 60 * 1000);
      campSheet.getRange(2, 2).setValue("GROUP_" + groupNumber + "_PAUSED_TIMEOUT");
      campSheet.getRange(2, 4).setValue(resumeTime);
      
      var trig = scheduleTrigger("executeGroup" + groupNumber + "WatchdogTrigger", 2 * 60 * 1000);
      campSheet.getRange(2, 6).setValue(trig.getUniqueId());

      return { allFinished: false, pausedForTimeout: true, sentInThisRun: sentInThisRun, isDryRun: isDryRun };
    }

    // 3. Process Row (Dry Run vs Real Send)
    var html = EmailTemplate.buildHtml(campaignData, email, name, webAppUrl);
    var timeStamp = Utilities.formatDate(new Date(), "Europe/Prague", "HH:mm");

    if (isDryRun) {
      // DRY RUN: DO NOT SEND EMAIL TO RECIPIENT!
      // Send sample to admin on very first processed row if requested
      if (sentInThisRun === 0 && campaignData.sendSamplePreview && groupNumber === 1) {
        try {
          GmailApp.sendEmail(adminEmail, "[DRY RUN VZOREK] " + subject, "", {
            htmlBody: html,
            name: "[DRY RUN] " + senderName
          });
        } catch (e) {
          Logger.log("Dry run sample send error: " + e);
        }
      }
      sheet.getRange(i + 2, 3).setValue("DRY RUN: OK (" + timeStamp + ")").setFontColor("#6b46c1").setFontWeight("bold");
    } else {
      // REAL SEND: Send actual email via GmailApp
      GmailApp.sendEmail(email, subject, "", {
        htmlBody: html,
        name: senderName
      });
      sheet.getRange(i + 2, 3).setValue("ODESLÁNO (" + timeStamp + ")").setFontColor("#2b6cb0").setFontWeight("bold");
    }
    sentInThisRun++;
  }

  return { allFinished: allFinished, sentInThisRun: sentInThisRun, isDryRun: isDryRun };
}

// ==========================================
// LAUNCH CAMPAIGN
// ==========================================
function launchCampaign(campaignData, options) {
  try {
    var isDryRun = false;
    var delayHours = CONFIG.DEFAULT_GROUP2_DELAY_HOURS;
    var sendSamplePreview = true;

    if (typeof options === "boolean") {
      isDryRun = options; // backwards compatibility
    } else if (typeof options === "object" && options !== null) {
      isDryRun = options.isDryRun === true;
      if (options.delayHours && !isNaN(Number(options.delayHours))) {
        delayHours = Number(options.delayHours);
      }
      sendSamplePreview = options.sendSamplePreview !== false;
    }

    campaignData.sendSamplePreview = sendSamplePreview;

    var ss = getSpreadsheet();
    initializeSheets(ss);
    clearAllCampaignTriggers();

    resetSentStatusForNewCampaign(ss.getSheetByName("Group 1"));
    resetSentStatusForNewCampaign(ss.getSheetByName("Group 2"));

    var subject = campaignData.subject || "Newsletter";
    var delayMs = isDryRun ? CONFIG.DRY_RUN_DELAY_MS : Math.round(delayHours * 3600 * 1000);
    var scheduledDate = new Date(Date.now() + delayMs);

    var campSheet = ss.getSheetByName("_Active_Campaign");
    campSheet.getRange("A2:H2").clearContent();
    campSheet.getRange(2, 1, 1, 8).setValues([[
      subject,
      isDryRun ? "GROUP_1_DRY_RUN_PROCESSING" : "GROUP_1_PROCESSING",
      new Date(),
      scheduledDate,
      isDryRun,
      "",
      JSON.stringify(campaignData),
      delayHours
    ]]);

    var result = processGroupWithWatchdog(1, isDryRun, delayHours);

    if (result.allFinished) {
      campSheet.getRange(2, 2).setValue(isDryRun ? "GROUP_1_DRY_RUN_DONE" : "GROUP_1_SENT");
      var trigger = scheduleTrigger("executeGroup2Trigger", delayMs);
      campSheet.getRange(2, 6).setValue(trigger.getUniqueId());
    }

    return {
      success: true,
      sentGroup1Count: result.sentInThisRun,
      allGroup1Finished: result.allFinished,
      isDryRun: isDryRun,
      delayHours: delayHours,
      scheduledTimeFormatted: Utilities.formatDate(scheduledDate, "Europe/Prague", "dd.MM.yyyy HH:mm:ss")
    };
  } catch (err) {
    Logger.log("launchCampaign error: " + err);
    return { success: false, error: err.toString() };
  }
}

function resetSentStatusForNewCampaign(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return;
  var rows = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < rows.length; i++) {
    var val = rows[i][0] ? rows[i][0].toString().trim().toUpperCase() : "";
    if (val.indexOf("ODESLÁNO") !== -1 || val.indexOf("DRY RUN") !== -1) {
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
  var row = campSheet.getRange(2, 1, 1, 8).getValues()[0];
  var isDryRun = row[4] === true || row[4] === "true";
  var delayHours = row[7] || CONFIG.DEFAULT_GROUP2_DELAY_HOURS;

  campSheet.getRange(2, 2).setValue(isDryRun ? "GROUP_1_DRY_RUN_PROCESSING" : "GROUP_1_PROCESSING");

  var result = processGroupWithWatchdog(1, isDryRun, delayHours);
  if (result.allFinished) {
    campSheet.getRange(2, 2).setValue(isDryRun ? "GROUP_1_DRY_RUN_DONE" : "GROUP_1_SENT");
    var delayMs = isDryRun ? CONFIG.DRY_RUN_DELAY_MS : Math.round(delayHours * 3600 * 1000);
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
  var row = campSheet.getRange(2, 1, 1, 8).getValues()[0];
  var isDryRun = row[4] === true || row[4] === "true";
  var delayHours = row[7] || CONFIG.DEFAULT_GROUP2_DELAY_HOURS;

  campSheet.getRange(2, 2).setValue(isDryRun ? "GROUP_2_DRY_RUN_PROCESSING" : "GROUP_2_PROCESSING");

  var result = processGroupWithWatchdog(2, isDryRun, delayHours);
  if (result.allFinished) {
    finishCampaign(ss, campSheet, isDryRun);
  }
}

function executeGroup2WatchdogTrigger() {
  clearSpecificTriggers("executeGroup2WatchdogTrigger");
  var ss = getSpreadsheet();
  var campSheet = ss.getSheetByName("_Active_Campaign");
  var row = campSheet.getRange(2, 1, 1, 8).getValues()[0];
  var isDryRun = row[4] === true || row[4] === "true";
  var delayHours = row[7] || CONFIG.DEFAULT_GROUP2_DELAY_HOURS;

  campSheet.getRange(2, 2).setValue(isDryRun ? "GROUP_2_DRY_RUN_PROCESSING" : "GROUP_2_PROCESSING");

  var result = processGroupWithWatchdog(2, isDryRun, delayHours);
  if (result.allFinished) {
    finishCampaign(ss, campSheet, isDryRun);
  }
}

function finishCampaign(ss, campSheet, isDryRun) {
  campSheet.getRange(2, 2).setValue(isDryRun ? "DRY_RUN_COMPLETED" : "COMPLETED");
  campSheet.getRange(2, 4).setValue(new Date());

  var subject = campSheet.getRange(2, 1).getValue();
  var g1Count = countStatus(ss.getSheetByName("Group 1"), isDryRun ? "DRY RUN: OK" : "ODESLÁNO");
  var g2Count = countStatus(ss.getSheetByName("Group 2"), isDryRun ? "DRY RUN: OK" : "ODESLÁNO");
  var adminEmail = CONFIG.ADMIN_EMAIL || Session.getEffectiveUser().getEmail();

  if (isDryRun) {
    GmailApp.sendEmail(
      adminEmail,
      "🧪 [DRY RUN HOTOVO] Test kampaně: " + subject,
      "Dobrý den,\n\n" +
      "úspěšně proběhla kompletní simulace (Dry Run) na celém vašem seznamu bez odesílání emailů příjemcům.\n\n" +
      "- Zkontrolováno Skupina 1: " + g1Count + " kontaktů\n" +
      "- Zkontrolováno Skupina 2: " + g2Count + " kontaktů\n" +
      "- Všechny šablony a proměnné byly úspěšně vyrenderovány.\n" +
      "- Zpoždění a triggery fungují bezchybně.\n\n" +
      "Vše je připraveno pro ostré odeslání!\n\nNewsletter Studio."
    );
  } else {
    GmailApp.sendEmail(
      adminEmail,
      "🎉 Newsletter kampaň dokončena: " + subject,
      "Dobrý den,\n\nkampaň pro obě skupiny byla úspěšně dokončena.\n\n" +
      "- Odesláno Skupině 1: " + g1Count + " emailů\n" +
      "- Odesláno Skupině 2: " + g2Count + " emailů\n\n" +
      "Všechny emaily byly spolehlivě doručeny bez překročení časových limitů.\n\nNewsletter Studio."
    );
  }
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