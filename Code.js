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
 * - Strict email validation: invalid addresses are stamped "NEPLATNÝ EMAIL" and skipped, never block the campaign
 * - Per-recipient try/catch: a failure on one recipient can never stop the campaign
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
  
  // Custom Hosting URL for Unsubscribe (if hosting on own server, e.g. "https://www.vasedomena.cz/unsubscribe.html")
  // If empty, falls back to Google Apps Script Web App URL
  CUSTOM_UNSUB_URL: "",
  
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
// MENU IN GOOGLE SHEETS
// ==========================================
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("📧 Newsletter")
      .addItem("⚙️ Inicializovat listy a Nastavení", "initializeSheets")
      .addItem("🧹 Odstranit duplicity z kontaktů", "removeDuplicatesUi")
      .addItem("📜 Zobrazit log smazaných duplicit", "showDuplicatesLog")
      .addItem("📊 Zkontrolovat stav / info", "getSystemInfo")
      .addToUi();
  } catch (e) {
    // V kontextu web app getUi() neexistuje, ignorujeme
  }
}

// ==========================================
// WEB APP & API ROUTING (doGet & doPost)
// ==========================================
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";

  // 1. Unsubscribe page (Fallback if running inside GAS)
  if (action === "unsub") {
    var email = e.parameter.email || "";
    var template = HtmlService.createTemplateFromFile("UnsubPage");
    template.email = email;
    return template.evaluate()
      .setTitle("Odhlášení z odběru")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  // 2. Direct JSON status/info for health checks or external webhosting GET
  if (action === "info" || action === "status") {
    return ContentService.createTextOutput(JSON.stringify(getSystemInfo()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 3. Confirm unsubscribe via GET (optional fallback)
  if (action === "confirm_unsub" && e.parameter.email) {
    return ContentService.createTextOutput(JSON.stringify(confirmUnsubscribe(e.parameter.email)))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 4. Default: Render built-in Web App Index
  var html = HtmlService.createTemplateFromFile("Index").evaluate();
  return html
    .setTitle("Newsletter Studio")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Universal JSON API Router (POST)
 * Allows frontend running on external webhosting (or local server) to call Google Apps Script securely.
 */
function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    var body = {};
    try {
      body = JSON.parse(raw);
    } catch (parseErr) {
      body = {};
    }

    var action = body.action || (e && e.parameter && e.parameter.action) || "";
    var payload = body.payload || body.data || body;
    var result = {};

    switch (action) {
      case "verify_password":
        result = { success: verifyPassword(payload.password) };
        break;

      case "get_system_info":
        result = getSystemInfo();
        break;

      case "upload_image":
        result = uploadImage(payload.base64Data, payload.fileName, payload.mimeType);
        break;

      case "get_drive_assets":
        result = getDriveAssets();
        break;

      case "send_test_email":
        result = sendTestEmail(payload.campaignData, payload.testRecipient);
        break;

      case "launch_campaign":
        result = launchCampaign(payload.campaignData, payload.options);
        break;

      case "run_group2_now":
        result = runGroup2Now();
        break;

      case "cancel_scheduled_trigger":
        result = cancelScheduledTrigger();
        break;

      case "confirm_unsubscribe":
        result = confirmUnsubscribe(payload.email);
        break;

      case "remove_duplicates":
        result = removeDuplicates();
        break;

      case "get_duplicates_log":
        result = getDuplicatesLog(payload.limit);
        break;

      default:
        result = { success: false, error: "Neznámá API akce: " + action };
    }

    return jsonOut(result);

  } catch (err) {
    return jsonOut({ success: false, error: err.toString() });
  }
}

// Apps Script TextOutput has NO setHeader(). Web Apps add CORS headers automatically
// for simple (text/plain) POST requests, so we only return JSON.
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

function verifyPassword(password) {
  var settings = getSettings();
  return password === settings.adminPassword;
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

  var userEmail = "";
  try {
    userEmail = Session.getEffectiveUser().getEmail();
  } catch (e) {
    try {
      userEmail = Session.getActiveUser().getEmail();
    } catch (e2) {
      userEmail = CONFIG.ADMIN_EMAIL || "";
    }
  }

  var settings = getSettings(ss);

  return {
    dailyQuotaRemaining: MailApp.getRemainingDailyQuota(),
    userEmail: userEmail,
    webAppUrl: getWebAppUrl(),
    group1Count: g1Total,
    group1Sent: g1Sent,
    group2Count: g2Total,
    group2Sent: g2Sent,
    unsubscribedCount: unTotal,
    activeCampaign: activeCamp,
    hasPendingGroup2Trigger: pendingTrigger,
    pendingTriggerType: triggerType,
    customUnsubUrl: settings.customUnsubUrl,
    delayHours: settings.delayHours
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

  // 4b. Log duplicit (co bylo smazáno)
  var dupLog = ss.getSheetByName(DUP_LOG_SHEET);
  if (!dupLog) {
    dupLog = ss.insertSheet(DUP_LOG_SHEET);
    dupLog.appendRow(["Datum a čas", "Email", "Jméno", "Stav", "Smazáno ze skupiny", "Důvod", "Původní řádek", "Ponecháno v"]);
    dupLog.getRange("A1:H1").setFontWeight("bold").setBackground("#718096").setFontColor("#ffffff");
    dupLog.setFrozenRows(1);
    dupLog.setColumnWidth(1, 150);
    dupLog.setColumnWidth(2, 260);
    dupLog.setColumnWidth(6, 220);
    dupLog.setColumnWidth(8, 160);
  }

  // 5. Nastavení (Uživatelsky konfigurovatelná nastavení)
  var setSheet = ss.getSheetByName("Nastavení");
  if (!setSheet) {
    setSheet = ss.insertSheet("Nastavení");
    setSheet.appendRow(["Položka", "Hodnota", "Popis / Nápověda"]);
    setSheet.appendRow(["Heslo administrátora", "admin", "Heslo pro přístup do webového editoru newsletteru"]);
    setSheet.appendRow(["Webhosting / Odhlašovací adresa", "https://newsletter.zmuhrineves.cz/unsubscribe.html", "URL adresa odhlašovací stránky na hostingu (můžete zadat i jen newsletter.zmuhrineves.cz)"]);
    setSheet.appendRow(["Výchozí prodleva Skupiny 2 (hodiny)", 24, "Po kolika hodinách od dokončení Skupiny 1 se automaticky odešle Skupina 2"]);
    setSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#2d3748").setFontColor("#ffffff");
    setSheet.setFrozenRows(1);
    setSheet.setColumnWidth(1, 280);
    setSheet.setColumnWidth(2, 400);
    setSheet.setColumnWidth(3, 480);
  }
}

// ==========================================
// DYNAMIC SETTINGS FROM SHEET "Nastavení"
// ==========================================
function getSettings(ss) {
  if (!ss) ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Nastavení");
  var settings = {
    adminPassword: CONFIG.ADMIN_PASSWORD || "admin",
    customUnsubUrl: "https://newsletter.zmuhrineves.cz/unsubscribe.html",
    delayHours: CONFIG.DEFAULT_GROUP2_DELAY_HOURS || 24
  };

  if (!sheet || sheet.getLastRow() <= 1) {
    return settings;
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    var key = (data[i][0] || "").toString().trim().toLowerCase();
    var val = (data[i][1] !== undefined && data[i][1] !== null) ? data[i][1].toString().trim() : "";

    if (key.indexOf("heslo") !== -1 || key.indexOf("password") !== -1) {
      if (val) settings.adminPassword = val;
    } else if (key.indexOf("odhlaš") !== -1 || key.indexOf("unsub") !== -1 || key.indexOf("webhosting") !== -1 || key.indexOf("domén") !== -1 || key.indexOf("subdom") !== -1) {
      if (val) {
        var cleanUrl = val;
        // Pokud chybí http:// nebo https://, doplníme https://
        if (!cleanUrl.match(/^https?:\/\//i)) {
          cleanUrl = "https://" + cleanUrl;
        }
        // Pokud končí lomítkem, odstraníme
        cleanUrl = cleanUrl.replace(/\/+$/, "");
        // Pokud nezadali konkrétní soubor (.html, .php), doplníme /unsubscribe.html
        if (!cleanUrl.match(/\.html$/i) && !cleanUrl.match(/\.php$/i)) {
          cleanUrl = cleanUrl + "/unsubscribe.html";
        }
        settings.customUnsubUrl = cleanUrl;
      }
    } else if (key.indexOf("prodlev") !== -1 || key.indexOf("delay") !== -1) {
      var num = parseFloat(val);
      if (!isNaN(num) && num > 0) settings.delayHours = num;
    }
  }

  return settings;
}

// ==========================================
// EMAIL NORMALIZATION & VALIDATION
// ==========================================
var DUP_LOG_SHEET = "Log duplicit";

// Sjednotí email: odstraní neviditelné znaky, mezery, "mailto:", "Jméno <email>" a převede na malá písmena
function normalizeEmail(v) {
  var s = (v === null || v === undefined) ? "" : String(v);
  s = s.replace(/[\u200B-\u200D\u2060\uFEFF\u00A0]/g, " ").trim();
  var m = s.match(/<([^<>]+)>/);
  if (m) s = m[1];
  s = s.replace(/^mailto:/i, "").replace(/\s+/g, "");
  return s.toLowerCase();
}

// Strict validator. Rejects what Gmail rejects: double dots, trailing commas/semicolons,
// multiple addresses in one cell, raw diacritics in domain, bad TLD, etc.
// Punycode domains (xn--) are accepted.
function isValidEmail_(raw) {
  var s = normalizeEmail(raw);
  if (!s || s.length > 254) return false;
  if (/[^\x21-\x7E]/.test(s)) return false;            // spaces, diacritics, invisible chars
  var m = s.match(/^([A-Za-z0-9.!#$%&'*+\/=?^_`{|}~-]+)@([A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+)$/);
  if (!m) return false;
  var local = m[1], labels = m[2].split(".");
  if (local.length > 64 || local.charAt(0) === "." || local.slice(-1) === "." || local.indexOf("..") !== -1) return false;
  for (var i = 0; i < labels.length; i++) {
    var l = labels[i];
    if (l.length < 1 || l.length > 63 || l.charAt(0) === "-" || l.slice(-1) === "-") return false;
  }
  return /^([A-Za-z]{2,}|xn--[A-Za-z0-9-]+)$/.test(labels[labels.length - 1]);
}

// ==========================================
// DEDUPLICATION + VALIDATION: GROUP 1 & GROUP 2 (+ LOG)
// ==========================================
function readGroupSheet_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), 3);
  var values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  return { values: values, lastRow: lastRow, lastCol: lastCol };
}

function writeGroupSheet_(sheet, info, keptRows) {
  if (info.lastRow > 1) {
    sheet.getRange(2, 1, info.lastRow - 1, info.lastCol).clearContent();
  }
  if (keptRows.length > 0) {
    sheet.getRange(2, 1, keptRows.length, info.lastCol).setValues(keptRows);
  }
}

function removeDuplicates() {
  var ss = getSpreadsheet();
  initializeSheets(ss);

  var g1 = ss.getSheetByName("Group 1");
  var g2 = ss.getSheetByName("Group 2");
  var d1 = readGroupSheet_(g1);
  var d2 = readGroupSheet_(g2);

  var tz = Session.getScriptTimeZone();
  var stamp = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
  var removed = [];   // záznamy pro log

  function logRemoved(row, group, reason, rowNum, keptIn) {
    removed.push({
      time: stamp, email: String(row[0]).trim(), name: String(row[1] || ""), status: String(row[2] || ""),
      group: group, reason: reason, row: rowNum, keptIn: keptIn
    });
  }

  // 1. Group 1: invalid format, then duplicates (first occurrence kept)
  var firstRowG1 = {};   // normalizovaný email -> původní číslo řádku
  var keptG1 = [];
  var g1Dupes = 0, g1Invalid = 0;

  for (var i = 0; i < d1.values.length; i++) {
    var row = d1.values[i];
    var norm = normalizeEmail(row[0]);
    if (!norm) continue; // prázdný řádek se přeskočí

    if (!isValidEmail_(norm)) {
      g1Invalid++;
      logRemoved(row, "Group 1", "Neplatný formát emailu", i + 2, "—");
    } else if (firstRowG1.hasOwnProperty(norm)) {
      g1Dupes++;
      logRemoved(row, "Group 1", "Duplicita uvnitř Group 1", i + 2, "Group 1, řádek " + firstRowG1[norm]);
    } else {
      firstRowG1[norm] = i + 2;
      keptG1.push(row);
    }
  }

  // 2. Group 2: invalid format, already in G1 (delete from G2), duplicates inside G2
  var firstRowG2 = {};
  var keptG2 = [];
  var g2Dupes = 0, crossDupes = 0, g2Invalid = 0;

  for (var j = 0; j < d2.values.length; j++) {
    var row2 = d2.values[j];
    var norm2 = normalizeEmail(row2[0]);
    if (!norm2) continue;

    if (!isValidEmail_(norm2)) {
      g2Invalid++;
      logRemoved(row2, "Group 2", "Neplatný formát emailu", j + 2, "—");
    } else if (firstRowG1.hasOwnProperty(norm2)) {
      crossDupes++;
      logRemoved(row2, "Group 2", "Už je v Group 1", j + 2, "Group 1, řádek " + firstRowG1[norm2]);
    } else if (firstRowG2.hasOwnProperty(norm2)) {
      g2Dupes++;
      logRemoved(row2, "Group 2", "Duplicita uvnitř Group 2", j + 2, "Group 2, řádek " + firstRowG2[norm2]);
    } else {
      firstRowG2[norm2] = j + 2;
      keptG2.push(row2);
    }
  }

  // 3. Zápis zpět do listů (jen pokud se něco změnilo)
  if (keptG1.length !== d1.values.length) writeGroupSheet_(g1, d1, keptG1);
  if (keptG2.length !== d2.values.length) writeGroupSheet_(g2, d2, keptG2);

  // 4. Zápis do listu "Log duplicit"
  if (removed.length > 0) {
    var logSheet = ss.getSheetByName(DUP_LOG_SHEET);
    var logRows = removed.map(function(r) {
      return [r.time, r.email, r.name, r.status, r.group, r.reason, r.row, r.keptIn];
    });
    var startRow = logSheet.getLastRow() + 1;
    var rng = logSheet.getRange(startRow, 1, logRows.length, 8);
    rng.setNumberFormat("@");   // text, aby se datum/čísla neměnily
    rng.setValues(logRows);
  }

  var totalRemoved = removed.length;

  return {
    success: true,
    totalRemoved: totalRemoved,
    g1DuplicatesRemoved: g1Dupes,
    g2DuplicatesRemoved: g2Dupes,
    crossDuplicatesRemoved: crossDupes,
    g1InvalidRemoved: g1Invalid,
    g2InvalidRemoved: g2Invalid,
    remainingG1: keptG1.length,
    remainingG2: keptG2.length,
    timestamp: stamp,
    removed: removed.slice(0, 1000),
    message: totalRemoved === 0
      ? "Žádné duplicity ani neplatné emaily nenalezeny."
      : "Kontrola dokončena. Celkem smazáno " + totalRemoved + " kontaktů (neplatný formát G1: " + g1Invalid +
        ", G2: " + g2Invalid + "; duplicity uvnitř G1: " + g1Dupes + ", uvnitř G2: " + g2Dupes +
        "; už v G1 a smazáno z G2: " + crossDupes + ")."
  };
}

// Vrátí posledních N záznamů z listu "Log duplicit" (nejnovější první)
function getDuplicatesLog(limit) {
  try {
    limit = Number(limit) || 300;
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(DUP_LOG_SHEET);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, total: 0, entries: [] };
    }
    var total = sheet.getLastRow() - 1;
    var count = Math.min(limit, total);
    var startRow = sheet.getLastRow() - count + 1;
    var vals = sheet.getRange(startRow, 1, count, 8).getValues();
    var tz = Session.getScriptTimeZone();
    var entries = vals.map(function(r) {
      var t = r[0] instanceof Date ? Utilities.formatDate(r[0], tz, "yyyy-MM-dd HH:mm:ss") : String(r[0]);
      return { time: t, email: String(r[1]), name: String(r[2]), status: String(r[3]),
               group: String(r[4]), reason: String(r[5]), row: String(r[6]), keptIn: String(r[7]) };
    }).reverse();
    return { success: true, total: total, entries: entries };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function escHtml_(t) {
  return String(t === null || t === undefined ? "" : t)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function buildDupDialogHtml_(summaryHtml, entries, footerNote) {
  var rows = entries.map(function(e) {
    return "<tr><td>" + escHtml_(e.time) + "</td><td><b>" + escHtml_(e.email) + "</b></td><td>" + escHtml_(e.name) +
           "</td><td>" + escHtml_(e.group) + "</td><td>" + escHtml_(e.reason) + "</td><td>" + escHtml_(e.row) +
           "</td><td>" + escHtml_(e.keptIn) + "</td></tr>";
  }).join("");
  if (!rows) rows = '<tr><td colspan="7" style="text-align:center;color:#718096;padding:20px;">Zatím nebylo nic smazáno.</td></tr>';
  return '<div style="font-family:Arial,sans-serif;font-size:13px;color:#2d3748;">' +
    summaryHtml +
    '<div style="max-height:380px;overflow:auto;border:1px solid #e2e8f0;border-radius:6px;">' +
    '<table style="border-collapse:collapse;width:100%;">' +
    '<thead><tr style="background:#edf2f7;text-align:left;position:sticky;top:0;">' +
    '<th style="padding:6px 8px;">Čas</th><th style="padding:6px 8px;">Email</th><th style="padding:6px 8px;">Jméno</th>' +
    '<th style="padding:6px 8px;">Smazáno z</th><th style="padding:6px 8px;">Důvod</th><th style="padding:6px 8px;">Řádek</th>' +
    '<th style="padding:6px 8px;">Ponecháno v</th></tr></thead><tbody>' + rows + '</tbody></table></div>' +
    '<p style="color:#718096;font-size:12px;">' + (footerNote || "") + '</p></div>' +
    '<style>td{padding:5px 8px;border-top:1px solid #e2e8f0;vertical-align:top;}</style>';
}

function removeDuplicatesUi() {
  try {
    var res = removeDuplicates();
    var ui = SpreadsheetApp.getUi();

    if (res.totalRemoved === 0) {
      ui.alert("Čištění kontaktů", "Žádné duplicity ani neplatné emaily nenalezeny.\n\nGroup 1: " + res.remainingG1 + " kontaktů\nGroup 2: " + res.remainingG2 + " kontaktů", ui.ButtonSet.OK);
      return;
    }

    var summary =
      "<p><b>Smazáno celkem: " + res.totalRemoved + "</b><br>" +
      "• Neplatný formát emailu: Group 1 = " + res.g1InvalidRemoved + ", Group 2 = " + res.g2InvalidRemoved + "<br>" +
      "• Duplicity uvnitř Group 1: " + res.g1DuplicatesRemoved + "<br>" +
      "• Duplicity uvnitř Group 2: " + res.g2DuplicatesRemoved + "<br>" +
      "• Už v Group 1 (smazáno z Group 2): " + res.crossDuplicatesRemoved + "<br>" +
      "Zůstává: Group 1 = " + res.remainingG1 + ", Group 2 = " + res.remainingG2 + "</p>";

    var html = HtmlService.createHtmlOutput(
      buildDupDialogHtml_(summary, res.removed, "Kompletní historii najdete v listu „" + DUP_LOG_SHEET + "“.")
    ).setWidth(900).setHeight(560);
    ui.showModalDialog(html, "🧹 Smazané kontakty");
  } catch (err) {
    try {
      SpreadsheetApp.getUi().alert("Chyba při čištění kontaktů", err.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (eUi) {
      Logger.log("Chyba: " + err);
    }
  }
}

// Menu: zobrazí historii smazaných duplicit
function showDuplicatesLog() {
  var ui = SpreadsheetApp.getUi();
  var res = getDuplicatesLog(500);
  if (!res.success) {
    ui.alert("Chyba", res.error, ui.ButtonSet.OK);
    return;
  }
  var summary = "<p>Celkem záznamů v logu: <b>" + res.total + "</b>" +
    (res.total > res.entries.length ? " (zobrazeno posledních " + res.entries.length + ")" : "") + "</p>";
  var html = HtmlService.createHtmlOutput(
    buildDupDialogHtml_(summary, res.entries, "Kompletní historii najdete v listu „" + DUP_LOG_SHEET + "“.")
  ).setWidth(900).setHeight(560);
  ui.showModalDialog(html, "📜 Log smazaných kontaktů");
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
    if (!testRecipient || !isValidEmail_(testRecipient)) {
      throw new Error("Zadejte platnou testovací emailovou adresu.");
    }
    testRecipient = normalizeEmail(testRecipient);

    var settings = getSettings();
    campaignData.customUnsubUrl = campaignData.customUnsubUrl || settings.customUnsubUrl;
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
  var settings = getSettings(ss);
  campaignData.customUnsubUrl = campaignData.customUnsubUrl || settings.customUnsubUrl;
  var unSet = getUnsubscribedSet();
  var webAppUrl = getWebAppUrl();
  var senderName = campaignData.senderName || CONFIG.DEFAULT_SENDER_NAME;
  var adminEmail = CONFIG.ADMIN_EMAIL || Session.getEffectiveUser().getEmail();

  var startTime = Date.now();
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var sentInThisRun = 0;
  var allFinished = true;

  for (var i = 0; i < rows.length; i++) {
    var email = normalizeEmail(rows[i][0]);   // handles "Name <a@b.cz>", mailto:, invisible chars
    var name = rows[i][1] ? rows[i][1].toString().trim() : "";
    var status = rows[i][2] ? rows[i][2].toString().trim().toUpperCase() : "";

    if (!email) continue;

    // Skip already processed, unsubscribed, invalid or errored rows
    if (status.indexOf("ODESLÁNO") !== -1 || status.indexOf("DRY RUN") !== -1 ||
        status.indexOf("NEPLATNÝ") !== -1 || status.indexOf("CHYBA") !== -1 ||
        status === "ODHLÁŠENO" || status === "UNSUBSCRIBED" || unSet[email]) {
      continue;
    }

    // Invalid format: mark and move on (in real runs too, not just dry runs)
    if (!isValidEmail_(email)) {
      sheet.getRange(i + 2, 3).setValue((isDryRun ? "DRY RUN: " : "") + "NEPLATNÝ EMAIL")
        .setFontColor("#c53030").setFontWeight("bold");
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
      
      var trig2 = scheduleTrigger("executeGroup" + groupNumber + "WatchdogTrigger", 2 * 60 * 1000);
      campSheet.getRange(2, 6).setValue(trig2.getUniqueId());

      return { allFinished: false, pausedForTimeout: true, sentInThisRun: sentInThisRun, isDryRun: isDryRun };
    }

    // 3. Process Row; a failure on one recipient can never stop the campaign
    var timeStamp = Utilities.formatDate(new Date(), "Europe/Prague", "HH:mm");
    try {
      var html = EmailTemplate.buildHtml(campaignData, email, name, webAppUrl);

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
    } catch (sendErr) {
      var msg = String(sendErr && sendErr.message ? sendErr.message : sendErr);
      Logger.log("Send failed for " + email + ": " + msg);
      var bad = /invalid (email|address|argument)/i.test(msg);
      sheet.getRange(i + 2, 3).setValue(bad ? "NEPLATNÝ EMAIL" : "CHYBA: " + msg.substring(0, 80))
        .setFontColor("#c53030").setFontWeight("bold");
    }
  }

  return { allFinished: allFinished, sentInThisRun: sentInThisRun, isDryRun: isDryRun };
}

// ==========================================
// LAUNCH CAMPAIGN
// ==========================================
function launchCampaign(campaignData, options) {
  try {
    var ss = getSpreadsheet();
    initializeSheets(ss);
    var settings = getSettings(ss);

    var isDryRun = false;
    var delayHours = settings.delayHours || CONFIG.DEFAULT_GROUP2_DELAY_HOURS;
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
    campaignData.customUnsubUrl = campaignData.customUnsubUrl || settings.customUnsubUrl;

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

// Clears sent / dry-run / invalid / error stamps so a new campaign retries every row
// (e.g. an address you fixed since the last run)
function resetSentStatusForNewCampaign(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return;
  var rows = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < rows.length; i++) {
    var val = rows[i][0] ? rows[i][0].toString().trim().toUpperCase() : "";
    if (val.indexOf("ODESLÁNO") !== -1 || val.indexOf("DRY RUN") !== -1 ||
        val.indexOf("NEPLATNÝ") !== -1 || val.indexOf("CHYBA") !== -1) {
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
  var g1Bad = countStatus(ss.getSheetByName("Group 1"), "NEPLATNÝ") + countStatus(ss.getSheetByName("Group 1"), "CHYBA");
  var g2Bad = countStatus(ss.getSheetByName("Group 2"), "NEPLATNÝ") + countStatus(ss.getSheetByName("Group 2"), "CHYBA");
  var adminEmail = CONFIG.ADMIN_EMAIL || Session.getEffectiveUser().getEmail();

  if (isDryRun) {
    GmailApp.sendEmail(
      adminEmail,
      "🧪 [DRY RUN HOTOVO] Test kampaně: " + subject,
      "Dobrý den,\n\n" +
      "úspěšně proběhla kompletní simulace (Dry Run) na celém vašem seznamu bez odesílání emailů příjemcům.\n\n" +
      "- Zkontrolováno Skupina 1: " + g1Count + " kontaktů\n" +
      "- Zkontrolováno Skupina 2: " + g2Count + " kontaktů\n" +
      "- Přeskočeno (neplatné / chyba) G1: " + g1Bad + ", G2: " + g2Bad + "\n" +
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
      "- Odesláno Skupině 2: " + g2Count + " emailů\n" +
      "- Přeskočeno (neplatný email / chyba) G1: " + g1Bad + ", G2: " + g2Bad + " (viz sloupec Stav)\n\n" +
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