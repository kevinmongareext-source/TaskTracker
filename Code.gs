/**
 * CADENCE — Google Sheets + Apps Script version
 * ------------------------------------------------
 * Stores Daily / Weekly / Monthly tasks as rows in three sheets,
 * plus a Settings sheet for your reminder email + send time.
 * Serves a web app UI (Index.html) that reads/writes the sheet live.
 *
 * SHEET COLUMNS (row 1 = header, created automatically):
 *   ID | Task | Urgency | Completed | CreatedAt | ReminderSentAt
 *
 * Urgency is one of: "Low", "Medium", "High"
 */

var CADENCES = ['Daily', 'Weekly', 'Monthly'];
var HEADERS = ['ID', 'Task', 'Urgency', 'Completed', 'CreatedAt', 'ReminderSentAt'];

// ---------- Setup ----------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cadence')
    .addItem('Open Cadence app', 'showSidebarLink_')
    .addSeparator()
    .addItem('Send reminder email now', 'sendDailyDigest')
    .addItem('Set up daily reminder trigger…', 'promptCreateTrigger_')
    .addItem('Remove reminder trigger', 'removeTriggers_')
    .addToUi();
  setupSheets_();
}

function setupSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  CADENCES.forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }
  });
  var settings = ss.getSheetByName('Settings');
  if (!settings) {
    settings = ss.insertSheet('Settings');
    settings.appendRow(['Key', 'Value']);
    settings.appendRow(['NotifyEmail', Session.getActiveUser().getEmail() || '']);
    settings.appendRow(['ReminderHour', '8']); // 24h, local script timezone
    settings.setFrozenRows(1);
    settings.getRange(1, 1, 1, 2).setFontWeight('bold');
  }
}

function showSidebarLink_() {
  var url = ScriptApp.getService().getUrl();
  var html = HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif;font-size:13px;">Deploy this project as a web app ' +
    '(Deploy &gt; New deployment &gt; Web app) to get a shareable link. ' +
    (url ? 'Current URL: <a href="' + url + '" target="_blank">' + url + '</a>' : 'Not deployed yet.') +
    '</p>'
  ).setWidth(360).setHeight(140);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cadence web app');
}

// ---------- Web app entry point ----------

function doGet(e) {
  setupSheets_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Cadence')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------- Helpers ----------

function sheet_(cadence) {
  var name = cadence.charAt(0).toUpperCase() + cadence.slice(1).toLowerCase();
  if (CADENCES.indexOf(name) === -1) throw new Error('Unknown cadence: ' + cadence);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) { setupSheets_(); sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
  return sheet;
}

function rowsToTasks_(values) {
  // values excludes header row
  return values
    .filter(function (r) { return r[0] !== '' && r[0] !== null; })
    .map(function (r) {
      return {
        id: String(r[0]),
        text: String(r[1]),
        urgency: r[2] || 'Medium',
        completed: r[3] === true || r[3] === 'TRUE',
        createdAt: r[4] instanceof Date ? r[4].toISOString() : String(r[4] || ''),
        reminderSentAt: r[5] instanceof Date ? r[5].toISOString() : String(r[5] || '')
      };
    });
}

// ---------- CRUD (called from Index.html via google.script.run) ----------

function getAllTasks() {
  var out = {};
  CADENCES.forEach(function (c) {
    var sheet = sheet_(c);
    var lastRow = sheet.getLastRow();
    var values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues() : [];
    out[c.toLowerCase()] = rowsToTasks_(values);
  });
  out.settings = getSettings();
  return out;
}

function addTask(cadence, text, urgency) {
  text = String(text || '').trim();
  if (!text) throw new Error('Task text is required.');
  urgency = ['Low', 'Medium', 'High'].indexOf(urgency) === -1 ? 'Medium' : urgency;
  var sheet = sheet_(cadence);
  var id = Utilities.getUuid();
  sheet.appendRow([id, text, urgency, false, new Date(), '']);
  return getAllTasks();
}

function findRowById_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2; // sheet row number
  }
  return -1;
}

function toggleTask(cadence, id) {
  var sheet = sheet_(cadence);
  var row = findRowById_(sheet, id);
  if (row === -1) throw new Error('Task not found.');
  var cell = sheet.getRange(row, 4);
  cell.setValue(!(cell.getValue() === true || cell.getValue() === 'TRUE'));
  return getAllTasks();
}

function updateUrgency(cadence, id, urgency) {
  urgency = ['Low', 'Medium', 'High'].indexOf(urgency) === -1 ? 'Medium' : urgency;
  var sheet = sheet_(cadence);
  var row = findRowById_(sheet, id);
  if (row === -1) throw new Error('Task not found.');
  sheet.getRange(row, 3).setValue(urgency);
  return getAllTasks();
}

function deleteTask(cadence, id) {
  var sheet = sheet_(cadence);
  var row = findRowById_(sheet, id);
  if (row === -1) throw new Error('Task not found.');
  sheet.deleteRow(row);
  return getAllTasks();
}

// ---------- Settings ----------

function getSettings() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!sheet) { setupSheets_(); sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings'); }
  var values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 2).getValues();
  var map = {};
  values.forEach(function (r) { if (r[0]) map[r[0]] = r[1]; });
  return {
    notifyEmail: map.NotifyEmail || Session.getActiveUser().getEmail() || '',
    reminderHour: Number(map.ReminderHour) || 8,
    triggerActive: ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'sendDailyDigest'; })
  };
}

function saveSettings(email, hour) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 2).getValues();
  var rowFor = {};
  values.forEach(function (r, i) { rowFor[r[0]] = i + 2; });
  if (rowFor.NotifyEmail) sheet.getRange(rowFor.NotifyEmail, 2).setValue(email);
  else sheet.appendRow(['NotifyEmail', email]);
  if (rowFor.ReminderHour) sheet.getRange(rowFor.ReminderHour, 2).setValue(hour);
  else sheet.appendRow(['ReminderHour', hour]);

  createDailyTrigger_(Number(hour));
  return getSettings();
}

// ---------- Reminder email ----------

var URGENCY_ORDER = { High: 0, Medium: 1, Low: 2 };
var URGENCY_EMOJI = { High: '🔴', Medium: '🟡', Low: '🟢' };

function sendDailyDigest() {
  var settings = getSettings();
  if (!settings.notifyEmail) return;

  var sections = [];
  var totalOpen = 0;

  CADENCES.forEach(function (c) {
    var sheet = sheet_(c);
    var lastRow = sheet.getLastRow();
    var values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues() : [];
    var tasks = rowsToTasks_(values).filter(function (t) { return !t.completed; });
    tasks.sort(function (a, b) { return URGENCY_ORDER[a.urgency] - URGENCY_ORDER[b.urgency]; });
    totalOpen += tasks.length;
    if (tasks.length > 0) {
      var lines = tasks.map(function (t) {
        return URGENCY_EMOJI[t.urgency] + ' [' + t.urgency + '] ' + t.text;
      });
      sections.push(c.toUpperCase() + ' (' + tasks.length + ' open)\n' + lines.join('\n'));
    }
  });

  if (totalOpen === 0) return; // nothing to remind about

  var body = 'Here\'s what\'s still open in Cadence:\n\n' + sections.join('\n\n') +
    '\n\n— Sent automatically by your Cadence tracker.';
  var highCount = 0;
  CADENCES.forEach(function (c) {
    var sheet = sheet_(c);
    var lastRow = sheet.getLastRow();
    var values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues() : [];
    highCount += rowsToTasks_(values).filter(function (t) { return !t.completed && t.urgency === 'High'; }).length;
  });
  var subject = (highCount > 0 ? '🔴 ' + highCount + ' urgent — ' : '') + 'Cadence: ' + totalOpen + ' open task' + (totalOpen === 1 ? '' : 's');

  MailApp.sendEmail(settings.notifyEmail, subject, body);

  // stamp ReminderSentAt on any task not yet stamped today (for your own records)
  var now = new Date();
  CADENCES.forEach(function (c) {
    var sheet = sheet_(c);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    var range = sheet.getRange(2, 4, lastRow - 1, 3); // Completed, CreatedAt, ReminderSentAt
    var values = range.getValues();
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] !== true) values[i][2] = now;
    }
    range.setValues(values);
  });
}

// ---------- Trigger management ----------

function removeTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendDailyDigest') ScriptApp.deleteTrigger(t);
  });
}

function createDailyTrigger_(hour) {
  removeTriggers_();
  ScriptApp.newTrigger('sendDailyDigest')
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();
}

function promptCreateTrigger_() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Daily reminder time', 'Send the digest email at which hour (0–23, 24h clock, your script timezone)?', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var hour = Number(resp.getResponseText());
  if (isNaN(hour) || hour < 0 || hour > 23) { ui.alert('Enter a number 0–23.'); return; }
  createDailyTrigger_(hour);
  ui.alert('Reminder trigger set for ' + hour + ':00 daily.');
}
