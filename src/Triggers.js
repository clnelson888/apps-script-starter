/**
 * Triggers.js — All installable trigger handlers.
 *
 * - onEditGlobal:              Yellow sheet "Approved" -> highlight member name
 * - processYellowSheetSubmit:  Log to Yellow Sheets tab (Status = Pending)
 * - processPinkSheetSubmit:    Log + mark cell Excused with pink highlight
 * - processLateCheckInSubmit:  Log + mark Present or Tardy based on threshold
 */

/* exported installTriggers, onEditGlobal,
   processYellowSheetSubmit, processPinkSheetSubmit, processLateCheckInSubmit,
   SHEET_YELLOW, SHEET_PINK, SHEET_LATE */

var SHEET_YELLOW = 'Yellow Sheets';
var SHEET_PINK = 'Pink Sheets';
var SHEET_LATE = 'Late Check-Ins';

/**
 * Install all needed triggers, removing stale duplicates first.
 */
function installTriggers(ss) {
  var handlerNames = ['onEditGlobal', 'processYellowSheetSubmit', 'processPinkSheetSubmit', 'processLateCheckInSubmit'];

  // Remove existing triggers for our handlers
  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    if (handlerNames.indexOf(triggers[t].getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(triggers[t]);
    }
  }

  // Spreadsheet onEdit
  ScriptApp.newTrigger('onEditGlobal').forSpreadsheet(ss).onEdit().create();

  // Form-level onFormSubmit triggers
  var formTriggers = [
    { key: 'YELLOW_FORM_ID', handler: 'processYellowSheetSubmit' },
    { key: 'PINK_FORM_ID', handler: 'processPinkSheetSubmit' },
    { key: 'LATE_FORM_ID', handler: 'processLateCheckInSubmit' },
  ];

  for (var i = 0; i < formTriggers.length; i++) {
    var formId = getConfig(formTriggers[i].key, ss);
    if (formId) {
      ScriptApp.newTrigger(formTriggers[i].handler).forForm(FormApp.openById(formId)).onFormSubmit().create();
    }
  }
}

// ─── Trigger handlers ───────────────────────────────────────────────────────

/**
 * Spreadsheet onEdit — watches Yellow Sheets for Status -> "Approved".
 * Highlights the member's name cell yellow in their section sheet.
 */
function onEditGlobal(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_YELLOW) return;

  var newVal = _str(e.value).toLowerCase();
  if (newVal !== 'approved') return;

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var statusCol = _headerIndex(headers, 'Status') + 1;
  if (e.range.getColumn() !== statusCol) return;

  var row = e.range.getRow();
  var rowData = data[row - 1];
  var name = _str(rowData[_headerIndex(headers, 'Full Name')]);
  var section = _str(rowData[_headerIndex(headers, 'Section')]);
  var conflictDays = _str(rowData[_headerIndex(headers, 'Conflict Days')]);

  if (!name || !section) return;

  var ss = e.source;
  var sectionSheet = ss.getSheetByName(section);
  if (!sectionSheet) return;

  var memberRow = findMemberRow(sectionSheet, name);
  if (memberRow < 0) return;

  var cfg = getAllConfig(ss);
  var nameCell = sectionSheet.getRange(memberRow, 1);
  nameCell.setBackground(cfg.COLOR_YELLOW);
  nameCell.setNote(conflictDays ? `Approved Class Conflict: ${conflictDays}` : 'Approved class conflict on file');
}

/**
 * Yellow Sheet form onFormSubmit — logs to Yellow Sheets tab.
 */
function processYellowSheetSubmit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_YELLOW);
  if (!sheet) return;

  var resp = parseFormResponse(e.response);

  // Conflict Days comes as an array from checkboxes — join to comma string
  var conflictDays = resp[Q_YELLOW_DAYS];
  if (Array.isArray(conflictDays)) {
    conflictDays = conflictDays.join(', ');
  }

  // Time items return "HH:MM" strings
  var startTime = resp[Q_YELLOW_START] || '';
  var endTime = resp[Q_YELLOW_END] || '';

  sheet.appendRow([
    resp[Q_YELLOW_NAME] || '',
    resp[Q_YELLOW_ENSEMBLE] || '',
    resp[Q_YELLOW_SECTION] || '',
    conflictDays || '',
    startTime,
    endTime,
    'Pending',
    resp[Q_YELLOW_NOTES] || '',
  ]);
}

/**
 * Pink Sheet form onFormSubmit — logs and marks cell Excused with pink highlight.
 */
function processPinkSheetSubmit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PINK);
  if (!sheet) return;

  var resp = parseFormResponse(e.response);
  var tz = ss.getSpreadsheetTimeZone();
  var cfg = getAllConfig(ss);

  var name = resp[Q_PINK_NAME] || '';
  var section = resp[Q_PINK_SECTION] || '';
  var absenceDate = resp[Q_PINK_DATE];
  var reason = resp[Q_PINK_REASON] || '';

  var absenceDateStr;
  if (absenceDate instanceof Date) {
    absenceDateStr = Utilities.formatDate(absenceDate, tz, 'M/d/yyyy');
  } else {
    absenceDateStr = _str(absenceDate);
  }

  // Log to Pink Sheets tab
  sheet.appendRow([name, resp[Q_PINK_ENSEMBLE] || '', section, absenceDateStr, reason]);

  // Mark the cell in the section sheet
  if (name && section && absenceDateStr) {
    var sectionSheet = ss.getSheetByName(section);
    if (!sectionSheet) return;

    var memberRow = findMemberRow(sectionSheet, name);
    if (memberRow < 0) return;

    var dateCol = getOrCreateDateColumn(sectionSheet, absenceDateStr, tz, cfg);
    var cell = sectionSheet.getRange(memberRow, dateCol);

    cell.setValue('Excused');
    cell.setBackground(cfg.COLOR_PINK);
    cell.setNote(`Excused: ${reason}`);
  }
}

/**
 * Late Check-In form onFormSubmit.
 * Marks the member Present or Tardy based on the late threshold.
 */
function processLateCheckInSubmit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_LATE);
  if (!sheet) return;

  var resp = parseFormResponse(e.response);
  var tz = ss.getSpreadsheetTimeZone();
  var cfg = getAllConfig(ss);

  var timestamp = resp.timestamp;
  var section = resp[Q_LATE_SECTION] || '';
  var name = resp[Q_LATE_NAME] || '';
  var reason = resp[Q_LATE_REASON] || '';
  var reasonOther = resp[Q_LATE_OTHER] || '';
  var fullReason = reason === 'Other (explain below)' && reasonOther ? `Other: ${reasonOther}` : reason;

  var arrivalTime = Utilities.formatDate(timestamp, tz, 'h:mm a');
  var today = Utilities.formatDate(timestamp, tz, 'M/d/yyyy');

  // Log to Late Check-Ins tab
  sheet.appendRow([name, section, arrivalTime, fullReason]);

  if (!name || !section) return;

  var status = _calculateLateStatus(timestamp, tz, cfg);

  var sectionSheet = ss.getSheetByName(section);
  if (!sectionSheet) return;

  var memberRow = findMemberRow(sectionSheet, name);
  if (memberRow < 0) return;

  var dateCol = getOrCreateDateColumn(sectionSheet, today, tz, cfg);
  var cell = sectionSheet.getRange(memberRow, dateCol);

  cell.setValue(status);
  var noteText = `Late arrival: ${arrivalTime}. Reason: ${fullReason}`;
  var existing = cell.getNote();
  cell.setNote(existing ? `${existing}\n${noteText}` : noteText);
}

// ─── Private helpers ────────────────────────────────────────────────────────

function _calculateLateStatus(arrivalTimestamp, _tz, cfg) {
  var startTimeStr = cfg.REHEARSAL_START_TIME || '15:30';
  var thresholdMin = parseInt(cfg.LATE_THRESHOLD_MIN, 10) || 45;

  var parts = startTimeStr.split(':');
  var startH = Number(parts[0]);
  var startM = Number(parts[1]);

  var arrivalDate = new Date(arrivalTimestamp);
  var cutoff = new Date(arrivalDate);
  cutoff.setHours(startH, startM + thresholdMin, 0, 0);

  return arrivalDate <= cutoff ? 'Present' : 'Tardy';
}

function _headerIndex(headers, name) {
  for (var i = 0; i < headers.length; i++) {
    if (_str(headers[i]) === name) return i;
  }
  return -1;
}
