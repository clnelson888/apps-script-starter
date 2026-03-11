/**
 * triggers.js — All installable trigger handlers.
 *
 * - onEditGlobal:              Yellow sheet "Approved" → highlight member name
 * - processYellowSheetSubmit:  Log to Yellow Sheets tab (Status = Pending)
 * - processPinkSheetSubmit:    Log + mark cell Excused with pink highlight
 * - processLateCheckInSubmit:  Log + mark Present or Tardy based on threshold
 */

import { SheetTable } from './db.js';
import { getConfig, getAllConfig } from './config.js';
import {
  findMemberRow, findDateColumn, getOrCreateDateColumn, getSectionSheets,
} from './roster.js';
import {
  parseFormResponse,
  Q_NAME, Q_SECTION, Q_LATE_SECTION, Q_LATE_NAME,
  Q_LATE_REASON, Q_LATE_OTHER, Q_CONFLICT_DESC,
  Q_CONFLICT_TIMES, Q_ABSENCE_DATE, Q_REASON,
} from './forms.js';

const SHEET_YELLOW = 'Yellow Sheets';
const SHEET_PINK   = 'Pink Sheets';
const SHEET_LATE   = 'Late Check-Ins';

// ─── Yellow Sheet Headers ───────────────────────────────────────────────────
const YELLOW_HEADERS = ['Timestamp', 'Full Name', 'Section', 'Conflict Description', 'Conflict Days/Times', 'Status', 'Notes'];
const PINK_HEADERS   = ['Timestamp', 'Full Name', 'Section', 'Absence Date', 'Reason'];
const LATE_HEADERS   = ['Timestamp', 'Full Name', 'Section', 'Arrival Time', 'Reason'];

/**
 * Ensure the three log sheets (Yellow, Pink, Late) exist with correct headers.
 */
export function ensureLogSheets(ss) {
  const cfg = getAllConfig(ss);
  _ensureSheet(ss, SHEET_YELLOW, YELLOW_HEADERS, cfg);
  _ensureSheet(ss, SHEET_PINK, PINK_HEADERS, cfg);
  _ensureSheet(ss, SHEET_LATE, LATE_HEADERS, cfg);

  // Yellow Sheets: add Status dropdown to column F
  const yellowSheet = ss.getSheetByName(SHEET_YELLOW);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Approved', 'Denied'], true)
    .setAllowInvalid(false)
    .build();
  yellowSheet.getRange(2, 6, yellowSheet.getMaxRows() - 1, 1).setDataValidation(statusRule);
}

/**
 * Install all needed triggers, removing stale duplicates first.
 */
export function installTriggers(ss) {
  const handlerNames = [
    'onEditGlobal',
    'processYellowSheetSubmit',
    'processPinkSheetSubmit',
    'processLateCheckInSubmit',
  ];

  // Remove existing triggers for our handlers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlerNames.includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Spreadsheet onEdit
  ScriptApp.newTrigger('onEditGlobal')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // Form-level onFormSubmit triggers
  const formTriggers = [
    { key: 'YELLOW_FORM_ID', handler: 'processYellowSheetSubmit' },
    { key: 'PINK_FORM_ID',   handler: 'processPinkSheetSubmit' },
    { key: 'LATE_FORM_ID',   handler: 'processLateCheckInSubmit' },
  ];

  formTriggers.forEach(({ key, handler }) => {
    const formId = getConfig(key, ss);
    if (formId) {
      ScriptApp.newTrigger(handler)
        .forForm(FormApp.openById(formId))
        .onFormSubmit()
        .create();
    }
  });
}

// ─── Trigger handlers (exported to global scope via index.js) ───────────────

/**
 * Spreadsheet onEdit — watches Yellow Sheets for Status → "Approved".
 * Highlights the member's name cell yellow in their section sheet
 * and adds a comment with the conflict times.
 */
export function onEditGlobal(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_YELLOW) return;

  const newVal = _str(e.value).toLowerCase();
  if (newVal !== 'approved') return;

  const data      = sheet.getDataRange().getValues();
  const headers   = data[0];
  const statusCol = headers.indexOf('Status') + 1;
  if (e.range.getColumn() !== statusCol) return;

  const row           = e.range.getRow();
  const rowData       = data[row - 1];
  const name          = _str(rowData[headers.indexOf('Full Name')]);
  const section       = _str(rowData[headers.indexOf('Section')]);
  const conflictTimes = _str(rowData[headers.indexOf('Conflict Days/Times')]);

  if (!name || !section) return;

  const ss           = e.source;
  const sectionSheet = ss.getSheetByName(section);
  if (!sectionSheet) return;

  const memberRow = findMemberRow(sectionSheet, name);
  if (memberRow < 0) return;

  const cfg       = getAllConfig(ss);
  const nameCell  = sectionSheet.getRange(memberRow, 1);
  nameCell.setBackground(cfg.COLOR_YELLOW);
  nameCell.setNote(
    conflictTimes
      ? `Approved Class Conflict: ${conflictTimes}`
      : 'Approved class conflict on file'
  );
}

/**
 * Yellow Sheet form onFormSubmit — logs to Yellow Sheets tab.
 */
export function processYellowSheetSubmit(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_YELLOW);
  if (!sheet) return;

  const resp = parseFormResponse(e.response);
  sheet.appendRow([
    resp.timestamp,
    resp[Q_NAME] || '',
    resp[Q_SECTION] || '',
    resp[Q_CONFLICT_DESC] || '',
    resp[Q_CONFLICT_TIMES] || '',
    'Pending',
    '',
  ]);
}

/**
 * Pink Sheet form onFormSubmit — logs and marks cell Excused with pink highlight.
 */
export function processPinkSheetSubmit(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PINK);
  if (!sheet) return;

  const resp = parseFormResponse(e.response);
  const tz   = ss.getSpreadsheetTimeZone();
  const cfg  = getAllConfig(ss);

  const name        = resp[Q_NAME] || '';
  const section     = resp[Q_SECTION] || '';
  const absenceDate = resp[Q_ABSENCE_DATE];
  const reason      = resp[Q_REASON] || '';

  const absenceDateStr = absenceDate instanceof Date
    ? Utilities.formatDate(absenceDate, tz, 'M/d/yyyy')
    : _str(absenceDate);

  // Log to Pink Sheets tab
  sheet.appendRow([resp.timestamp, name, section, absenceDateStr, reason]);

  // Mark the cell in the section sheet
  if (name && section && absenceDateStr) {
    const sectionSheet = ss.getSheetByName(section);
    if (!sectionSheet) return;

    const memberRow = findMemberRow(sectionSheet, name);
    if (memberRow < 0) return;

    const dateCol = getOrCreateDateColumn(sectionSheet, absenceDateStr, tz, cfg);
    const cell    = sectionSheet.getRange(memberRow, dateCol);

    cell.setValue('Excused');
    cell.setBackground(cfg.COLOR_PINK);
    cell.setNote(`Excused: ${reason}`);
  }
}

/**
 * Late Check-In form onFormSubmit.
 * Marks the member Present or Tardy based on the late threshold,
 * and adds a comment with the arrival time and reason.
 */
export function processLateCheckInSubmit(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LATE);
  if (!sheet) return;

  const resp = parseFormResponse(e.response);
  const tz   = ss.getSpreadsheetTimeZone();
  const cfg  = getAllConfig(ss);

  const timestamp   = resp.timestamp;
  const section     = resp[Q_LATE_SECTION] || '';
  const name        = resp[Q_LATE_NAME] || '';
  const reason      = resp[Q_LATE_REASON] || '';
  const reasonOther = resp[Q_LATE_OTHER] || '';
  const fullReason  = reason === 'Other (explain below)' && reasonOther
    ? `Other: ${reasonOther}`
    : reason;

  const arrivalTime = Utilities.formatDate(timestamp, tz, 'h:mm a');
  const today       = Utilities.formatDate(timestamp, tz, 'M/d/yyyy');

  // Log to Late Check-Ins tab
  sheet.appendRow([timestamp, name, section, arrivalTime, fullReason]);

  if (!name || !section) return;

  // Determine Present vs Tardy based on threshold
  const status = _calculateLateStatus(timestamp, tz, cfg);

  const sectionSheet = ss.getSheetByName(section);
  if (!sectionSheet) return;

  const memberRow = findMemberRow(sectionSheet, name);
  if (memberRow < 0) return;

  const dateCol = getOrCreateDateColumn(sectionSheet, today, tz, cfg);
  const cell    = sectionSheet.getRange(memberRow, dateCol);

  cell.setValue(status);
  const noteText = `Late arrival: ${arrivalTime}. Reason: ${fullReason}`;
  const existing = cell.getNote();
  cell.setNote(existing ? `${existing}\n${noteText}` : noteText);
}

// ─── Private helpers ────────────────────────────────────────────────────────

/**
 * Determine whether a late arrival is "Present" or "Tardy" based on
 * the REHEARSAL_START_TIME and LATE_THRESHOLD_MIN from the Data tab.
 */
function _calculateLateStatus(arrivalTimestamp, tz, cfg) {
  const startTimeStr = cfg.REHEARSAL_START_TIME || '15:30';
  const thresholdMin = parseInt(cfg.LATE_THRESHOLD_MIN, 10) || 45;

  // Parse the rehearsal start time
  const [startH, startM] = startTimeStr.split(':').map(Number);

  // Build a Date for "rehearsal start + threshold" on the same day
  const arrivalDate = new Date(arrivalTimestamp);
  const cutoff = new Date(arrivalDate);
  cutoff.setHours(startH, startM + thresholdMin, 0, 0);

  return arrivalDate <= cutoff ? 'Present' : 'Tardy';
}

function _ensureSheet(ss, name, headers, cfg) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  const hdrRange = sheet.getRange(1, 1, 1, headers.length);
  hdrRange.setBackground(cfg.COLOR_HEADER)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  return sheet;
}

function _str(val) {
  return val === undefined || val === null ? '' : String(val).trim();
}
