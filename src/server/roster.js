/**
 * roster.js — Section sheet management and roster synchronization.
 *
 * The Database tab is the source of truth for member names/sections.
 * syncRoster() creates/updates per-section tabs and refreshes all
 * form dropdowns to use canonical names.
 */

import { SheetTable } from './db.js';
import { getConfig, getAllConfig } from './config.js';

// Tabs that are never treated as section sheets
const SYSTEM_SHEETS = ['Data', 'Database', 'Yellow Sheets', 'Pink Sheets', 'Late Check-Ins'];

const SHEET_DATABASE = 'Database';
const DB_HEADERS = ['Last Name', 'First Name', 'Full Name', 'Section', 'Email', 'Instrument', 'Active'];

// Attendance dropdown values applied to every date cell
const ATTENDANCE_VALUES = ['Present', 'Tardy', 'Absent', 'Excused'];

/**
 * Ensure the Database tab exists with correct headers.
 */
export function ensureDatabaseSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_DATABASE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DATABASE);
    sheet.appendRow(DB_HEADERS);
    sheet.setFrozenRows(1);
  }
  const cfg = getAllConfig(ss);
  const hdrRange = sheet.getRange(1, 1, 1, DB_HEADERS.length);
  hdrRange.setBackground(cfg.COLOR_HEADER)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  return sheet;
}

/**
 * Sync Database → per-section roster sheets.
 * - Creates section tabs that don't exist.
 * - Adds new members; removes deactivated/deleted members.
 * - Preserves existing date columns and attendance data.
 * - Refreshes all form dropdowns.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
export function doSyncRoster(ss) {
  const dbSheet = ss.getSheetByName(SHEET_DATABASE);
  if (!dbSheet) {
    _alert('Database tab not found. Run Setup first.');
    return;
  }

  const db = new SheetTable(dbSheet);
  const allMembers = db.where({ Active: 'TRUE' });

  if (allMembers.length === 0) {
    _alert('No active members found in the Database tab.');
    return;
  }

  // Group by section
  const bySection = {};
  allMembers.forEach(member => {
    const section = _str(member['Section']);
    if (!section) return;
    if (!bySection[section]) bySection[section] = [];
    const fullName = _str(member['Full Name']) || `${_str(member['Last Name'])}, ${_str(member['First Name'])}`;
    bySection[section].push(fullName);
  });

  const cfg = getAllConfig(ss);

  // Sync each section sheet
  Object.entries(bySection).forEach(([section, names]) => {
    _syncSectionSheet(ss, section, names, cfg);
  });

  // Remove section sheets for sections that no longer exist
  const activeSections = new Set(Object.keys(bySection));
  getSectionSheets(ss).forEach(sheet => {
    if (!activeSections.has(sheet.getName())) {
      // Don't auto-delete — just log. Staff can delete manually.
      Logger.log(`Section sheet "${sheet.getName()}" has no active members.`);
    }
  });
}

/**
 * Sync a single section sheet with the given member names.
 * Preserves all date columns and their data.
 */
function _syncSectionSheet(ss, section, names, cfg) {
  let sheet = ss.getSheetByName(section);

  if (!sheet) {
    sheet = ss.insertSheet(section);
    const hdr = sheet.getRange(1, 1);
    hdr.setValue('Name');
    _styleHeaderCell(hdr, cfg);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 200);
  }

  // Build map of existing names: normalized → row number
  const data = sheet.getDataRange().getValues();
  const existingMap = new Map();
  for (let r = 1; r < data.length; r++) {
    const cell = _str(data[r][0]);
    if (cell) existingMap.set(_normalizeName(cell), { row: r + 1, display: cell });
  }

  // Desired names set (normalized)
  const desiredSet = new Set(names.map(n => _normalizeName(n)));

  // Remove rows for members no longer active (bottom-up to preserve indices)
  const rowsToDelete = [];
  existingMap.forEach(({ row }, key) => {
    if (!desiredSet.has(key)) rowsToDelete.push(row);
  });
  rowsToDelete.sort((a, b) => b - a);
  rowsToDelete.forEach(row => sheet.deleteRow(row));

  // Refresh the map after deletions
  const refreshed = sheet.getDataRange().getValues();
  const refreshedSet = new Set();
  for (let r = 1; r < refreshed.length; r++) {
    const cell = _str(refreshed[r][0]);
    if (cell) refreshedSet.add(_normalizeName(cell));
  }

  // Append new members
  let nextRow = refreshed.length + 1;
  names.sort().forEach(name => {
    const key = _normalizeName(name);
    if (!refreshedSet.has(key)) {
      sheet.getRange(nextRow, 1).setValue(name);
      // Apply attendance dropdown to any existing date columns for this row
      const lastCol = sheet.getLastColumn();
      if (lastCol >= 2) {
        _applyAttendanceValidation(sheet, nextRow, 2, lastCol);
      }
      nextRow++;
    }
  });
}

/**
 * Add a rehearsal date column to all section sheets.
 *
 * @param {string} dateStr  ISO date "YYYY-MM-DD"
 * @param {string} [timeStr]  Optional time label e.g. "3:30 PM"
 */
export function doAddRehearsalDate(dateStr, timeStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const cfg = getAllConfig(ss);

  const [y, m, d] = dateStr.split('-').map(Number);
  const date = new Date(y, m - 1, d, 12, 0, 0);
  const fmt = Utilities.formatDate(date, tz, 'M/d/yyyy');

  getSectionSheets(ss).forEach(sheet => {
    if (_findDateColumn(sheet, fmt, tz) > 0) return; // already exists

    const lastRow = sheet.getLastRow();
    const newCol = sheet.getLastColumn() + 1;

    // Header cell
    const cell = sheet.getRange(1, newCol);
    cell.setValue(date);
    cell.setNumberFormat('M/d');
    _styleHeaderCell(cell, cfg);
    sheet.setColumnWidth(newCol, 55);
    if (timeStr) cell.setNote(`Rehearsal time: ${timeStr}`);

    // Apply attendance dropdown to all member rows in this column
    if (lastRow >= 2) {
      _applyAttendanceValidation(sheet, 2, newCol, newCol, lastRow);
    }
  });

  _alert(`${fmt}${timeStr ? ' @ ' + timeStr : ''} added to all section sheets.`);
}

/**
 * Apply the attendance dropdown validation to a range of cells.
 */
function _applyAttendanceValidation(sheet, startRow, startCol, endCol, endRow) {
  endRow = endRow || startRow;
  const range = sheet.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ATTENDANCE_VALUES, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

// ─── Shared utilities (used by other modules) ───────────────────────────────

/**
 * Get all section sheets (excludes system tabs).
 */
export function getSectionSheets(ss) {
  return ss.getSheets().filter(s => !SYSTEM_SHEETS.includes(s.getName()));
}

/**
 * Find the 1-based column number for a formatted date in a sheet header.
 * Returns -1 if not found.
 */
export function findDateColumn(sheet, dateStr, tz) {
  return _findDateColumn(sheet, dateStr, tz);
}

function _findDateColumn(sheet, dateStr, tz) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let c = 1; c < headers.length; c++) {
    if (!headers[c]) continue;
    const d = new Date(headers[c]);
    if (!isNaN(d.getTime())) {
      const formatted = Utilities.formatDate(d, tz, 'M/d/yyyy');
      if (formatted === dateStr) return c + 1;
    }
  }
  return -1;
}

/**
 * Find or create a date column in a section sheet.
 * Returns the 1-based column number.
 */
export function getOrCreateDateColumn(sheet, dateStr, tz, cfg) {
  const existing = _findDateColumn(sheet, dateStr, tz);
  if (existing > 0) return existing;

  const lastRow = sheet.getLastRow();
  const newCol = sheet.getLastColumn() + 1;
  const cell = sheet.getRange(1, newCol);
  const d = new Date(dateStr);
  cell.setValue(d);
  cell.setNumberFormat('M/d');
  _styleHeaderCell(cell, cfg);
  sheet.setColumnWidth(newCol, 55);

  // Apply dropdown to all member rows
  if (lastRow >= 2) {
    _applyAttendanceValidation(sheet, 2, newCol, newCol, lastRow);
  }

  return newCol;
}

/**
 * Find a member's 1-based row in a section sheet.
 * Column A stores "Last, First". Input can be "First Last" or "Last, First".
 * Returns -1 if not found.
 */
export function findMemberRow(sheet, name) {
  const data = sheet.getDataRange().getValues();
  const target = _normalizeName(name);
  for (let r = 1; r < data.length; r++) {
    if (_normalizeName(_str(data[r][0])) === target) return r + 1;
  }
  return -1;
}

/**
 * Get all active member names from the Database as "Last, First" strings.
 */
export function getActiveMemberNames(ss) {
  const dbSheet = ss.getSheetByName(SHEET_DATABASE);
  if (!dbSheet) return [];
  const db = new SheetTable(dbSheet);
  return db.where({ Active: 'TRUE' })
    .map(r => _str(r['Full Name']) || `${_str(r['Last Name'])}, ${_str(r['First Name'])}`)
    .filter(n => n && n !== ', ')
    .sort();
}

/**
 * Get all active sections from the Database.
 */
export function getActiveSections(ss) {
  const dbSheet = ss.getSheetByName(SHEET_DATABASE);
  if (!dbSheet) return [];
  const db = new SheetTable(dbSheet);
  return db.distinct('Section', { Active: 'TRUE' });
}

/**
 * Get active member names for a specific section.
 */
export function getMembersBySection(ss, section) {
  const dbSheet = ss.getSheetByName(SHEET_DATABASE);
  if (!dbSheet) return [];
  const db = new SheetTable(dbSheet);
  return db.where({ Section: section, Active: 'TRUE' })
    .map(r => _str(r['Full Name']) || `${_str(r['Last Name'])}, ${_str(r['First Name'])}`)
    .filter(n => n && n !== ', ')
    .sort();
}

// ─── Private helpers ────────────────────────────────────────────────────────

/**
 * Normalize a name to "first last" lowercase for comparison.
 * Handles both "Last, First" and "First Last".
 */
function _normalizeName(name) {
  const s = String(name).toLowerCase().trim();
  const commaIdx = s.indexOf(',');
  if (commaIdx > 0) {
    return `${s.slice(commaIdx + 1).trim()} ${s.slice(0, commaIdx).trim()}`;
  }
  return s;
}

function _styleHeaderCell(cell, cfg) {
  const color = (cfg && cfg.COLOR_HEADER) || '#1155CC';
  cell.setBackground(color)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

function _str(val) {
  return val === undefined || val === null ? '' : String(val).trim();
}

function _alert(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (_) {
    Logger.log(msg);
  }
}

export { SYSTEM_SHEETS, SHEET_DATABASE, DB_HEADERS, ATTENDANCE_VALUES };
