/**
 * Roster.js — Database sheet management and roster synchronization.
 *
 * The Database tab is the source of truth for member names and sections.
 * doSyncRoster() creates/updates per-section tabs from the Database.
 */

/* exported ensureDatabaseSheet, doSyncRoster, getSectionSheets, getActiveMemberNames,
   getActiveSections, getMembersBySection, findMemberRow,
   SYSTEM_SHEETS, SHEET_DATABASE, DB_HEADERS, ATTENDANCE_VALUES */

var SYSTEM_SHEETS = ['Data', 'Database', 'Yellow Sheets', 'Pink Sheets', 'Late Check-Ins'];
var SHEET_DATABASE = 'Database';
var DB_HEADERS = ['Last Name', 'First Name', 'Full Name', 'Section', 'Email', 'Instrument', 'Active'];
var ATTENDANCE_VALUES = ['Present', 'Tardy', 'Absent', 'Excused'];

/**
 * Ensure the Database tab exists with correct headers.
 */
function ensureDatabaseSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_DATABASE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DATABASE);
    sheet.appendRow(DB_HEADERS);
    sheet.setFrozenRows(1);
  }
  try {
    var cfg = getAllConfig(ss);
    var hdrRange = sheet.getRange(1, 1, 1, DB_HEADERS.length);
    hdrRange.setBackground(cfg.COLOR_HEADER).setFontColor('#FFFFFF').setFontWeight('bold');
  } catch (e) {
    Logger.log(`Could not style Database header: ${e.message}`);
  }
  return sheet;
}

/**
 * Sync Database -> per-section roster sheets.
 * Creates section tabs that don't exist, adds/removes members, preserves date columns.
 */
function doSyncRoster(ss) {
  var dbSheet = ss.getSheetByName(SHEET_DATABASE);
  if (!dbSheet) {
    _alert('Database tab not found. Run Setup first.');
    return;
  }

  var allMembers = _getActiveMembers(dbSheet);
  if (allMembers.length === 0) {
    _alert('No active members found in the Database tab.');
    return;
  }

  // Group by section
  var bySection = {};
  for (var i = 0; i < allMembers.length; i++) {
    var member = allMembers[i];
    var section = member.section;
    if (!section) continue;
    if (!bySection[section]) bySection[section] = [];
    bySection[section].push(member.fullName);
  }

  var cfg = getAllConfig(ss);

  // Sync each section sheet
  var sections = Object.keys(bySection);
  for (var s = 0; s < sections.length; s++) {
    _syncSectionSheet(ss, sections[s], bySection[sections[s]], cfg);
  }

  // Log sections with no active members (don't auto-delete)
  var activeSections = {};
  for (var k = 0; k < sections.length; k++) {
    activeSections[sections[k]] = true;
  }
  var sectionSheets = getSectionSheets(ss);
  for (var j = 0; j < sectionSheets.length; j++) {
    if (!activeSections[sectionSheets[j].getName()]) {
      Logger.log(`Section sheet "${sectionSheets[j].getName()}" has no active members.`);
    }
  }
}

/**
 * Get all section sheets (excludes system tabs).
 */
function getSectionSheets(ss) {
  return ss.getSheets().filter((s) => {
    return SYSTEM_SHEETS.indexOf(s.getName()) === -1;
  });
}

/**
 * Get all active member names from the Database as "Last, First" strings.
 */
function getActiveMemberNames(ss) {
  var dbSheet = ss.getSheetByName(SHEET_DATABASE);
  if (!dbSheet) return [];
  var names = _getActiveMembers(dbSheet)
    .map((m) => {
      return m.fullName;
    })
    .filter((n) => {
      return n && n !== ', ';
    });
  return _dedupe(names).sort();
}

/**
 * Get all active sections from the Database.
 */
function getActiveSections(ss) {
  var dbSheet = ss.getSheetByName(SHEET_DATABASE);
  if (!dbSheet) return [];
  var members = _getActiveMembers(dbSheet);
  var seen = {};
  var result = [];
  for (var i = 0; i < members.length; i++) {
    var sec = members[i].section;
    if (sec && !seen[sec]) {
      seen[sec] = true;
      result.push(sec);
    }
  }
  return result.sort();
}

/**
 * Get active member names for a specific section.
 */
function getMembersBySection(ss, section) {
  var dbSheet = ss.getSheetByName(SHEET_DATABASE);
  if (!dbSheet) return [];
  var names = _getActiveMembers(dbSheet)
    .filter((m) => {
      return m.section === section;
    })
    .map((m) => {
      return m.fullName;
    })
    .filter((n) => {
      return n && n !== ', ';
    });
  return _dedupe(names).sort();
}

/**
 * Find a member's 1-based row in a section sheet.
 * Column A stores "Last, First". Returns -1 if not found.
 */
function findMemberRow(sheet, name) {
  var data = sheet.getDataRange().getValues();
  var target = _normalizeName(name);
  for (var r = 1; r < data.length; r++) {
    if (_normalizeName(_str(data[r][0])) === target) return r + 1;
  }
  return -1;
}

// ─── Private helpers ────────────────────────────────────────────────────────

/**
 * Read Database sheet and return active members as objects.
 */
function _getActiveMembers(dbSheet) {
  var data = dbSheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0].map((h) => {
    return String(h).trim();
  });
  var colIdx = {};
  for (var c = 0; c < headers.length; c++) {
    colIdx[headers[c]] = c;
  }

  var result = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var active = _str(row[colIdx['Active']]).toUpperCase();
    if (active !== 'TRUE') continue;

    var fullName = _str(row[colIdx['Full Name']]);
    if (!fullName) {
      var last = _str(row[colIdx['Last Name']]);
      var first = _str(row[colIdx['First Name']]);
      fullName = `${last}, ${first}`;
    }

    result.push({
      fullName,
      section: _str(row[colIdx['Section']]),
    });
  }
  return result;
}

/**
 * Sync a single section sheet with the given member names.
 */
function _syncSectionSheet(ss, section, names, cfg) {
  var sheet = ss.getSheetByName(section);

  if (!sheet) {
    sheet = ss.insertSheet(section);
    var hdr = sheet.getRange(1, 1);
    hdr.setValue('Name');
    _styleHeaderCell(hdr, cfg);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 200);
  }

  // Build map of existing names: normalized -> row number
  var data = sheet.getDataRange().getValues();
  var existingMap = {};
  for (var r = 1; r < data.length; r++) {
    var cell = _str(data[r][0]);
    if (cell) existingMap[_normalizeName(cell)] = r + 1;
  }

  // Desired names set (normalized)
  var desiredSet = {};
  for (var i = 0; i < names.length; i++) {
    desiredSet[_normalizeName(names[i])] = true;
  }

  // Remove rows for members no longer active (bottom-up)
  var rowsToDelete = [];
  var keys = Object.keys(existingMap);
  for (var k = 0; k < keys.length; k++) {
    if (!desiredSet[keys[k]]) rowsToDelete.push(existingMap[keys[k]]);
  }
  rowsToDelete.sort((a, b) => {
    return b - a;
  });
  for (var d = 0; d < rowsToDelete.length; d++) {
    try {
      sheet.deleteRow(rowsToDelete[d]);
    } catch (e) {
      Logger.log(`Could not delete row ${rowsToDelete[d]} on ${section}: ${e.message}`);
    }
  }

  // Refresh map after deletions
  var refreshed = sheet.getDataRange().getValues();
  var refreshedSet = {};
  for (var rr = 1; rr < refreshed.length; rr++) {
    var c2 = _str(refreshed[rr][0]);
    if (c2) refreshedSet[_normalizeName(c2)] = true;
  }

  // Append new members
  var sortedNames = names.slice().sort();
  var nextRow = refreshed.length + 1;
  for (var n = 0; n < sortedNames.length; n++) {
    var key = _normalizeName(sortedNames[n]);
    if (!refreshedSet[key]) {
      sheet.getRange(nextRow, 1).setValue(sortedNames[n]);
      var lastCol = sheet.getLastColumn();
      if (lastCol >= 2) {
        _tryApplyAttendanceValidation(sheet, nextRow, 2, lastCol);
      }
      nextRow++;
    }
  }
}

/**
 * Normalize a name to "first last" lowercase for comparison.
 */
function _normalizeName(name) {
  var s = String(name).toLowerCase().trim();
  var commaIdx = s.indexOf(',');
  if (commaIdx > 0) {
    return `${s.slice(commaIdx + 1).trim()} ${s.slice(0, commaIdx).trim()}`;
  }
  return s;
}

function _styleHeaderCell(cell, cfg) {
  var color = (cfg && cfg.COLOR_HEADER) || '#1155CC';
  cell.setBackground(color).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
}

function _dedupe(arr) {
  var seen = {};
  var result = [];
  for (var i = 0; i < arr.length; i++) {
    if (!seen[arr[i]]) {
      seen[arr[i]] = true;
      result.push(arr[i]);
    }
  }
  return result;
}

function _str(val) {
  return val === undefined || val === null ? '' : String(val).trim();
}

function _alert(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (_e) {
    Logger.log(msg);
  }
}
