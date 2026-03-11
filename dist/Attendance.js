/**
 * Attendance.js — Add rehearsal dates and attendance column helpers.
 *
 * Handles adding date columns to section sheets and applying
 * attendance dropdown validation (skipped on structured table sheets).
 */

/* exported doAddRehearsalDate, getOrCreateDateColumn, findDateColumn */

/**
 * Add a rehearsal date column to all section sheets.
 * @param {string} dateStr  ISO date "YYYY-MM-DD"
 * @param {string} [timeStr]  Optional time label e.g. "3:30 PM"
 */
function doAddRehearsalDate(dateStr, timeStr) {
  var ss = _getSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var cfg = getAllConfig(ss);

  var parts = dateStr.split('-');
  var date = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 12, 0, 0);
  var fmt = Utilities.formatDate(date, tz, 'M/d/yyyy');

  var sheets = getSectionSheets(ss);
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if (_findDateCol(sheet, fmt, tz) > 0) continue;

    var lastRow = sheet.getLastRow();
    var newCol = sheet.getLastColumn() + 1;

    var cell = sheet.getRange(1, newCol);
    try {
      cell.setValue(date);
      cell.setNumberFormat('M/d');
      _styleHeaderCell(cell, cfg);
      sheet.setColumnWidth(newCol, 55);
      if (timeStr) cell.setNote(`Rehearsal time: ${timeStr}`);
    } catch (e) {
      Logger.log(`Could not style header on ${sheet.getName()}: ${e.message}`);
    }

    if (lastRow >= 2) {
      _tryApplyAttendanceValidation(sheet, 2, newCol, newCol, lastRow);
    }
  }

  _alert(`${fmt + (timeStr ? ` @ ${timeStr}` : '')} added to all section sheets.`);
}

/**
 * Find the 1-based column number for a formatted date in a sheet header.
 * Returns -1 if not found.
 */
function findDateColumn(sheet, dateStr, tz) {
  return _findDateCol(sheet, dateStr, tz);
}

/**
 * Find or create a date column in a section sheet.
 * Returns the 1-based column number.
 */
function getOrCreateDateColumn(sheet, dateStr, tz, cfg) {
  var existing = _findDateCol(sheet, dateStr, tz);
  if (existing > 0) return existing;

  var lastRow = sheet.getLastRow();
  var newCol = sheet.getLastColumn() + 1;
  var cell = sheet.getRange(1, newCol);
  try {
    var d = new Date(dateStr);
    cell.setValue(d);
    cell.setNumberFormat('M/d');
    _styleHeaderCell(cell, cfg);
    sheet.setColumnWidth(newCol, 55);
  } catch (e) {
    Logger.log(`Could not style date column on ${sheet.getName()}: ${e.message}`);
  }

  if (lastRow >= 2) {
    _tryApplyAttendanceValidation(sheet, 2, newCol, newCol, lastRow);
  }

  return newCol;
}

// ─── Private helpers ────────────────────────────────────────────────────────

function _findDateCol(sheet, dateStr, tz) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var c = 1; c < headers.length; c++) {
    if (!headers[c]) continue;
    var d = new Date(headers[c]);
    if (!isNaN(d.getTime())) {
      var formatted = Utilities.formatDate(d, tz, 'M/d/yyyy');
      if (formatted === dateStr) return c + 1;
    }
  }
  return -1;
}

/**
 * Try to apply attendance dropdown validation. Silently skips if the
 * sheet uses structured tables with typed columns.
 */
function _tryApplyAttendanceValidation(sheet, startRow, startCol, endCol, endRow) {
  endRow = endRow || startRow;
  try {
    var range = sheet.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(ATTENDANCE_VALUES, true)
      .setAllowInvalid(false)
      .build();
    range.setDataValidation(rule);
  } catch (e) {
    // Structured tables manage their own column types — skip gracefully
    Logger.log(`Skipped validation on ${sheet.getName()}: ${e.message}`);
  }
}
