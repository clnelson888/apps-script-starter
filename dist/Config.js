/**
 * Config.js — Reads/writes system configuration from the "Data" tab.
 *
 * The Data tab stores key-value pairs in columns A (key) and B (value).
 * All form IDs, colors, thresholds, and settings live here instead of
 * being hardcoded in the script.
 */

/* exported getConfig, getAllConfig, setConfig, ensureDataSheet, _getSpreadsheet, DATA_SHEET_NAME, CONFIG_DEFAULTS */

var DATA_SHEET_NAME = 'Data';

var CONFIG_DEFAULTS = {
  YELLOW_FORM_ID: '',
  PINK_FORM_ID: '',
  LATE_FORM_ID: '',
  COLOR_YELLOW: '#FFD966',
  COLOR_PINK: '#FF91A4',
  COLOR_HEADER: '#1155CC',
  REHEARSAL_START_TIME: '15:30',
  LATE_THRESHOLD_MIN: '45',
  CONCERN_INCLUDE_TARDY: 'TRUE',
};

/**
 * Get a single config value by key.
 * Falls back to CONFIG_DEFAULTS if the key exists there but not in the sheet.
 */
function getConfig(key, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) return CONFIG_DEFAULTS[key] || '';

  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      return String(data[i][1]).trim();
    }
  }
  return CONFIG_DEFAULTS[key] || '';
}

/**
 * Get all config values as a plain object.
 */
function getAllConfig(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var result = {};
  var key;
  for (key in CONFIG_DEFAULTS) {
    result[key] = CONFIG_DEFAULTS[key];
  }

  var sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) return result;

  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    key = String(data[i][0]).trim();
    if (key) result[key] = String(data[i][1]).trim();
  }
  return result;
}

/**
 * Set a single config value. Creates/updates the row for that key.
 */
function setConfig(key, value, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

/**
 * Ensure the Data tab exists with all default keys populated.
 * Does not overwrite existing values.
 */
function ensureDataSheet(ss) {
  var sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DATA_SHEET_NAME);
    var hdr = sheet.getRange(1, 1, 1, 2);
    hdr.setValues([['Key', 'Value']]);
    hdr.setBackground('#1155CC').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 240);
    sheet.setColumnWidth(2, 300);
  }

  var data = sheet.getDataRange().getValues();
  var existingKeys = {};
  for (var i = 0; i < data.length; i++) {
    existingKeys[String(data[i][0]).trim()] = true;
  }

  var keys = Object.keys(CONFIG_DEFAULTS);
  for (var j = 0; j < keys.length; j++) {
    if (!existingKeys[keys[j]]) {
      sheet.appendRow([keys[j], CONFIG_DEFAULTS[keys[j]]]);
    }
  }
}

/**
 * Get the spreadsheet reliably, even from google.script.run (HTML dialog) contexts
 * where getActiveSpreadsheet() may return the wrong spreadsheet.
 * Uses the SPREADSHEET_ID stored in script properties by setupSystem().
 */
function _getSpreadsheet() {
  var storedId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (storedId) {
    return SpreadsheetApp.openById(storedId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}
