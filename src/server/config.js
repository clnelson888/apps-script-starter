/**
 * config.js — Reads/writes system configuration from the "Data" tab.
 *
 * The Data tab stores key-value pairs in columns A (key) and B (value).
 * All form IDs, colors, thresholds, and settings live here instead of
 * being hardcoded in the script.
 */

const DATA_SHEET_NAME = 'Data';

// Default values applied when setupSystem() creates the Data tab
const DEFAULTS = {
  YELLOW_FORM_ID:       '',
  PINK_FORM_ID:         '',
  LATE_FORM_ID:         '',
  ROSTER_FORM_ID:       '',
  COLOR_YELLOW:         '#FFD966',
  COLOR_PINK:           '#FF91A4',
  COLOR_HEADER:         '#1155CC',
  REHEARSAL_START_TIME:  '15:30',
  LATE_THRESHOLD_MIN:   '45',
  CONCERN_INCLUDE_TARDY: 'TRUE',
};

/**
 * Get a single config value by key.
 * Falls back to DEFAULTS if the key exists there but not in the sheet.
 * Returns empty string if not found anywhere.
 *
 * @param {string} key
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @returns {string}
 */
export function getConfig(key, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) return DEFAULTS[key] || '';

  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      return String(data[i][1]).trim();
    }
  }
  return DEFAULTS[key] || '';
}

/**
 * Get all config values as a plain object.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @returns {Object<string, string>}
 */
export function getAllConfig(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const result = { ...DEFAULTS };
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) return result;

  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    if (key) result[key] = String(data[i][1]).trim();
  }
  return result;
}

/**
 * Set a single config value. Creates/updates the row for that key.
 *
 * @param {string} key
 * @param {string} value
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 */
export function setConfig(key, value, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  // Key not found — append new row
  sheet.appendRow([key, value]);
}

/**
 * Ensure the Data tab exists with all default keys populated.
 * Does not overwrite existing values.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
export function ensureDataSheet(ss) {
  let sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DATA_SHEET_NAME);
    // Header row
    const hdr = sheet.getRange(1, 1, 1, 2);
    hdr.setValues([['Key', 'Value']]);
    hdr.setBackground('#1155CC')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 240);
    sheet.setColumnWidth(2, 300);
  }

  // Ensure all default keys exist
  const data = sheet.getDataRange().getValues();
  const existingKeys = new Set(data.map(row => String(row[0]).trim()));

  Object.entries(DEFAULTS).forEach(([key, defaultVal]) => {
    if (!existingKeys.has(key)) {
      sheet.appendRow([key, defaultVal]);
    }
  });
}

export { DATA_SHEET_NAME, DEFAULTS };
