/**
 * db.js — Thin query layer over Google Sheets tabs.
 *
 * SheetTable wraps a single sheet, maps column headers to indices,
 * and provides object-based read/write operations so calling code
 * never deals with raw array indices.
 *
 * Usage:
 *   const table = new SheetTable(ss.getSheetByName('Database'));
 *   const members = table.where({ Section: 'Trumpet', Active: 'TRUE' });
 *   const member  = table.findOne({ 'Full Name': 'Doe, John' });
 *   table.insert({ 'First Name': 'John', 'Last Name': 'Doe', Section: 'Trumpet' });
 *   table.update(member._row, { Section: 'Trombone' });
 */

export class SheetTable {
  /**
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
   */
  constructor(sheet) {
    this.sheet = sheet;
    this._refreshHeaders();
  }

  /** Re-read the header row to pick up any structural changes. */
  _refreshHeaders() {
    if (!this.sheet || this.sheet.getLastColumn() === 0) {
      this._headers = [];
      this._headerMap = {};
      return;
    }
    this._headers = this.sheet
      .getRange(1, 1, 1, this.sheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h).trim());
    this._headerMap = {};
    this._headers.forEach((name, idx) => {
      if (name) this._headerMap[name] = idx;
    });
  }

  /** Get the list of header names. */
  get headers() {
    return this._headers;
  }

  /**
   * Get the 0-based column index for a header name.
   * Returns -1 if not found.
   */
  colIndex(headerName) {
    const idx = this._headerMap[headerName];
    return idx !== undefined ? idx : -1;
  }

  /**
   * Return all data rows as objects keyed by header name.
   * Each object has an extra `_row` property (1-based sheet row number).
   *
   * @returns {Array<Object>}
   */
  getAll() {
    const lastRow = this.sheet.getLastRow();
    if (lastRow < 2) return [];
    const lastCol = this.sheet.getLastColumn();
    if (lastCol < 1) return [];

    const data = this.sheet
      .getRange(2, 1, lastRow - 1, lastCol)
      .getValues();

    return data.map((row, idx) => {
      const obj = { _row: idx + 2 }; // 1-based row in the sheet
      this._headers.forEach((name, colIdx) => {
        if (name) obj[name] = row[colIdx];
      });
      return obj;
    });
  }

  /**
   * Return rows matching ALL key-value pairs in the filter.
   * Comparison is case-insensitive string match.
   *
   * @param {Object} filter  e.g. { Section: 'Trumpet', Active: 'TRUE' }
   * @returns {Array<Object>}
   */
  where(filter) {
    const entries = Object.entries(filter).map(([k, v]) => [k, _norm(v)]);
    return this.getAll().filter(row =>
      entries.every(([key, val]) => _norm(row[key]) === val)
    );
  }

  /**
   * Return the first row matching the filter, or null.
   *
   * @param {Object} filter
   * @returns {Object|null}
   */
  findOne(filter) {
    const entries = Object.entries(filter).map(([k, v]) => [k, _norm(v)]);
    const all = this.getAll();
    for (const row of all) {
      if (entries.every(([key, val]) => _norm(row[key]) === val)) {
        return row;
      }
    }
    return null;
  }

  /**
   * Append a new row. Keys in `obj` must match header names.
   * Missing headers are left blank.
   *
   * @param {Object} obj
   */
  insert(obj) {
    const rowData = this._headers.map(name => obj[name] !== undefined ? obj[name] : '');
    this.sheet.appendRow(rowData);
  }

  /**
   * Update specific fields in an existing row.
   *
   * @param {number} sheetRow  1-based row number (from `_row`)
   * @param {Object} updates   e.g. { Section: 'Trombone' }
   */
  update(sheetRow, updates) {
    Object.entries(updates).forEach(([key, value]) => {
      const colIdx = this.colIndex(key);
      if (colIdx < 0) return;
      this.sheet.getRange(sheetRow, colIdx + 1).setValue(value);
    });
  }

  /**
   * Get distinct values for a column, optionally filtered.
   *
   * @param {string} columnName
   * @param {Object} [filter]
   * @returns {string[]}
   */
  distinct(columnName, filter) {
    const rows = filter ? this.where(filter) : this.getAll();
    const seen = new Set();
    rows.forEach(row => {
      const val = _norm(row[columnName]);
      if (val) seen.add(String(row[columnName]).trim());
    });
    return [...seen].sort();
  }
}

/** Normalize a value for comparison: trim, lowercase, stringify. */
function _norm(val) {
  if (val === undefined || val === null) return '';
  return String(val).trim().toLowerCase();
}
