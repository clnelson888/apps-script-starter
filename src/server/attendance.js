/**
 * KSUMB Attendance System
 * ──────────────────────────────────────────────────────────────────────────
 * One Google Spreadsheet acts as the central hub:
 *   • "Database"      — member roster (source of truth)
 *   • "Yellow Sheets" — recurring class conflict log (staff approves)
 *   • "Pink Sheets"   — single excused absence log (auto-processed)
 *   • "Late Check-Ins"— late arrival log (auto-processed)
 *   • [Section] tabs  — per-section attendance grids
 *
 * Three Google Forms are created/managed by this script:
 *   • Yellow Sheet Form  — class conflict submission
 *   • Pink Sheet Form    — excused absence submission
 *   • Late Check-In Form — fast check-in with section → name routing
 *
 * Trigger architecture:
 *   • onEditGlobal           → spreadsheet installable onEdit
 *   • processYellowSheetSubmit → yellow form installable onFormSubmit
 *   • processPinkSheetSubmit   → pink form installable onFormSubmit
 *   • processLateCheckInSubmit → late check-in form installable onFormSubmit
 */

// ─── Sheet names ──────────────────────────────────────────────────────────────
const SHEET_DATABASE = 'Database';
const SHEET_YELLOW   = 'Yellow Sheets';
const SHEET_PINK     = 'Pink Sheets';
const SHEET_LATE     = 'Late Check-Ins';

// Non-section sheets that should never be treated as a section roster
const SYSTEM_SHEETS  = [SHEET_DATABASE, SHEET_YELLOW, SHEET_PINK, SHEET_LATE];

// ─── Colors ───────────────────────────────────────────────────────────────────
const COLOR_YELLOW  = '#FFD966'; // approved class conflict highlight (name row)
const COLOR_PINK    = '#FF91A4'; // excused absence highlight (date cell)
const COLOR_HEADER  = '#1155CC'; // column/row header fill
const COLOR_WHITE   = '#FFFFFF';

// ─── Script property keys (store created form IDs) ────────────────────────────
const PROP_YELLOW_FORM_ID = 'KSUMB_YELLOW_FORM_ID';
const PROP_PINK_FORM_ID   = 'KSUMB_PINK_FORM_ID';
const PROP_LATE_FORM_ID   = 'KSUMB_LATE_FORM_ID';

// ─── Form question titles (must be consistent between build/parse) ────────────
const Q_NAME           = 'Your Full Name';
const Q_SECTION        = 'Your Section';
const Q_LATE_SECTION   = 'What is your section?';
const Q_LATE_NAME      = 'Your Name';
const Q_LATE_REASON    = 'Reason for late arrival';
const Q_LATE_OTHER     = 'If "Other", please explain:';
const Q_CONFLICT_DESC  = 'Conflict description (class name, course number, etc.)';
const Q_CONFLICT_TIMES = 'Days and times of conflict (e.g., "MWF 3:30–4:20pm")';
const Q_ABSENCE_DATE   = 'Date of Absence';
const Q_REASON         = 'Reason';

// ─────────────────────────────────────────────────────────────────────────────
//  PUBLIC FUNCTIONS (exported via src/index.js → global scope in GAS)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Simple onOpen trigger — adds the KSUMB menu to the spreadsheet UI.
 * This is a simple (non-installable) trigger so it runs automatically
 * whenever the spreadsheet is opened.
 */
export function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎺 KSUMB Attendance')
    .addItem('📋 Show Form URLs', 'showFormUrls')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('⚙️ Setup & Maintenance')
        .addItem('Run Initial Setup (first time only)', 'setupSystem')
        .addItem('Sync Roster from Database', 'syncRoster')
        .addItem('Refresh Form Dropdowns', 'updateForms')
        .addItem('Add Rehearsal Date…', 'addRehearsalDate')
    )
    .addSeparator()
    .addItem('📊 Generate Conflict Report', 'generateConflictReport')
    .addToUi();
}

/**
 * Displays the published URLs for all three managed forms.
 * Useful after setup or when sharing links with section leaders.
 */
export function showFormUrls() {
  const props = PropertiesService.getScriptProperties();
  const yId   = props.getProperty(PROP_YELLOW_FORM_ID);
  const pId   = props.getProperty(PROP_PINK_FORM_ID);
  const lId   = props.getProperty(PROP_LATE_FORM_ID);

  const fmt = (id, label) => {
    if (!id) return `${label}: Not created yet — run Setup first`;
    try {
      return `${label}:\n${FormApp.openById(id).getPublishedUrl()}`;
    } catch (_) {
      return `${label}: Form not found (ID: ${id})`;
    }
  };

  const msg = [
    fmt(yId, 'Yellow Sheet (Class Conflict)'),
    '',
    fmt(pId, 'Pink Sheet (Excused Absence)'),
    '',
    fmt(lId, 'Late Check-In'),
  ].join('\n');

  SpreadsheetApp.getUi().alert('Form URLs', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * ONE-TIME SETUP — run manually from the Apps Script menu or editor.
 * Creates system sheets, builds/updates all three forms, and installs triggers.
 */
export function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _ensureSystemSheets(ss);
  _createOrUpdateForms(ss);
  _installTriggers(ss);

  const props  = PropertiesService.getScriptProperties();
  const yId    = props.getProperty(PROP_YELLOW_FORM_ID);
  const pId    = props.getProperty(PROP_PINK_FORM_ID);
  const lId    = props.getProperty(PROP_LATE_FORM_ID);

  const yUrl = yId ? FormApp.openById(yId).getPublishedUrl() : 'not created';
  const pUrl = pId ? FormApp.openById(pId).getPublishedUrl() : 'not created';
  const lUrl = lId ? FormApp.openById(lId).getPublishedUrl() : 'not created';

  const msg =
    'KSUMB Attendance System setup complete!\n\n' +
    'Form URLs (share these with your section):\n\n' +
    `Yellow Sheet (Class Conflict):\n${yUrl}\n\n` +
    `Pink Sheet (Excused Absence):\n${pUrl}\n\n` +
    `Late Check-In:\n${lUrl}`;

  SpreadsheetApp.getUi().alert(msg);
}

/**
 * Syncs the Database tab to per-section roster sheets.
 * • Creates a sheet for each section if it does not exist.
 * • Adds new members; marks removed members with strikethrough.
 * • Preserves existing date columns, highlights, and notes.
 * • Refreshes name dropdowns on all three forms.
 */
export function syncRoster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const dbData = _getDbData(ss);
  if (!dbData) return;

  const { headers, rows } = dbData;
  const colFirst   = headers.indexOf('First Name');
  const colLast    = headers.indexOf('Last Name');
  const colSection = headers.indexOf('Section');

  if (colFirst < 0 || colLast < 0 || colSection < 0) {
    _alert('Database is missing required columns (First Name, Last Name, Section).');
    return;
  }

  // Group members by section
  const sectionsMap = {};
  rows.forEach(row => {
    const section = _str(row[colSection]);
    const first   = _str(row[colFirst]);
    const last    = _str(row[colLast]);
    if (!section || (!first && !last)) return;
    if (!sectionsMap[section]) sectionsMap[section] = [];
    sectionsMap[section].push({ first, last });
  });

  Object.entries(sectionsMap).forEach(([section, members]) => {
    _syncSectionSheet(ss, section, members);
  });

  // Rebuild form dropdowns to reflect current roster
  _updateAllFormDropdowns(ss);

  Logger.log('Roster sync complete.');
}

/**
 * When called with no arguments (from the menu), shows a date/time picker
 * dialog. When called with arguments (from the dialog via google.script.run),
 * adds the date column directly to all section sheets.
 *
 * @param {string} [dateStr] ISO date string "YYYY-MM-DD" from the dialog.
 * @param {string} [timeStr] Optional time label, e.g. "3:30 PM".
 */
export function addRehearsalDate(dateStr, timeStr) {
  if (!dateStr) {
    const html = HtmlService
      .createHtmlOutputFromFile('add-date-dialog')
      .setWidth(420)
      .setHeight(310);
    SpreadsheetApp.getUi().showModalDialog(html, 'Add Rehearsal Date');
    return;
  }
  _doAddRehearsalDate(dateStr, timeStr || '');
}

/** Internal: actually adds the date column to every section sheet. */
function _doAddRehearsalDate(dateStr, timeStr) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const tz  = ss.getSpreadsheetTimeZone();

  // Parse date components to avoid UTC-midnight timezone issues
  const [y, m, d] = dateStr.split('-').map(Number);
  const date = new Date(y, m - 1, d, 12, 0, 0);
  const fmt  = Utilities.formatDate(date, tz, 'M/d/yyyy');

  _getSectionSheets(ss).forEach(sheet => {
    if (_findDateColumn(sheet, fmt, tz) > 0) return; // already exists

    const newCol = sheet.getLastColumn() + 1;
    const cell   = sheet.getRange(1, newCol);
    cell.setValue(date);
    cell.setNumberFormat('M/d');
    _styleHeaderCell(cell);
    sheet.setColumnWidth(newCol, 55);
    if (timeStr) cell.setNote(`Rehearsal time: ${timeStr}`);
  });

  Logger.log(`Rehearsal date ${fmt} added.`);
  _alert(`✓ ${fmt}${timeStr ? ' @ ' + timeStr : ''} added to all section sheets.`);
}

/**
 * Refreshes name and section dropdowns on all forms from the current Database.
 * Run after roster changes if setupSystem() has already been called.
 */
export function updateForms() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _updateAllFormDropdowns(ss);
}

/**
 * Generates a "List of Concerns" for the director — members who appear absent
 * for today without an approved conflict (yellow) or excused absence (pink).
 * Displays a modal dialog and logs the output.
 */
export function generateConflictReport() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const tz   = ss.getSpreadsheetTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'M/d/yyyy');

  const lines = [];
  lines.push(`<b>Attendance Concerns — ${today}</b>`);
  lines.push('<hr>');

  let anyMissing = false;

  _getSectionSheets(ss).forEach(sheet => {
    const data     = sheet.getDataRange().getValues();
    const headers  = data[0];
    const dateCol  = _findDateColumn(sheet, today, tz); // 1-indexed; -1 if not found

    const missing = [];

    for (let r = 1; r < data.length; r++) {
      // Name column is a single "Last, First" cell
      const nameCell = _str(data[r][0]);
      if (!nameCell) continue;

      // Skip if name row is yellow → approved class conflict
      const nameBg = sheet.getRange(r + 1, 1).getBackground().toLowerCase();
      if (nameBg === COLOR_YELLOW.toLowerCase()) continue;

      if (dateCol > 0) {
        const cell   = sheet.getRange(r + 1, dateCol);
        const cellBg = cell.getBackground().toLowerCase();
        const note   = cell.getNote();
        // dateCol is 1-indexed; data is 0-indexed
        const val    = _str(data[r][dateCol - 1]);

        // Pink = excused
        if (cellBg === COLOR_PINK.toLowerCase()) continue;
        // Has a late check-in note
        if (note && note.startsWith('Late arrival')) continue;
        // Cell marked present (non-empty value set by section leader)
        if (val) continue;
      }

      missing.push(`&nbsp;&nbsp;${nameCell}`);
    }

    if (missing.length > 0) {
      anyMissing = true;
      lines.push(`<b>${sheet.getName()}</b>`);
      missing.forEach(n => lines.push(n));
      lines.push('');
    }
  });

  if (!anyMissing) {
    lines.push('No unexcused absences found for today.');
  }

  const html = HtmlService
    .createHtmlOutput(`<div style="font-family:monospace;font-size:13px;line-height:1.6">${lines.join('<br>')}</div>`)
    .setWidth(520)
    .setHeight(450)
    .setTitle(`Attendance Concerns — ${today}`);

  SpreadsheetApp.getUi().showModalDialog(html, `Attendance Concerns — ${today}`);
}

// ─────────────────────────────────────────────────────────────────────────────
//  TRIGGER HANDLERS (must be global — exported via index.js)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Spreadsheet installable onEdit trigger.
 * Watches the Yellow Sheets tab for Status → "Approved" and highlights the
 * member's name row in their section sheet.
 */
export function onEditGlobal(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_YELLOW) return;

  const newVal = _str(e.value).toLowerCase();
  if (newVal !== 'approved') return;

  const data       = sheet.getDataRange().getValues();
  const headers    = data[0];
  const statusCol  = headers.indexOf('Status') + 1;
  const col        = e.range.getColumn();
  if (col !== statusCol) return;

  const row            = e.range.getRow();
  const rowData        = data[row - 1];
  const nameCol        = headers.indexOf('Full Name');
  const sectionCol     = headers.indexOf('Section');
  const conflictCol    = headers.indexOf('Conflict Days/Times');

  const name          = nameCol >= 0    ? _str(rowData[nameCol])     : '';
  const section       = sectionCol >= 0 ? _str(rowData[sectionCol])  : '';
  const conflictTimes = conflictCol >= 0 ? _str(rowData[conflictCol]) : '';

  if (!name || !section) return;

  const ss           = e.source;
  const sectionSheet = ss.getSheetByName(section);
  if (!sectionSheet) return;

  const memberRow = _findMemberRow(sectionSheet, name);
  if (memberRow < 0) return;

  const nameCell = sectionSheet.getRange(memberRow, 1);
  nameCell.setBackground(COLOR_YELLOW);
  const noteText = conflictTimes
    ? `Approved Class Conflict: ${conflictTimes}`
    : 'Approved class conflict on file';
  nameCell.setNote(noteText);
}

/**
 * Yellow Sheet form installable onFormSubmit trigger.
 * Logs the submission to the Yellow Sheets tab (Status = "Pending").
 * Staff must manually change Status to "Approved" — onEditGlobal handles that.
 */
export function processYellowSheetSubmit(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_YELLOW);
  if (!sheet) return;

  const resp          = _parseFormResponse(e.response);
  const timestamp     = resp.timestamp;
  const name          = resp[Q_NAME]           || '';
  const section       = resp[Q_SECTION]        || '';
  const conflictDesc  = resp[Q_CONFLICT_DESC]  || '';
  const conflictTimes = resp[Q_CONFLICT_TIMES] || '';

  sheet.appendRow([timestamp, name, section, conflictDesc, conflictTimes, 'Pending', '']);
}

/**
 * Pink Sheet form installable onFormSubmit trigger.
 * Logs the submission to the Pink Sheets tab and immediately applies the
 * pink excused-absence highlight to the appropriate date cell.
 */
export function processPinkSheetSubmit(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PINK);
  if (!sheet) return;

  const resp        = _parseFormResponse(e.response);
  const tz          = ss.getSpreadsheetTimeZone();
  const timestamp   = resp.timestamp;
  const name        = resp[Q_NAME]         || '';
  const section     = resp[Q_SECTION]      || '';
  const absenceDate = resp[Q_ABSENCE_DATE]; // Date object or string
  const reason      = resp[Q_REASON]       || '';

  const absenceDateStr = absenceDate instanceof Date
    ? Utilities.formatDate(absenceDate, tz, 'M/d/yyyy')
    : _str(absenceDate);

  sheet.appendRow([timestamp, name, section, absenceDateStr, reason]);

  if (name && section && absenceDateStr) {
    _applyDateCellMark(ss, name, section, absenceDateStr, COLOR_PINK, `Excused Absence: ${reason}`);
  }
}

/**
 * Late Check-In form installable onFormSubmit trigger.
 * Logs the submission to the Late Check-Ins tab and adds an arrival-time
 * note to today's date cell in the member's section sheet.
 */
export function processLateCheckInSubmit(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LATE);
  if (!sheet) return;

  const resp        = _parseFormResponse(e.response);
  const tz          = ss.getSpreadsheetTimeZone();
  const timestamp   = resp.timestamp;                // Date object
  const section     = resp[Q_LATE_SECTION] || '';
  const name        = resp[Q_LATE_NAME]    || '';
  const reason      = resp[Q_LATE_REASON]  || '';
  const reasonOther = resp[Q_LATE_OTHER]   || '';
  const fullReason  = reason === 'Other (explain below)' && reasonOther
    ? `Other: ${reasonOther}`
    : reason;
  const arrivalTime = Utilities.formatDate(timestamp, tz, 'h:mm a');
  const today       = Utilities.formatDate(timestamp, tz, 'M/d/yyyy');

  sheet.appendRow([timestamp, name, section, arrivalTime, fullReason]);

  if (name && section) {
    _applyDateCellMark(
      ss, name, section, today,
      null,
      `Late arrival: ${arrivalTime}. Reason: ${fullReason}`
    );
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  PRIVATE HELPERS
// ─────────────────────────────────────────────────────────────────────────────

/** Ensure all four system sheets exist with correct headers and frozen row. */
function _ensureSystemSheets(ss) {
  _ensureSheet(ss, SHEET_DATABASE, [
    'First Name', 'Last Name', 'School Email', 'Section', 'Section Leader', 'Instrument'
  ]);
  _ensureSheet(ss, SHEET_YELLOW, [
    'Timestamp', 'Full Name', 'Section', 'Conflict Description', 'Conflict Days/Times', 'Status', 'Notes'
  ]);
  _ensureSheet(ss, SHEET_PINK, [
    'Timestamp', 'Full Name', 'Section', 'Absence Date', 'Reason'
  ]);
  _ensureSheet(ss, SHEET_LATE, [
    'Timestamp', 'Full Name', 'Section', 'Arrival Time', 'Reason'
  ]);

  // Add a Status dropdown to the Yellow Sheets tab
  const yellowSheet = ss.getSheetByName(SHEET_YELLOW);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Approved', 'Denied'], true)
    .setAllowInvalid(false)
    .build();
  // Apply to the entire Status column (F) below the header
  yellowSheet.getRange(2, 6, yellowSheet.getMaxRows() - 1, 1).setDataValidation(statusRule);
}

/** Create a sheet if it does not exist; style the header row. */
function _ensureSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  // Style header
  const hdrRange = sheet.getRange(1, 1, 1, headers.length);
  hdrRange.setBackground(COLOR_HEADER)
    .setFontColor(COLOR_WHITE)
    .setFontWeight('bold');
  return sheet;
}

/**
 * Sync one section's members into its roster sheet.
 *
 * • Names are stored as a single "Last, First" cell in column A.
 * • Members no longer in the database have their rows deleted.
 * • All date columns (column B onward) are deleted on every sync so the
 *   sheet starts fresh — add dates back with "Add Rehearsal Date…".
 * • New members are appended; existing members are left in place.
 */
function _syncSectionSheet(ss, section, members) {
  let sheet = ss.getSheetByName(section);
  if (!sheet) {
    sheet = ss.insertSheet(section);
    const hdr = sheet.getRange(1, 1);
    hdr.setValue('Name');
    _styleHeaderCell(hdr);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 200);
  }

  // ── 1. Delete all date columns (column B onward) ──────────────────────────
  const lastCol = sheet.getLastColumn();
  if (lastCol >= 2) {
    sheet.deleteColumns(2, lastCol - 1);
  }

  // ── 2. Build map of existing members: "Last, First" → 1-based row ─────────
  const existing = sheet.getDataRange().getValues();
  const existingMap = new Map(); // normalized key → row number
  for (let r = 1; r < existing.length; r++) {
    const cell = _str(existing[r][0]);
    if (cell) existingMap.set(_normalizeName(cell), r + 1);
  }

  // Database set (normalized "first last" keys)
  const dbSet = new Set(members.map(m => _normalizeName(`${m.last}, ${m.first}`)));

  // ── 3. Delete rows for members no longer in the database ─────────────────
  // Collect row numbers first, then delete bottom-up to preserve indices.
  const rowsToDelete = [];
  existingMap.forEach((rowNum, key) => {
    if (!dbSet.has(key)) rowsToDelete.push(rowNum);
  });
  rowsToDelete.sort((a, b) => b - a); // descending
  rowsToDelete.forEach(rowNum => sheet.deleteRow(rowNum));

  // ── 4. Refresh existing map after deletions ───────────────────────────────
  const refreshed = sheet.getDataRange().getValues();
  const refreshedMap = new Map();
  for (let r = 1; r < refreshed.length; r++) {
    const cell = _str(refreshed[r][0]);
    if (cell) refreshedMap.set(_normalizeName(cell), r + 1);
  }

  // ── 5. Append new members ─────────────────────────────────────────────────
  let nextRow = refreshed.length + 1;
  members.forEach(({ first, last }) => {
    const display = `${last}, ${first}`;
    const key     = _normalizeName(display);
    if (!refreshedMap.has(key)) {
      sheet.getRange(nextRow, 1).setValue(display);
      refreshedMap.set(key, nextRow);
      nextRow++;
    }
  });
}

/** Install all needed triggers, removing stale duplicates first. */
function _installTriggers(ss) {
  const handlerNames = [
    'onEditGlobal',
    'processYellowSheetSubmit',
    'processPinkSheetSubmit',
    'processLateCheckInSubmit',
  ];

  // Remove any existing triggers for our handlers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlerNames.includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Spreadsheet onEdit (for yellow sheet approval)
  ScriptApp.newTrigger('onEditGlobal')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // Form-level onFormSubmit triggers
  const props = PropertiesService.getScriptProperties();

  const yId = props.getProperty(PROP_YELLOW_FORM_ID);
  const pId = props.getProperty(PROP_PINK_FORM_ID);
  const lId = props.getProperty(PROP_LATE_FORM_ID);

  if (yId) {
    ScriptApp.newTrigger('processYellowSheetSubmit')
      .forForm(FormApp.openById(yId))
      .onFormSubmit()
      .create();
  }
  if (pId) {
    ScriptApp.newTrigger('processPinkSheetSubmit')
      .forForm(FormApp.openById(pId))
      .onFormSubmit()
      .create();
  }
  if (lId) {
    ScriptApp.newTrigger('processLateCheckInSubmit')
      .forForm(FormApp.openById(lId))
      .onFormSubmit()
      .create();
  }
}

/** Create or open all three managed forms, then build their questions. */
function _createOrUpdateForms(ss) {
  const props = PropertiesService.getScriptProperties();

  const yellowForm = _getOrCreateForm(props, PROP_YELLOW_FORM_ID, 'KSUMB Yellow Sheet — Class Conflict');
  const pinkForm   = _getOrCreateForm(props, PROP_PINK_FORM_ID,   'KSUMB Pink Sheet — Excused Absence');
  const lateForm   = _getOrCreateForm(props, PROP_LATE_FORM_ID,   'KSUMB Late Check-In');

  _buildYellowForm(yellowForm, ss);
  _buildPinkForm(pinkForm, ss);
  _buildLateForm(lateForm, ss);
}

/** Open an existing form by stored ID, or create a new one. */
function _getOrCreateForm(props, propKey, title) {
  const storedId = props.getProperty(propKey);
  if (storedId) {
    try {
      return FormApp.openById(storedId);
    } catch (_) {
      // Form was deleted — fall through to recreate
    }
  }
  const form = FormApp.create(title);
  props.setProperty(propKey, form.getId());
  return form;
}

/** Build / rebuild the Yellow Sheet (class conflict) form. */
function _buildYellowForm(form, ss) {
  form.getItems().forEach(item => form.deleteItem(item));
  form.setTitle('KSUMB Yellow Sheet — Recurring Class Conflict');
  form.setDescription(
    'Submit this form to report a recurring class conflict with rehearsal. ' +
    'Your request will be reviewed and approved by staff. ' +
    'You will continue to be marked absent until approved.'
  );
  form.setConfirmationMessage('Your class conflict has been submitted and is pending staff approval.');
  form.setCollectEmail(true);

  const names    = _getAllMemberNames(ss);
  const sections = _getAllSections(ss);

  const nameItem = form.addListItem();
  nameItem.setTitle(Q_NAME).setRequired(true);
  if (names.length > 0) nameItem.setChoiceValues(names);

  const sectionItem = form.addListItem();
  sectionItem.setTitle(Q_SECTION).setRequired(true);
  if (sections.length > 0) sectionItem.setChoiceValues(sections);

  form.addTextItem()
    .setTitle(Q_CONFLICT_DESC)
    .setRequired(true);

  form.addTextItem()
    .setTitle(Q_CONFLICT_TIMES)
    .setHelpText('Example: MWF 3:30–4:20pm')
    .setRequired(true);
}

/** Build / rebuild the Pink Sheet (excused absence) form. */
function _buildPinkForm(form, ss) {
  form.getItems().forEach(item => form.deleteItem(item));
  form.setTitle('KSUMB Pink Sheet — Excused Absence');
  form.setDescription(
    'Submit this form to request a single-day excused absence from rehearsal. ' +
    'Submit as early as possible. One pink sheet per rehearsal date.'
  );
  form.setConfirmationMessage('Your excused absence has been submitted.');
  form.setCollectEmail(true);

  const names    = _getAllMemberNames(ss);
  const sections = _getAllSections(ss);

  const nameItem = form.addListItem();
  nameItem.setTitle(Q_NAME).setRequired(true);
  if (names.length > 0) nameItem.setChoiceValues(names);

  const sectionItem = form.addListItem();
  sectionItem.setTitle(Q_SECTION).setRequired(true);
  if (sections.length > 0) sectionItem.setChoiceValues(sections);

  form.addDateItem()
    .setTitle(Q_ABSENCE_DATE)
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle(Q_REASON)
    .setRequired(true);
}

/**
 * Build / rebuild the Late Check-In form with section-based page routing.
 * Page 0: section selection (routes to a section-specific page)
 * Pages 1…N: one page per section with name dropdown + reason
 *
 * Because this form is rebuilt whenever the roster changes, section leaders
 * can always find their members' names.
 */
function _buildLateForm(form, ss) {
  form.getItems().forEach(item => form.deleteItem(item));
  form.setTitle('KSUMB Late Check-In');
  form.setDescription(
    'Arrived after 3:30pm? Fill this out quickly — have your section and name ready.'
  );
  form.setConfirmationMessage('You have been checked in. Welcome to rehearsal!');

  const dbData = _getDbData(ss);
  if (!dbData) return; // No members yet — build an empty form shell

  const { headers, rows } = dbData;
  const colFirst   = headers.indexOf('First Name');
  const colLast    = headers.indexOf('Last Name');
  const colSection = headers.indexOf('Section');

  // Build section → sorted names map
  const sectionsMap = {};
  rows.forEach(row => {
    const section = _str(row[colSection]);
    const first   = _str(row[colFirst]);
    const last    = _str(row[colLast]);
    if (!section || (!first && !last)) return;
    if (!sectionsMap[section]) sectionsMap[section] = [];
    sectionsMap[section].push(`${first} ${last}`);
  });

  const sections = Object.keys(sectionsMap).sort();
  if (sections.length === 0) return;

  // ── Page 0: section selection ─────────────────────────────────────────────
  const sectionItem = form.addMultipleChoiceItem();
  sectionItem.setTitle(Q_LATE_SECTION).setRequired(true);

  // ── One page per section ──────────────────────────────────────────────────
  const sectionChoices = [];

  sections.forEach(section => {
    const pageBreak = form.addPageBreakItem();
    pageBreak.setTitle(`${section} — Select Your Name`);

    const nameItem = form.addListItem();
    nameItem.setTitle(Q_LATE_NAME).setRequired(true);
    const sortedNames = (sectionsMap[section] || []).sort();
    if (sortedNames.length > 0) nameItem.setChoiceValues(sortedNames);

    const reasonItem = form.addMultipleChoiceItem();
    reasonItem.setTitle(Q_LATE_REASON).setRequired(true);
    reasonItem.setChoiceValues(['Class', 'Traffic / Parking', 'Work', 'Other (explain below)']);

    form.addTextItem()
      .setTitle(Q_LATE_OTHER)
      .setHelpText('Only needed if you selected "Other" above.')
      .setRequired(false);

    // Route from section page back to submit (go to end of form)
    pageBreak.setGoToPage(FormApp.PageNavigationType.SUBMIT);

    sectionChoices.push(
      sectionItem.createChoice(section, pageBreak)
    );
  });

  sectionItem.setChoices(sectionChoices);
}

/**
 * Refresh name and section lists on all three forms after a roster change.
 * Only rebuilds if forms have been created.
 */
function _updateAllFormDropdowns(ss) {
  const props   = PropertiesService.getScriptProperties();
  const yId     = props.getProperty(PROP_YELLOW_FORM_ID);
  const pId     = props.getProperty(PROP_PINK_FORM_ID);
  const lId     = props.getProperty(PROP_LATE_FORM_ID);

  const names    = _getAllMemberNames(ss);
  const sections = _getAllSections(ss);

  // Helper: update the first ListItem named Q_NAME
  function updateNameDropdown(form) {
    const items = form.getItems(FormApp.ItemType.LIST);
    items.forEach(item => {
      if (item.getTitle() === Q_NAME && names.length > 0) {
        item.asListItem().setChoiceValues(names);
      }
      if (item.getTitle() === Q_SECTION && sections.length > 0) {
        item.asListItem().setChoiceValues(sections);
      }
    });
  }

  try {
    if (yId) updateNameDropdown(FormApp.openById(yId));
    if (pId) updateNameDropdown(FormApp.openById(pId));
  } catch (err) {
    Logger.log(`updateAllFormDropdowns error: ${err.message}`);
  }

  // Late check-in has section routing — needs a full rebuild
  try {
    if (lId) _buildLateForm(FormApp.openById(lId), ss);
  } catch (err) {
    Logger.log(`updateLateCheckInForm error: ${err.message}`);
  }
}

// ─── Sheet / data utilities ───────────────────────────────────────────────────

function _getDbData(ss) {
  const dbSheet = ss.getSheetByName(SHEET_DATABASE);
  if (!dbSheet) {
    Logger.log('Database sheet not found.');
    return null;
  }
  const data = dbSheet.getDataRange().getValues();
  if (data.length < 2) return null;
  return { headers: data[0], rows: data.slice(1) };
}

function _getAllMemberNames(ss) {
  const dbData = _getDbData(ss);
  if (!dbData) return [];
  const { headers, rows } = dbData;
  const colFirst = headers.indexOf('First Name');
  const colLast  = headers.indexOf('Last Name');
  if (colFirst < 0 || colLast < 0) return [];
  return rows
    .map(r => `${_str(r[colFirst])} ${_str(r[colLast])}`.trim())
    .filter(n => n)
    .sort();
}

function _getAllSections(ss) {
  const dbData = _getDbData(ss);
  if (!dbData) return [];
  const { headers, rows } = dbData;
  const colSection = headers.indexOf('Section');
  if (colSection < 0) return [];
  const sections = new Set();
  rows.forEach(r => { const s = _str(r[colSection]); if (s) sections.add(s); });
  return [...sections].sort();
}

function _getSectionSheets(ss) {
  return ss.getSheets().filter(s => !SYSTEM_SHEETS.includes(s.getName()));
}

/**
 * Find the 1-based column number for a date in a sheet's header row.
 * Returns -1 if not found.
 */
function _findDateColumn(sheet, dateStr, tz) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  // Column A (index 0) is the Name column; dates start at index 1 (column B)
  for (let c = 1; c < headers.length; c++) {
    if (!headers[c]) continue;
    const d = new Date(headers[c]);
    if (!isNaN(d.getTime())) {
      const formatted = Utilities.formatDate(d, tz, 'M/d/yyyy');
      if (formatted === dateStr) return c + 1; // return 1-indexed column number
    }
  }
  return -1;
}

/**
 * Find (or create) a date column in a section sheet.
 * Returns the 1-based column number.
 */
function _getOrCreateDateColumn(sheet, dateStr, tz) {
  const existing = _findDateColumn(sheet, dateStr, tz);
  if (existing > 0) return existing;

  const newCol = sheet.getLastColumn() + 1;
  const cell   = sheet.getRange(1, newCol);
  const d      = new Date(dateStr);
  cell.setValue(d);
  cell.setNumberFormat('M/d');
  _styleHeaderCell(cell);
  sheet.setColumnWidth(newCol, 55);
  return newCol;
}

/**
 * Find a member's 1-based row in a section sheet.
 * Column A stores "Last, First". Input can be "First Last" or "Last, First"
 * — both are normalized to "first last" for comparison.
 * Returns -1 if not found.
 */
function _findMemberRow(sheet, name) {
  const data   = sheet.getDataRange().getValues();
  const target = _normalizeName(name);
  for (let r = 1; r < data.length; r++) {
    if (_normalizeName(_str(data[r][0])) === target) return r + 1;
  }
  return -1;
}

/**
 * Normalize a name to "first last" (lowercase, no comma) for fuzzy matching.
 * Handles both "Last, First" and "First Last" input formats.
 */
function _normalizeName(name) {
  const s = name.toLowerCase().trim();
  const commaIdx = s.indexOf(',');
  if (commaIdx > 0) {
    // "last, first" → "first last"
    return `${s.slice(commaIdx + 1).trim()} ${s.slice(0, commaIdx).trim()}`;
  }
  return s; // already "first last"
}

/**
 * Apply a background color and/or note to a member's date cell.
 * Creates the date column if it does not yet exist.
 */
function _applyDateCellMark(ss, name, section, dateStr, color, note) {
  const sectionSheet = ss.getSheetByName(section);
  if (!sectionSheet) {
    Logger.log(`Section sheet not found: "${section}"`);
    return;
  }

  const memberRow = _findMemberRow(sectionSheet, name);
  if (memberRow < 0) {
    Logger.log(`Member not found: "${name}" in "${section}"`);
    return;
  }

  const tz      = ss.getSpreadsheetTimeZone();
  const dateCol = _getOrCreateDateColumn(sectionSheet, dateStr, tz);
  const cell    = sectionSheet.getRange(memberRow, dateCol);

  if (color) cell.setBackground(color);
  if (note) {
    const existing = cell.getNote();
    cell.setNote(existing ? `${existing}\n${note}` : note);
  }
}

// ─── Form response parsing ────────────────────────────────────────────────────

/**
 * Convert a FormResponse into a plain object keyed by question title.
 * The special key "timestamp" holds the submission Date.
 */
function _parseFormResponse(formResponse) {
  const result    = { timestamp: formResponse.getTimestamp() };
  formResponse.getItemResponses().forEach(ir => {
    const title    = ir.getItem().getTitle();
    const response = ir.getResponse();
    // DateItem responses are Date objects; everything else is a string
    result[title]  = response;
  });
  return result;
}

// ─── Style helpers ────────────────────────────────────────────────────────────

function _styleHeaderCell(cell) {
  cell.setBackground(COLOR_HEADER)
    .setFontColor(COLOR_WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

// ─── String utility ───────────────────────────────────────────────────────────

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
