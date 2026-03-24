/**
 * Menu.js — Custom menu and menu action wrappers.
 *
 * This file defines onOpen() and all menu-callable functions.
 * Each wrapper gets the active spreadsheet and delegates to
 * the core function in the appropriate module.
 */

/* exported onOpen, showFormUrls, setupSystem, syncRoster, addRehearsalDate, updateForms, generateConcernList */

/**
 * Simple trigger: builds the KSUMB Attendance custom menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('KSUMB Attendance')
    .addItem('Show Form URLs', 'showFormUrls')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu('Setup & Maintenance')
        .addItem('Run Initial Setup (first time only)', 'setupSystem')
        .addItem('Sync Roster from Database', 'syncRoster')
        .addItem('Refresh Form Dropdowns', 'updateForms')
        .addItem('Add Rehearsal Date...', 'addRehearsalDate')
    )
    .addSeparator()
    .addItem('Concern List', 'generateConcernList')
    .addToUi();
}

function showFormUrls() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var urls = getFormUrls(ss);
  var keys = Object.keys(urls);
  var lines = [];
  for (var i = 0; i < keys.length; i++) {
    lines.push(`${keys[i]}:\n${urls[keys[i]]}`);
  }
  SpreadsheetApp.getUi().alert('Form URLs', lines.join('\n\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function setupSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Store spreadsheet ID for use in google.script.run contexts (e.g. HTML dialogs)
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  ensureDataSheet(ss);
  ensureDatabaseSheet(ss);
  createForms(ss);
  installTriggers(ss);

  var urls = getFormUrls(ss);
  var keys = Object.keys(urls);
  var lines = [];
  for (var i = 0; i < keys.length; i++) {
    lines.push(`${keys[i]}:\n${urls[keys[i]]}`);
  }
  var msg = `KSUMB Attendance System setup complete!\n\nForm URLs (share these with your section):\n\n${lines.join('\n\n')}`;

  SpreadsheetApp.getUi().alert(msg);
}

function syncRoster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  doSyncRoster(ss);
  updateAllFormDropdowns(ss);
  SpreadsheetApp.getUi().alert('Roster sync complete. Form dropdowns updated.');
}

function addRehearsalDate(dateStr, timeStr, ssId) {
  if (!dateStr) {
    // Pass the spreadsheet ID into the dialog template so it can send it back
    // via google.script.run (avoids stale script-property issues)
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tpl = HtmlService.createTemplateFromFile('add-date-dialog');
    tpl.ssId = ss.getId();
    var html = tpl.evaluate().setWidth(420).setHeight(310);
    SpreadsheetApp.getUi().showModalDialog(html, 'Add Rehearsal Date');
    return;
  }
  doAddRehearsalDate(dateStr, timeStr || '', ssId);
}

function updateForms() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  updateAllFormDropdowns(ss);
  SpreadsheetApp.getUi().alert('Form dropdowns refreshed.');
}

function generateConcernList() {
  doGenerateConcernList();
}
