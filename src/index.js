/**
 * KSUMB Attendance System — Main entry point
 *
 * All named exports become global functions in Google Apps Script
 * via the viteExposeGasFunctions plugin.
 */

import { ensureDataSheet } from './server/config.js';
import { ensureDatabaseSheet, doSyncRoster, doAddRehearsalDate } from './server/roster.js';
import { createOrUpdateForms, updateAllFormDropdowns, getFormUrls } from './server/forms.js';
import { ensureLogSheets, installTriggers, onEditGlobal, processYellowSheetSubmit, processPinkSheetSubmit, processLateCheckInSubmit } from './server/triggers.js';
import { doGenerateConcernList } from './server/reports.js';

// ─── Simple trigger: builds custom menu on spreadsheet open ─────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('KSUMB Attendance')
    .addItem('Show Form URLs', 'showFormUrls')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Setup & Maintenance')
        .addItem('Run Initial Setup (first time only)', 'setupSystem')
        .addItem('Sync Roster from Database', 'syncRoster')
        .addItem('Refresh Form Dropdowns', 'updateForms')
        .addItem('Add Rehearsal Date...', 'addRehearsalDate')
    )
    .addSeparator()
    .addItem('Concern List', 'generateConcernList')
    .addToUi();
}

// ─── Menu actions ───────────────────────────────────────────────────────────

function showFormUrls() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const urls = getFormUrls(ss);
  const msg  = Object.entries(urls)
    .map(([label, url]) => `${label}:\n${url}`)
    .join('\n\n');
  SpreadsheetApp.getUi().alert('Form URLs', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureDataSheet(ss);
  ensureDatabaseSheet(ss);
  ensureLogSheets(ss);
  createOrUpdateForms(ss);
  installTriggers(ss);

  const urls = getFormUrls(ss);
  const msg = 'KSUMB Attendance System setup complete!\n\n' +
    'Form URLs (share these with your section):\n\n' +
    Object.entries(urls)
      .map(([label, url]) => `${label}:\n${url}`)
      .join('\n\n');

  SpreadsheetApp.getUi().alert(msg);
}

function syncRoster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  doSyncRoster(ss);
  updateAllFormDropdowns(ss);
  SpreadsheetApp.getUi().alert('Roster sync complete. Form dropdowns updated.');
}

function addRehearsalDate(dateStr, timeStr) {
  if (!dateStr) {
    const html = HtmlService
      .createHtmlOutputFromFile('add-date-dialog')
      .setWidth(420)
      .setHeight(310);
    SpreadsheetApp.getUi().showModalDialog(html, 'Add Rehearsal Date');
    return;
  }
  doAddRehearsalDate(dateStr, timeStr || '');
}

function updateForms() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  updateAllFormDropdowns(ss);
  SpreadsheetApp.getUi().alert('Form dropdowns refreshed.');
}

function generateConcernList() {
  doGenerateConcernList();
}

// ─── Exports (become global GAS functions) ──────────────────────────────────

export {
  // Simple trigger
  onOpen,

  // Menu actions
  showFormUrls,
  setupSystem,
  syncRoster,
  addRehearsalDate,
  updateForms,
  generateConcernList,

  // Installable trigger handlers
  onEditGlobal,
  processYellowSheetSubmit,
  processPinkSheetSubmit,
  processLateCheckInSubmit,
};
