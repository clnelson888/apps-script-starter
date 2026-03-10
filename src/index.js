import {
  onOpen,
  showFormUrls,
  setupSystem,
  syncRoster,
  addRehearsalDate,
  updateForms,
  generateConflictReport,
  onEditGlobal,
  processYellowSheetSubmit,
  processPinkSheetSubmit,
  processLateCheckInSubmit,
} from './server/attendance.js';

export {
  // ── Simple triggers (auto-run by Apps Script) ──────────────────────────────
  onOpen,         // Builds the custom menu on spreadsheet open

  // ── Menu actions ───────────────────────────────────────────────────────────
  showFormUrls,         // Display published form URLs
  setupSystem,          // Create sheets, forms, and triggers
  syncRoster,           // Sync Database → section sheets + refresh form dropdowns
  addRehearsalDate,     // Add today's date column to all section sheets
  updateForms,          // Manually refresh form name/section dropdowns
  generateConflictReport, // Show absence list for today

  // ── Installable trigger handlers ───────────────────────────────────────────
  onEditGlobal,             // onEdit: yellow sheet approval
  processYellowSheetSubmit, // onFormSubmit: yellow sheet form
  processPinkSheetSubmit,   // onFormSubmit: pink sheet form
  processLateCheckInSubmit, // onFormSubmit: late check-in form
};
