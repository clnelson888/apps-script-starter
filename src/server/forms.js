/**
 * forms.js — Build and update the Yellow, Pink, and Late Check-In forms.
 *
 * Forms use dropdown lists populated from the Database tab so members
 * never type their name free-form. This prevents misspelling issues.
 */

import { getConfig, setConfig, getAllConfig } from './config.js';
import { getActiveMemberNames, getActiveSections, getMembersBySection } from './roster.js';

// ─── Form question titles (consistent between build and parse) ──────────────
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

/**
 * Create or update all three managed forms.
 * Stores form IDs in the Data tab.
 */
export function createOrUpdateForms(ss) {
  const yellowForm = _getOrCreateForm(ss, 'YELLOW_FORM_ID', 'KSUMB Yellow Sheet — Class Conflict');
  const pinkForm   = _getOrCreateForm(ss, 'PINK_FORM_ID',   'KSUMB Pink Sheet — Excused Absence');
  const lateForm   = _getOrCreateForm(ss, 'LATE_FORM_ID',   'KSUMB Late Check-In');

  _buildYellowForm(yellowForm, ss);
  _buildPinkForm(pinkForm, ss);
  _buildLateForm(lateForm, ss);
}

/**
 * Refresh name/section dropdowns on all forms from the current Database.
 */
export function updateAllFormDropdowns(ss) {
  const names    = getActiveMemberNames(ss);
  const sections = getActiveSections(ss);

  // Yellow and Pink: update list items
  ['YELLOW_FORM_ID', 'PINK_FORM_ID'].forEach(key => {
    const formId = getConfig(key, ss);
    if (!formId) return;
    try {
      const form = FormApp.openById(formId);
      const items = form.getItems(FormApp.ItemType.LIST);
      items.forEach(item => {
        const title = item.getTitle();
        if (title === Q_NAME && names.length > 0) {
          item.asListItem().setChoiceValues(names);
        }
        if (title === Q_SECTION && sections.length > 0) {
          item.asListItem().setChoiceValues(sections);
        }
      });
    } catch (err) {
      Logger.log(`Failed to update ${key}: ${err.message}`);
    }
  });

  // Late check-in has section routing — rebuild entirely
  const lateId = getConfig('LATE_FORM_ID', ss);
  if (lateId) {
    try {
      _buildLateForm(FormApp.openById(lateId), ss);
    } catch (err) {
      Logger.log(`Failed to update late form: ${err.message}`);
    }
  }
}

/**
 * Get the published URLs for all forms.
 * Returns an object with form name → URL (or status message).
 */
export function getFormUrls(ss) {
  const result = {};
  const formKeys = {
    'Yellow Sheet (Class Conflict)': 'YELLOW_FORM_ID',
    'Pink Sheet (Excused Absence)':  'PINK_FORM_ID',
    'Late Check-In':                 'LATE_FORM_ID',
  };

  Object.entries(formKeys).forEach(([label, key]) => {
    const id = getConfig(key, ss);
    if (!id) {
      result[label] = 'Not created yet — run Setup first';
    } else {
      try {
        result[label] = FormApp.openById(id).getPublishedUrl();
      } catch (_) {
        result[label] = `Form not found (ID: ${id})`;
      }
    }
  });

  return result;
}

/**
 * Parse a FormResponse event into a plain object keyed by question title.
 * Special key "timestamp" holds the submission Date.
 */
export function parseFormResponse(formResponse) {
  const result = { timestamp: formResponse.getTimestamp() };
  formResponse.getItemResponses().forEach(ir => {
    result[ir.getItem().getTitle()] = ir.getResponse();
  });
  return result;
}

// ─── Private: form building ─────────────────────────────────────────────────

function _getOrCreateForm(ss, configKey, title) {
  const storedId = getConfig(configKey, ss);
  if (storedId) {
    try {
      return FormApp.openById(storedId);
    } catch (_) {
      // Form was deleted — recreate
    }
  }
  const form = FormApp.create(title);
  setConfig(configKey, form.getId(), ss);
  return form;
}

function _buildYellowForm(form, ss) {
  form.getItems().forEach(item => form.deleteItem(item));
  form.setTitle('KSUMB Yellow Sheet — Recurring Class Conflict');
  form.setDescription(
    'Submit this form to report a recurring class conflict with rehearsal. ' +
    'Your request will be reviewed and approved by staff.'
  );
  form.setConfirmationMessage('Your class conflict has been submitted and is pending staff approval.');
  form.setCollectEmail(true);

  const names    = getActiveMemberNames(ss);
  const sections = getActiveSections(ss);

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

function _buildPinkForm(form, ss) {
  form.getItems().forEach(item => form.deleteItem(item));
  form.setTitle('KSUMB Pink Sheet — Excused Absence');
  form.setDescription(
    'Submit this form to request a single-day excused absence from rehearsal. ' +
    'Submit as early as possible. One pink sheet per rehearsal date.'
  );
  form.setConfirmationMessage('Your excused absence has been submitted.');
  form.setCollectEmail(true);

  const names    = getActiveMemberNames(ss);
  const sections = getActiveSections(ss);

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

function _buildLateForm(form, ss) {
  form.getItems().forEach(item => form.deleteItem(item));
  form.setTitle('KSUMB Late Check-In');
  form.setDescription(
    'Arrived late? Fill this out quickly — have your section and name ready.'
  );
  form.setConfirmationMessage('You have been checked in. Welcome to rehearsal!');

  const sections = getActiveSections(ss);
  if (sections.length === 0) return;

  // Page 0: section selection
  const sectionItem = form.addMultipleChoiceItem();
  sectionItem.setTitle(Q_LATE_SECTION).setRequired(true);

  const sectionChoices = [];

  sections.forEach(section => {
    const pageBreak = form.addPageBreakItem();
    pageBreak.setTitle(`${section} — Select Your Name`);

    const nameItem = form.addListItem();
    nameItem.setTitle(Q_LATE_NAME).setRequired(true);
    const sectionNames = getMembersBySection(ss, section);
    if (sectionNames.length > 0) nameItem.setChoiceValues(sectionNames);

    const reasonItem = form.addMultipleChoiceItem();
    reasonItem.setTitle(Q_LATE_REASON).setRequired(true);
    reasonItem.setChoiceValues(['Class', 'Traffic / Parking', 'Work', 'Other (explain below)']);

    form.addTextItem()
      .setTitle(Q_LATE_OTHER)
      .setHelpText('Only needed if you selected "Other" above.')
      .setRequired(false);

    pageBreak.setGoToPage(FormApp.PageNavigationType.SUBMIT);

    sectionChoices.push(sectionItem.createChoice(section, pageBreak));
  });

  sectionItem.setChoices(sectionChoices);
}

export {
  Q_NAME, Q_SECTION, Q_LATE_SECTION, Q_LATE_NAME,
  Q_LATE_REASON, Q_LATE_OTHER, Q_CONFLICT_DESC,
  Q_CONFLICT_TIMES, Q_ABSENCE_DATE, Q_REASON,
};
