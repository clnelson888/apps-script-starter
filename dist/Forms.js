/**
 * Forms.js — Build and update the Yellow, Pink, and Late Check-In forms.
 *
 * Forms use dropdown lists populated from the Database tab so members
 * never type their name free-form.
 */

/* exported createForms, updateAllFormDropdowns, getFormUrls, parseFormResponse,
   Q_YELLOW_NAME, Q_YELLOW_ENSEMBLE, Q_YELLOW_SECTION, Q_YELLOW_DAYS,
   Q_YELLOW_START, Q_YELLOW_END, Q_YELLOW_NOTES,
   Q_PINK_NAME, Q_PINK_ENSEMBLE, Q_PINK_SECTION, Q_PINK_DATE, Q_PINK_REASON,
   Q_LATE_SECTION, Q_LATE_NAME, Q_LATE_REASON, Q_LATE_OTHER */

// ─── Form question titles ───────────────────────────────────────────────────

// Yellow Sheet (Class Conflict)
var Q_YELLOW_NAME = 'Your Full Name';
var Q_YELLOW_ENSEMBLE = 'Ensemble';
var Q_YELLOW_SECTION = 'Your Section';
var Q_YELLOW_DAYS = 'Conflict Days';
var Q_YELLOW_START = 'Conflict Start Time';
var Q_YELLOW_END = 'Conflict End Time';
var Q_YELLOW_NOTES = 'Notes';

// Pink Sheet (Excused Absence)
var Q_PINK_NAME = 'Your Full Name';
var Q_PINK_ENSEMBLE = 'Ensemble';
var Q_PINK_SECTION = 'Your Section';
var Q_PINK_DATE = 'Date of Absence';
var Q_PINK_REASON = 'Reason';

// Late Check-In
var Q_LATE_SECTION = 'What is your section?';
var Q_LATE_NAME = 'Your Name';
var Q_LATE_REASON = 'Reason for late arrival';
var Q_LATE_OTHER = 'If "Other", please explain:';

var DAYS_OF_WEEK = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
var ENSEMBLES = ['KSUMB'];

/**
 * Create all three forms from scratch. Stores new IDs in Data tab.
 * Deletes old forms referenced in Data tab first.
 */
function createForms(ss) {
  _deleteOldForm(ss, 'YELLOW_FORM_ID');
  _deleteOldForm(ss, 'PINK_FORM_ID');
  _deleteOldForm(ss, 'LATE_FORM_ID');

  var names = getActiveMemberNames(ss);
  var sections = getActiveSections(ss);

  _createYellowForm(ss, names, sections);
  _createPinkForm(ss, names, sections);
  _createLateForm(ss, sections);
}

/**
 * Refresh name/section dropdowns on all forms from the current Database.
 */
function updateAllFormDropdowns(ss) {
  var names = getActiveMemberNames(ss);
  var sections = getActiveSections(ss);

  // Yellow and Pink: update list items
  var formKeys = ['YELLOW_FORM_ID', 'PINK_FORM_ID'];
  for (var f = 0; f < formKeys.length; f++) {
    var formId = getConfig(formKeys[f], ss);
    if (!formId) continue;
    try {
      var form = FormApp.openById(formId);
      var items = form.getItems(FormApp.ItemType.LIST);
      for (var i = 0; i < items.length; i++) {
        var title = items[i].getTitle();
        if ((title === Q_YELLOW_NAME || title === Q_PINK_NAME) && names.length > 0) {
          items[i].asListItem().setChoiceValues(names);
        }
        if ((title === Q_YELLOW_SECTION || title === Q_PINK_SECTION) && sections.length > 0) {
          items[i].asListItem().setChoiceValues(sections);
        }
      }
    } catch (err) {
      Logger.log(`Failed to update ${formKeys[f]}: ${err.message}`);
    }
  }

  // Late check-in: rebuild entirely (has section routing)
  var lateId = getConfig('LATE_FORM_ID', ss);
  if (lateId) {
    try {
      var lateForm = FormApp.openById(lateId);
      _rebuildLateForm(lateForm, ss, sections);
    } catch (err) {
      Logger.log(`Failed to update late form: ${err.message}`);
    }
  }
}

/**
 * Get the published URLs for all forms.
 */
function getFormUrls(ss) {
  var result = {};
  var formKeys = {
    'Yellow Sheet (Class Conflict)': 'YELLOW_FORM_ID',
    'Pink Sheet (Excused Absence)': 'PINK_FORM_ID',
    'Late Check-In': 'LATE_FORM_ID',
  };

  var labels = Object.keys(formKeys);
  for (var i = 0; i < labels.length; i++) {
    var label = labels[i];
    var id = getConfig(formKeys[label], ss);
    if (!id) {
      result[label] = 'Not created yet — run Setup first';
    } else {
      try {
        result[label] = FormApp.openById(id).getPublishedUrl();
      } catch (_e) {
        result[label] = `Form not found (ID: ${id})`;
      }
    }
  }
  return result;
}

/**
 * Parse a FormResponse into a plain object keyed by question title.
 */
function parseFormResponse(formResponse) {
  var result = { timestamp: formResponse.getTimestamp() };
  var itemResponses = formResponse.getItemResponses();
  for (var i = 0; i < itemResponses.length; i++) {
    result[itemResponses[i].getItem().getTitle()] = itemResponses[i].getResponse();
  }
  return result;
}

// ─── Private: form building ─────────────────────────────────────────────────

function _deleteOldForm(ss, configKey) {
  var formId = getConfig(configKey, ss);
  if (!formId) return;
  try {
    DriveApp.getFileById(formId).setTrashed(true);
  } catch (_e) {
    // Form already deleted or inaccessible
  }
  setConfig(configKey, '', ss);
}

function _createYellowForm(ss, names, sections) {
  var form = FormApp.create('KSUMB Yellow Sheet — Class Conflict');
  setConfig('YELLOW_FORM_ID', form.getId(), ss);

  form.setDescription(
    'Submit this form to report a recurring class conflict with rehearsal. ' +
      'Your request will be reviewed and approved by staff.'
  );
  form.setConfirmationMessage('Your class conflict has been submitted and is pending staff approval.');
  form.setCollectEmail(true);

  var nameItem = form.addListItem();
  nameItem.setTitle(Q_YELLOW_NAME).setRequired(true);
  if (names.length > 0) nameItem.setChoiceValues(names);

  var ensembleItem = form.addListItem();
  ensembleItem.setTitle(Q_YELLOW_ENSEMBLE).setRequired(true);
  ensembleItem.setChoiceValues(ENSEMBLES);

  var sectionItem = form.addListItem();
  sectionItem.setTitle(Q_YELLOW_SECTION).setRequired(true);
  if (sections.length > 0) sectionItem.setChoiceValues(sections);

  var daysItem = form.addCheckboxItem();
  daysItem.setTitle(Q_YELLOW_DAYS).setRequired(true);
  daysItem.setChoiceValues(DAYS_OF_WEEK);

  form.addTimeItem().setTitle(Q_YELLOW_START).setRequired(true);
  form.addTimeItem().setTitle(Q_YELLOW_END).setRequired(true);

  form.addTextItem().setTitle(Q_YELLOW_NOTES).setRequired(false);
}

function _createPinkForm(ss, names, sections) {
  var form = FormApp.create('KSUMB Pink Sheet — Excused Absence');
  setConfig('PINK_FORM_ID', form.getId(), ss);

  form.setDescription(
    'Submit this form to request a single-day excused absence from rehearsal. ' +
      'Submit as early as possible. One pink sheet per rehearsal date.'
  );
  form.setConfirmationMessage('Your excused absence has been submitted.');
  form.setCollectEmail(true);

  var nameItem = form.addListItem();
  nameItem.setTitle(Q_PINK_NAME).setRequired(true);
  if (names.length > 0) nameItem.setChoiceValues(names);

  var ensembleItem = form.addListItem();
  ensembleItem.setTitle(Q_PINK_ENSEMBLE).setRequired(true);
  ensembleItem.setChoiceValues(ENSEMBLES);

  var sectionItem = form.addListItem();
  sectionItem.setTitle(Q_PINK_SECTION).setRequired(true);
  if (sections.length > 0) sectionItem.setChoiceValues(sections);

  form.addDateItem().setTitle(Q_PINK_DATE).setRequired(true);

  form.addParagraphTextItem().setTitle(Q_PINK_REASON).setRequired(true);
}

function _createLateForm(ss, sections) {
  var form = FormApp.create('KSUMB Late Check-In');
  setConfig('LATE_FORM_ID', form.getId(), ss);

  form.setDescription('Arrived late? Fill this out quickly — have your section and name ready.');
  form.setConfirmationMessage('You have been checked in. Welcome to rehearsal!');

  _rebuildLateForm(form, ss, sections);
}

function _rebuildLateForm(form, ss, sections) {
  // Clear existing items
  var existingItems = form.getItems();
  for (var i = existingItems.length - 1; i >= 0; i--) {
    form.deleteItem(existingItems[i]);
  }

  if (!sections || sections.length === 0) {
    sections = getActiveSections(ss);
  }
  if (sections.length === 0) return;

  var sectionItem = form.addMultipleChoiceItem();
  sectionItem.setTitle(Q_LATE_SECTION).setRequired(true);

  var sectionChoices = [];
  for (var s = 0; s < sections.length; s++) {
    var section = sections[s];
    var pageBreak = form.addPageBreakItem();
    pageBreak.setTitle(`${section} — Select Your Name`);

    var nameItem = form.addListItem();
    nameItem.setTitle(Q_LATE_NAME).setRequired(true);
    var sectionNames = getMembersBySection(ss, section);
    if (sectionNames.length > 0) nameItem.setChoiceValues(sectionNames);

    var reasonItem = form.addMultipleChoiceItem();
    reasonItem.setTitle(Q_LATE_REASON).setRequired(true);
    reasonItem.setChoiceValues(['Class', 'Traffic / Parking', 'Work', 'Other (explain below)']);

    form
      .addTextItem()
      .setTitle(Q_LATE_OTHER)
      .setHelpText('Only needed if you selected "Other" above.')
      .setRequired(false);

    pageBreak.setGoToPage(FormApp.PageNavigationType.SUBMIT);

    sectionChoices.push(sectionItem.createChoice(section, pageBreak));
  }

  sectionItem.setChoices(sectionChoices);
}
