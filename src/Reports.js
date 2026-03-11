/**
 * Reports.js — Concern list and attendance reporting.
 *
 * The concern list shows all members who are not "Present" or "Excused"
 * for today's rehearsal, grouped by section.
 */

/* exported doGenerateConcernList */

/**
 * Generate and display the concern list for today's rehearsal.
 */
function doGenerateConcernList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var cfg = getAllConfig(ss);
  var today = Utilities.formatDate(new Date(), tz, 'M/d/yyyy');

  var includeTardy = (cfg.CONCERN_INCLUDE_TARDY || 'TRUE').toUpperCase() === 'TRUE';

  var sections = [];
  var sectionSheets = getSectionSheets(ss);

  for (var i = 0; i < sectionSheets.length; i++) {
    var sheet = sectionSheets[i];
    var dateCol = findDateColumn(sheet, today, tz);
    if (dateCol < 0) continue;

    var data = sheet.getDataRange().getValues();
    var absent = [];
    var tardy = [];

    for (var r = 1; r < data.length; r++) {
      var name = _str(data[r][0]);
      if (!name) continue;

      var cellValue = _str(data[r][dateCol - 1]).toLowerCase();

      if (cellValue === 'present' || cellValue === 'excused') continue;

      if (cellValue === 'tardy') {
        if (includeTardy) tardy.push(name);
      } else {
        absent.push(name);
      }
    }

    if (absent.length > 0 || tardy.length > 0) {
      sections.push({ name: sheet.getName(), absent, tardy });
    }
  }

  // Build HTML output
  var lines = [];
  lines.push(`<b>Attendance Concerns \u2014 ${today}</b>`);
  lines.push('<hr>');

  if (sections.length === 0) {
    lines.push('No concerns found for today.');
  } else {
    for (var s = 0; s < sections.length; s++) {
      var sec = sections[s];
      lines.push(`<b>${sec.name}</b>`);
      if (sec.absent.length > 0) {
        lines.push(`&nbsp;&nbsp;<i>Absent:</i> ${sec.absent.join(' &middot; ')}`);
      }
      if (sec.tardy.length > 0) {
        lines.push(`&nbsp;&nbsp;<i>Tardy:</i> ${sec.tardy.join(' &middot; ')}`);
      }
      lines.push('');
    }
  }

  var html = HtmlService.createHtmlOutput(
    `<div style="font-family:monospace;font-size:13px;line-height:1.8">${lines.join('<br>')}</div>`
  )
    .setWidth(560)
    .setHeight(480)
    .setTitle(`Attendance Concerns \u2014 ${today}`);

  SpreadsheetApp.getUi().showModalDialog(html, `Attendance Concerns \u2014 ${today}`);
}
