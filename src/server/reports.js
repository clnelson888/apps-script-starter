/**
 * reports.js — Concern list dashboard and future reporting.
 *
 * The concern list shows all members who are not "Present" or "Excused"
 * for a given rehearsal date, grouped by section with Absent and Tardy
 * separated.
 */

import { getAllConfig } from './config.js';
import { getSectionSheets, findDateColumn } from './roster.js';

/**
 * Generate and display the concern list for today's rehearsal.
 * Groups results by section, with Absent and Tardy separated.
 */
export function doGenerateConcernList() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const tz    = ss.getSpreadsheetTimeZone();
  const cfg   = getAllConfig(ss);
  const today = Utilities.formatDate(new Date(), tz, 'M/d/yyyy');

  const includeTardy = (cfg.CONCERN_INCLUDE_TARDY || 'TRUE').toUpperCase() === 'TRUE';

  const sections = [];

  getSectionSheets(ss).forEach(sheet => {
    const dateCol = findDateColumn(sheet, today, tz);
    if (dateCol < 0) return; // No column for today — skip

    const data   = sheet.getDataRange().getValues();
    const absent = [];
    const tardy  = [];

    for (let r = 1; r < data.length; r++) {
      const name = _str(data[r][0]);
      if (!name) continue;

      const cellValue = _str(data[r][dateCol - 1]).toLowerCase();

      // Skip Present and Excused — they're accounted for
      if (cellValue === 'present' || cellValue === 'excused') continue;

      if (cellValue === 'tardy') {
        if (includeTardy) tardy.push(name);
      } else {
        // Empty cell or "Absent" — both count as absent
        absent.push(name);
      }
    }

    if (absent.length > 0 || tardy.length > 0) {
      sections.push({ name: sheet.getName(), absent, tardy });
    }
  });

  // Build HTML output
  const lines = [];
  lines.push(`<b>Attendance Concerns — ${today}</b>`);
  lines.push('<hr>');

  if (sections.length === 0) {
    lines.push('No concerns found for today.');
  } else {
    sections.forEach(({ name, absent, tardy }) => {
      lines.push(`<b>${name}</b>`);
      if (absent.length > 0) {
        lines.push(`&nbsp;&nbsp;<i>Absent:</i> ${absent.join(' &middot; ')}`);
      }
      if (tardy.length > 0) {
        lines.push(`&nbsp;&nbsp;<i>Tardy:</i> ${tardy.join(' &middot; ')}`);
      }
      lines.push('');
    });
  }

  const html = HtmlService
    .createHtmlOutput(
      `<div style="font-family:monospace;font-size:13px;line-height:1.8">${lines.join('<br>')}</div>`
    )
    .setWidth(560)
    .setHeight(480)
    .setTitle(`Attendance Concerns — ${today}`);

  SpreadsheetApp.getUi().showModalDialog(html, `Attendance Concerns — ${today}`);
}
