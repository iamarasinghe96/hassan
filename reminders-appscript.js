/**
 * Lion Gutter Cleaning — Client Reminder System
 * Google Apps Script — paste into Extensions > Apps Script in your Google Sheet
 *
 * SHEET COLUMNS (Sheet name: "Clients")
 * A  Client Name
 * B  Phone
 * C  Email
 * D  Address
 * E  Services Done
 * F  Job Date           (format: DD/MM/YYYY)
 * G  Next Reminder Date (format: DD/MM/YYYY — auto-filled by script)
 * H  Reminder Interval  (months — default 12)
 * I  Notes
 * J  Status             (Active / Reminded / Booked / Skip)
 * K  Last Contacted     (auto-filled by script)
 */

const SHEET_NAME = 'Clients';
const LEADS_SHEET_NAME = 'Leads';
const HASSAN_EMAIL = 'iamarasinghe96@gmail.com';
const DAYS_BEFORE_DUE = 7; // send reminder this many days before due date

/**
 * Receives form submissions from the website and logs them to the Leads sheet.
 * Deploy this script as a Web App (Execute as: Me, Who has access: Anyone).
 *
 * LEADS SHEET COLUMNS:
 * A  Submitted At
 * B  Name
 * C  Phone
 * D  Email
 * E  Address
 * F  Services
 * G  Storeys
 * H  Notes
 * I  Status  (New / Contacted / Booked / No Answer)
 */
function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LEADS_SHEET_NAME);

  // Create Leads sheet with headers if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(LEADS_SHEET_NAME);
    sheet.appendRow(['Submitted At', 'Name', 'Phone', 'Email', 'Address', 'Services', 'Storeys', 'Notes', 'Status']);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  const p = e.parameter;
  sheet.appendRow([
    new Date(),
    p.name    || '',
    p.phone   || '',
    p.email   || '',
    p.address || '',
    p.services || '',
    p.storeys  || '',
    p.notes    || '',
    'New'
  ]);

  // Also notify Hassan by email
  MailApp.sendEmail({
    to: HASSAN_EMAIL,
    subject: '🦁 New quote request — ' + (p.name || 'Unknown'),
    htmlBody:
      '<h2 style="color:#1a2e4a;font-family:sans-serif;">New Quote Request</h2>' +
      '<div style="font-family:sans-serif;border:1px solid #e0e0e0;border-radius:10px;padding:16px 20px;">' +
      '<p><b>Name:</b> ' + (p.name || '—') + '</p>' +
      '<p><b>Phone:</b> ' + (p.phone || '—') + '</p>' +
      '<p><b>Email:</b> ' + (p.email || '—') + '</p>' +
      '<p><b>Address:</b> ' + (p.address || '—') + '</p>' +
      '<p><b>Services:</b> ' + (p.services || '—') + '</p>' +
      '<p><b>Storeys:</b> ' + (p.storeys || '—') + '</p>' +
      '<p><b>Notes:</b> ' + (p.notes || '—') + '</p>' +
      '<div style="margin-top:16px;">' +
      '<a href="https://wa.me/61' + (p.phone || '').replace(/\D/g,'').replace(/^0/,'') + '" style="background:#25D366;color:white;padding:10px 18px;border-radius:8px;text-decoration:none;font-weight:600;font-size:14px;">WhatsApp ' + (p.name ? p.name.split(' ')[0] : '') + '</a>' +
      '</div></div>'
  });

  return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Run this once to install the daily trigger.
 * Go to Extensions > Apps Script > Run > setupDailyTrigger
 */
function setupDailyTrigger() {
  // Remove any existing triggers first
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('checkReminders')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  Logger.log('Daily reminder trigger set for 8am every day.');
}

/**
 * Auto-fill Next Reminder Date when Job Date or Interval is entered.
 * Triggered automatically on edit.
 */
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return; // skip header

  // If Job Date (col F=6) or Interval (col H=8) changed, recalculate Next Reminder Date
  if (col === 6 || col === 8) {
    fillNextReminderDate(sheet, row);
  }
}

function fillNextReminderDate(sheet, row) {
  const jobDateVal = sheet.getRange(row, 6).getValue();
  const intervalVal = sheet.getRange(row, 8).getValue();

  if (!jobDateVal) return;

  const jobDate = new Date(jobDateVal);
  const months = intervalVal ? parseInt(intervalVal) : 12;

  const nextDate = new Date(jobDate);
  nextDate.setMonth(nextDate.getMonth() + months);

  sheet.getRange(row, 7).setValue(nextDate);
  sheet.getRange(row, 7).setNumberFormat('DD/MM/YYYY');
}

/**
 * Main function — checks all clients and emails Hassan when reminders are due.
 * Runs automatically each morning via the daily trigger.
 */
function checkReminders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log('Sheet "' + SHEET_NAME + '" not found.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const dueClients = [];

  data.forEach(function(row, i) {
    const sheetRow = i + 2;
    const name       = row[0];
    const phone      = row[1];
    const email      = row[2];
    const address    = row[3];
    const services   = row[4];
    const jobDate    = row[5];
    const nextRemind = row[6];
    const notes      = row[8];
    const status     = (row[9] || '').toString().trim().toLowerCase();

    if (!name || !nextRemind) return;
    if (status === 'skip' || status === 'booked') return;

    const reminderDate = new Date(nextRemind);
    reminderDate.setHours(0, 0, 0, 0);

    const daysUntilDue = Math.round((reminderDate - today) / (1000 * 60 * 60 * 24));

    if (daysUntilDue <= DAYS_BEFORE_DUE && daysUntilDue >= 0) {
      dueClients.push({
        name, phone, email, address, services,
        jobDate: jobDate ? Utilities.formatDate(new Date(jobDate), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '—',
        reminderDate: Utilities.formatDate(reminderDate, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
        daysUntilDue,
        notes,
        sheetRow
      });
    }
  });

  if (dueClients.length === 0) {
    Logger.log('No reminders due today.');
    return;
  }

  sendReminderEmail(dueClients);

  // Update status to "Reminded" and log last contacted date
  dueClients.forEach(function(c) {
    sheet.getRange(c.sheetRow, 10).setValue('Reminded');
    sheet.getRange(c.sheetRow, 11).setValue(new Date());
    sheet.getRange(c.sheetRow, 11).setNumberFormat('DD/MM/YYYY');
  });

  Logger.log('Reminder email sent for ' + dueClients.length + ' client(s).');
}

function sendReminderEmail(clients) {
  const subject = '🦁 Lion Gutter — ' + clients.length + ' client reminder(s) due';

  let html = '<h2 style="color:#1a2e4a;font-family:sans-serif;">Client Reminders Due</h2>';
  html += '<p style="font-family:sans-serif;color:#555;">The following clients are due for a follow-up within the next ' + DAYS_BEFORE_DUE + ' days:</p>';

  clients.forEach(function(c) {
    const waLink = 'https://wa.me/61' + c.phone.replace(/\D/g, '').replace(/^0/, '') +
      '?text=' + encodeURIComponent(
        "Hi " + c.name.split(' ')[0] + "! It's Hassan from Lion Gutter Cleaning. " +
        "Just checking in — it's been about a year since we last cleaned your gutters at " + c.address + ". " +
        "Would you like to book in for another clean? Happy to give you a free quote."
      );

    html += '<div style="border:1px solid #e0e0e0;border-radius:10px;padding:16px 20px;margin:16px 0;font-family:sans-serif;">';
    html += '<h3 style="margin:0 0 8px;color:#1a2e4a;">' + c.name + '</h3>';
    html += '<p style="margin:2px 0;color:#555;"><b>Phone:</b> ' + c.phone + '</p>';
    if (c.email) html += '<p style="margin:2px 0;color:#555;"><b>Email:</b> ' + c.email + '</p>';
    html += '<p style="margin:2px 0;color:#555;"><b>Address:</b> ' + c.address + '</p>';
    html += '<p style="margin:2px 0;color:#555;"><b>Services done:</b> ' + (c.services || '—') + '</p>';
    html += '<p style="margin:2px 0;color:#555;"><b>Last job:</b> ' + c.jobDate + '</p>';
    html += '<p style="margin:2px 0;color:#555;"><b>Reminder due:</b> ' + c.reminderDate + ' (' + (c.daysUntilDue === 0 ? 'today' : 'in ' + c.daysUntilDue + ' day(s)') + ')</p>';
    if (c.notes) html += '<p style="margin:6px 0 0;color:#888;font-size:13px;"><b>Notes:</b> ' + c.notes + '</p>';
    html += '<div style="margin-top:14px;">';
    html += '<a href="' + waLink + '" style="background:#25D366;color:white;padding:10px 18px;border-radius:8px;text-decoration:none;font-weight:600;font-size:14px;margin-right:10px;">WhatsApp ' + c.name.split(' ')[0] + '</a>';
    if (c.email) {
      const mailLink = 'mailto:' + c.email + '?subject=' + encodeURIComponent('Time for your annual gutter clean?') +
        '&body=' + encodeURIComponent("Hi " + c.name.split(' ')[0] + ",\n\nIt's Hassan from Lion Gutter Cleaning. It's been about a year since we last serviced your gutters at " + c.address + ".\n\nWould you like to book in for another clean? Happy to pop around for a free assessment.\n\nCheers,\nHassan\n0423 540 616");
      html += '<a href="' + mailLink + '" style="background:#1a2e4a;color:white;padding:10px 18px;border-radius:8px;text-decoration:none;font-weight:600;font-size:14px;">Email ' + c.name.split(' ')[0] + '</a>';
    }
    html += '</div></div>';
  });

  html += '<p style="font-family:sans-serif;color:#aaa;font-size:12px;margin-top:24px;">Update the Status column to "Booked" or "Skip" in the sheet to stop future reminders for a client.</p>';

  MailApp.sendEmail({
    to: HASSAN_EMAIL,
    subject: subject,
    htmlBody: html
  });
}

/**
 * Utility: Run this manually to backfill Next Reminder Dates for all existing rows.
 */
function backfillAllDates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  for (let row = 2; row <= lastRow; row++) {
    fillNextReminderDate(sheet, row);
  }
  Logger.log('Backfill complete for ' + (lastRow - 1) + ' rows.');
}
