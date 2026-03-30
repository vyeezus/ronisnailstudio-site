/**
 * --- Roni's Nail Studio: Master Google Apps Script V8 (Luxury Masterpiece) ---
 * 100% Fixed Data + Detailed Confirmed Emails + Zero Emojis + TEST PREVIEWER 💎✨
 */

const CALENDAR_ID = 'f9c38dc209bf435115238aba24b24be51b7e4e2f05f3e3f9c08b9077a78c33b3@group.calendar.google.com';
const PERSONAL_CALENDAR_ID = 'nguyenveronica0108@gmail.com'; 
const SHEET_NAME = 'Bookings';
const PENDING_COLOR = '5'; // Yellow
const MY_EMAIL = 'ronisnailstudio@gmail.com';
const SPREADSHEET_ID = '16IJ_aJlAXWrF6UpM_g4Oia8rGGyAcmeR898STRX_5tc'; 
/**
 * Web app URL — must be the /exec deployment (not /dev; /dev requires developer login).
 * After edits: Deploy → Manage deployments → Edit → New version → Deploy.
 */
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzdT_rV3dR7Th4VHeLE3uJcyTPr4bI-6uy-_Im6xz-nZ0rGPToj85zy7Is7LmpNVS0Wwg/exec';

/**
 * Approve/Decline / Confirm / Cancel links — proxied on your site (not script.google.com in the browser).
 * Change if you use a different domain.
 */
const BOOKING_ACTION_HTML_BASE = 'https://ronisnailstudio.com/api/booking';

/**
 * 2-day reminder “Reschedule” button — dedicated page on your site.
 */
const RESCHEDULE_PAGE_BASE = 'https://ronisnailstudio.com/reschedule.html';

/** Sheet columns A–K: … eventId (I), actionToken (J), reminderSent (K) for 2-day email */
function generateActionToken() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Sheet date column: full calendar date (avoid raw Date string in emails). */
function formatSheetDateForEmail(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy');
  }
  return String(value).trim();
}

/**
 * Sheet time column: often stored as a time-only Date (shows as Dec 30, 1899 when stringified).
 * Always format as h:mm a in the script timezone.
 */
function formatSheetTimeForEmail(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'h:mm a');
  }
  return String(value).trim();
}

function emailDetailRow(label, value) {
  const v = escapeHtml(value || '—');
  const l = escapeHtml(label);
  return '<tr><td style="padding:10px 0;border-bottom:1px solid #eee;vertical-align:top;width:100px;"><span style="color:#666;font-size:12px;text-transform:uppercase;letter-spacing:0.06em;">' + l + '</span></td><td style="padding:10px 0;border-bottom:1px solid #eee;color:#1a1a1a;font-size:17px;font-weight:500;">' + v + '</td></tr>';
}

/**
 * Plain HTML only — no HtmlService. Google’s HtmlService shell injects JS that sends the browser
 * to script.google.com, which breaks the Netlify proxy (address bar must stay on your domain).
 * Through the proxy, ContentService renders as a normal web page.
 */
function htmlPage(title, bodyHtml) {
  const html = '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>' + escapeHtml(title) + '</title></head><body style="font-family:system-ui,sans-serif;text-align:center;padding:2rem;line-height:1.5;">' + bodyHtml + '</body></html>';
  return ContentService.createTextOutput(html).setMimeType(ContentService.MimeType.HTML);
}

function buildBookingActionUrl(queryString) {
  const base = BOOKING_ACTION_HTML_BASE || SCRIPT_URL;
  if (base.indexOf('?') >= 0) return base + '&' + queryString;
  return base + '?' + queryString;
}

function getCRMSpreadsheet() {
  try { return SpreadsheetApp.openById(SPREADSHEET_ID); } catch (e) { return SpreadsheetApp.getActiveSpreadsheet(); }
}

/** Compare sheet E column to calendar start (yyyy-MM-dd in script TZ). */
function syncSheetDateToYyyyMmDd_(value) {
  const tz = Session.getScriptTimeZone();
  if (value instanceof Date) {
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }
  const s = String(value == null ? '' : value).trim();
  if (!s) return '';
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, tz, 'yyyy-MM-dd');
  }
  return s.split('T')[0].split(' ')[0];
}

/** Normalize time strings so sheet "2:00 PM" matches calendar formatting. */
function normalizeTimeToken_(t) {
  return String(t == null ? '' : t)
    .trim()
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, '')
    .toLowerCase();
}

/**
 * --- Calendar → Sheet sync (drag/drop in Google Calendar) ---
 * If the event start in your studio calendar no longer matches columns E–F, updates the sheet,
 * clears the 2-day reminder flag (K), and emails the client.
 *
 * Automatic runs: run installCalendarSyncTrigger() once in the editor (every 10 minutes).
 */
function syncCalendarToSpreadsheet() {
  const tz = Session.getScriptTimeZone();
  const ss = getCRMSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  for (let i = 1; i < data.length; i++) {
    const status = data[i][7];
    const eventId = data[i][8];
    if (
      (status === 'CONFIRMED' || status === 'CLIENT_CONFIRMED' || status === 'PENDING') &&
      eventId &&
      String(eventId).indexOf('pending') < 0
    ) {
      const event = calendar.getEventById(eventId);
      if (!event) {
        sheet.getRange(i + 1, 8).setValue('CANCELLED');
        const email = data[i][6];
        if (email) {
          const cn = escapeHtml(data[i][1]);
          const body =
            '<div style="font-family: sans-serif; padding: 40px; max-width: 500px; margin: auto; border: 1px solid #f0f0f0; border-radius: 12px;"><h2 style="color: #1a1a1a; text-align: center; font-weight: 500;">Appointment Cancelled</h2><p style="text-align: center;">Hello ' +
            cn +
            ", your appointment at Roni's Nail Studio has been cancelled.</p></div>";
          MailApp.sendEmail({ to: email, name: "Roni's Nail Studio", subject: "Appointment Cancelled: Roni's Nail Studio", htmlBody: body });
        }
        continue;
      }
      const sheetDateYmd = syncSheetDateToYyyyMmDd_(data[i][4]);
      const sheetTimeNorm = normalizeTimeToken_(formatSheetTimeForEmail(data[i][5]));
      const calStart = event.getStartTime();
      const calDateYmd = Utilities.formatDate(calStart, tz, 'yyyy-MM-dd');
      const calTimeNorm = normalizeTimeToken_(Utilities.formatDate(calStart, tz, 'h:mm a'));
      if (calDateYmd !== sheetDateYmd || calTimeNorm !== sheetTimeNorm) {
        const calTimeDisplay = Utilities.formatDate(calStart, tz, 'h:mm a');
        sheet.getRange(i + 1, 5).setValue(calStart);
        sheet.getRange(i + 1, 6).setValue(calTimeDisplay);
        sheet.getRange(i + 1, 11).setValue('');
        SpreadsheetApp.flush();
        const neatDate = Utilities.formatDate(calStart, tz, 'EEEE, MMMM d, yyyy');
        const clientEmail = data[i][6];
        const service = data[i][3];
        if (clientEmail) {
          const html = getCalendarUpdatedClientEmailHtml(data[i][1], neatDate, calTimeDisplay, service);
          MailApp.sendEmail({
            to: clientEmail,
            name: "Roni's Nail Studio",
            subject: "Appointment time updated — Roni's Nail Studio",
            htmlBody: html,
          });
        }
      }
    }
  }
}

/**
 * One-time setup: creates a time-driven trigger to run syncCalendarToSpreadsheet every 10 minutes.
 * Apps Script editor → select installCalendarSyncTrigger → Run. (Skip if trigger already listed under Triggers.)
 */
function installCalendarSyncTrigger() {
  const fn = 'syncCalendarToSpreadsheet';
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === fn) {
      Logger.log('Trigger already exists for ' + fn);
      return;
    }
  }
  ScriptApp.newTrigger(fn).timeBased().everyMinutes(10).create();
  Logger.log('Installed time trigger: ' + fn + ' every 10 minutes');
}

function tokenMatches(stored, provided) {
  if (stored === undefined || stored === null || provided === undefined || provided === null) return false;
  return String(stored).trim() === String(provided).trim();
}

function doGet(e) {
  const action = e.parameter.action;
  const eventId = e.parameter.eventId;
  const token = e.parameter.token;

  if (action === 'accept' || action === 'reject') {
    if (!eventId || !token) {
      return htmlPage('Invalid link', '<h2>Invalid link</h2><p>This approval link is incomplete. Open the message from your latest booking email.</p>');
    }
    const ss = getCRMSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let clientName = "Client", clientEmail = "", service = "", dateVal = "", timeStr = "";
    let rowStatus = "";
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][8]) === String(eventId)) {
          if (!tokenMatches(data[i][9], token)) {
            return htmlPage('Invalid link', '<h2>Invalid or expired link</h2><p>This link does not match the booking. Use the Approve or Decline buttons in the booking email.</p>');
          }
          rowIndex = i + 1;
          rowStatus = data[i][7];
          clientName = data[i][1];
          service = data[i][3];
          dateVal = data[i][4];
          timeStr = data[i][5];
          clientEmail = data[i][6];
          break;
        }
    }
    if (rowIndex < 0) {
      return htmlPage('Not found', '<h2>Booking not found</h2><p>No matching request. It may have been removed.</p>');
    }
    if (rowStatus !== 'PENDING') {
      return htmlPage('Already handled', '<h2>Already handled</h2><p>This request was already approved or declined.</p>');
    }
    if (action === 'accept') {
        const cal = CalendarApp.getCalendarById(CALENDAR_ID);
        const ev = cal.getEventById(eventId);
        if (ev) { 
          const nEv = cal.createEvent(clientName, ev.getStartTime(), ev.getEndTime(), { description: ev.getDescription() }); 
          ev.deleteEvent(); 
          sheet.getRange(rowIndex, 9).setValue(nEv.getId()); 
        }
        sheet.getRange(rowIndex, 8).setValue('CONFIRMED');
        
        const neatD = formatSheetDateForEmail(dateVal);
        const neatTime = formatSheetTimeForEmail(timeStr);
        const acceptEmailHtml = getConfirmedEmailHtml(clientName, neatD, neatTime, service);
        MailApp.sendEmail({ to: clientEmail, name: "Roni's Nail Studio", subject: "Appointment Confirmed: Roni's Nail Studio", htmlBody: acceptEmailHtml });
        SpreadsheetApp.flush();
        return htmlPage('Accepted', '<h2>Request accepted</h2><p>The client has been notified.</p>');
    }
    if (action === 'reject') {
       const cal = CalendarApp.getCalendarById(CALENDAR_ID); const ev = cal.getEventById(eventId); if (ev) ev.deleteEvent();
       sheet.getRange(rowIndex, 8).setValue('REJECTED');
       const neatD = formatSheetDateForEmail(dateVal);
       const neatTime = formatSheetTimeForEmail(timeStr);
       const declinedHtml = getDeclinedEmailHtml(clientName, neatD, neatTime, service);
       if (clientEmail) {
         MailApp.sendEmail({ to: clientEmail, name: "Roni's Nail Studio", subject: "Update on your booking request: Roni's Nail Studio", htmlBody: declinedHtml });
       }
       SpreadsheetApp.flush();
       return htmlPage('Declined', '<h2>Request declined</h2><p>The client has been notified.</p>');
    }
    return htmlPage('Not supported', '<h2>Unsupported action</h2>');
  }

  if (action === 'client_confirm' || action === 'client_cancel') {
    if (!eventId || !token) {
      return htmlPage('Invalid link', '<h2>Invalid link</h2><p>This link is incomplete.</p>');
    }
    const ss = getCRMSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let clientName = "Client", clientEmail = "", service = "", dateVal = "", timeStr = "";
    let rowStatus = "";
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][8]) === String(eventId)) {
        if (!tokenMatches(data[i][9], token)) {
          return htmlPage('Invalid link', '<h2>Invalid or expired link</h2><p>Use the buttons in your reminder email.</p>');
        }
        rowIndex = i + 1;
        rowStatus = data[i][7];
        clientName = data[i][1];
        service = data[i][3];
        dateVal = data[i][4];
        timeStr = data[i][5];
        clientEmail = data[i][6];
        break;
      }
    }
    if (rowIndex < 0) {
      return htmlPage('Not found', '<h2>Booking not found</h2>');
    }
    if (rowStatus !== 'CONFIRMED') {
      if (rowStatus === 'CLIENT_CONFIRMED') {
        return htmlPage('Confirmed', '<h2>Already confirmed</h2><p>We already have your confirmation. See you soon!</p>');
      }
      if (rowStatus === 'CANCELLED' || rowStatus === 'REJECTED') {
        return htmlPage('Cancelled', '<h2>No longer active</h2><p>This appointment is no longer on the schedule.</p>');
      }
      return htmlPage('Invalid', '<h2>Link not valid</h2><p>This reminder link only applies to confirmed appointments.</p>');
    }
    if (action === 'client_confirm') {
      sheet.getRange(rowIndex, 8).setValue('CLIENT_CONFIRMED');
      SpreadsheetApp.flush();
      return htmlPage('Thank you', '<h2>Thank you!</h2><p>Your appointment is confirmed. We\'ll see you soon.</p>');
    }
    if (action === 'client_cancel') {
      const cal = CalendarApp.getCalendarById(CALENDAR_ID);
      const ev = cal.getEventById(eventId);
      if (ev) ev.deleteEvent();
      sheet.getRange(rowIndex, 8).setValue('CANCELLED');
      SpreadsheetApp.flush();
      const neatD = formatSheetDateForEmail(dateVal);
      const neatTime = formatSheetTimeForEmail(timeStr);
      const ownerRows = emailDetailRow('Date', neatD) + emailDetailRow('Time', neatTime) + emailDetailRow('Service', service);
      const ownerBody = '<div style="font-family:sans-serif;padding:24px;max-width:480px;margin:auto;"><h2 style="font-weight:500;">Client cancelled (2-day reminder)</h2><p><strong>' + escapeHtml(clientName) + '</strong> cancelled their upcoming appointment.</p><table style="width:100%;border-collapse:collapse;margin-top:16px;background:#fafafa;border-radius:8px;"><tbody>' + ownerRows + '</tbody></table><p style="margin-top:16px;color:#666;">Client email: ' + escapeHtml(clientEmail || '—') + '</p></div>';
      MailApp.sendEmail({ to: MY_EMAIL, subject: 'Client cancelled appointment: ' + clientName, htmlBody: ownerBody });
      if (clientEmail) {
        const clientNote = '<div style="font-family:sans-serif;padding:32px;max-width:480px;margin:auto;"><p>Hi ' + escapeHtml(clientName) + ',</p><p>Your appointment on <strong>' + escapeHtml(neatD) + '</strong> at <strong>' + escapeHtml(neatTime) + '</strong> has been cancelled as requested.</p><p>If you\'d like to book again, you can do so on our website anytime.</p><p style="margin-top:24px;color:#999;font-size:13px;">Roni\'s Nail Studio</p></div>';
        MailApp.sendEmail({ to: clientEmail, name: "Roni's Nail Studio", subject: 'Appointment cancelled: Roni\'s Nail Studio', htmlBody: clientNote });
      }
      return htmlPage('Cancelled', '<h2>Appointment cancelled</h2><p>We\'ve sent a confirmation to your email.</p>');
    }
  }

  if (action === 'reschedule_meta') {
    const eventId = e.parameter.eventId;
    const token = e.parameter.token;
    const callback = e.parameter.callback;
    function wrapRescheduleMeta_(obj) {
      const json = JSON.stringify(obj);
      if (callback) {
        return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
    }
    if (!eventId || !token) {
      return wrapRescheduleMeta_({ ok: false, error: 'missing_params' });
    }
    const ss = getCRMSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    let found = null;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][8]) !== String(eventId)) continue;
      if (!tokenMatches(data[i][9], token)) {
        return wrapRescheduleMeta_({ ok: false, error: 'invalid_token' });
      }
      found = { row: i + 1, status: data[i][7], service: data[i][3], clientName: data[i][1], phone: data[i][2], email: data[i][6] };
      break;
    }
    if (!found) {
      return wrapRescheduleMeta_({ ok: false, error: 'not_found' });
    }
    if (found.status !== 'CONFIRMED' && found.status !== 'CLIENT_CONFIRMED') {
      return wrapRescheduleMeta_({ ok: false, error: 'not_confirmed' });
    }
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    let ev;
    try {
      ev = calendar.getEventById(eventId);
    } catch (err) {
      ev = null;
    }
    if (!ev) {
      return wrapRescheduleMeta_({ ok: false, error: 'event_missing' });
    }
    const durMs = ev.getEndTime().getTime() - ev.getStartTime().getTime();
    const durationMinutes = Math.max(30, Math.round(durMs / 60000));
    return wrapRescheduleMeta_({
      ok: true,
      service: found.service,
      clientName: found.clientName,
      phone: found.phone,
      email: found.email,
      durationMinutes: durationMinutes,
    });
  }

  if (!action) {
    const busCal = CalendarApp.getCalendarById(CALENDAR_ID);
    const perCal = CalendarApp.getCalendarById(PERSONAL_CALENDAR_ID);
    const now = new Date();
    const three = new Date();
    three.setMonth(now.getMonth() + 3);
    const ignoreEventId = e.parameter.ignoreEventId ? String(e.parameter.ignoreEventId) : '';
    const allBusy = [];
    function pushBusy_(cal) {
      if (!cal) return;
      cal.getEvents(now, three).forEach(function (v) {
        if (ignoreEventId && String(v.getId()) === ignoreEventId) return;
        allBusy.push({ start: v.getStartTime().toISOString(), end: v.getEndTime().toISOString() });
      });
    }
    pushBusy_(busCal);
    pushBusy_(perCal);
    const json = JSON.stringify(allBusy);
    return ContentService.createTextOutput(e.parameter.callback ? e.parameter.callback + '(' + json + ')' : json).setMimeType(e.parameter.callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
  }
  return htmlPage('Bad request', '<h2>Bad request</h2>');
}

function handleReschedulePost(d) {
  if (!d.eventId || !d.token || !d.date || !d.time) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'missing_fields' })).setMimeType(ContentService.MimeType.JSON);
  }
  const ss = getCRMSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  let clientName = '';
  let clientEmail = '';
  let phone = '';
  let service = '';
  let rowStatus = '';
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][8]) !== String(d.eventId)) continue;
    if (!tokenMatches(data[i][9], d.token)) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'invalid_token' })).setMimeType(ContentService.MimeType.JSON);
    }
    rowIndex = i + 1;
    clientName = data[i][1];
    phone = data[i][2];
    service = data[i][3];
    clientEmail = data[i][6];
    rowStatus = data[i][7];
    break;
  }
  if (rowIndex < 0) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'not_found' })).setMimeType(ContentService.MimeType.JSON);
  }
  if (rowStatus !== 'CONFIRMED' && rowStatus !== 'CLIENT_CONFIRMED') {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'not_confirmed' })).setMimeType(ContentService.MimeType.JSON);
  }
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  const ev = cal.getEventById(d.eventId);
  if (!ev) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'event_missing' })).setMimeType(ContentService.MimeType.JSON);
  }
  const durMs = ev.getEndTime().getTime() - ev.getStartTime().getTime();
  const start = new Date(d.date.toString().split('T')[0] + 'T' + convertTo24Hour(d.time));
  const newEnd = new Date(start.getTime() + durMs);
  ev.setTime(start, newEnd);
  ev.setTitle(clientName);
  sheet.getRange(rowIndex, 5).setValue(start);
  const neatTimeStr = Utilities.formatDate(start, Session.getScriptTimeZone(), 'h:mm a');
  sheet.getRange(rowIndex, 6).setValue(neatTimeStr);
  sheet.getRange(rowIndex, 11).setValue('');
  SpreadsheetApp.flush();
  const neatDate = Utilities.formatDate(start, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy');
  const clientHtml = getRescheduledClientEmailHtml(clientName, neatDate, neatTimeStr, service);
  if (clientEmail) {
    MailApp.sendEmail({ to: clientEmail, name: "Roni's Nail Studio", subject: "Appointment rescheduled — Roni's Nail Studio", htmlBody: clientHtml });
  }
  const ownerRows = emailDetailRow('Client', clientName) + emailDetailRow('Date', neatDate) + emailDetailRow('Time', neatTimeStr) + emailDetailRow('Service', service);
  const ownerBody = '<div style="font-family:sans-serif;padding:24px;max-width:480px;margin:auto;"><h2 style="font-weight:500;">Appointment rescheduled</h2><table style="width:100%;border-collapse:collapse;margin-top:16px;background:#fafafa;border-radius:8px;"><tbody>' + ownerRows + '</tbody></table><p style="margin-top:16px;color:#666;">' + escapeHtml(clientEmail || '') + '</p></div>';
  MailApp.sendEmail({ to: MY_EMAIL, subject: 'Rescheduled: ' + clientName, htmlBody: ownerBody });
  return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.reschedule === true) {
      return handleReschedulePost(d);
    }
    const actionToken = generateActionToken();
    const ss = getCRMSpreadsheet(); let s = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    const row = s.getLastRow() + 1;
    s.appendRow([new Date(), d.clientName, d.phone, d.service, d.date, d.time, d.email, 'PENDING', '', actionToken, '']);
    SpreadsheetApp.flush();
    const c = CalendarApp.getCalendarById(CALENDAR_ID);
    const start = new Date(d.date.toString().split('T')[0] + 'T' + convertTo24Hour(d.time));
    const ev = c.createEvent('PENDING: ' + d.clientName, start, new Date(start.getTime() + 3600000), { description: `Phone: ${d.phone}\nEmail: ${d.email}\nService: ${d.service}` });
    ev.setColor(PENDING_COLOR); s.getRange(row, 9).setValue(ev.getId());
    
    // Formatting for Owner Request
    let rawDate = new Date(d.date);
    let neatDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy");
    
    const requestHtml = getRequestEmailHtml(d.clientName, d.service, d.phone, d.email, neatDate, d.time, ev.getId(), actionToken);
    MailApp.sendEmail({ to: MY_EMAIL, subject: "New Booking Request: " + d.clientName, htmlBody: requestHtml });
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) { return ContentService.createTextOutput(JSON.stringify({ status: 'error' })); }
}

/**
 * --- BRANDED EMAIL TEMPLATES ---
 */
function getConfirmedEmailHtml(name, date, time, service) {
  const n = escapeHtml(name);
  const rows = emailDetailRow('Date', date) + emailDetailRow('Time', time) + emailDetailRow('Service', service);
  return `
    <div style="font-family: sans-serif; padding: 40px; max-width: 500px; margin: auto; border: 1px solid #f0f0f0; border-radius: 12px; background-color: #ffffff;">
      <p style="font-size: 16px; color: #1a1a1a; line-height: 1.6;">Hi ${n},</p>
      <p style="font-size: 16px; color: #1a1a1a; line-height: 1.6; margin-bottom: 20px;">Your appointment at Roni's Nail Studio is <strong>confirmed</strong>. Here are the details:</p>
      <table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:8px;margin:0 0 24px 0;" cellpadding="0" cellspacing="0" role="presentation"><tbody>
        ${rows}
      </tbody></table>
      <p style="font-size: 16px; color: #1a1a1a; line-height: 1.6;">Thank you for booking with me &lt;3</p>
      <hr style="border: none; border-top: 1px solid #f0f0f0; margin: 40px 0;">
      <p style="font-size: 13px; color: #999; text-align: center;">Roni's Nail Studio</p>
    </div>`;
}

function getRescheduledClientEmailHtml(name, date, time, service) {
  const n = escapeHtml(name);
  const rows = emailDetailRow('Date', date) + emailDetailRow('Time', time) + emailDetailRow('Service', service);
  return (
    '<div style="font-family: sans-serif; padding: 40px; max-width: 500px; margin: auto; border: 1px solid #f0f0f0; border-radius: 12px; background-color: #ffffff;">' +
    '<p style="font-size: 16px; color: #1a1a1a; line-height: 1.6;">Hi ' +
    n +
    ',</p>' +
    '<p style="font-size: 16px; color: #1a1a1a; line-height: 1.6; margin-bottom: 20px;">Your appointment has been <strong>rescheduled</strong>. Here are your updated details:</p>' +
    '<table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:8px;margin:0 0 24px 0;" cellpadding="0" cellspacing="0" role="presentation"><tbody>' +
    rows +
    '</tbody></table>' +
    '<p style="font-size: 16px; color: #1a1a1a; line-height: 1.6;">We look forward to seeing you.</p>' +
    '<hr style="border: none; border-top: 1px solid #f0f0f0; margin: 40px 0;">' +
    '<p style="font-size: 13px; color: #999; text-align: center;">Roni\'s Nail Studio</p>' +
    '</div>'
  );
}

/** Client email when you move the event in Google Calendar (syncCalendarToSpreadsheet). */
function getCalendarUpdatedClientEmailHtml(name, date, time, service) {
  const n = escapeHtml(name);
  const rows = emailDetailRow('Date', date) + emailDetailRow('Time', time) + emailDetailRow('Service', service);
  return (
    '<div style="font-family: sans-serif; padding: 40px; max-width: 500px; margin: auto; border: 1px solid #f0f0f0; border-radius: 12px; background-color: #ffffff;">' +
    '<p style="font-size: 16px; color: #1a1a1a; line-height: 1.6;">Hi ' +
    n +
    ',</p>' +
    '<p style="font-size: 16px; color: #1a1a1a; line-height: 1.6; margin-bottom: 20px;">Your appointment time has been <strong>updated</strong>. Here are your current details:</p>' +
    '<table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:8px;margin:0 0 24px 0;" cellpadding="0" cellspacing="0" role="presentation"><tbody>' +
    rows +
    '</tbody></table>' +
    '<p style="font-size: 15px; color: #555; line-height: 1.6;">If this doesn\'t work for you, reply to this email or use the reschedule link in your reminder when you receive it.</p>' +
    '<hr style="border: none; border-top: 1px solid #f0f0f0; margin: 40px 0;">' +
    '<p style="font-size: 13px; color: #999; text-align: center;">Roni\'s Nail Studio</p>' +
    '</div>'
  );
}

function getDeclinedEmailHtml(name, date, time, service) {
  const n = escapeHtml(name);
  const rows = emailDetailRow('Date', date) + emailDetailRow('Time', time) + emailDetailRow('Service', service);
  return `
    <div style="font-family: sans-serif; padding: 40px; max-width: 500px; margin: auto; border: 1px solid #f0f0f0; border-radius: 12px; background-color: #ffffff;">
      <p style="font-size: 16px; color: #1a1a1a; line-height: 1.6;">Hi ${n},</p>
      <p style="font-size: 16px; color: #1a1a1a; line-height: 1.6; margin-bottom: 20px;">Thank you for your interest in booking with Roni's Nail Studio. Unfortunately, we're unable to accommodate this request. Details below:</p>
      <table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:8px;margin:0 0 24px 0;" cellpadding="0" cellspacing="0" role="presentation"><tbody>
        ${rows}
      </tbody></table>
      <p style="font-size: 16px; color: #1a1a1a; line-height: 1.6;">You're welcome to book another time on our website when it works for you.</p>
      <p style="font-size: 16px; color: #1a1a1a; line-height: 1.6; margin-top: 16px;">If you have any questions, please feel free to reach out to me directly.</p>
      <hr style="border: none; border-top: 1px solid #f0f0f0; margin: 40px 0;">
      <p style="font-size: 13px; color: #999; text-align: center;">Roni's Nail Studio</p>
    </div>`;
}

function buildRescheduleUrl(eventId, token) {
  if (!RESCHEDULE_PAGE_BASE) return '';
  const qs = 'eventId=' + encodeURIComponent(eventId) + '&token=' + encodeURIComponent(token);
  return RESCHEDULE_PAGE_BASE.indexOf('?') >= 0 ? RESCHEDULE_PAGE_BASE + '&' + qs : RESCHEDULE_PAGE_BASE + '?' + qs;
}

function getTwoDayReminderEmailHtml(name, date, time, service, eventId, token) {
  const n = escapeHtml(name);
  const qConfirm = 'action=client_confirm&eventId=' + encodeURIComponent(eventId) + '&token=' + encodeURIComponent(token);
  const qCancel = 'action=client_cancel&eventId=' + encodeURIComponent(eventId) + '&token=' + encodeURIComponent(token);
  const urlConfirm = buildBookingActionUrl(qConfirm);
  const urlCancel = buildBookingActionUrl(qCancel);
  const urlReschedule = buildRescheduleUrl(eventId, token);
  const rescheduleBlock = urlReschedule
    ? '<a href="' + urlReschedule + '" style="background-color:#fff;color:#111;border:1px solid #111;padding:12px;text-decoration:none;border-radius:8px;font-weight:500;display:block;margin-bottom:12px;">Reschedule</a>'
    : '<p style="font-size:14px;color:#666;">To reschedule, visit our website and book a new appointment time.</p>';
  return `
    <div style="font-family: sans-serif; padding: 32px; max-width: 450px; margin: auto; border: 1px solid #eaeaea; border-radius: 12px;">
      <h2 style="color: #111; font-weight: 500; font-size: 20px; text-align: center;">Appointment in 2 days</h2>
      <p style="font-size: 16px; color: #1a1a1a; line-height: 1.6;">Hi ${n},</p>
      <p style="font-size: 16px; color: #1a1a1a; line-height: 1.6;">This is a friendly reminder about your upcoming visit. Please confirm, reschedule, or cancel below.</p>
      <table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:8px;margin:20px 0;" cellpadding="0" cellspacing="0" role="presentation"><tbody>
        ${emailDetailRow('Date', date) + emailDetailRow('Time', time) + emailDetailRow('Service', service)}
      </tbody></table>
      <div style="text-align: center;">
        <a href="${urlConfirm}" style="background-color: #111; color: white; padding: 14px; text-decoration: none; border-radius: 8px; font-weight: 500; display: block; margin-bottom: 12px;">Confirm</a>
        ${rescheduleBlock}
        <a href="${urlCancel}" style="background-color: #fff; color: #dc3545; border: 1px solid #dc3545; padding: 12px; text-decoration: none; border-radius: 8px; display: block;">Cancel</a>
      </div>
      <p style="font-size: 13px; color: #999; text-align: center; margin-top: 24px;">Roni's Nail Studio</p>
    </div>`;
}

/**
 * Sends the 2-day reminder once per booking (column K = SENT).
 * In Apps Script: Triggers → Add trigger → choose this function → time-driven → daily.
 */
function sendTwoDayReminders() {
  const ss = getCRMSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const today = new Date();
  const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  for (let i = 1; i < data.length; i++) {
    const status = data[i][7];
    if (status !== 'CONFIRMED') continue;
    if (data[i][10] === 'SENT') continue;
    const eventId = data[i][8];
    const tok = data[i][9];
    if (!eventId || !tok) continue;
    let ev;
    try {
      ev = calendar.getEventById(eventId);
    } catch (err) {
      continue;
    }
    if (!ev) continue;
    const start = ev.getStartTime();
    const apptDay = new Date(start.getFullYear(), start.getMonth(), start.getDate());
    const daysUntil = Math.round((apptDay - todayStart) / 86400000);
    if (daysUntil !== 2) continue;

    const clientName = data[i][1];
    const service = data[i][3];
    const dateVal = data[i][4];
    const timeStr = data[i][5];
    const clientEmail = data[i][6];
    if (!clientEmail) continue;

    const neatD = formatSheetDateForEmail(dateVal);
    const neatTime = formatSheetTimeForEmail(timeStr);
    const html = getTwoDayReminderEmailHtml(clientName, neatD, neatTime, service, eventId, tok);
    MailApp.sendEmail({
      to: clientEmail,
      name: "Roni's Nail Studio",
      subject: "Reminder: your appointment in 2 days — Roni's Nail Studio",
      htmlBody: html,
    });
    sheet.getRange(i + 1, 11).setValue('SENT');
    SpreadsheetApp.flush();
  }
}

function getRequestEmailHtml(name, service, phone, email, date, time, eventId, actionToken) {
  const q = 'action=accept&eventId=' + encodeURIComponent(eventId) + '&token=' + encodeURIComponent(actionToken);
  const qr = 'action=reject&eventId=' + encodeURIComponent(eventId) + '&token=' + encodeURIComponent(actionToken);
  const acc = buildBookingActionUrl(q);
  const rej = buildBookingActionUrl(qr);
  return `
    <div style="font-family: sans-serif; padding: 32px; max-width: 450px; margin: auto; border: 1px solid #eaeaea; border-radius: 12px;">
      <h2 style="color: #111; font-weight: 500; font-size: 20px; text-align: center;">New Booking Request</h2>
      <div style="background: #fafafa; padding: 20px; border-radius: 8px; margin: 20px 0; font-size: 15px; line-height: 1.6;">
        <strong>Client:</strong> ${name}<br>
        <strong>Service:</strong> ${service}<br>
        <strong>Phone:</strong> ${phone}<br>
        <strong>Email:</strong> ${email}<br>
        <strong>Date:</strong> ${date}<br>
        <strong>Time:</strong> ${time}
      </div>
      <div style="text-align: center;">
        <a href="${acc}" style="background-color: #111; color: white; padding: 14px; text-decoration: none; border-radius: 8px; font-weight: 500; display: block; margin-bottom: 12px;">Approve Request</a>
        <a href="${rej}" style="background-color: #fff; color: #dc3545; border: 1px solid #dc3545; padding: 12px; text-decoration: none; border-radius: 8px; display: block;">Decline Request</a>
      </div>
    </div>`;
}

/**
 * --- TEST / PREVIEW EMAILS ---
 * In Apps Script: select the function → Run. Authorize MailApp if prompted.
 */

/**
 * Layout only: Confirm / Cancel / Reschedule links use fake IDs (clicks will show invalid/not found).
 */
function testTwoDayReminderPreview() {
  const html = getTwoDayReminderEmailHtml(
    'Vy (Preview)',
    'Monday, April 7, 2026',
    '2:00 PM',
    'Gel Manicure',
    'PREVIEW_NO_REAL_EVENT',
    'preview_invalid_token'
  );
  MailApp.sendEmail({
    to: MY_EMAIL,
    subject: "PREVIEW: 2-day reminder (links are dummy — safe to click)",
    htmlBody:
      '<p style="font-family:sans-serif;font-size:13px;color:#666;">Dummy links only. Use <strong>testTwoDayReminderLiveToStudio</strong> to test working buttons on a real booking.</p>' +
      html,
  });
  Logger.log('Sent 2-day reminder PREVIEW to ' + MY_EMAIL);
}

/**
 * Sends the real 2-day reminder email to MY_EMAIL only, using the latest CONFIRMED row
 * that has eventId + token. Links work like production. WARNING: Cancel removes the calendar event
 * and sets the row to CANCELLED; Confirm sets CLIENT_CONFIRMED. Does not set reminder column K.
 */
function testTwoDayReminderLiveToStudio() {
  const ss = getCRMSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][7] !== 'CONFIRMED') continue;
    const eventId = data[i][8];
    const tok = data[i][9];
    if (!eventId || !tok) continue;
    const clientName = data[i][1];
    const service = data[i][3];
    const neatD = formatSheetDateForEmail(data[i][4]);
    const neatTime = formatSheetTimeForEmail(data[i][5]);
    const html = getTwoDayReminderEmailHtml(clientName, neatD, neatTime, service, eventId, tok);
    MailApp.sendEmail({
      to: MY_EMAIL,
      name: "Roni's Nail Studio",
      subject: 'TEST (live links): 2-day reminder — ' + clientName + ' — Cancel will cancel for real',
      htmlBody:
        '<p style="font-family:sans-serif;font-size:14px;color:#b45309;background:#fffbeb;padding:12px;border-radius:8px;border:1px solid #fcd34d;"><strong>Test email.</strong> Sent to studio only. Buttons use real booking data — <strong>Cancel</strong> will cancel this appointment.</p>' +
        html,
    });
    Logger.log('Sent LIVE test 2-day reminder for sheet row ' + (i + 1) + ' to ' + MY_EMAIL);
    return;
  }
  throw new Error('No CONFIRMED row with eventId and token. Approve a booking first.');
}

/**
 * Run in the editor: request + confirmed + 2-day preview (dummy links for 2-day).
 */
function testEmailPreview() {
  const mockName = "Vy Nguyen (Test)";
  const mockService = "Structured Gel New Set + Tier 2 Art";
  const mockDate = "Monday, April 6, 2026";
  const mockTime = "11:30 AM";
  const mockPhone = "555-0199";
  const mockEmail = "test@example.com";

  const reqHtml = getRequestEmailHtml(mockName, mockService, mockPhone, mockEmail, mockDate, mockTime, "test_event_id", "preview_only_invalid_token");
  MailApp.sendEmail({ to: MY_EMAIL, subject: "PREVIEW: New Booking Request", htmlBody: reqHtml });

  const confHtml = getConfirmedEmailHtml(mockName, mockDate, mockTime, mockService);
  MailApp.sendEmail({ to: MY_EMAIL, subject: "PREVIEW: Appointment Confirmed", htmlBody: confHtml });

  const twoDayHtml = getTwoDayReminderEmailHtml(mockName, mockDate, mockTime, mockService, "PREVIEW_NO_REAL_EVENT", "preview_invalid_token");
  MailApp.sendEmail({
    to: MY_EMAIL,
    subject: "PREVIEW: 2-day reminder (dummy links)",
    htmlBody: twoDayHtml,
  });

  Logger.log("Sent 3 preview emails to: " + MY_EMAIL);
}

function convertTo24Hour(timeStr) {
  const [time, modifier] = timeStr.split(' '); let [hours, minutes] = time.split(':');
  if (hours === '12') hours = '00';
  if (modifier === 'PM' || modifier === 'pm') hours = parseInt(hours, 10) + 12;
  return `${hours.toString().padStart(2, '0')}:${minutes}:00`;
}
