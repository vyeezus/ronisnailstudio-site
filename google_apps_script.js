/**
 * --- Roni's Nail Studio: Master Google Apps Script V8 (Luxury Masterpiece) ---
 * 100% Fixed Data + Detailed Confirmed Emails + Zero Emojis + TEST PREVIEWER 💎✨
 */

const CALENDAR_ID = 'f9c38dc209bf435115238aba24b24be51b7e4e2f05f3e3f9c08b9077a78c33b3@group.calendar.google.com';
const PERSONAL_CALENDAR_ID = 'nguyenveronica0108@gmail.com'; 
const SHEET_NAME = 'Bookings';
/** Optional tab: columns day (0–6 Sun–Sat), start, end (integers). If missing, booking uses script defaults. */
const HOURS_SHEET_NAME = 'StudioHours';
const PENDING_COLOR = '5'; // Yellow
/**
 * Prepended to calendar titles only after the *client* confirms (2-day reminder Confirm, or accepting an alternate time).
 * Plain Unicode dingbat (U+2727), not an emoji codepoint — usually renders as a small four-point sparkle.
 */
const CLIENT_CONFIRMED_CAL_PREFIX = '\u2727 '; // ✧

function clientConfirmedCalendarEventTitle_(clientName) {
  const n = String(clientName == null ? '' : clientName).trim() || 'Client';
  return CLIENT_CONFIRMED_CAL_PREFIX + n;
}

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

/** Owner proposes alternate time (from booking-request email). */
const OWNER_MODIFY_PAGE_BASE = 'https://ronisnailstudio.com/owner-modify-request.html';

/** Owner declines a request with an optional note (from booking-request email). */
const OWNER_REJECT_PAGE_BASE = 'https://ronisnailstudio.com/reject-booking.html';

/** Column P: optional owner decline reason (only set when status REJECTED via decline form). */
const SHEET_COL_OWNER_DECLINE_REASON = 16;

/** Sheet: A–N as before; O = client notes; L/M/N = proposed date, time, mod client token; P = owner decline reason (if any). */
function generateActionToken() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Optional booking notes: trim, cap length, strip control chars (column O + calendar description). */
function sanitizeClientNotes_(raw) {
  let t = String(raw == null ? '' : raw).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  t = t.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, '');
  t = t.trim();
  if (t.length > 2000) t = t.substring(0, 2000);
  return t;
}

/** Owner decline note → sheet column P + client email (trim, cap length, strip controls). */
function sanitizeOwnerDeclineReason_(raw) {
  let t = String(raw == null ? '' : raw).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  t = t.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, '');
  t = t.trim();
  if (t.length > 1500) t = t.substring(0, 1500);
  return t;
}

/** Sheet date column: full calendar date (avoid raw Date string in emails). */
function formatSheetDateForEmail(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy');
  }
  const s = String(value).trim().split('T')[0].split(' ')[0];
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const noon = parseYmdAndTimeLocal_(s, '12:00:00');
    if (!isNaN(noon.getTime())) {
      return Utilities.formatDate(noon, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy');
    }
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

/**
 * Service text is comma-separated (public booking) or middle-dot-separated (admin page).
 * Calendar "location" shows start time + base service only (no design tier, no foreign soak-off add-on).
 */
function splitServiceSegments_(serviceBlob) {
  const s = String(serviceBlob == null ? '' : serviceBlob).trim();
  if (!s) return [];
  return s
    .split(/\s*[,，]\s*|\s*·\s*/)
    .map(function (p) {
      return String(p).trim();
    })
    .filter(function (p) {
      return p.length > 0;
    });
}

function isDesignTierOrSoakoffSegment_(segment) {
  const t = String(segment || '').trim();
  if (!t) return true;
  const low = t.toLowerCase();
  if (low.indexOf('foreign soak-off') >= 0) return true;
  if (/^tier\s*\d+/i.test(t)) return true;
  return false;
}

/** Labels like "Structured Gel New Set" / "Gel X Medium" — excludes Tier N and foreign soak-off. */
function baseServiceLabelForCalendar_(serviceBlob) {
  const segments = splitServiceSegments_(serviceBlob);
  const bases = segments.filter(function (p) {
    return !isDesignTierOrSoakoffSegment_(p);
  });
  if (bases.length > 0) {
    return bases.join(' + ');
  }
  if (segments.length > 0) {
    return segments.join(' + ');
  }
  return '';
}

/**
 * Calendar location reads cleaner in mixed case (some feeds or views show SHOUTY text).
 * Title-cases each word; keeps " + " between multiple base labels.
 */
function humanizeCalendarServiceWords_(s) {
  const t = String(s == null ? '' : s).trim();
  if (!t) return '';
  return t
    .split(/\s*\+\s*/)
    .map(function (seg) {
      return seg
        .trim()
        .split(/\s+/)
        .map(function (w) {
          if (!w) return w;
          return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
        })
        .join(' ');
    })
    .filter(function (x) {
      return x.length > 0;
    })
    .join(' + ');
}

function calendarLocationFromTimeAndService_(timeDisplay, serviceBlob) {
  const time = String(timeDisplay == null ? '' : timeDisplay).trim();
  const base = humanizeCalendarServiceWords_(baseServiceLabelForCalendar_(serviceBlob));
  if (time && base) return time + ' ' + base;
  if (time) return time;
  return base;
}

function applyBookingLocationToEvent_(ev, timeDisplay, serviceBlob) {
  if (!ev || typeof ev.setLocation !== 'function') return;
  const loc = calendarLocationFromTimeAndService_(timeDisplay, serviceBlob);
  if (loc) ev.setLocation(loc);
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

/** Booking POST sends this; also embedded in event description for reschedule if calendar length is wrong. */
function clampDurationMinutes_(n) {
  const x = Number(n);
  if (isNaN(x)) return 60;
  return Math.max(15, Math.min(480, Math.round(x)));
}

/** yyyy-MM-dd in script TZ: same calendar date exactly one month after today (Mar 31 → Apr 30). */
function maxPublicBookingDateYmd_() {
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const p = todayStr.split('-');
  var cal = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10), 12, 0, 0);
  cal.setMonth(cal.getMonth() + 1);
  return Utilities.formatDate(cal, tz, 'yyyy-MM-dd');
}

/**
 * Public website: appointment date must be on or before “one month from today” (script TZ).
 * Earliest bookable day is still enforced by isDateMeetingBookingLeadTime_ (not today/tomorrow).
 * Admin/owner POST handlers do not use this.
 */
function isDateInPublicBookingWindow_(yyyyMmDd) {
  const raw = String(yyyyMmDd == null ? '' : yyyyMmDd).trim().split('T')[0].split(' ')[0];
  const m = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return false;
  const maxD = maxPublicBookingDateYmd_();
  return raw <= maxD;
}

/** Earliest yyyy-MM-dd for public booking/reschedule (script TZ): not same-day or next calendar day. */
function minPublicBookingLeadDateYmd_() {
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const p = todayStr.split('-');
  var cal = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10), 12, 0, 0);
  cal.setDate(cal.getDate() + 2);
  return Utilities.formatDate(cal, tz, 'yyyy-MM-dd');
}

function isDateMeetingBookingLeadTime_(yyyyMmDd) {
  const s = String(yyyyMmDd == null ? '' : yyyyMmDd).trim().split('T')[0].split(' ')[0];
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  return s >= minPublicBookingLeadDateYmd_();
}

function parseDurationMinutesFromDescription_(description) {
  if (!description) return 0;
  const m = String(description).match(/DurationMinutes:\s*(\d+)/i);
  if (!m) return 0;
  const v = parseInt(m[1], 10);
  return isNaN(v) ? 0 : v;
}

/** Use the longer of calendar block length and DurationMinutes: line (fixes old 1h defaults vs 1h15 services). */
function effectiveDurationMinutesFromEvent_(ev) {
  const durMs = ev.getEndTime().getTime() - ev.getStartTime().getTime();
  const fromCal = Math.round(durMs / 60000);
  const fromDesc = parseDurationMinutesFromDescription_(ev.getDescription());
  return Math.max(30, Math.max(fromCal, fromDesc));
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

/** Column H status: trim + uppercase so "Confirmed " or formulas still match. */
function normalizeSheetStatus_(raw) {
  return String(raw == null ? '' : raw)
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

/**
 * If Advanced Calendar API is enabled (Editor → Services → Google Calendar API), this is authoritative
 * for deleted / cancelled events. CalendarApp alone can still return trashed events from getEventById().
 * Returns: 'active' | 'gone' | 'unknown' (unknown = API off or non-404 error — use CalendarApp heuristic).
 */
function calendarApiEventStatus_(eventId) {
  const want = String(eventId || '').trim();
  if (!want) return 'gone';
  try {
    if (typeof Calendar === 'undefined' || !Calendar.Events || typeof Calendar.Events.get !== 'function') {
      return 'unknown';
    }
  } catch (e) {
    return 'unknown';
  }
  try {
    const apiEv = Calendar.Events.get(CALENDAR_ID, want);
    if (!apiEv) return 'gone';
    if (apiEv.status === 'cancelled') return 'gone';
    return 'active';
  } catch (err) {
    const msg = String(err.message || err).toLowerCase();
    if (
      msg.indexOf('not found') >= 0 ||
      msg.indexOf('404') >= 0 ||
      msg.indexOf('requested entity was not found') >= 0
    ) {
      return 'gone';
    }
    return 'unknown';
  }
}

/**
 * Deleted events often stay in Calendar "trash" but disappear from normal listings.
 * getEventById() can still return them — so we only trust an event that also appears in getEvents(),
 * unless the Advanced Calendar API says the event is active (then we skip the list check).
 */
function getActiveBookingEvent_(calendar, eventId) {
  const want = String(eventId || '').trim();
  if (!want) return null;

  const api = calendarApiEventStatus_(want);
  if (api === 'gone') return null;

  let ev = null;
  try {
    ev = calendar.getEventById(want);
  } catch (err) {
    return null;
  }
  if (!ev) return null;

  if (api === 'active') {
    return ev;
  }

  try {
    const center = ev.getStartTime();
    const from = new Date(center.getTime() - 120 * 86400000);
    const to = new Date(center.getTime() + 120 * 86400000);
    const listed = calendar.getEvents(from, to);
    for (let i = 0; i < listed.length; i++) {
      if (String(listed[i].getId()) === want) return ev;
    }
  } catch (err2) {
    return null;
  }
  return null;
}

/**
 * --- Calendar → Sheet sync ---
 * - You moved the event: sheet + client “time updated” email.
 * - You deleted the event (on THIS studio calendar — CALENDAR_ID): sheet → CANCELLED, client email.
 * Best setup: (1) installStudioCalendarOnChangeTrigger() — must use studio calendar CALENDAR_ID (not
 * personal). (2) installCalendarSyncTrigger() — time-based backup; deletes sometimes don’t fire
 * onEventUpdated reliably on group calendars. (3) One deployment only (Head or version, not both).
 * Optional: Editor → Services → add “Google Calendar API” so deleted/cancelled events are detected reliably.
 *
 * Race guard: snapshots sheet rows at START; re-reads status + event id before cancelling. Approve must
 * write the new calendar id + CONFIRMED and flush before deleting the old PENDING event, and use the
 * same script lock as this sync — otherwise a sync can see “event gone” while the row is still PENDING
 * and send a false “Appointment Cancelled” email to the client.
 */
function syncCalendarToSpreadsheet() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('syncCalendarToSpreadsheet: skipped (another sync is running — avoids duplicate emails)');
    return;
  }
  try {
    syncCalendarToSpreadsheetBody_();
  } finally {
    lock.releaseLock();
  }
}

function syncCalendarToSpreadsheetBody_() {
  const tz = Session.getScriptTimeZone();
  const ss = getCRMSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  let qualifying = 0;
  let stillOnCalendar = 0;
  let markedCancelled = 0;
  let timeUpdated = 0;
  Logger.log(
    'syncCalendarToSpreadsheet: START sheet=' +
      SHEET_NAME +
      ' dataRows=' +
      (data.length - 1) +
      ' calendarId=' +
      CALENDAR_ID
  );
  for (let i = 1; i < data.length; i++) {
    try {
      const status = normalizeSheetStatus_(data[i][7]);
      const eventId = data[i][8];
      if (
        (status === 'CONFIRMED' || status === 'CLIENT_CONFIRMED' || status === 'PENDING' || status === 'MOD_PROPOSED') &&
        eventId &&
        String(eventId).indexOf('pending') < 0
      ) {
        qualifying++;
        let event = getActiveBookingEvent_(calendar, eventId);
        if (!event) {
          const statusFresh = normalizeSheetStatus_(sheet.getRange(i + 1, 8).getValue());
          if (statusFresh === 'CANCELLED') {
            Logger.log('syncCalendarToSpreadsheet: row ' + (i + 1) + ' already CANCELLED — skip (no duplicate email)');
            continue;
          }
          if (statusFresh === 'CONFIRMED' || statusFresh === 'CLIENT_CONFIRMED') {
            Logger.log(
              'syncCalendarToSpreadsheet: row ' +
                (i + 1) +
                ' skip cancel — sheet is now ' +
                statusFresh +
                ' (approve likely raced this sync; snapshot was ' +
                status +
                ')'
            );
            continue;
          }
          const eventIdFresh = String(sheet.getRange(i + 1, 9).getValue() || '').trim();
          if (eventIdFresh && eventIdFresh !== String(eventId).trim()) {
            event = getActiveBookingEvent_(calendar, eventIdFresh);
            if (event) {
              Logger.log(
                'syncCalendarToSpreadsheet: row ' +
                  (i + 1) +
                  ' skip cancel — event id was updated mid-sync (new id still on calendar)'
              );
            }
          }
        }
        if (!event) {
          const statusFresh = normalizeSheetStatus_(sheet.getRange(i + 1, 8).getValue());
          if (statusFresh === 'CONFIRMED' || statusFresh === 'CLIENT_CONFIRMED') {
            Logger.log(
              'syncCalendarToSpreadsheet: row ' +
                (i + 1) +
                ' skip cancel — re-check: now ' +
                statusFresh +
                ' after event id retry'
            );
            continue;
          }
          markedCancelled++;
          const clientName = data[i][1];
          const clientEmail = data[i][6];
          const service = data[i][3];
          const neatD = formatSheetDateForEmail(data[i][4]);
          const neatTime = formatSheetTimeForEmail(data[i][5]);
          Logger.log(
            'syncCalendarToSpreadsheet: row ' +
              (i + 1) +
              ' no active event → CANCELLED (status was ' +
              status +
              ') eventId=' +
              String(eventId).substring(0, 40) +
              '…'
          );
          sheet.getRange(i + 1, 8).setValue('CANCELLED');
          sheet.getRange(i + 1, 11).setValue('');
          SpreadsheetApp.flush();
          if (clientEmail) {
            const html = getCalendarDeletedClientEmailHtml(clientName, neatD, neatTime, service);
            MailApp.sendEmail({
              to: clientEmail,
              name: "Roni's Nail Studio",
              subject: "Appointment Cancelled: Roni's Nail Studio",
              htmlBody: html,
            });
            Logger.log('syncCalendarToSpreadsheet: cancellation email sent for row ' + (i + 1) + ' eventId=' + eventId);
          } else {
            Logger.log('syncCalendarToSpreadsheet: row ' + (i + 1) + ' event missing but no client email — no email sent');
          }
          continue;
        }
        stillOnCalendar++;
        const sheetDateYmd = syncSheetDateToYyyyMmDd_(data[i][4]);
        const sheetTimeNorm = normalizeTimeToken_(formatSheetTimeForEmail(data[i][5]));
        const calStart = event.getStartTime();
        const calDateYmd = Utilities.formatDate(calStart, tz, 'yyyy-MM-dd');
        const calTimeNorm = normalizeTimeToken_(Utilities.formatDate(calStart, tz, 'h:mm a'));
        if (calDateYmd !== sheetDateYmd || calTimeNorm !== sheetTimeNorm) {
          timeUpdated++;
          const calTimeDisplay = Utilities.formatDate(calStart, tz, 'h:mm a');
          sheet.getRange(i + 1, 5).setValue(calStart);
          sheet.getRange(i + 1, 6).setValue(calTimeDisplay);
          sheet.getRange(i + 1, 11).setValue('');
          SpreadsheetApp.flush();
          const neatDate = Utilities.formatDate(calStart, tz, 'EEEE, MMMM d, yyyy');
          const clientEmail = data[i][6];
          const service = data[i][3];
          applyBookingLocationToEvent_(event, calTimeDisplay, service);
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
    } catch (rowErr) {
      Logger.log('syncCalendarToSpreadsheet row ' + (i + 1) + ' skipped: ' + rowErr);
    }
  }
  Logger.log(
    'syncCalendarToSpreadsheet: DONE qualifying=' +
      qualifying +
      ' stillOnCalendar=' +
      stillOnCalendar +
      ' markedCancelled=' +
      markedCancelled +
      ' timeUpdated=' +
      timeUpdated
  );
}

/**
 * Time-driven backup for syncCalendarToSpreadsheet (does NOT remove calendar on-change triggers).
 * Important: installCalendarSyncTrigger used to bail out if *any* trigger existed — so with only a
 * Calendar trigger, no clock backup was added. Deletes on group calendars often need this backup.
 * Personal @gmail: may fall back to hourly.
 */
function installCalendarSyncTrigger() {
  const fn = 'syncCalendarToSpreadsheet';
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() !== fn) continue;
    if (triggers[i].getEventType() === ScriptApp.EventType.CLOCK) {
      Logger.log('Time-based trigger already exists for ' + fn);
      return;
    }
  }
  try {
    ScriptApp.newTrigger(fn).timeBased().everyMinutes(10).create();
    Logger.log('Installed time trigger: ' + fn + ' every 10 minutes');
  } catch (e) {
    ScriptApp.newTrigger(fn).timeBased().everyHours(1).create();
    Logger.log('10-minute trigger not allowed for this account; installed hourly instead. Reason: ' + e);
  }
}

/**
 * Calendar “event updated” trigger on the STUDIO calendar (CALENDAR_ID). Run this if your manual
 * trigger was tied to the wrong calendar — deletes won’t sync if the trigger watches personal, not studio.
 */
function installStudioCalendarOnChangeTrigger() {
  const fn = 'syncCalendarToSpreadsheet';
  const wantId = String(CALENDAR_ID).toLowerCase().replace(/\s/g, '');
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() !== fn) continue;
    if (triggers[i].getEventType() !== ScriptApp.EventType.ON_EVENT_UPDATED) continue;
    const src = String(triggers[i].getTriggerSourceId() || '')
      .toLowerCase()
      .replace(/\s/g, '');
    if (src === wantId) {
      Logger.log('Studio calendar on-change trigger already present (source ' + triggers[i].getTriggerSourceId() + ')');
      return;
    }
  }
  // forUserCalendar(id) accepts email or full calendar ID (e.g. …@group.calendar.google.com). forCalendar() is not on TriggerBuilder.
  ScriptApp.newTrigger(fn).forUserCalendar(CALENDAR_ID).onEventUpdated().create();
  Logger.log('Installed onEventUpdated trigger on studio calendar: ' + CALENDAR_ID);
}

/** Removes only ON_EVENT_UPDATED triggers for sync, then adds one on CALENDAR_ID. */
function reinstallStudioCalendarOnChangeTrigger() {
  const fn = 'syncCalendarToSpreadsheet';
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = triggers.length - 1; i >= 0; i--) {
    if (triggers[i].getHandlerFunction() !== fn) continue;
    if (triggers[i].getEventType() === ScriptApp.EventType.ON_EVENT_UPDATED) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger(fn).forUserCalendar(CALENDAR_ID).onEventUpdated().create();
  Logger.log('Reinstalled onEventUpdated on studio calendar: ' + CALENDAR_ID);
}

/**
 * Removes only CLOCK triggers for syncCalendarToSpreadsheet, then installs a fresh time-based one.
 * Does not remove calendar on-change triggers.
 */
function reinstallCalendarSyncTrigger() {
  const fn = 'syncCalendarToSpreadsheet';
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = triggers.length - 1; i >= 0; i--) {
    if (triggers[i].getHandlerFunction() !== fn) continue;
    if (triggers[i].getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  try {
    ScriptApp.newTrigger(fn).timeBased().everyMinutes(10).create();
    Logger.log('Reinstalled CLOCK: ' + fn + ' every 10 minutes');
  } catch (e) {
    ScriptApp.newTrigger(fn).timeBased().everyHours(1).create();
    Logger.log('Reinstalled CLOCK: ' + fn + ' every 1 hour (10 min not allowed). Reason: ' + e);
  }
}

/**
 * Run once from the editor: logs whether Calendar + sheet open, API service, and installed triggers.
 * Open Executions → select this run → View → to read the log.
 */
function validateCalendarSyncSetup() {
  const lines = [];
  lines.push('--- validateCalendarSyncSetup ---');
  try {
    const cal = CalendarApp.getCalendarById(CALENDAR_ID);
    lines.push('Calendar OK: ' + cal.getName() + ' (' + CALENDAR_ID + ')');
  } catch (e) {
    lines.push('Calendar ERROR (check CALENDAR_ID + script owner access): ' + e);
  }
  try {
    const ss = getCRMSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    lines.push('Sheet OK: ' + ss.getName() + ' tab=' + sh.getName() + ' rows=' + sh.getLastRow());
  } catch (e2) {
    lines.push('Sheet ERROR: ' + e2);
  }
  try {
    if (typeof Calendar !== 'undefined' && Calendar.Events && typeof Calendar.Events.get === 'function') {
      lines.push('Advanced Calendar API: ENABLED (good for delete/cancel detection)');
    } else {
      lines.push('Advanced Calendar API: NOT enabled → Editor → Services → add Google Calendar API');
    }
  } catch (e3) {
    lines.push('Advanced Calendar API: NOT enabled');
  }
  const triggers = ScriptApp.getProjectTriggers();
  let syncTriggers = 0;
  let hasClock = false;
  let hasCal = false;
  const wantCalNorm = String(CALENDAR_ID).toLowerCase().replace(/\s/g, '');
  let calendarTriggerMatchesStudio = false;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() !== 'syncCalendarToSpreadsheet') continue;
    syncTriggers++;
    const et = triggers[i].getEventType();
    if (et === ScriptApp.EventType.CLOCK) hasClock = true;
    if (et === ScriptApp.EventType.ON_EVENT_UPDATED) {
      hasCal = true;
      let src = '';
      try {
        src = triggers[i].getTriggerSourceId();
      } catch (e) {
        src = '';
      }
      if (String(src).toLowerCase().replace(/\s/g, '') === wantCalNorm) {
        calendarTriggerMatchesStudio = true;
      }
      lines.push(
        'Trigger: syncCalendarToSpreadsheet | eventType=' +
          et +
          (src ? ' | calendarSourceId=' + src : '')
      );
    } else {
      lines.push('Trigger: syncCalendarToSpreadsheet | eventType=' + et);
    }
  }
  if (syncTriggers === 0) {
    lines.push('WARNING: no trigger — run installStudioCalendarOnChangeTrigger() and installCalendarSyncTrigger()');
  } else {
    if (!hasCal) {
      lines.push('TIP: no ON_EVENT_UPDATED — run installStudioCalendarOnChangeTrigger() for fast updates');
    } else if (!calendarTriggerMatchesStudio) {
      lines.push(
        '!!! FIX THIS: Calendar trigger is on the WRONG calendar (e.g. personal @gmail). studio CALENDAR_ID must match calendarSourceId.'
      );
      lines.push('!!! Run reinstallStudioCalendarOnChangeTrigger() once, then validate again.');
    }
    if (!hasClock) {
      lines.push('TIP: no CLOCK backup — run installCalendarSyncTrigger() so deletes still sync if calendar trigger misses');
    }
  }
  const out = lines.join('\n');
  Logger.log(out);
  return out;
}

/**
 * One-time setup: daily trigger for sendTwoDayReminders (Confirm / Reschedule / Cancel email).
 * Runs at 8:00 AM in the project time zone (File → Project settings → Time zone).
 * Apps Script editor → select installTwoDayReminderTrigger → Run.
 */
function installTwoDayReminderTrigger() {
  const fn = 'sendTwoDayReminders';
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === fn) {
      Logger.log('Trigger already exists for ' + fn);
      return;
    }
  }
  ScriptApp.newTrigger(fn).timeBased().everyDays(1).atHour(8).create();
  Logger.log('Installed daily trigger: ' + fn + ' at 8:00 AM (project time zone)');
}

function tokenMatches(stored, provided) {
  if (stored === undefined || stored === null || provided === undefined || provided === null) return false;
  return String(stored).trim() === String(provided).trim();
}

function defaultWorkHoursObject_() {
  return { '1': { start: 11, end: 18 }, '2': { start: 11, end: 18 }, '3': { start: 9, end: 16 }, '5': { start: 9, end: 16 } };
}

function getWorkHoursObjectFromSheet_() {
  try {
    const ss = getCRMSpreadsheet();
    const sh = ss.getSheetByName(HOURS_SHEET_NAME);
    if (!sh) return null;
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return null;
    const out = {};
    for (let i = 1; i < data.length; i++) {
      const rawD = data[i][0];
      if (rawD === '' || rawD === null || rawD === undefined) continue;
      if (typeof rawD === 'object') continue;
      const d = parseInt(String(rawD), 10);
      const st = Number(data[i][1]);
      const en = Number(data[i][2]);
      if (isNaN(d) || d < 0 || d > 6) continue;
      if (!isFinite(st) || !isFinite(en) || st !== Math.floor(st) || en !== Math.floor(en)) continue;
      if (st < 0 || st > 23 || en < 1 || en > 24 || st >= en) continue;
      out[String(d)] = { start: st, end: en };
    }
    return Object.keys(out).length ? out : null;
  } catch (err) {
    return null;
  }
}

function getWorkHoursPayload_() {
  const fromSheet = getWorkHoursObjectFromSheet_();
  if (fromSheet) return fromSheet;
  return defaultWorkHoursObject_();
}

function handleAdminSetWorkHours(d) {
  const secret = PropertiesService.getScriptProperties().getProperty('BOOKING_ADMIN_SECRET');
  if (!secret || String(d.adminSecret || '').trim() !== String(secret).trim()) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'unauthorized' })).setMimeType(ContentService.MimeType.JSON);
  }
  if (!d.hours || typeof d.hours !== 'object' || Array.isArray(d.hours)) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'bad_hours' })).setMimeType(ContentService.MimeType.JSON);
  }
  const ss = getCRMSpreadsheet();
  let sh = ss.getSheetByName(HOURS_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(HOURS_SHEET_NAME);
  sh.clear();
  sh.appendRow(['day', 'start', 'end']);
  const rows = [];
  for (let day = 0; day <= 6; day++) {
    const h = d.hours[day] !== undefined && d.hours[day] !== null ? d.hours[day] : d.hours[String(day)];
    if (!h || h.open !== true) continue;
    const st = Math.floor(Number(h.start));
    const en = Math.floor(Number(h.end));
    if (!isFinite(st) || !isFinite(en) || st < 0 || st > 23 || en < 1 || en > 24 || st >= en) continue;
    rows.push([day, st, en]);
  }
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  SpreadsheetApp.flush();
  return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
}

/** Run once in the editor: replace the string with a long random secret, then Run. Do not commit the real secret. */
function setBookingAdminSecret() {
  PropertiesService.getScriptProperties().setProperty('BOOKING_ADMIN_SECRET', 'REPLACE_WITH_LONG_RANDOM_SECRET');
}

/** Optional: create StudioHours rows from the same defaults as the booking page (before first save from the admin page). */
function seedStudioHoursSheet() {
  const ss = getCRMSpreadsheet();
  let sh = ss.getSheetByName(HOURS_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(HOURS_SHEET_NAME);
  sh.clear();
  sh.appendRow(['day', 'start', 'end']);
  const def = defaultWorkHoursObject_();
  const rows = [];
  Object.keys(def).forEach(function (k) {
    rows.push([parseInt(k, 10), def[k].start, def[k].end]);
  });
  rows.sort(function (a, b) {
    return a[0] - b[0];
  });
  if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
}

function doGet(e) {
  const action = e.parameter.action;
  const eventId = e.parameter.eventId;
  const token = e.parameter.token;

  if (action === 'work_hours') {
    const callback = e.parameter.callback;
    const payload = getWorkHoursPayload_();
    const json = JSON.stringify(payload);
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }

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
    if (rowStatus !== 'PENDING' && rowStatus !== 'MOD_PROPOSED') {
      return htmlPage('Already handled', '<h2>Already handled</h2><p>This request was already approved or declined.</p>');
    }
    if (action === 'reject') {
      const rb = String(OWNER_REJECT_PAGE_BASE || '').trim();
      if (rb) {
        const sep = rb.indexOf('?') >= 0 ? '&' : '?';
        const target =
          rb +
          sep +
          'eventId=' +
          encodeURIComponent(String(eventId)) +
          '&token=' +
          encodeURIComponent(String(token));
        const html =
          '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Decline request</title></head><body><p>Loading…</p><script>location.replace(' +
          JSON.stringify(target) +
          ');</script><p style="font-family:system-ui,sans-serif;text-align:center;padding:2rem"><a href=' +
          JSON.stringify(target) +
          '>Continue</a></p></body></html>';
        return ContentService.createTextOutput(html).setMimeType(ContentService.MimeType.HTML);
      }
    }
    if (action === 'accept') {
        if (rowStatus === 'MOD_PROPOSED') {
          return htmlPage(
            'Waiting on client',
            '<h2>Alternate time pending</h2><p>You already suggested a different time. Wait for the client to confirm it, or use <strong>Decline</strong> to cancel this request.</p>'
          );
        }
        const lock = LockService.getScriptLock();
        if (!lock.tryLock(20000)) {
          return htmlPage('Busy', '<h2>Please try again</h2><p>Another update is in progress. Wait a moment and tap Approve again.</p>');
        }
        try {
          const cal = CalendarApp.getCalendarById(CALENDAR_ID);
          const ev = cal.getEventById(eventId);
          if (ev) {
            const nEv = cal.createEvent(clientName, ev.getStartTime(), ev.getEndTime(), { description: ev.getDescription() });
            applyBookingLocationToEvent_(
              nEv,
              Utilities.formatDate(nEv.getStartTime(), Session.getScriptTimeZone(), 'h:mm a'),
              service
            );
            sheet.getRange(rowIndex, 9).setValue(nEv.getId());
            sheet.getRange(rowIndex, 8).setValue('CONFIRMED');
            SpreadsheetApp.flush();
            ev.deleteEvent();
          } else {
            sheet.getRange(rowIndex, 8).setValue('CONFIRMED');
            SpreadsheetApp.flush();
          }
        } finally {
          lock.releaseLock();
        }
        const neatD = formatSheetDateForEmail(dateVal);
        const neatTime = formatSheetTimeForEmail(timeStr);
        const acceptEmailHtml = getConfirmedEmailHtml(clientName, neatD, neatTime, service);
        MailApp.sendEmail({ to: clientEmail, name: "Roni's Nail Studio", subject: "Appointment Confirmed: Roni's Nail Studio", htmlBody: acceptEmailHtml });
        SpreadsheetApp.flush();
        return htmlPage('Accepted', '<h2>Request accepted</h2><p>The client has been notified.</p>');
    }
    if (action === 'reject') {
        const lock = LockService.getScriptLock();
        if (!lock.tryLock(20000)) {
          return htmlPage('Busy', '<h2>Please try again</h2><p>Another update is in progress. Wait a moment and try again.</p>');
        }
        try {
          sheet.getRange(rowIndex, 8).setValue('REJECTED');
          sheet.getRange(rowIndex, 12).setValue('');
          sheet.getRange(rowIndex, 13).setValue('');
          sheet.getRange(rowIndex, 14).setValue('');
          sheet.getRange(rowIndex, SHEET_COL_OWNER_DECLINE_REASON).setValue('');
          SpreadsheetApp.flush();
          const cal = CalendarApp.getCalendarById(CALENDAR_ID);
          const ev = cal.getEventById(eventId);
          if (ev) ev.deleteEvent();
        } finally {
          lock.releaseLock();
        }
        const neatD = formatSheetDateForEmail(dateVal);
        const neatTime = formatSheetTimeForEmail(timeStr);
        const declinedHtml = getDeclinedEmailHtml(clientName, neatD, neatTime, service, '');
        if (clientEmail) {
          MailApp.sendEmail({ to: clientEmail, name: "Roni's Nail Studio", subject: "Update on your booking request: Roni's Nail Studio", htmlBody: declinedHtml });
        }
        SpreadsheetApp.flush();
        return htmlPage('Declined', '<h2>Request declined</h2><p>The client has been notified.</p>');
    }
    return htmlPage('Not supported', '<h2>Unsupported action</h2>');
  }

  if (action === 'client_accept_mod' || action === 'client_decline_mod') {
    if (!eventId || !token) {
      return htmlPage('Invalid link', '<h2>Invalid link</h2><p>This link is incomplete.</p>');
    }
    const ss = getCRMSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let clientName = 'Client';
    let clientEmail = '';
    let service = '';
    let dateVal = '';
    let timeStr = '';
    let rowStatus = '';
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][8]) !== String(eventId)) continue;
      if (!tokenMatches(data[i][13], token)) {
        return htmlPage('Invalid link', '<h2>Invalid or expired link</h2><p>Use the buttons in the alternate-time email.</p>');
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
    if (rowIndex < 0) {
      return htmlPage('Not found', '<h2>Booking not found</h2>');
    }
    if (rowStatus !== 'MOD_PROPOSED') {
      return htmlPage('Not available', '<h2>Link not valid</h2><p>This alternate-time request is no longer active.</p>');
    }
    const proposedDateVal = data[rowIndex - 1][11];
    const proposedTimeRaw = data[rowIndex - 1][12];
    if (proposedTimeRaw === '' || proposedTimeRaw === null || proposedTimeRaw === undefined) {
      return htmlPage('Error', '<h2>Something went wrong</h2><p>Missing proposed time. Please contact the studio.</p>');
    }
    const tz = Session.getScriptTimeZone();
    const datePart = syncSheetDateToYyyyMmDd_(proposedDateVal);
    const propTimeStr = formatSheetTimeForEmail(proposedTimeRaw);
    if (!propTimeStr) {
      return htmlPage('Error', '<h2>Something went wrong</h2><p>Missing proposed time. Please contact the studio.</p>');
    }
    if (!datePart || !/^\d{4}-\d{2}-\d{2}$/.test(datePart)) {
      return htmlPage('Error', '<h2>Invalid date</h2><p>Please contact the studio.</p>');
    }
    if (action === 'client_decline_mod') {
      const lock = LockService.getScriptLock();
      if (!lock.tryLock(20000)) {
        return htmlPage('Busy', '<h2>Please try again</h2><p>Another update is in progress.</p>');
      }
      try {
        sheet.getRange(rowIndex, 8).setValue('MOD_DECLINED');
        sheet.getRange(rowIndex, 12).setValue('');
        sheet.getRange(rowIndex, 13).setValue('');
        sheet.getRange(rowIndex, 14).setValue('');
        SpreadsheetApp.flush();
        const cal = CalendarApp.getCalendarById(CALENDAR_ID);
        const ev = cal.getEventById(eventId);
        if (ev) ev.deleteEvent();
      } finally {
        lock.releaseLock();
      }
      const neatD = formatSheetDateForEmail(dateVal);
      const neatTime = formatSheetTimeForEmail(timeStr);
      if (clientEmail) {
        const declineHtml = getAlternateDeclinedClientEmailHtml(clientName, neatD, neatTime, service);
        MailApp.sendEmail({
          to: clientEmail,
          name: "Roni's Nail Studio",
          subject: "Alternate time — Roni's Nail Studio",
          htmlBody: declineHtml,
        });
      }
      MailApp.sendEmail({
        to: MY_EMAIL,
        subject: 'Client declined alternate time: ' + clientName,
        htmlBody:
          '<p style="font-family:sans-serif;">' +
          escapeHtml(clientName) +
          ' declined the proposed alternate time. Original request was ' +
          escapeHtml(neatD) +
          ' at ' +
          escapeHtml(neatTime) +
          '.</p>',
      });
      return htmlPage('Recorded', '<h2>Thanks for letting us know</h2><p>You can book another time on our website whenever you like.</p>');
    }
    const start = parseYmdAndTimeLocal_(datePart, convertTo24Hour(propTimeStr));
    if (isNaN(start.getTime())) {
      return htmlPage('Error', '<h2>Invalid time</h2><p>Please contact the studio.</p>');
    }
    const cal = CalendarApp.getCalendarById(CALENDAR_ID);
    const ev = cal.getEventById(eventId);
    if (!ev) {
      return htmlPage('Error', '<h2>Calendar error</h2><p>Please contact the studio.</p>');
    }
    const durMin = effectiveDurationMinutesFromEvent_(ev);
    const newEnd = new Date(start.getTime() + durMin * 60000);
    if (slotOverlapsExistingCalendarEvents_(start, newEnd, [eventId])) {
      return htmlPage(
        'Error',
        '<h2>Time no longer available</h2><p>That slot was filled before your confirmation. Please contact the studio to choose another time.</p>'
      );
    }
    const neatTimeDisplay = Utilities.formatDate(start, tz, 'h:mm a');
    const nEv = cal.createEvent(clientConfirmedCalendarEventTitle_(clientName), start, newEnd, { description: ev.getDescription() });
    applyBookingLocationToEvent_(nEv, neatTimeDisplay, service);
    const lockAlt = LockService.getScriptLock();
    if (!lockAlt.tryLock(20000)) {
      nEv.deleteEvent();
      return htmlPage('Busy', '<h2>Please try again</h2><p>Another update is in progress.</p>');
    }
    try {
      sheet.getRange(rowIndex, 9).setValue(nEv.getId());
      sheet.getRange(rowIndex, 5).setValue(start);
      sheet.getRange(rowIndex, 6).setValue(neatTimeDisplay);
      sheet.getRange(rowIndex, 8).setValue('CONFIRMED');
      sheet.getRange(rowIndex, 12).setValue('');
      sheet.getRange(rowIndex, 13).setValue('');
      sheet.getRange(rowIndex, 14).setValue('');
      SpreadsheetApp.flush();
      ev.deleteEvent();
    } finally {
      lockAlt.releaseLock();
    }
    const neatDate = Utilities.formatDate(start, tz, 'EEEE, MMMM d, yyyy');
    const acceptHtml = getConfirmedEmailHtml(clientName, neatDate, neatTimeDisplay, service);
    if (clientEmail) {
      MailApp.sendEmail({ to: clientEmail, name: "Roni's Nail Studio", subject: "Appointment Confirmed: Roni's Nail Studio", htmlBody: acceptHtml });
    }
    MailApp.sendEmail({
      to: MY_EMAIL,
      subject: 'Client accepted alternate time: ' + clientName,
      htmlBody:
        '<p style="font-family:sans-serif;"><strong>' +
        escapeHtml(clientName) +
        '</strong> accepted the proposed time: <strong>' +
        escapeHtml(neatDate) +
        '</strong> at <strong>' +
        escapeHtml(neatTimeDisplay) +
        '</strong>.</p>',
    });
    return htmlPage('Confirmed', '<h2>You\'re all set!</h2><p>Your appointment is confirmed. We\'ll see you then.</p>');
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
        return htmlPage('Cancelled', '<h2>Cancelled</h2><p>This appointment has been cancelled.</p>');
      }
      return htmlPage('Invalid', '<h2>Link not valid</h2><p>This reminder link only applies to confirmed appointments.</p>');
    }
    if (action === 'client_confirm') {
      const lock = LockService.getScriptLock();
      if (!lock.tryLock(15000)) {
        return htmlPage('Busy', '<h2>Please try again</h2><p>Another update is in progress.</p>');
      }
      try {
        const statusNow = normalizeSheetStatus_(sheet.getRange(rowIndex, 8).getValue());
        if (statusNow === 'CLIENT_CONFIRMED') {
          return htmlPage('Confirmed', '<h2>Already confirmed</h2><p>We already have your confirmation. See you soon!</p>');
        }
        if (statusNow !== 'CONFIRMED') {
          if (statusNow === 'CANCELLED' || statusNow === 'REJECTED') {
            return htmlPage('Cancelled', '<h2>Cancelled</h2><p>This appointment has been cancelled.</p>');
          }
          return htmlPage('Invalid', '<h2>Link not valid</h2><p>This reminder link only applies to confirmed appointments.</p>');
        }
        sheet.getRange(rowIndex, 8).setValue('CLIENT_CONFIRMED');
        const calConfirm = CalendarApp.getCalendarById(CALENDAR_ID);
        const evConfirm = calConfirm.getEventById(eventId);
        if (evConfirm) {
          evConfirm.setTitle(clientConfirmedCalendarEventTitle_(clientName));
          applyBookingLocationToEvent_(
            evConfirm,
            Utilities.formatDate(evConfirm.getStartTime(), Session.getScriptTimeZone(), 'h:mm a'),
            service
          );
        }
        SpreadsheetApp.flush();
        const neatD = formatSheetDateForEmail(dateVal);
        const neatTime = formatSheetTimeForEmail(timeStr);
        const ownerRows = emailDetailRow('Date', neatD) + emailDetailRow('Time', neatTime) + emailDetailRow('Service', service);
        const ownerBody =
          '<div style="font-family:sans-serif;padding:24px;max-width:480px;margin:auto;"><h2 style="font-weight:500;">Client confirmed (2-day reminder)</h2><p><strong>' +
          escapeHtml(clientName) +
          '</strong> tapped <strong>Confirm</strong> on their reminder email.</p><table style="width:100%;border-collapse:collapse;margin-top:16px;background:#fafafa;border-radius:8px;"><tbody>' +
          ownerRows +
          '</tbody></table><p style="margin-top:16px;color:#666;">Client email: ' +
          escapeHtml(clientEmail || '—') +
          '</p></div>';
        MailApp.sendEmail({ to: MY_EMAIL, subject: 'Client confirmed appointment: ' + clientName, htmlBody: ownerBody });
        return htmlPage('Thank you', '<h2>Thank you!</h2><p>Your appointment is confirmed. We\'ll see you soon.</p>');
      } finally {
        lock.releaseLock();
      }
    }
    if (action === 'client_cancel') {
      const lock = LockService.getScriptLock();
      if (!lock.tryLock(15000)) {
        return htmlPage('Busy', '<h2>Please try again</h2><p>Another update is in progress.</p>');
      }
      try {
        const statusNow = normalizeSheetStatus_(sheet.getRange(rowIndex, 8).getValue());
        if (statusNow === 'CANCELLED') {
          return htmlPage('Cancelled', '<h2>Appointment cancelled</h2><p>We\'ve sent a confirmation to your email.</p>');
        }
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
      } finally {
        lock.releaseLock();
      }
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
    const durationMinutes = effectiveDurationMinutesFromEvent_(ev);
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

/**
 * Owner declines from reject-booking.html: optional note to client + column P on sheet.
 */
function handleOwnerRejectBooking(d) {
  const eventId = String(d.eventId || '').trim();
  const token = String(d.ownerToken != null && d.ownerToken !== '' ? d.ownerToken : d.token || '').trim();
  const reason = sanitizeOwnerDeclineReason_(d.declineReason);
  if (!eventId || !token) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'missing_fields' })).setMimeType(ContentService.MimeType.JSON);
  }
  const ss = getCRMSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  let clientName = 'Client';
  let clientEmail = '';
  let service = '';
  let dateVal = '';
  let timeStr = '';
  let rowStatus = '';
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][8]) !== String(eventId)) continue;
    if (!tokenMatches(data[i][9], token)) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'invalid_token' })).setMimeType(ContentService.MimeType.JSON);
    }
    rowIndex = i + 1;
    rowStatus = normalizeSheetStatus_(data[i][7]);
    clientName = data[i][1];
    clientEmail = data[i][6];
    service = data[i][3];
    dateVal = data[i][4];
    timeStr = data[i][5];
    break;
  }
  if (rowIndex < 0) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'not_found' })).setMimeType(ContentService.MimeType.JSON);
  }
  if (rowStatus !== 'PENDING' && rowStatus !== 'MOD_PROPOSED') {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'already_handled' })).setMimeType(ContentService.MimeType.JSON);
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'server_busy' })).setMimeType(ContentService.MimeType.JSON);
  }
  try {
    sheet.getRange(rowIndex, 8).setValue('REJECTED');
    sheet.getRange(rowIndex, 12).setValue('');
    sheet.getRange(rowIndex, 13).setValue('');
    sheet.getRange(rowIndex, 14).setValue('');
    sheet.getRange(rowIndex, SHEET_COL_OWNER_DECLINE_REASON).setValue(reason);
    SpreadsheetApp.flush();
    const cal = CalendarApp.getCalendarById(CALENDAR_ID);
    const ev = cal.getEventById(eventId);
    if (ev) ev.deleteEvent();
  } finally {
    lock.releaseLock();
  }
  const neatD = formatSheetDateForEmail(dateVal);
  const neatTime = formatSheetTimeForEmail(timeStr);
  const declinedHtml = getDeclinedEmailHtml(clientName, neatD, neatTime, service, reason);
  if (clientEmail) {
    MailApp.sendEmail({
      to: clientEmail,
      name: "Roni's Nail Studio",
      subject: "Update on your booking request: Roni's Nail Studio",
      htmlBody: declinedHtml,
    });
  }
  SpreadsheetApp.flush();
  return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
}

function handleOwnerProposeAlternate(d) {
  if (!d.eventId || !d.ownerToken || !d.proposedDate || !d.proposedTime) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'missing_fields' })).setMimeType(ContentService.MimeType.JSON);
  }
  const ss = getCRMSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  let clientEmail = '';
  let clientName = '';
  let service = '';
  let origDateVal = '';
  let origTimeStr = '';
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][8]) !== String(d.eventId)) continue;
    if (!tokenMatches(data[i][9], d.ownerToken)) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'invalid_token' })).setMimeType(ContentService.MimeType.JSON);
    }
    rowIndex = i + 1;
    const st = data[i][7];
    if (st !== 'PENDING') {
      if (st === 'MOD_PROPOSED') {
        return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'already_proposed' })).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'not_pending' })).setMimeType(ContentService.MimeType.JSON);
    }
    clientName = data[i][1];
    clientEmail = data[i][6];
    service = data[i][3];
    origDateVal = data[i][4];
    origTimeStr = data[i][5];
    break;
  }
  if (rowIndex < 0) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'not_found' })).setMimeType(ContentService.MimeType.JSON);
  }
  const pds = d.proposedDate.toString().split('T')[0];
  const propTimeTrim = String(d.proposedTime).trim();
  const testStart = parseYmdAndTimeLocal_(pds, convertTo24Hour(propTimeTrim));
  if (isNaN(testStart.getTime())) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'bad_datetime' })).setMimeType(ContentService.MimeType.JSON);
  }
  const calPropose = CalendarApp.getCalendarById(CALENDAR_ID);
  const pendingEvPropose = calPropose.getEventById(d.eventId);
  const proposeDurMin = pendingEvPropose ? effectiveDurationMinutesFromEvent_(pendingEvPropose) : 60;
  const proposeSlotEnd = new Date(testStart.getTime() + proposeDurMin * 60000);
  if (slotOverlapsExistingCalendarEvents_(testStart, proposeSlotEnd, [d.eventId])) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'slot_unavailable' })).setMimeType(ContentService.MimeType.JSON);
  }
  const tz = Session.getScriptTimeZone();
  const pdate = new Date(pds + 'T12:00:00');
  const modTok = generateActionToken();
  sheet.getRange(rowIndex, 12).setValue(pdate);
  sheet.getRange(rowIndex, 13).setValue(propTimeTrim);
  sheet.getRange(rowIndex, 14).setValue(modTok);
  sheet.getRange(rowIndex, 8).setValue('MOD_PROPOSED');
  SpreadsheetApp.flush();
  const origNeatD = formatSheetDateForEmail(origDateVal);
  const origNeatT = formatSheetTimeForEmail(origTimeStr);
  const newNeatD = Utilities.formatDate(pdate, tz, 'EEEE, MMMM d, yyyy');
  if (clientEmail) {
    const html = getAlternateProposalClientEmailHtml(
      clientName,
      origNeatD,
      origNeatT,
      newNeatD,
      propTimeTrim,
      service,
      d.eventId,
      modTok
    );
    MailApp.sendEmail({
      to: clientEmail,
      name: "Roni's Nail Studio",
      subject: "Suggested appointment time — Roni's Nail Studio",
      htmlBody: html,
    });
  }
  MailApp.sendEmail({
    to: MY_EMAIL,
    subject: 'Alternate time proposed to client: ' + clientName,
    htmlBody:
      '<p style="font-family:sans-serif;">Proposed <strong>' +
      escapeHtml(newNeatD) +
      ' at ' +
      escapeHtml(propTimeTrim) +
      '</strong> to ' +
      escapeHtml(clientName) +
      ' (was ' +
      escapeHtml(origNeatD) +
      ' at ' +
      escapeHtml(origNeatT) +
      ').</p>',
  });
  return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
}

function handleReschedulePost(d) {
  if (!d.eventId || !d.token || !d.date || !d.time) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'missing_fields' })).setMimeType(ContentService.MimeType.JSON);
  }
  const rs = d.date ? String(d.date).split('T')[0].trim() : '';
  if (!rs || !/^\d{4}-\d{2}-\d{2}$/.test(rs)) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'bad_date' })).setMimeType(ContentService.MimeType.JSON);
  }
  if (!isDateInPublicBookingWindow_(rs)) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'date_outside_booking_window' })).setMimeType(ContentService.MimeType.JSON);
  }
  if (!isDateMeetingBookingLeadTime_(rs)) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'date_too_soon' })).setMimeType(ContentService.MimeType.JSON);
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
  const durMin = effectiveDurationMinutesFromEvent_(ev);
  const start = parseYmdAndTimeLocal_(rs, convertTo24Hour(d.time));
  if (isNaN(start.getTime())) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'bad_datetime' })).setMimeType(ContentService.MimeType.JSON);
  }
  if (!isBookingWithinStudioHours_(start, durMin)) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'invalid_day_or_time' })).setMimeType(ContentService.MimeType.JSON);
  }
  const newEnd = new Date(start.getTime() + durMin * 60000);
  if (slotOverlapsExistingCalendarEvents_(start, newEnd, [d.eventId])) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'slot_unavailable' })).setMimeType(ContentService.MimeType.JSON);
  }
  ev.setTime(start, newEnd);
  ev.setTitle(clientName);
  const neatTimeStr = Utilities.formatDate(start, Session.getScriptTimeZone(), 'h:mm a');
  applyBookingLocationToEvent_(ev, neatTimeStr, service);
  sheet.getRange(rowIndex, 5).setValue(start);
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

/** Stable keys so +1 vs 10-digit phone and email both dedupe the same person. */
function ownerLookupDedupeKeys_(email, phone) {
  const keys = [];
  const em = String(email == null ? '' : email)
    .trim()
    .toLowerCase();
  if (em && em.indexOf('@') > 0) keys.push('e:' + em);
  let dig = String(phone == null ? '' : phone).replace(/\D/g, '');
  if (dig.length === 11 && dig.charAt(0) === '1') dig = dig.slice(1);
  if (dig.length >= 10) keys.push('p:' + dig.slice(-10));
  return keys;
}

/** Admin page: search Bookings by client name (substring). Dedupes by email and/or normalized phone. */
function handleOwnerClientLookup(d) {
  const secret = PropertiesService.getScriptProperties().getProperty('BOOKING_ADMIN_SECRET');
  if (!secret || String(d.adminSecret || '').trim() !== String(secret).trim()) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'unauthorized' })).setMimeType(ContentService.MimeType.JSON);
  }
  const q = String(d.query || '')
    .trim()
    .toLowerCase();
  if (q.length < 2) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', matches: [] })).setMimeType(ContentService.MimeType.JSON);
  }
  const ss = getCRMSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const seen = Object.create(null);
  const matches = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const name = data[i][1];
    const phone = data[i][2];
    const email = data[i][6];
    const ts = data[i][0];
    if (name == null || String(name).trim() === '') continue;
    if (String(name).toLowerCase().indexOf(q) < 0) continue;
    const ph = String(phone == null ? '' : phone).trim();
    const keys = ownerLookupDedupeKeys_(email, phone);
    if (keys.length === 0) continue;
    let duplicate = false;
    for (let k = 0; k < keys.length; k++) {
      if (seen[keys[k]]) {
        duplicate = true;
        break;
      }
    }
    if (duplicate) continue;
    for (let k = 0; k < keys.length; k++) seen[keys[k]] = true;
    let lastBooked = '';
    if (ts instanceof Date && !isNaN(ts.getTime())) {
      lastBooked = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'MMM d, yyyy');
    }
    matches.push({
      name: String(name).trim(),
      phone: ph,
      email: String(email || '').trim(),
      lastBooked: lastBooked,
    });
    if (matches.length >= 15) break;
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', matches: matches })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Admin page: append CONFIRMED row, create calendar event (not PENDING), send client confirmation + owner note.
 */
function handleOwnerDirectBooking(d) {
  const secret = PropertiesService.getScriptProperties().getProperty('BOOKING_ADMIN_SECRET');
  if (!secret || String(d.adminSecret || '').trim() !== String(secret).trim()) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'unauthorized' })).setMimeType(ContentService.MimeType.JSON);
  }
  const clientName = String(d.clientName || '').trim();
  const phone = String(d.phone || '').trim();
  const email = String(d.email || '').trim();
  const service = String(d.service || '').trim();
  if (!clientName || !email || !service) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'missing_fields' })).setMimeType(ContentService.MimeType.JSON);
  }
  const dateStr = d.date ? String(d.date).split('T')[0].trim() : '';
  const timeStr = String(d.time || '').trim();
  if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr) || !timeStr) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'bad_datetime' })).setMimeType(ContentService.MimeType.JSON);
  }
  const start = parseYmdAndTimeLocal_(dateStr, convertTo24Hour(timeStr));
  if (isNaN(start.getTime())) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'bad_datetime' })).setMimeType(ContentService.MimeType.JSON);
  }
  const durMin = clampDurationMinutes_(d.durationMinutes);
  const end = new Date(start.getTime() + durMin * 60000);
  const actionToken = generateActionToken();
  const ss = getCRMSpreadsheet();
  const s = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const row = s.getLastRow() + 1;
  s.appendRow([new Date(), clientName, phone, service, start, timeStr, email, 'CONFIRMED', '', actionToken, '', '', '', '', '']);
  const c = CalendarApp.getCalendarById(CALENDAR_ID);
  const desc =
    'Phone: ' +
    phone +
    '\nEmail: ' +
    email +
    '\nService: ' +
    service +
    '\nDurationMinutes: ' +
    durMin +
    '\nBooked by: owner (admin page)';
  const ev = c.createEvent(clientName, start, end, { description: desc });
  const tz = Session.getScriptTimeZone();
  const neatTime = Utilities.formatDate(start, tz, 'h:mm a');
  applyBookingLocationToEvent_(ev, neatTime, service);
  s.getRange(row, 9).setValue(ev.getId());
  SpreadsheetApp.flush();
  const neatD = Utilities.formatDate(start, tz, 'EEEE, MMMM d, yyyy');
  const acceptEmailHtml = getConfirmedEmailHtml(clientName, neatD, neatTime, service);
  MailApp.sendEmail({
    to: email,
    name: "Roni's Nail Studio",
    subject: "Appointment Confirmed: Roni's Nail Studio",
    htmlBody: acceptEmailHtml,
  });
  const ownerRows =
    emailDetailRow('Client', clientName) +
    emailDetailRow('When', neatD + ' · ' + neatTime) +
    emailDetailRow('Service', service) +
    emailDetailRow('Email', email);
  MailApp.sendEmail({
    to: MY_EMAIL,
    subject: 'Studio booked (admin): ' + clientName,
    htmlBody:
      '<p style="font-family:sans-serif;">You added a confirmed appointment from the admin page.</p><table style="width:100%;border-collapse:collapse;margin-top:12px;background:#fafafa;border-radius:8px;"><tbody>' +
      ownerRows +
      '</tbody></table>',
  });
  return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Admin page: JSON calendar (studio + personal). No iframe.
 * POST { ownerCalendarWeek: true, adminSecret,
 *   calView?: 'week' | 'month' (default week),
 *   weekStart?: 'yyyy-MM-dd' (any day in week → Sunday),
 *   month?: 'yyyy-MM' (for month view) }
 */
function handleOwnerCalendarWeek(d) {
  const secret = PropertiesService.getScriptProperties().getProperty('BOOKING_ADMIN_SECRET');
  if (!secret || String(d.adminSecret || '').trim() !== String(secret).trim()) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'unauthorized' })).setMimeType(ContentService.MimeType.JSON);
  }
  const tz = Session.getScriptTimeZone();
  const viewRaw = String(d.calView == null ? 'week' : d.calView)
    .trim()
    .toLowerCase();

  if (viewRaw === 'month') {
    let year;
    let month0;
    const mk = d.month ? String(d.month).trim() : '';
    if (/^\d{4}-\d{2}$/.test(mk)) {
      const pp = mk.split('-');
      year = parseInt(pp[0], 10);
      month0 = parseInt(pp[1], 10) - 1;
    } else {
      const now = new Date();
      year = now.getFullYear();
      month0 = now.getMonth();
    }
    if (isNaN(year) || month0 < 0 || month0 > 11) {
      const now2 = new Date();
      year = now2.getFullYear();
      month0 = now2.getMonth();
    }
    const monthStart = new Date(year, month0, 1, 0, 0, 0);
    const monthEndExclusive = new Date(year, month0 + 1, 1, 0, 0, 0);
    const events = collectCalendarWeekEvents_(monthStart, monthEndExclusive);
    const built = buildMonthDaysPayload_(year, month0, tz);
    return ContentService.createTextOutput(
      JSON.stringify({
        status: 'success',
        timeZone: tz,
        calView: 'month',
        month: built.monthKey,
        leadingBlankDays: built.leadingBlankDays,
        days: built.days,
        events: events,
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  let weekStart;
  const ws = d.weekStart ? String(d.weekStart).split('T')[0].trim() : '';
  if (ws && /^\d{4}-\d{2}-\d{2}$/.test(ws)) {
    const anchor = parseYmdAndTimeLocal_(ws, '12:00:00');
    if (!isNaN(anchor.getTime())) {
      const dow = anchor.getDay();
      weekStart = new Date(anchor.getFullYear(), anchor.getMonth(), anchor.getDate() - dow, 0, 0, 0);
    }
  }
  if (!weekStart || isNaN(weekStart.getTime())) {
    const todayYmd = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const anchor = parseYmdAndTimeLocal_(todayYmd, '12:00:00');
    const dow = anchor.getDay();
    weekStart = new Date(anchor.getFullYear(), anchor.getMonth(), anchor.getDate() - dow, 0, 0, 0);
  }
  const weekEndExclusive = new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate() + 7, 0, 0, 0);
  const events = collectCalendarWeekEvents_(weekStart, weekEndExclusive);
  const days = buildWeekDaysPayload_(weekStart, tz);
  return ContentService.createTextOutput(
    JSON.stringify({
      status: 'success',
      timeZone: tz,
      calView: 'week',
      days: days,
      events: events,
    })
  ).setMimeType(ContentService.MimeType.JSON);
}

function buildMonthDaysPayload_(year, month0, tz) {
  const leadingBlankDays = new Date(year, month0, 1, 0, 0, 0).getDay();
  const lastDayNum = new Date(year, month0 + 1, 0).getDate();
  const days = [];
  for (var dom = 1; dom <= lastDayNum; dom++) {
    const ymd = Utilities.formatDate(new Date(year, month0, dom), tz, 'yyyy-MM-dd');
    const dayStart = parseYmdAndTimeLocal_(ymd, '00:00:00');
    const nextDay = new Date(year, month0, dom + 1, 0, 0, 0);
    const label = Utilities.formatDate(new Date(year, month0, dom), tz, 'EEE, MMM d');
    days.push({
      ymd: ymd,
      startMs: dayStart.getTime(),
      endMs: nextDay.getTime(),
      label: label,
      dayOfMonth: dom,
    });
  }
  const monthKey = Utilities.formatDate(new Date(year, month0, 1), tz, 'yyyy-MM');
  return { leadingBlankDays: leadingBlankDays, days: days, monthKey: monthKey };
}

function buildWeekDaysPayload_(weekStart, tz) {
  const days = [];
  var cursor = new Date(weekStart.getTime());
  for (var i = 0; i < 7; i++) {
    const ymd = Utilities.formatDate(cursor, tz, 'yyyy-MM-dd');
    const dayStart = parseYmdAndTimeLocal_(ymd, '00:00:00');
    const next = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate() + 1, 0, 0, 0);
    const label = Utilities.formatDate(cursor, tz, 'EEE, MMM d');
    days.push({
      ymd: ymd,
      startMs: dayStart.getTime(),
      endMs: next.getTime(),
      label: label,
    });
    cursor = next;
  }
  return days;
}

function collectCalendarWeekEvents_(rangeStart, rangeEndExclusive) {
  const out = [];
  function addCal(calId, key) {
    const id = String(calId == null ? '' : calId).trim();
    if (!id) return;
    try {
      const cal = CalendarApp.getCalendarById(id);
      if (!cal) return;
      const evs = cal.getEvents(rangeStart, rangeEndExclusive);
      for (var i = 0; i < evs.length; i++) {
        const ev = evs[i];
        out.push({
          start: ev.getStartTime().getTime(),
          end: ev.getEndTime().getTime(),
          title: String(ev.getTitle() || '').trim() || '(No title)',
          allDay: ev.isAllDayEvent(),
          calendar: key,
        });
      }
    } catch (err) {
      Logger.log('collectCalendarWeekEvents_ ' + key + ' ' + err);
    }
  }
  addCal(CALENDAR_ID, 'studio');
  addCal(PERSONAL_CALENDAR_ID, 'personal');
  out.sort(function (a, b) {
    return a.start - b.start;
  });
  return out;
}

/** POST JSON flags: some proxies/clients send "true" as a string — strict === true misses and falls through to public booking → bad_date. */
function isJsonBoolTrue_(v) {
  return v === true || v === 'true' || v === 1 || v === '1';
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (isJsonBoolTrue_(d.ownerCalendarWeek)) {
      return handleOwnerCalendarWeek(d);
    }
    if (isJsonBoolTrue_(d.adminSetWorkHours)) {
      return handleAdminSetWorkHours(d);
    }
    if (isJsonBoolTrue_(d.ownerLookupClient)) {
      return handleOwnerClientLookup(d);
    }
    if (isJsonBoolTrue_(d.ownerDirectBooking)) {
      return handleOwnerDirectBooking(d);
    }
    if (isJsonBoolTrue_(d.ownerProposeAlternate)) {
      return handleOwnerProposeAlternate(d);
    }
    if (isJsonBoolTrue_(d.ownerRejectBooking)) {
      return handleOwnerRejectBooking(d);
    }
    if (isJsonBoolTrue_(d.reschedule)) {
      return handleReschedulePost(d);
    }
    const pubDateStr = d.date ? String(d.date).split('T')[0].trim() : '';
    if (!pubDateStr || !/^\d{4}-\d{2}-\d{2}$/.test(pubDateStr)) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'bad_date' })).setMimeType(ContentService.MimeType.JSON);
    }
    if (!isDateInPublicBookingWindow_(pubDateStr)) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'date_outside_booking_window' })).setMimeType(ContentService.MimeType.JSON);
    }
    if (!isDateMeetingBookingLeadTime_(pubDateStr)) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'date_too_soon' })).setMimeType(ContentService.MimeType.JSON);
    }
    const timeTrim = String(d.time || '').trim();
    if (!timeTrim) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'bad_time' })).setMimeType(ContentService.MimeType.JSON);
    }
    const start = parseYmdAndTimeLocal_(pubDateStr, convertTo24Hour(timeTrim));
    if (isNaN(start.getTime())) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'bad_datetime' })).setMimeType(ContentService.MimeType.JSON);
    }
    const durMin = clampDurationMinutes_(d.durationMinutes);
    if (!isBookingWithinStudioHours_(start, durMin)) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'invalid_day_or_time' })).setMimeType(ContentService.MimeType.JSON);
    }
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(25000)) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'server_busy' })).setMimeType(ContentService.MimeType.JSON);
    }
    try {
      const end = new Date(start.getTime() + durMin * 60000);
      if (slotOverlapsExistingCalendarEvents_(start, end, [])) {
        return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'slot_unavailable' })).setMimeType(ContentService.MimeType.JSON);
      }
      const actionToken = generateActionToken();
      const clientNotes = sanitizeClientNotes_(d.clientNotes);
      const ss = getCRMSpreadsheet();
      let s = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
      const row = s.getLastRow() + 1;
      s.appendRow([new Date(), d.clientName, d.phone, d.service, pubDateStr, timeTrim, d.email, 'PENDING', '', actionToken, '', '', '', '', clientNotes]);
      SpreadsheetApp.flush();
      const c = CalendarApp.getCalendarById(CALENDAR_ID);
      const desc =
        'Phone: ' +
        d.phone +
        '\nEmail: ' +
        d.email +
        '\nService: ' +
        d.service +
        '\nDurationMinutes: ' +
        durMin +
        (clientNotes ? '\nClient notes: ' + clientNotes : '');
      const ev = c.createEvent('PENDING: ' + d.clientName, start, end, { description: desc });
      ev.setColor(PENDING_COLOR);
      const tz = Session.getScriptTimeZone();
      const neatTimeEmail = Utilities.formatDate(start, tz, 'h:mm a');
      applyBookingLocationToEvent_(ev, neatTimeEmail, d.service);
      s.getRange(row, 9).setValue(ev.getId());
      const neatDate = Utilities.formatDate(start, tz, 'EEEE, MMMM d, yyyy');
      const requestHtml = getRequestEmailHtml(d.clientName, d.service, d.phone, d.email, neatDate, neatTimeEmail, ev.getId(), actionToken, clientNotes);
      MailApp.sendEmail({ to: MY_EMAIL, subject: "New Booking Request: " + d.clientName, htmlBody: requestHtml });

      return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    const msg = err && err.message ? String(err.message) : 'server_error';
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: msg })).setMimeType(ContentService.MimeType.JSON);
  }
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

/** Client email when the calendar event was removed (you deleted it); syncCalendarToSpreadsheet. */
function getCalendarDeletedClientEmailHtml(name, date, time, service) {
  const n = escapeHtml(name);
  const rows = emailDetailRow('Date', date) + emailDetailRow('Time', time) + emailDetailRow('Service', service);
  return (
    '<div style="font-family: sans-serif; padding: 40px; max-width: 500px; margin: auto; border: 1px solid #f0f0f0; border-radius: 12px; background-color: #ffffff;">' +
    '<p style="font-size: 16px; color: #1a1a1a; line-height: 1.6;">Hi ' +
    n +
    ',</p>' +
    '<p style="font-size: 16px; color: #1a1a1a; line-height: 1.6; margin-bottom: 20px;">Your appointment below has been <strong>cancelled</strong>.</p>' +
    '<table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:8px;margin:0 0 24px 0;" cellpadding="0" cellspacing="0" role="presentation"><tbody>' +
    rows +
    '</tbody></table>' +
    '<p style="font-size: 16px; color: #1a1a1a; line-height: 1.6;">If you have questions or would like to book another time, please reach out or use our website.</p>' +
    '<hr style="border: none; border-top: 1px solid #f0f0f0; margin: 40px 0;">' +
    '<p style="font-size: 13px; color: #999; text-align: center;">Roni\'s Nail Studio</p>' +
    '</div>'
  );
}

function getDeclinedEmailHtml(name, date, time, service, declineReasonNote) {
  const n = escapeHtml(name);
  const rows = emailDetailRow('Date', date) + emailDetailRow('Time', time) + emailDetailRow('Service', service);
  const noteTrim = String(declineReasonNote == null ? '' : declineReasonNote).trim();
  const reasonBlock = noteTrim
    ? '<p style="font-size: 15px; color: #1a1a1a; line-height: 1.65; margin: 0 0 20px 0; padding: 16px; background: #fafafa; border-radius: 8px; border-left: 3px solid #b76e7a;"><strong>Note from the studio:</strong><br><span style="white-space:pre-wrap;">' +
      escapeHtml(noteTrim) +
      '</span></p>'
    : '';
  return `
    <div style="font-family: sans-serif; padding: 40px; max-width: 500px; margin: auto; border: 1px solid #f0f0f0; border-radius: 12px; background-color: #ffffff;">
      <p style="font-size: 16px; color: #1a1a1a; line-height: 1.6;">Hi ${n},</p>
      <p style="font-size: 16px; color: #1a1a1a; line-height: 1.6; margin-bottom: 20px;">Thank you for your interest in booking with Roni's Nail Studio. Unfortunately, we're unable to accommodate this request. Details below:</p>
      ${reasonBlock}
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
 * Schedule: run installTwoDayReminderTrigger() once, or Triggers → Add trigger → time-driven → daily.
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

function getAlternateProposalClientEmailHtml(name, origDate, origTime, newDate, newTime, service, eventId, modToken) {
  const n = escapeHtml(name);
  const qOk = 'action=client_accept_mod&eventId=' + encodeURIComponent(eventId) + '&token=' + encodeURIComponent(modToken);
  const qNo = 'action=client_decline_mod&eventId=' + encodeURIComponent(eventId) + '&token=' + encodeURIComponent(modToken);
  const urlOk = buildBookingActionUrl(qOk);
  const urlNo = buildBookingActionUrl(qNo);
  const rows =
    emailDetailRow('Service', service) +
    emailDetailRow('You requested', origDate + ' · ' + origTime) +
    emailDetailRow('Suggested time', newDate + ' · ' + newTime);
  return (
    '<div style="font-family: sans-serif; padding: 40px; max-width: 500px; margin: auto; border: 1px solid #f0f0f0; border-radius: 12px; background: #fff;">' +
    '<p style="font-size: 16px; color: #1a1a1a;">Hi ' +
    n +
    ',</p>' +
    '<p style="font-size: 16px; color: #1a1a1a; line-height: 1.6;">Roni suggested a different time that may work better. If it works for you, confirm below. If not, you can always pick another slot on our website.</p>' +
    '<table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:8px;margin:20px 0;" cellpadding="0" cellspacing="0" role="presentation"><tbody>' +
    rows +
    '</tbody></table>' +
    '<div style="text-align: center;">' +
    '<a href="' +
    urlOk +
    '" style="background-color:#111;color:#fff;padding:14px;text-decoration:none;border-radius:8px;font-weight:500;display:block;margin-bottom:12px;">Yes, that works</a>' +
    '<a href="' +
    urlNo +
    '" style="background-color:#fff;color:#666;border:1px solid #ccc;padding:12px;text-decoration:none;border-radius:8px;display:block;">No thanks</a>' +
    '</div>' +
    '<p style="font-size:13px;color:#999;text-align:center;margin-top:24px;">Roni\'s Nail Studio</p></div>'
  );
}

function getAlternateDeclinedClientEmailHtml(name, origDate, origTime, service) {
  const n = escapeHtml(name);
  const rows = emailDetailRow('Your original request', origDate + ' · ' + origTime) + emailDetailRow('Service', service);
  return (
    '<div style="font-family: sans-serif; padding: 40px; max-width: 500px; margin: auto; border: 1px solid #f0f0f0; border-radius: 12px;">' +
    '<p style="font-size: 16px; color: #1a1a1a;">Hi ' +
    n +
    ',</p>' +
    '<p style="font-size: 16px; color: #1a1a1a; line-height: 1.6;">No problem — we won\'t hold that alternate time. Whenever you\'re ready, you can submit a new booking on our website.</p>' +
    '<table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:8px;margin:20px 0;" cellpadding="0" cellspacing="0" role="presentation"><tbody>' +
    rows +
    '</tbody></table>' +
    '<p style="font-size:13px;color:#999;text-align:center;margin-top:24px;">Roni\'s Nail Studio</p></div>'
  );
}

function getRequestEmailHtml(name, service, phone, email, date, time, eventId, actionToken, clientNotes) {
  const q = 'action=accept&eventId=' + encodeURIComponent(eventId) + '&token=' + encodeURIComponent(actionToken);
  const qr = 'action=reject&eventId=' + encodeURIComponent(eventId) + '&token=' + encodeURIComponent(actionToken);
  const acc = buildBookingActionUrl(q);
  const rejBase = String(OWNER_REJECT_PAGE_BASE || '').trim();
  const rej = rejBase
    ? rejBase + (rejBase.indexOf('?') >= 0 ? '&' : '?') + 'eventId=' + encodeURIComponent(eventId) + '&token=' + encodeURIComponent(actionToken)
    : buildBookingActionUrl(qr);
  const modUrl = OWNER_MODIFY_PAGE_BASE + '?eventId=' + encodeURIComponent(eventId) + '&token=' + encodeURIComponent(actionToken);
  const notesTrim = String(clientNotes == null ? '' : clientNotes).trim();
  const notesBlock = notesTrim
    ? '<p style="margin:12px 0 0 0;padding-top:12px;border-top:1px solid #e8e8e8;"><strong>Notes / requests:</strong><br><span style="white-space:pre-wrap;color:#333;">' +
      escapeHtml(notesTrim) +
      '</span></p>'
    : '';
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
        ${notesBlock}
      </div>
      <div style="text-align: center;">
        <a href="${acc}" style="background-color: #111; color: white; padding: 14px; text-decoration: none; border-radius: 8px; font-weight: 500; display: block; margin-bottom: 12px;">Approve Request</a>
        <a href="${modUrl}" style="background-color: #fff; color: #111; border: 1px solid #111; padding: 12px; text-decoration: none; border-radius: 8px; font-weight: 500; display: block; margin-bottom: 12px;">Suggest a different time</a>
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

  const reqHtml = getRequestEmailHtml(
    mockName,
    mockService,
    mockPhone,
    mockEmail,
    mockDate,
    mockTime,
    'test_event_id',
    'preview_only_invalid_token',
    'Sample client note (preview only).'
  );
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

/**
 * Build a local wall-clock Date (script project timezone behavior via JS Date parts).
 * Avoids new Date('yyyy-MM-ddTHH:mm:ss') which can be Invalid in some Apps Script cases.
 */
function parseYmdAndTimeLocal_(yyyyMmDd, hhMmSs) {
  const dp = String(yyyyMmDd).split('-');
  const tp = String(hhMmSs || '0:0:0').split(':');
  const y = parseInt(dp[0], 10);
  const mo = parseInt(dp[1], 10) - 1;
  const d = parseInt(dp[2], 10);
  const h = parseInt(tp[0], 10);
  const mi = parseInt(tp[1], 10) || 0;
  if (isNaN(y) || isNaN(mo) || isNaN(d) || isNaN(h) || isNaN(mi)) return new Date(NaN);
  return new Date(y, mo, d, h, mi, 0);
}

/**
 * True if start + duration fits inside studio hours for that calendar day (script project time zone).
 * Matches public booking rules; owner direct booking skips this check.
 */
function isBookingWithinStudioHours_(start, durationMinutes) {
  if (!(start instanceof Date) || isNaN(start.getTime())) return false;
  const dm = Number(durationMinutes);
  if (!isFinite(dm) || dm <= 0) return false;
  const whAll = getWorkHoursPayload_();
  const dow = start.getDay();
  const wh = whAll[String(dow)];
  if (!wh) return false;
  const y = start.getFullYear();
  const mo = start.getMonth();
  const day = start.getDate();
  const workStart = new Date(y, mo, day, wh.start, 0, 0);
  const workEnd = new Date(y, mo, day, wh.end, 0, 0);
  const slotEnd = new Date(start.getTime() + dm * 60000);
  return start >= workStart && slotEnd <= workEnd;
}

/**
 * True if [slotStart, slotEnd) overlaps any event on studio or personal calendars (same sources as public busy JSONP).
 * @param {Date} slotStart
 * @param {Date} slotEnd exclusive end instant
 * @param {Array<string>=} ignoreEventIds event IDs to ignore (e.g. appointment being moved)
 */
function slotOverlapsExistingCalendarEvents_(slotStart, slotEnd, ignoreEventIds) {
  const skip = Object.create(null);
  const ids = ignoreEventIds || [];
  for (var i = 0; i < ids.length; i++) {
    const id = String(ids[i] == null ? '' : ids[i]).trim();
    if (id) skip[id] = true;
  }
  if (!(slotStart instanceof Date) || !(slotEnd instanceof Date) || isNaN(slotStart.getTime()) || isNaN(slotEnd.getTime())) {
    return true;
  }
  if (slotEnd <= slotStart) {
    return true;
  }

  function overlapsOnCalendar(calId) {
    const cid = String(calId == null ? '' : calId).trim();
    if (!cid) return false;
    try {
      const cal = CalendarApp.getCalendarById(cid);
      if (!cal) return false;
      const evs = cal.getEvents(slotStart, slotEnd);
      for (var j = 0; j < evs.length; j++) {
        const ev = evs[j];
        if (skip[String(ev.getId())]) continue;
        const es = ev.getStartTime();
        const ee = ev.getEndTime();
        if (es < slotEnd && ee > slotStart) return true;
      }
    } catch (err) {
      Logger.log('slotOverlapsExistingCalendarEvents_ ' + cid + ' ' + err);
    }
    return false;
  }

  if (overlapsOnCalendar(CALENDAR_ID)) return true;
  if (overlapsOnCalendar(PERSONAL_CALENDAR_ID)) return true;
  return false;
}

function convertTo24Hour(timeStr) {
  const s = String(timeStr == null ? '' : timeStr)
    .trim()
    .replace(/\s+/g, ' ');
  if (!s) return '00:00:00';
  const bits = s.split(' ');
  const modifier = bits.length > 1 ? bits[bits.length - 1].toUpperCase() : '';
  const timePart = bits.length > 1 ? bits.slice(0, -1).join(' ') : bits[0];
  let [hours, minutes] = timePart.split(':');
  if (minutes === undefined) minutes = '00';
  else minutes = String(minutes).replace(/\D/g, '').slice(0, 2) || '00';
  hours = parseInt(hours, 10);
  minutes = parseInt(minutes, 10);
  if (isNaN(hours) || isNaN(minutes)) return '00:00:00';
  let h = hours;
  if (modifier === 'PM' && hours !== 12) h = hours + 12;
  if (modifier === 'AM' && hours === 12) h = 0;
  if (modifier !== 'AM' && modifier !== 'PM') h = hours;
  return `${String(h).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:00`;
}
