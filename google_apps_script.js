/**
 * --- Roni's Nail Studio: Google Apps Script V8 ---
 * Full merged version with:
 * - faster owner accept flow
 * - shorter lock waits on web actions
 * - cached public availability feed
 * - optional archiving of past appointments
 */

const CALENDAR_ID = 'f9c38dc209bf435115238aba24b24be51b7e4e2f05f3e3f9c08b9077a78c33b3@group.calendar.google.com';
const PERSONAL_CALENDAR_ID = 'nguyenveronica0108@gmail.com';
const SHEET_NAME = 'Bookings';
const ARCHIVE_SHEET_NAME = 'BookingsArchive';
const HOURS_SHEET_NAME = 'StudioHours';
const PENDING_COLOR = '5';
const CLIENT_CONFIRMED_CAL_PREFIX = '\u2727 ';
const MY_EMAIL = 'ronisnailstudio@gmail.com';
const TWO_DAY_REMINDER_EMAIL_SUBJECT = "Please confirm your appointment — Roni's Nail Studio";
const SPREADSHEET_ID = '16IJ_aJlAXWrF6UpM_g4Oia8rGGyAcmeR898STRX_5tc';
const PROP_DISABLE_CALENDAR_SYNC = 'DISABLE_CALENDAR_SYNC';
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzdT_rV3dR7Th4VHeLE3uJcyTPr4bI-6uy-_Im6xz-nZ0rGPToj85zy7Is7LmpNVS0Wwg/exec';
const BOOKING_ACTION_HTML_BASE = 'https://ronisnailstudio.com/api/booking';
const RESCHEDULE_PAGE_BASE = 'https://ronisnailstudio.com/reschedule.html';
const OWNER_MODIFY_PAGE_BASE = 'https://ronisnailstudio.com/owner-modify-request.html';
const OWNER_REJECT_PAGE_BASE = 'https://ronisnailstudio.com/reject-booking.html';
const SHEET_COL_OWNER_DECLINE_REASON = 16;
const ARCHIVE_AFTER_PAST_DAYS = 1;

function clientConfirmedCalendarEventTitle_(clientName) {
  const n = String(clientName == null ? '' : clientName).trim() || 'Client';
  return CLIENT_CONFIRMED_CAL_PREFIX + n;
}

function generateActionToken() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function sanitizeClientNotes_(raw) {
  let t = String(raw == null ? '' : raw).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  t = t.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, '');
  t = t.trim();
  if (t.length > 2000) t = t.substring(0, 2000);
  return t;
}

function sanitizeOwnerDeclineReason_(raw) {
  let t = String(raw == null ? '' : raw).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  t = t.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, '');
  t = t.trim();
  if (t.length > 1500) t = t.substring(0, 1500);
  return t;
}

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

function formatSheetTimeForEmail(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'h:mm a');
  }
  return String(value).trim();
}

function normalizeBulletSeparators_(serviceBlob) {
  return String(serviceBlob == null ? '' : serviceBlob)
    .replace(/\u2022/g, '\u00B7')
    .replace(/\r\n/g, '\n')
    .replace(/\s*\|\s*/g, ' \u00B7 ');
}

function splitServiceSegments_(serviceBlob) {
  const s = normalizeBulletSeparators_(serviceBlob).trim();
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
  if (low.indexOf('gel x removal') >= 0) return true;
  if (/^tier\s*\d+/i.test(t)) return true;
  return false;
}

function baseServiceLabelForCalendar_(serviceBlob) {
  const segments = splitServiceSegments_(serviceBlob);
  const bases = segments.filter(function (p) {
    return !isDesignTierOrSoakoffSegment_(p);
  });
  if (bases.length > 0) return bases.join(' + ');
  if (segments.length > 0) return segments.join(' + ');
  return '';
}

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
  const svcNorm = normalizeBulletSeparators_(serviceBlob);
  let loc = calendarLocationFromTimeAndService_(timeDisplay, svcNorm);
  const t = String(timeDisplay == null ? '' : timeDisplay).trim();
  const raw = String(svcNorm || '').trim();
  if (!loc) {
    if (t && raw) loc = t + ' ' + raw.substring(0, 140);
    else if (raw) loc = raw.substring(0, 200);
    else if (t) loc = t;
  }
  if (loc) ev.setLocation(String(loc).substring(0, 200));
  if (typeof ev.getLocation === 'function' && raw) {
    const check = String(ev.getLocation() || '').trim();
    if (!check || /^studio$/i.test(check)) {
      let rescue = calendarLocationFromTimeAndService_(t, svcNorm);
      if (!rescue || /^studio$/i.test(String(rescue).trim())) {
        rescue = (t ? t + ' ' : '') + raw.substring(0, 140);
      }
      if (String(rescue).trim() && !/^studio$/i.test(String(rescue).trim())) {
        ev.setLocation(String(rescue).trim().substring(0, 200));
      }
    }
  }
}

function emailDetailRow(label, value) {
  const v = escapeHtml(value || '—');
  const l = escapeHtml(label);
  return '<tr><td style="padding:10px 0;border-bottom:1px solid #eee;vertical-align:top;width:100px;"><span style="color:#666;font-size:12px;text-transform:uppercase;letter-spacing:0.06em;">' + l + '</span></td><td style="padding:10px 0;border-bottom:1px solid #eee;color:#1a1a1a;font-size:17px;font-weight:500;">' + v + '</td></tr>';
}

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
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

function clampDurationMinutes_(n) {
  const x = Number(n);
  if (isNaN(x)) return 60;
  return Math.max(15, Math.min(480, Math.round(x)));
}

function maxPublicBookingDateYmd_() {
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const p = todayStr.split('-');
  var cal = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10), 12, 0, 0);
  cal.setMonth(cal.getMonth() + 1);
  return Utilities.formatDate(cal, tz, 'yyyy-MM-dd');
}

function isDateInPublicBookingWindow_(yyyyMmDd) {
  const raw = String(yyyyMmDd == null ? '' : yyyyMmDd).trim().split('T')[0].split(' ')[0];
  const m = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return false;
  const maxD = maxPublicBookingDateYmd_();
  return raw <= maxD;
}

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

function effectiveDurationMinutesFromEvent_(ev) {
  const durMs = ev.getEndTime().getTime() - ev.getStartTime().getTime();
  const fromCal = Math.round(durMs / 60000);
  const fromDesc = parseDurationMinutesFromDescription_(ev.getDescription());
  return Math.max(30, Math.max(fromCal, fromDesc));
}

function syncSheetDateToYyyyMmDd_(value) {
  const tz = Session.getScriptTimeZone();
  if (value instanceof Date) return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  const s = String(value == null ? '' : value).trim();
  if (!s) return '';
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, tz, 'yyyy-MM-dd');
  return s.split('T')[0].split(' ')[0];
}

function normalizeTimeToken_(t) {
  return String(t == null ? '' : t).trim().replace(/\u00a0/g, ' ').replace(/\s+/g, '').toLowerCase();
}

function normalizeSheetStatus_(raw) {
  return String(raw == null ? '' : raw).trim().replace(/\s+/g, ' ').toUpperCase();
}

function calendarApiEventStatus_(eventId) {
  const want = String(eventId || '').trim();
  if (!want) return 'gone';
  try {
    if (typeof Calendar === 'undefined' || !Calendar.Events || typeof Calendar.Events.get !== 'function') return 'unknown';
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
    if (msg.indexOf('not found') >= 0 || msg.indexOf('404') >= 0 || msg.indexOf('requested entity was not found') >= 0) {
      return 'gone';
    }
    return 'unknown';
  }
}

function normalizeCalendarEventIdForCompare_(rawId) {
  var s = '';
  if (rawId != null) s = String(rawId);
  s = s.trim().toLowerCase();
  if (s.endsWith('@google.com')) s = s.slice(0, -'@google.com'.length);
  return s;
}

function calendarEventIdsMatch_(idA, idB) {
  const a = normalizeCalendarEventIdForCompare_(idA);
  const b = normalizeCalendarEventIdForCompare_(idB);
  if (!a || !b) return false;
  if (a === b) return true;
  if (b.length > a.length && b.slice(0, a.length + 1) === a + '_') return true;
  if (a.length > b.length && a.slice(0, b.length + 1) === b + '_') return true;
  return false;
}

function getActiveBookingEvent_(calendar, eventId, sheetStartHint) {
  const want = String(eventId || '').trim();
  if (!want) return null;
  let ev = null;
  try {
    ev = calendar.getEventById(want);
  } catch (err) {
    ev = null;
  }
  const api = calendarApiEventStatus_(want);
  if (api === 'active') {
    if (ev) return ev;
    Logger.log('getActiveBookingEvent_: Calendar API active but CalendarApp.getEventById returned null; id=' + String(want).substring(0, 56));
    return null;
  }
  if (api === 'gone' && !ev) return null;
  if (!ev) return null;
  if (api === 'unknown') {
    Logger.log(
      'getActiveBookingEvent_: API unknown - trusting getEventById (skip expensive listing verification). id=' +
        String(want).substring(0, 48)
    );
    return ev;
  }
  if (api === 'gone') {
    Logger.log('getActiveBookingEvent_: API reports gone/not found but CalendarApp has event; verifying with getEvents. id=' + String(want).substring(0, 56));
  }

  function listedInRange_(from, to) {
    try {
      const listed = calendar.getEvents(from, to);
      for (let i = 0; i < listed.length; i++) {
        if (calendarEventIdsMatch_(listed[i].getId(), want)) return true;
      }
    } catch (err2) {
      return false;
    }
    return false;
  }

  const center = ev.getStartTime();
  const hints = [];
  if (center && !isNaN(center.getTime())) hints.push(center);
  if (sheetStartHint && !isNaN(sheetStartHint.getTime())) hints.push(sheetStartHint);
  for (let hi = 0; hi < hints.length; hi++) {
    const c = hints[hi];
    let from = new Date(c.getTime() - 120 * 86400000);
    let to = new Date(c.getTime() + 120 * 86400000);
    if (listedInRange_(from, to)) return ev;
  }
  if (center && !isNaN(center.getTime())) {
    const from = new Date(center.getTime() - 400 * 86400000);
    const to = new Date(center.getTime() + 400 * 86400000);
    if (listedInRange_(from, to)) return ev;
  }
  Logger.log('getActiveBookingEvent_: getEventById ok but ID not found in getEvents listing. id=' + String(want).substring(0, 56));
  return null;
}

function appointmentRowStartMs_(dateVal, timeVal) {
  const ymd = syncSheetDateToYyyyMmDd_(dateVal);
  if (!ymd || !/^\d{4}-\d{2}-\d{2}$/.test(ymd)) return NaN;
  const timeDisp = formatSheetTimeForEmail(timeVal);
  if (!timeDisp) return NaN;
  const d = parseYmdAndTimeLocal_(ymd, convertTo24Hour(timeDisp));
  if (isNaN(d.getTime())) return NaN;
  return d.getTime();
}

function isCalendarToSheetSyncPaused_() {
  var v = String(PropertiesService.getScriptProperties().getProperty(PROP_DISABLE_CALENDAR_SYNC) || '').trim().toLowerCase();
  return v === '1' || v === 'true' || v === 'yes';
}

function pauseCalendarToSheetSync() {
  PropertiesService.getScriptProperties().setProperty(PROP_DISABLE_CALENDAR_SYNC, '1');
  Logger.log('Calendar->sheet LIVE sync is PAUSED. Triggers will no-op until resumeCalendarToSheetSync(). Dry run still runs.');
}

function resumeCalendarToSheetSync() {
  PropertiesService.getScriptProperties().deleteProperty(PROP_DISABLE_CALENDAR_SYNC);
  Logger.log('Calendar->sheet live sync RESUMED.');
}

function tokenMatches(stored, provided) {
  if (stored === undefined || stored === null || provided === undefined || provided === null) return false;
  return String(stored).trim() === String(provided).trim();
}

function defaultWorkHoursObject_() {
  return { '1': { start: 11, end: 18 }, '2': { start: 11, end: 18 }, '3': { start: 9, end: 16 }, '5': { start: 9, end: 16 } };
}

function sheetCellToScheduleKey_(raw) {
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return { kind: 'date', ymd: Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd') };
  }
  const s = String(raw == null ? '' : raw).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return { kind: 'date', ymd: s };
  const d = parseInt(s, 10);
  if (!isNaN(d) && d >= 0 && d <= 6) return { kind: 'dow', dow: d };
  return { kind: 'skip' };
}

function getWorkHoursFromSheet_() {
  try {
    const ss = getCRMSpreadsheet();
    const sh = ss.getSheetByName(HOURS_SHEET_NAME);
    if (!sh) return null;
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return null;
    const weekly = {};
    const dateOverrides = {};
    for (let i = 1; i < data.length; i++) {
      const rawD = data[i][0];
      if (rawD === '' || rawD === null || rawD === undefined) continue;
      const st = Number(data[i][1]);
      const en = Number(data[i][2]);
      if (!isFinite(st) || !isFinite(en) || st !== Math.floor(st) || en !== Math.floor(en)) continue;
      if (st < 0 || st > 23 || en < 1 || en > 24 || st >= en) continue;
      const key = sheetCellToScheduleKey_(rawD);
      if (key.kind === 'dow') weekly[String(key.dow)] = { start: st, end: en };
      else if (key.kind === 'date') dateOverrides[key.ymd] = { start: st, end: en };
    }
    if (!Object.keys(weekly).length && !Object.keys(dateOverrides).length) return null;
    return { weekly: weekly, dateOverrides: dateOverrides };
  } catch (err) {
    return null;
  }
}

function getWorkHoursPayload_() {
  const fromSheet = getWorkHoursFromSheet_();
  if (fromSheet) {
    return { weekly: fromSheet.weekly, dateOverrides: fromSheet.dateOverrides || {} };
  }
  return { weekly: defaultWorkHoursObject_(), dateOverrides: {} };
}
/**
 * Merge guide
 * 1. Add the new constants and helper functions from this file near the top of your Apps Script.
 * 2. Replace the matching functions in your existing script with the versions in this file.
 * 3. Deploy a new /exec version after pasting.
 *
 * Goals of this patch:
 * - Keep the normal sheet / calendar / email behavior.
 * - Stop web requests from waiting 15s-120s on script locks.
 * - Make owner "Accept" use a fast GET page plus a retried POST action.
 * - Cache the public availability feeds briefly and only scan the real booking window.
 */

const WEB_ACTION_LOCK_TIMEOUT_MS = 5000;
const OWNER_ACCEPT_LOCK_TIMEOUT_MS = 8000;
const SYNC_LOCK_TIMEOUT_MS = 5000;
const PUBLIC_BUSY_CACHE_TTL_SECONDS = 30;
const WORK_HOURS_CACHE_TTL_SECONDS = 60;
const SYNC_PAST_CONFIRMED_LOOKBACK_DAYS = 1;
const PROP_SYNC_SUPPRESS_UNTIL_MS = 'SYNC_SUPPRESS_UNTIL_MS';
const INTERNAL_CALENDAR_SYNC_SUPPRESS_MS = 15000;
const SYNC_TIME_BUDGET_MS = 12000;
const SYNC_DRY_RUN_TIME_BUDGET_MS = 20000;

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function markInternalCalendarMutation_() {
  try {
    const untilMs = Date.now() + INTERNAL_CALENDAR_SYNC_SUPPRESS_MS;
    PropertiesService.getScriptProperties().setProperty(PROP_SYNC_SUPPRESS_UNTIL_MS, String(untilMs));
  } catch (err) {
    Logger.log('markInternalCalendarMutation_: ' + err);
  }
}

function isInternalCalendarSyncSuppressedNow_() {
  try {
    const raw = String(PropertiesService.getScriptProperties().getProperty(PROP_SYNC_SUPPRESS_UNTIL_MS) || '').trim();
    if (!raw) return false;
    const untilMs = parseInt(raw, 10);
    if (!isFinite(untilMs) || untilMs <= 0) {
      PropertiesService.getScriptProperties().deleteProperty(PROP_SYNC_SUPPRESS_UNTIL_MS);
      return false;
    }
    if (Date.now() >= untilMs) {
      PropertiesService.getScriptProperties().deleteProperty(PROP_SYNC_SUPPRESS_UNTIL_MS);
      return false;
    }
    return true;
  } catch (err) {
    Logger.log('isInternalCalendarSyncSuppressedNow_: ' + err);
    return false;
  }
}

function clearBookingEndpointCaches_() {
  try {
    CacheService.getScriptCache().removeAll([
      'booking:workHours:v1',
      'booking:busy:v2:-',
    ]);
  } catch (err) {
    Logger.log('clearBookingEndpointCaches_: ' + err);
  }
}

function cachedWorkHoursPayloadJson_() {
  const cache = CacheService.getScriptCache();
  const key = 'booking:workHours:v1';
  const hit = cache.get(key);
  if (hit) return hit;
  const json = JSON.stringify(getWorkHoursPayload_());
  cache.put(key, json, WORK_HOURS_CACHE_TTL_SECONDS);
  return json;
}

function publicBusyRangeEnd_() {
  const maxYmd = maxPublicBookingDateYmd_();
  const lastBookable = parseYmdAndTimeLocal_(maxYmd, '12:00:00');
  if (isNaN(lastBookable.getTime())) {
    const fallback = new Date();
    fallback.setMonth(fallback.getMonth() + 1);
    return fallback;
  }
  return new Date(
    lastBookable.getFullYear(),
    lastBookable.getMonth(),
    lastBookable.getDate() + 1,
    0,
    0,
    0
  );
}

function buildPublicBusyArray_(ignoreEventId) {
  const busCal = CalendarApp.getCalendarById(CALENDAR_ID);
  const perCal = CalendarApp.getCalendarById(PERSONAL_CALENDAR_ID);
  const now = new Date();
  const rangeEnd = publicBusyRangeEnd_();
  const ignore = String(ignoreEventId || '').trim();
  const allBusy = [];

  function pushBusy_(cal) {
    if (!cal) return;
    cal.getEvents(now, rangeEnd).forEach(function (v) {
      if (ignore && String(v.getId()) === ignore) return;
      allBusy.push({
        start: v.getStartTime().toISOString(),
        end: v.getEndTime().toISOString(),
      });
    });
  }

  pushBusy_(busCal);
  pushBusy_(perCal);
  return allBusy;
}

function cachedPublicBusyJson_(ignoreEventId) {
  const ignore = String(ignoreEventId || '').trim();
  if (ignore) {
    return JSON.stringify(buildPublicBusyArray_(ignore));
  }
  const cache = CacheService.getScriptCache();
  const key = 'booking:busy:v2:-';
  const hit = cache.get(key);
  if (hit) return hit;
  const json = JSON.stringify(buildPublicBusyArray_(''));
  cache.put(key, json, PUBLIC_BUSY_CACHE_TTL_SECONDS);
  return json;
}

function handleWorkHoursGet_(e) {
  const callback = e.parameter.callback;
  const json = cachedWorkHoursPayloadJson_();
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function handlePublicBusyFeedGet_(e) {
  const callback = e.parameter.callback;
  const ignoreEventId = e.parameter.ignoreEventId ? String(e.parameter.ignoreEventId) : '';
  const json = cachedPublicBusyJson_(ignoreEventId);
  return ContentService.createTextOutput(callback ? callback + '(' + json + ')' : json).setMimeType(
    callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON
  );
}

function ownerAcceptProcessingPage_(eventId, token) {
  const eventIdJson = JSON.stringify(String(eventId || ''));
  const tokenJson = JSON.stringify(String(token || ''));
  const bodyHtml = `
    <div style="max-width:480px;margin:0 auto;">
      <h2 style="font-weight:600;margin:0 0 12px 0;">Approving request</h2>
      <p style="color:#555;margin:0 0 20px 0;">Please keep this page open while we update the booking, calendar, and client email.</p>
      <div id="msg" style="padding:16px;border:1px solid #eee;border-radius:12px;background:#fafafa;">Working...</div>
    </div>
    <script>
      (function () {
        var eventId = ${eventIdJson};
        var token = ${tokenJson};
        var msg = document.getElementById('msg');

        function setBox(kind, text) {
          var styles = {
            info: 'background:#fafafa;border:1px solid #eee;color:#222;',
            ok: 'background:rgba(100,140,100,0.12);border:1px solid rgba(100,140,100,0.25);color:#2d5a2d;',
            err: 'background:rgba(183,94,122,0.10);border:1px solid rgba(183,94,122,0.25);color:#8b3a52;'
          };
          msg.style.cssText = 'padding:16px;border-radius:12px;' + (styles[kind] || styles.info);
          msg.textContent = text;
        }

        function friendly(code, warning) {
          if (warning === 'email_failed') {
            return 'The request was accepted and the calendar was updated, but the confirmation email could not be sent automatically. Please contact the client manually.';
          }
          if (code === 'already_proposed') {
            return 'You already suggested a different time for this request. Wait for the client to respond, or decline the request instead.';
          }
          if (code === 'already_handled') {
            return 'This request was already approved or declined.';
          }
          if (code === 'invalid_token') {
            return 'This approval link is invalid or expired. Open the latest booking request email and try again.';
          }
          if (code === 'not_found') {
            return 'We could not find this request. It may already have been removed.';
          }
          if (code === 'event_missing') {
            return 'Google Calendar no longer has this pending hold. Refresh your booking sheet before trying again.';
          }
          if (code === 'server_busy') {
            return 'The booking system is busy right now. Retrying automatically in 8 seconds...';
          }
          return 'Something went wrong. Please try again in a moment.';
        }

        function submitOwnerAccept(attempt) {
          var n = typeof attempt === 'number' ? attempt : 0;
          setBox('info', n > 0 ? 'Retrying approval... (attempt ' + (n + 1) + ' of 6)' : 'Approving request...');
          fetch(window.location.origin + '/api/booking', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              ownerAcceptBooking: true,
              eventId: eventId,
              ownerToken: token
            })
          })
            .then(function (r) {
              return r.text().then(function (t) {
                var j = {};
                try {
                  j = JSON.parse(String(t || '').trim().replace(/^\\uFEFF/, ''));
                } catch (e) {
                  j = { _parseError: true, _raw: t };
                }
                return { ok: r.ok, status: r.status, j: j };
              });
            })
            .then(function (x) {
              var status = x.j && x.j.status;
              var warning = x.j && x.j.warning;
              if (x.ok && status === 'success') {
                setBox('ok', friendly('', warning) || 'Request accepted. The client has been notified.');
                if (warning !== 'email_failed') {
                  setBox('ok', 'Request accepted. The client has been notified.');
                }
                return;
              }
              var code = (x.j && x.j.message) || '';
              if (code === 'server_busy' && n < 5) {
                setBox('info', friendly(code) + ' (attempt ' + (n + 1) + ' of 6)');
                window.setTimeout(function () {
                  submitOwnerAccept(n + 1);
                }, 8000);
                return;
              }
              setBox('err', friendly(code));
            })
            .catch(function () {
              if (n < 2) {
                setBox('info', 'Connection was interrupted. Retrying automatically...');
                window.setTimeout(function () {
                  submitOwnerAccept(n + 1);
                }, 5000);
                return;
              }
              setBox(
                'err',
                'We could not reach the booking server. If the client still received the confirmation email, the approval may already have succeeded, so please check the sheet before trying again.'
              );
            });
        }

        submitOwnerAccept(0);
      })();
    </script>
  `;
  return htmlPage('Approve request', bodyHtml);
}

function handleOwnerActionGet_(e) {
  const action = String(e.parameter.action || '').trim();
  const eventId = String(e.parameter.eventId || '').trim();
  const token = String(e.parameter.token || '').trim();

  if (!eventId || !token) {
    return htmlPage(
      'Invalid link',
      '<h2>Invalid link</h2><p>This approval link is incomplete. Open the message from your latest booking email.</p>'
    );
  }

  if (action === 'reject') {
    const rb = String(OWNER_REJECT_PAGE_BASE || '').trim();
    if (rb) {
      const sep = rb.indexOf('?') >= 0 ? '&' : '?';
      const target =
        rb +
        sep +
        'eventId=' +
        encodeURIComponent(eventId) +
        '&token=' +
        encodeURIComponent(token);
      const html =
        '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Decline request</title></head><body><p>Loading...</p><script>location.replace(' +
        JSON.stringify(target) +
        ');</script><p style="font-family:system-ui,sans-serif;text-align:center;padding:2rem"><a href=' +
        JSON.stringify(target) +
        '>Continue</a></p></body></html>';
      return ContentService.createTextOutput(html).setMimeType(ContentService.MimeType.HTML);
    }
    return htmlPage(
      'Decline unavailable',
      '<h2>Decline page not configured</h2><p>Please use the latest reject link from your booking email.</p>'
    );
  }

  if (action === 'accept') {
    return ownerAcceptProcessingPage_(eventId, token);
  }

  return htmlPage('Not supported', '<h2>Unsupported action</h2>');
}

function findBookingRowIndexByEventId_(sheet, eventId) {
  const want = String(eventId || '').trim();
  if (!sheet || !want) return -1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  try {
    const finder = sheet.getRange(2, 9, lastRow - 1, 1).createTextFinder(want).matchEntireCell(true);
    const match = finder.findNext();
    if (match) return match.getRow();
  } catch (err) {
    Logger.log('findBookingRowIndexByEventId_: text finder failed: ' + err);
  }

  try {
    const values = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0] || '').trim() === want) return i + 2;
    }
  } catch (fallbackErr) {
    Logger.log('findBookingRowIndexByEventId_: fallback scan failed: ' + fallbackErr);
  }
  return -1;
}

function getBookingRowSnapshot_(sheet, rowIndex) {
  if (!sheet || rowIndex < 2) return null;
  const width = Math.max(SHEET_COL_OWNER_DECLINE_REASON, 15);
  const row = sheet.getRange(rowIndex, 1, 1, width).getValues()[0];
  return {
    rowIndex: rowIndex,
    clientName: row[1],
    phone: row[2],
    service: row[3],
    dateVal: row[4],
    timeStr: row[5],
    clientEmail: row[6],
    rowStatus: normalizeSheetStatus_(row[7]),
    eventId: String(row[8] || '').trim(),
    actionToken: String(row[9] || '').trim(),
  };
}

function handleOwnerAcceptBooking(d) {
  const eventId = String(d.eventId || '').trim();
  const token = String(d.ownerToken != null && d.ownerToken !== '' ? d.ownerToken : d.token || '').trim();
  if (!eventId || !token) {
    return jsonResponse_({ status: 'error', message: 'missing_fields' });
  }

  const ss = getCRMSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const rowIndex = findBookingRowIndexByEventId_(sheet, eventId);
  if (rowIndex < 0) {
    return jsonResponse_({ status: 'error', message: 'not_found' });
  }
  const snapshot = getBookingRowSnapshot_(sheet, rowIndex);
  if (!snapshot) {
    return jsonResponse_({ status: 'error', message: 'not_found' });
  }
  if (!tokenMatches(snapshot.actionToken, token)) {
    return jsonResponse_({ status: 'error', message: 'invalid_token' });
  }
  const clientName = String(snapshot.clientName || '').trim() || 'Client';
  const clientEmail = String(snapshot.clientEmail || '').trim();
  const service = snapshot.service;
  const dateVal = snapshot.dateVal;
  const timeStr = snapshot.timeStr;
  const rowStatus = snapshot.rowStatus;

  if (rowStatus === 'MOD_PROPOSED') {
    return jsonResponse_({ status: 'error', message: 'already_proposed' });
  }
  if (rowStatus !== 'PENDING') {
    return jsonResponse_({ status: 'error', message: 'already_handled' });
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(OWNER_ACCEPT_LOCK_TIMEOUT_MS)) {
    return jsonResponse_({ status: 'error', message: 'server_busy' });
  }

  try {
    const statusFresh = normalizeSheetStatus_(sheet.getRange(rowIndex, 8).getValue());
    if (statusFresh === 'MOD_PROPOSED') {
      return jsonResponse_({ status: 'error', message: 'already_proposed' });
    }
    if (statusFresh !== 'PENDING') {
      return jsonResponse_({ status: 'error', message: 'already_handled' });
    }

    const currentEventId = String(sheet.getRange(rowIndex, 9).getValue() || eventId).trim() || eventId;
    const cal = CalendarApp.getCalendarById(CALENDAR_ID);
    const ev = cal.getEventById(currentEventId);
    if (!ev) {
      Logger.log('owner accept: pending calendar event not found; not marking CONFIRMED (eventId=' + currentEventId + ')');
      return jsonResponse_({ status: 'error', message: 'event_missing' });
    }

    markInternalCalendarMutation_();
    const nEv = cal.createEvent(clientName, ev.getStartTime(), ev.getEndTime(), {
      description: ev.getDescription(),
    });
    applyBookingLocationToEvent_(
      nEv,
      Utilities.formatDate(nEv.getStartTime(), Session.getScriptTimeZone(), 'h:mm a'),
      service
    );

    sheet.getRange(rowIndex, 9).setValue(nEv.getId());
    sheet.getRange(rowIndex, 8).setValue('CONFIRMED');
    SpreadsheetApp.flush();

    try {
      ev.deleteEvent();
    } catch (deleteErr) {
      Logger.log('owner accept: could not delete pending hold after confirm: ' + deleteErr);
    }
  } finally {
    lock.releaseLock();
  }

  let warning = '';
  if (clientEmail) {
    try {
      const neatD = formatSheetDateForEmail(dateVal);
      const neatTime = formatSheetTimeForEmail(timeStr);
      const acceptEmailHtml = getConfirmedEmailHtml(clientName, neatD, neatTime, service);
      MailApp.sendEmail({
        to: clientEmail,
        name: "Roni's Nail Studio",
        subject: "Appointment Confirmed: Roni's Nail Studio",
        htmlBody: acceptEmailHtml,
      });
    } catch (mailErr) {
      warning = 'email_failed';
      Logger.log('owner accept: confirmation email failed: ' + mailErr);
    }
  }

  return jsonResponse_({ status: 'success', warning: warning });
}

function handlePublicBookingPost_(d) {
  const pubDateStr = d.date ? String(d.date).split('T')[0].trim() : '';
  if (!pubDateStr || !/^\d{4}-\d{2}-\d{2}$/.test(pubDateStr)) {
    return jsonResponse_({ status: 'error', message: 'bad_date' });
  }
  if (!isDateInPublicBookingWindow_(pubDateStr)) {
    return jsonResponse_({ status: 'error', message: 'date_outside_booking_window' });
  }
  if (!isDateMeetingBookingLeadTime_(pubDateStr)) {
    return jsonResponse_({ status: 'error', message: 'date_too_soon' });
  }

  const timeTrim = String(d.time || '').trim();
  if (!timeTrim) {
    return jsonResponse_({ status: 'error', message: 'bad_time' });
  }

  const start = parseYmdAndTimeLocal_(pubDateStr, convertTo24Hour(timeTrim));
  if (isNaN(start.getTime())) {
    return jsonResponse_({ status: 'error', message: 'bad_datetime' });
  }

  const durMin = clampDurationMinutes_(d.durationMinutes);
  if (!isBookingWithinStudioHours_(start, durMin)) {
    return jsonResponse_({ status: 'error', message: 'invalid_day_or_time' });
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(WEB_ACTION_LOCK_TIMEOUT_MS)) {
    return jsonResponse_({ status: 'error', message: 'server_busy' });
  }

  try {
    const end = new Date(start.getTime() + durMin * 60000);
    if (slotOverlapsExistingCalendarEvents_(start, end, [])) {
      return jsonResponse_({ status: 'error', message: 'slot_unavailable' });
    }

    const actionToken = generateActionToken();
    const clientNotes = sanitizeClientNotes_(d.clientNotes);
    const ss = getCRMSpreadsheet();
    const s = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    const row = s.getLastRow() + 1;

    s.appendRow([
      new Date(),
      d.clientName,
      d.phone,
      d.service,
      pubDateStr,
      timeTrim,
      d.email,
      'PENDING',
      '',
      actionToken,
      '',
      '',
      '',
      '',
      clientNotes,
    ]);
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
    markInternalCalendarMutation_();
    const ev = c.createEvent('PENDING: ' + d.clientName, start, end, { description: desc });
    ev.setColor(PENDING_COLOR);

    const tz = Session.getScriptTimeZone();
    const neatTimeEmail = Utilities.formatDate(start, tz, 'h:mm a');
    applyBookingLocationToEvent_(ev, neatTimeEmail, d.service);
    s.getRange(row, 9).setValue(ev.getId());
    SpreadsheetApp.flush();
    clearBookingEndpointCaches_();

    const neatDate = Utilities.formatDate(start, tz, 'EEEE, MMMM d, yyyy');
    const requestHtml = getRequestEmailHtml(
      d.clientName,
      d.service,
      d.phone,
      d.email,
      neatDate,
      neatTimeEmail,
      ev.getId(),
      actionToken,
      clientNotes
    );
    MailApp.sendEmail({
      to: MY_EMAIL,
      subject: 'New Booking Request: ' + d.clientName,
      htmlBody: requestHtml,
    });

    return jsonResponse_({ status: 'success' });
  } finally {
    lock.releaseLock();
  }
}

function syncCalendarToSpreadsheet() {
  if (isCalendarToSheetSyncPaused_()) {
    Logger.log(
      'syncCalendarToSpreadsheet: skipped (pause is ON - run resumeCalendarToSheetSync when safe, or was left on after restore)'
    );
    return;
  }
  if (isInternalCalendarSyncSuppressedNow_()) {
    Logger.log('syncCalendarToSpreadsheet: skipped (recent internal calendar mutation - avoiding self-trigger storm)');
    return;
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(SYNC_LOCK_TIMEOUT_MS)) {
    Logger.log('syncCalendarToSpreadsheet: skipped (another action is running - avoids web timeouts and duplicate emails)');
    return;
  }
  try {
    syncCalendarToSpreadsheetBody_(false);
  } finally {
    lock.releaseLock();
  }
}

function syncCalendarToSpreadsheetDryRun() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(SYNC_LOCK_TIMEOUT_MS)) {
    Logger.log('syncCalendarToSpreadsheetDryRun: skipped (another sync or web action is running)');
    return;
  }
  try {
    syncCalendarToSpreadsheetBody_(true);
  } finally {
    lock.releaseLock();
  }
}

function syncCalendarToSpreadsheetBody_(dryRun) {
  dryRun = dryRun === true;
  const startedAtMs = Date.now();
  const timeBudgetMs = dryRun ? SYNC_DRY_RUN_TIME_BUDGET_MS : SYNC_TIME_BUDGET_MS;
  const tz = Session.getScriptTimeZone();
  const ss = getCRMSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const pastConfirmedCutoffMs = Date.now() - SYNC_PAST_CONFIRMED_LOOKBACK_DAYS * 86400000;
  let qualifying = 0;
  let stillOnCalendar = 0;
  let markedCancelled = 0;
  let timeUpdated = 0;
  let dryRunWouldEmailCancel = 0;
  let dryRunWouldEmailTimeUpdate = 0;
  let changedAvailability = false;

  Logger.log(
    'syncCalendarToSpreadsheet: START ' +
      (dryRun ? '[DRY RUN - no sheet/email/calendar edits] ' : '') +
      'sheet=' +
      SHEET_NAME +
      ' dataRows=' +
      (data.length - 1) +
      ' calendarId=' +
      CALENDAR_ID
  );

  for (let i = 1; i < data.length; i++) {
    if (Date.now() - startedAtMs > timeBudgetMs) {
      Logger.log(
        'syncCalendarToSpreadsheet: stopping early after ' +
          (Date.now() - startedAtMs) +
          ' ms to avoid long lock contention; rowsProcessed=' +
          (i - 1)
      );
      break;
    }
    try {
      const status = normalizeSheetStatus_(data[i][7]);
      const eventId = data[i][8];
      if (
        (status === 'CONFIRMED' || status === 'CLIENT_CONFIRMED' || status === 'PENDING' || status === 'MOD_PROPOSED') &&
        eventId &&
        String(eventId).indexOf('pending') < 0
      ) {
        qualifying++;
        const sheetStartMs = appointmentRowStartMs_(data[i][4], data[i][5]);
        if (
          (status === 'CONFIRMED' || status === 'CLIENT_CONFIRMED') &&
          !isNaN(sheetStartMs) &&
          sheetStartMs < pastConfirmedCutoffMs
        ) {
          continue;
        }

        const sheetStartHint = isNaN(sheetStartMs) ? null : new Date(sheetStartMs);
        let event = getActiveBookingEvent_(calendar, eventId, sheetStartHint);
        if (!event) {
          const statusFresh = normalizeSheetStatus_(sheet.getRange(i + 1, 8).getValue());
          if (statusFresh === 'CANCELLED') {
            Logger.log('syncCalendarToSpreadsheet: row ' + (i + 1) + ' already CANCELLED - skip (no duplicate email)');
            continue;
          }
          if (
            (status === 'PENDING' || status === 'MOD_PROPOSED') &&
            (statusFresh === 'CONFIRMED' || statusFresh === 'CLIENT_CONFIRMED')
          ) {
            Logger.log(
              'syncCalendarToSpreadsheet: row ' +
                (i + 1) +
                ' skip cancel - sheet moved to ' +
                statusFresh +
                ' during sync (snapshot was ' +
                status +
                ')'
            );
            continue;
          }
          const eventIdFresh = String(sheet.getRange(i + 1, 9).getValue() || '').trim();
          if (eventIdFresh && eventIdFresh !== String(eventId).trim()) {
            event = getActiveBookingEvent_(calendar, eventIdFresh, sheetStartHint);
            if (event) {
              Logger.log(
                'syncCalendarToSpreadsheet: row ' +
                  (i + 1) +
                  ' skip cancel - event id was updated mid-sync (new id still on calendar)'
              );
            }
          }
        }

        if (!event) {
          const apiStillActive = calendarApiEventStatus_(String(eventId).trim());
          if (apiStillActive === 'active') {
            Logger.log(
              'syncCalendarToSpreadsheet: row ' +
                (i + 1) +
                ' skip cancel - Calendar API still shows event active (CalendarApp lookup failed; avoid false mass-cancel)'
            );
            stillOnCalendar++;
            continue;
          }

          markedCancelled++;
          changedAvailability = true;
          const clientName = data[i][1];
          const clientEmail = data[i][6];
          const service = data[i][3];
          const neatD = formatSheetDateForEmail(data[i][4]);
          const neatTime = formatSheetTimeForEmail(data[i][5]);
          const startMs = sheetStartMs;
          const apptAlreadyPassed = !isNaN(startMs) && startMs < Date.now();

          Logger.log(
            (dryRun ? '[DRY RUN] ' : '') +
              'syncCalendarToSpreadsheet: row ' +
              (i + 1) +
              ' no active event -> ' +
              (dryRun ? 'WOULD set CANCELLED' : 'CANCELLED') +
              ' (status was ' +
              status +
              ') eventId=' +
              String(eventId).substring(0, 40) +
              '...'
          );

          if (dryRun) {
            if (clientEmail && !apptAlreadyPassed) {
              dryRunWouldEmailCancel++;
              Logger.log(
                '[DRY RUN] row ' +
                  (i + 1) +
                  ' WOULD send cancellation email to ' +
                  String(clientEmail) +
                  ' (not sent)'
              );
            } else if (clientEmail && apptAlreadyPassed) {
              Logger.log('[DRY RUN] row ' + (i + 1) + ' would set CANCELLED only; no email (appointment in the past)');
            } else {
              Logger.log('[DRY RUN] row ' + (i + 1) + ' would set CANCELLED; no client email on row');
            }
            continue;
          }

          sheet.getRange(i + 1, 8).setValue('CANCELLED');
          sheet.getRange(i + 1, 11).setValue('');
          SpreadsheetApp.flush();
          if (clientEmail && !apptAlreadyPassed) {
            const html = getCalendarDeletedClientEmailHtml(clientName, neatD, neatTime, service);
            MailApp.sendEmail({
              to: clientEmail,
              name: "Roni's Nail Studio",
              subject: 'Appointment Cancelled: Roni\'s Nail Studio',
              htmlBody: html,
            });
            Logger.log('syncCalendarToSpreadsheet: cancellation email sent for row ' + (i + 1) + ' eventId=' + eventId);
          } else if (clientEmail && apptAlreadyPassed) {
            Logger.log(
              'syncCalendarToSpreadsheet: row ' +
                (i + 1) +
                ' marked CANCELLED; no client email (appointment start already in the past)'
            );
          } else {
            Logger.log('syncCalendarToSpreadsheet: row ' + (i + 1) + ' event missing but no client email - no email sent');
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
          changedAvailability = true;
          const calTimeDisplay = Utilities.formatDate(calStart, tz, 'h:mm a');
          if (dryRun) {
            const clientEmailUp = data[i][6];
            if (clientEmailUp) dryRunWouldEmailTimeUpdate++;
            Logger.log(
              '[DRY RUN] row ' +
                (i + 1) +
                ' WOULD update sheet date/time from calendar to ' +
                calDateYmd +
                ' ' +
                calTimeDisplay +
                '; ' +
                (clientEmailUp ? 'WOULD email time-updated to ' + String(clientEmailUp) : 'no client email') +
                '; WOULD applyBookingLocationToEvent_ (not run)'
            );
            continue;
          }

          sheet.getRange(i + 1, 5).setValue(calStart);
          sheet.getRange(i + 1, 6).setValue(calTimeDisplay);
          sheet.getRange(i + 1, 11).setValue('');
          SpreadsheetApp.flush();
          const neatDate = Utilities.formatDate(calStart, tz, 'EEEE, MMMM d, yyyy');
          const clientEmail = data[i][6];
          const service = data[i][3];
          markInternalCalendarMutation_();
          applyBookingLocationToEvent_(event, calTimeDisplay, service);
          if (clientEmail) {
            const html = getCalendarUpdatedClientEmailHtml(data[i][1], neatDate, calTimeDisplay, service);
            MailApp.sendEmail({
              to: clientEmail,
              name: "Roni's Nail Studio",
              subject: 'Appointment time updated - Roni\'s Nail Studio',
              htmlBody: html,
            });
          }
        }
      }
    } catch (rowErr) {
      Logger.log('syncCalendarToSpreadsheet row ' + (i + 1) + ' skipped: ' + rowErr);
    }
  }

  if (!dryRun && changedAvailability) {
    clearBookingEndpointCaches_();
  }

  Logger.log(
    'syncCalendarToSpreadsheet: DONE ' +
      (dryRun ? '[DRY RUN - no changes applied] ' : '') +
      'qualifying=' +
      qualifying +
      ' stillOnCalendar=' +
      stillOnCalendar +
      ' markedCancelled=' +
      markedCancelled +
      ' timeUpdated=' +
      timeUpdated +
      (dryRun
        ? ' dryRunWouldEmailCancel=' + dryRunWouldEmailCancel + ' dryRunWouldEmailTimeUpdate=' + dryRunWouldEmailTimeUpdate
        : '')
  );
}

function handleAdminSetWorkHours(d) {
  const secret = PropertiesService.getScriptProperties().getProperty('BOOKING_ADMIN_SECRET');
  if (!secret || String(d.adminSecret || '').trim() !== String(secret).trim()) {
    return jsonResponse_({ status: 'error', message: 'unauthorized' });
  }
  if (!d.hours || typeof d.hours !== 'object' || Array.isArray(d.hours)) {
    return jsonResponse_({ status: 'error', message: 'bad_hours' });
  }

  const overrideObj =
    d.dateOverrides && typeof d.dateOverrides === 'object' && !Array.isArray(d.dateOverrides) ? d.dateOverrides : {};
  const ss = getCRMSpreadsheet();
  let sh = ss.getSheetByName(HOURS_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(HOURS_SHEET_NAME);
  sh.clear();
  sh.appendRow(['day', 'start', 'end']);

  const rows = [];
  let anyOpen = false;
  for (let day = 0; day <= 6; day++) {
    const h = d.hours[day] !== undefined && d.hours[day] !== null ? d.hours[day] : d.hours[String(day)];
    if (!h || h.open !== true) continue;
    const st = Math.floor(Number(h.start));
    const en = Math.floor(Number(h.end));
    if (!isFinite(st) || !isFinite(en) || st < 0 || st > 23 || en < 1 || en > 24 || st >= en) continue;
    rows.push([day, st, en]);
    anyOpen = true;
  }

  const overrideKeys = Object.keys(overrideObj).sort();
  for (let i = 0; i < overrideKeys.length; i++) {
    const ymd = overrideKeys[i];
    if (!/^\d{4}-\d{2}-\d{2}$/.test(ymd)) continue;
    const oh = overrideObj[ymd];
    if (!oh || oh.open === false) continue;
    const ost = Math.floor(Number(oh.start));
    const oen = Math.floor(Number(oh.end));
    if (!isFinite(ost) || !isFinite(oen) || ost < 0 || ost > 23 || oen < 1 || oen > 24 || ost >= oen) continue;
    rows.push([ymd, ost, oen]);
    anyOpen = true;
  }

  if (!anyOpen) {
    return jsonResponse_({ status: 'error', message: 'nothing_open' });
  }
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  SpreadsheetApp.flush();
  clearBookingEndpointCaches_();
  return jsonResponse_({ status: 'success' });
}

function handleOwnerRejectBooking(d) {
  const eventId = String(d.eventId || '').trim();
  const token = String(d.ownerToken != null && d.ownerToken !== '' ? d.ownerToken : d.token || '').trim();
  const reason = sanitizeOwnerDeclineReason_(d.declineReason);
  if (!eventId || !token) {
    return jsonResponse_({ status: 'error', message: 'missing_fields' });
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
    if (String(data[i][8]) !== eventId) continue;
    if (!tokenMatches(data[i][9], token)) {
      return jsonResponse_({ status: 'error', message: 'invalid_token' });
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
    return jsonResponse_({ status: 'error', message: 'not_found' });
  }
  if (rowStatus !== 'PENDING' && rowStatus !== 'MOD_PROPOSED') {
    return jsonResponse_({ status: 'error', message: 'already_handled' });
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(WEB_ACTION_LOCK_TIMEOUT_MS)) {
    return jsonResponse_({ status: 'error', message: 'server_busy' });
  }

  try {
    const statusFresh = normalizeSheetStatus_(sheet.getRange(rowIndex, 8).getValue());
    if (statusFresh !== 'PENDING' && statusFresh !== 'MOD_PROPOSED') {
      return jsonResponse_({ status: 'error', message: 'already_handled' });
    }
    sheet.getRange(rowIndex, 8).setValue('REJECTED');
    sheet.getRange(rowIndex, 12).setValue('');
    sheet.getRange(rowIndex, 13).setValue('');
    sheet.getRange(rowIndex, 14).setValue('');
    sheet.getRange(rowIndex, SHEET_COL_OWNER_DECLINE_REASON).setValue(reason);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  const ev = cal.getEventById(eventId);
  if (ev) markInternalCalendarMutation_();
  if (ev) ev.deleteEvent();
  clearBookingEndpointCaches_();

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
  return jsonResponse_({ status: 'success' });
}

function handleReschedulePost(d) {
  if (!d.eventId || !d.token || !d.date || !d.time) {
    return jsonResponse_({ status: 'error', message: 'missing_fields' });
  }

  const rs = d.date ? String(d.date).split('T')[0].trim() : '';
  if (!rs || !/^\d{4}-\d{2}-\d{2}$/.test(rs)) {
    return jsonResponse_({ status: 'error', message: 'bad_date' });
  }
  if (!isDateInPublicBookingWindow_(rs)) {
    return jsonResponse_({ status: 'error', message: 'date_outside_booking_window' });
  }
  if (!isDateMeetingBookingLeadTime_(rs)) {
    return jsonResponse_({ status: 'error', message: 'date_too_soon' });
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
      return jsonResponse_({ status: 'error', message: 'invalid_token' });
    }
    rowIndex = i + 1;
    clientName = data[i][1];
    phone = data[i][2];
    service = data[i][3];
    clientEmail = data[i][6];
    rowStatus = normalizeSheetStatus_(data[i][7]);
    break;
  }

  if (rowIndex < 0) {
    return jsonResponse_({ status: 'error', message: 'not_found' });
  }
  if (rowStatus !== 'CONFIRMED' && rowStatus !== 'CLIENT_CONFIRMED') {
    return jsonResponse_({ status: 'error', message: 'not_confirmed' });
  }

  const start = parseYmdAndTimeLocal_(rs, convertTo24Hour(d.time));
  if (isNaN(start.getTime())) {
    return jsonResponse_({ status: 'error', message: 'bad_datetime' });
  }
  if (!isBookingWithinStudioHours_(start, 1)) {
    return jsonResponse_({ status: 'error', message: 'invalid_day_or_time' });
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(WEB_ACTION_LOCK_TIMEOUT_MS)) {
    return jsonResponse_({ status: 'error', message: 'server_busy' });
  }

  let neatTimeStr = '';
  let neatDate = '';
  try {
    const statusFresh = normalizeSheetStatus_(sheet.getRange(rowIndex, 8).getValue());
    if (statusFresh !== 'CONFIRMED' && statusFresh !== 'CLIENT_CONFIRMED') {
      return jsonResponse_({ status: 'error', message: 'not_confirmed' });
    }

    const currentEventId = String(sheet.getRange(rowIndex, 9).getValue() || d.eventId).trim() || String(d.eventId);
    const cal = CalendarApp.getCalendarById(CALENDAR_ID);
    const ev = cal.getEventById(currentEventId);
    if (!ev) {
      return jsonResponse_({ status: 'error', message: 'event_missing' });
    }

    const durMin = effectiveDurationMinutesFromEvent_(ev);
    if (!isBookingWithinStudioHours_(start, durMin)) {
      return jsonResponse_({ status: 'error', message: 'invalid_day_or_time' });
    }
    const newEnd = new Date(start.getTime() + durMin * 60000);
    if (slotOverlapsExistingCalendarEvents_(start, newEnd, [currentEventId])) {
      return jsonResponse_({ status: 'error', message: 'slot_unavailable' });
    }

    markInternalCalendarMutation_();
    ev.setTime(start, newEnd);
    ev.setTitle(clientName);
    neatTimeStr = Utilities.formatDate(start, Session.getScriptTimeZone(), 'h:mm a');
    applyBookingLocationToEvent_(ev, neatTimeStr, service);
    sheet.getRange(rowIndex, 5).setValue(start);
    sheet.getRange(rowIndex, 6).setValue(neatTimeStr);
    sheet.getRange(rowIndex, 11).setValue('');
    SpreadsheetApp.flush();
    clearBookingEndpointCaches_();
    neatDate = Utilities.formatDate(start, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy');
  } finally {
    lock.releaseLock();
  }

  const clientHtml = getRescheduledClientEmailHtml(clientName, neatDate, neatTimeStr, service);
  if (clientEmail) {
    MailApp.sendEmail({
      to: clientEmail,
      name: "Roni's Nail Studio",
      subject: 'Appointment rescheduled - Roni\'s Nail Studio',
      htmlBody: clientHtml,
    });
  }
  const ownerRows =
    emailDetailRow('Client', clientName) +
    emailDetailRow('Date', neatDate) +
    emailDetailRow('Time', neatTimeStr) +
    emailDetailRow('Service', service);
  const ownerBody =
    '<div style="font-family:sans-serif;padding:24px;max-width:480px;margin:auto;"><h2 style="font-weight:500;">Appointment rescheduled</h2><table style="width:100%;border-collapse:collapse;margin-top:16px;background:#fafafa;border-radius:8px;"><tbody>' +
    ownerRows +
    '</tbody></table><p style="margin-top:16px;color:#666;">' +
    escapeHtml(clientEmail || '') +
    '</p></div>';
  MailApp.sendEmail({
    to: MY_EMAIL,
    subject: 'Rescheduled: ' + clientName,
    htmlBody: ownerBody,
  });
  return jsonResponse_({ status: 'success' });
}

function doGet(e) {
  const action = e.parameter.action;
  const eventId = e.parameter.eventId;
  const token = e.parameter.token;

  if (action === 'work_hours') {
    return handleWorkHoursGet_(e);
  }

  if (action === 'accept' || action === 'reject') {
    return handleOwnerActionGet_(e);
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
        if (ev) markInternalCalendarMutation_();
        if (ev) ev.deleteEvent();
      } finally {
        lock.releaseLock();
      }
      clearBookingEndpointCaches_();
      const neatD = formatSheetDateForEmail(dateVal);
      const neatTime = formatSheetTimeForEmail(timeStr);
      if (clientEmail) {
        const declineHtml = getAlternateDeclinedClientEmailHtml(clientName, neatD, neatTime, service);
        MailApp.sendEmail({
          to: clientEmail,
          name: "Roni's Nail Studio",
          subject: 'Alternate time - Roni\'s Nail Studio',
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
    markInternalCalendarMutation_();
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
    clearBookingEndpointCaches_();
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
    let clientName = 'Client', clientEmail = '', service = '', dateVal = '', timeStr = '';
    let rowStatus = '';
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
          markInternalCalendarMutation_();
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
        if (ev) markInternalCalendarMutation_();
        if (ev) ev.deleteEvent();
        sheet.getRange(rowIndex, 8).setValue('CANCELLED');
        SpreadsheetApp.flush();
        clearBookingEndpointCaches_();
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
    return handlePublicBusyFeedGet_(e);
  }
  return htmlPage('Bad request', '<h2>Bad request</h2>');
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
    if (isJsonBoolTrue_(d.ownerAcceptBooking)) {
      return handleOwnerAcceptBooking(d);
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
    return handlePublicBookingPost_(d);
  } catch (err) {
    const msg = err && err.message ? String(err.message) : 'server_error';
    return jsonResponse_({ status: 'error', message: msg });
  }
}

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

function installStudioCalendarOnChangeTrigger() {
  const fn = 'syncCalendarToSpreadsheet';
  const wantId = String(CALENDAR_ID).toLowerCase().replace(/\s/g, '');
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() !== fn) continue;
    if (triggers[i].getEventType() !== ScriptApp.EventType.ON_EVENT_UPDATED) continue;
    const src = String(triggers[i].getTriggerSourceId() || '').toLowerCase().replace(/\s/g, '');
    if (src === wantId) {
      Logger.log('Studio calendar on-change trigger already present (source ' + triggers[i].getTriggerSourceId() + ')');
      return;
    }
  }
  ScriptApp.newTrigger(fn).forUserCalendar(CALENDAR_ID).onEventUpdated().create();
  Logger.log('Installed onEventUpdated trigger on studio calendar: ' + CALENDAR_ID);
}

function reinstallStudioCalendarOnChangeTrigger() {
  const fn = 'syncCalendarToSpreadsheet';
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = triggers.length - 1; i >= 0; i--) {
    if (triggers[i].getHandlerFunction() !== fn) continue;
    if (triggers[i].getEventType() === ScriptApp.EventType.ON_EVENT_UPDATED) ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger(fn).forUserCalendar(CALENDAR_ID).onEventUpdated().create();
  Logger.log('Reinstalled onEventUpdated on studio calendar: ' + CALENDAR_ID);
}

function reinstallCalendarSyncTrigger() {
  const fn = 'syncCalendarToSpreadsheet';
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = triggers.length - 1; i >= 0; i--) {
    if (triggers[i].getHandlerFunction() !== fn) continue;
    if (triggers[i].getEventType() === ScriptApp.EventType.CLOCK) ScriptApp.deleteTrigger(triggers[i]);
  }
  try {
    ScriptApp.newTrigger(fn).timeBased().everyMinutes(10).create();
    Logger.log('Reinstalled CLOCK: ' + fn + ' every 10 minutes');
  } catch (e) {
    ScriptApp.newTrigger(fn).timeBased().everyHours(1).create();
    Logger.log('Reinstalled CLOCK: ' + fn + ' every 1 hour (10 min not allowed). Reason: ' + e);
  }
}

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

function installArchiveOldBookingsTrigger() {
  const fn = 'archiveEligibleBookings';
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === fn) {
      Logger.log('Trigger already exists for ' + fn);
      return;
    }
  }
  ScriptApp.newTrigger(fn).timeBased().everyDays(1).atHour(3).create();
  Logger.log('Installed daily trigger: ' + fn + ' at 3:00 AM (project time zone)');
}

function reinstallArchiveOldBookingsTrigger() {
  const fn = 'archiveEligibleBookings';
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = triggers.length - 1; i >= 0; i--) {
    if (triggers[i].getHandlerFunction() === fn) ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger(fn).timeBased().everyDays(1).atHour(3).create();
  Logger.log('Reinstalled daily trigger: ' + fn + ' at 3:00 AM');
}

function installAllBookingAutomationTriggers() {
  installStudioCalendarOnChangeTrigger();
  installCalendarSyncTrigger();
  installTwoDayReminderTrigger();
  installArchiveOldBookingsTrigger();
}

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
    const arch = ss.getSheetByName(ARCHIVE_SHEET_NAME);
    lines.push('Archive sheet: ' + (arch ? 'present rows=' + arch.getLastRow() : 'not created yet'));
  } catch (e2) {
    lines.push('Sheet ERROR: ' + e2);
  }
  try {
    if (typeof Calendar !== 'undefined' && Calendar.Events && typeof Calendar.Events.get === 'function') {
      lines.push('Advanced Calendar API: ENABLED');
    } else {
      lines.push('Advanced Calendar API: NOT enabled -> Editor -> Services -> add Google Calendar API');
    }
  } catch (e3) {
    lines.push('Advanced Calendar API: NOT enabled');
  }

  const triggers = ScriptApp.getProjectTriggers();
  let hasClock = false;
  let hasCal = false;
  let hasReminder = false;
  let hasArchive = false;
  const wantCalNorm = String(CALENDAR_ID).toLowerCase().replace(/\s/g, '');
  let calendarTriggerMatchesStudio = false;
  for (let i = 0; i < triggers.length; i++) {
    const fn = triggers[i].getHandlerFunction();
    const et = triggers[i].getEventType();
    if (fn === 'syncCalendarToSpreadsheet') {
      if (et === ScriptApp.EventType.CLOCK) hasClock = true;
      if (et === ScriptApp.EventType.ON_EVENT_UPDATED) {
        hasCal = true;
        let src = '';
        try {
          src = triggers[i].getTriggerSourceId();
        } catch (e) {
          src = '';
        }
        if (String(src).toLowerCase().replace(/\s/g, '') === wantCalNorm) calendarTriggerMatchesStudio = true;
        lines.push('Trigger: ' + fn + ' | eventType=' + et + (src ? ' | calendarSourceId=' + src : ''));
      } else {
        lines.push('Trigger: ' + fn + ' | eventType=' + et);
      }
    }
    if (fn === 'sendTwoDayReminders') hasReminder = true;
    if (fn === 'archiveEligibleBookings') hasArchive = true;
  }
  if (!hasCal) lines.push('TIP: no ON_EVENT_UPDATED -> run installStudioCalendarOnChangeTrigger()');
  else if (!calendarTriggerMatchesStudio) lines.push('FIX: Calendar trigger is on the wrong calendar. Run reinstallStudioCalendarOnChangeTrigger().');
  if (!hasClock) lines.push('TIP: no CLOCK backup -> run installCalendarSyncTrigger()');
  if (!hasReminder) lines.push('TIP: no daily reminder trigger -> run installTwoDayReminderTrigger()');
  if (!hasArchive) lines.push('TIP: no archive trigger -> run installArchiveOldBookingsTrigger()');
  const out = lines.join('\n');
  Logger.log(out);
  return out;
}

function setBookingAdminSecret() {
  PropertiesService.getScriptProperties().setProperty('BOOKING_ADMIN_SECRET', 'REPLACE_WITH_LONG_RANDOM_SECRET');
}

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
  rows.sort(function (a, b) { return a[0] - b[0]; });
  if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
}

function ensureArchiveSheet_() {
  const ss = getCRMSpreadsheet();
  let arch = ss.getSheetByName(ARCHIVE_SHEET_NAME);
  const source = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (!arch) arch = ss.insertSheet(ARCHIVE_SHEET_NAME);
  const sourceHeader = source.getRange(1, 1, 1, Math.max(source.getLastColumn(), 1)).getValues()[0];
  const width = Math.max(sourceHeader.length, arch.getLastColumn(), 1);
  if (arch.getLastRow() === 0) {
    const header = sourceHeader.slice();
    while (header.length < width) header.push('');
    arch.getRange(1, 1, 1, width).setValues([header]);
  } else if (arch.getLastColumn() < width) {
    arch.insertColumnsAfter(arch.getLastColumn(), width - arch.getLastColumn());
  }
  return arch;
}

function padRowToWidth_(row, width) {
  const out = row.slice();
  while (out.length < width) out.push('');
  return out;
}

function isAppointmentRowArchiveEligible_(row, nowMs) {
  const status = normalizeSheetStatus_(row[7]);
  if (!status) return false;
  const startMs = appointmentRowStartMs_(row[4], row[5]);
  if (isNaN(startMs)) return false;
  const cutoffMs = nowMs - ARCHIVE_AFTER_PAST_DAYS * 86400000;
  if (startMs >= cutoffMs) return false;
  return true;
}

function archiveEligibleBookings() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log('archiveEligibleBookings: skipped (another action is running)');
    return;
  }
  try {
    const ss = getCRMSpreadsheet();
    const live = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    if (live.getLastRow() < 2) {
      Logger.log('archiveEligibleBookings: no data rows');
      return;
    }
    const arch = ensureArchiveSheet_();
    const data = live.getDataRange().getValues();
    const nowMs = Date.now();
    const moveRows = [];
    const deleteRowNums = [];
    const width = Math.max(live.getLastColumn(), arch.getLastColumn(), data[0].length, 16);
    for (let i = 1; i < data.length; i++) {
      if (!isAppointmentRowArchiveEligible_(data[i], nowMs)) continue;
      moveRows.push(padRowToWidth_(data[i], width));
      deleteRowNums.push(i + 1);
    }
    if (!moveRows.length) {
      Logger.log('archiveEligibleBookings: nothing to archive');
      return;
    }
    arch.getRange(arch.getLastRow() + 1, 1, moveRows.length, width).setValues(moveRows);
    for (let j = deleteRowNums.length - 1; j >= 0; j--) {
      live.deleteRow(deleteRowNums[j]);
    }
    SpreadsheetApp.flush();
    Logger.log('archiveEligibleBookings: moved ' + moveRows.length + ' rows to ' + ARCHIVE_SHEET_NAME);
  } finally {
    lock.releaseLock();
  }
}

function ownerLookupDedupeKeys_(email, phone) {
  const keys = [];
  const em = String(email == null ? '' : email).trim().toLowerCase();
  if (em && em.indexOf('@') > 0) keys.push('e:' + em);
  let dig = String(phone == null ? '' : phone).replace(/\D/g, '');
  if (dig.length === 11 && dig.charAt(0) === '1') dig = dig.slice(1);
  if (dig.length >= 10) keys.push('p:' + dig.slice(-10));
  return keys;
}

function getBookingSheetsForLookup_() {
  const ss = getCRMSpreadsheet();
  const sheets = [];
  const live = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (live) sheets.push(live);
  const arch = ss.getSheetByName(ARCHIVE_SHEET_NAME);
  if (arch) sheets.push(arch);
  return sheets;
}

function handleOwnerClientLookup(d) {
  const secret = PropertiesService.getScriptProperties().getProperty('BOOKING_ADMIN_SECRET');
  if (!secret || String(d.adminSecret || '').trim() !== String(secret).trim()) {
    return jsonResponse_({ status: 'error', message: 'unauthorized' });
  }
  const q = String(d.query || '').trim().toLowerCase();
  if (q.length < 2) return jsonResponse_({ status: 'success', matches: [] });

  const rows = [];
  getBookingSheetsForLookup_().forEach(function (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const name = data[i][1];
      if (name == null || String(name).trim() === '') continue;
      if (String(name).toLowerCase().indexOf(q) < 0) continue;
      const ts = data[i][0];
      let tsMs = 0;
      if (ts instanceof Date && !isNaN(ts.getTime())) tsMs = ts.getTime();
      else {
        const startMs = appointmentRowStartMs_(data[i][4], data[i][5]);
        if (!isNaN(startMs)) tsMs = startMs;
      }
      rows.push({
        name: String(name).trim(),
        phone: String(data[i][2] || '').trim(),
        email: String(data[i][6] || '').trim(),
        tsMs: tsMs,
      });
    }
  });

  rows.sort(function (a, b) { return b.tsMs - a.tsMs; });
  const seen = Object.create(null);
  const matches = [];
  for (let i = 0; i < rows.length; i++) {
    const item = rows[i];
    const keys = ownerLookupDedupeKeys_(item.email, item.phone);
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
    matches.push({
      name: item.name,
      phone: item.phone,
      email: item.email,
      lastBooked: item.tsMs ? Utilities.formatDate(new Date(item.tsMs), Session.getScriptTimeZone(), 'MMM d, yyyy') : '',
    });
    if (matches.length >= 15) break;
  }
  return jsonResponse_({ status: 'success', matches: matches });
}

function handleOwnerProposeAlternate(d) {
  if (!d.eventId || !d.ownerToken || !d.proposedDate || !d.proposedTime) {
    return jsonResponse_({ status: 'error', message: 'missing_fields' });
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
    if (!tokenMatches(data[i][9], d.ownerToken)) return jsonResponse_({ status: 'error', message: 'invalid_token' });
    rowIndex = i + 1;
    const st = data[i][7];
    if (st !== 'PENDING') {
      if (st === 'MOD_PROPOSED') return jsonResponse_({ status: 'error', message: 'already_proposed' });
      return jsonResponse_({ status: 'error', message: 'not_pending' });
    }
    clientName = data[i][1];
    clientEmail = data[i][6];
    service = data[i][3];
    origDateVal = data[i][4];
    origTimeStr = data[i][5];
    break;
  }
  if (rowIndex < 0) return jsonResponse_({ status: 'error', message: 'not_found' });
  const pds = d.proposedDate.toString().split('T')[0];
  const propTimeTrim = String(d.proposedTime).trim();
  const testStart = parseYmdAndTimeLocal_(pds, convertTo24Hour(propTimeTrim));
  if (isNaN(testStart.getTime())) return jsonResponse_({ status: 'error', message: 'bad_datetime' });
  const calPropose = CalendarApp.getCalendarById(CALENDAR_ID);
  const pendingEvPropose = calPropose.getEventById(d.eventId);
  const proposeDurMin = pendingEvPropose ? effectiveDurationMinutesFromEvent_(pendingEvPropose) : 60;
  const proposeSlotEnd = new Date(testStart.getTime() + proposeDurMin * 60000);
  if (slotOverlapsExistingCalendarEvents_(testStart, proposeSlotEnd, [d.eventId])) {
    return jsonResponse_({ status: 'error', message: 'slot_unavailable' });
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
    const html = getAlternateProposalClientEmailHtml(clientName, origNeatD, origNeatT, newNeatD, propTimeTrim, service, d.eventId, modTok);
    MailApp.sendEmail({ to: clientEmail, name: "Roni's Nail Studio", subject: "Suggested appointment time — Roni's Nail Studio", htmlBody: html });
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
  return jsonResponse_({ status: 'success' });
}

function handleOwnerDirectBooking(d) {
  const secret = PropertiesService.getScriptProperties().getProperty('BOOKING_ADMIN_SECRET');
  if (!secret || String(d.adminSecret || '').trim() !== String(secret).trim()) {
    return jsonResponse_({ status: 'error', message: 'unauthorized' });
  }
  const clientName = String(d.clientName || '').trim();
  const phone = String(d.phone || '').trim();
  const email = String(d.email || '').trim();
  const service = String(d.service || '').trim();
  if (!clientName || !email || !service) return jsonResponse_({ status: 'error', message: 'missing_fields' });
  const dateStr = d.date ? String(d.date).split('T')[0].trim() : '';
  const timeStr = String(d.time || '').trim();
  if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr) || !timeStr) return jsonResponse_({ status: 'error', message: 'bad_datetime' });
  const start = parseYmdAndTimeLocal_(dateStr, convertTo24Hour(timeStr));
  if (isNaN(start.getTime())) return jsonResponse_({ status: 'error', message: 'bad_datetime' });
  const durMin = clampDurationMinutes_(d.durationMinutes);
  const end = new Date(start.getTime() + durMin * 60000);
  const actionToken = generateActionToken();
  const ss = getCRMSpreadsheet();
  const s = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const row = s.getLastRow() + 1;
  s.appendRow([new Date(), clientName, phone, service, start, timeStr, email, 'CONFIRMED', '', actionToken, '', '', '', '', '']);
  const c = CalendarApp.getCalendarById(CALENDAR_ID);
  const desc =
    'Client: ' + clientName +
    '\nPhone: ' + phone +
    '\nEmail: ' + email +
    '\nService: ' + service +
    '\nDuration: ' + durMin + ' minutes' +
    '\n\nBooked via Roni\'s Nail Studio website (admin booking page).' +
    '\n(Google Calendar may list the Google account that runs this script as the creator; that is normal for studio automation.)' +
    '\nDurationMinutes: ' + durMin;
  markInternalCalendarMutation_();
  const ev = c.createEvent(clientName, start, end, { description: desc });
  if (ev && typeof ev.setDescription === 'function') ev.setDescription(desc);
  const tz = Session.getScriptTimeZone();
  const neatTime = Utilities.formatDate(start, tz, 'h:mm a');
  applyBookingLocationToEvent_(ev, neatTime, service);
  s.getRange(row, 9).setValue(ev.getId());
  SpreadsheetApp.flush();
  clearBookingEndpointCaches_();
  const neatD = Utilities.formatDate(start, tz, 'EEEE, MMMM d, yyyy');
  const acceptEmailHtml = getConfirmedEmailHtml(clientName, neatD, neatTime, service);
  MailApp.sendEmail({ to: email, name: "Roni's Nail Studio", subject: "Appointment Confirmed: Roni's Nail Studio", htmlBody: acceptEmailHtml });
  const ownerRows = emailDetailRow('Client', clientName) + emailDetailRow('When', neatD + ' · ' + neatTime) + emailDetailRow('Service', service) + emailDetailRow('Email', email);
  MailApp.sendEmail({
    to: MY_EMAIL,
    subject: 'Studio booked (admin): ' + clientName,
    htmlBody: '<p style="font-family:sans-serif;">You added a confirmed appointment from the admin page.</p><table style="width:100%;border-collapse:collapse;margin-top:12px;background:#fafafa;border-radius:8px;"><tbody>' + ownerRows + '</tbody></table>',
  });
  return jsonResponse_({ status: 'success' });
}

function handleOwnerCalendarWeek(d) {
  const secret = PropertiesService.getScriptProperties().getProperty('BOOKING_ADMIN_SECRET');
  if (!secret || String(d.adminSecret || '').trim() !== String(secret).trim()) {
    return jsonResponse_({ status: 'error', message: 'unauthorized' });
  }
  const tz = Session.getScriptTimeZone();
  const viewRaw = String(d.calView == null ? 'week' : d.calView).trim().toLowerCase();
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
    return jsonResponse_({
      status: 'success',
      timeZone: tz,
      calView: 'month',
      month: built.monthKey,
      leadingBlankDays: built.leadingBlankDays,
      days: built.days,
      events: events,
    });
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
  return jsonResponse_({ status: 'success', timeZone: tz, calView: 'week', days: days, events: events });
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
    days.push({ ymd: ymd, startMs: dayStart.getTime(), endMs: nextDay.getTime(), label: label, dayOfMonth: dom });
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
    days.push({ ymd: ymd, startMs: dayStart.getTime(), endMs: next.getTime(), label: label });
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
        out.push({ start: ev.getStartTime().getTime(), end: ev.getEndTime().getTime(), title: String(ev.getTitle() || '').trim() || '(No title)', allDay: ev.isAllDayEvent(), calendar: key });
      }
    } catch (err) {
      Logger.log('collectCalendarWeekEvents_ ' + key + ' ' + err);
    }
  }
  addCal(CALENDAR_ID, 'studio');
  addCal(PERSONAL_CALENDAR_ID, 'personal');
  out.sort(function (a, b) { return a.start - b.start; });
  return out;
}

function isJsonBoolTrue_(v) {
  return v === true || v === 'true' || v === 1 || v === '1';
}

function getConfirmedEmailHtml(name, date, time, service) {
  const n = escapeHtml(name);
  const rows = emailDetailRow('Date', date) + emailDetailRow('Time', time) + emailDetailRow('Service', service);
  return '<div style="font-family:sans-serif;padding:40px;max-width:500px;margin:auto;border:1px solid #f0f0f0;border-radius:12px;background-color:#ffffff;"><p style="font-size:16px;color:#1a1a1a;line-height:1.6;">Hi ' + n + ',</p><p style="font-size:16px;color:#1a1a1a;line-height:1.6;margin-bottom:20px;">Your appointment at Roni\'s Nail Studio is <strong>confirmed</strong>. Here are the details:</p><table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:8px;margin:0 0 24px 0;" cellpadding="0" cellspacing="0" role="presentation"><tbody>' + rows + '</tbody></table><p style="font-size:16px;color:#1a1a1a;line-height:1.6;">Thank you for booking with me &lt;3</p><hr style="border:none;border-top:1px solid #f0f0f0;margin:40px 0;"><p style="font-size:13px;color:#999;text-align:center;">Roni\'s Nail Studio</p></div>';
}

function getRescheduledClientEmailHtml(name, date, time, service) {
  const n = escapeHtml(name);
  const rows = emailDetailRow('Date', date) + emailDetailRow('Time', time) + emailDetailRow('Service', service);
  return '<div style="font-family:sans-serif;padding:40px;max-width:500px;margin:auto;border:1px solid #f0f0f0;border-radius:12px;background-color:#ffffff;"><p style="font-size:16px;color:#1a1a1a;line-height:1.6;">Hi ' + n + ',</p><p style="font-size:16px;color:#1a1a1a;line-height:1.6;margin-bottom:20px;">Your appointment has been <strong>rescheduled</strong>. Here are your updated details:</p><table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:8px;margin:0 0 24px 0;" cellpadding="0" cellspacing="0" role="presentation"><tbody>' + rows + '</tbody></table><p style="font-size:16px;color:#1a1a1a;line-height:1.6;">We look forward to seeing you.</p><hr style="border:none;border-top:1px solid #f0f0f0;margin:40px 0;"><p style="font-size:13px;color:#999;text-align:center;">Roni\'s Nail Studio</p></div>';
}

function getCalendarUpdatedClientEmailHtml(name, date, time, service) {
  const n = escapeHtml(name);
  const rows = emailDetailRow('Date', date) + emailDetailRow('Time', time) + emailDetailRow('Service', service);
  return '<div style="font-family:sans-serif;padding:40px;max-width:500px;margin:auto;border:1px solid #f0f0f0;border-radius:12px;background-color:#ffffff;"><p style="font-size:16px;color:#1a1a1a;line-height:1.6;">Hi ' + n + ',</p><p style="font-size:16px;color:#1a1a1a;line-height:1.6;margin-bottom:20px;">Your appointment time has been <strong>updated</strong>. Here are your current details:</p><table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:8px;margin:0 0 24px 0;" cellpadding="0" cellspacing="0" role="presentation"><tbody>' + rows + '</tbody></table><p style="font-size:15px;color:#555;line-height:1.6;">If this doesn\'t work for you, reply to this email or use the reschedule link in your reminder when you receive it.</p><hr style="border:none;border-top:1px solid #f0f0f0;margin:40px 0;"><p style="font-size:13px;color:#999;text-align:center;">Roni\'s Nail Studio</p></div>';
}

function getCalendarDeletedClientEmailHtml(name, date, time, service) {
  const n = escapeHtml(name);
  const rows = emailDetailRow('Date', date) + emailDetailRow('Time', time) + emailDetailRow('Service', service);
  return '<div style="font-family:sans-serif;padding:40px;max-width:500px;margin:auto;border:1px solid #f0f0f0;border-radius:12px;background-color:#ffffff;"><p style="font-size:16px;color:#1a1a1a;line-height:1.6;">Hi ' + n + ',</p><p style="font-size:16px;color:#1a1a1a;line-height:1.6;margin-bottom:20px;">Your appointment below has been <strong>cancelled</strong>.</p><table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:8px;margin:0 0 24px 0;" cellpadding="0" cellspacing="0" role="presentation"><tbody>' + rows + '</tbody></table><p style="font-size:16px;color:#1a1a1a;line-height:1.6;">If you have questions or would like to book another time, please reach out or use our website.</p><hr style="border:none;border-top:1px solid #f0f0f0;margin:40px 0;"><p style="font-size:13px;color:#999;text-align:center;">Roni\'s Nail Studio</p></div>';
}

function getDeclinedEmailHtml(name, date, time, service, declineReasonNote) {
  const n = escapeHtml(name);
  const rows = emailDetailRow('Date', date) + emailDetailRow('Time', time) + emailDetailRow('Service', service);
  const noteTrim = String(declineReasonNote == null ? '' : declineReasonNote).trim();
  const reasonBlock = noteTrim ? '<p style="font-size:15px;color:#1a1a1a;line-height:1.65;margin:0 0 20px 0;padding:16px;background:#fafafa;border-radius:8px;border-left:3px solid #b76e7a;"><strong>Note from the studio:</strong><br><span style="white-space:pre-wrap;">' + escapeHtml(noteTrim) + '</span></p>' : '';
  return '<div style="font-family:sans-serif;padding:40px;max-width:500px;margin:auto;border:1px solid #f0f0f0;border-radius:12px;background-color:#ffffff;"><p style="font-size:16px;color:#1a1a1a;line-height:1.6;">Hi ' + n + ',</p><p style="font-size:16px;color:#1a1a1a;line-height:1.6;margin-bottom:20px;">Thank you for your interest in booking with Roni\'s Nail Studio. Unfortunately, we\'re unable to accommodate this request. Details below:</p>' + reasonBlock + '<table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:8px;margin:0 0 24px 0;" cellpadding="0" cellspacing="0" role="presentation"><tbody>' + rows + '</tbody></table><p style="font-size:16px;color:#1a1a1a;line-height:1.6;">You\'re welcome to book another time on our website when it works for you.</p><p style="font-size:16px;color:#1a1a1a;line-height:1.6;margin-top:16px;">If you have any questions, please feel free to reach out to me directly.</p><hr style="border:none;border-top:1px solid #f0f0f0;margin:40px 0;"><p style="font-size:13px;color:#999;text-align:center;">Roni\'s Nail Studio</p></div>';
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
  const rescheduleBlock = urlReschedule ? '<a href="' + urlReschedule + '" style="background-color:#fff;color:#111;border:1px solid #111;padding:12px;text-decoration:none;border-radius:8px;font-weight:500;display:block;margin-bottom:12px;">Reschedule</a>' : '<p style="font-size:14px;color:#666;">To reschedule, visit our website and book a new appointment time.</p>';
  return '<div style="font-family:sans-serif;padding:32px;max-width:450px;margin:auto;border:1px solid #eaeaea;border-radius:12px;"><h2 style="color:#111;font-weight:600;font-size:20px;text-align:center;line-height:1.35;">Please confirm your appointment</h2><p style="font-size:14px;color:#555;text-align:center;margin:8px 0 20px 0;">Your visit is in 2 days -- please let us know of any changes.</p><p style="font-size:16px;color:#1a1a1a;line-height:1.6;">Hi ' + n + ',</p><p style="font-size:16px;color:#1a1a1a;line-height:1.6;">Please tap <strong>Confirm</strong> if you&rsquo;re all set. If you need to change plans, use <strong>Reschedule</strong> or <strong>Cancel</strong>.</p><table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:8px;margin:20px 0;" cellpadding="0" cellspacing="0" role="presentation"><tbody>' + (emailDetailRow('Date', date) + emailDetailRow('Time', time) + emailDetailRow('Service', service)) + '</tbody></table><div style="text-align:center;"><a href="' + urlConfirm + '" style="background-color:#111;color:white;padding:14px;text-decoration:none;border-radius:8px;font-weight:600;display:block;margin-bottom:12px;">Confirm</a>' + rescheduleBlock + '<a href="' + urlCancel + '" style="background-color:#fff;color:#dc3545;border:1px solid #dc3545;padding:12px;text-decoration:none;border-radius:8px;display:block;">Cancel</a></div><p style="font-size:13px;color:#999;text-align:center;margin-top:24px;">Roni\'s Nail Studio</p></div>';
}

function getAlternateProposalClientEmailHtml(name, origDate, origTime, newDate, newTime, service, eventId, modToken) {
  const n = escapeHtml(name);
  const qOk = 'action=client_accept_mod&eventId=' + encodeURIComponent(eventId) + '&token=' + encodeURIComponent(modToken);
  const qNo = 'action=client_decline_mod&eventId=' + encodeURIComponent(eventId) + '&token=' + encodeURIComponent(modToken);
  const urlOk = buildBookingActionUrl(qOk);
  const urlNo = buildBookingActionUrl(qNo);
  const rows = emailDetailRow('Service', service) + emailDetailRow('You requested', origDate + ' · ' + origTime) + emailDetailRow('Suggested time', newDate + ' · ' + newTime);
  return '<div style="font-family:sans-serif;padding:40px;max-width:500px;margin:auto;border:1px solid #f0f0f0;border-radius:12px;background:#fff;"><p style="font-size:16px;color:#1a1a1a;">Hi ' + n + ',</p><p style="font-size:16px;color:#1a1a1a;line-height:1.6;">Roni suggested a different time that may work better. If it works for you, confirm below. If not, you can always pick another slot on our website.</p><table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:8px;margin:20px 0;" cellpadding="0" cellspacing="0" role="presentation"><tbody>' + rows + '</tbody></table><div style="text-align:center;"><a href="' + urlOk + '" style="background-color:#111;color:#fff;padding:14px;text-decoration:none;border-radius:8px;font-weight:500;display:block;margin-bottom:12px;">Yes, that works</a><a href="' + urlNo + '" style="background-color:#fff;color:#666;border:1px solid #ccc;padding:12px;text-decoration:none;border-radius:8px;display:block;">No thanks</a></div><p style="font-size:13px;color:#999;text-align:center;margin-top:24px;">Roni\'s Nail Studio</p></div>';
}

function getAlternateDeclinedClientEmailHtml(name, origDate, origTime, service) {
  const n = escapeHtml(name);
  const rows = emailDetailRow('Your original request', origDate + ' · ' + origTime) + emailDetailRow('Service', service);
  return '<div style="font-family:sans-serif;padding:40px;max-width:500px;margin:auto;border:1px solid #f0f0f0;border-radius:12px;"><p style="font-size:16px;color:#1a1a1a;">Hi ' + n + ',</p><p style="font-size:16px;color:#1a1a1a;line-height:1.6;">No problem — we won\'t hold that alternate time. Whenever you\'re ready, you can submit a new booking on our website.</p><table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:8px;margin:20px 0;" cellpadding="0" cellspacing="0" role="presentation"><tbody>' + rows + '</tbody></table><p style="font-size:13px;color:#999;text-align:center;margin-top:24px;">Roni\'s Nail Studio</p></div>';
}

function getRequestEmailHtml(name, service, phone, email, date, time, eventId, actionToken, clientNotes) {
  const q = 'action=accept&eventId=' + encodeURIComponent(eventId) + '&token=' + encodeURIComponent(actionToken);
  const qr = 'action=reject&eventId=' + encodeURIComponent(eventId) + '&token=' + encodeURIComponent(actionToken);
  const acc = buildBookingActionUrl(q);
  const rejBase = String(OWNER_REJECT_PAGE_BASE || '').trim();
  const rej = rejBase ? rejBase + (rejBase.indexOf('?') >= 0 ? '&' : '?') + 'eventId=' + encodeURIComponent(eventId) + '&token=' + encodeURIComponent(actionToken) : buildBookingActionUrl(qr);
  const modUrl = OWNER_MODIFY_PAGE_BASE + '?eventId=' + encodeURIComponent(eventId) + '&token=' + encodeURIComponent(actionToken);
  const notesTrim = String(clientNotes == null ? '' : clientNotes).trim();
  const notesBlock = notesTrim ? '<p style="margin:12px 0 0 0;padding-top:12px;border-top:1px solid #e8e8e8;"><strong>Notes / requests:</strong><br><span style="white-space:pre-wrap;color:#333;">' + escapeHtml(notesTrim) + '</span></p>' : '';
  return '<div style="font-family:sans-serif;padding:32px;max-width:450px;margin:auto;border:1px solid #eaeaea;border-radius:12px;"><h2 style="color:#111;font-weight:500;font-size:20px;text-align:center;">New Booking Request</h2><div style="background:#fafafa;padding:20px;border-radius:8px;margin:20px 0;font-size:15px;line-height:1.6;"><strong>Client:</strong> ' + name + '<br><strong>Service:</strong> ' + service + '<br><strong>Phone:</strong> ' + phone + '<br><strong>Email:</strong> ' + email + '<br><strong>Date:</strong> ' + date + '<br><strong>Time:</strong> ' + time + notesBlock + '</div><div style="text-align:center;"><a href="' + acc + '" style="background-color:#111;color:white;padding:14px;text-decoration:none;border-radius:8px;font-weight:500;display:block;margin-bottom:12px;">Approve Request</a><a href="' + modUrl + '" style="background-color:#fff;color:#111;border:1px solid #111;padding:12px;text-decoration:none;border-radius:8px;font-weight:500;display:block;margin-bottom:12px;">Suggest a different time</a><a href="' + rej + '" style="background-color:#fff;color:#dc3545;border:1px solid #dc3545;padding:12px;text-decoration:none;border-radius:8px;display:block;">Decline Request</a></div></div>';
}

/**
 * 2-day reminder: use calendar start for "is it 2 days away?" AND for the email body.
 * Sheet columns E/F can lag after you move an event in Google Calendar (stale Tuesday in email, Monday on calendar).
 */
function sendTwoDayReminders() {
  const ss = getCRMSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  for (let i = 1; i < data.length; i++) {
    const status = normalizeSheetStatus_(data[i][7]);
    if (status !== 'CONFIRMED' && status !== 'CLIENT_CONFIRMED') continue;
    if (data[i][10] === 'SENT') continue;
    const eventId = data[i][8];
    const tok = data[i][9];
    if (!eventId || !tok) continue;
    const rowStartMs = appointmentRowStartMs_(data[i][4], data[i][5]);
    const rowStartHint = isNaN(rowStartMs) ? null : new Date(rowStartMs);
    let ev;
    try {
      ev = getActiveBookingEvent_(calendar, eventId, rowStartHint);
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
    const clientEmail = data[i][6];
    if (!clientEmail) continue;
    const neatD = Utilities.formatDate(start, tz, 'EEEE, MMMM d, yyyy');
    const neatTime = Utilities.formatDate(start, tz, 'h:mm a');
    const sheetDateYmd = syncSheetDateToYyyyMmDd_(data[i][4]);
    const sheetTimeNorm = normalizeTimeToken_(formatSheetTimeForEmail(data[i][5]));
    const calDateYmd = Utilities.formatDate(start, tz, 'yyyy-MM-dd');
    const calTimeNorm = normalizeTimeToken_(Utilities.formatDate(start, tz, 'h:mm a'));
    if (calDateYmd !== sheetDateYmd || calTimeNorm !== sheetTimeNorm) {
      sheet.getRange(i + 1, 5).setValue(start);
      sheet.getRange(i + 1, 6).setValue(neatTime);
      SpreadsheetApp.flush();
      Logger.log(
        'sendTwoDayReminders: row ' + (i + 1) + ' sheet date/time synced from calendar before reminder email'
      );
    }
    const html = getTwoDayReminderEmailHtml(clientName, neatD, neatTime, service, eventId, tok);
    MailApp.sendEmail({ to: clientEmail, name: "Roni's Nail Studio", subject: TWO_DAY_REMINDER_EMAIL_SUBJECT, htmlBody: html });
    sheet.getRange(i + 1, 11).setValue('SENT');
    SpreadsheetApp.flush();
  }
}

function testTwoDayReminderPreview() {
  const html = getTwoDayReminderEmailHtml('Vy (Preview)', 'Monday, April 7, 2026', '2:00 PM', 'Gel Manicure', 'PREVIEW_NO_REAL_EVENT', 'preview_invalid_token');
  MailApp.sendEmail({
    to: MY_EMAIL,
    name: "Roni's Nail Studio",
    subject: TWO_DAY_REMINDER_EMAIL_SUBJECT,
    htmlBody: html + '<p style="font-size:11px;color:#b0b0b0;text-align:center;margin-top:28px;line-height:1.5;">Preview to your studio inbox — button links are placeholders (safe to click). For real links, run <strong>testTwoDayReminderLiveToStudio</strong>.</p>',
  });
  Logger.log('Sent 2-day reminder PREVIEW to ' + MY_EMAIL + ' (same subject line as clients)');
}

function testTwoDayReminderLiveToStudio() {
  const ss = getCRMSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  const tz = Session.getScriptTimeZone();
  for (let i = data.length - 1; i >= 1; i--) {
    const st = normalizeSheetStatus_(data[i][7]);
    if (st !== 'CONFIRMED' && st !== 'CLIENT_CONFIRMED') continue;
    const eventId = data[i][8];
    const tok = data[i][9];
    if (!eventId || !tok) continue;
    const clientName = data[i][1];
    const service = data[i][3];
    let neatD;
    let neatTime;
    let evTest = null;
    try {
      const hintMs = appointmentRowStartMs_(data[i][4], data[i][5]);
      const hint = isNaN(hintMs) ? null : new Date(hintMs);
      evTest = getActiveBookingEvent_(cal, eventId, hint);
    } catch (eTest) {
      evTest = null;
    }
    if (evTest) {
      const stEv = evTest.getStartTime();
      neatD = Utilities.formatDate(stEv, tz, 'EEEE, MMMM d, yyyy');
      neatTime = Utilities.formatDate(stEv, tz, 'h:mm a');
    } else {
      neatD = formatSheetDateForEmail(data[i][4]);
      neatTime = formatSheetTimeForEmail(data[i][5]);
    }
    const html = getTwoDayReminderEmailHtml(clientName, neatD, neatTime, service, eventId, tok);
    MailApp.sendEmail({
      to: MY_EMAIL,
      name: "Roni's Nail Studio",
      subject: TWO_DAY_REMINDER_EMAIL_SUBJECT,
      htmlBody: '<p style="font-family:sans-serif;font-size:14px;color:#b45309;background:#fffbeb;padding:12px;border-radius:8px;border:1px solid #fcd34d;"><strong>Test only — sent to studio inbox.</strong> Subject line matches what clients see. Buttons use real data: <strong>Cancel</strong> cancels this appointment.</p>' + html,
    });
    Logger.log('Sent LIVE test 2-day reminder for sheet row ' + (i + 1) + ' to ' + MY_EMAIL);
    return;
  }
  throw new Error('No CONFIRMED row with eventId and token. Approve a booking first.');
}

function testEmailPreview() {
  const mockName = 'Vy Nguyen (Test)';
  const mockService = 'Structured Gel New Set + Tier 2 Art';
  const mockDate = 'Monday, April 6, 2026';
  const mockTime = '11:30 AM';
  const mockPhone = '555-0199';
  const mockEmail = 'test@example.com';
  const reqHtml = getRequestEmailHtml(mockName, mockService, mockPhone, mockEmail, mockDate, mockTime, 'test_event_id', 'preview_only_invalid_token', 'Sample client note (preview only).');
  MailApp.sendEmail({ to: MY_EMAIL, subject: 'PREVIEW: New Booking Request', htmlBody: reqHtml });
  const confHtml = getConfirmedEmailHtml(mockName, mockDate, mockTime, mockService);
  MailApp.sendEmail({ to: MY_EMAIL, subject: 'PREVIEW: Appointment Confirmed', htmlBody: confHtml });
  const twoDayHtml = getTwoDayReminderEmailHtml(mockName, mockDate, mockTime, mockService, 'PREVIEW_NO_REAL_EVENT', 'preview_invalid_token');
  MailApp.sendEmail({ to: MY_EMAIL, name: "Roni's Nail Studio", subject: TWO_DAY_REMINDER_EMAIL_SUBJECT, htmlBody: twoDayHtml + '<p style="font-size:11px;color:#b0b0b0;text-align:center;margin-top:28px;">Preview — dummy links only.</p>' });
  Logger.log('Sent 3 preview emails to: ' + MY_EMAIL);
}

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

function isBookingWithinStudioHours_(start, durationMinutes) {
  if (!(start instanceof Date) || isNaN(start.getTime())) return false;
  const dm = Number(durationMinutes);
  if (!isFinite(dm) || dm <= 0) return false;
  const whAll = getWorkHoursPayload_();
  const ymd = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let wh = whAll.dateOverrides && whAll.dateOverrides[ymd] ? whAll.dateOverrides[ymd] : null;
  if (!wh && whAll.weekly) {
    const dow = start.getDay();
    wh = whAll.weekly[String(dow)];
  }
  if (!wh) return false;
  const y = start.getFullYear();
  const mo = start.getMonth();
  const day = start.getDate();
  const workStart = new Date(y, mo, day, wh.start, 0, 0);
  const workEnd = new Date(y, mo, day, wh.end, 0, 0);
  const slotEnd = new Date(start.getTime() + dm * 60000);
  return start >= workStart && slotEnd <= workEnd;
}

function slotOverlapsExistingCalendarEvents_(slotStart, slotEnd, ignoreEventIds) {
  const skip = Object.create(null);
  const ids = ignoreEventIds || [];
  for (var i = 0; i < ids.length; i++) {
    const id = String(ids[i] == null ? '' : ids[i]).trim();
    if (id) skip[id] = true;
  }
  if (!(slotStart instanceof Date) || !(slotEnd instanceof Date) || isNaN(slotStart.getTime()) || isNaN(slotEnd.getTime())) return true;
  if (slotEnd <= slotStart) return true;
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
  const s = String(timeStr == null ? '' : timeStr).trim().replace(/\s+/g, ' ');
  if (!s) return '00:00:00';
  const bits = s.split(' ');
  const modifier = bits.length > 1 ? bits[bits.length - 1].toUpperCase() : '';
  const timePart = bits.length > 1 ? bits.slice(0, -1).join(' ') : bits[0];
  let parts = timePart.split(':');
  let hours = parts[0];
  let minutes = parts[1];
  if (minutes === undefined) minutes = '00';
  else minutes = String(minutes).replace(/\D/g, '').slice(0, 2) || '00';
  hours = parseInt(hours, 10);
  minutes = parseInt(minutes, 10);
  if (isNaN(hours) || isNaN(minutes)) return '00:00:00';
  let h = hours;
  if (modifier === 'PM' && hours !== 12) h = hours + 12;
  if (modifier === 'AM' && hours === 12) h = 0;
  if (modifier !== 'AM' && modifier !== 'PM') h = hours;
  return String(h).padStart(2, '0') + ':' + String(minutes).padStart(2, '0') + ':00';
}user_query>