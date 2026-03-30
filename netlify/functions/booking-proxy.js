/**
 * Proxies Google Apps Script (GET: calendar JSONP, approve/decline HTML; POST: new booking).
 * APPS_SCRIPT_EXEC_URL must match SCRIPT_URL in google_apps_script.js.
 */
const APPS_SCRIPT_EXEC_URL =
  'https://script.google.com/macros/s/AKfycbzdT_rV3dR7Th4VHeLE3uJcyTPr4bI-6uy-_Im6xz-nZ0rGPToj85zy7Is7LmpNVS0Wwg/exec';

const FETCH_TIMEOUT_MS = 8000;
/** POST (bookings, owner alternate-time) can exceed 8s while MailApp + Sheets run on Google. */
const POST_FETCH_TIMEOUT_MS = 45000;

const BROWSER_HEADERS = {
  'User-Agent':
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
  Accept: 'text/html,application/xhtml+xml,application/json,application/javascript,*/*;q=0.8',
  'Accept-Language': 'en-US,en;q=0.9',
};

function getQueryString(event) {
  if (typeof event.rawQuery === 'string' && event.rawQuery.length) {
    return event.rawQuery;
  }
  const q = event.queryStringParameters;
  if (q && typeof q === 'object' && Object.keys(q).length) {
    return Object.keys(q)
      .map((k) => `${encodeURIComponent(k)}=${encodeURIComponent(q[k] == null ? '' : String(q[k]))}`)
      .join('&');
  }
  const mv = event.multiValueQueryStringParameters;
  if (mv && typeof mv === 'object' && Object.keys(mv).length) {
    const parts = [];
    for (const k of Object.keys(mv)) {
      const vals = mv[k];
      if (Array.isArray(vals)) {
        vals.forEach((v) => parts.push(`${encodeURIComponent(k)}=${encodeURIComponent(v == null ? '' : String(v))}`));
      }
    }
    return parts.join('&');
  }
  return '';
}

function stripMetaRefreshTags(html) {
  return html.replace(/<meta[^>]*http-equiv\s*=\s*["']?\s*refresh\s*["']?[^>]*>/gi, '');
}

async function fetchGoogleAppsScript(targetUrl) {
  const ac = new AbortController();
  const timer = setTimeout(() => ac.abort(), FETCH_TIMEOUT_MS);
  try {
    const res = await fetch(targetUrl, {
      method: 'GET',
      redirect: 'follow',
      headers: BROWSER_HEADERS,
      signal: ac.signal,
    });
    const contentType = res.headers.get('content-type') || 'text/html; charset=utf-8';
    const body = await res.text();
    return { ok: true, contentType, body };
  } catch (e) {
    const name = e && e.name === 'AbortError' ? 'timeout' : 'fetch';
    return {
      ok: false,
      contentType: 'text/html; charset=utf-8',
      body:
        '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="font-family:system-ui,sans-serif;text-align:center;padding:2rem">' +
        '<h2>Could not reach Google</h2><p>' +
        name +
        '</p></body></html>',
    };
  } finally {
    clearTimeout(timer);
  }
}

exports.handler = async (event) => {
  const method = event.httpMethod || 'GET';

  if (method === 'OPTIONS') {
    return { statusCode: 204, headers: {} };
  }

  if (method === 'POST') {
    let rawBody = event.body || '{}';
    if (event.isBase64Encoded && typeof rawBody === 'string') {
      rawBody = Buffer.from(rawBody, 'base64').toString('utf8');
    }
    const ac = new AbortController();
    const timer = setTimeout(() => ac.abort(), POST_FETCH_TIMEOUT_MS);
    try {
      const res = await fetch(APPS_SCRIPT_EXEC_URL, {
        method: 'POST',
        redirect: 'follow',
        headers: {
          ...BROWSER_HEADERS,
          'Content-Type': 'application/json',
        },
        body: rawBody,
        signal: ac.signal,
      });
      const text = await res.text();
      return {
        statusCode: res.ok ? 200 : res.status,
        headers: {
          'Content-Type': res.headers.get('content-type') || 'application/json; charset=utf-8',
          'Cache-Control': 'no-store',
          'X-Booking-Proxy': '1',
        },
        body: text,
      };
    } catch (e) {
      return {
        statusCode: 502,
        headers: { 'Content-Type': 'application/json; charset=utf-8' },
        body: JSON.stringify({
          status: 'error',
          reason: e && e.name === 'AbortError' ? 'timeout' : 'fetch',
        }),
      };
    } finally {
      clearTimeout(timer);
    }
  }

  if (method !== 'GET' && method !== 'HEAD') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  const qs = getQueryString(event);
  const qObj = event.queryStringParameters || {};

  if (qObj.__ping === '1') {
    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json', 'X-Booking-Proxy': 'ping' },
      body: JSON.stringify({ ok: true, message: 'booking-proxy is running' }),
    };
  }

  const targetUrl = qs ? `${APPS_SCRIPT_EXEC_URL}?${qs}` : APPS_SCRIPT_EXEC_URL;

  try {
    const result = await fetchGoogleAppsScript(targetUrl);
    let body = stripMetaRefreshTags(result.body || '');
    let outType = result.contentType || 'text/html; charset=utf-8';
    if (/text\/html/i.test(outType) && !/charset=/i.test(outType)) {
      outType = outType.replace(/text\/html/i, 'text/html; charset=utf-8');
    }
    const trimmed = body.trim();
    if (/^<(!DOCTYPE|html)/i.test(trimmed)) {
      outType = 'text/html; charset=utf-8';
    }
    if (/application\/javascript|text\/javascript/i.test(outType) || /^\s*[\w.]+\s*\(/.test(trimmed)) {
      outType = 'application/javascript; charset=utf-8';
    }

    if (method === 'HEAD') {
      return {
        statusCode: 200,
        headers: {
          'Content-Type': outType,
          'Cache-Control': 'no-store',
          'X-Content-Type-Options': 'nosniff',
          'X-Booking-Proxy': '1',
        },
        body: '',
      };
    }

    return {
      statusCode: 200,
      headers: {
        'Content-Type': outType,
        'Cache-Control': 'no-store',
        'X-Content-Type-Options': 'nosniff',
        'X-Booking-Proxy': '1',
      },
      body,
    };
  } catch (_err) {
    return {
      statusCode: 502,
      headers: { 'Content-Type': 'text/html; charset=utf-8', 'X-Booking-Proxy': 'error' },
      body: '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="font-family:sans-serif;text-align:center;padding:2rem"><h2>Proxy error</h2></body></html>',
    };
  }
};
