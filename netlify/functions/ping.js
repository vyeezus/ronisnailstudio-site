/** Instant response — proves Netlify Functions + redirects work. GET https://yoursite.com/api/ping */
exports.handler = async () => ({
  statusCode: 200,
  headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' },
  body: JSON.stringify({ ok: true, at: new Date().toISOString() }),
});
