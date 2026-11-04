// Label Sheet Keeper server, a Cloudflare Worker at api.<subdomain>.workers.dev.
// For now it only answers /health, so deploys can be checked end to end.

const VERSION = '0.1.0';

// Only the panel's own address may call the Worker from a browser.
const ALLOWED_ORIGINS = ['https://labelsheetkeeper.github.io'];

function corsHeaders(origin) {
  const h = { Vary: 'Origin' };
  if (ALLOWED_ORIGINS.includes(origin)) {
    h['Access-Control-Allow-Origin'] = origin;
    h['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS';
    h['Access-Control-Allow-Headers'] = 'Content-Type';
    h['Access-Control-Max-Age'] = '86400';
  }
  return h;
}

function json(body, status, origin) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', ...corsHeaders(origin) }
  });
}

export default {
  async fetch(request) {
    const url = new URL(request.url);
    const origin = request.headers.get('Origin') || '';

    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: corsHeaders(origin) });

    if (url.pathname === '/' || url.pathname === '/health') {
      if (request.method !== 'GET') return json({ error: 'method not allowed' }, 405, origin);
      return json({ ok: true, service: 'label-sheet-keeper', version: VERSION }, 200, origin);
    }

    return json({ error: 'not found' }, 404, origin);
  }
};
