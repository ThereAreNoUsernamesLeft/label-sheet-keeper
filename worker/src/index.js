// Label Sheet Keeper server, a Cloudflare Worker at api.<subdomain>.workers.dev.
//
// Phone mailbox. The panel links a phone with POST /pair and gets a mailbox
// ("box") and a read key. The phone only ever knows the box, so it can send
// labels (POST /drop/<box>) but can't read them. The panel lists, fetches and
// deletes them with the read key. Labels are kept 7 days at most.
//
//   GET    /health               status
//   POST   /pair                 new mailbox → { box, key, send }        (panel only)
//   POST   /drop/<box>           send a label (raw PDF or picture body)  (phone)
//   GET    /box/<box>            labels waiting                          (read key)
//   GET    /box/<box>/<id>       one label's file                        (read key)
//   DELETE /box/<box>/<id>       picked up                               (read key)
//   DELETE /box/<box>            unlink: the mailbox and its labels      (read key)
//
// Storage is the LABELS KV namespace: "box:<box>" holds the mailbox record,
// "file:<box>:<id>" each label's bytes.

const VERSION = '0.2.0';

// Only the panel's own address may call the Worker from a browser.
const ALLOWED_ORIGINS = ['https://labelsheetkeeper.github.io'];

const MAX_BYTES = 10 * 1024 * 1024;   // per label
const MAX_WAITING = 20;               // labels waiting in one mailbox
const MAX_PER_DAY = 60;               // labels sent to one mailbox per day
const KEEP_SECONDS = 7 * 24 * 3600;   // a label not picked up is deleted after this

const BOX_RE = /^[A-Za-z0-9_-]{22}$/;
const ID_RE = /^[A-Za-z0-9_-]{16}$/;

const EXT = { 'application/pdf': 'pdf', 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp', 'image/heic': 'heic' };

// ---------- helpers ----------
function corsHeaders(origin) {
  const h = { Vary: 'Origin' };
  if (ALLOWED_ORIGINS.includes(origin)) {
    h['Access-Control-Allow-Origin'] = origin;
    h['Access-Control-Allow-Methods'] = 'GET, POST, DELETE, OPTIONS';
    h['Access-Control-Allow-Headers'] = 'Content-Type, Authorization, X-File-Name';
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

// The iPhone Shortcut asks for ?reply=text so its notification shows a sentence.
function message(text, status, origin, asText) {
  if (!asText) return json(status < 300 ? { ok: true, message: text } : { error: text }, status, origin);
  return new Response(text, { status, headers: { 'Content-Type': 'text/plain; charset=utf-8', 'Cache-Control': 'no-store', ...corsHeaders(origin) } });
}

function randomId(bytes) {
  const a = new Uint8Array(bytes);
  crypto.getRandomValues(a);
  return btoa(String.fromCharCode(...a)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

async function sha256(text) {
  const d = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  return [...new Uint8Array(d)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

function sameText(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== b.length) return false;
  let x = 0;
  for (let i = 0; i < a.length; i++) x |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return x === 0;
}

// What the file really is, from its first bytes (the phone's own label isn't trusted).
function sniff(b) {
  const ascii = (from, to) => String.fromCharCode(...b.subarray(from, to));
  if (b.length >= 8 && b[0] === 0x89 && ascii(1, 4) === 'PNG') return 'image/png';
  if (b.length >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'image/jpeg';
  if (b.length >= 6 && /^GIF8[79]a$/.test(ascii(0, 6))) return 'image/gif';
  if (b.length >= 12 && ascii(0, 4) === 'RIFF' && ascii(8, 12) === 'WEBP') return 'image/webp';
  if (b.length >= 12 && ascii(4, 8) === 'ftyp' && /^(heic|heix|hevc|hevx|mif1|msf1|heim|heis)$/.test(ascii(8, 12))) return 'image/heic';
  if (ascii(0, Math.min(b.length, 1024)).indexOf('%PDF-') >= 0) return 'application/pdf';
  return '';
}

function fileName(request, type) {
  let name = '';
  try { name = decodeURIComponent(request.headers.get('X-File-Name') || ''); } catch (e) { name = ''; }
  name = name.replace(/^.*[\\/]/, '').replace(/[\u0000-\u001f\u007f]/g, '').trim().slice(0, 80);
  return name || 'Label from your phone.' + EXT[type];
}

const boxKey = (box) => 'box:' + box;
const fileKey = (box, id) => 'file:' + box + ':' + id;
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// KV takes one write a second to the same key, so a label sent (or picked up)
// right after another waits its turn.
async function putBox(env, box, rec) {
  for (let tries = 1; ; tries++) {
    try { return await env.LABELS.put(boxKey(box), JSON.stringify(rec)); } catch (e) {
      if (tries >= 4) throw e;
      await sleep(1100);
    }
  }
}

async function getBox(env, box) {
  if (!BOX_RE.test(box)) return null;
  return env.LABELS.get(boxKey(box), 'json');
}

function current(rec) {
  const oldest = Date.now() - KEEP_SECONDS * 1000;
  return (rec.items || []).filter((it) => it.at > oldest);
}

async function withKey(request, env, box) {
  const rec = await getBox(env, box);
  if (!rec) return null;
  const m = /^Bearer\s+([A-Za-z0-9_-]{20,})$/.exec(request.headers.get('Authorization') || '');
  if (!m) return null;
  return sameText(await sha256(m[1]), rec.key) ? rec : null;
}

// ---------- routes ----------
async function pair(request, env, url, origin) {
  if (!ALLOWED_ORIGINS.includes(origin)) return json({ error: 'Link your phone from the Label Sheet Keeper panel.' }, 403, origin);
  const box = randomId(16), key = randomId(24);
  await putBox(env, box, { key: await sha256(key), items: [], created: Date.now(), day: '', sent: 0 });
  return json({ box, key, send: url.origin + '/drop/' + box }, 201, origin);
}

async function drop(request, env, box, origin, asText) {
  const rec = await getBox(env, box);
  if (!rec) return message('This sending link isn’t active. Link your phone again from the panel’s Settings.', 404, origin, asText);
  if (Number(request.headers.get('Content-Length') || 0) > MAX_BYTES) return message('That file is over 10 MB.', 413, origin, asText);
  const buf = await request.arrayBuffer();
  if (!buf.byteLength) return message('The file was empty.', 400, origin, asText);
  if (buf.byteLength > MAX_BYTES) return message('That file is over 10 MB.', 413, origin, asText);
  const type = sniff(new Uint8Array(buf, 0, Math.min(buf.byteLength, 1024)));
  if (!type) return message('Send a PDF or a picture (PNG, JPEG or HEIC).', 415, origin, asText);

  rec.items = current(rec);
  if (rec.items.length >= MAX_WAITING) return message('Too many labels are waiting. Open the panel to pick them up first.', 429, origin, asText);
  const today = new Date().toISOString().slice(0, 10);
  if (rec.day !== today) { rec.day = today; rec.sent = 0; }
  if (rec.sent >= MAX_PER_DAY) return message('That’s the most labels for one day. Try again tomorrow.', 429, origin, asText);

  const id = randomId(12), name = fileName(request, type);
  await env.LABELS.put(fileKey(box, id), buf, { expirationTtl: KEEP_SECONDS, metadata: { type, name } });
  rec.items.push({ id, type, name, size: buf.byteLength, at: Date.now() });
  rec.sent++;
  await putBox(env, box, rec);
  return message('Sent to Label Sheet Keeper.', 201, origin, asText);
}

async function boxRoute(request, env, box, id, origin) {
  const rec = await withKey(request, env, box);
  if (!rec) return json({ error: 'not linked' }, 401, origin);
  const items = current(rec);

  if (!id && request.method === 'GET') {
    return json({ items: items.map(({ id, type, name, size, at }) => ({ id, type, name, size, at })) }, 200, origin);
  }
  if (!id && request.method === 'DELETE') {
    await Promise.all((rec.items || []).map((it) => env.LABELS.delete(fileKey(box, it.id))));
    await env.LABELS.delete(boxKey(box));
    return json({ ok: true }, 200, origin);
  }
  if (!id || !ID_RE.test(id)) return json({ error: 'not found' }, 404, origin);

  if (request.method === 'GET') {
    const got = await env.LABELS.getWithMetadata(fileKey(box, id), 'arrayBuffer');
    if (!got || !got.value) return json({ error: 'This label is no longer waiting.' }, 404, origin);
    const type = (got.metadata && got.metadata.type) || 'application/octet-stream';
    return new Response(got.value, { status: 200, headers: { 'Content-Type': type, 'Cache-Control': 'no-store', ...corsHeaders(origin) } });
  }
  if (request.method === 'DELETE') {
    await env.LABELS.delete(fileKey(box, id));
    rec.items = items.filter((it) => it.id !== id);
    await putBox(env, box, rec);
    return json({ ok: true, waiting: rec.items.length }, 200, origin);
  }
  return json({ error: 'method not allowed' }, 405, origin);
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const origin = request.headers.get('Origin') || '';
    const parts = url.pathname.split('/').filter(Boolean);

    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: corsHeaders(origin) });

    try {
      if (url.pathname === '/' || url.pathname === '/health') {
        if (request.method !== 'GET') return json({ error: 'method not allowed' }, 405, origin);
        return json({ ok: true, service: 'label-sheet-keeper', version: VERSION }, 200, origin);
      }
      if (parts[0] === 'pair' && parts.length === 1) {
        if (request.method !== 'POST') return json({ error: 'method not allowed' }, 405, origin);
        return await pair(request, env, url, origin);
      }
      if (parts[0] === 'drop' && parts.length === 2) {
        if (request.method !== 'POST') return json({ error: 'method not allowed' }, 405, origin);
        return await drop(request, env, parts[1], origin, url.searchParams.get('reply') === 'text');
      }
      if (parts[0] === 'box' && (parts.length === 2 || parts.length === 3)) {
        return await boxRoute(request, env, parts[1], parts[2] || '', origin);
      }
      return json({ error: 'not found' }, 404, origin);
    } catch (e) {
      return json({ error: 'Something went wrong on the server. Try again.' }, 500, origin);
    }
  }
};
