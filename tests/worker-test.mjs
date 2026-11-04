// Checks the Cloudflare Worker (worker/src/index.js) without deploying it,
// using an in-memory stand-in for its KV storage.
// Run from the repository folder: node tests/worker-test.mjs
import assert from 'node:assert';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import worker from '../worker/src/index.js';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const PANEL = 'https://labelsheetkeeper.github.io';
let n = 0;
const ok = (c, m) => { assert(c, m); n++; };

class FakeKV {
  constructor() { this.m = new Map(); }
  async get(key, type) {
    const v = this.m.get(key);
    if (!v) return null;
    return type === 'json' ? JSON.parse(v.value) : v.value;
  }
  async getWithMetadata(key) {
    const v = this.m.get(key);
    return v ? { value: v.value, metadata: v.metadata || null } : { value: null, metadata: null };
  }
  async put(key, value, opts = {}) {
    // Like KV's one-write-a-second limit on a key: fail the next N writes to box records.
    if (this.busy && key.startsWith('box:')) { this.busy--; throw new Error('KV PUT failed: 429 Too Many Requests'); }
    this.m.set(key, { value, metadata: opts.metadata, ttl: opts.expirationTtl });
  }
  async delete(key) { this.m.delete(key); }
  keys(prefix) { return [...this.m.keys()].filter((k) => k.startsWith(prefix)); }
}
const kv = new FakeKV();
const env = { LABELS: kv };
const call = (p, init = {}) => worker.fetch(new Request('https://api.example.workers.dev' + p, init), env);

// ---- health and browser access ----
let r = await call('/health', { headers: { Origin: PANEL } });
let body = await r.json();
ok(r.status === 200 && body.ok === true && body.service === 'label-sheet-keeper', 'health answers ' + JSON.stringify(body));
ok(r.headers.get('Access-Control-Allow-Origin') === PANEL, 'the panel may call it');
ok(r.headers.get('Cache-Control') === 'no-store', 'answers are not cached');
r = await call('/health', { headers: { Origin: 'https://example.com' } });
ok(r.status === 200 && r.headers.get('Access-Control-Allow-Origin') === null, 'other websites get no browser access');
r = await call('/box/x', { method: 'OPTIONS', headers: { Origin: PANEL, 'Access-Control-Request-Method': 'DELETE' } });
ok(r.status === 204 && /DELETE/.test(r.headers.get('Access-Control-Allow-Methods')) && /Authorization/.test(r.headers.get('Access-Control-Allow-Headers')), 'browser pre-check allows the panel’s requests');
ok((await call('/health', { method: 'POST' })).status === 405, 'health is read-only');
ok((await call('/nothing-here')).status === 404, 'unknown paths are not found');

// ---- linking a phone ----
r = await call('/pair', { method: 'POST', headers: { Origin: 'https://example.com' } });
ok(r.status === 403, 'only the panel can link a phone');
r = await call('/pair', { method: 'POST', headers: { Origin: PANEL } });
const link = await r.json();
ok(r.status === 201 && /^[A-Za-z0-9_-]{22}$/.test(link.box) && /^[A-Za-z0-9_-]{32}$/.test(link.key), 'linking gives a mailbox and a read key ' + JSON.stringify(link));
ok(link.send === 'https://api.example.workers.dev/drop/' + link.box, 'and the phone’s sending address');
const rec = JSON.parse(kv.m.get('box:' + link.box).value);
ok(rec.key !== link.key && rec.key.length === 64, 'the read key itself is never stored, only its hash');
const auth = { Authorization: 'Bearer ' + link.key, Origin: PANEL };

// ---- sending labels ----
const pdf = fs.readFileSync(path.join(HERE, 'fixtures/sample-return-label.pdf'));
const png = fs.readFileSync(path.join(HERE, 'fixtures/mark-ship-good.png'));
r = await call('/drop/' + link.box, { method: 'POST', body: pdf, headers: { 'Content-Type': 'application/octet-stream', 'X-File-Name': encodeURIComponent('Amazon return.pdf') } });
ok(r.status === 201 && (await r.json()).ok === true, 'a PDF can be sent');
kv.busy = 1;
r = await call('/drop/' + link.box + '?reply=text', { method: 'POST', body: png });
ok(r.status === 201 && (await r.text()) === 'Sent to Label Sheet Keeper.', 'the Shortcut gets a sentence back');
ok(kv.busy === 0 && JSON.parse(kv.m.get('box:' + link.box).value).items.length === 2, 'a label sent right after another waits its turn and is kept');
r = await call('/drop/' + link.box + '?reply=text', { method: 'POST', body: 'hello, not a label' });
ok(r.status === 415 && /PDF or a picture/.test(await r.text()), 'other files are refused');
r = await call('/drop/' + link.box, { method: 'POST', body: new Uint8Array(0) });
ok(r.status === 400, 'an empty file is refused');
r = await call('/drop/' + link.box, { method: 'POST', body: new Uint8Array(10 * 1024 * 1024 + 1) });
ok(r.status === 413, 'a file over 10 MB is refused');
r = await call('/drop/AAAAAAAAAAAAAAAAAAAAAA?reply=text', { method: 'POST', body: pdf });
ok(r.status === 404 && /Link your phone again/.test(await r.text()), 'an unknown sending link says how to fix it');
ok(kv.keys('file:' + link.box).every((k) => kv.m.get(k).ttl === 7 * 24 * 3600), 'labels are kept 7 days at most');

// ---- the panel picks them up ----
r = await call('/box/' + link.box, { headers: { Origin: PANEL } });
ok(r.status === 401, 'listing needs the read key');
r = await call('/box/' + link.box, { headers: { Authorization: 'Bearer ' + 'x'.repeat(32), Origin: PANEL } });
ok(r.status === 401, 'a wrong read key is refused');
r = await call('/box/' + link.box, { headers: auth });
let list = (await r.json()).items;
ok(r.status === 200 && list.length === 2, 'both labels are waiting ' + JSON.stringify(list));
ok(list[0].name === 'Amazon return.pdf' && list[0].type === 'application/pdf' && list[1].type === 'image/png' && list[1].name === 'Label from your phone.png', 'names and real file types are kept');
r = await call('/box/' + link.box + '/' + list[0].id, { headers: auth });
const got = new Uint8Array(await r.arrayBuffer());
ok(r.status === 200 && r.headers.get('Content-Type') === 'application/pdf' && got.length === pdf.length && got[0] === pdf[0], 'the panel gets the file back');
r = await call('/box/' + link.box + '/' + list[0].id, { headers: { Origin: PANEL } });
ok(r.status === 401, 'files need the read key too');
r = await call('/box/' + link.box + '/' + list[0].id, { method: 'DELETE', headers: auth });
ok(r.status === 200 && (await r.json()).waiting === 1, 'a picked-up label is deleted');
ok(!kv.m.has('file:' + link.box + ':' + list[0].id), 'its file is gone from storage');
r = await call('/box/' + link.box + '/' + list[0].id, { headers: auth });
ok(r.status === 404, 'and can’t be fetched again');

// ---- limits ----
const stored = JSON.parse(kv.m.get('box:' + link.box).value);
stored.items.push({ id: 'oldoldoldoldoldo', type: 'image/png', name: 'old', size: 1, at: Date.now() - 8 * 24 * 3600 * 1000 });
await kv.put('box:' + link.box, JSON.stringify(stored));
list = (await (await call('/box/' + link.box, { headers: auth })).json()).items;
ok(list.length === 1 && list.every((it) => it.id !== 'oldoldoldoldoldo'), 'labels older than 7 days aren’t listed');
for (let i = 0; i < 19; i++) await call('/drop/' + link.box, { method: 'POST', body: png });
r = await call('/drop/' + link.box + '?reply=text', { method: 'POST', body: png });
ok(r.status === 429 && /Too many labels are waiting/.test(await r.text()), 'at most 20 labels wait at once');

// ---- unlinking ----
r = await call('/box/' + link.box, { method: 'DELETE', headers: auth });
ok(r.status === 200 && !kv.m.has('box:' + link.box) && kv.keys('file:' + link.box).length === 0, 'unlinking deletes the mailbox and every label');
r = await call('/drop/' + link.box, { method: 'POST', body: pdf });
ok(r.status === 404, 'the old sending link stops working');

console.log('all', n, 'worker checks passed');
