// Checks the Cloudflare Worker's request handling without deploying it.
// Run from the repository folder: node tests/worker-test.mjs
import assert from 'node:assert';
import worker from '../worker/src/index.js';

const PANEL = 'https://labelsheetkeeper.github.io';
let n = 0;
const ok = (c, m) => { assert(c, m); n++; };
const call = (path, init = {}) => worker.fetch(new Request('https://api.example.workers.dev' + path, init));

let r = await call('/health', { headers: { Origin: PANEL } });
let body = await r.json();
ok(r.status === 200 && body.ok === true && body.service === 'label-sheet-keeper', 'health answers ' + JSON.stringify(body));
ok(r.headers.get('Access-Control-Allow-Origin') === PANEL, 'the panel may call it');
ok(r.headers.get('Cache-Control') === 'no-store', 'answers are not cached');

r = await call('/health', { headers: { Origin: 'https://example.com' } });
ok(r.status === 200 && r.headers.get('Access-Control-Allow-Origin') === null, 'other websites get no browser access');

r = await call('/health', { method: 'OPTIONS', headers: { Origin: PANEL, 'Access-Control-Request-Method': 'POST' } });
ok(r.status === 204 && r.headers.get('Access-Control-Allow-Methods').includes('POST'), 'browser pre-check is answered');

r = await call('/health', { method: 'POST' });
ok(r.status === 405, 'health is read-only');

r = await call('/nothing-here');
ok(r.status === 404, 'unknown paths are not found');

console.log('all', n, 'worker checks passed');
