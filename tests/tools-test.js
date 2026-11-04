const T = require(require('path').join(__dirname, '..', 'lsk-tools.js'));
const assert = require('assert');
let n = 0; const ok = (c, m) => { assert(c, m); n++; };
// CSV with quotes, embedded newline and comma
let rows = T.parseDelimited('Name,Street,City,State,ZIP\n"Doe, Jane","12 Elm St\nUnit 4",Boston,MA,2108\n');
ok(rows.length === 2 && rows[1][0] === 'Doe, Jane' && rows[1][1] === '12 Elm St\nUnit 4', 'csv quotes');
// Excel paste (tabs)
rows = T.parseDelimited('Jane Doe\tAcme\t12 Elm St\t\tBoston\tMA\t02108\nJohn Roe\t\t9 Oak Ave\tApt 2\tCambridge\tMA\t02139');
ok(rows.length === 2 && rows[1][3] === 'Apt 2', 'tsv');
// Header mapping: Google Contacts style
let g = T.rowsToAddresses(T.parseDelimited('First Name,Last Name,Organization Name,Address 1 - Street,Address 1 - City,Address 1 - Region,Address 1 - Postal Code,Address 1 - Country\nJane,Doe,,12 Elm St,Boston,MA,2108,United States\nNo,Address,,,,,,\n'));
ok(g.length === 2, 'google rows ' + g.length);
let lines = T.addressLines(g[0]);
ok(JSON.stringify(lines) === JSON.stringify(['Jane Doe','12 Elm St','Boston, MA 02108']), 'google lines ' + JSON.stringify(lines));
ok(T.addressLines(g[1]).length === 1, 'name-only row stays name only');
// Formatted address column
let fm = T.rowsToAddresses(T.parseDelimited('Name,Address 1 - Formatted\n"Ann Lee","5 Pine Rd\nSalem, MA 01970"\n'));
ok(JSON.stringify(T.addressLines(fm[0])) === JSON.stringify(['Ann Lee','5 Pine Rd','Salem, MA 01970']), 'formatted ' + JSON.stringify(T.addressLines(fm[0])));
// Positional (no header)
let pos = T.rowsToAddresses([['Jane Doe','','12 Elm St','Apt 3','Boston','MA','2108']]);
ok(JSON.stringify(T.addressLines(pos[0])) === JSON.stringify(['Jane Doe','12 Elm St Apt 3','Boston, MA 02108']), 'positional ' + JSON.stringify(T.addressLines(pos[0])));
// vCard (Apple style, folded line, item prefix)
let vc = T.parseVCard('BEGIN:VCARD\r\nVERSION:3.0\r\nN:Doe;Jane;;;\r\nFN:Jane Doe\r\nORG:Acme Inc;\r\nitem1.ADR;type=HOME;type=pref:;;12 Elm St\\nUnit 4;Boston;MA;02108;United\r\n  States\r\nEND:VCARD\r\nBEGIN:VCARD\r\nFN:No Addr\r\nEND:VCARD\r\n');
ok(vc.length === 1, 'vcard count ' + vc.length);
ok(JSON.stringify(vc[0].lines) === JSON.stringify(['Jane Doe','Acme Inc','12 Elm St','Unit 4','Boston, MA 02108']), 'vcard lines ' + JSON.stringify(vc[0].lines));
// lines -> fields
let f = T.linesToFields(['Jane Doe','12 Elm St Apt 3','Boston, MA 02108']);
ok(f.name === 'Jane Doe' && f.street === '12 Elm St Apt 3' && f.city === 'Boston' && f.state === 'MA' && f.zip === '02108', 'linesToFields ' + JSON.stringify(f));
// Canada
let ca = T.addressLines({ name: 'Jane Roe', street: '123 Main St', apt: '5', city: 'Toronto', state: 'on', zip: 'm5v3l9' });
ok(JSON.stringify(ca) === JSON.stringify(['Jane Roe', '5-123 Main St', 'Toronto ON  M5V 3L9', 'CANADA']), 'canada from us ' + JSON.stringify(ca));
ca = T.addressLines({ name: 'Jane Roe', street: '123 Main St', apt: 'Suite 200', city: 'Toronto', state: 'ON', zip: 'M5V 3L9' }, { home: 'CA' });
ok(JSON.stringify(ca) === JSON.stringify(['Jane Roe', '123 Main St Suite 200', 'Toronto ON  M5V 3L9']), 'canada from canada ' + JSON.stringify(ca));
let us2 = T.addressLines({ name: 'Ann Lee', street: '5 Pine Rd', city: 'Salem', state: 'MA', zip: '1970' }, { home: 'CA' });
ok(JSON.stringify(us2) === JSON.stringify(['Ann Lee', '5 Pine Rd', 'Salem, MA 01970', 'USA']), 'us from canada ' + JSON.stringify(us2));
ok(T.countryOf({ state: 'CA', zip: '90210' }) === 'US' && T.countryOf({ state: 'QC' }) === 'CA' && T.countryOf({ zip: 'H2X 1Y4' }) === 'CA', 'countryOf');
ok(T.fixPostal('k1a0b1') === 'K1A 0B1', 'fixPostal');
let cf = T.linesToFields(['Jane Roe', '5-123 Main St', 'Toronto ON  M5V 3L9', 'CANADA']);
ok(cf.city === 'Toronto' && cf.state === 'ON' && cf.zip === 'M5V 3L9' && cf.street === '5-123 Main St', 'linesToFields canada ' + JSON.stringify(cf));
ok(T.regionCode('Massachusetts') === 'MA' && T.regionCode('Québec') === 'QC' && T.regionCode('on') === 'ON' && T.regionCode('Nouvelle-Écosse') === 'NS', 'regionCode');
// Fit: fake measure = 0.5 * size per character
const measure = (text, css) => { const size = parseFloat(css.match(/([\d.]+)pt/)[1]); return text.length * size * 0.5; };
const box = { width: 160, height: 60 };
const short = [{ text: 'Jane Doe', size: 10 }, { text: '12 Elm St', size: 10 }, { text: 'Boston, MA 02108', size: 10 }];
ok(T.fits(short, box, measure), 'short fits');
const long = short.concat([{ text: 'Second line of a much longer company name here', size: 10 }, { text: 'Extra line', size: 10 }]);
ok(!T.fits(long, box, measure), 'long does not fit');
const s = T.fitScale(long, box, measure);
ok(s > 0.5 && s < 1, 'scale ' + s);
const scaled = long.map(p => ({ ...p, size: p.size * s }));
ok(T.fits(scaled, box, measure), 'scaled fits');
ok(T.fitScale([{ text: 'x'.repeat(2000), size: 10 }], box, measure) === null, 'impossible -> null');
// Border finder: 400x500 white page, 4x6 box border 3px thick at (40,40)-(240,340), text and a full-width rule below
const w = 400, h = 500, gray = new Uint8Array(w * h).fill(255);
const rect = (x0, y0, x1, y1) => { for (let y = y0; y <= y1; y++) for (let x = x0; x <= x1; x++) gray[y * w + x] = 0; };
rect(40, 40, 240, 42); rect(40, 338, 240, 340); rect(40, 40, 42, 340); rect(238, 40, 240, 340);
rect(60, 100, 220, 102); // inner rule
rect(20, 420, 380, 421); // instructions rule outside the box
for (let i = 0; i < 30; i++) rect(60 + i * 5, 200, 61 + i * 5, 260); // barcode bars
const b = T.findLabelBorder(gray, w, h);
ok(b && Math.abs(b.x - 40) <= 2 && Math.abs(b.y - 40) <= 2 && Math.abs(b.w - 201) <= 3 && Math.abs(b.h - 301) <= 3, 'border ' + JSON.stringify(b));
// No border -> null
const plain = new Uint8Array(w * h).fill(255); for (let i = 0; i < 10; i++) for (let x = 50; x < 150; x++) plain[(60 + i * 10) * w + x] = 0;
ok(T.findLabelBorder(plain, w, h) === null, 'no border -> null');
console.log('all', n, 'tool checks passed');
