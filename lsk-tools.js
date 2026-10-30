/*
  Label Sheet Keeper — helpers that don't touch Word:
    reading address lists (pasted cells, CSV, Excel, contact cards),
    turning an address into label lines,
    estimating whether text fits on a label,
    finding a shipping label's printed border on a page image.
*/
(function (root) {
  'use strict';

  // ---------- delimited text (CSV, or cells pasted from Excel/Numbers) ----------
  function parseDelimited(text, delim) {
    text = String(text || '').replace(/^﻿/, '');
    if (!delim) {
      var firstLine = text.split(/\r?\n/)[0] || '';
      delim = firstLine.indexOf('\t') >= 0 ? '\t' : (firstLine.split(';').length > firstLine.split(',').length ? ';' : ',');
    }
    var rows = [], row = [], cell = '', i = 0, q = false, ch;
    while (i < text.length) {
      ch = text[i];
      if (q) {
        if (ch === '"') {
          if (text[i + 1] === '"') { cell += '"'; i += 2; continue; }
          q = false; i++; continue;
        }
        cell += ch; i++; continue;
      }
      if (ch === '"' && cell === '') { q = true; i++; continue; }
      if (ch === delim) { row.push(cell); cell = ''; i++; continue; }
      if (ch === '\r') { i++; continue; }
      if (ch === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; i++; continue; }
      cell += ch; i++;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows.filter(function (r) { return r.some(function (c) { return String(c).trim() !== ''; }); });
  }

  // ---------- Excel (.xlsx): first sheet, cell values as text ----------
  function colIndex(ref) {
    var m = /^([A-Z]+)/.exec(ref || ''), n = 0;
    if (!m) return 0;
    for (var i = 0; i < m[1].length; i++) n = n * 26 + (m[1].charCodeAt(i) - 64);
    return n - 1;
  }
  function readXlsx(data, JSZipLib) {
    var Z = JSZipLib || root.JSZip;
    if (!Z) return Promise.reject(new Error('The file reader didn’t load.'));
    var parse = function (s) { return new root.DOMParser().parseFromString(s, 'application/xml'); };
    var zip;
    return Z.loadAsync(data).then(function (z) {
      zip = z;
      var wb = zip.file('xl/workbook.xml'), rels = zip.file('xl/_rels/workbook.xml.rels'), ss = zip.file('xl/sharedStrings.xml');
      if (!wb) throw new Error('That isn’t an Excel workbook.');
      return Promise.all([wb.async('string'), rels ? rels.async('string') : '', ss ? ss.async('string') : '']);
    }).then(function (parts) {
      var wbDoc = parse(parts[0]);
      var sheet = wbDoc.getElementsByTagName('sheet')[0];
      var rid = sheet && (sheet.getAttribute('r:id') || sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'));
      var path = 'xl/worksheets/sheet1.xml';
      if (rid && parts[1]) {
        var rel = Array.prototype.filter.call(parse(parts[1]).getElementsByTagName('Relationship'), function (r) { return r.getAttribute('Id') === rid; })[0];
        if (rel) { var t = rel.getAttribute('Target').replace(/^\//, ''); path = t.indexOf('xl/') === 0 ? t : 'xl/' + t; }
      }
      var strings = [];
      if (parts[2]) {
        Array.prototype.forEach.call(parse(parts[2]).getElementsByTagName('si'), function (si) {
          strings.push(Array.prototype.map.call(si.getElementsByTagName('t'), function (t) { return t.textContent; }).join(''));
        });
      }
      var f = zip.file(path);
      if (!f) throw new Error('The workbook’s first sheet couldn’t be read.');
      return f.async('string').then(function (xml) {
        var rows = [];
        Array.prototype.forEach.call(parse(xml).getElementsByTagName('row'), function (r) {
          var out = [];
          Array.prototype.forEach.call(r.getElementsByTagName('c'), function (c) {
            var type = c.getAttribute('t'), v = c.getElementsByTagName('v')[0], val = '';
            if (type === 's') val = strings[Number(v && v.textContent)] || '';
            else if (type === 'inlineStr') val = Array.prototype.map.call(c.getElementsByTagName('t'), function (t) { return t.textContent; }).join('');
            else if (type === 'b') val = v && v.textContent === '1' ? 'TRUE' : 'FALSE';
            else val = v ? v.textContent : '';
            out[colIndex(c.getAttribute('r'))] = val;
          });
          for (var i = 0; i < out.length; i++) if (out[i] === undefined) out[i] = '';
          rows.push(out);
        });
        return rows.filter(function (r) { return r.some(function (c) { return String(c).trim() !== ''; }); });
      });
    });
  }

  // ---------- contact cards (.vcf from Mac Contacts or Google Contacts) ----------
  function parseVCard(text) {
    var lines = String(text || '').replace(/\r\n[ \t]/g, '').replace(/\n[ \t]/g, '').split(/\r?\n/);
    var out = [], cur = null;
    var unesc = function (s) { return String(s || '').replace(/\\n/gi, '\n').replace(/\\([,;\\])/g, '$1').trim(); };
    lines.forEach(function (line) {
      if (/^BEGIN:VCARD/i.test(line)) { cur = { name: '', n: '', org: '', adr: [] }; return; }
      if (/^END:VCARD/i.test(line)) { if (cur) out.push(cur); cur = null; return; }
      if (!cur) return;
      var idx = line.indexOf(':'); if (idx < 0) return;
      var key = line.slice(0, idx).toUpperCase(), val = line.slice(idx + 1);
      var prop = key.split(';')[0].replace(/^ITEM\d+\./, '');
      if (prop === 'FN') cur.name = unesc(val);
      else if (prop === 'N') { var p = val.split(';'); cur.n = [unesc(p[1]), unesc(p[2]), unesc(p[0])].filter(Boolean).join(' '); }
      else if (prop === 'ORG') cur.org = unesc(val.split(';')[0]);
      else if (prop === 'ADR') {
        var a = val.split(';');
        cur.adr.push({ street: unesc(a[2]), apt: unesc(a[1]), city: unesc(a[3]), state: unesc(a[4]), zip: unesc(a[5]), country: unesc(a[6]) });
      }
    });
    var contacts = [];
    out.forEach(function (c) {
      c.adr.forEach(function (a) {
        var addr = { name: c.name || c.n, company: c.org, street: a.street, apt: a.apt, city: a.city, state: a.state, zip: a.zip, country: a.country };
        var ls = addressLines(addr);
        if (ls.length > 1) contacts.push({ name: ls[0], lines: ls, fields: addr });
      });
    });
    return contacts;
  }

  // ---------- rows with or without a header -> addresses ----------
  var FIELDS = ['name', 'company', 'street', 'apt', 'city', 'state', 'zip', 'country'];
  var ALIASES = {
    name: ['name', 'full name', 'display name', 'recipient', 'contact', 'ship to', 'to'],
    first: ['first name', 'given name', 'first'],
    middle: ['middle name', 'additional name'],
    last: ['last name', 'family name', 'surname', 'last'],
    company: ['company', 'organization', 'organization name', 'organization 1 - name', 'business', 'company name'],
    street: ['street', 'address', 'address 1', 'address line 1', 'street address', 'address 1 - street', 'home street', 'business street', 'mailing street', 'address1', 'street 1'],
    apt: ['apt', 'apt/suite', 'apt / suite', 'suite', 'unit', 'address 2', 'address line 2', 'address2', 'apartment', 'address 1 - extended address', 'home street 2', 'street 2'],
    city: ['city', 'town', 'address 1 - city', 'home city', 'business city', 'mailing city'],
    state: ['state', 'st', 'province', 'region', 'address 1 - region', 'home state', 'business state', 'mailing state'],
    zip: ['zip', 'zip code', 'zipcode', 'postal code', 'postcode', 'address 1 - postal code', 'home postal code', 'business postal code', 'mailing zip'],
    country: ['country', 'address 1 - country', 'home country/region', 'country/region', 'business country/region'],
    formatted: ['address 1 - formatted', 'formatted address', 'home address', 'full address', 'mailing address']
  };
  function norm(s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').replace(/[._]/g, ' ').trim(); }

  // Which column holds which field, if the first row is a header row.
  function headerMap(row) {
    var map = {}, hits = 0;
    (row || []).forEach(function (cell, i) {
      var h = norm(cell);
      Object.keys(ALIASES).forEach(function (k) {
        if (map[k] === undefined && ALIASES[k].indexOf(h) >= 0) { map[k] = i; hits++; }
      });
    });
    return hits >= 2 ? map : null;
  }

  function rowsToAddresses(rows) {
    if (!rows || !rows.length) return [];
    var map = headerMap(rows[0]);
    var body = map ? rows.slice(1) : rows;
    return body.map(function (r) {
      var get = function (k) { return map && map[k] !== undefined ? String(r[map[k]] || '').trim() : ''; };
      var a = {};
      if (map) {
        a.name = get('name') || [get('first'), get('middle'), get('last')].filter(Boolean).join(' ');
        FIELDS.slice(1).forEach(function (k) { a[k] = get(k); });
        if (!a.street && !a.city && get('formatted')) {
          var parts = get('formatted').split(/\r?\n/).map(function (s) { return s.trim(); }).filter(Boolean);
          a.extra = parts;
        }
      } else {
        FIELDS.forEach(function (k, i) { a[k] = String(r[i] === undefined ? '' : r[i]).trim(); });
      }
      return a;
    }).filter(function (a) { return addressLines(a).length > 0; });
  }

  var US = /^(us|usa|u\.s\.a?\.?|united states( of america)?)$/i;
  var CANADA = /^(ca|can|canada)$/i;
  var PROVINCES = ['AB', 'BC', 'MB', 'NB', 'NL', 'NS', 'NT', 'NU', 'ON', 'PE', 'QC', 'SK', 'YT'];
  var CA_POSTAL = /^([ABCEGHJ-NPRSTVXY]\d[ABCEGHJ-NPRSTV-Z])[ -]?(\d[ABCEGHJ-NPRSTV-Z]\d)$/i;
  // Canadian postal code in Canada Post's form: "M5V 3L9".
  function fixPostal(code) {
    var m = CA_POSTAL.exec(String(code || '').trim());
    return m ? (m[1] + ' ' + m[2]).toUpperCase() : String(code || '').trim();
  }
  // 'US', 'CA', or another country's name as given.
  function countryOf(a) {
    var c = String(a.country || '').trim();
    if (c) return US.test(c) ? 'US' : CANADA.test(c) ? 'CA' : c;
    if (PROVINCES.indexOf(String(a.state || '').trim().toUpperCase()) >= 0 || CA_POSTAL.test(String(a.zip || '').trim())) return 'CA';
    return 'US';
  }
  function fixZip(zip, state) {
    zip = String(zip || '').trim();
    if (CA_POSTAL.test(zip)) return fixPostal(zip);
    if (/^\d{3,4}$/.test(zip) && (!state || /^[A-Za-z]{2}$/.test(state))) zip = ('00000' + zip).slice(-5);
    return zip;
  }
  // An address (fields) -> the lines printed on a label.
  // opts.home: 'US' (default) or 'CA', the country the mail is sent from. Mail
  // to the other country gets a country line at the bottom.
  // Canadian addresses follow Canada Post: "5-123 Main St" for a unit, and
  // "Toronto ON  M5V 3L9" (two spaces before the postal code).
  function addressLines(a, opts) {
    var home = opts && opts.home === 'CA' ? 'CA' : 'US';
    var country = countryOf(a);
    var lines = [];
    var push = function (s) { s = String(s || '').trim(); if (s) lines.push(s); };
    push(a.name);
    if (a.company && norm(a.company) !== norm(a.name)) push(a.company);
    if (a.extra && a.extra.length) a.extra.forEach(push);
    var street = String(a.street || '').trim(), apt = String(a.apt || '').trim();
    if (street && apt) {
      if (country === 'CA' && /^\d+[A-Z]?$/i.test(apt) && /^\d/.test(street)) push(apt.toUpperCase() + '-' + street);
      else push(street + (/^(#|apt|suite|ste|unit|fl|floor|rm|room)/i.test(apt) ? ' ' : ', ') + apt);
    } else { String(street).split(/\r?\n/).forEach(push); push(apt); }
    var city = String(a.city || '').trim(), state = String(a.state || '').trim();
    if (country === 'CA') push([city, [state.toUpperCase(), fixPostal(a.zip)].filter(Boolean).join('  ')].filter(Boolean).join(' '));
    else push([city, [state, fixZip(a.zip, a.state)].filter(Boolean).join(' ')].filter(Boolean).join(', '));
    if (!lines.length) return lines;
    if (country === 'CA') { if (home !== 'CA') push('CANADA'); }
    else if (country === 'US') { if (home === 'CA') push('USA'); }
    else push(a.country);
    return lines;
  }

  // ---------- will this text fit? ----------
  // paras: [{ text, size, bold, italic, font, before, after, indent, lineSpacing }]
  // box: { width, height } in points. measure(text, fontCss) -> width in points.
  function lineHeightFactor(font) {
    var f = String(font || '').toLowerCase();
    if (/calibri|aptos/.test(f)) return 1.22;
    if (/helvetica/.test(f)) return 1.2;
    if (/times|georgia/.test(f)) return 1.15;
    return 1.15;
  }
  function fontCss(p, scale) {
    var size = (p.size || 10) * (scale || 1);
    return (p.italic ? 'italic ' : '') + (p.bold ? 'bold ' : '') + size.toFixed(2) + 'pt "' + (p.font || 'Arial') + '", Arial, sans-serif';
  }
  function wrapCount(text, width, css, measure) {
    var words = String(text || '').split(/\s+/).filter(Boolean);
    if (!words.length) return { lines: 1, tooWide: false };
    var lines = 1, cur = '', tooWide = false;
    words.forEach(function (w) {
      if (measure(w, css) > width) tooWide = true;
      var test = cur ? cur + ' ' + w : w;
      if (cur && measure(test, css) > width) { lines++; cur = w; } else cur = test;
    });
    return { lines: lines, tooWide: tooWide };
  }
  function textHeight(paras, box, measure, scale) {
    scale = scale || 1;
    var h = 0, tooWide = false;
    paras.forEach(function (p) {
      var size = (p.size || 10) * scale;
      var mult = p.lineSpacing && p.lineSpacing > 0 ? p.lineSpacing / 12 : 1;
      var w = box.width - (p.indent || 0);
      var r = wrapCount(p.text, w, fontCss(p, scale), measure);
      if (r.tooWide) tooWide = true;
      h += (p.before || 0) + r.lines * size * lineHeightFactor(p.font) * mult + (p.after || 0);
    });
    return { height: h, tooWide: tooWide };
  }
  function fits(paras, box, measure) {
    var r = textHeight(paras, box, measure, 1);
    return !r.tooWide && r.height <= box.height + 0.5;
  }
  // Largest scale (<= 1) at which the text fits, or null if even 5pt text won't.
  function fitScale(paras, box, measure) {
    if (fits(paras, box, measure)) return 1;
    var maxSize = Math.max.apply(null, paras.map(function (p) { return p.size || 10; }));
    var lo = 5 / maxSize, hi = 1;
    var ok = function (s) { var r = textHeight(paras, box, measure, s); return !r.tooWide && r.height <= box.height + 0.5; };
    if (!ok(lo)) return null;
    for (var i = 0; i < 14; i++) { var mid = (lo + hi) / 2; if (ok(mid)) lo = mid; else hi = mid; }
    return lo;
  }

  // ---------- find a shipping label's printed border ----------
  // gray: Uint8Array of w*h luminance values. Returns {x, y, w, h} in those
  // pixels, or null. Looks for the largest rectangle made of long dark lines.
  function findLabelBorder(gray, w, h) {
    var DARK = 110, GAP = 2;
    var dark = function (x, y) { return gray[y * w + x] < DARK; };
    function runs(len, other, at) {
      var best = null, start = -1, lastDark = -10;
      for (var i = 0; i < len; i++) {
        if (at(i)) { if (start < 0 || i - lastDark > GAP + 1) start = i; lastDark = i; if (!best || lastDark - start > best.b - best.a) best = { a: start, b: lastDark }; }
      }
      return best;
    }
    function lines(count, len, minLen, at) {
      var found = [];
      for (var k = 0; k < count; k++) {
        var r = runs(len, k, function (i) { return at(i, k); });
        if (r && r.b - r.a >= minLen) found.push({ k: k, a: r.a, b: r.b });
      }
      // Merge neighbors (a thick line spans several rows or columns).
      var merged = [];
      found.forEach(function (l) {
        var m = merged[merged.length - 1];
        if (m && l.k - m.k2 <= 2 && Math.abs(l.a - m.a) <= 6 && Math.abs(l.b - m.b) <= 6) { m.k2 = l.k; m.a = Math.min(m.a, l.a); m.b = Math.max(m.b, l.b); }
        else merged.push({ k1: l.k, k2: l.k, a: l.a, b: l.b });
      });
      return merged;
    }
    var H = lines(h, w, w * 0.2, function (x, y) { return dark(x, y); });
    var V = lines(w, h, h * 0.12, function (y, x) { return dark(x, y); });
    if (H.length < 2 || V.length < 2) return null;
    var tolX = w * 0.025, tolY = h * 0.02, best = null;
    for (var i = 0; i < H.length; i++) {
      for (var j = i + 1; j < H.length; j++) {
        var t = H[i], b = H[j];
        if (Math.abs(t.a - b.a) > tolX || Math.abs(t.b - b.b) > tolX) continue;
        var top = t.k1, bottom = b.k2, left = Math.min(t.a, b.a), right = Math.max(t.b, b.b);
        if (bottom - top < h * 0.12) continue;
        var span = bottom - top;
        var hasSide = function (x) {
          return V.some(function (v) { return Math.abs((v.k1 + v.k2) / 2 - x) <= tolX && v.a <= top + tolY + span * 0.1 && v.b >= bottom - tolY - span * 0.1; });
        };
        if (!hasSide(left) || !hasSide(right)) continue;
        var area = (right - left) * span;
        if ((right - left) > w * 0.97 && span > h * 0.97) continue;   // the page edge itself
        if (!best || area > best.area) best = { x: left, y: top, w: right - left + 1, h: span + 1, area: area };
      }
    }
    return best ? { x: best.x, y: best.y, w: best.w, h: best.h } : null;
  }

  // Label lines (or a formatted address) -> fields, best effort:
  // first line is the name, a last line like "City, ST 12345" or
  // "City ON  M5V 3L9" is split up, and a country line is set aside.
  function linesToFields(lines) {
    lines = (lines || []).map(function (l) { return String(l).trim(); }).filter(Boolean);
    var a = { name: lines[0] || '', company: '', street: '', apt: '', city: '', state: '', zip: '', country: '' };
    var rest = lines.slice(1);
    if (rest.length > 1 && (US.test(rest[rest.length - 1]) || CANADA.test(rest[rest.length - 1]))) a.country = rest.pop();
    if (rest.length && !/\d/.test(rest[0]) && rest.length > 2) a.company = rest.shift();
    var last = rest[rest.length - 1] || '';
    var m = /^(.*?),?\s+([A-Za-z]{2})\.?\s+(\d{5}(?:-\d{4})?|[A-Za-z]\d[A-Za-z][ -]?\d[A-Za-z]\d)$/.exec(last);
    if (m) { a.city = m[1].replace(/,$/, ''); a.state = m[2].toUpperCase(); a.zip = fixZip(m[3]); rest.pop(); }
    a.street = rest.join(', ');
    return a;
  }

  // ---------- looking up addresses online ----------
  function fetchJson(url, ms) {
    var ctl = typeof AbortController !== 'undefined' ? new AbortController() : null;
    var timer = ctl ? setTimeout(function () { ctl.abort(); }, ms || 6000) : 0;
    return fetch(url, ctl ? { signal: ctl.signal } : undefined).then(function (r) {
      clearTimeout(timer);
      if (!r.ok) { var e = new Error('HTTP ' + r.status); e.status = r.status; throw e; }
      return r.json();
    }, function (e) { clearTimeout(timer); throw e; });
  }
  var GEO = 'https://api.geoapify.com/v1/geocode/autocomplete?format=json&lang=en';
  // City and state/province for a US ZIP or Canadian postal code. Uses the
  // Geoapify key when there is one, otherwise the free Zippopotam.us service.
  // Resolves to { city, state, country } or null.
  function lookupPostal(code, key) {
    code = String(code || '').trim();
    var us = /^(\d{5})(?:-\d{4})?$/.exec(code), ca = CA_POSTAL.exec(code);
    if ((!us && !ca) || typeof fetch === 'undefined') return Promise.resolve(null);
    var country = us ? 'US' : 'CA';
    var free = function () {
      return fetchJson('https://api.zippopotam.us/' + (us ? 'us/' + us[1] : 'ca/' + ca[1].toUpperCase())).then(function (j) {
        var p = j && j.places && j.places[0];
        if (!p) return null;
        return { city: String(p['place name'] || '').replace(/\s*\(.*\)\s*$/, '').trim(), state: String(p['state abbreviation'] || '').toUpperCase(), country: country };
      }).catch(function () { return null; });
    };
    if (!key) return free();
    return fetchJson(GEO + '&type=postcode&limit=1&filter=countrycode:' + country.toLowerCase() + '&text=' + encodeURIComponent(us ? us[1] : fixPostal(code)) + '&apiKey=' + encodeURIComponent(key))
      .then(function (j) {
        var r = j && j.results && j.results[0];
        return r && r.city ? { city: r.city, state: String(r.state_code || '').toUpperCase(), country: country } : null;
      }).catch(function () { return null; }).then(function (r) { return r || free(); });
  }
  // State and province names -> two-letter codes (Photon gives full names).
  var REGIONS = {
    alabama: 'AL', alaska: 'AK', arizona: 'AZ', arkansas: 'AR', california: 'CA', colorado: 'CO', connecticut: 'CT', delaware: 'DE',
    'district of columbia': 'DC', florida: 'FL', georgia: 'GA', hawaii: 'HI', idaho: 'ID', illinois: 'IL', indiana: 'IN', iowa: 'IA',
    kansas: 'KS', kentucky: 'KY', louisiana: 'LA', maine: 'ME', maryland: 'MD', massachusetts: 'MA', michigan: 'MI', minnesota: 'MN',
    mississippi: 'MS', missouri: 'MO', montana: 'MT', nebraska: 'NE', nevada: 'NV', 'new hampshire': 'NH', 'new jersey': 'NJ',
    'new mexico': 'NM', 'new york': 'NY', 'north carolina': 'NC', 'north dakota': 'ND', ohio: 'OH', oklahoma: 'OK', oregon: 'OR',
    pennsylvania: 'PA', 'rhode island': 'RI', 'south carolina': 'SC', 'south dakota': 'SD', tennessee: 'TN', texas: 'TX', utah: 'UT',
    vermont: 'VT', virginia: 'VA', washington: 'WA', 'west virginia': 'WV', wisconsin: 'WI', wyoming: 'WY', 'puerto rico': 'PR',
    alberta: 'AB', 'british columbia': 'BC', manitoba: 'MB', 'new brunswick': 'NB', 'newfoundland and labrador': 'NL', 'nova scotia': 'NS',
    'northwest territories': 'NT', nunavut: 'NU', ontario: 'ON', 'prince edward island': 'PE', quebec: 'QC', saskatchewan: 'SK', yukon: 'YT',
    'colombie-britannique': 'BC', 'nouveau-brunswick': 'NB', 'terre-neuve-et-labrador': 'NL', 'nouvelle-ecosse': 'NS',
    'territoires du nord-ouest': 'NT', 'ile-du-prince-edouard': 'PE'
  };
  function regionCode(name) {
    var s = String(name || '').trim();
    if (/^[A-Za-z]{2}$/.test(s)) return s.toUpperCase();
    var k = s.toLowerCase();
    if (k.normalize) k = k.normalize('NFD').replace(/[̀-ͯ]/g, '');
    return REGIONS[k] || s;
  }

  // Address suggestions in the US and Canada while typing a street. With a
  // Geoapify key they come from Geoapify; otherwise from Photon by Komoot,
  // a free service with no key (please keep use light). Both use
  // OpenStreetMap data. Resolves to [{ label, sub, street, city, state, zip,
  // country, source }]; rejects with err.status (401/403: bad key, 429: too many).
  function suggestAddresses(text, key) { return key ? geoapifySuggest(text, key) : photonSuggest(text); }
  function geoapifySuggest(text, key) {
    return fetchJson(GEO + '&limit=5&filter=countrycode:us,ca&text=' + encodeURIComponent(text) + '&apiKey=' + encodeURIComponent(key)).then(function (j) {
      return ((j && j.results) || []).filter(function (r) { return r.street || r.address_line1; }).map(function (r) {
        var country = String(r.country_code || '').toUpperCase() === 'CA' ? 'CA' : 'US';
        var street = r.street ? [r.housenumber, r.street].filter(Boolean).join(' ') : String(r.address_line1 || '');
        return {
          label: r.address_line1 || street, sub: r.address_line2 || '', street: street, city: r.city || '',
          state: String(r.state_code || '').toUpperCase(), zip: country === 'CA' ? fixPostal(r.postcode) : String(r.postcode || ''),
          country: country, source: 'geoapify'
        };
      });
    });
  }
  var PHOTON = 'https://photon.komoot.io/api/?limit=8&lang=en&layer=house&layer=street&countrycode=US&countrycode=CA&q=';
  function photonSuggest(text) {
    // A street without that house number in the map still gets the number typed.
    var typedNumber = (/^\s*(\d+[A-Za-z]?(?:-\d+)?)\s/.exec(text) || [])[1] || '';
    return fetchJson(PHOTON + encodeURIComponent(text)).then(function (j) {
      var seen = {};
      return ((j && j.features) || []).map(function (f) {
        var p = f.properties || {}, country = String(p.countrycode || '').toUpperCase();
        if (country !== 'US' && country !== 'CA') return null;
        var streetName = p.street || (p.type === 'street' ? p.name : '');
        if (!streetName) return null;
        var street = [p.housenumber || typedNumber, streetName].filter(Boolean).join(' ');
        var city = p.city || p.locality || p.district || '', state = regionCode(p.state);
        var zip = country === 'CA' ? fixPostal(p.postcode) : String(p.postcode || '');
        var sub = [city, [state, zip].filter(Boolean).join(' ')].filter(Boolean).join(', ') + (country === 'CA' ? ', Canada' : '');
        var id = (street + '|' + sub).toLowerCase();
        if (seen[id]) return null;
        seen[id] = true;
        return { label: street, sub: sub, street: street, city: city, state: state, zip: zip, country: country, source: 'photon' };
      }).filter(Boolean).slice(0, 5);
    });
  }

  // ---------- reordering labels ----------
  // A store link -> where the Reorder button goes. Amazon and Walmart links
  // can put the labels straight in the cart; other stores open the product
  // page. Resolves nothing; returns { url, cart, store } or null for a bad link.
  function reorderLink(link, packs) {
    var u;
    try { u = new URL(String(link || '').trim()); } catch (e) { return null; }
    if (!/^https?:$/.test(u.protocol)) return null;
    packs = Math.max(1, Math.min(99, Math.round(Number(packs) || 1)));
    var host = u.hostname.toLowerCase().replace(/^www\./, ''), m;
    if (/^(smile\.)?amazon\.(com|ca|co\.uk|de|fr|it|es|nl|se|pl|com\.mx|com\.au|co\.jp|in|sg|ae|sa|com\.br|com\.tr|com\.be)$/.test(host)) {
      m = /\/(?:dp|gp\/product|gp\/aw\/d|exec\/obidos\/asin|o\/asin)\/([A-Z0-9]{10})(?:[/?]|$)/i.exec(u.pathname + '/');
      if (m) return { url: 'https://www.' + host.replace(/^smile\./, '') + '/gp/aws/cart/add.html?ASIN.1=' + m[1].toUpperCase() + '&Quantity.1=' + packs, cart: true, store: 'Amazon' };
      return { url: u.href, cart: false, store: 'Amazon' };
    }
    if (host === 'walmart.com') {
      m = /\/ip\/(?:[^/]+\/)?(\d{5,})(?:[/?]|$)/.exec(u.pathname + '/');
      if (m) return { url: 'https://www.walmart.com/sc/cart/addToCart?items=' + m[1] + (packs > 1 ? encodeURIComponent('|' + packs) : ''), cart: true, store: 'Walmart' };
      return { url: u.href, cart: false, store: 'Walmart' };
    }
    var names = { 'staples.com': 'Staples', 'staples.ca': 'Staples', 'officedepot.com': 'Office Depot', 'target.com': 'Target', 'avery.com': 'Avery', 'avery.ca': 'Avery', 'costco.com': 'Costco', 'costco.ca': 'Costco', 'bestbuy.com': 'Best Buy', 'samsclub.com': 'Sam’s Club', 'bjs.com': 'BJ’s', 'uline.com': 'Uline' };
    return { url: u.href, cart: false, store: names[host] || host };
  }

  // ---------- numbers, QR codes and barcodes ----------
  // "Box {n} of {total}" -> "Box 3 of 12". pad: digits for {n}.
  function numberText(pattern, n, totalN, pad) {
    var num = String(n);
    if (pad > num.length) num = new Array(pad - num.length + 1).join('0') + num;
    return String(pattern || '{n}').replace(/\{n\}/gi, num).replace(/\{total\}/gi, String(totalN));
  }
  // QR code modules (true = dark) using qrcode-generator (MIT, Kazuhiko Arase),
  // loaded separately as the global "qrcode". Error correction M; UTF-8.
  function qrMatrix(text) {
    var lib = root.qrcode;
    if (!lib) throw new Error('The QR code maker didn’t load.');
    lib.stringToBytes = lib.stringToBytesFuncs['UTF-8'];
    var q = lib(0, 'M');
    q.addData(String(text), 'Byte');
    q.make();
    var n = q.getModuleCount(), rows = [];
    for (var r = 0; r < n; r++) { var row = []; for (var c = 0; c < n; c++) row.push(q.isDark(r, c)); rows.push(row); }
    return rows;
  }
  // Code 128 bar patterns (1 = bar, 0 = space), the same table JsBarcode
  // (MIT) uses. Index 103-105 are the start codes, 106 the stop.
  var C128 = [
    '11011001100', '11001101100', '11001100110', '10010011000', '10010001100', '10001001100', '10011001000', '10011000100', '10001100100', '11001001000',
    '11001000100', '11000100100', '10110011100', '10011011100', '10011001110', '10111001100', '10011101100', '10011100110', '11001110010', '11001011100',
    '11001001110', '11011100100', '11001110100', '11101101110', '11101001100', '11100101100', '11100100110', '11101100100', '11100110100', '11100110010',
    '11011011000', '11011000110', '11000110110', '10100011000', '10001011000', '10001000110', '10110001000', '10001101000', '10001100010', '11010001000',
    '11000101000', '11000100010', '10110111000', '10110001110', '10001101110', '10111011000', '10111000110', '10001110110', '11101110110', '11010001110',
    '11000101110', '11011101000', '11011100010', '11011101110', '11101011000', '11101000110', '11100010110', '11101101000', '11101100010', '11100011010',
    '11101111010', '11001000010', '11110001010', '10100110000', '10100001100', '10010110000', '10010000110', '10000101100', '10000100110', '10110010000',
    '10110000100', '10011010000', '10011000010', '10000110100', '10000110010', '11000010010', '11001010000', '11110111010', '11000010100', '10001111010',
    '10100111100', '10010111100', '10010011110', '10111100100', '10011110100', '10011110010', '11110100100', '11110010100', '11110010010', '11011011110',
    '11011110110', '11110110110', '10101111000', '10100011110', '10001011110', '10111101000', '10111100010', '11110101000', '11110100010', '10111011110',
    '10111101110', '11101011110', '11110101110', '11010000100', '11010010000', '11010011100', '1100011101011'
  ];
  // Text -> the barcode's modules as a string of 1s and 0s, or null if the
  // text has characters Code 128 can't hold. All-digit text of even length
  // uses code set C (half as wide); anything else uses set B.
  function code128(text) {
    text = String(text);
    if (!text) return null;
    var vals = [], start;
    if (/^\d+$/.test(text) && text.length % 2 === 0) {
      start = 105;
      for (var i = 0; i < text.length; i += 2) vals.push(Number(text.substr(i, 2)));
    } else {
      start = 104;
      for (var j = 0; j < text.length; j++) {
        var c = text.charCodeAt(j);
        if (c < 32 || c > 126) return null;
        vals.push(c - 32);
      }
    }
    var sum = start;
    vals.forEach(function (v, k) { sum += v * (k + 1); });
    return C128[start] + vals.map(function (v) { return C128[v]; }).join('') + C128[sum % 103] + C128[106];
  }

  // ---------- measuring a scanned test page ----------
  // The outline of the label grid on a scan of the test page: the centers of
  // its outermost long lines, in pixels. gray: one byte per pixel.
  function findGrid(gray, w, h) {
    var rows = new Float32Array(h), cols = new Float32Array(w);
    for (var y = 0; y < h; y++) {
      for (var x = 0; x < w; x++) {
        if (gray[y * w + x] < 160) { rows[y]++; cols[x]++; }
      }
    }
    // The first (or last) thin dark line. Runs that touch the edge of the
    // scan or are thick (the scanner lid's shadow, a dark border) are skipped.
    var line = function (counts, len, fromEnd) {
      var n = counts.length, maxRun = Math.max(3, n * 0.012), i = fromEnd ? n - 1 : 0, step = fromEnd ? -1 : 1;
      while (i >= 0 && i < n) {
        if (counts[i] / len > 0.35) {
          var a = i, b = i;
          while (b + step >= 0 && b + step < n && counts[b + step] / len > 0.35) b += step;
          var lo = Math.min(a, b), hi = Math.max(a, b);
          if (lo > 0 && hi < n - 1 && hi - lo + 1 <= maxRun) return (lo + hi) / 2;
          i = b + step;
          continue;
        }
        i += step;
      }
      return -1;
    };
    var top = line(rows, w, false), bottom = line(rows, w, true), left = line(cols, h, false), right = line(cols, h, true);
    if (top < 0 || left < 0 || bottom <= top || right <= left) return null;
    return { left: left, top: top, right: right, bottom: bottom };
  }
  // From the grid found on a scan (in pixels) -> the printer adjustment.
  // f: the label size; nudge: the adjustment the test page was made with.
  // Returns { nudge, error, widthRatio } in inches, or { problem } text.
  function alignFromScan(grid, imgW, imgH, f, nudge) {
    nudge = nudge || { x: 0, y: 0 };
    var pageW = f.page.width, pageH = f.page.height;
    if (Math.abs(imgH / imgW - pageH / pageW) > 0.04) return { problem: 'The scan isn’t the shape of a whole ' + pageW + ' × ' + pageH + ' in page. Scan the whole page at ' + (pageW === 8.5 && pageH === 11 ? 'Letter' : 'full-page') + ' size.' };
    var ppi = imgW / pageW;
    var expW = (f.across - 1) * f.pitch.across + f.label.width;
    var gotW = (grid.right - grid.left) / ppi, ratio = gotW / expW;
    if (Math.abs(ratio - 1) > 0.03) {
      return { problem: ratio < 0.85 || ratio > 1.15 ? 'This doesn’t look like the test page for this size. Make a test page from this panel, print it, and scan that.'
        : 'The test page printed ' + Math.round(ratio * 100) + '% of full size. Print it again at 100% scale (turn off “Scale to fit” in the print settings), then scan it.' };
    }
    var ex = f.firstLabel.left + nudge.x, ey = f.firstLabel.top + nudge.y;
    var err = { x: grid.left / ppi - ex, y: grid.top / ppi - ey };
    if (Math.abs(err.x) > 0.5 || Math.abs(err.y) > 0.5) return { problem: 'The labels on the scan are more than half an inch from where they should be. Check that the page was scanned right side up, against the corner of the glass.' };
    var round = function (v) { return Math.round(v * 100) / 100 || 0; };
    return { nudge: { x: round(nudge.x - err.x), y: round(nudge.y - err.y) }, error: { x: round(err.x), y: round(err.y) }, widthRatio: ratio };
  }

  // Label content (a flat OOXML package) that is nothing but a table, as
  // version 1.3's Copy left behind: replace the table with the content of its
  // non-empty cells. Anything else comes back unchanged.
  var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  function kids(node, name) {
    return Array.prototype.filter.call(node.childNodes, function (k) { return k.nodeType === 1 && k.namespaceURI === W && (!name || k.localName === name); });
  }
  function hasContent(node) {
    if (Array.prototype.some.call(node.getElementsByTagNameNS(W, 't'), function (t) { return /\S/.test(t.textContent); })) return true;
    return ['drawing', 'pict', 'object'].some(function (n) { return node.getElementsByTagNameNS(W, n).length > 0; });
  }
  function unwrapTable(ooxml) {
    if (!ooxml || ooxml.indexOf('<w:tbl') < 0 || typeof DOMParser === 'undefined') return ooxml;
    var doc;
    try { doc = new DOMParser().parseFromString(ooxml, 'application/xml'); } catch (e) { return ooxml; }
    if (!doc || doc.getElementsByTagName('parsererror').length) return ooxml;
    var body = doc.getElementsByTagNameNS(W, 'body')[0];
    if (!body) return ooxml;
    var changed = false;
    for (var round = 0; round < 3; round++) {
      var main = kids(body).filter(function (k) { return k.localName === 'tbl' || k.localName === 'sdt' || (k.localName === 'p' && hasContent(k)); });
      if (main.length !== 1 || main[0].localName !== 'tbl') break;
      var keep = [];
      kids(main[0], 'tr').forEach(function (tr) {
        kids(tr, 'tc').forEach(function (tc) {
          if (hasContent(tc)) kids(tc).forEach(function (k) { if (k.localName !== 'tcPr') keep.push(k); });
        });
      });
      var sect = kids(body, 'sectPr')[0] || null;
      kids(body).forEach(function (k) { if (k !== sect) body.removeChild(k); });
      if (!keep.length) keep.push(doc.createElementNS(W, 'w:p'));
      keep.forEach(function (k) { body.insertBefore(k, sect); });
      changed = true;
    }
    return changed ? new XMLSerializer().serializeToString(doc) : ooxml;
  }

  var api = {
    parseDelimited: parseDelimited, readXlsx: readXlsx, parseVCard: parseVCard,
    headerMap: headerMap, rowsToAddresses: rowsToAddresses, addressLines: addressLines, fixZip: fixZip, linesToFields: linesToFields,
    fontCss: fontCss, textHeight: textHeight, fits: fits, fitScale: fitScale, findLabelBorder: findLabelBorder,
    unwrapTable: unwrapTable, countryOf: countryOf, fixPostal: fixPostal, regionCode: regionCode, lookupPostal: lookupPostal, suggestAddresses: suggestAddresses,
    reorderLink: reorderLink, numberText: numberText, qrMatrix: qrMatrix, code128: code128, findGrid: findGrid, alignFromScan: alignFromScan,
    FIELDS: FIELDS, PROVINCES: PROVINCES
  };
  root.LSK_TOOLS = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
