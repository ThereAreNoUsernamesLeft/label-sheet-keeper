/*
  Label Sheet Keeper — the Word side panel.

  Words used throughout (code and screen):
    label            one spot on the sheet, numbered 1, 2, 3… across each row
    open             blank and still on the sheet
    used             already peeled off (locked and left blank in the document)
    ready to print   has text or a picture that hasn't been printed yet
    trip             one pass of the physical sheet through the printer
*/
(function () {
  'use strict';

  var VERSION = '1.3';
  var ADDIN_ID = '8d491c85-a76d-4d26-a6f5-b7b2dd4a1711';   // must match manifest.xml <Id>
  var F = window.LSK_FORMATS;
  var S = window.LSK_SHEET;
  var T = window.LSK_TOOLS;

  // Stored in the document (Office settings travel with the file)
  var DOC = {
    history: 'lskHistory',     // [{at, nums, text}]
    lastPrint: 'lskLastPrint', // last "printed" time the panel has seen
    trips: 'lskTrips',         // trips this sheet has made through the printer
    spots: 'lskSpots',         // {n: {at, t, m}} when each used label was printed or marked
    snap: 'lskSnapshot',       // {u: [used], r: {n: first line}} for repairs
    id: 'lskSheetId',
    autoShow: 'Office.AutoShowTaskpaneWithDocument'
  };
  // Stored on this Mac
  var LS = {
    from: 'lsk.fromAddress', autoMark: 'lsk.autoMark', side: 'lsk.side.', fillTab: 'lsk.fillTab', shipTab: 'lsk.shipTab',
    addinRef: 'lsk.addinRef', feedArrow: 'lsk.feedArrow', tripLimit: 'lsk.tripLimit', nudge: 'lsk.nudge.', testCheck: 'lsk.testCheck.',
    warnPass: 'lsk.warnPassHidden', last: 'lsk.last.', book: 'lsk.addressBook', design: 'lsk.design', logo: 'lsk.logo',
    custom: 'lsk.customFormats', sheets: 'lsk.sheets', tableInfo: 'lsk.tableInfo', tableDraft: 'lsk.tableDraft'
  };
  var USED_TAG = 'lsk-used';

  // A missing element (for example while Word still shows an older copy of
  // the panel page after an update) gets a stand-in so the panel keeps working.
  function $(id) { return document.getElementById(id) || document.createElement('div'); }

  var st = { kind: 'loading', fmt: null, used: {}, ready: {}, texts: {}, pics: {}, over: {}, here: 0, header: null, margins: null, applied: null };
  var undo = null;            // last print marked used: { nums, cells: [{n, ooxml}], at, trips }
  var pendingConfirm = null;  // action for the confirm bar's main button
  var choosing = false;       // size chooser opened from a sheet
  var chosenId = F.formats.length ? F.formats[0].id : '';
  var chain = Promise.resolve();
  var busy = false;

  // ---------- small helpers ----------
  function store(key, value) {
    try {
      if (value === undefined) return localStorage.getItem(key);
      if (value === null) localStorage.removeItem(key); else localStorage.setItem(key, value);
    } catch (e) { /* storage full or blocked */ }
    return null;
  }
  function storeJson(key, value) {
    if (value === undefined) { try { return JSON.parse(store(key) || 'null'); } catch (e) { return null; } }
    store(key, value === null ? null : JSON.stringify(value));
    return null;
  }
  function plural(n, w, suffix) { return n + ' ' + w + (n === 1 ? '' : (suffix || 's')); }
  function cap(t) { return t.charAt(0).toUpperCase() + t.slice(1); }
  function ranges(list) {
    var a = list.slice().sort(function (x, y) { return x - y; }), out = [], i = 0;
    while (i < a.length) {
      var j = i;
      while (j + 1 < a.length && a[j + 1] === a[j] + 1) j++;
      out.push(i === j ? String(a[i]) : a[i] + '–' + a[j]);
      i = j + 1;
    }
    return out.join(', ');
  }
  function labelList(list) { return (list.length === 1 ? 'label ' : 'labels ') + ranges(list); }
  function total() { return st.fmt ? st.fmt.across * st.fmt.down : 0; }
  function keys(obj) { return Object.keys(obj || {}).map(Number).sort(function (a, b) { return a - b; }); }
  function openList() { var a = []; for (var n = 1; n <= total(); n++) if (!st.used[n] && !st.ready[n]) a.push(n); return a; }
  function nextOpen() { var o = openList(); return o.length ? o[0] : 0; }
  function userError(msg) { var e = new Error(msg); e.userMessage = msg; return e; }
  function fmtDate(ms) {
    try { return new Date(ms).toLocaleString(undefined, { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' }); }
    catch (e) { return String(new Date(ms)); }
  }
  function fmtDay(ms) {
    try { return new Date(ms).toLocaleDateString(undefined, { month: 'short', day: 'numeric' }); } catch (e) { return ''; }
  }
  function timeOnly(ms) {
    try { return new Date(ms).toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' }); } catch (e) { return ''; }
  }
  function firstLine(t) {
    var l = String(t || '').split(/[\r\n\u000b]+/).map(function (s) { return s.trim(); }).filter(Boolean);
    return l.length ? l[0].slice(0, 48) : '';
  }
  function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }
  function el(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text !== undefined) e.textContent = text; return e; }
  function settings() { try { return Office.context.document.settings; } catch (e) { return null; } }
  function getSetting(k) { var s = settings(); try { return s ? s.get(k) : null; } catch (e) { return null; } }
  function setSetting(k, v) { var s = settings(); try { if (s) { s.set(k, v); s.saveAsync(function () {}); } } catch (e) { /* ignore */ } }
  function docUrl() { try { return Office.context.document.url || ''; } catch (e) { return ''; } }
  function fileName(url) {
    url = url || docUrl();
    if (!url) return '';
    var name = url.split(/[\\/]/).pop();
    try { name = decodeURIComponent(name); } catch (e) { /* keep as is */ }
    return name;
  }

  // ---------- label sizes ----------
  function customFormats() {
    var list = storeJson(LS.custom) || [];
    return list.filter(function (f) { return f && !F.problems(f).length; }).map(function (f) { f.custom = true; return F.withDefaults(f); });
  }
  function allFormats() { return F.formats.concat(customFormats()); }
  function findFormat(id) { return allFormats().filter(function (f) { return f.id === id; })[0] || null; }
  // A sheet's own copy of its measurements (stored when it was made).
  function formatFromSpec(meta) {
    var json = '';
    for (var i = 1; meta[S.SPEC_PROP + i] !== undefined; i++) json += meta[S.SPEC_PROP + i];
    if (!json) return null;
    try { var f = JSON.parse(json); return F.problems(f).length ? null : F.withDefaults(f); } catch (e) { return null; }
  }
  function sheetLayout(meta, f) {
    try {
      var l = JSON.parse(meta[S.LAYOUT_PROP] || 'null');
      if (l && typeof l.top === 'number') return { margins: { top: l.top, left: l.left }, applied: { x: l.nx || 0, y: l.ny || 0 } };
    } catch (e) { /* older sheet */ }
    // Sheets from before version 1.3 used the first label's position as the margins.
    return { margins: { top: f.firstLabel.top, left: f.firstLabel.left }, applied: { x: 0, y: 0 } };
  }

  // Per-Mac settings
  function feedArrowOn() { return store(LS.feedArrow) !== 'off'; }
  function autoMarkOn() { return store(LS.autoMark) !== 'off'; }
  function tripLimit() { var n = Math.round(Number(store(LS.tripLimit))); return n >= 1 && n <= 20 ? n : 3; }
  function nudgeFor(id) { var v = storeJson(LS.nudge + id); return v && typeof v.x === 'number' ? v : { x: 0, y: 0 }; }
  function trips() { return Number(getSetting(DOC.trips)) || 0; }
  function spots() { var s = getSetting(DOC.spots); return s && typeof s === 'object' ? s : {}; }
  function history() { var h = getSetting(DOC.history); return Array.isArray(h) ? h : []; }

  // Sideways 2-per-sheet labels can be called "top" and "bottom" once the
  // person says which half of the physical sheet label 1 printed on.
  function sideApplies(f) { return !!f && f.page.width > f.page.height && f.across === 2 && f.down === 1; }
  function getSide(f) { return sideApplies(f) ? (store(LS.side + f.id) || '') : ''; }
  function labelName(n) {
    var f = st.fmt;
    if (!sideApplies(f)) return 'label ' + n;
    var side = getSide(f);
    if (!side) return 'label ' + n + (n === 1 ? ' (left)' : ' (right)');
    return n === (side === 'top' ? 1 : 2) ? 'top label' : 'bottom label';
  }
  function theName(n) { var nm = labelName(n); return /^label/.test(nm) ? nm : 'the ' + nm; }
  function sideWord(n) { var side = getSide(st.fmt); return side ? (n === (side === 'top' ? 1 : 2) ? 'Top' : 'Bottom') : ''; }

  // ---------- status, errors, confirm ----------
  function say(msg, tone) {
    var e = $('say');
    e.textContent = msg || '';
    e.className = 'say' + (tone === 'warn' ? ' warn' : '');
  }
  function showError(e) {
    var box = $('error');
    if (e && e.userMessage) { say(e.userMessage, 'warn'); box.hidden = true; return; }
    box.textContent = 'Word reported a problem' + (e && e.message ? ': ' + e.message : '') + '. Try again, or close and reopen the panel.';
    box.hidden = false;
    if (window.console) console.error(e, e && e.debugInfo);
  }
  function ask(msg, yesLabel, action, danger) {
    pendingConfirm = action;
    $('confirm-msg').textContent = msg;
    $('confirm-yes').textContent = yesLabel;
    $('confirm-yes').className = 'btn sm ' + (danger ? 'danger' : 'primary');
    $('confirm').className = 'confirm' + (danger ? ' danger' : '');
    $('confirm').hidden = false;
    try { $('confirm').scrollIntoView({ block: 'nearest' }); } catch (e) { /* old browser */ }
    $('confirm-no').focus();
  }
  function closeAsk() { pendingConfirm = null; $('confirm').hidden = true; }

  // Word tasks run one at a time. If one fails, the error is shown and the
  // caller's follow-up steps don't run.
  function queue(task) {
    var p = chain.then(function () { busy = true; $('error').hidden = true; return task(); });
    chain = p.then(function () { busy = false; }, function (e) { busy = false; showError(e); });
    return p.then(function (v) { return v; }, function () { return new Promise(function () {}); });
  }
  // Same, for tasks that change the document: save afterwards.
  function queueChange(task) { return queue(task).then(function (v) { autosave(); return v; }); }

  // ---------- saving ----------
  var saveTimer = 0, needsFirstSave = false;
  // A document that has never been saved is left alone, because saving it
  // would ask for a name; the panel shows a reminder instead.
  function autosave() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () {
      if (!docUrl()) { needsFirstSave = true; $('save-hint').hidden = false; return; }
      Word.run(function (ctx) { ctx.document.save(); return ctx.sync(); }).catch(function () {});
    }, 700);
  }

  // ---------- opening the panel automatically ----------
  // Word opens a panel by itself only in a document that has the add-in
  // attached and tagged. This document gets tagged when the panel reads it.
  // New sheets are built from scratch, so the panel copies the add-in
  // reference Word wrote into this document and builds it into each new sheet.
  function tagForAutoOpen() { if (getSetting(DOC.autoShow) !== true) setSetting(DOC.autoShow, true); }

  var refState = 'unknown';
  function cachedRef() { return storeJson(LS.addinRef); }
  function readOwnFile() {
    return new Promise(function (resolve, reject) {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) { reject(res.error); return; }
        var file = res.value, count = file.sliceCount, parts = [];
        (function next(i) {
          file.getSliceAsync(i, function (r) {
            if (r.status !== Office.AsyncResultStatus.Succeeded) { file.closeAsync(); reject(r.error); return; }
            parts.push(r.value.data);
            if (i + 1 < count) { next(i + 1); return; }
            file.closeAsync();
            var size = parts.reduce(function (a, p) { return a + p.length; }, 0), out = new Uint8Array(size), off = 0;
            parts.forEach(function (p) { out.set(p, off); off += p.length; });
            resolve(out);
          });
        })(0);
      });
    });
  }
  function findAddinRef() {
    var c = cachedRef();
    if (c) { refState = 'found'; return Promise.resolve(c); }
    if (!window.JSZip || !window.DOMParser) { refState = 'failed: file tools didn’t load'; return Promise.resolve(null); }
    return readOwnFile().then(function (bytes) { return window.JSZip.loadAsync(bytes); }).then(function (zip) {
      var names = Object.keys(zip.files).filter(function (n) { return /^word\/webextensions\/webextension\d*\.xml$/i.test(n); });
      return Promise.all(names.map(function (n) { return zip.file(n).async('string'); }));
    }).then(function (xmls) {
      for (var i = 0; i < xmls.length; i++) {
        var doc = new DOMParser().parseFromString(xmls[i], 'application/xml');
        var r = doc.getElementsByTagNameNS(S.WE_NS, 'reference')[0];
        if (r && (r.getAttribute('id') || '').toLowerCase().indexOf(ADDIN_ID) >= 0) {
          var ref = { id: r.getAttribute('id'), version: r.getAttribute('version') || '1.0.0.0', store: r.getAttribute('store') || '', storeType: r.getAttribute('storeType') || '' };
          storeJson(LS.addinRef, ref);
          refState = 'found';
          return ref;
        }
      }
      refState = 'missing';
      return null;
    }).catch(function (e) {
      refState = 'failed: ' + ((e && e.message) || 'couldn’t read this document');
      return null;
    });
  }

  // ---------- will the text fit? ----------
  var measureCtx = null;
  function measure(text, css) {
    if (!measureCtx) measureCtx = document.createElement('canvas').getContext('2d');
    measureCtx.font = css;
    return measureCtx.measureText(text).width * 0.75;   // CSS px -> points
  }
  function rowHeightIn(f) {
    var minRow = Infinity;
    S.layout(f).rows.forEach(function (r) { if (r.label) minRow = Math.min(minRow, r.h / 1440); });
    return minRow;
  }
  function textBox(f) {
    return {
      width: (f.label.width - f.padding.left - f.padding.right) * 72 - 1,
      height: rowHeightIn(f) * 72 - (f.align === 'top' ? 0.12 * 72 : 4)
    };
  }
  var PARA_PROPS = 'items/text,items/font/name,items/font/size,items/font/bold,items/font/italic,items/spaceBefore,items/spaceAfter,items/leftIndent,items/lineSpacing';
  function paraSpecs(items, f) {
    return items.map(function (p) {
      return {
        text: p.text, font: p.font.name || f.font.name, size: p.font.size || f.font.size, bold: !!p.font.bold, italic: !!p.font.italic,
        before: p.spaceBefore || 0, after: p.spaceAfter || 0, indent: Math.max(0, p.leftIndent || 0), lineSpacing: p.lineSpacing || 12
      };
    });
  }

  // ---------- reading the document ----------
  function readState() {
    return Word.run(function (ctx) {
      var props = ctx.document.properties.customProperties;
      props.load('items/key,items/value');
      var tables = ctx.document.body.tables;
      tables.load('items/rowCount');
      var paras = ctx.document.body.paragraphs;
      paras.load('items/tableNestingLevel,items/text');
      var bodyCCs = ctx.document.body.contentControls;
      bodyCCs.load('items/tag');
      var header = ctx.document.sections.getFirst().getHeader('Primary');
      header.load('text');
      return ctx.sync().then(function () {
        var meta = {};
        props.items.forEach(function (p) { meta[p.key] = String(p.value); });
        var id = meta[S.PROP];
        if (!id) return { kind: 'none' };
        var fmt = findFormat(id) || formatFromSpec(meta);
        if (!fmt) {
          return { kind: 'broken', fixable: false, reason: 'This sheet uses a label size (“' + id + '”) that isn’t on this Mac anymore. Make a new sheet, or add the size again.' };
        }
        var lay = sheetLayout(meta, fmt);
        var shape = S.tableShape(fmt);
        var spacers = bodyCCs.items.filter(function (cc) { return cc.tag === S.SPACER_TAG; }).length;
        var outside = paras.items.filter(function (p) { return p.tableNestingLevel === 0; });
        var damage = '';
        if (!tables.items.length) damage = 'The labels are missing from this document.';
        else if (tables.items[0].rowCount !== shape.rows) damage = 'Rows of labels were added or removed, so the labels no longer line up with your sheet.';
        else if (outside.length > spacers + 1 || outside.some(function (p) { return p.text.trim(); })) damage = 'There’s text or a blank line outside the labels, which pushes them out of place.';
        var base = { fmt: fmt, margins: lay.margins, applied: lay.applied, header: /▲/.test(header.text || ''), headerOther: !!(header.text || '').trim() && !/▲/.test(header.text || '') };
        if (damage) { base.kind = 'broken'; base.fixable = true; base.reason = damage; return base; }

        var table = tables.items[0];
        table.load('values');
        table.rows.load('items/cellCount');
        var pics = ctx.document.body.inlinePictures; pics.load('items');
        var selCell = ctx.document.getSelection().parentTableCellOrNullObject; selCell.load('rowIndex,cellIndex');
        return ctx.sync().then(function () {
          if (table.rows.items.some(function (r) { return r.cellCount !== shape.cols; })) {
            base.kind = 'broken'; base.fixable = true; base.reason = 'Labels were added, removed or merged in a row, so they no longer line up with your sheet.';
            return base;
          }
          var usedCells = [], picCells = [];
          bodyCCs.items.forEach(function (cc) {
            if ((cc.tag || '').indexOf(USED_TAG) === 0) { var c = cc.parentTableCellOrNullObject; c.load('rowIndex,cellIndex'); usedCells.push(c); }
          });
          pics.items.forEach(function (p) { var c = p.parentTableCellOrNullObject; c.load('rowIndex,cellIndex'); picCells.push(c); });
          return ctx.sync().then(function () {
            var used = {}, ready = {}, texts = {}, picsAt = {};
            usedCells.forEach(function (c) { if (!c.isNullObject) { var n = S.labelAt(fmt, c.rowIndex, c.cellIndex); if (n) used[n] = true; } });
            var vals = table.values || [];
            for (var n = 1; n <= fmt.across * fmt.down; n++) {
              if (used[n]) continue;
              var ci = S.cellOf(fmt, n);
              var t = String((vals[ci.row] && vals[ci.row][ci.cell]) || '').replace(/ /g, ' ').trim();
              if (t) { ready[n] = true; texts[n] = t; }
            }
            picCells.forEach(function (c) {
              if (c.isNullObject) return;
              var k = S.labelAt(fmt, c.rowIndex, c.cellIndex);
              if (k && !used[k]) { ready[k] = true; picsAt[k] = true; texts[k] = texts[k] || '(picture)'; }
            });
            var here = selCell.isNullObject ? 0 : S.labelAt(fmt, selCell.rowIndex, selCell.cellIndex);
            // Check whether text fits on each label that's ready to print.
            var check = keys(ready).filter(function (k) { return !picsAt[k]; });
            var lists = check.map(function (k) {
              var cell = S.cellOf(fmt, k), p = table.getCell(cell.row, cell.cell).body.paragraphs;
              p.load(PARA_PROPS);
              return p;
            });
            return ctx.sync().then(function () {
              var over = {}, box = textBox(fmt);
              check.forEach(function (k, i) { try { if (!T.fits(paraSpecs(lists[i].items, fmt), box, measure)) over[k] = true; } catch (e) { /* skip */ } });
              base.kind = 'sheet';
              base.used = used; base.ready = ready; base.texts = texts; base.pics = picsAt; base.over = over; base.here = here;
              return base;
            });
          });
        });
      });
    });
  }

  var lastSnap = '';
  function apply(s) {
    st = s;
    if (s.kind === 'sheet') {
      tagForAutoOpen();
      if (!getSetting(DOC.id)) setSetting(DOC.id, uid());
      var snap = JSON.stringify({ u: keys(st.used), r: keys(st.ready).reduce(function (o, n) { o[n] = firstLine(st.texts[n]); return o; }, {}) });
      if (snap !== lastSnap) { lastSnap = snap; setSetting(DOC.snap, JSON.parse(snap)); }
      updateRegistry();
      syncHeader();
    } else if (s.kind !== 'broken') st.fmt = null;
    render();
  }
  function refresh() { return queue(function () { return readState().then(apply); }); }

  var selTimer = 0;
  function onSelectionChanged() {
    clearTimeout(selTimer);
    selTimer = setTimeout(function () { if (!busy) refresh(); }, 400);
  }

  // ---------- the "▲ top" mark in the sheet's blank top strip ----------
  function arrowFits() { return !!st.fmt && F.hasStrip(st.fmt) && st.margins && st.margins.top >= S.STRIP_MARGIN - 0.001; }
  function setHeader(on) {
    return Word.run(function (ctx) {
      var h = ctx.document.sections.getFirst().getHeader('Primary');
      h.clear();
      if (on) {
        var p = h.paragraphs.getFirst();
        p.insertText(S.ARROW_TEXT, 'Replace');
        p.alignment = 'Centered';
        p.spaceBefore = S.ARROW_BEFORE_PT;
        p.spaceAfter = 0;
        p.font.set({ name: 'Arial', size: S.ARROW_SIZE_PT, bold: true, italic: false, color: '#555555' });
      }
      return ctx.sync();
    });
  }
  var headerBusy = false;
  function syncHeader() {
    if (headerBusy || !arrowFits() || st.headerOther) return;
    var want = feedArrowOn();
    if (want === st.header) return;
    headerBusy = true;
    queueChange(function () { return setHeader(want); }).then(function () { st.header = want; headerBusy = false; render(); });
  }

  // ---------- your sheets (kept on this Mac) ----------
  function registry() { return storeJson(LS.sheets) || {}; }
  function updateRegistry() {
    var id = getSetting(DOC.id);
    if (!id || !st.fmt) return;
    var reg = registry(), open = openList().length, ready = keys(st.ready).length;
    if (!open && !ready) delete reg[id];
    else reg[id] = { id: id, fmtId: st.fmt.id, fmtName: st.fmt.name, file: fileName(), url: docUrl(), open: open, ready: ready, trips: trips(), at: Date.now() };
    storeJson(LS.sheets, reg);
  }
  function sheetEntries(exceptCurrent) {
    var reg = registry(), me = getSetting(DOC.id);
    return Object.keys(reg).map(function (k) { return reg[k]; })
      .filter(function (e) { return e && (e.open > 0 || e.ready > 0) && !(exceptCurrent && e.id === me); })
      .sort(function (a, b) { return a.fmtName < b.fmtName ? -1 : a.fmtName > b.fmtName ? 1 : b.at - a.at; });
  }
  function openSheetFile(entry) {
    queue(function () {
      return Word.run(function (ctx) {
        if (!entry.url || typeof ctx.application.openDocument !== 'function') throw userError('Open “' + (entry.file || 'that sheet') + '” from File › Open Recent in Word.');
        ctx.application.openDocument(entry.url);
        return ctx.sync();
      }).catch(function (e) {
        if (e && e.userMessage) throw e;
        throw userError('Word couldn’t open it from here. Open “' + (entry.file || 'that sheet') + '” from File › Open Recent.');
      });
    });
  }
  function renderSheetList(ul, entries, showCurrent) {
    ul.innerHTML = '';
    var me = getSetting(DOC.id), limit = tripLimit();
    entries.forEach(function (e) {
      var li = el('li', showCurrent && e.id === me ? 'here' : '');
      li.appendChild(el('b', '', e.fmtName));
      var t = Math.min(e.trips || 0, limit);
      li.appendChild(el('span', 'meta', plural(e.open, 'open label') + (e.ready ? ' · ' + e.ready + ' ready to print' : '') +
        ' · ' + t + ' of ' + limit + ' trips · ' + (showCurrent && e.id === me ? 'this sheet' : (e.file || 'not saved yet') + ' · ' + fmtDay(e.at))));
      var acts = el('span', 'acts');
      if (!(showCurrent && e.id === me)) {
        var open = el('button', 'btn ghost sm', 'Open'); open.type = 'button';
        open.addEventListener('click', function () { openSheetFile(e); });
        var x = el('button', 'btn link', 'Remove'); x.type = 'button'; x.title = 'Remove from this list';
        x.addEventListener('click', function () { var reg = registry(); delete reg[e.id]; storeJson(LS.sheets, reg); render(); });
        acts.appendChild(open); acts.appendChild(x);
      }
      li.appendChild(acts);
      ul.appendChild(li);
    });
  }

  // ---------- changing the document ----------
  function cellBody(ctx, n) {
    var ci = S.cellOf(st.fmt, n);
    return ctx.document.body.tables.getFirst().getCell(ci.row, ci.cell).body;
  }

  // A used label is emptied and locked with a tagged content control, so it
  // prints nothing and can't be typed in by accident.
  function lockAsUsed(ctx, nums) {
    var made = nums.map(function (n) {
      var body = cellBody(ctx, n);
      body.clear();
      var cc = body.insertText(' ', 'Start').insertContentControl();
      cc.tag = USED_TAG;
      cc.title = 'Used';
      cc.cannotEdit = true;
      cc.cannotDelete = true;
      return cc;
    });
    return ctx.sync().then(function () {
      // On-screen look only (tags and colors never print). Skipped if this Word can't show them.
      made.forEach(function (cc) { cc.appearance = 'Tags'; cc.color = '#7F8C8D'; });
      return ctx.sync().catch(function () {});
    });
  }
  function unlock(ctx, nums) {
    var lists = nums.map(function (n) { var cs = cellBody(ctx, n).contentControls; cs.load('items/tag'); return cs; });
    return ctx.sync().then(function () {
      var mine = [];
      lists.forEach(function (cs) {
        cs.items.forEach(function (cc) {
          if ((cc.tag || '').indexOf(USED_TAG) === 0) { cc.cannotDelete = false; cc.cannotEdit = false; mine.push(cc); }
        });
      });
      return ctx.sync().then(function () { mine.forEach(function (cc) { cc.delete(false); }); return ctx.sync(); });
    });
  }
  function setSpots(nums, rec) {
    var s = spots();
    nums.forEach(function (n) { if (rec) s[n] = rec(n); else delete s[n]; });
    setSetting(DOC.spots, s);
  }
  function markUsed(nums) {
    return Word.run(function (ctx) { return lockAsUsed(ctx, nums); }).then(function () {
      var at = Date.now();
      setSpots(nums, function () { return { at: at, m: 1 }; });
    });
  }
  function markOpen(nums) {
    return Word.run(function (ctx) { return unlock(ctx, nums); }).then(function () { setSpots(nums, null); });
  }

  // Remove the empty paragraph Word sometimes adds after pasted content.
  function trimAfterPaste(ctx, bodies, wantCount) {
    var lists = bodies.map(function (b) { var p = b.paragraphs; p.load('items/text'); return p; });
    return ctx.sync().then(function () {
      lists.forEach(function (p) {
        if (p.items.length > wantCount) { var last = p.items[p.items.length - 1]; if (!last.text) last.delete(); }
      });
      return ctx.sync();
    });
  }

  // Copy one label's content (by number) to other labels.
  function copyFrom(src, targets) {
    return Word.run(function (ctx) {
      var sb = cellBody(ctx, src), ox = sb.getOoxml(), sp = sb.paragraphs;
      sp.load('items');
      return ctx.sync().then(function () {
        var bodies = targets.map(function (t) { var b = cellBody(ctx, t); b.insertOoxml(ox.value, 'Replace'); return b; });
        return ctx.sync().then(function () { return trimAfterPaste(ctx, bodies, sp.items.length); });
      });
    });
  }
  // The label the cursor is in, if it's ready to print.
  function cursorLabel() {
    return Word.run(function (ctx) {
      var c = ctx.document.getSelection().parentTableCellOrNullObject;
      c.load('rowIndex,cellIndex');
      return ctx.sync().then(function () {
        if (c.isNullObject) throw userError('Click inside a label in the document first.');
        var n = S.labelAt(st.fmt, c.rowIndex, c.cellIndex);
        if (!n) throw userError('The cursor is in the gap between labels. Click inside a label.');
        if (st.used[n]) throw userError(cap(labelName(n)) + ' is used. Click inside the label that has your text.');
        if (!st.ready[n]) throw userError(cap(labelName(n)) + ' is empty. Type your label there first.');
        return n;
      });
    });
  }
  // The label's text and basic formatting, for building new sheets.
  function readLabelParas(n) {
    return Word.run(function (ctx) {
      var p = cellBody(ctx, n).paragraphs;
      p.load('items/text,items/alignment,items/font/name,items/font/size,items/font/bold,items/font/italic');
      return ctx.sync().then(function () {
        return p.items.map(function (q) {
          var a = String(q.alignment || '').toLowerCase();
          return { t: q.text, size: q.font.size || st.fmt.font.size, b: !!q.font.bold, i: !!q.font.italic, font: q.font.name || undefined, align: a === 'centered' ? 'center' : a === 'right' ? 'right' : 'left' };
        });
      });
    });
  }

  // Write lines of text (and optionally a picture) into labels.
  // paras: [{ t, size, b, i, color, align, indent, before }]; pic: { b64, w, h, align } in points
  function writeLabels(nums, paras, pic) {
    return writeEach(nums.map(function (n) { return { n: n, paras: paras, pic: pic }; }));
  }
  // items: [{ n, paras, pic }] — each label gets its own content, in one pass.
  function writeEach(items) {
    var f = st.fmt, topPt = f.align === 'top' ? f.topSpace * 72 : 0;
    var alignOf = function (a) { return a === 'center' ? 'Centered' : a === 'right' ? 'Right' : 'Left'; };
    return Word.run(function (ctx) {
      items.forEach(function (item) {
        var n = item.n, paras = item.paras || [], pic = item.pic;
        var body = cellBody(ctx, n);
        body.clear();
        var first = body.paragraphs.getFirst();
        if (pic) {
          var img = first.insertInlinePictureFromBase64(pic.b64, 'Start');
          img.width = pic.w; img.height = pic.h;
          first.alignment = alignOf(pic.align);
          first.spaceBefore = topPt; first.spaceAfter = 2;
        }
        paras.forEach(function (p, i) {
          var para;
          if (i === 0 && !pic) { para = first; para.insertText(p.t, 'Replace'); }
          else para = body.insertParagraph(p.t, 'End');
          para.alignment = alignOf(p.align);
          para.spaceBefore = p.before !== undefined ? p.before : (i === 0 && !pic ? topPt : 0);
          para.spaceAfter = 0;
          para.leftIndent = p.indent || 0;
          para.firstLineIndent = 0;
          para.font.set({ name: p.font || f.font.name, size: p.size || f.font.size, bold: !!p.b, italic: !!p.i, color: p.color || '#000000' });
        });
      });
      return ctx.sync();
    });
  }

  // Mark every ready-to-print label used: one more trip for the sheet.
  function markPrinted() {
    var nums = keys(st.ready);
    return Word.run(function (ctx) {
      var oxs = nums.map(function (n) { return cellBody(ctx, n).getOoxml(); });
      return ctx.sync().then(function () {
        undo = { nums: nums, cells: nums.map(function (n, i) { return { n: n, ooxml: oxs[i].value }; }), at: Date.now() };
        return lockAsUsed(ctx, nums);
      });
    }).then(function () {
      var at = undo.at, texts = [];
      nums.forEach(function (n) { var t = firstLine(st.texts[n]); if (t && texts.indexOf(t) < 0) texts.push(t); });
      var h = history();
      h.unshift({ at: at, nums: nums, text: texts.length ? texts[0] + (texts.length > 1 ? ' and ' + plural(texts.length - 1, 'other') : '') : '' });
      setSetting(DOC.history, h.slice(0, 200));
      setSetting(DOC.trips, trips() + 1);
      setSpots(nums, function (n) { return { at: at, t: firstLine(st.texts[n]) }; });
      // Keep the last label for "Reprint last label".
      var last = undo.cells[undo.cells.length - 1];
      if (last) {
        store(LS.last + st.fmt.id, null);
        storeJson(LS.last + st.fmt.id, { ooxml: last.ooxml.length < 3e6 ? last.ooxml : '', paras: (last.ooxml.match(/<w:p[ >]/g) || []).length, text: firstLine(st.texts[last.n]) || 'picture', at: at });
      }
      return nums;
    });
  }

  function undoPrinted() {
    var u = undo;
    return Word.run(function (ctx) {
      return unlock(ctx, u.nums).then(function () {
        u.cells.forEach(function (c) { cellBody(ctx, c.n).insertOoxml(c.ooxml, 'Replace'); });
        return ctx.sync();
      });
    }).then(function () {
      var h = history();
      if (h.length && h[0].at === u.at) { h.shift(); setSetting(DOC.history, h); }
      setSetting(DOC.trips, Math.max(0, trips() - 1));
      setSpots(u.nums, null);
      undo = null;
    });
  }

  function reprintLast() {
    var last = storeJson(LS.last + st.fmt.id), n = nextOpen();
    if (!last || !last.ooxml) return Promise.reject(userError('There’s no saved label to reprint yet.'));
    if (!n) return Promise.reject(userError('There are no open labels on this sheet.'));
    return Word.run(function (ctx) {
      var b = cellBody(ctx, n);
      b.insertOoxml(last.ooxml, 'Replace');
      return ctx.sync().then(function () { return trimAfterPaste(ctx, [b], last.paras || 99); });
    }).then(function () { return n; });
  }

  // ---------- making documents ----------
  function createDoc(f, opts) {
    return S.buildDocx(f, null, opts).then(function (b64) {
      return Word.run(function (ctx) {
        var doc = ctx.application.createDocument(b64);
        return ctx.sync().then(function () { doc.open(); return ctx.sync(); });
      });
    });
  }
  // prefills: one { n: [paras] } object per new sheet
  function createSheets(f, prefills) {
    return findAddinRef().then(function (ref) {
      var list = prefills && prefills.length ? prefills : [null];
      return list.reduce(function (p, pre) {
        return p.then(function () { return createDoc(f, { addinRef: ref, feedArrow: feedArrowOn(), nudge: nudgeFor(f.id), prefill: pre || undefined }); });
      }, Promise.resolve()).then(function () { return !!ref; });
    });
  }
  // Split many labels' paragraphs into new sheets.
  function spillPrefills(f, labels) {
    var per = f.across * f.down, out = [];
    for (var i = 0; i < labels.length; i += per) {
      var pre = {};
      labels.slice(i, i + per).forEach(function (paras, j) { pre[j + 1] = paras; });
      out.push(pre);
    }
    return out;
  }
  function makeTestPage(f) {
    return createDoc(f, { test: true, nudge: nudgeFor(f.id) }).then(function () { store(LS.testCheck + f.id, 'pending'); });
  }

  // ---------- repair / rebuild in place ----------
  // Rebuilds the label table (and the spacer above it), keeping used labels
  // and labels that are ready to print. Also applies the printer adjustment.
  function rebuildSheet(f, margins, nudge) {
    var report = { used: 0, ready: 0, lost: 0 };
    return Word.run(function (ctx) {
      var body = ctx.document.body;
      var tables = body.tables; tables.load('items/rowCount');
      var ccs = body.contentControls; ccs.load('items/tag');
      return ctx.sync().then(function () {
        var shape = S.tableShape(f), t = tables.items[0];
        if (t) { t.load('values'); t.rows.load('items/cellCount'); }
        return ctx.sync().then(function () {
          var intact = !!t && t.rowCount === shape.rows && t.rows.items.every(function (r) { return r.cellCount === shape.cols; });
          var snap = getSetting(DOC.snap) || { u: [], r: {} };
          var used = {}, grabbed = [];
          var usedCells = [];
          ccs.items.forEach(function (cc) {
            if ((cc.tag || '').indexOf(USED_TAG) === 0) { var c = cc.parentTableCellOrNullObject; c.load('rowIndex,cellIndex'); usedCells.push(c); }
          });
          return ctx.sync().then(function () {
            if (intact) usedCells.forEach(function (c) { if (!c.isNullObject) { var n = S.labelAt(f, c.rowIndex, c.cellIndex); if (n) used[n] = true; } });
            else (snap.u || []).forEach(function (n) { used[n] = true; });
            // Collect content from cells that have text.
            var vals = t ? (t.values || []) : [];
            if (t) {
              vals.forEach(function (row, ri) {
                (row || []).forEach(function (v, ci) {
                  var text = String(v || '').replace(/ /g, ' ').trim();
                  if (!text) return;
                  var n = intact ? S.labelAt(f, ri, ci) : 0;
                  if (intact && (!n || used[n])) return;
                  grabbed.push({ n: n, text: text, ox: t.getCell(ri, ci).body.getOoxml() });
                });
              });
            }
            return ctx.sync();
          }).then(function () {
            // Without an intact table, match content to labels by its first line.
            var taken = {};
            grabbed.forEach(function (g) {
              if (g.n) { taken[g.n] = true; return; }
              var line = firstLine(g.text);
              Object.keys(snap.r || {}).some(function (k) {
                var n = Number(k);
                if (!taken[n] && !used[n] && snap.r[k] === line) { g.n = n; taken[n] = true; return true; }
                return false;
              });
            });
            grabbed.forEach(function (g) {
              if (g.n) return;
              for (var n = 1; n <= f.across * f.down; n++) if (!used[n] && !taken[n]) { g.n = n; taken[n] = true; return; }
              report.lost++;
            });
            ccs.items.forEach(function (cc) { cc.cannotDelete = false; cc.cannotEdit = false; });
            return ctx.sync();
          }).then(function () {
            body.clear();
            return ctx.sync();
          }).then(function () {
            body.insertOoxml(S.tablePackage(f, { margins: margins, nudge: nudge }), 'Start');
            return ctx.sync();
          }).then(function () {
            // Keep only one (tiny) paragraph after the table.
            var paras = body.paragraphs; paras.load('items/tableNestingLevel,items/text');
            return ctx.sync().then(function () {
              var items = paras.items, lastIn = -1;
              items.forEach(function (p, i) { if (p.tableNestingLevel > 0) lastIn = i; });
              var after = items.filter(function (p, i) { return i > lastIn && p.tableNestingLevel === 0; });
              after.slice(0, -1).forEach(function (p) { if (!p.text) p.delete(); });
              if (after.length) { var lastP = after[after.length - 1]; lastP.font.size = 1; lastP.spaceBefore = 0; lastP.spaceAfter = 0; }
              return ctx.sync();
            });
          }).then(function () {
            var prev = st.fmt; st.fmt = f;
            var usedNums = keys(used);
            return lockAsUsed(ctx, usedNums).then(function () {
              grabbed.forEach(function (g) { if (g.n) cellBody(ctx, g.n).insertOoxml(g.ox.value, 'Replace'); });
              ctx.document.properties.customProperties.add(S.LAYOUT_PROP, JSON.stringify({ v: 2, top: margins.top, left: margins.left, nx: nudge.x, ny: nudge.y }));
              report.used = usedNums.length;
              report.ready = grabbed.filter(function (g) { return g.n; }).length;
              return ctx.sync();
            }).then(function () { st.fmt = prev; }, function (e) { st.fmt = prev; throw e; });
          });
        });
      });
    }).then(function () { return report; });
  }

  // ---------- noticing prints ----------
  // Word stamps a "last printed" time on the document when it prints. The
  // panel checks it every few seconds; when it moves forward, the labels that
  // were ready to print have just been printed.
  var lastPrintSeen = null, watchTimer = 0;
  function readPrintTime() {
    return Word.run(function (ctx) {
      var props = ctx.document.properties;
      props.load('lastPrintDate');
      return ctx.sync().then(function () {
        var ms = props.lastPrintDate ? new Date(props.lastPrintDate).getTime() : 0;
        return isFinite(ms) && ms > 0 && new Date(ms).getFullYear() > 1990 ? ms : 0;
      });
    });
  }
  function checkPrinted() {
    if (busy || st.kind !== 'sheet') return;
    readPrintTime().then(function (ms) {
      if (lastPrintSeen === null) {
        // First check: a print counts as new only if it's later than the last
        // one this document recorded.
        var saved = Number(getSetting(DOC.lastPrint)) || 0;
        lastPrintSeen = saved || ms;
        if (!saved) setSetting(DOC.lastPrint, ms);
        showLastPrint(ms);
        if (!(saved && ms > saved)) return;
      }
      if (ms <= lastPrintSeen) return;
      lastPrintSeen = ms;
      setSetting(DOC.lastPrint, ms);
      showLastPrint(ms);
      onPrinted(ms);
    }).catch(function () { /* best effort; the button still works */ });
  }
  function showLastPrint(ms) { $('print-watch').textContent = ms ? 'Last printed ' + fmtDate(ms) + '.' : 'Not printed yet.'; }
  function onPrinted(ms) {
    queue(function () { return readState().then(apply); }).then(function () {
      var nums = keys(st.ready);
      if (!nums.length) return;
      var when = timeOnly(ms) ? ' at ' + timeOnly(ms) : '';
      var doMark = function () {
        queueChange(markPrinted).then(function (done) {
          return refresh().then(function () {
            say('Word printed this sheet' + when + ', so ' + labelList(done) + (done.length === 1 ? ' is' : ' are') + ' now marked used.');
          });
        });
      };
      if (autoMarkOn()) doMark();
      else ask('Word printed this sheet' + when + '. Mark ' + labelList(nums) + ' used?', 'Mark used', doMark);
    });
  }
  function startPrintWatch() {
    clearInterval(watchTimer);
    watchTimer = setInterval(checkPrinted, 2500);
    document.addEventListener('visibilitychange', function () { if (!document.hidden) checkPrinted(); });
    setTimeout(checkPrinted, 800);
  }

  // ---------- shipping label files ----------
  var DPI = 300;
  var ship = { doc: null, page: 1, pages: 1, src: null, crop: null, turn: 0, found: false };

  // Usable area inside one label, in inches.
  function contentBox(f) {
    var rowH = rowHeightIn(f), top = f.align === 'top' ? f.topSpace : 0.3;
    return { w: f.label.width - f.padding.left - f.padding.right, h: Math.max(0.5, rowH - top - 0.3), top: top, rowH: rowH };
  }

  var pdfLib = null;
  function loadPdfLib() {
    if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
    if (pdfLib) return pdfLib;
    var fail = function () { pdfLib = null; throw userError('The PDF reader didn’t load. Check your internet connection and try again.'); };
    pdfLib = new Promise(function (resolve, reject) {
      var s = document.createElement('script');
      s.src = 'vendor/pdfjs/pdf.js';
      s.onload = function () { if (window.pdfjsLib) resolve(window.pdfjsLib); else reject(); };
      s.onerror = reject;
      document.head.appendChild(s);
    }).then(function (lib) { lib.GlobalWorkerOptions.workerSrc = 'vendor/pdfjs/pdf.worker.js'; return lib; }, fail);
    return pdfLib;
  }
  function fileError(e) {
    say((e && e.userMessage) || 'That file couldn’t be opened. Try a PDF, PNG, or JPEG.', 'warn');
    if (window.console && e && !e.userMessage) console.error(e);
  }
  function handleFile(file) {
    if (!file) return;
    say('Opening ' + (file.name || 'the picture') + '…');
    var isPdf = /pdf$/i.test(file.type || '') || /\.pdf$/i.test(file.name || '');
    var reader = new FileReader();
    reader.onerror = function () { fileError(); };
    if (isPdf) {
      reader.onload = function () {
        loadPdfLib().then(function (lib) {
          return lib.getDocument({ data: new Uint8Array(reader.result), standardFontDataUrl: 'vendor/pdfjs/standard_fonts/', isEvalSupported: false }).promise;
        }).then(function (doc) {
          ship.doc = doc; ship.pages = doc.numPages; ship.page = 1;
          return renderPdfPage();
        }).then(afterSource).catch(fileError);
      };
      reader.readAsArrayBuffer(file);
    } else {
      reader.onload = function () {
        var img = new Image();
        img.onload = function () {
          var c = document.createElement('canvas');
          c.width = img.naturalWidth; c.height = img.naturalHeight;
          var x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(img, 0, 0);
          ship.doc = null; ship.pages = 1; ship.page = 1; ship.src = c;
          afterSource();
        };
        img.onerror = function () { fileError(userError('That picture couldn’t be opened. Try a PDF, PNG, or JPEG.')); };
        img.src = reader.result;
      };
      reader.readAsDataURL(file);
    }
  }
  function renderPdfPage() {
    return ship.doc.getPage(ship.page).then(function (page) {
      var one = page.getViewport({ scale: 1 });
      var scale = DPI / 72, maxPx = 14e6;
      if (one.width * one.height * scale * scale > maxPx) scale = Math.sqrt(maxPx / (one.width * one.height));
      var vp = page.getViewport({ scale: scale });
      var c = document.createElement('canvas');
      c.width = Math.floor(vp.width); c.height = Math.floor(vp.height);
      var ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
      return page.render({ canvasContext: ctx, viewport: vp }).promise.then(function () { ship.src = c; });
    });
  }
  function afterSource() {
    ship.turn = 0;
    var found = findLabel(ship.src);
    ship.crop = found || trimBox(ship.src);
    ship.found = !!found;
    setFillTab('file');
    $('crop').hidden = false;
    drawCrop();
    if (ship.pages > 1 && !found) say('This PDF has ' + ship.pages + ' pages. Use the arrows to find the page with the label.');
    else if (found) say('Found the label' + (ship.pages > 1 ? ' on page ' + ship.page + ' of ' + ship.pages : '') + '. Check the preview, then put it on the sheet.');
    else say('Drag across the label if the page has anything else on it.');
  }
  // Smart crop: the label's printed border, if it has one.
  function findLabel(src) {
    var w = Math.min(900, src.width), k = src.width / w, h = Math.max(1, Math.round(src.height / k));
    var c = document.createElement('canvas'); c.width = w; c.height = h;
    var x = c.getContext('2d'); x.drawImage(src, 0, 0, w, h);
    var d = x.getImageData(0, 0, w, h).data, gray = new Uint8Array(w * h);
    for (var i = 0; i < w * h; i++) gray[i] = (d[i * 4] * 299 + d[i * 4 + 1] * 587 + d[i * 4 + 2] * 114) / 1000;
    var r = T.findLabelBorder(gray, w, h);
    if (!r) return null;
    var pad = 3;
    var bx = Math.max(0, Math.floor((r.x - pad) * k)), by = Math.max(0, Math.floor((r.y - pad) * k));
    return { x: bx, y: by, w: Math.min(src.width - bx, Math.ceil((r.w + 2 * pad) * k)), h: Math.min(src.height - by, Math.ceil((r.h + 2 * pad) * k)) };
  }
  // Box around everything that isn't white, with a small margin.
  function trimBox(src) {
    var w = Math.min(500, src.width), k = src.width / w, h = Math.max(1, Math.round(src.height / k));
    var c = document.createElement('canvas'); c.width = w; c.height = h;
    var x = c.getContext('2d'); x.drawImage(src, 0, 0, w, h);
    var d = x.getImageData(0, 0, w, h).data, minX = w, minY = h, maxX = -1, maxY = -1;
    for (var y = 0; y < h; y++) {
      for (var i = 0; i < w; i++) {
        var o = (y * w + i) * 4;
        if (d[o] + d[o + 1] + d[o + 2] < 690) { if (i < minX) minX = i; if (i > maxX) maxX = i; if (y < minY) minY = y; if (y > maxY) maxY = y; }
      }
    }
    if (maxX < 0) return { x: 0, y: 0, w: src.width, h: src.height };
    var m = 4;
    minX = Math.max(0, minX - m); minY = Math.max(0, minY - m); maxX = Math.min(w - 1, maxX + m); maxY = Math.min(h - 1, maxY + m);
    var bx = Math.floor(minX * k), by = Math.floor(minY * k);
    return { x: bx, y: by, w: Math.min(src.width - bx, Math.ceil((maxX - minX + 1) * k)), h: Math.min(src.height - by, Math.ceil((maxY - minY + 1) * k)) };
  }
  function placeBox(cr) {
    var box = $('crop-box');
    box.hidden = false;
    box.style.left = (cr.x / ship.src.width * 100) + '%';
    box.style.top = (cr.y / ship.src.height * 100) + '%';
    box.style.width = (cr.w / ship.src.width * 100) + '%';
    box.style.height = (cr.h / ship.src.height * 100) + '%';
  }
  function drawCrop() {
    if (!ship.src) return;
    var cv = $('crop-canvas'), dpr = window.devicePixelRatio || 1;
    var dispW = $('crop-wrap').clientWidth || 280;
    cv.width = Math.round(dispW * dpr);
    cv.height = Math.round(dispW * ship.src.height / ship.src.width * dpr);
    var x = cv.getContext('2d');
    x.fillStyle = '#fff'; x.fillRect(0, 0, cv.width, cv.height);
    x.drawImage(ship.src, 0, 0, cv.width, cv.height);
    placeBox(ship.crop);
    $('page-nav').hidden = ship.pages < 2;
    $('pg-label').textContent = 'Page ' + ship.page + ' of ' + ship.pages;
    $('pg-prev').disabled = ship.page <= 1;
    $('pg-next').disabled = ship.page >= ship.pages;
    drawFit();
  }
  // The chosen area, turned to fit the label best, at up to about 2400 px.
  function output() {
    var cr = ship.crop, box = contentBox(st.fmt);
    var straight = Math.min(box.w / cr.w, box.h / cr.h), turned = Math.min(box.w / cr.h, box.h / cr.w);
    var rot = ((turned > straight * 1.05 ? 90 : 0) + ship.turn) % 360;
    var k = Math.min(1, 2400 / Math.max(cr.w, cr.h));
    var w = Math.round(cr.w * k), h = Math.round(cr.h * k), swap = rot % 180 !== 0;
    var c = document.createElement('canvas');
    c.width = swap ? h : w; c.height = swap ? w : h;
    var x = c.getContext('2d');
    x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
    x.translate(c.width / 2, c.height / 2);
    x.rotate(rot * Math.PI / 180);
    x.drawImage(ship.src, cr.x, cr.y, cr.w, cr.h, -w / 2, -h / 2, w, h);
    var fit = Math.min(box.w / c.width, box.h / c.height);
    return { canvas: c, wIn: c.width * fit, hIn: c.height * fit, box: box };
  }
  function drawFit() {
    if (!ship.src || !st.fmt) return;
    var out = output(), f = st.fmt, box = out.box;
    var cv = $('fit-canvas'), pxPerIn = 40;
    cv.width = Math.round(f.label.width * pxPerIn); cv.height = Math.round(box.rowH * pxPerIn);
    var x = cv.getContext('2d');
    x.fillStyle = '#fff'; x.fillRect(0, 0, cv.width, cv.height);
    var left = (f.padding.left + (box.w - out.wIn) / 2) * pxPerIn;
    var top = (f.align === 'top' ? box.top : (box.rowH - out.hIn) / 2) * pxPerIn;
    x.drawImage(out.canvas, left, top, out.wIn * pxPerIn, out.hIn * pxPerIn);
  }

  // ---------- addresses ----------
  function book() { var b = storeJson(LS.book); return Array.isArray(b) ? b : []; }
  function addToBook(list) {
    var b = book(), seen = {}, added = 0;
    b.forEach(function (e) { seen[e.lines.join('\n').toLowerCase()] = true; });
    list.forEach(function (lines) {
      lines = lines.map(function (s) { return String(s).trim(); }).filter(Boolean);
      var key = lines.join('\n').toLowerCase();
      if (!lines.length || seen[key]) return;
      seen[key] = true; added++;
      b.push({ id: uid(), name: lines[0], lines: lines });
    });
    b.sort(function (x, y) { return x.name.localeCompare(y.name); });
    storeJson(LS.book, b);
    return added;
  }
  function textLines(t) { return String(t || '').split(/\r?\n/).map(function (s) { return s.trim(); }).filter(Boolean); }
  function fromAddress() { return store(LS.from) || getSetting(LS.from) || ''; }
  function saveFrom(v) { store(LS.from, v); setSetting(LS.from, v); }

  var DESIGNS = [
    { id: 'plain', name: 'Plain' },
    { id: 'bold', name: 'Bold name' },
    { id: 'returnto', name: 'Return to' },
    { id: 'logo', name: 'With logo' }
  ];
  function currentDesign() {
    var d = store(LS.design) || 'plain';
    if (d === 'logo' && !logoFits(st.fmt)) d = 'plain';
    return DESIGNS.some(function (x) { return x.id === d; }) ? d : 'plain';
  }
  function logoFits(f) { return !!f && rowHeightIn(f) >= 0.9; }
  // Lines in a design -> paragraphs (and a picture for the logo design).
  function designParas(lines, design, f) {
    var base = f.font.size;
    if (design === 'bold') return { paras: lines.map(function (t, i) { return { t: t, size: i ? base : base + 1, b: !i }; }) };
    if (design === 'returnto') return { paras: [{ t: 'RETURN TO', size: Math.max(5, Math.round(base * 0.6)), b: true, color: '#444444' }].concat(lines.map(function (t) { return { t: t, size: base }; })) };
    if (design === 'logo') {
      var logo = storeJson(LS.logo), paras = lines.map(function (t) { return { t: t, size: Math.max(6, base - 1) }; });
      if (!logo || !logo.b64) return { paras: paras };
      var maxH = Math.min(rowHeightIn(f) * 0.38, 0.6) * 72, maxW = (f.label.width - f.padding.left - f.padding.right) * 72;
      var s = Math.min(maxH / logo.h, maxW / logo.w);
      return { paras: paras, pic: { b64: logo.b64, w: logo.w * s, h: logo.h * s } };
    }
    return { paras: lines.map(function (t) { return { t: t, size: base }; }) };
  }
  // Shipping layout: FROM block at the top, larger SHIP TO block below.
  function shippingParas(fromText, toText, f) {
    var topPt = f.align === 'top' ? f.topSpace * 72 : 0, fromLines = textLines(fromText), paras = [];
    if (fromLines.length) {
      paras.push({ t: 'FROM', size: 8, b: true, before: topPt });
      fromLines.forEach(function (l) { paras.push({ t: l, size: 11 }); });
    }
    paras.push({ t: 'SHIP TO', size: 9, b: true, before: fromLines.length ? 90 : topPt + 90, indent: 36 });
    textLines(toText).forEach(function (l, i) { paras.push({ t: l, size: 16, b: i === 0, indent: 36 }); });
    return paras;
  }

  // ---------- fit text ----------
  function fitLabels(nums) {
    var f = st.fmt, box = textBox(f), tooLong = [], fitted = [];
    return Word.run(function (ctx) {
      var lists = nums.map(function (n) { var p = cellBody(ctx, n).paragraphs; p.load(PARA_PROPS); return p; });
      return ctx.sync().then(function () {
        nums.forEach(function (n, i) {
          var specs = paraSpecs(lists[i].items, f);
          var s = T.fitScale(specs, box, measure);
          if (s === null) { tooLong.push(n); return; }
          if (s >= 0.999) return;
          lists[i].items.forEach(function (p, j) { p.font.size = Math.max(5, Math.floor(specs[j].size * s * 2) / 2); });
          fitted.push(n);
        });
        return ctx.sync();
      });
    }).then(function () { return { fitted: fitted, tooLong: tooLong }; });
  }

  // ---------- table window ----------
  var tableDlg = null;
  function openTableWindow() {
    if (!Office.context.ui || !Office.context.ui.displayDialogAsync) { say('This version of Word can’t open the table window. Update Word and try again.', 'warn'); return; }
    if (tableDlg) { say('The table window is already open.'); return; }
    storeJson(LS.tableInfo, { open: openList().length, name: st.fmt.name, per: total() });
    var url = new URL('table.html', location.href.split('?')[0]).href + '?t=' + Date.now();
    Office.context.ui.displayDialogAsync(url, { height: 70, width: 60, displayInIframe: false }, function (res) {
      if (res.status !== Office.AsyncResultStatus.Succeeded) { say('The table window couldn’t open' + (res.error && res.error.message ? ': ' + res.error.message : '') + '.', 'warn'); return; }
      tableDlg = res.value;
      tableDlg.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        var msg = null;
        try { msg = JSON.parse(arg.message); } catch (e) { return; }
        if (msg && msg.type === 'fill') { closeTable(); fillFromRows(msg.rows || []); }
        else if (msg && msg.type === 'close') closeTable();
      });
      tableDlg.addEventHandler(Office.EventType.DialogEventReceived, function () { tableDlg = null; });
    });
  }
  function closeTable() { try { if (tableDlg) tableDlg.close(); } catch (e) { /* already closed */ } tableDlg = null; }

  // rows: [[name, company, street, apt, city, state, zip], ...]
  function fillFromRows(rows) {
    var f = st.fmt;
    var labels = rows.map(function (r) {
      return T.addressLines({ name: r[0], company: r[1], street: r[2], apt: r[3], city: r[4], state: r[5], zip: r[6] });
    }).filter(function (l) { return l.length; });
    if (!labels.length) { say('The table is empty. Add at least one address.', 'warn'); return; }
    var design = currentDesign() === 'bold' ? 'bold' : 'plain';
    var parasFor = function (lines) { return designParas(lines, design, f).paras; };
    var open = openList();
    var here = labels.slice(0, open.length), rest = labels.slice(open.length);
    var fillHere = function () {
      if (!here.length) return Promise.resolve();
      return queueChange(function () {
        return writeEach(here.map(function (lines, i) { return { n: open[i], paras: parasFor(lines) }; }));
      }).then(refresh);
    };
    if (!rest.length) {
      fillHere().then(function () { say('Filled ' + labelList(open.slice(0, here.length)) + ' from the table.'); });
      return;
    }
    var sheetsNeeded = Math.ceil(rest.length / total());
    ask('You have ' + plural(labels.length, 'address', 'es') + ' but ' + plural(open.length, 'open label') + '. Fill the open labels here and put the other ' + rest.length + ' on ' + plural(sheetsNeeded, 'new sheet') + '?',
      'Fill and make ' + plural(sheetsNeeded, 'sheet'), function () {
        fillHere().then(function () {
          return queue(function () { return createSheets(f, spillPrefills(f, rest.map(parasFor))); });
        }).then(function () {
          say('Filled ' + (here.length ? labelList(open.slice(0, here.length)) + ' here and ' : '') + 'made ' + plural(sheetsNeeded, 'new sheet') + ' for the other ' + rest.length + '.');
        });
      });
  }

  // ---------- rendering ----------
  function show(view) {
    ['v-loading', 'v-outside', 'v-old', 'v-none', 'v-broken', 'v-sheet'].forEach(function (id) { $(id).hidden = id !== view; });
  }
  function render() {
    if (st.kind === 'none' || choosing) { renderChooser(); show('v-none'); return; }
    if (st.kind === 'broken') {
      $('broken-msg').textContent = st.reason;
      $('repair-btn').hidden = !st.fixable;
      $('repair-note').hidden = !st.fixable;
      show('v-broken');
      return;
    }
    if (st.kind !== 'sheet') return;
    var f = st.fmt;
    show('v-sheet');
    $('fmt-name').textContent = f.name;
    $('fmt-detail').textContent = f.detail;
    renderMap();
    renderTally();
    $('first-row').hidden = total() <= 2;
    $('first-n').max = String(total());
    renderFill();
    renderPrint();
    $('save-hint').hidden = !(needsFirstSave && !docUrl());
    var entries = sheetEntries(false);
    $('sheets-count').textContent = entries.length;
    renderSheetList($('sheets-list'), entries, true);
    renderHistory();
    renderSettings();
    renderAbout();
  }

  function thumb(f, width) {
    var t = el('span', 'thumb');
    width = width || 30;
    t.style.width = width + 'px';
    t.style.height = Math.round(width * f.page.height / f.page.width) + 'px';
    t.setAttribute('aria-hidden', 'true');
    for (var n = 0; n < f.across * f.down; n++) {
      var r = Math.floor(n / f.across), c = n % f.across, i = document.createElement('i');
      i.style.left = ((f.firstLabel.left + c * f.pitch.across) / f.page.width * 100) + '%';
      i.style.top = ((f.firstLabel.top + r * f.pitch.down) / f.page.height * 100) + '%';
      i.style.width = (f.label.width / f.page.width * 100) + '%';
      i.style.height = (f.label.height / f.page.height * 100) + '%';
      t.appendChild(i);
    }
    return t;
  }

  function renderChooser() {
    var list = $('fmt-list'), fmts = allFormats();
    if (!fmts.some(function (f) { return f.id === chosenId; })) chosenId = fmts.length ? fmts[0].id : '';
    list.innerHTML = '';
    fmts.forEach(function (f) {
      var lab = el('label', 'fmt');
      var inp = document.createElement('input'); inp.type = 'radio'; inp.name = 'fmt'; inp.value = f.id; inp.checked = f.id === chosenId;
      inp.addEventListener('change', function () { chosenId = f.id; });
      var text = el('span');
      text.appendChild(el('b', '', f.name));
      text.appendChild(el('span', 'd', f.detail));
      lab.appendChild(inp); lab.appendChild(thumb(f)); lab.appendChild(text);
      if (f.custom) {
        var x = el('button', 'x', '×'); x.type = 'button'; x.title = 'Remove this size'; x.setAttribute('aria-label', 'Remove ' + f.name);
        x.addEventListener('click', function (e) {
          e.preventDefault();
          ask('Remove the size “' + f.name + '” from this Mac? Sheets you already made with it keep working.', 'Remove size', function () {
            storeJson(LS.custom, (storeJson(LS.custom) || []).filter(function (c) { return c.id !== f.id; }));
            render();
          }, true);
        });
        lab.appendChild(x);
      } else lab.appendChild(el('span'));
      list.appendChild(lab);
    });
    var links = $('tpl-links'); links.innerHTML = '';
    F.formats.forEach(function (f, i) {
      if (i) links.appendChild(document.createTextNode(' · '));
      var a = el('a', '', f.name); a.href = 'templates/' + f.id + '.docx'; a.target = '_blank'; a.rel = 'noopener';
      links.appendChild(a);
    });
    $('fmt-problems').hidden = !F.invalid.length;
    $('fmt-problems').textContent = F.invalid.map(function (x) { return 'The size “' + x.id + '” in formats.js was skipped: ' + x.problems.join('; ') + '.'; }).join(' ');
    $('choose-title').textContent = choosing ? 'Make another sheet' : 'Make a label sheet';
    $('choose-lead').textContent = choosing
      ? 'Choose the labels for the new sheet. It opens as a new Word document.'
      : 'This document isn’t a label sheet yet. Choose the labels you have, and the sheet opens as a new Word document laid out to match them.';
    $('back-btn').hidden = !(choosing && st.kind === 'sheet');
    var entries = sheetEntries(true);
    $('finish-first').hidden = !entries.length;
    renderSheetList($('finish-list'), entries, false);
    renderSizeForm();
  }

  var mapBuiltFor = '';
  function renderMap() {
    var f = st.fmt, map = $('map'), sp = spots();
    if (mapBuiltFor !== f.id + total()) {
      map.innerHTML = '';
      map.style.aspectRatio = f.page.width + ' / ' + f.page.height;
      for (var n = 1; n <= total(); n++) {
        var r = Math.floor((n - 1) / f.across), c = (n - 1) % f.across;
        var b = el('button', 'cell'); b.type = 'button'; b.dataset.n = n;
        b.style.left = ((f.firstLabel.left + c * f.pitch.across) / f.page.width * 100) + '%';
        b.style.top = ((f.firstLabel.top + r * f.pitch.down) / f.page.height * 100) + '%';
        b.style.width = (f.label.width / f.page.width * 100) + '%';
        b.style.height = (f.label.height / f.page.height * 100) + '%';
        b.style.borderRadius = f.corner ? (f.corner / f.page.width * 100) + 'cqw' : '0';
        b.innerHTML = '<span class="t"></span><span class="num" aria-hidden="true">' + n + '</span><span class="side" aria-hidden="true"></span>';
        map.appendChild(b);
      }
      mapBuiltFor = f.id + total();
    }
    Array.prototype.forEach.call(map.children, function (b) {
      var n = +b.dataset.n;
      var state = st.used[n] ? 'used' : st.ready[n] ? 'ready' : 'open';
      b.className = 'cell ' + state + (st.over[n] ? ' over' : '') + (st.here === n ? ' here' : '');
      b.querySelector('.t').textContent = state === 'ready' ? String(st.texts[n] || '').split(/[\r\n\u000b]+/).slice(0, 3).join('\n') : '';
      b.querySelector('.side').textContent = sideWord(n);
      var tip = '';
      if (state === 'used') {
        var rec = sp[n];
        tip = rec ? (rec.m ? 'Marked used ' + fmtDate(rec.at) : 'Printed ' + fmtDate(rec.at) + (rec.t ? ' · “' + rec.t + '”' : '')) : 'Used';
      } else if (state === 'ready') tip = st.over[n] ? 'Ready to print, but the text may be cut off' : 'Ready to print';
      else tip = 'Open';
      b.title = tip;
      b.setAttribute('aria-label', cap(labelName(n)) + ', ' + (state === 'ready' ? 'ready to print: ' + firstLine(st.texts[n]) + (st.over[n] ? ', text may be cut off' : '') : tip.toLowerCase()));
    });
  }

  function renderTally() {
    var t = $('tally'); t.innerHTML = '';
    var counts = { open: openList().length, used: keys(st.used).length, ready: keys(st.ready).length };
    [['open', 'open'], ['used', 'used'], ['ready', 'ready to print']].forEach(function (p) {
      var s = el('span');
      s.innerHTML = '<i class="sw ' + p[0] + '"></i><b>' + counts[p[0]] + '</b>';
      s.appendChild(document.createTextNode(p[1]));
      t.appendChild(s);
    });
    var lim = tripLimit(), tr = trips();
    var tripsEl = el('span', 'trips' + (tr >= lim - 1 && tr > 0 ? ' last' : ''), tr + ' of ' + lim + ' trips');
    tripsEl.title = 'Times this sheet has gone through the printer';
    t.appendChild(tripsEl);
  }

  function renderHere() {
    var box = $('here'), n = st.here;
    box.innerHTML = '';
    if (!n) { box.textContent = 'The cursor isn’t in a label.'; return; }
    box.appendChild(document.createTextNode('The cursor is in '));
    box.appendChild(el('b', '', labelName(n)));
    box.appendChild(document.createTextNode(st.used[n] ? ', which is used.' : st.ready[n] ? (st.over[n] ? '. Its text may be cut off.' : '.') : ', which is empty.'));
  }

  function fillTabs() {
    return st.fmt.shipping ? [['file', 'Label file'], ['addr', 'Address']] : [['copy', 'Copy'], ['addr', 'Address'], ['table', 'Table']];
  }
  function currentTab() {
    var tabs = fillTabs(), saved = store(st.fmt.shipping ? LS.shipTab : LS.fillTab);
    return tabs.some(function (t) { return t[0] === saved; }) ? saved : tabs[0][0];
  }
  function setFillTab(id) { store(st.fmt.shipping ? LS.shipTab : LS.fillTab, id); if (st.kind === 'sheet') renderFill(); }

  function renderFill() {
    var f = st.fmt, tab = currentTab(), n = nextOpen();
    $('s2-title').textContent = f.shipping ? 'Add a shipping label' : 'Fill labels';
    var tabsEl = $('fill-tabs');
    tabsEl.innerHTML = '';
    fillTabs().forEach(function (t) {
      var b = el('button', '', t[1]); b.type = 'button'; b.setAttribute('role', 'tab');
      b.setAttribute('aria-selected', String(t[0] === tab));
      b.setAttribute('aria-controls', 'pane-' + t[0]);
      b.addEventListener('click', function () { setFillTab(t[0]); });
      tabsEl.appendChild(b);
    });
    ['copy', 'file', 'addr', 'table'].forEach(function (id) { $('pane-' + id).hidden = id !== tab; });

    // Shipping sheets: which label is next, and which half is which.
    var applies = sideApplies(f), side = getSide(f);
    $('ship-target').hidden = !f.shipping;
    $('ship-target').textContent = n ? 'Next open: ' + theName(n) + '.' : 'There are no open labels on this sheet. When you load a fresh sheet, click Reset for a fresh sheet.';
    $('side-row').hidden = !applies;
    $('side-q').textContent = side
      ? 'Label 1 prints on the ' + side + ' half of your sheet. Change this if it’s wrong:'
      : 'Print the test page once. Which half of your sheet did Label 1 come out on?';
    $('side-top').setAttribute('aria-pressed', String(side === 'top'));
    $('side-bottom').setAttribute('aria-pressed', String(side === 'bottom'));

    if (tab === 'copy') renderHere();
    if (tab === 'file') { $('ship-place').disabled = !n; if (ship.src) drawFit(); }
    if (tab === 'addr') renderAddress();
    if (tab === 'table') {
      $('table-open').disabled = !n;
      $('table-msg').textContent = n ? plural(openList().length, 'open label') + ' on this sheet. Extra addresses go on new sheets.' : 'There are no open labels on this sheet.';
    }

    // Text that may not fit
    var over = keys(st.over), warn = $('fit-warn');
    warn.hidden = !over.length || f.shipping;
    warn.innerHTML = '';
    if (over.length) {
      warn.appendChild(document.createTextNode('Text may be cut off on ' + labelList(over) + '. '));
      var fit = el('button', 'btn link inline', over.length === 1 ? 'Fit it' : 'Fit them'); fit.type = 'button';
      fit.addEventListener('click', function () { doFit(over); });
      warn.appendChild(fit);
    }
    // Reprint last label
    var last = storeJson(LS.last + f.id), rp = $('reprint-btn');
    rp.hidden = !(last && last.ooxml && n);
    if (!rp.hidden) rp.textContent = 'Reprint last label (“' + last.text + '”)';
  }

  function pickerOptions(select, includeMine) {
    var b = book(), mine = fromAddress();
    select.innerHTML = '';
    var o0 = el('option', '', b.length || (includeMine && mine) ? 'Choose a saved address…' : 'No saved addresses yet'); o0.value = '';
    select.appendChild(o0);
    if (includeMine && mine) { var om = el('option', '', 'My return address'); om.value = 'mine'; select.appendChild(om); }
    b.forEach(function (e) { var o = el('option', '', e.name + (e.lines[1] ? ' · ' + e.lines[1] : '')); o.value = e.id; select.appendChild(o); });
  }
  function renderAddress() {
    var f = st.fmt, n = nextOpen();
    $('addr-plain').hidden = !!f.shipping;
    $('addr-ship').hidden = !f.shipping;
    if (f.shipping) {
      pickerOptions($('to-pick'), false);
      if (document.activeElement !== $('from-addr')) $('from-addr').value = fromAddress();
      $('ship-addr-place').disabled = !n;
    } else {
      pickerOptions($('addr-pick'), true);
      var box = $('designs'), d = currentDesign();
      box.innerHTML = '';
      DESIGNS.forEach(function (x) {
        if (x.id === 'logo' && !logoFits(f)) return;
        var b = el('button', 'design'); b.type = 'button'; b.setAttribute('role', 'radio'); b.setAttribute('aria-checked', String(x.id === d));
        var mini = el('span', 'mini');
        if (x.id === 'returnto') mini.appendChild(el('small', '', 'RETURN TO'));
        if (x.id === 'logo') mini.appendChild(el('i'));
        var nm = x.id === 'bold' ? el('b', '', 'Your Name') : document.createTextNode('Your Name');
        mini.appendChild(nm); mini.appendChild(el('br')); mini.appendChild(document.createTextNode('123 Main St'));
        b.appendChild(mini); b.appendChild(document.createTextNode(x.name));
        b.addEventListener('click', function () { store(LS.design, x.id); renderAddress(); });
        box.appendChild(b);
      });
      var logo = storeJson(LS.logo);
      $('logo-row').hidden = d !== 'logo';
      $('logo-label').textContent = logo ? 'Change logo' : 'Choose a logo';
      $('logo-msg').textContent = logo ? 'Logo saved on this Mac.' : 'PNG or JPEG. Saved on this Mac.';
      $('addr-one').disabled = !n; $('addr-all').disabled = !n;
    }
    var b2 = book();
    $('book-count').textContent = b2.length;
    var ul = $('book-list'); ul.innerHTML = '';
    if (!b2.length) ul.appendChild(el('li', '', 'No saved addresses yet.'));
    b2.forEach(function (e) {
      var li = el('li'), span = el('span', '', e.lines.join(', '));
      span.title = e.lines.join('\n');
      var x = el('button', '', '×'); x.type = 'button'; x.setAttribute('aria-label', 'Delete ' + e.name);
      x.addEventListener('click', function () { storeJson(LS.book, book().filter(function (k) { return k.id !== e.id; })); renderAddress(); say('Deleted ' + e.name + ' from the address book.'); });
      li.appendChild(span); li.appendChild(x); ul.appendChild(li);
    });
  }

  function renderPrint() {
    var f = st.fmt, ready = keys(st.ready).length, open = openList().length, tr = trips(), lim = tripLimit();
    // Feeding guide
    var side = getSide(f), text;
    if (sideApplies(f)) {
      text = side ? 'Feed the sheet print side up. Label 1 prints on the ' + side + ' half. Choose Labels as the paper type.'
        : 'Feed the sheet print side up and choose Labels as the paper type. Print a test page to see which half is label 1.';
    } else {
      text = 'Feed the sheet top edge first, print side up, and choose Labels as the paper type.' +
        (arrowFits() && feedArrowOn() ? ' The small ▲ in the top strip marks the top.' : '');
    }
    $('feed-text').textContent = text;
    // Avery's one-pass note: shown before a sheet's second trip.
    $('warn-pass').hidden = !(tr >= 1 && ready > 0 && store(LS.warnPass) !== '1');
    // After a test page
    $('test-check').hidden = store(LS.testCheck + f.id) !== 'pending';
    // Last trip: fill the rest so nothing goes to waste
    var lastTrip = !f.shipping && tr === lim - 1 && open > 0 && ready > 0;
    $('use-up').hidden = !lastTrip;
    if (lastTrip) $('use-up-msg').textContent = 'This is this sheet’s last trip through the printer. Fill the ' + plural(open, 'open label') + ' too, so nothing goes to waste?';
    $('use-up-mine').hidden = !fromAddress();
    // Trip limit reached
    $('trips-out').hidden = tr < lim;
    $('trips-out').textContent = 'This sheet has gone through the printer ' + plural(tr, 'time') + ', your limit. Use the rest of its labels by hand, or print anyway at your own risk.';
    $('printed-btn').textContent = ready ? 'I printed these (' + ready + ')' : 'I printed these';
    $('printed-btn').disabled = !ready;
    $('undo-row').hidden = !undo;
    if (undo) $('undo-msg').textContent = cap(labelList(undo.nums)) + (undo.nums.length === 1 ? ' was' : ' were') + ' marked used.';
  }

  function renderHistory() {
    var ul = $('hist'); ul.innerHTML = '';
    var h = history();
    if (!h.length) { ul.appendChild(el('li', 'txt', 'Nothing printed yet.')); return; }
    h.slice(0, 20).forEach(function (e) {
      var li = el('li');
      li.appendChild(el('span', 'when', fmtDate(e.at)));
      li.appendChild(el('span', '', cap(labelList(e.nums || []))));
      if (e.text) li.appendChild(el('span', 'txt', '“' + e.text + '”'));
      ul.appendChild(li);
    });
  }

  function renderSettings() {
    var f = st.fmt, nd = nudgeFor(f.id);
    $('auto-mark').checked = autoMarkOn();
    $('arrow-row').hidden = !(F.hasStrip(f));
    $('feed-arrow').checked = feedArrowOn();
    $('feed-arrow').disabled = !arrowFits() || !!st.headerOther;
    if (document.activeElement !== $('trip-limit')) $('trip-limit').value = tripLimit();
    if (document.activeElement !== $('nudge-x')) $('nudge-x').value = nd.x;
    if (document.activeElement !== $('nudge-y')) $('nudge-y').value = nd.y;
    var ap = st.applied || { x: 0, y: 0 };
    $('nudge-apply').hidden = Math.abs(ap.x - nd.x) < 0.001 && Math.abs(ap.y - nd.y) < 0.001;
  }

  function renderAbout() {
    var tagged = getSetting(DOC.autoShow) === true, saved = !!docUrl(), ref = cachedRef(), nd = nudgeFor(st.fmt.id), ap = st.applied || { x: 0, y: 0 };
    var rows = [
      ['Version', VERSION],
      ['This sheet', tagged ? (saved ? 'Opens with the panel' : 'Opens with the panel once you save it') : 'Not set to open with the panel yet'],
      ['New sheets', ref ? 'Open with the panel' : refState === 'missing' ? 'Need one more step: save this document, close it, open it again, then make the sheet'
        : refState.indexOf('failed') === 0 ? 'Can’t open with the panel (' + refState.slice(8) + ')' : 'Checking…'],
      ['Trips', trips() + ' of ' + tripLimit()],
      ['Adjustment', nd.x || nd.y ? 'right ' + nd.x + ' in, down ' + nd.y + ' in' + (ap.x === nd.x && ap.y === nd.y ? '' : ' (not yet on this sheet)') : 'None'],
      ['Add-in link', ref ? (ref.storeType || '?') + ' · ' + (ref.store || '?') : '—']
    ];
    var dl = $('about-list'); dl.innerHTML = '';
    rows.forEach(function (r) { dl.appendChild(el('dt', '', r[0])); dl.appendChild(el('dd', '', r[1])); });
  }

  // ---------- add your own size ----------
  function sizeFromForm() {
    var num = function (id) { return Number($(id).value); };
    var w = num('sz-w'), h = num('sz-h'), sideways = $('sz-sideways').checked, shipping = $('sz-ship').checked;
    var detail = (w + '″ × ' + h + '″') + ($('sz-detail').value.trim() ? ' · ' + $('sz-detail').value.trim() : '');
    var font = h < 0.6 ? 7 : h < 1.2 ? 10 : h < 2.5 ? 12 : 14;
    return {
      id: 'custom-' + uid(), name: $('sz-name').value.trim() || 'My labels', detail: detail,
      page: sideways ? { width: 11, height: 8.5 } : { width: 8.5, height: 11 },
      across: Math.round(num('sz-across')), down: Math.round(num('sz-down')),
      label: { width: w, height: h },
      firstLabel: { left: num('sz-left'), top: num('sz-top') },
      pitch: { across: w + num('sz-gapx'), down: h + num('sz-gapy') },
      padding: { left: Math.min(0.1, w * 0.05), right: Math.min(0.1, w * 0.05) },
      align: shipping ? 'top' : 'center', topSpace: 0.45,
      font: { name: 'Arial', size: font }, corner: w < 5 ? 0.0625 : 0, shipping: shipping
    };
  }
  function sizeProblems(f) {
    var bad = [];
    if (!(f.across >= 1 && f.down >= 1)) bad.push('Enter how many labels go across and down.');
    if (!(f.label.width > 0 && f.label.height > 0)) bad.push('Enter the label width and height.');
    [f.firstLabel.left, f.firstLabel.top, f.pitch.across - f.label.width, f.pitch.down - f.label.height].forEach(function (v) { if (!(v >= 0)) bad.push('Distances can’t be negative.'); });
    if (bad.length) return bad.filter(function (v, i, a) { return a.indexOf(v) === i; });
    return F.problems(f).map(function (p) {
      if (/right edge/.test(p)) return 'The labels run past the right edge of the page. Check the widths and gaps.';
      if (/bottom edge/.test(p)) return 'The labels run past the bottom of the page. Check the heights and gaps.';
      return cap(p) + '.';
    });
  }
  function renderSizeForm() {
    var f = sizeFromForm(), bad = sizeProblems(f), holder = $('sz-thumb');
    holder.innerHTML = '';
    if (!bad.length) holder.appendChild(thumb(F.withDefaults(JSON.parse(JSON.stringify(f))), 84));
    $('sz-msg').textContent = bad.length ? bad.join(' ') : (f.across * f.down) + ' labels, ' + f.label.width + '″ × ' + f.label.height + '″. Print a test page to check it before saving.';
    $('sz-msg').className = bad.length ? 'warn' : 'hint';
    $('sz-save').disabled = !!bad.length;
    $('sz-test').disabled = !!bad.length;
  }

  // ---------- wiring ----------
  function wire() {
    // Choosing and making sheets
    $('create-btn').addEventListener('click', function () {
      var f = findFormat(chosenId);
      if (!f) { say('Choose a label size first.', 'warn'); return; }
      var make = function () {
        queue(function () { return createSheets(f, null); }).then(function (withPanel) {
          var msg = 'Your new sheet (' + f.name + ') opened in its own window' +
            (withPanel ? ', with this panel.' : '. Click Label Sheet on its Home tab to open the panel there.');
          if (st.kind === 'sheet') { choosing = false; render(); say(msg); }
          else { $('made').hidden = false; $('made-msg').textContent = msg; say(''); }
        });
      };
      var lim = tripLimit();
      var left = sheetEntries(true).filter(function (e) { return e.fmtId === f.id && e.open > 0 && (e.trips || 0) < lim; })[0];
      if (left) ask('You still have ' + plural(left.open, 'open label') + ' on “' + (left.file || 'a sheet you haven’t saved') + '” (' + f.name + '). Finish that sheet first?', 'Make a new sheet anyway', make);
      else make();
    });
    $('broken-new').addEventListener('click', function () { choosing = true; render(); });
    $('back-btn').addEventListener('click', function () { choosing = false; render(); });
    $('other-btn').addEventListener('click', function () { choosing = true; $('made').hidden = true; render(); });
    $('refresh-btn').addEventListener('click', function () { refresh().then(function () { say('Up to date.'); }); });
    $('reset-btn').addEventListener('click', function () {
      var used = keys(st.used);
      if (!used.length && !trips()) { say('Every label on this sheet is already open.'); return; }
      ask('Loaded a fresh sheet? All ' + plural(used.length, 'used label') + ' become open again and the trip count starts over. Labels ready to print stay as they are.', 'Reset', function () {
        undo = null;
        queueChange(function () { return markOpen(used).then(function () { setSetting(DOC.trips, 0); setSetting(DOC.spots, {}); }); }).then(refresh).then(function () { say('Reset. Every label without text is open.'); });
      });
    });
    var repair = function (nudge) {
      var f = st.fmt, margins = st.margins || { top: f.firstLabel.top, left: f.firstLabel.left };
      queueChange(function () { return rebuildSheet(f, margins, nudge); }).then(function (r) {
        return refresh().then(function () {
          say('Rebuilt the labels. Kept ' + plural(r.used, 'used label') + ' and ' + plural(r.ready, 'label') + ' ready to print.' + (r.lost ? ' ' + plural(r.lost, 'label') + ' with text didn’t fit back on the sheet.' : ''), r.lost ? 'warn' : '');
        });
      });
    };
    $('repair-btn').addEventListener('click', function () { repair(st.applied || { x: 0, y: 0 }); });
    $('rebuild-btn').addEventListener('click', function () {
      ask('Repair rebuilds this sheet’s label layout. Used labels and labels ready to print are kept. Continue?', 'Repair', function () { repair(st.applied || { x: 0, y: 0 }); });
    });
    $('test-btn').addEventListener('click', function () {
      if (!st.fmt) return;
      var f = st.fmt;
      queue(function () { return makeTestPage(f); }).then(function () {
        render();
        say('A test page opened in its own window. Print it on scrap paper at 100% scale.');
      });
    });

    // Step 1: mark used labels
    $('map').addEventListener('click', function (ev) {
      var b = ev.target.closest('.cell'); if (!b || busy) return;
      var n = +b.dataset.n;
      if (st.used[n]) {
        queueChange(function () { return markOpen([n]); }).then(refresh).then(function () { say(cap(labelName(n)) + ' is open again.'); });
      } else if (st.ready[n]) {
        ask(cap(labelName(n)) + ' has text on it. Clear it and mark it used?', 'Clear and mark used', function () {
          queueChange(function () { return markUsed([n]); }).then(refresh).then(function () { say(cap(labelName(n)) + ' marked used.'); });
        }, true);
      } else {
        queueChange(function () { return markUsed([n]); }).then(refresh).then(function () { say(cap(labelName(n)) + ' marked used.'); });
      }
    });
    $('first-btn').addEventListener('click', function () {
      var max = total(), k = Math.round(Number($('first-n').value));
      if (!(k >= 1 && k <= max)) { say('Enter a number from 1 to ' + max + '.', 'warn'); return; }
      var nums = []; for (var n = 1; n <= k; n++) if (!st.used[n]) nums.push(n);
      if (!nums.length) { say((k === 1 ? 'Label 1 is' : 'Labels 1–' + k + ' are') + ' already marked used.'); return; }
      var run = function () { queueChange(function () { return markUsed(nums); }).then(refresh).then(function () { say('Marked ' + labelList(nums) + ' used.'); }); };
      var withText = nums.filter(function (x) { return st.ready[x]; });
      if (withText.length) ask(cap(labelList(withText)) + (withText.length === 1 ? ' has' : ' have') + ' text on it. Clear and mark used?', 'Clear and mark used', run, true);
      else run();
    });

    // Step 2: copy
    function doCopy(k, all) {
      queue(function () { return readState().then(apply).then(cursorLabel); }).then(function (src) {
        var open = openList(), targets = open.slice(0, k), short = all ? 0 : Math.max(0, k - open.length);
        if (!targets.length && !short) { say('There are no open labels left on this sheet.', 'warn'); return; }
        var copyHere = function () {
          if (!targets.length) return Promise.resolve();
          return queueChange(function () { return copyFrom(src, targets); }).then(refresh);
        };
        if (!short) { copyHere().then(function () { say('Copied ' + labelName(src) + ' to ' + labelList(targets) + '.'); }); return; }
        ask(open.length
          ? 'Only ' + plural(open.length, 'open label') + ' left. Copy to those and put the other ' + plural(short, 'copy', 'ies').replace('copyies', 'copies') + ' on a new sheet?'
          : 'There are no open labels left. Put ' + plural(short, 'copy', 'ies').replace('copyies', 'copies') + ' on a new sheet?', 'Copy and make a sheet', function () {
          copyHere().then(function () { return queue(function () { return readLabelParas(src); }); }).then(function (paras) {
            var copies = []; for (var i = 0; i < short; i++) copies.push(paras);
            return queue(function () { return createSheets(st.fmt, spillPrefills(st.fmt, copies)); });
          }).then(function () { say('Copied ' + labelName(src) + (targets.length ? ' to ' + labelList(targets) + ' and' : '') + ' to a new sheet.'); });
        });
      });
    }
    $('copy-btn').addEventListener('click', function () {
      var k = Math.round(Number($('copies').value));
      if (!(k >= 1)) { say('Enter how many copies you want.', 'warn'); return; }
      doCopy(k, false);
    });
    $('fill-all').addEventListener('click', function () { doCopy(total(), true); });
    // TEMPORARY: triple-click "About this panel" details to run Word checks (lab.js).
    $('about-list').addEventListener('click', function (e) {
      if (e.detail < 3) return;
      var s = document.createElement('script'); s.src = 'lab.js?t=' + Date.now(); document.body.appendChild(s);
    });
    $('fit-btn').addEventListener('click', function () {
      queue(function () { return readState().then(apply); }).then(function () {
        var n = st.here && st.ready[st.here] ? st.here : 0;
        var nums = n ? [n] : keys(st.over);
        if (!nums.length) { say('Click inside a label with text first. Labels whose text may be cut off get a red outline on the map.', 'warn'); return; }
        doFit(nums);
      });
    });

    // Step 2: shipping label files, address, table
    wireShip();
    wireAddress();
    $('table-open').addEventListener('click', openTableWindow);
    $('reprint-btn').addEventListener('click', function () {
      queueChange(reprintLast).then(function (n) { return refresh().then(function () { say('Put the last label on ' + theName(n) + ' again.'); }); });
    });

    // Step 3: print
    $('printed-btn').addEventListener('click', function () {
      if (!keys(st.ready).length) { say('No labels are ready to print.', 'warn'); return; }
      queueChange(markPrinted).then(function (nums) {
        return refresh().then(function () { say('Marked ' + labelList(nums) + ' used. ' + plural(openList().length, 'open label') + ' left.'); });
      });
      readPrintTime().then(function (ms) { if (ms) { lastPrintSeen = Math.max(lastPrintSeen || 0, ms); setSetting(DOC.lastPrint, lastPrintSeen); } }).catch(function () {});
    });
    $('undo-btn').addEventListener('click', function () {
      if (!undo) return;
      queueChange(undoPrinted).then(refresh).then(function () { say('Undone. Those labels have their text back.'); });
    });
    $('warn-pass-hide').addEventListener('click', function () { store(LS.warnPass, '1'); render(); });
    $('test-yes').addEventListener('click', function () { store(LS.testCheck + st.fmt.id, null); render(); say('You’re set for this printer and these labels.'); });
    $('test-no').addEventListener('click', function () {
      var box = $('settings'); box.open = true;
      var nd = $('nudge'); nd.classList.remove('flash'); void nd.offsetWidth; nd.classList.add('flash');
      try { nd.scrollIntoView({ block: 'center' }); } catch (e) { /* old browser */ }
      $('nudge-x').focus();
      say('Enter how far the boxes were off, then click Save. Print another test page to check.');
    });
    $('use-up-mine').addEventListener('click', function () {
      var lines = textLines(fromAddress()), open = openList();
      if (!lines.length) { setFillTab('addr'); say('Enter your return address first.', 'warn'); return; }
      var d = designParas(lines, currentDesign(), st.fmt);
      queueChange(function () { return writeLabels(open, d.paras, d.pic); }).then(refresh).then(function () { say('Filled ' + labelList(open) + ' with your return address.'); });
    });
    $('use-up-copy').addEventListener('click', function () {
      var src = st.here && st.ready[st.here] ? st.here : keys(st.ready)[0], open = openList();
      queueChange(function () { return copyFrom(src, open); }).then(refresh).then(function () { say('Filled ' + labelList(open) + ' with copies of ' + labelName(src) + '.'); });
    });

    // Settings
    $('auto-mark').addEventListener('change', function () { store(LS.autoMark, $('auto-mark').checked ? 'on' : 'off'); });
    $('feed-arrow').addEventListener('change', function () { store(LS.feedArrow, $('feed-arrow').checked ? 'on' : 'off'); syncHeader(); render(); });
    $('trip-limit').addEventListener('change', function () {
      var n = Math.round(Number($('trip-limit').value));
      if (!(n >= 1 && n <= 20)) { say('Enter a number of trips from 1 to 20.', 'warn'); return; }
      store(LS.tripLimit, String(n)); updateRegistry(); render();
    });
    $('nudge-save').addEventListener('click', function () {
      var x = Number($('nudge-x').value), y = Number($('nudge-y').value);
      if (!(Math.abs(x) <= 0.5 && Math.abs(y) <= 0.5)) { say('Enter adjustments between -0.5 and 0.5 inches.', 'warn'); return; }
      storeJson(LS.nudge + st.fmt.id, { x: Math.round(x * 1000) / 1000, y: Math.round(y * 1000) / 1000 });
      store(LS.testCheck + st.fmt.id, null);
      render();
      say('Saved. New sheets and test pages for this size use it.' + ($('nudge-apply').hidden ? '' : ' Click Apply to this sheet to move these labels too.'));
    });
    $('nudge-apply').addEventListener('click', function () {
      var nd = nudgeFor(st.fmt.id), L = S.layout(st.fmt, { margins: st.margins, nudge: nd });
      if (L.clampedUp) say('These labels can’t move that far up on this sheet. They were moved as far as possible; new sheets get the full adjustment.', 'warn');
      repair(nd);
    });

    $('confirm-yes').addEventListener('click', function () { var a = pendingConfirm; closeAsk(); if (a) a(); });
    $('confirm-no').addEventListener('click', closeAsk);
    wireFileLabels();

    // Add your own size
    ['sz-name', 'sz-detail', 'sz-across', 'sz-down', 'sz-w', 'sz-h', 'sz-top', 'sz-left', 'sz-gapx', 'sz-gapy', 'sz-sideways', 'sz-ship'].forEach(function (id) {
      $(id).addEventListener('input', renderSizeForm);
      $(id).addEventListener('change', renderSizeForm);
    });
    $('sz-test').addEventListener('click', function () {
      var f = F.withDefaults(sizeFromForm());
      queue(function () { return createDoc(f, { test: true }); }).then(function () { say('A test page for your size opened in its own window. Print it on scrap paper at 100% scale.'); });
    });
    $('sz-save').addEventListener('click', function () {
      var f = sizeFromForm();
      if (sizeProblems(f).length) return;
      var list = storeJson(LS.custom) || []; list.push(f); storeJson(LS.custom, list);
      chosenId = f.id; $('add-size').open = false;
      render();
      say('Saved “' + f.name + '” on this Mac. It’s selected; click Make the sheet.');
    });
  }

  function doFit(nums) {
    queueChange(function () { return fitLabels(nums); }).then(function (r) {
      return refresh().then(function () {
        var parts = [];
        if (r.fitted.length) parts.push('Shrank the text on ' + labelList(r.fitted) + ' to fit.');
        if (r.tooLong.length) parts.push(cap(labelList(r.tooLong)) + (r.tooLong.length === 1 ? ' has' : ' have') + ' too much text to fit even at 5 pt. Shorten it.');
        if (!parts.length) parts.push('The text already fits.');
        say(parts.join(' '), r.tooLong.length ? 'warn' : '');
      });
    });
  }

  function wireShip() {
    [$('side-top'), $('side-bottom')].forEach(function (btn) {
      btn.addEventListener('click', function () {
        if (!st.fmt) return;
        store(LS.side + st.fmt.id, btn.dataset.side);
        render();
        say('Got it. Label 1 is the ' + btn.dataset.side + ' label on your sheet.');
      });
    });
    var input = $('ship-file');
    input.addEventListener('change', function () { handleFile(input.files && input.files[0]); input.value = ''; });
    var card = $('pane-file');
    card.addEventListener('dragover', function (e) { e.preventDefault(); card.classList.add('drop'); });
    card.addEventListener('dragleave', function () { card.classList.remove('drop'); });
    card.addEventListener('drop', function (e) {
      e.preventDefault(); card.classList.remove('drop');
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) handleFile(f);
    });
    // Paste a screenshot: ⌘V anywhere in the panel (outside text boxes).
    document.addEventListener('paste', function (e) {
      if (!st.fmt || !st.fmt.shipping || st.kind !== 'sheet') return;
      var tgt = e.target;
      if (tgt && /^(TEXTAREA|INPUT|SELECT)$/.test(tgt.tagName)) return;
      var items = (e.clipboardData && e.clipboardData.items) || [];
      for (var i = 0; i < items.length; i++) {
        if (items[i].kind === 'file' && /^image\//.test(items[i].type)) {
          var file = items[i].getAsFile();
          if (file) { e.preventDefault(); handleFile(file); return; }
        }
      }
      say('There’s no picture on the clipboard. Copy a label image first, then press ⌘V here.', 'warn');
    });

    $('pg-prev').addEventListener('click', function () { if (ship.doc && ship.page > 1) { ship.page--; renderPdfPage().then(afterSource).catch(fileError); } });
    $('pg-next').addEventListener('click', function () { if (ship.doc && ship.page < ship.pages) { ship.page++; renderPdfPage().then(afterSource).catch(fileError); } });
    $('crop-find').addEventListener('click', function () {
      if (!ship.src) return;
      var found = findLabel(ship.src);
      ship.crop = found || trimBox(ship.src);
      drawCrop();
      say(found ? 'Found the label’s border.' : 'No printed border found, so the white edges were trimmed instead. Drag across the label if needed.', found ? '' : 'warn');
    });
    $('crop-all').addEventListener('click', function () { if (ship.src) { ship.crop = { x: 0, y: 0, w: ship.src.width, h: ship.src.height }; drawCrop(); } });
    $('crop-turn').addEventListener('click', function () { ship.turn = (ship.turn + 90) % 360; drawFit(); });

    // Drag a rectangle on the page preview to choose the label area.
    var wrap = $('crop-wrap'), drag = null;
    function frac(e) {
      var r = wrap.getBoundingClientRect();
      return { x: Math.min(1, Math.max(0, (e.clientX - r.left) / r.width)), y: Math.min(1, Math.max(0, (e.clientY - r.top) / r.height)) };
    }
    function rectFrom(a, b) {
      return { x: Math.min(a.x, b.x) * ship.src.width, y: Math.min(a.y, b.y) * ship.src.height, w: Math.abs(a.x - b.x) * ship.src.width, h: Math.abs(a.y - b.y) * ship.src.height };
    }
    wrap.addEventListener('pointerdown', function (e) {
      if (!ship.src) return;
      drag = { a: frac(e), b: frac(e) };
      try { wrap.setPointerCapture(e.pointerId); } catch (err) { /* not supported */ }
      e.preventDefault();
    });
    wrap.addEventListener('pointermove', function (e) { if (!drag) return; drag.b = frac(e); placeBox(rectFrom(drag.a, drag.b)); });
    wrap.addEventListener('pointerup', function (e) {
      if (!drag) return;
      drag.b = frac(e);
      var r = rectFrom(drag.a, drag.b);
      drag = null;
      if (r.w > ship.src.width * 0.03 && r.h > ship.src.height * 0.03) ship.crop = { x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.w), h: Math.round(r.h) };
      drawCrop();
    });
    wrap.addEventListener('pointercancel', function () { drag = null; drawCrop(); });

    $('ship-place').addEventListener('click', function () {
      if (!ship.src) { say('Choose a PDF or picture first.', 'warn'); return; }
      var n = nextOpen();
      if (!n) { say('There are no open labels on this sheet.', 'warn'); return; }
      var out = output(), b64 = out.canvas.toDataURL('image/png').split(',')[1];
      queueChange(function () { return writeLabels([n], [], { b64: b64, w: out.wIn * 72, h: out.hIn * 72, align: 'center' }); }).then(refresh).then(function () {
        say('Placed on ' + theName(n) + '. Print with ⌘P.');
      });
    });
  }

  function wireAddress() {
    // Address labels (with a design)
    $('addr-pick').addEventListener('change', function () {
      var v = $('addr-pick').value;
      if (v === 'mine') $('addr-text').value = fromAddress();
      else { var e = book().filter(function (k) { return k.id === v; })[0]; if (e) $('addr-text').value = e.lines.join('\n'); }
    });
    function placeAddr(all) {
      var lines = textLines($('addr-text').value);
      if (!lines.length) { say('Type an address or choose a saved one first.', 'warn'); $('addr-text').focus(); return; }
      var nums = all ? openList() : [nextOpen()].filter(Boolean);
      if (!nums.length) { say('There are no open labels on this sheet.', 'warn'); return; }
      var d = designParas(lines, currentDesign(), st.fmt);
      if (currentDesign() === 'logo' && !d.pic) { say('Choose a logo first, or pick another design.', 'warn'); return; }
      queueChange(function () { return writeLabels(nums, d.paras, d.pic); }).then(refresh).then(function () { say('Put the address on ' + labelList(nums) + '.'); });
    }
    $('addr-one').addEventListener('click', function () { placeAddr(false); });
    $('addr-all').addEventListener('click', function () { placeAddr(true); });
    $('addr-save').addEventListener('click', function () {
      var lines = textLines($('addr-text').value);
      if (!lines.length) { say('Type an address first.', 'warn'); return; }
      say(addToBook([lines]) ? 'Saved ' + lines[0] + ' in the address book.' : 'That address is already in the address book.');
      renderAddress();
    });
    $('addr-mine').addEventListener('click', function () {
      var v = $('addr-text').value.replace(/\s+$/, '');
      if (!v.trim()) { say('Type your return address first.', 'warn'); return; }
      saveFrom(v); renderAddress(); say('Saved as your return address.');
    });
    var logoIn = $('logo-file');
    logoIn.addEventListener('change', function () {
      var file = logoIn.files && logoIn.files[0]; logoIn.value = '';
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function () {
        var img = new Image();
        img.onload = function () {
          var s = Math.min(1, 600 / Math.max(img.naturalWidth, img.naturalHeight));
          var c = document.createElement('canvas'); c.width = Math.round(img.naturalWidth * s); c.height = Math.round(img.naturalHeight * s);
          var x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(img, 0, 0, c.width, c.height);
          storeJson(LS.logo, { b64: c.toDataURL('image/png').split(',')[1], w: c.width, h: c.height });
          renderAddress(); say('Logo saved.');
        };
        img.onerror = function () { say('That picture couldn’t be opened. Try a PNG or JPEG.', 'warn'); };
        img.src = reader.result;
      };
      reader.readAsDataURL(file);
    });

    // Shipping address labels (From / Ship to)
    var from = $('from-addr');
    from.value = fromAddress();
    from.addEventListener('input', function () { saveFrom(from.value); });
    $('to-pick').addEventListener('change', function () {
      var e = book().filter(function (k) { return k.id === $('to-pick').value; })[0];
      if (e) $('to-addr').value = e.lines.join('\n');
    });
    $('to-save').addEventListener('click', function () {
      var lines = textLines($('to-addr').value);
      if (!lines.length) { say('Type the Ship to address first.', 'warn'); return; }
      say(addToBook([lines]) ? 'Saved ' + lines[0] + ' in the address book.' : 'That address is already in the address book.');
      renderAddress();
    });
    $('ship-addr-place').addEventListener('click', function () {
      var to = $('to-addr').value.replace(/\s+$/, '');
      if (!to.trim()) { say('Type the Ship to address first.', 'warn'); $('to-addr').focus(); return; }
      var n = nextOpen();
      if (!n) { say('There are no open labels on this sheet.', 'warn'); return; }
      var fromText = from.value.replace(/\s+$/, '');
      saveFrom(fromText);
      queueChange(function () { return writeLabels([n], shippingParas(fromText, to, st.fmt)); }).then(refresh).then(function () {
        $('to-addr').value = '';
        say('Address label placed on ' + theName(n) + '. You can edit it in the document.');
      });
    });

    // Address book import and backup
    var bookIn = $('book-file');
    bookIn.addEventListener('change', function () {
      var file = bookIn.files && bookIn.files[0]; bookIn.value = '';
      if (!file) return;
      var name = (file.name || '').toLowerCase(), reader = new FileReader();
      var done = function (lists) {
        var added = addToBook(lists);
        renderAddress();
        say(lists.length ? 'Added ' + plural(added, 'address', 'es') + ' to the address book' + (lists.length > added ? ' (' + (lists.length - added) + ' were already there).' : '.') : 'No addresses with a street and city were found in that file.', lists.length ? '' : 'warn');
      };
      var fromRows = function (rows) { return T.rowsToAddresses(rows).map(T.addressLines).filter(function (l) { return l.length > 1; }); };
      if (/\.xlsx$/.test(name)) {
        reader.onload = function () { T.readXlsx(reader.result).then(function (rows) { done(fromRows(rows)); }).catch(function () { say('That Excel file couldn’t be read.', 'warn'); }); };
        reader.readAsArrayBuffer(file);
      } else {
        reader.onload = function () {
          var text = String(reader.result || '');
          if (/\.vcf$/.test(name) || /BEGIN:VCARD/i.test(text)) done(T.parseVCard(text).map(function (c) { return c.lines; }));
          else done(fromRows(T.parseDelimited(text)));
        };
        reader.readAsText(file);
      }
    });
    $('book-export').addEventListener('click', function () {
      var b = book();
      if (!b.length) { say('The address book is empty.', 'warn'); return; }
      var q = function (s) { return '"' + String(s).replace(/"/g, '""') + '"'; };
      var csv = 'Name,Address\n' + b.map(function (e) { return q(e.name) + ',' + q(e.lines.slice(1).join('\n')); }).join('\n');
      var fallback = function () {
        var ta = $('book-csv'); ta.hidden = false; ta.value = csv; ta.focus(); ta.select();
        say('Press ⌘C to copy the selected addresses.');
      };
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(csv).then(function () { say('Copied ' + plural(b.length, 'address', 'es') + '. Paste them into Excel or a text file to keep a backup.'); }, fallback);
      } else fallback();
    });
  }

  // Labels that stand in for file pickers open them with Enter or Space too.
  function wireFileLabels() {
    Array.prototype.forEach.call(document.querySelectorAll('label[for][tabindex]'), function (lab) {
      lab.addEventListener('keydown', function (e) {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); var inp = document.getElementById(lab.getAttribute('for')); if (inp) inp.click(); }
      });
    });
  }

  // ---------- start ----------
  function demo() {
    // Preview outside Word: taskpane.html?demo (address sheet), ?demo=ship, ?demo&choose
    var shipDemo = /demo=ship/.test(location.search);
    var f = shipDemo ? F.formats.filter(function (x) { return x.shipping; })[0] : F.formats[0];
    st = { kind: 'sheet', fmt: f, used: {}, ready: {}, texts: {}, pics: {}, over: {}, here: shipDemo ? 1 : 8, header: true, margins: { top: S.STRIP_MARGIN, left: f.firstLabel.left }, applied: { x: 0, y: 0 } };
    if (!shipDemo) {
      for (var n = 1; n <= 7; n++) st.used[n] = true;
      [8, 9, 10].forEach(function (k) { st.ready[k] = true; st.texts[k] = 'Sample Name\n123 Main St\nBoston, MA 02101'; });
      st.over[10] = true;
    }
    if (/choose/.test(location.search)) st.kind = 'none';
    window.LSK_TEST = { ship: ship, output: output, handleFile: handleFile, st: function () { return st; } };
    render();
  }

  function start() {
    if (typeof Office === 'undefined') { show('v-outside'); return; }
    Office.onReady(function (info) {
      if (!info || info.host !== Office.HostType.Word) {
        if (/[?&]demo\b/.test(location.search)) { wire(); demo(); } else show('v-outside');
        return;
      }
      if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) { show('v-old'); return; }
      wire();
      refresh();
      startPrintWatch();
      // Learn how Word attaches this add-in, so new sheets can open with the panel.
      findAddinRef().then(function () { if (st.kind === 'sheet') renderAbout(); });
      try { Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged); } catch (e) { /* older Word */ }
      window.LSK_TEST = { st: function () { return st; }, ship: ship, output: output, handleFile: handleFile, fillFromRows: fillFromRows, checkPrinted: checkPrinted };
    });
  }
  start();
})();
