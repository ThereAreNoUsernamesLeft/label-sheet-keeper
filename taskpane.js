/*
  Label Sheet Keeper — the Word side panel.

  Words used throughout (code and screen):
    label            one spot on the sheet, numbered 1, 2, 3… across each row
    open             blank and still on the sheet
    used             already peeled off (locked and left blank in the document)
    ready to print   has text or a picture that hasn't been printed yet
*/
(function () {
  'use strict';

  var VERSION = '1.2';
  var ADDIN_ID = '8d491c85-a76d-4d26-a6f5-b7b2dd4a1711';   // must match manifest.xml <Id>
  var F = window.LSK_FORMATS;
  var S = window.LSK_SHEET;

  // Names used to store things
  var USED_TAG = 'lsk-used';            // content control on a used label
  var HISTORY_KEY = 'lskHistory';       // document setting: print history
  var PRINT_KEY = 'lskLastPrint';       // document setting: last print time seen
  var AUTOSHOW_KEY = 'Office.AutoShowTaskpaneWithDocument';
  var LS = {                            // this Mac's browser storage
    from: 'lsk.fromAddress',
    autoMark: 'lsk.autoMark',
    side: 'lsk.side.',
    shipTab: 'lsk.shipTab',
    addinRef: 'lsk.addinRef'
  };

  // A missing element (for example while Word still shows an older copy of
  // the panel page after an update) gets a stand-in so the panel keeps working.
  function $(id) { return document.getElementById(id) || document.createElement('div'); }

  var st = { kind: 'loading', fmt: null, used: {}, ready: {}, texts: {}, here: 0 };
  var undo = null;            // last print marked used: { nums, cells: [{n, ooxml}], at }
  var pendingConfirm = null;  // action for the confirm bar's main button
  var choosing = false;       // size chooser opened from a sheet
  var chosenId = F.formats.length ? F.formats[0].id : '';
  var chain = Promise.resolve();
  var busy = false;

  // ---------- small helpers ----------
  function store(key, value) {
    try {
      if (value === undefined) return localStorage.getItem(key);
      localStorage.setItem(key, value);
    } catch (e) {}
    return null;
  }
  function plural(n, w) { return n + ' ' + w + (n === 1 ? '' : 's'); }
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
  function keys(obj) { return Object.keys(obj).map(Number).sort(function (a, b) { return a - b; }); }
  function openList() { var a = []; for (var n = 1; n <= total(); n++) if (!st.used[n] && !st.ready[n]) a.push(n); return a; }
  function nextOpen() { var o = openList(); return o.length ? o[0] : 0; }
  function userError(msg) { var e = new Error(msg); e.userMessage = msg; return e; }
  function fmtDate(ms) {
    try { return new Date(ms).toLocaleString(undefined, { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' }); }
    catch (e) { return String(new Date(ms)); }
  }
  function timeOnly(ms) {
    try { return new Date(ms).toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' }); } catch (e) { return ''; }
  }
  function firstLine(t) {
    var l = String(t || '').split(/[\r\n\u000b]+/).map(function (s) { return s.trim(); }).filter(Boolean);
    return l.length ? l[0].slice(0, 48) : '';
  }
  function settings() { try { return Office.context.document.settings; } catch (e) { return null; } }
  function getSetting(k) { var s = settings(); try { return s ? s.get(k) : null; } catch (e) { return null; } }
  function setSetting(k, v) { var s = settings(); try { if (s) { s.set(k, v); s.saveAsync(function () {}); } } catch (e) {} }

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
    var el = $('say');
    el.textContent = msg || '';
    el.className = 'say' + (tone === 'warn' ? ' warn' : '');
  }
  function showError(e) {
    var el = $('error');
    if (e && e.userMessage) { say(e.userMessage, 'warn'); el.hidden = true; return; }
    el.textContent = 'Word reported a problem' + (e && e.message ? ': ' + e.message : '') + '. Try again, or close and reopen the panel.';
    el.hidden = false;
    if (window.console) console.error(e, e && e.debugInfo);
  }
  function ask(msg, yesLabel, action, danger) {
    pendingConfirm = action;
    $('confirm-msg').textContent = msg;
    $('confirm-yes').textContent = yesLabel;
    $('confirm-yes').className = 'btn sm ' + (danger ? 'danger' : 'primary');
    $('confirm').className = 'confirm' + (danger ? ' danger' : '');
    $('confirm').hidden = false;
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
  function docUrl() { try { return Office.context.document.url || ''; } catch (e) { return ''; } }
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
  function tagForAutoOpen() { if (getSetting(AUTOSHOW_KEY) !== true) setSetting(AUTOSHOW_KEY, true); }

  var refState = 'unknown';   // 'found' | 'missing' | 'failed: …'
  function cachedRef() { try { return JSON.parse(store(LS.addinRef) || 'null'); } catch (e) { return null; } }

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
          store(LS.addinRef, JSON.stringify(ref));
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

  // ---------- reading the document ----------
  function readState() {
    return Word.run(function (ctx) {
      var props = ctx.document.properties.customProperties;
      props.load('items/key,items/value');
      var tables = ctx.document.body.tables;
      tables.load('items/rowCount');
      return ctx.sync().then(function () {
        var id = '';
        props.items.forEach(function (p) { if (p.key === S.PROP) id = String(p.value); });
        if (!id) return { kind: 'none' };
        var fmt = F.getFormat(id);
        if (!fmt) return { kind: 'broken', reason: 'This sheet uses a label size (“' + id + '”) that’s no longer in the size list. Add it back to formats.js on GitHub, or make a new sheet.' };
        if (!tables.items.length || tables.items[0].rowCount !== S.layout(fmt).rows.length) {
          return { kind: 'broken', reason: 'The label layout in this document was changed, so it no longer lines up with your sheet. Make a new sheet.' };
        }
        var table = tables.items[0];
        table.load('values');
        var ccs = ctx.document.body.contentControls; ccs.load('items/tag');
        var pics = ctx.document.body.inlinePictures; pics.load('items');
        var selCell = ctx.document.getSelection().parentTableCellOrNullObject; selCell.load('rowIndex,cellIndex');
        return ctx.sync().then(function () {
          var usedCells = [], picCells = [];
          ccs.items.forEach(function (cc) {
            if ((cc.tag || '').indexOf(USED_TAG) === 0) { var c = cc.parentTableCellOrNullObject; c.load('rowIndex,cellIndex'); usedCells.push(c); }
          });
          pics.items.forEach(function (p) { var c = p.parentTableCellOrNullObject; c.load('rowIndex,cellIndex'); picCells.push(c); });
          return ctx.sync().then(function () {
            var used = {}, ready = {}, texts = {};
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
              var n = S.labelAt(fmt, c.rowIndex, c.cellIndex);
              if (n && !used[n]) { ready[n] = true; texts[n] = texts[n] || '(picture)'; }
            });
            var here = selCell.isNullObject ? 0 : S.labelAt(fmt, selCell.rowIndex, selCell.cellIndex);
            return { kind: 'sheet', fmt: fmt, used: used, ready: ready, texts: texts, here: here };
          });
        });
      });
    });
  }

  function apply(s) {
    st = s;
    if (s.kind === 'sheet') tagForAutoOpen();
    else st.fmt = null;
    render();
  }
  function refresh() { return queue(function () { return readState().then(apply); }); }

  var selTimer = 0;
  function onSelectionChanged() {
    clearTimeout(selTimer);
    selTimer = setTimeout(function () { if (!busy) refresh(); }, 400);
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

  function markUsed(nums) { return Word.run(function (ctx) { return lockAsUsed(ctx, nums); }); }
  function markOpen(nums) { return Word.run(function (ctx) { return unlock(ctx, nums); }); }

  // Copy the label the cursor is in to the next open labels.
  function copyHere(copies) {
    return Word.run(function (ctx) {
      var fmt = st.fmt;
      var c = ctx.document.getSelection().parentTableCellOrNullObject;
      c.load('rowIndex,cellIndex');
      return ctx.sync().then(function () {
        if (c.isNullObject) throw userError('Click inside a label in the document first, then copy.');
        var n = S.labelAt(fmt, c.rowIndex, c.cellIndex);
        if (!n) throw userError('The cursor is in the gap between labels. Click inside a label.');
        if (st.used[n]) throw userError(cap(labelName(n)) + ' is used. Click inside the label that has your text.');
        if (!st.ready[n]) throw userError(cap(labelName(n)) + ' is empty. Type your label there first.');
        var targets = openList().slice(0, copies);
        if (!targets.length) throw userError('There are no open labels left on this sheet.');
        var ox = c.body.getOoxml();
        var srcParas = c.body.paragraphs; srcParas.load('items');
        return ctx.sync().then(function () {
          var bodies = targets.map(function (t) { var b = cellBody(ctx, t); b.insertOoxml(ox.value, 'Replace'); return b; });
          return ctx.sync().then(function () {
            // Word sometimes adds an empty paragraph after pasted content; remove it.
            var paraLists = bodies.map(function (b) { var p = b.paragraphs; p.load('items/text'); return p; });
            return ctx.sync().then(function () {
              paraLists.forEach(function (p) {
                if (p.items.length > srcParas.items.length) { var last = p.items[p.items.length - 1]; if (!last.text) last.delete(); }
              });
              return ctx.sync();
            });
          }).then(function () { return { from: n, targets: targets }; });
        });
      });
    }).catch(function (e) {
      if (e && e.userMessage) throw e;
      return copyPlain(copies);   // fallback: text and basic formatting only
    });
  }

  function copyPlain(copies) {
    return Word.run(function (ctx) {
      var fmt = st.fmt;
      var c = ctx.document.getSelection().parentTableCellOrNullObject;
      c.load('rowIndex,cellIndex');
      return ctx.sync().then(function () {
        if (c.isNullObject) throw userError('Click inside a label in the document first, then copy.');
        var n = S.labelAt(fmt, c.rowIndex, c.cellIndex);
        var paras = c.body.paragraphs;
        paras.load('items/text,items/alignment,items/font/name,items/font/size,items/font/bold,items/font/italic,items/font/color');
        return ctx.sync().then(function () {
          var targets = openList().slice(0, copies);
          targets.forEach(function (t) {
            var body = cellBody(ctx, t);
            body.clear();
            paras.items.forEach(function (sp, i) {
              var p = i === 0 ? body.paragraphs.getFirst() : body.insertParagraph('', 'End');
              p.insertText(sp.text, 'Replace');
              p.alignment = sp.alignment;
              p.font.set({ name: sp.font.name, size: sp.font.size, bold: sp.font.bold, italic: sp.font.italic, color: sp.font.color });
            });
          });
          return ctx.sync().then(function () { return { from: n, targets: targets, plain: true }; });
        });
      });
    });
  }

  // Mark every ready-to-print label used, keeping a copy for Undo.
  function markPrinted() {
    var nums = keys(st.ready);
    return Word.run(function (ctx) {
      var oxs = nums.map(function (n) { return cellBody(ctx, n).getOoxml(); });
      return ctx.sync().then(function () {
        undo = { nums: nums, cells: nums.map(function (n, i) { return { n: n, ooxml: oxs[i].value }; }), at: Date.now() };
        return lockAsUsed(ctx, nums);
      });
    }).then(function () {
      var texts = [];
      nums.forEach(function (n) { var t = firstLine(st.texts[n]); if (t && texts.indexOf(t) < 0) texts.push(t); });
      var h = history();
      h.unshift({ at: undo.at, nums: nums, text: texts.length ? texts[0] + (texts.length > 1 ? ' and ' + plural(texts.length - 1, 'other') : '') : '' });
      setSetting(HISTORY_KEY, h.slice(0, 200));
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
      if (h.length && h[0].at === u.at) { h.shift(); setSetting(HISTORY_KEY, h); }
      undo = null;
    });
  }

  function history() { var h = getSetting(HISTORY_KEY); return Array.isArray(h) ? h : []; }

  function createDoc(f, opts) {
    return S.buildDocx(f, null, opts).then(function (b64) {
      return Word.run(function (ctx) {
        var doc = ctx.application.createDocument(b64);
        return ctx.sync().then(function () { doc.open(); return ctx.sync(); });
      });
    });
  }
  function createSheet(id) {
    var f = F.getFormat(id);
    if (!f) return Promise.reject(userError('Choose a label size first.'));
    return findAddinRef().then(function (ref) { return createDoc(f, { addinRef: ref }).then(function () { return !!ref; }); });
  }

  // ---------- noticing prints ----------
  // Word stamps a "last printed" time on the document when it prints. The
  // panel checks it every few seconds; when it moves forward, the labels that
  // were ready to print have just been printed.
  var lastPrintSeen = null, watchTimer = 0;
  function autoMarkOn() { return store(LS.autoMark) !== 'off'; }

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
        var saved = Number(getSetting(PRINT_KEY)) || 0;
        lastPrintSeen = saved || ms;
        if (!saved) setSetting(PRINT_KEY, ms);
        showLastPrint(ms);
        if (!(saved && ms > saved)) return;
      }
      if (ms <= lastPrintSeen) return;
      lastPrintSeen = ms;
      setSetting(PRINT_KEY, ms);
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
      if (autoMarkOn()) {
        queueChange(markPrinted).then(function (done) {
          return refresh().then(function () {
            say('Word printed this sheet' + when + ', so ' + labelList(done) + (done.length === 1 ? ' is' : ' are') + ' now marked used.');
          });
        });
      } else {
        ask('Word printed this sheet' + when + '. Mark ' + labelList(nums) + ' used?', 'Mark used', function () {
          queueChange(markPrinted).then(function (done) { return refresh().then(function () { say('Marked ' + labelList(done) + ' used.'); }); });
        });
      }
    });
  }

  function startPrintWatch() {
    clearInterval(watchTimer);
    watchTimer = setInterval(checkPrinted, 2500);
    document.addEventListener('visibilitychange', function () { if (!document.hidden) checkPrinted(); });
    setTimeout(checkPrinted, 800);
  }

  // ---------- shipping labels: label files ----------
  var DPI = 300;
  var ship = { doc: null, page: 1, pages: 1, src: null, crop: null, turn: 0 };

  // Usable area inside one label, in inches: inside the side padding, below
  // the top space, and clear of the printer's no-print strip at the bottom.
  function contentBox(f) {
    var minRow = Infinity;
    S.layout(f).rows.forEach(function (r) { if (r.label) minRow = Math.min(minRow, r.h / 1440); });
    var top = f.align === 'top' ? f.topSpace : 0.3;
    return { w: f.label.width - f.padding.left - f.padding.right, h: Math.max(0.5, minRow - top - 0.3), top: top, rowH: minRow };
  }

  var pdfLib = null;
  function loadPdfLib() {
    if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
    if (pdfLib) return pdfLib;
    var fail = function () { pdfLib = null; throw userError('The PDF reader didn’t load. Check your internet connection and try again.'); };
    pdfLib = new Promise(function (resolve, reject) {
      var s = document.createElement('script');
      s.src = 'vendor/pdfjs/pdf.js';
      s.onload = function () { window.pdfjsLib ? resolve(window.pdfjsLib) : reject(); };
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
    say('Opening ' + file.name + '…');
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
    ship.crop = trimBox(ship.src);
    $('crop').hidden = false;
    drawCrop();
    say(ship.pages > 1 ? 'This PDF has ' + ship.pages + ' pages. Use the arrows to find the page with the label.' : 'Drag across the label if the page has anything else on it.');
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

  function placeImage(n, b64, wPt, hPt) {
    var f = st.fmt;
    return Word.run(function (ctx) {
      var body = cellBody(ctx, n);
      body.clear();
      var pic = body.insertInlinePictureFromBase64(b64, 'Start');
      pic.width = wPt;
      pic.height = hPt;
      var p = body.paragraphs.getFirst();
      p.alignment = 'Centered';
      p.spaceBefore = f.align === 'top' ? f.topSpace * 72 : 0;
      p.spaceAfter = 0;
      return ctx.sync();
    });
  }

  // ---------- shipping labels: address labels ----------
  function placeAddress(n, fromText, toText) {
    var f = st.fmt, topPt = f.align === 'top' ? f.topSpace * 72 : 0;
    var fromLines = fromText ? fromText.split(/\r?\n/) : [];
    var lines = [];
    if (fromLines.length) {
      lines.push({ t: 'FROM', size: 8, bold: true, before: topPt });
      fromLines.forEach(function (l) { lines.push({ t: l, size: 11 }); });
    }
    lines.push({ t: 'SHIP TO', size: 9, bold: true, before: fromLines.length ? 90 : topPt + 90, indent: 36 });
    toText.split(/\r?\n/).forEach(function (l, i) { lines.push({ t: l, size: 16, bold: i === 0, indent: 36 }); });
    return Word.run(function (ctx) {
      var body = cellBody(ctx, n);
      body.clear();
      lines.forEach(function (ln, i) {
        var p = i === 0 ? body.paragraphs.getFirst() : body.insertParagraph(ln.t, 'End');
        if (i === 0) p.insertText(ln.t, 'Replace');
        p.alignment = 'Left';
        p.spaceBefore = ln.before || 0;
        p.spaceAfter = 0;
        p.leftIndent = ln.indent || 0;
        p.firstLineIndent = 0;
        p.font.set({ name: f.font.name, size: ln.size, bold: !!ln.bold, italic: false, color: '#000000' });
      });
      return ctx.sync();
    });
  }

  function loadFrom() { return store(LS.from) || getSetting(LS.from) || ''; }
  function saveFrom(v) { store(LS.from, v); setSetting(LS.from, v); }

  // ---------- rendering ----------
  function show(view) {
    ['v-loading', 'v-outside', 'v-old', 'v-none', 'v-broken', 'v-sheet'].forEach(function (id) { $(id).hidden = id !== view; });
  }

  function render() {
    if (st.kind === 'none' || choosing) { renderChooser(); show('v-none'); return; }
    if (st.kind === 'broken') { $('broken-msg').textContent = st.reason; show('v-broken'); return; }
    if (st.kind !== 'sheet') return;
    var f = st.fmt, ready = keys(st.ready).length;
    show('v-sheet');
    $('fmt-name').textContent = f.name;
    $('fmt-detail').textContent = f.detail;
    renderMap();
    renderTally();
    $('first-row').hidden = total() <= 2;
    $('first-n').max = String(total());
    $('step-copy').hidden = f.shipping;
    $('step-ship').hidden = !f.shipping;
    if (f.shipping) renderShip(); else renderHere();
    $('copies').max = String(total());
    $('printed-btn').textContent = ready ? 'I printed these (' + ready + ')' : 'I printed these';
    $('printed-btn').disabled = !ready;
    $('undo-row').hidden = !undo;
    if (undo) $('undo-msg').textContent = cap(labelList(undo.nums)) + (undo.nums.length === 1 ? ' was' : ' were') + ' marked used.';
    $('save-hint').hidden = !(needsFirstSave && !docUrl());
    renderHistory();
    renderAbout();
  }

  function thumb(f) {
    var t = document.createElement('span');
    t.className = 'thumb';
    t.style.height = Math.round(30 * f.page.height / f.page.width) + 'px';
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
    var list = $('fmt-list');
    list.innerHTML = '';
    F.formats.forEach(function (f) {
      var lab = document.createElement('label'); lab.className = 'fmt';
      var inp = document.createElement('input'); inp.type = 'radio'; inp.name = 'fmt'; inp.value = f.id; inp.checked = f.id === chosenId;
      inp.addEventListener('change', function () { chosenId = f.id; });
      var text = document.createElement('span');
      var b = document.createElement('b'); b.textContent = f.name;
      var d = document.createElement('span'); d.textContent = f.detail;
      text.appendChild(b); text.appendChild(d);
      lab.appendChild(inp); lab.appendChild(thumb(f)); lab.appendChild(text);
      list.appendChild(lab);
    });
    var links = $('tpl-links'); links.innerHTML = '';
    F.formats.forEach(function (f, i) {
      if (i) links.appendChild(document.createTextNode(' · '));
      var a = document.createElement('a'); a.href = 'templates/' + f.id + '.docx'; a.target = '_blank'; a.rel = 'noopener'; a.textContent = f.name;
      links.appendChild(a);
    });
    $('fmt-problems').hidden = !F.invalid.length;
    $('fmt-problems').textContent = F.invalid.map(function (x) { return 'The size “' + x.id + '” in formats.js was skipped: ' + x.problems.join('; ') + '.'; }).join(' ');
    $('choose-title').textContent = choosing ? 'Make another sheet' : 'Make a label sheet';
    $('choose-lead').textContent = choosing
      ? 'Choose the labels for the new sheet. It opens as a new Word document.'
      : 'This document isn’t a label sheet yet. Choose the labels you have, and the sheet opens as a new Word document laid out to match them.';
    $('back-btn').hidden = !(choosing && st.kind === 'sheet');
  }

  var mapBuiltFor = '';
  function renderMap() {
    var f = st.fmt, map = $('map');
    if (mapBuiltFor !== f.id) {
      map.innerHTML = '';
      map.style.aspectRatio = f.page.width + ' / ' + f.page.height;
      for (var n = 1; n <= total(); n++) {
        var r = Math.floor((n - 1) / f.across), c = (n - 1) % f.across;
        var b = document.createElement('button');
        b.type = 'button'; b.className = 'cell'; b.dataset.n = n;
        b.style.left = ((f.firstLabel.left + c * f.pitch.across) / f.page.width * 100) + '%';
        b.style.top = ((f.firstLabel.top + r * f.pitch.down) / f.page.height * 100) + '%';
        b.style.width = (f.label.width / f.page.width * 100) + '%';
        b.style.height = (f.label.height / f.page.height * 100) + '%';
        b.style.borderRadius = f.corner ? (f.corner / f.page.width * 100) + 'cqw' : '0';
        b.innerHTML = '<span class="t"></span><span class="num" aria-hidden="true">' + n + '</span><span class="side" aria-hidden="true"></span>';
        map.appendChild(b);
      }
      mapBuiltFor = f.id;
    }
    Array.prototype.forEach.call(map.children, function (b) {
      var n = +b.dataset.n;
      var state = st.used[n] ? 'used' : st.ready[n] ? 'ready' : 'open';
      b.className = 'cell ' + state + (st.here === n ? ' here' : '');
      b.querySelector('.t').textContent = state === 'ready' ? String(st.texts[n] || '').split(/[\r\n\u000b]+/).slice(0, 3).join('\n') : '';
      b.querySelector('.side').textContent = sideWord(n);
      b.setAttribute('aria-label', cap(labelName(n)) + ', ' + (state === 'ready' ? 'ready to print: ' + firstLine(st.texts[n]) : state));
    });
  }

  function renderTally() {
    var t = $('tally'); t.innerHTML = '';
    var counts = { open: openList().length, used: keys(st.used).length, ready: keys(st.ready).length };
    [['open', 'open'], ['used', 'used'], ['ready', 'ready to print']].forEach(function (p) {
      var s = document.createElement('span');
      s.innerHTML = '<i class="sw ' + p[0] + '"></i><b>' + counts[p[0]] + '</b>';
      s.appendChild(document.createTextNode(p[1]));
      t.appendChild(s);
    });
  }

  function renderHere() {
    var el = $('here'), n = st.here;
    el.innerHTML = '';
    if (!n) { el.textContent = 'The cursor isn’t in a label.'; return; }
    el.appendChild(document.createTextNode('The cursor is in '));
    var b = document.createElement('b'); b.textContent = labelName(n); el.appendChild(b);
    el.appendChild(document.createTextNode(st.used[n] ? ', which is used.' : st.ready[n] ? '.' : ', which is empty.'));
  }

  function renderShip() {
    var n = nextOpen(), f = st.fmt, applies = sideApplies(f), side = getSide(f);
    $('ship-target').textContent = n
      ? 'Next open: ' + theName(n) + '.'
      : 'There are no open labels on this sheet. When you load a fresh sheet, click Reset for a fresh sheet.';
    $('side-row').hidden = !applies;
    $('side-q').textContent = side
      ? 'Label 1 prints on the ' + side + ' half of your sheet. Change this if it’s wrong:'
      : 'Print the test page once. Which half of your sheet did Label 1 come out on?';
    $('side-top').setAttribute('aria-pressed', String(side === 'top'));
    $('side-bottom').setAttribute('aria-pressed', String(side === 'bottom'));
    var tab = store(LS.shipTab) === 'addr' ? 'addr' : 'file';
    $('tab-file').setAttribute('aria-selected', String(tab === 'file'));
    $('tab-addr').setAttribute('aria-selected', String(tab === 'addr'));
    $('pane-file').hidden = tab !== 'file';
    $('pane-addr').hidden = tab !== 'addr';
    $('ship-place').disabled = !n;
    $('addr-place').disabled = !n;
    if (ship.src) drawFit();
  }

  function renderHistory() {
    var ul = $('hist'); ul.innerHTML = '';
    var h = history();
    if (!h.length) { ul.innerHTML = '<li class="txt">Nothing printed yet.</li>'; return; }
    h.slice(0, 20).forEach(function (e) {
      var li = document.createElement('li');
      var w = document.createElement('span'); w.className = 'when'; w.textContent = fmtDate(e.at);
      var what = document.createElement('span'); what.textContent = cap(labelList(e.nums || []));
      li.appendChild(w); li.appendChild(what);
      if (e.text) { var t = document.createElement('span'); t.className = 'txt'; t.textContent = '“' + e.text + '”'; li.appendChild(t); }
      ul.appendChild(li);
    });
  }

  function renderAbout() {
    var tagged = getSetting(AUTOSHOW_KEY) === true, saved = !!docUrl(), ref = cachedRef();
    var rows = [
      ['Version', VERSION],
      ['This sheet', tagged ? (saved ? 'Opens with the panel' : 'Opens with the panel once you save it') : 'Not set to open with the panel yet'],
      ['New sheets', ref ? 'Open with the panel' : refState === 'missing' ? 'Need one more step: save this document, close it, open it again, then make the sheet'
        : refState.indexOf('failed') === 0 ? 'Can’t open with the panel (' + refState.slice(8) + ')' : 'Checking…'],
      ['Add-in link', ref ? (ref.storeType || '?') + ' · ' + (ref.store || '?') : '—']
    ];
    var dl = $('about-list'); dl.innerHTML = '';
    rows.forEach(function (r) {
      var dt = document.createElement('dt'); dt.textContent = r[0];
      var dd = document.createElement('dd'); dd.textContent = r[1];
      dl.appendChild(dt); dl.appendChild(dd);
    });
  }

  // ---------- wiring ----------
  function wire() {
    // Sheets
    $('create-btn').addEventListener('click', function () {
      var f = F.getFormat(chosenId);
      queue(function () { return createSheet(chosenId); }).then(function (withPanel) {
        var msg = 'Your new sheet (' + f.name + ') opened in its own window' +
          (withPanel ? ', with this panel.' : '. Click Label Sheet on its Home tab to open the panel there.');
        if (st.kind === 'sheet') { choosing = false; render(); say(msg); }
        else { $('made').hidden = false; $('made-msg').textContent = msg; say(''); }
        renderAbout();
      });
    });
    $('test-btn').addEventListener('click', function () {
      if (!st.fmt) return;
      var f = st.fmt;
      queue(function () { return createDoc(f, { test: true }); }).then(function () {
        say('A test page opened in its own window. Print it on plain paper at 100% scale.');
      });
    });
    $('broken-new').addEventListener('click', function () { choosing = true; render(); });
    $('back-btn').addEventListener('click', function () { choosing = false; render(); });
    $('other-btn').addEventListener('click', function () { choosing = true; $('made').hidden = true; render(); });
    $('refresh-btn').addEventListener('click', function () { refresh().then(function () { say('Up to date.'); }); });
    $('reset-btn').addEventListener('click', function () {
      var used = keys(st.used);
      if (!used.length) { say('Every label on this sheet is already open.'); return; }
      ask('Loaded a fresh sheet? All ' + plural(used.length, 'used label') + ' become open again. Labels ready to print stay as they are.', 'Reset', function () {
        undo = null;
        queueChange(function () { return markOpen(used); }).then(refresh).then(function () { say('Reset. Every label without text is open.'); });
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
      var withText = nums.filter(function (n) { return st.ready[n]; });
      if (withText.length) ask(cap(labelList(withText)) + (withText.length === 1 ? ' has' : ' have') + ' text on it. Clear and mark used?', 'Clear and mark used', run, true);
      else run();
    });

    // Step 2: fill labels
    function doCopy(k, all) {
      queueChange(function () { return readState().then(apply).then(function () { return copyHere(k); }); }).then(function (res) {
        return refresh().then(function () {
          var short = !all && res.targets.length < k ? ' Only ' + plural(res.targets.length, 'open label') + ' were left.' : '';
          var plain = res.plain ? ' Pictures and mixed formatting weren’t copied.' : '';
          say('Copied ' + labelName(res.from) + ' to ' + labelList(res.targets) + '.' + plain + short, short || plain ? 'warn' : '');
        });
      });
    }
    $('copy-btn').addEventListener('click', function () {
      var k = Math.round(Number($('copies').value));
      if (!(k >= 1)) { say('Enter how many copies you want.', 'warn'); return; }
      doCopy(k, false);
    });
    $('fill-all').addEventListener('click', function () { doCopy(total(), true); });

    // Step 2 (shipping)
    wireShip();

    // Step 3: print
    var auto = $('auto-mark');
    auto.checked = autoMarkOn();
    auto.addEventListener('change', function () { store(LS.autoMark, auto.checked ? 'on' : 'off'); });
    $('printed-btn').addEventListener('click', function () {
      if (!keys(st.ready).length) { say('No labels are ready to print.', 'warn'); return; }
      queueChange(markPrinted).then(function (nums) {
        return refresh().then(function () { say('Marked ' + labelList(nums) + ' used. ' + plural(openList().length, 'open label') + ' left.'); });
      });
      readPrintTime().then(function (ms) { if (ms) { lastPrintSeen = Math.max(lastPrintSeen || 0, ms); setSetting(PRINT_KEY, lastPrintSeen); } }).catch(function () {});
    });
    $('undo-btn').addEventListener('click', function () {
      if (!undo) return;
      queueChange(undoPrinted).then(refresh).then(function () { say('Undone. Those labels have their text back.'); });
    });

    $('confirm-yes').addEventListener('click', function () { var a = pendingConfirm; closeAsk(); if (a) a(); });
    $('confirm-no').addEventListener('click', closeAsk);
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
    [['tab-file', 'file'], ['tab-addr', 'addr']].forEach(function (t) {
      $(t[0]).addEventListener('click', function () { store(LS.shipTab, t[1]); render(); });
    });

    var input = $('ship-file');
    input.addEventListener('change', function () { handleFile(input.files && input.files[0]); input.value = ''; });
    $('ship-file-label').addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
    var card = $('pane-file');
    card.addEventListener('dragover', function (e) { e.preventDefault(); card.classList.add('drop'); });
    card.addEventListener('dragleave', function () { card.classList.remove('drop'); });
    card.addEventListener('drop', function (e) {
      e.preventDefault(); card.classList.remove('drop');
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) handleFile(f);
    });

    $('pg-prev').addEventListener('click', function () { if (ship.doc && ship.page > 1) { ship.page--; renderPdfPage().then(afterSource).catch(fileError); } });
    $('pg-next').addEventListener('click', function () { if (ship.doc && ship.page < ship.pages) { ship.page++; renderPdfPage().then(afterSource).catch(fileError); } });
    $('crop-auto').addEventListener('click', function () { if (ship.src) { ship.crop = trimBox(ship.src); drawCrop(); } });
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
      try { wrap.setPointerCapture(e.pointerId); } catch (err) {}
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
      queueChange(function () { return placeImage(n, b64, out.wIn * 72, out.hIn * 72); }).then(refresh).then(function () {
        say('Placed on ' + theName(n) + '. Print with ⌘P.');
      });
    });

    var from = $('from-addr');
    from.value = loadFrom();
    from.addEventListener('input', function () { saveFrom(from.value); });
    $('addr-place').addEventListener('click', function () {
      var to = $('to-addr').value.replace(/\s+$/, '');
      if (!to.trim()) { say('Type the Ship to address first.', 'warn'); $('to-addr').focus(); return; }
      var n = nextOpen();
      if (!n) { say('There are no open labels on this sheet.', 'warn'); return; }
      var fromText = from.value.replace(/\s+$/, '');
      saveFrom(fromText);
      queueChange(function () { return placeAddress(n, fromText, to); }).then(refresh).then(function () {
        $('to-addr').value = '';
        say('Address label placed on ' + theName(n) + '. You can edit it in the document.');
      });
    });
  }

  // ---------- start ----------
  function demo() {
    // Preview outside Word: taskpane.html?demo (address sheet) or ?demo=ship.
    var shipDemo = /demo=ship/.test(location.search);
    var f = shipDemo ? F.formats.filter(function (x) { return x.shipping; })[0] : F.formats[0];
    st = { kind: 'sheet', fmt: f, used: {}, ready: {}, texts: {}, here: shipDemo ? 1 : 8 };
    if (!shipDemo) {
      for (var n = 1; n <= 7; n++) st.used[n] = true;
      [8, 9, 10].forEach(function (n) { st.ready[n] = true; st.texts[n] = 'Sample Name\n123 Main St\nBoston, MA 02101'; });
    }
    if (/choose/.test(location.search)) st.kind = 'none';
    window.LSK_TEST = { ship: ship, output: output, handleFile: handleFile };
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
      try { Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged); } catch (e) {}
    });
  }
  start();
})();
