/* Label Sheet Keeper — Word side panel. */
(function () {
  'use strict';

  var F = window.LSK_FORMATS;
  var S = window.LSK_SHEET;
  var USED_TAG = 'lsk-used';
  var HISTORY_KEY = 'lskHistory';
  // A missing element (for example, Word still showing an older copy of the
  // panel page for a few minutes after an update) gets a stand-in so the
  // panel keeps working instead of stopping.
  var $ = function (id) { return document.getElementById(id) || document.createElement('div'); };

  var st = { kind: 'loading', fmt: null, used: {}, filled: {}, texts: {}, here: 0 };
  var undo = null;           // { nums: [], cells: [{n, ooxml}], at }
  var pendingConfirm = null; // function to run on "yes"
  var choosing = false;      // showing the size chooser from a sheet
  var chosenId = F && F.formats.length ? F.formats[0].id : '';
  var chain = Promise.resolve();
  var busy = false;

  // ---------- small helpers ----------
  function show(view) {
    ['v-loading', 'v-outside', 'v-old', 'v-none', 'v-broken', 'v-sheet'].forEach(function (id) { $(id).hidden = id !== view; });
  }
  function say(msg, tone) { var el = $('say'); if (!el) return; el.textContent = msg || ''; el.className = 'say' + (tone === 'warn' ? ' warn' : ''); }
  function plural(n, w) { return n + ' ' + w + (n === 1 ? '' : 's'); }
  function ranges(list) {
    var a = list.slice().sort(function (x, y) { return x - y; }), out = [], i = 0;
    while (i < a.length) { var j = i; while (j + 1 < a.length && a[j + 1] === a[j] + 1) j++; out.push(i === j ? String(a[i]) : a[i] + '–' + a[j]); i = j + 1; }
    return out.join(', ');
  }
  function labelWord(list) { return (list.length === 1 ? 'label ' : 'labels ') + ranges(list); }
  function total() { return st.fmt ? st.fmt.across * st.fmt.down : 0; }
  function keys(obj) { return Object.keys(obj).map(Number).sort(function (a, b) { return a - b; }); }
  function openList() { var a = []; for (var n = 1; n <= total(); n++) if (!st.used[n] && !st.filled[n]) a.push(n); return a; }
  function userError(msg) { var e = new Error(msg); e.userMessage = msg; return e; }
  function fmtDate(ms) {
    try { return new Date(ms).toLocaleString(undefined, { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' }); }
    catch (e) { return String(new Date(ms)); }
  }
  function firstLine(t) { var l = String(t || '').split(/[\r\n\u000b]+/).map(function (s) { return s.trim(); }).filter(Boolean); return l.length ? l[0].slice(0, 48) : ''; }

  function showError(e) {
    var el = $('error');
    var msg = (e && e.userMessage) ? e.userMessage
      : 'Word reported a problem' + (e && e.message ? ': ' + e.message : '') + '. Try again, or close and reopen the panel.';
    if (e && e.userMessage) { say(msg, 'warn'); el.hidden = true; }
    else { el.textContent = msg; el.hidden = false; }
    if (window.console && e && !e.userMessage) console.error(e, e && e.debugInfo);
  }
  function clearError() { $('error').hidden = true; }

  // Run Word tasks one at a time so they never overlap. If a task fails, the
  // error is shown and the caller's follow-up steps don't run.
  function queue(task) {
    var p = chain.then(function () { busy = true; clearError(); return task(); });
    chain = p.then(function () { busy = false; }, function (e) { busy = false; showError(e); });
    return p.then(function (v) { return v; }, function () { return new Promise(function () {}); });
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
        if (!fmt) return { kind: 'broken', reason: 'This sheet uses a label size (“' + id + '”) that isn’t in the size list anymore. Add it back to formats.js, or make a new sheet.' };
        var L = S.layout(fmt);
        if (!tables.items.length || tables.items[0].rowCount !== L.rows.length) {
          return { kind: 'broken', reason: 'The label table in this document was changed or removed, so the spots no longer line up with the sheet. Make a new label sheet.' };
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
            var used = {}, filled = {}, texts = {};
            usedCells.forEach(function (c) { if (!c.isNullObject) { var n = S.labelAt(fmt, c.rowIndex, c.cellIndex); if (n) used[n] = true; } });
            var vals = table.values || [];
            for (var n = 1; n <= fmt.across * fmt.down; n++) {
              if (used[n]) continue;
              var ci = S.cellOf(fmt, n);
              var t = String((vals[ci.row] && vals[ci.row][ci.cell]) || '').replace(/ /g, ' ').trim();
              if (t) { filled[n] = true; texts[n] = t; }
            }
            picCells.forEach(function (c) {
              if (c.isNullObject) return;
              var n = S.labelAt(fmt, c.rowIndex, c.cellIndex);
              if (n && !used[n]) { filled[n] = true; texts[n] = texts[n] || '(picture)'; }
            });
            var here = selCell.isNullObject ? 0 : S.labelAt(fmt, selCell.rowIndex, selCell.cellIndex);
            return { kind: 'sheet', fmt: fmt, used: used, filled: filled, texts: texts, here: here };
          });
        });
      });
    });
  }

  function refresh() {
    return queue(function () { return readState().then(apply); });
  }
  function apply(s) {
    st = s;
    if (s.kind !== 'sheet') st.fmt = null;
    render();
  }
  var selTimer = 0;
  function onSelectionChanged() {
    clearTimeout(selTimer);
    selTimer = setTimeout(function () { if (!busy) refresh(); }, 400);
  }

  // ---------- changing the document ----------
  function lockAsUsed(ctx, table, nums) {
    var made = [];
    nums.forEach(function (n) {
      var ci = S.cellOf(st.fmt, n);
      var cell = table.getCell(ci.row, ci.cell);
      cell.body.clear();
      var r = cell.body.insertText(' ', 'Start');
      var cc = r.insertContentControl();
      cc.tag = USED_TAG;
      cc.title = 'Used';
      cc.cannotEdit = true;
      cc.cannotDelete = true;
      made.push(cc);
    });
    return ctx.sync().then(function () {
      // On-screen look only; tags and colors never print. Ignore if this Word can't show them.
      made.forEach(function (cc) { cc.appearance = 'Tags'; cc.color = '#7F8C8D'; });
      return ctx.sync().catch(function () {});
    });
  }

  function unlock(ctx, table, nums) {
    var lists = nums.map(function (n) {
      var ci = S.cellOf(st.fmt, n);
      var cs = table.getCell(ci.row, ci.cell).body.contentControls;
      cs.load('items/tag');
      return cs;
    });
    return ctx.sync().then(function () {
      var mine = [];
      lists.forEach(function (cs) { cs.items.forEach(function (cc) { if ((cc.tag || '').indexOf(USED_TAG) === 0) { cc.cannotDelete = false; cc.cannotEdit = false; mine.push(cc); } }); });
      return ctx.sync().then(function () { mine.forEach(function (cc) { cc.delete(false); }); return ctx.sync(); });
    });
  }

  function markUsed(nums) {
    return Word.run(function (ctx) { return lockAsUsed(ctx, ctx.document.body.tables.getFirst(), nums); });
  }
  function markOpen(nums) {
    return Word.run(function (ctx) { return unlock(ctx, ctx.document.body.tables.getFirst(), nums); });
  }

  function copyHere(copies) {
    return Word.run(function (ctx) {
      var fmt = st.fmt;
      var c = ctx.document.getSelection().parentTableCellOrNullObject;
      c.load('rowIndex,cellIndex');
      return ctx.sync().then(function () {
        if (c.isNullObject) throw userError('Click inside a label in the document first, then copy.');
        var n = S.labelAt(fmt, c.rowIndex, c.cellIndex);
        if (!n) throw userError('The cursor is in the gap between labels. Click inside a label.');
        if (st.used[n]) throw userError('Label ' + n + ' is used. Click inside the label that has your text.');
        if (!st.filled[n]) throw userError('Label ' + n + ' is empty. Type your label there first.');
        var targets = openList().filter(function (t) { return t !== n; }).slice(0, copies);
        if (!targets.length) throw userError('There are no open labels left on this sheet.');
        var ox = c.body.getOoxml();
        var srcParas = c.body.paragraphs; srcParas.load('items');
        return ctx.sync().then(function () {
          var table = ctx.document.body.tables.getFirst();
          var cells = targets.map(function (t) {
            var ci = S.cellOf(fmt, t);
            var tc = table.getCell(ci.row, ci.cell);
            tc.body.insertOoxml(ox.value, 'Replace');
            return tc;
          });
          return ctx.sync().then(function () {
            // Word sometimes adds an empty paragraph after pasted content; remove it.
            var paraLists = cells.map(function (tc) { var p = tc.body.paragraphs; p.load('items/text'); return p; });
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
      return copyPlain(copies); // fallback: copy text and basic formatting
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
          var targets = openList().filter(function (t) { return t !== n; }).slice(0, copies);
          var table = ctx.document.body.tables.getFirst();
          targets.forEach(function (t) {
            var ci = S.cellOf(fmt, t);
            var body = table.getCell(ci.row, ci.cell).body;
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

  function printedNow() {
    var nums = keys(st.filled);
    return Word.run(function (ctx) {
      var table = ctx.document.body.tables.getFirst();
      var oxs = nums.map(function (n) { var ci = S.cellOf(st.fmt, n); return table.getCell(ci.row, ci.cell).body.getOoxml(); });
      return ctx.sync().then(function () {
        undo = { nums: nums, cells: nums.map(function (n, i) { return { n: n, ooxml: oxs[i].value }; }), at: Date.now() };
        return lockAsUsed(ctx, table, nums);
      });
    }).then(function () {
      var texts = [];
      nums.forEach(function (n) { var t = firstLine(st.texts[n]); if (t && texts.indexOf(t) < 0) texts.push(t); });
      var summary = texts.length ? texts[0] + (texts.length > 1 ? ' and ' + plural(texts.length - 1, 'other') : '') : '';
      var h = history();
      h.unshift({ at: undo.at, nums: nums, text: summary });
      saveHistory(h.slice(0, 200));
      return nums;
    });
  }

  function undoPrinted() {
    var u = undo;
    return Word.run(function (ctx) {
      var table = ctx.document.body.tables.getFirst();
      return unlock(ctx, table, u.nums).then(function () {
        u.cells.forEach(function (c) { var ci = S.cellOf(st.fmt, c.n); table.getCell(ci.row, ci.cell).body.insertOoxml(c.ooxml, 'Replace'); });
        return ctx.sync();
      });
    }).then(function () {
      var h = history();
      if (h.length && h[0].at === u.at) { h.shift(); saveHistory(h); }
      undo = null;
    });
  }

  function createSheet(id, opts) {
    var f = F.getFormat(id);
    if (!f) return Promise.reject(userError('Pick a label size first.'));
    return S.buildDocx(f, null, opts).then(function (b64) {
      return Word.run(function (ctx) {
        var doc = ctx.application.createDocument(b64);
        return ctx.sync().then(function () { doc.open(); return ctx.sync(); });
      });
    });
  }

  // History lives in this document's settings, so it travels with the file.
  function history() {
    try { var h = Office.context.document.settings.get(HISTORY_KEY); return Array.isArray(h) ? h : []; } catch (e) { return []; }
  }
  function saveHistory(h) {
    try { Office.context.document.settings.set(HISTORY_KEY, h); Office.context.document.settings.saveAsync(function () {}); } catch (e) {}
  }

  // ---------- rendering ----------
  function render() {
    if (st.kind === 'none' || choosing) { renderChooser(); show('v-none'); $('fmt-name').textContent = ''; return; }
    if (st.kind === 'broken') { $('broken-msg').textContent = st.reason; show('v-broken'); $('fmt-name').textContent = ''; return; }
    if (st.kind !== 'sheet') return;
    $('fmt-name').textContent = st.fmt.name + ' · ' + st.fmt.detail;
    show('v-sheet');
    renderMap();
    renderTally();
    renderHere();
    renderHistory();
    $('step-copy').hidden = st.fmt.shipping;
    $('step-ship').hidden = !st.fmt.shipping;
    if (st.fmt.shipping) { renderShip(); if (ship.src) drawFit(); }
    var filled = keys(st.filled).length;
    $('printed-btn').textContent = filled ? 'I printed these (' + filled + ')' : 'I printed these';
    $('printed-btn').disabled = !filled;
    $('first-n').max = String(total());
    $('copies').max = String(total());
    $('undo-row').hidden = !undo;
    if (undo) $('undo-msg').textContent = 'Marked ' + labelWord(undo.nums) + ' used.';
  }

  function renderChooser() {
    var list = $('fmt-list');
    list.innerHTML = '';
    F.formats.forEach(function (f) {
      var lab = document.createElement('label'); lab.className = 'fmt';
      var inp = document.createElement('input'); inp.type = 'radio'; inp.name = 'fmt'; inp.value = f.id; inp.checked = f.id === chosenId;
      inp.addEventListener('change', function () { chosenId = f.id; });
      var b = document.createElement('b'); b.textContent = f.name;
      var s = document.createElement('span'); s.textContent = f.detail;
      lab.appendChild(inp); lab.appendChild(b); lab.appendChild(s);
      list.appendChild(lab);
    });
    var links = $('tpl-links'); links.innerHTML = '';
    F.formats.forEach(function (f, i) {
      if (i) links.appendChild(document.createTextNode(' · '));
      var a = document.createElement('a'); a.href = 'templates/' + f.id + '.docx'; a.target = '_blank'; a.rel = 'noopener'; a.textContent = f.name;
      links.appendChild(a);
    });
    var bad = $('fmt-problems');
    bad.hidden = !F.invalid.length;
    bad.textContent = F.invalid.map(function (x) { return 'Size “' + x.id + '” in formats.js was skipped: ' + x.problems.join('; ') + '.'; }).join(' ');
    var lead = document.querySelector('#v-none .lead');
    lead.textContent = choosing ? 'Pick the label size for the new sheet.' : 'This document isn’t a label sheet yet. Pick your label size to make one.';
    $('back-btn').hidden = !(choosing && st.kind !== 'none');
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
        var t = document.createElement('span'); t.className = 't';
        var num = document.createElement('span'); num.className = 'num'; num.textContent = n; num.setAttribute('aria-hidden', 'true');
        b.appendChild(t); b.appendChild(num);
        map.appendChild(b);
      }
      mapBuiltFor = f.id;
    }
    Array.prototype.forEach.call(map.children, function (b) {
      var n = +b.dataset.n;
      var state = st.used[n] ? 'used' : st.filled[n] ? 'filled' : 'open';
      b.className = 'cell ' + state + (st.here === n ? ' here' : '');
      b.firstChild.textContent = state === 'filled' ? String(st.texts[n] || '').split(/[\r\n\u000b]+/).slice(0, 3).join('\n') : '';
      b.setAttribute('aria-label', 'Label ' + n + ', ' + (state === 'filled' ? 'has text: ' + firstLine(st.texts[n]) : state));
    });
  }

  function renderTally() {
    var t = $('tally'); t.innerHTML = '';
    var counts = { open: openList().length, used: keys(st.used).length, filled: keys(st.filled).length };
    [['open', 'open'], ['used', 'used'], ['filled', 'filled, not printed']].forEach(function (p) {
      var s = document.createElement('span');
      var sw = document.createElement('i'); sw.className = 'sw ' + p[0];
      var b = document.createElement('b'); b.textContent = counts[p[0]];
      s.appendChild(sw); s.appendChild(b); s.appendChild(document.createTextNode(p[1]));
      t.appendChild(s);
    });
  }

  function renderHere() {
    var el = $('here');
    el.innerHTML = '';
    if (st.here) {
      el.appendChild(document.createTextNode('Cursor is in '));
      var b = document.createElement('b'); b.textContent = 'label ' + st.here; el.appendChild(b);
      el.appendChild(document.createTextNode(st.used[st.here] ? ' (used).' : st.filled[st.here] ? '.' : ' (empty).'));
    } else {
      el.textContent = 'Cursor isn’t in a label.';
    }
  }

  function renderHistory() {
    var ul = $('hist'); ul.innerHTML = '';
    var h = history();
    if (!h.length) { var li0 = document.createElement('li'); li0.className = 'txt'; li0.textContent = 'No prints marked yet.'; ul.appendChild(li0); return; }
    h.slice(0, 20).forEach(function (e) {
      var li = document.createElement('li');
      var w = document.createElement('span'); w.className = 'when'; w.textContent = fmtDate(e.at);
      var what = document.createElement('span'); what.textContent = labelWord(e.nums || []);
      li.appendChild(w); li.appendChild(what);
      if (e.text) { var t = document.createElement('span'); t.className = 'txt'; t.textContent = '“' + e.text + '”'; li.appendChild(t); }
      ul.appendChild(li);
    });
  }

  function ask(msg, yesLabel, action) {
    pendingConfirm = action;
    $('confirm-msg').textContent = msg;
    $('confirm-yes').textContent = yesLabel;
    $('confirm').hidden = false;
    $('confirm-no').focus();
  }
  function closeAsk() { pendingConfirm = null; $('confirm').hidden = true; }

  // ---------- wiring ----------
  function wire() {
    $('create-btn').addEventListener('click', function () {
      queue(function () { return createSheet(chosenId); }).then(function () {
        choosing = false;
        var f = F.getFormat(chosenId);
        say('');
        var msg = 'A new ' + f.name.toLowerCase() + ' document opened in its own window. Click Label Sheet on its Home tab to track it.';
        if (st.kind === 'sheet') { render(); say(msg); } else { document.querySelector('#v-none .lead').textContent = msg; }
      });
    });
    $('test-btn').addEventListener('click', function () {
      if (!st.fmt) return;
      queue(function () { return createSheet(st.fmt.id, { test: true }); }).then(function () {
        say('A test page opened in its own window. Print it on plain paper at 100% scale.');
      });
    });
    $('broken-new').addEventListener('click', function () { choosing = true; render(); });
    $('back-btn').addEventListener('click', function () { choosing = false; render(); });
    $('other-btn').addEventListener('click', function () { choosing = true; render(); });
    $('refresh-btn').addEventListener('click', function () { refresh().then(function () { say('Up to date.'); }); });

    $('map').addEventListener('click', function (ev) {
      var b = ev.target.closest('.cell'); if (!b || busy) return;
      var n = +b.dataset.n;
      if (st.used[n]) {
        queue(function () { return markOpen([n]); }).then(refresh).then(function () { say('Label ' + n + ' is open again.'); });
      } else if (st.filled[n]) {
        ask('Label ' + n + ' has text. Clear it and mark the spot used?', 'Clear and mark used', function () {
          queue(function () { return markUsed([n]); }).then(refresh).then(function () { say('Label ' + n + ' marked used.'); });
        });
      } else {
        queue(function () { return markUsed([n]); }).then(refresh).then(function () { say('Label ' + n + ' marked used.'); });
      }
    });

    $('first-btn').addEventListener('click', function () {
      var max = total(), k = Math.round(Number($('first-n').value));
      if (!(k >= 1 && k <= max)) { say('Enter a number from 1 to ' + max + '.', 'warn'); return; }
      var nums = []; for (var n = 1; n <= k; n++) if (!st.used[n]) nums.push(n);
      if (!nums.length) { say((k === 1 ? 'Label 1 was' : 'Labels 1–' + k + ' were') + ' already marked used.'); return; }
      var run = function () { queue(function () { return markUsed(nums); }).then(refresh).then(function () { say('Marked ' + labelWord(nums) + ' used.'); }); };
      var withText = nums.filter(function (n) { return st.filled[n]; });
      if (withText.length) ask(labelWord(withText).replace(/^l/, 'L') + ' ' + (withText.length === 1 ? 'has' : 'have') + ' text. Clear and mark them used?', 'Clear and mark used', run);
      else run();
    });

    $('copy-btn').addEventListener('click', function () {
      var k = Math.round(Number($('copies').value));
      if (!(k >= 1)) { say('Enter how many copies you want.', 'warn'); return; }
      queue(function () { return readState().then(apply).then(function () { return copyHere(k); }); })
        .then(function (res) {
          return refresh().then(function () {
            var note = res.targets.length < k ? ' Only ' + plural(res.targets.length, 'open spot') + ' were left.' : '';
            say('Copied label ' + res.from + ' to ' + labelWord(res.targets) + '.' + (res.plain ? ' Pictures and mixed formatting weren’t copied.' : '') + note, note || res.plain ? 'warn' : '');
          });
        });
    });

    $('printed-btn').addEventListener('click', function () {
      if (!keys(st.filled).length) { say('No labels have text to mark as printed.', 'warn'); return; }
      queue(printedNow).then(function (nums) { return refresh().then(function () { say('Marked ' + labelWord(nums) + ' used. ' + plural(openList().length, 'open spot') + ' left.'); }); });
      readPrintTime().then(function (ms) { if (ms) { lastPrintSeen = Math.max(lastPrintSeen || 0, ms); savePrintTime(lastPrintSeen); } }).catch(function () {});
    });
    $('undo-btn').addEventListener('click', function () {
      if (!undo) return;
      queue(undoPrinted).then(refresh).then(function () { say('Undone. Those labels have their text back.'); });
    });

    $('reset-btn').addEventListener('click', function () {
      var used = keys(st.used);
      if (!used.length) { say('Every spot on this sheet is already open.'); return; }
      ask('Start a new sheet? All ' + plural(used.length, 'used spot') + ' become open again. Labels with text stay as they are.', 'Start new sheet', function () {
        undo = null;
        queue(function () { return markOpen(used); }).then(refresh).then(function () { say('New sheet. All spots without text are open.'); });
      });
    });

    var auto = $('auto-mark');
    auto.checked = autoMarkOn();
    auto.addEventListener('change', function () { setAutoMark(auto.checked); });
    $('confirm-yes').addEventListener('click', function () { var a = pendingConfirm; closeAsk(); if (a) a(); });
    $('confirm-no').addEventListener('click', closeAsk);
    wireShip();
  }

  // ---------- noticing prints ----------
  // Word stamps a "last printed" time on the document whenever it prints.
  // The panel checks that time every few seconds; when it moves forward,
  // the filled labels were just printed.
  var PRINT_KEY = 'lskLastPrint';
  var AUTO_KEY = 'lsk.autoMark';
  var lastPrintSeen = null;   // ms; null until the first check
  var watchTimer = 0;

  function autoMarkOn() { try { return localStorage.getItem(AUTO_KEY) !== 'off'; } catch (e) { return true; } }
  function setAutoMark(on) { try { localStorage.setItem(AUTO_KEY, on ? 'on' : 'off'); } catch (e) {} }
  function savedPrintTime() { try { return Number(Office.context.document.settings.get(PRINT_KEY)) || 0; } catch (e) { return 0; } }
  function savePrintTime(ms) {
    try { Office.context.document.settings.set(PRINT_KEY, ms); Office.context.document.settings.saveAsync(function () {}); } catch (e) {}
  }
  function timeOnly(ms) {
    try { return new Date(ms).toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' }); } catch (e) { return ''; }
  }

  function readPrintTime() {
    return Word.run(function (ctx) {
      var props = ctx.document.properties;
      props.load('lastPrintDate');
      return ctx.sync().then(function () {
        var d = props.lastPrintDate;
        var ms = d ? new Date(d).getTime() : 0;
        return isFinite(ms) && ms > 0 && new Date(ms).getFullYear() > 1990 ? ms : 0;
      });
    });
  }

  function checkPrinted() {
    if (busy || st.kind !== 'sheet') return;
    readPrintTime().then(function (ms) {
      if (lastPrintSeen === null) {
        // First look: anything printed before this panel opened counts as
        // new only if it's later than the last print this document recorded.
        var saved = savedPrintTime();
        lastPrintSeen = saved || ms;
        if (!saved) savePrintTime(ms);
        if (!(saved && ms > saved)) { showWatch(ms); return; }
      }
      if (ms <= lastPrintSeen) return;
      lastPrintSeen = ms;
      savePrintTime(ms);
      showWatch(ms);
      onPrinted(ms);
    }).catch(function () { /* checking is best-effort; the button still works */ });
  }

  function showWatch(ms) {
    $('print-watch').textContent = ms ? 'Last printed ' + fmtDate(ms) + '.' : 'Not printed yet.';
  }

  function onPrinted(ms) {
    queue(function () { return readState().then(apply); }).then(function () {
      var nums = keys(st.filled);
      if (!nums.length) return;
      var when = timeOnly(ms);
      if (autoMarkOn()) {
        queue(printedNow).then(function (done) {
          return refresh().then(function () {
            say('Word printed this sheet' + (when ? ' at ' + when : '') + ', so ' + labelWord(done) + (done.length === 1 ? ' is' : ' are') + ' now marked used. Click Undo if they didn’t come out right.');
          });
        });
      } else {
        ask('Word printed this sheet' + (when ? ' at ' + when : '') + '. Mark ' + labelWord(nums) + ' used?', 'Mark used', function () {
          queue(printedNow).then(function (done) { return refresh().then(function () { say('Marked ' + labelWord(done) + ' used.'); }); });
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

  // ---------- shipping labels ----------
  var FROM_KEY = 'lsk.fromAddress';
  var DPI = 300;
  var ship = { doc: null, page: 1, pages: 1, src: null, crop: null, turn: 0 };

  function nextOpen() { var o = openList(); return o.length ? o[0] : 0; }

  // Usable area inside one label, in inches: inside the side padding, below
  // the top space, and clear of the printer's no-print strip at the bottom.
  function contentBox(f) {
    var L = S.layout(f), minRow = Infinity;
    L.rows.forEach(function (r) { if (r.label) minRow = Math.min(minRow, r.h / 1440); });
    var top = f.align === 'top' ? f.topSpace : 0.3;
    return { w: f.label.width - f.padding.left - f.padding.right, h: Math.max(0.5, minRow - top - 0.3), top: top, rowH: minRow };
  }

  var pdfLibPromise = null;
  function loadPdfLib() {
    if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
    if (pdfLibPromise) return pdfLibPromise;
    pdfLibPromise = new Promise(function (resolve, reject) {
      var s = document.createElement('script');
      s.src = 'vendor/pdfjs/pdf.js';
      s.onload = function () {
        if (!window.pdfjsLib) { pdfLibPromise = null; reject(userError('The PDF reader didn’t load. Check your internet connection and try again.')); return; }
        window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdfjs/pdf.worker.js';
        resolve(window.pdfjsLib);
      };
      s.onerror = function () { pdfLibPromise = null; reject(userError('The PDF reader didn’t load. Check your internet connection and try again.')); };
      document.head.appendChild(s);
    });
    return pdfLibPromise;
  }

  function fileError(e) {
    say((e && e.userMessage) || 'That file couldn’t be opened. Try a PDF, PNG or JPEG.', 'warn');
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
        img.onerror = function () { fileError(userError('That picture couldn’t be opened. Try a PDF, PNG or JPEG.')); };
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
    say(ship.pages > 1 ? 'This PDF has ' + ship.pages + ' pages. Use the arrows to pick the page with the label.' : 'Drag across the label if the page has extra instructions on it.');
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

  // The cropped label, turned to fit the label shape best, at up to ~2400 px.
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
      var ci = S.cellOf(f, n);
      var body = ctx.document.body.tables.getFirst().getCell(ci.row, ci.cell).body;
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

  function placeAddress(n, fromText, toText) {
    var f = st.fmt, topPt = f.align === 'top' ? f.topSpace * 72 : 0;
    var lines = [];
    var fromLines = fromText ? fromText.split(/\r?\n/) : [];
    var toLines = toText.split(/\r?\n/);
    if (fromLines.length) {
      lines.push({ t: 'FROM', size: 8, bold: true, before: topPt });
      fromLines.forEach(function (l) { lines.push({ t: l, size: 11 }); });
    }
    lines.push({ t: 'SHIP TO', size: 9, bold: true, before: fromLines.length ? 90 : topPt + 90, indent: 36 });
    toLines.forEach(function (l, i) { lines.push({ t: l, size: 16, bold: i === 0, indent: 36 }); });
    return Word.run(function (ctx) {
      var ci = S.cellOf(f, n);
      var body = ctx.document.body.tables.getFirst().getCell(ci.row, ci.cell).body;
      body.clear();
      var first = body.paragraphs.getFirst();
      lines.forEach(function (ln, i) {
        var p;
        if (i === 0) { p = first; p.insertText(ln.t, 'Replace'); }
        else p = body.insertParagraph(ln.t, 'End');
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

  function loadFrom() {
    var v = '';
    try { v = localStorage.getItem(FROM_KEY) || ''; } catch (e) {}
    if (!v) { try { v = Office.context.document.settings.get(FROM_KEY) || ''; } catch (e) {} }
    return v;
  }
  function saveFrom(v) {
    try { localStorage.setItem(FROM_KEY, v); } catch (e) {}
    try { Office.context.document.settings.set(FROM_KEY, v); Office.context.document.settings.saveAsync(function () {}); } catch (e) {}
  }

  function renderShip() {
    var n = nextOpen();
    var side = st.fmt.across === 2 && st.fmt.down === 1 ? (n === 1 ? ' (left)' : ' (right)') : '';
    $('ship-target').textContent = n
      ? 'Next open: label ' + n + side + '. Print the test page once to see which half of your sheet that is.'
      : 'There are no open labels on this sheet. When you load a fresh sheet, click Start a new sheet.';
    $('ship-place').disabled = !n;
    $('addr-place').disabled = !n;
  }

  function wireShip() {
    var input = $('ship-file');
    input.addEventListener('change', function () { handleFile(input.files && input.files[0]); input.value = ''; });
    $('ship-file-label').addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
    var card = $('step-ship');
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
    function endDrag(e) {
      if (!drag) return;
      drag.b = frac(e);
      var r = rectFrom(drag.a, drag.b);
      drag = null;
      if (r.w > ship.src.width * 0.03 && r.h > ship.src.height * 0.03) ship.crop = { x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.w), h: Math.round(r.h) };
      drawCrop();
    }
    wrap.addEventListener('pointerup', endDrag);
    wrap.addEventListener('pointercancel', function () { drag = null; drawCrop(); });

    $('ship-place').addEventListener('click', function () {
      if (!ship.src) { say('Choose a PDF or picture first.', 'warn'); return; }
      var n = nextOpen();
      if (!n) { say('There are no open labels on this sheet.', 'warn'); return; }
      var out = output();
      var b64 = out.canvas.toDataURL('image/png').split(',')[1];
      queue(function () { return placeImage(n, b64, out.wIn * 72, out.hIn * 72); }).then(refresh).then(function () {
        say('Placed on label ' + n + '. Print with ⌘P, then click I printed these.');
      });
    });

    var from = $('from-addr');
    from.value = loadFrom();
    from.addEventListener('input', function () { saveFrom(from.value); });
    $('addr-place').addEventListener('click', function () {
      var to = $('to-addr').value.replace(/\s+$/, '');
      if (!to.trim()) { say('Type the ship-to address first.', 'warn'); $('to-addr').focus(); return; }
      var n = nextOpen();
      if (!n) { say('There are no open labels on this sheet.', 'warn'); return; }
      var fromText = from.value.replace(/\s+$/, '');
      saveFrom(fromText);
      queue(function () { return placeAddress(n, fromText, to); }).then(refresh).then(function () {
        $('to-addr').value = '';
        say('Address label placed on label ' + n + '. You can edit it in the document.');
      });
    });
  }

  // ---------- start ----------
  function demo() {
    // Preview outside Word (?demo): shows the panel with sample data.
    var shipDemo = /demo=ship/.test(location.search);
    var f = shipDemo ? F.formats.filter(function (x) { return x.shipping; })[0] : F.formats[0];
    st = { kind: 'sheet', fmt: f, used: {}, filled: {}, texts: {}, here: shipDemo ? 1 : 8 };
    if (!shipDemo) {
      for (var n = 1; n <= 7; n++) st.used[n] = true;
      [8, 9, 10].forEach(function (n) { st.filled[n] = true; st.texts[n] = 'Sample Name\n123 Main St\nBoston, MA 02101'; });
    }
    window.Office = window.Office || {};
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
      try { Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged); } catch (e) {}
    });
  }
  start();
})();
