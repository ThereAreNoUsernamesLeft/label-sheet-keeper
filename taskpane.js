/* Label Sheet Keeper — Word side panel. */
(function () {
  'use strict';

  var F = window.LSK_FORMATS;
  var S = window.LSK_SHEET;
  var USED_TAG = 'lsk-used';
  var HISTORY_KEY = 'lskHistory';
  var $ = function (id) { return document.getElementById(id); };

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

  function createSheet(id) {
    var f = F.getFormat(id);
    if (!f) return Promise.reject(userError('Pick a label size first.'));
    return S.buildDocx(f).then(function (b64) {
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

    $('confirm-yes').addEventListener('click', function () { var a = pendingConfirm; closeAsk(); if (a) a(); });
    $('confirm-no').addEventListener('click', closeAsk);
  }

  // ---------- start ----------
  function demo() {
    // Preview outside Word (?demo): shows the panel with sample data.
    var f = F.formats[0];
    st = { kind: 'sheet', fmt: f, used: {}, filled: {}, texts: {}, here: 8 };
    for (var n = 1; n <= 7; n++) st.used[n] = true;
    [8, 9, 10].forEach(function (n) { st.filled[n] = true; st.texts[n] = 'Sample Name\n123 Main St\nBoston, MA 02101'; });
    window.Office = window.Office || {};
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
      try { Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged); } catch (e) {}
    });
  }
  start();
})();
