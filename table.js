/*
  Label Sheet Keeper — the table window.
  One row per label, in order: row 1 goes on the first open label, and so on.
  The first column shows the label each row goes on; a "New sheet" divider
  marks where rows run past the open labels and onto new sheets. An empty
  row leaves its label open.
*/
(function () {
  'use strict';

  var T = window.LSK_TOOLS;
  var COLS = ['name', 'company', 'street', 'apt', 'city', 'state', 'zip'];
  var DRAFT = 'lsk.tableDraft', INFO = 'lsk.tableInfo', BOOK = 'lsk.addressBook';
  var rowsEl = document.getElementById('rows');

  function $(id) { return document.getElementById(id); }
  function store(k, v) { try { if (v === undefined) return localStorage.getItem(k); localStorage.setItem(k, v); } catch (e) { /* blocked */ } return null; }
  function json(k) { try { return JSON.parse(store(k) || 'null'); } catch (e) { return null; } }
  function plural(n, w, s) { return n + ' ' + w + (n === 1 ? '' : (s || 's')); }
  function msg(text, warn) { $('msg').textContent = text || ''; $('msg').className = 'msg' + (warn ? ' warn' : ''); }

  // ---------- rows ----------
  function dataRows() { return Array.prototype.filter.call(rowsEl.children, function (tr) { return tr.classList.contains('row'); }); }
  function addRow(values) {
    var tr = document.createElement('tr');
    tr.className = 'row';
    var num = document.createElement('td');
    tr.appendChild(num);
    COLS.forEach(function (c, i) {
      var td = document.createElement('td'), inp = document.createElement('input');
      inp.type = 'text'; inp.value = (values && values[i]) || ''; inp.setAttribute('aria-label', c);
      inp.autocomplete = 'off'; inp.spellcheck = false;
      td.appendChild(inp); tr.appendChild(td);
    });
    var xd = document.createElement('td'), x = document.createElement('button');
    x.type = 'button'; x.className = 'x'; x.textContent = '×'; x.setAttribute('aria-label', 'Delete row');
    x.addEventListener('click', function () { tr.remove(); if (!dataRows().length) addRows(1); update(); });
    xd.appendChild(x); tr.appendChild(xd);
    rowsEl.appendChild(tr);
    return tr;
  }
  function addRows(n) { for (var i = 0; i < n; i++) addRow(); }
  function rowValues(tr) { return Array.prototype.map.call(tr.querySelectorAll('input'), function (i) { return i.value.trim(); }); }
  // All rows in order, without the empty ones at the end.
  function readRows() {
    var rows = dataRows().map(rowValues);
    while (rows.length && !rows[rows.length - 1].some(Boolean)) rows.pop();
    return rows;
  }
  function filledRows() { return readRows().filter(function (r) { return r.some(Boolean); }); }
  function setRows(rows) {
    var info = json(INFO) || {};
    rowsEl.innerHTML = '';
    rows.forEach(function (r) { addRow(r); });
    // Start with one row per open label (or 10 for a full sheet), plus a few spare.
    var want = rows.length ? Math.max(rows.length + 3, info.open || 0) : (info.open || 10);
    addRows(want - rows.length);
    update();
  }
  // Put new rows after the last row with something in it.
  function appendRows(rows) { setRows(readRows().concat(rows)); }

  function update() {
    var info = json(INFO) || {}, open = info.open, per = info.per || 30, labels = info.labels || [];
    var known = typeof open === 'number';
    Array.prototype.slice.call(rowsEl.querySelectorAll('tr.divider')).forEach(function (d) { d.remove(); });
    var filled = 0, last = -1;
    dataRows().forEach(function (tr, i) {
      if (rowValues(tr).some(Boolean)) { filled++; last = i; }
      var extra = known ? i - open : -1;
      tr.firstChild.textContent = !known ? String(i + 1) : extra < 0 ? String(labels[i] || i + 1) : String(extra % per + 1);
      tr.classList.toggle('spill', extra >= 0);
      if (extra >= 0 && extra % per === 0) {
        var d = document.createElement('tr'), td = document.createElement('td');
        d.className = 'divider';
        td.colSpan = COLS.length + 2;
        td.textContent = 'New sheet ' + (extra / per + 1) + ' · rows below go on a new sheet that opens in its own window';
        d.appendChild(td);
        rowsEl.insertBefore(d, tr);
      }
    });
    var extraSlots = known ? Math.max(0, last + 1 - open) : 0;
    $('count').textContent = plural(filled, 'address', 'es') + (known ? ' · ' + plural(open, 'open label') + ' on your sheet' : '') +
      (extraSlots ? ' · ' + (extraSlots === 1 ? '1 row goes' : extraSlots + ' rows go') + ' on ' + plural(Math.ceil(extraSlots / per), 'new sheet') : '');
    $('fill').disabled = !filled;
    store(DRAFT, JSON.stringify(readRows()));
  }

  // ---------- turning other data into rows ----------
  function fieldsToRow(a) {
    if (a.extra && a.extra.length && !a.street && !a.city) {
      var f = T.linesToFields([a.name].concat(a.extra));
      f.company = a.company || f.company;
      a = f;
    }
    return COLS.map(function (c) { return c === 'zip' ? T.fixZip(a.zip, a.state) : (a[c] || ''); });
  }
  function rowsFromGrid(grid) {
    // A header row is recognized by its names; otherwise columns go in order.
    if (T.headerMap(grid[0])) return T.rowsToAddresses(grid).map(fieldsToRow);
    return grid.map(function (r) { return COLS.map(function (c, i) { return String(r[i] === undefined ? '' : r[i]).trim(); }); });
  }

  // Paste: a block of cells fills the grid from the cell you're in.
  rowsEl.addEventListener('paste', function (e) {
    var target = e.target;
    if (!target || target.tagName !== 'INPUT') return;
    var text = e.clipboardData && e.clipboardData.getData('text/plain');
    if (!text || (text.indexOf('\t') < 0 && text.indexOf('\n') < 0)) return;   // a single value: normal paste
    e.preventDefault();
    var grid = T.parseDelimited(text.replace(/\r?\n$/, ''), text.indexOf('\t') >= 0 ? '\t' : ',');
    if (T.headerMap(grid[0])) { appendRows(rowsFromGrid(grid)); msg('Pasted ' + plural(grid.length - 1, 'row') + ' and matched the columns by their headings.'); return; }
    var tr = target.closest('tr'), startCol = Array.prototype.indexOf.call(tr.querySelectorAll('input'), target);
    var rows = dataRows(), rowIndex = rows.indexOf(tr);
    grid.forEach(function (cells, r) {
      var row = rows[rowIndex + r] || addRow();
      var inputs = row.querySelectorAll('input');
      cells.forEach(function (v, c) { if (inputs[startCol + c]) inputs[startCol + c].value = String(v).trim(); });
    });
    if (dataRows().length - rowIndex - grid.length < 3) addRows(5);
    update();
    msg('Pasted ' + plural(grid.length, 'row') + '.');
  });

  // Enter moves down a row, like a spreadsheet.
  rowsEl.addEventListener('keydown', function (e) {
    if (e.key !== 'Enter' || e.target.tagName !== 'INPUT') return;
    e.preventDefault();
    var tr = e.target.closest('tr'), col = Array.prototype.indexOf.call(tr.querySelectorAll('input'), e.target);
    var rows = dataRows(), next = rows[rows.indexOf(tr) + 1] || addRow();
    next.querySelectorAll('input')[col].focus();
    update();
  });
  rowsEl.addEventListener('input', update);

  // Open a file
  $('file').addEventListener('change', function () {
    var file = $('file').files && $('file').files[0]; $('file').value = '';
    if (!file) return;
    var name = (file.name || '').toLowerCase(), reader = new FileReader();
    var done = function (rows) {
      rows = rows.filter(function (r) { return r.some(Boolean); });
      if (!rows.length) { msg('No addresses were found in that file.', true); return; }
      appendRows(rows);
      msg('Added ' + plural(rows.length, 'row') + ' from ' + file.name + '. Check the columns before filling.');
    };
    if (/\.xlsx$/.test(name)) {
      reader.onload = function () { T.readXlsx(reader.result).then(function (grid) { done(rowsFromGrid(grid)); }).catch(function () { msg('That Excel file couldn’t be read. Try saving it as CSV.', true); }); };
      reader.readAsArrayBuffer(file);
    } else {
      reader.onload = function () {
        var text = String(reader.result || '');
        if (/\.vcf$/.test(name) || /BEGIN:VCARD/i.test(text)) done(T.parseVCard(text).map(function (c) { return fieldsToRow(c.fields); }));
        else done(rowsFromGrid(T.parseDelimited(text)));
      };
      reader.readAsText(file);
    }
  });
  $('file-label').addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('file').click(); } });

  // Saved addresses
  function fillBookPicker() {
    var b = json(BOOK) || [], sel = $('book');
    sel.innerHTML = '';
    var o = document.createElement('option'); o.value = ''; o.textContent = b.length ? 'Add a saved address…' : 'No saved addresses';
    sel.appendChild(o);
    if (b.length) { var all = document.createElement('option'); all.value = '*'; all.textContent = 'Add all ' + b.length; sel.appendChild(all); }
    b.forEach(function (e) { var op = document.createElement('option'); op.value = e.id; op.textContent = e.name; sel.appendChild(op); });
    sel.disabled = !b.length;
  }
  $('book').addEventListener('change', function () {
    var b = json(BOOK) || [], v = $('book').value;
    var pick = v === '*' ? b : b.filter(function (e) { return e.id === v; });
    if (pick.length) { appendRows(pick.map(function (e) { return fieldsToRow(T.linesToFields(e.lines)); })); msg('Added ' + plural(pick.length, 'saved address', 'es') + '.'); }
    $('book').value = '';
  });

  $('more').addEventListener('click', function () { addRows(10); update(); });
  $('clear').addEventListener('click', function () {
    if ($('clear').dataset.sure !== '1') { $('clear').dataset.sure = '1'; $('clear').textContent = 'Click again to clear'; setTimeout(function () { $('clear').dataset.sure = ''; $('clear').textContent = 'Clear'; }, 3000); return; }
    $('clear').dataset.sure = ''; $('clear').textContent = 'Clear';
    setRows([]); msg('Cleared.');
  });

  function send(message) {
    var ui = window.Office && Office.context && Office.context.ui;
    if (ui && typeof ui.messageParent === 'function') ui.messageParent(JSON.stringify(message));
    else msg('This window works with the Label Sheet Keeper panel in Word.', true);
  }
  $('fill').addEventListener('click', function () {
    var rows = readRows();
    if (!filledRows().length) { msg('Add at least one address first.', true); return; }
    store(DRAFT, '[]');
    send({ type: 'fill', rows: rows });
  });
  $('cancel').addEventListener('click', function () { send({ type: 'close' }); });

  // ---------- start ----------
  function start() {
    var info = json(INFO) || {};
    $('info').textContent = (info.name ? info.name + '. ' : '') + 'One row per label; the first column shows which label. Type addresses, paste cells from Excel or Numbers, or open a file.';
    var draft = json(DRAFT);
    setRows(Array.isArray(draft) ? draft : []);
    fillBookPicker();
    var first = rowsEl.querySelector('input'); if (first) first.focus();
  }
  if (window.Office && Office.onReady) Office.onReady(function () { start(); });
  else start();
  window.LSK_TABLE = { readRows: filledRows, rows: readRows, setRows: setRows };
})();
