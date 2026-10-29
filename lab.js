/* Temporary test helper, loaded only by triple-clicking "About this panel".
   On a shipping test sheet it loads sample data; on any other test sheet it
   removes what testing left in this Mac's panel storage. Will be removed. */
(function () {
  'use strict';
  var out = [], st = window.LSK_TEST && window.LSK_TEST.st();
  var box = document.createElement('pre');
  box.style.cssText = 'position:fixed;left:4px;right:4px;top:4px;z-index:9999;margin:0;padding:6px;background:#fff;color:#000;font:10px/1.25 Menlo,monospace;white-space:pre-wrap;border:2px solid #c00';
  box.onclick = function () { box.remove(); };
  document.body.appendChild(box);
  function log(s) { out.push(s); box.textContent = out.join('\n') + '\n(click to close)'; }
  function json(k) { try { return JSON.parse(localStorage.getItem(k) || 'null'); } catch (e) { return null; } }
  var myId = Office.context.document.settings.get('lskSheetId');
  var ids = json('lsk.labIds') || [];
  if (myId && ids.indexOf(myId) < 0) { ids.push(myId); localStorage.setItem('lsk.labIds', JSON.stringify(ids)); }
  if (!st || !st.fmt) { log('No sheet.'); return; }

  if (st.fmt.shipping) {
    document.getElementById('to-addr').value = 'Returns Center\n100 Warehouse Rd\nSeattle, WA 98109';
    var c = document.createElement('canvas'); c.width = 800; c.height = 1200;
    var x = c.getContext('2d');
    x.fillStyle = '#fff'; x.fillRect(0, 0, 800, 1200);
    x.strokeStyle = '#000'; x.lineWidth = 8; x.strokeRect(80, 80, 640, 1040);
    x.fillStyle = '#000'; x.font = 'bold 48px Arial'; x.fillText('TEST LABEL', 260, 240);
    x.font = '32px Arial'; x.fillText('Not a real shipping label', 220, 320);
    for (var i = 0; i < 12; i++) x.fillRect(160 + i * 40, 900, i % 3 ? 14 : 24, 140);
    c.toBlob(function (blob) {
      window.LSK_TEST.handleFile(new File([blob], 'test-label.png', { type: 'image/png' }));
      log('Loaded a test label picture and a Ship to address. Test sheet ids: ' + ids.length);
    }, 'image/png');
    return;
  }

  // Clean up after testing.
  var reg = json('lsk.sheets') || {}, removed = [];
  ids.forEach(function (id) { if (reg[id]) { delete reg[id]; removed.push('sheet ' + id); } });
  localStorage.setItem('lsk.sheets', JSON.stringify(reg));
  ['lsk.last.' + st.fmt.id, 'lsk.last.shipping-2-sideways', 'lsk.design', 'lsk.tableDraft', 'lsk.tableInfo', 'lsk.labIds'].forEach(function (k) {
    if (localStorage.getItem(k) !== null) { localStorage.removeItem(k); removed.push(k); }
  });
  log('Removed: ' + (removed.join(', ') || 'nothing'));
  log('Sheets still listed: ' + Object.keys(reg).map(function (k) { return reg[k].file || '(unsaved)'; }).join(', '));
})();
