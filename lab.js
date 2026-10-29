/* Temporary checks of how Word for Mac handles label content. Loaded only by
   triple-clicking the version number in "About this panel". Will be removed. */
(function () {
  'use strict';
  var S = window.LSK_SHEET, out = [];
  var st = window.LSK_TEST && window.LSK_TEST.st();
  var box = document.getElementById('lsk-lab') || document.createElement('pre');
  box.id = 'lsk-lab';
  box.style.cssText = 'position:fixed;left:4px;right:4px;top:4px;bottom:4px;z-index:9999;margin:0;padding:6px;overflow:auto;background:#fff;color:#000;font:10px/1.25 Menlo,monospace;white-space:pre-wrap;border:2px solid #c00';
  box.onclick = function () { box.remove(); };
  document.body.appendChild(box);
  function log(s) { out.push(s); box.textContent = out.join('\n'); }
  if (!st || !st.fmt) { log('No sheet.'); return; }
  var f = st.fmt;
  function cell(ctx, n) { var c = S.cellOf(f, n); return ctx.document.body.tables.getFirst().getCell(c.row, c.cell).body; }
  function shape(xml) {
    var m = xml.match(/<w:body>([\s\S]*)<\/w:body>/), b = m ? m[1] : '';
    var cnt = function (re) { return (b.match(re) || []).length; };
    return 'len ' + xml.length + ' tbl ' + cnt(/<w:tbl>/g) + ' tr ' + cnt(/<w:tr[ >]/g) + ' tc ' + cnt(/<w:tc>/g) + ' p ' + cnt(/<w:p[ >]/g) +
      ' first ' + (b.replace(/^\s+/, '').slice(0, 12));
  }
  function pkg(paras) {
    var body = paras.map(function (p, i) { return S.paraXml(f, p, i === 0); }).join('');
    return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
      '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>' +
      '</pkg:xmlData></pkg:part>' +
      '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + body + '</w:body></w:document>' +
      '</pkg:xmlData></pkg:part></pkg:package>';
  }
  var P3 = [{ t: 'AAA small', size: 6, b: true }, { t: 'BBB plain', size: 10 }, { t: 'CCC indent', size: 12, indent: 36 }];
  function report(ctx, n, tag) {
    var ps = cell(ctx, n).paragraphs;
    ps.load('items/text,items/font/size,items/font/bold,items/leftIndent,items/tableNestingLevel,items/spaceBefore');
    return ctx.sync().then(function () {
      log(tag + ' L' + n + ': ' + ps.items.length + ' paras');
      ps.items.forEach(function (p) {
        log('   [' + JSON.stringify(p.text) + ' sz' + p.font.size + (p.font.bold ? ' B' : '') + ' ind' + p.leftIndent + ' bef' + p.spaceBefore + ' lvl' + p.tableNestingLevel + ']');
      });
    });
  }
  function step(name, fn) {
    return function () {
      return Word.run(fn).catch(function (e) { log(name + ' ERROR ' + (e && (e.code || '') + ' ' + (e.message || e))); });
    };
  }
  var src = Number(window.LSK_LAB_SRC || 5);
  log('Lab ' + new Date().toLocaleTimeString() + ' fmt ' + f.id + ' src L' + src + ' ua ' + navigator.userAgent.slice(0, 60));
  var PNG = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAICAIAAAB/FOjAAAAAFUlEQVR4nGNgYGD4z0AEQCRDg8YGAJ4GAP8gOZBRAAAAAElFTkSuQmCC';
  Promise.resolve()
    .then(step('E7', function (ctx) {
      var b = cell(ctx, 30); b.clear();
      var p0 = b.paragraphs.getFirst();
      var a = p0.insertParagraph('', 'Before');
      var img = a.insertInlinePictureFromBase64(PNG, 'Start'); img.width = 40; img.height = 20;
      p0.insertText('After picture', 'Start');
      var pics = b.inlinePictures; pics.load('items');
      return ctx.sync().then(function () { log('E7 pictures in L30: ' + pics.items.length); return report(ctx, 30, 'E7 picture paragraph'); });
    }))
    .then(function () { log('done. Click to close.'); });
})();
