/*
  Builds a Word document (.docx) for a label sheet size from formats.js.
  The page margins and a borderless table with exact row heights put every
  label cell exactly where the physical label sits on the sheet.
  Also maps label numbers (1, 2, 3... across each row) to table cells.
*/
(function (root) {
  'use strict';

  var TW = 1440;            // twips per inch
  var TRAIL = 60;           // room kept for Word's required last paragraph (twips)
  var PROP = 'LabelSheetFormat';

  function tw(inches) { return Math.round(inches * TW); }
  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  // Columns and rows of the Word table, including thin spacer columns/rows
  // where the sheet has gaps between labels.
  function layout(f) {
    var gapX = Math.max(0, f.pitch.across - f.label.width);
    var gapY = Math.max(0, f.pitch.down - f.label.height);
    var hasGapX = f.across > 1 && tw(gapX) > 0;
    var hasGapY = f.down > 1 && tw(gapY) > 0;
    var cols = [], rows = [], i;
    for (i = 0; i < f.across; i++) {
      cols.push({ label: true, w: tw(f.label.width) });
      if (hasGapX && i < f.across - 1) cols.push({ label: false, w: tw(gapX) });
    }
    for (i = 0; i < f.down; i++) {
      rows.push({ label: true, h: tw(f.label.height) });
      if (hasGapY && i < f.down - 1) rows.push({ label: false, h: tw(gapY) });
    }
    var pageW = tw(f.page.width), pageH = tw(f.page.height);
    var left = tw(f.firstLabel.left), top = tw(f.firstLabel.top);
    var tableW = cols.reduce(function (a, c) { return a + c.w; }, 0);
    var tableH = rows.reduce(function (a, r) { return a + r.h; }, 0);
    // Word always needs one paragraph after a table. When labels run to the
    // bottom edge, trim the last row a hair so that paragraph fits on page 1.
    var over = top + tableH + TRAIL - pageH;
    if (over > 0) { rows[rows.length - 1].h -= over; tableH -= over; }
    return {
      cols: cols, rows: rows, hasGapX: hasGapX, hasGapY: hasGapY,
      pageW: pageW, pageH: pageH, left: left, top: top,
      right: Math.max(0, pageW - left - tableW),
      bottom: Math.max(0, pageH - top - tableH - TRAIL),
      tableW: tableW
    };
  }

  function count(f) { return f.across * f.down; }

  // Label number (1-based, across each row from the top) -> table row/cell index
  function cellOf(f, n) {
    var L = layout(f);
    var r = Math.floor((n - 1) / f.across), c = (n - 1) % f.across;
    return { row: L.hasGapY ? r * 2 : r, cell: L.hasGapX ? c * 2 : c };
  }
  // Table row/cell index -> label number, or 0 for a spacer cell
  function labelAt(f, row, cell) {
    var L = layout(f);
    if (L.hasGapY) { if (row % 2) return 0; row = row / 2; }
    if (L.hasGapX) { if (cell % 2) return 0; cell = cell / 2; }
    if (row >= f.down || cell >= f.across) return 0;
    return row * f.across + cell + 1;
  }

  function documentXml(f, opts) {
    opts = opts || {};
    var L = layout(f);
    var pad = f.padding;
    var vAlign = f.align === 'top' ? 'top' : f.align === 'bottom' ? 'bottom' : 'center';
    var x = [];
    x.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    x.push('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>');
    x.push('<w:tbl><w:tblPr>');
    x.push('<w:tblW w:w="' + L.tableW + '" w:type="dxa"/>');
    x.push('<w:tblInd w:w="0" w:type="dxa"/>');
    var b = opts.outline ? 'w:val="single" w:sz="4" w:space="0" w:color="999999"' : 'w:val="nil"';
    x.push('<w:tblBorders><w:top ' + b + '/><w:left ' + b + '/><w:bottom ' + b + '/><w:right ' + b + '/><w:insideH ' + b + '/><w:insideV ' + b + '/></w:tblBorders>');
    x.push('<w:tblLayout w:type="fixed"/>');
    x.push('<w:tblCellMar><w:top w:w="' + tw(pad.top) + '" w:type="dxa"/><w:left w:w="' + tw(pad.left) + '" w:type="dxa"/><w:bottom w:w="' + tw(pad.bottom) + '" w:type="dxa"/><w:right w:w="' + tw(pad.right) + '" w:type="dxa"/></w:tblCellMar>');
    x.push('</w:tblPr><w:tblGrid>');
    L.cols.forEach(function (c) { x.push('<w:gridCol w:w="' + c.w + '"/>'); });
    x.push('</w:tblGrid>');
    var n = 0;
    L.rows.forEach(function (r) {
      x.push('<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="' + r.h + '" w:hRule="exact"/></w:trPr>');
      L.cols.forEach(function (c) {
        var isLabel = r.label && c.label;
        x.push('<w:tc><w:tcPr><w:tcW w:w="' + c.w + '" w:type="dxa"/>');
        if (!isLabel) x.push('<w:tcMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>');
        x.push('<w:vAlign w:val="' + (isLabel ? vAlign : 'top') + '"/></w:tcPr>');
        if (isLabel) {
          n++;
          if (opts.numbers) x.push('<w:p><w:r><w:t>Label ' + n + '</w:t></w:r></w:p>');
          else x.push('<w:p/>');
        } else {
          x.push('<w:p><w:pPr><w:spacing w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>');
        }
        x.push('</w:tc>');
      });
      x.push('</w:tr>');
    });
    x.push('</w:tbl>');
    // Required last paragraph: hidden, 1pt, so it never makes a second page.
    x.push('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:vanish/><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>');
    x.push('<w:sectPr><w:pgSz w:w="' + L.pageW + '" w:h="' + L.pageH + '"/>');
    x.push('<w:pgMar w:top="' + L.top + '" w:right="' + L.right + '" w:bottom="' + L.bottom + '" w:left="' + L.left + '" w:header="0" w:footer="0" w:gutter="0"/>');
    x.push('<w:cols w:space="720"/></w:sectPr>');
    x.push('</w:body></w:document>');
    return x.join('');
  }

  function stylesXml(f) {
    var font = esc(f.font.name), sz = Math.round(f.font.size * 2);
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="' + font + '" w:hAnsi="' + font + '" w:eastAsia="' + font + '" w:cs="' + font + '"/>' +
      '<w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
      '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/>' +
      '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
      '</w:styles>';
  }

  var SETTINGS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' +
    '</w:settings>';

  var CONTENT_TYPES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>' +
    '</Types>';

  var ROOT_RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>' +
    '</Relationships>';

  var DOC_RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>' +
    '</Relationships>';

  function coreXml(f) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:title>' + esc('Label sheet – ' + f.name) + '</dc:title><dc:creator>Label Sheet Keeper</dc:creator></cp:coreProperties>';
  }

  function customXml(f) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="' + PROP + '"><vt:lpwstr>' + esc(f.id) + '</vt:lpwstr></property>' +
      '</Properties>';
  }

  // Returns a Promise of the .docx as base64 (type 'base64') or a Node buffer (type 'nodebuffer').
  function buildDocx(f, JSZipLib, opts) {
    opts = opts || {};
    var Z = JSZipLib || root.JSZip;
    if (!Z) return Promise.reject(new Error('The file builder did not load. Check your internet connection and reopen the panel.'));
    var zip = new Z();
    zip.file('[Content_Types].xml', CONTENT_TYPES);
    zip.file('_rels/.rels', ROOT_RELS);
    zip.file('docProps/core.xml', coreXml(f));
    zip.file('docProps/custom.xml', customXml(f));
    zip.file('word/document.xml', documentXml(f, opts));
    zip.file('word/_rels/document.xml.rels', DOC_RELS);
    zip.file('word/styles.xml', stylesXml(f));
    zip.file('word/settings.xml', SETTINGS);
    return zip.generateAsync({ type: opts.type || 'base64', compression: 'DEFLATE' });
  }

  var api = { PROP: PROP, layout: layout, count: count, cellOf: cellOf, labelAt: labelAt, buildDocx: buildDocx, documentXml: documentXml };
  root.LSK_SHEET = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
