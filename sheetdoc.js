/*
  Builds Word documents for a label sheet size from formats.js.

  Page layout of a sheet (from the top of the page):
    top page margin        fixed per sheet: 0.32″ on sheets with a blank top
                           strip (room for the small feed arrow in the
                           header), otherwise 0
    spacer paragraph       locked, exact height; together with the margin it
                           puts the first row exactly where the labels start.
                           Changing its height moves every label up or down
                           without touching the page margins.
    label table            borderless, exact row heights; thin spacer rows
                           and columns where the sheet has gaps
    last paragraph         Word requires one after a table; kept tiny

  The left page margin is the first label's left edge; a table indent moves
  labels left or right later. Each sheet stores its size and its margins in
  document properties, so it keeps working even if its size is later removed
  from the list.
*/
(function (root) {
  'use strict';

  var TW = 1440;                 // twips per inch
  var TRAIL = 432;               // room below the labels for Word's last paragraph (0.3″)
  var STRIP_MARGIN = 0.32;       // top page margin on sheets with a blank top strip
  var PROP = 'LabelSheetFormat';
  var LAYOUT_PROP = 'LabelSheetLayout';
  var SPEC_PROP = 'LabelSheetSpec';
  var SPACER_TAG = 'lsk-spacer';
  var ARROW_TEXT = '▲ TOP · FEED THIS EDGE INTO THE PRINTER FIRST · PRINT SIDE UP';
  var TEST_ARROW_TEXT = '▲ TOP OF PAGE · THIS EDGE WENT INTO THE PRINTER FIRST';
  var ARROW_BEFORE_PT = 15;      // space above the arrow text; clears the 5 mm edge printers skip
  var ARROW_SIZE_PT = 6;

  function tw(inches) { return Math.round(inches * TW); }
  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function hasStrip(f) { return !!f && f.page.height >= f.page.width && f.firstLabel.top >= 0.4; }

  // The page margins a new sheet gets (inches). Older sheets used the first
  // label's position as the margins; pass those in opts.margins instead.
  function defaultMargins(f, nudge) {
    var left = f.firstLabel.left + ((nudge && nudge.x) || 0);
    return { top: hasStrip(f) ? STRIP_MARGIN : 0, left: Math.max(0, left) };
  }

  // Columns and rows of the label table, plus where it sits on the page.
  // opts: { nudge: {x, y} inches, margins: {top, left} inches (the sheet's
  // actual page margins, for rebuilding an existing sheet in place) }
  function layout(f, opts) {
    opts = opts || {};
    var nudge = opts.nudge || { x: 0, y: 0 };
    var m = opts.margins || defaultMargins(f, nudge);
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
    var wantTop = tw(f.firstLabel.top + (nudge.y || 0));
    var wantLeft = tw(f.firstLabel.left + (nudge.x || 0));
    var top = tw(m.top), left = tw(m.left);
    // The spacer makes up the difference between the margin and where the
    // labels should start. It can't be negative: labels can't move above the
    // top margin (on older sheets the margin is where the labels start, and
    // on some sizes the labels start at the very top of the page).
    var spacer = Math.max(0, wantTop - top);
    var clampedUp = wantTop < top;
    var indent = wantLeft - left;
    var tableW = cols.reduce(function (a, c) { return a + c.w; }, 0);
    var tableH = rows.reduce(function (a, r) { return a + r.h; }, 0);
    // Word needs one paragraph after the table. When labels run to the bottom
    // edge, trim the last row so that paragraph still fits on page 1 (that
    // strip is in the printer's no-print zone anyway).
    var over = top + spacer + tableH + TRAIL - pageH;
    if (over > 0) { rows[rows.length - 1].h -= over; tableH -= over; }
    return {
      cols: cols, rows: rows, hasGapX: hasGapX, hasGapY: hasGapY,
      pageW: pageW, pageH: pageH, top: top, left: left, spacer: spacer, indent: indent,
      right: Math.max(0, pageW - left - tableW), bottom: 0, tableW: tableW,
      clampedUp: clampedUp, margins: m
    };
  }

  function count(f) { return f.across * f.down; }

  // Label number (1-based, across each row from the top) <-> table row/cell.
  function cellOf(f, n) {
    var L = layout(f);
    var r = Math.floor((n - 1) / f.across), c = (n - 1) % f.across;
    return { row: L.hasGapY ? r * 2 : r, cell: L.hasGapX ? c * 2 : c };
  }
  function labelAt(f, row, cell) {
    var L = layout(f);
    if (L.hasGapY) { if (row % 2) return 0; row = row / 2; }
    if (L.hasGapX) { if (cell % 2) return 0; cell = cell / 2; }
    if (row >= f.down || cell >= f.across) return 0;
    return row * f.across + cell + 1;
  }
  function tableShape(f) { var L = layout(f); return { rows: L.rows.length, cols: L.cols.length }; }

  // One paragraph of label text. p: { t, size, b, i, align, before, indent, font, color }
  function paraXml(f, p, first) {
    var ppr = '';
    if (first && f.align === 'top' && p.before === undefined) ppr += '<w:pStyle w:val="LabelTop"/>';
    if (p.before !== undefined) ppr += '<w:spacing w:before="' + Math.round(p.before * 20) + '" w:after="0"/>';
    if (p.indent) ppr += '<w:ind w:left="' + Math.round(p.indent * 20) + '"/>';
    if (p.align && p.align !== 'left') ppr += '<w:jc w:val="' + (p.align === 'center' ? 'center' : p.align) + '"/>';
    var rpr = '';
    if (p.font) rpr += '<w:rFonts w:ascii="' + esc(p.font) + '" w:hAnsi="' + esc(p.font) + '" w:cs="' + esc(p.font) + '"/>';
    if (p.b) rpr += '<w:b/>';
    if (p.i) rpr += '<w:i/>';
    if (p.color) rpr += '<w:color w:val="' + esc(String(p.color).replace('#', '')) + '"/>';
    if (p.size) rpr += '<w:sz w:val="' + Math.round(p.size * 2) + '"/><w:szCs w:val="' + Math.round(p.size * 2) + '"/>';
    return '<w:p>' + (ppr ? '<w:pPr>' + ppr + '</w:pPr>' : '') +
      (p.t ? '<w:r>' + (rpr ? '<w:rPr>' + rpr + '</w:rPr>' : '') + '<w:t xml:space="preserve">' + esc(p.t) + '</w:t></w:r>' : '') + '</w:p>';
  }

  function spacerXml(L) {
    if (!L.spacer) return '';
    return '<w:sdt><w:sdtPr><w:alias w:val="Label position"/><w:tag w:val="' + SPACER_TAG + '"/><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent>' +
      '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="' + Math.max(1, L.spacer) + '" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>' +
      '</w:sdtContent></w:sdt>';
  }

  // The label table. opts: { outline, numbers, prefill: { n: [para] } }
  function tableXml(f, L, opts) {
    opts = opts || {};
    var pad = f.padding;
    var vAlign = f.align === 'top' ? 'top' : f.align === 'bottom' ? 'bottom' : 'center';
    var x = [];
    x.push('<w:tbl><w:tblPr>');
    x.push('<w:tblW w:w="' + L.tableW + '" w:type="dxa"/>');
    x.push('<w:tblInd w:w="' + L.indent + '" w:type="dxa"/>');
    // Dark, thin lines: easy to see against a light, and to find on a scan.
    var b = opts.outline ? 'w:val="single" w:sz="6" w:space="0" w:color="000000"' : 'w:val="nil"';
    x.push('<w:tblBorders><w:top ' + b + '/><w:left ' + b + '/><w:bottom ' + b + '/><w:right ' + b + '/><w:insideH ' + b + '/><w:insideV ' + b + '/></w:tblBorders>');
    x.push('<w:tblLayout w:type="fixed"/>');
    // Side padding only. Word for Mac can add top/bottom cell padding on top of
    // an exact row height, which makes rows drift down the page.
    x.push('<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="' + tw(pad.left) + '" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="' + tw(pad.right) + '" w:type="dxa"/></w:tblCellMar>');
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
          // On top-aligned sheets the first line uses the LabelTop style, which
          // starts it topSpace below the label's edge.
          var pStyle = f.align === 'top' ? '<w:pStyle w:val="LabelTop"/>' : '';
          var pre = opts.prefill && opts.prefill[n];
          if (opts.numbers) x.push('<w:p><w:pPr>' + pStyle + '<w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Label ' + n + '</w:t></w:r></w:p>');
          else if (pre && pre.length) pre.forEach(function (p, i) { x.push(paraXml(f, p, i === 0)); });
          else x.push(pStyle ? '<w:p><w:pPr>' + pStyle + '</w:pPr></w:p>' : '<w:p/>');
        } else {
          x.push('<w:p><w:pPr><w:spacing w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>');
        }
        x.push('</w:tc>');
      });
      x.push('</w:tr>');
    });
    x.push('</w:tbl>');
    return x.join('');
  }

  var NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

  function documentXml(f, opts) {
    opts = opts || {};
    var L = layout(f, opts);
    var x = [];
    x.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    x.push('<w:document ' + NS + '><w:body>');
    x.push(spacerXml(L));
    x.push(tableXml(f, L, opts));
    // Required last paragraph: hidden, 1pt, so it never makes a second page.
    x.push('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:vanish/><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>');
    var orient = L.pageW > L.pageH ? ' w:orient="landscape"' : '';
    x.push('<w:sectPr>');
    if (opts.parts) x.push('<w:headerReference w:type="default" r:id="rId3"/><w:footerReference w:type="default" r:id="rId4"/>');
    x.push('<w:pgSz w:w="' + L.pageW + '" w:h="' + L.pageH + '"' + orient + '/>');
    x.push('<w:pgMar w:top="' + L.top + '" w:right="' + L.right + '" w:bottom="' + L.bottom + '" w:left="' + L.left + '" w:header="0" w:footer="0" w:gutter="0"/>');
    x.push('<w:cols w:space="720"/></w:sectPr>');
    x.push('</w:body></w:document>');
    return x.join('');
  }

  // The header holds the feed arrow, or else one tiny empty paragraph. The
  // footer is always one tiny empty paragraph. With top or bottom margins of
  // 0, Word makes room for a header or footer even when there is none, which
  // would push full-height labels onto a second page.
  var TINY_P = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';
  function headerXml(text) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ' + NS + '>' + (text ?
      '<w:p><w:pPr><w:spacing w:before="' + ARROW_BEFORE_PT * 20 + '" w:after="0"/><w:jc w:val="center"/></w:pPr>' +
      '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:color w:val="555555"/><w:sz w:val="' + ARROW_SIZE_PT * 2 + '"/><w:szCs w:val="' + ARROW_SIZE_PT * 2 + '"/></w:rPr>' +
      '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>' : TINY_P) + '</w:hdr>';
  }
  function footerXml() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ' + NS + '>' + TINY_P + '</w:ftr>';
  }

  // A flat package for Word's insertOoxml: the spacer and the label table,
  // used to rebuild a sheet in place (repair, printer adjustment).
  function tablePackage(f, opts) {
    var L = layout(f, opts);
    var body = spacerXml(L) + tableXml(f, L, opts) + '<w:p/>';
    return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
      '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>' +
      '</pkg:xmlData></pkg:part>' +
      '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>' +
      '<w:document ' + NS + '><w:body>' + body + '</w:body></w:document>' +
      '</pkg:xmlData></pkg:part></pkg:package>';
  }

  function stylesXml(f) {
    var font = esc(f.font.name), sz = Math.round(f.font.size * 2);
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="' + font + '" w:hAnsi="' + font + '" w:eastAsia="' + font + '" w:cs="' + font + '"/>' +
      '<w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
      '<w:style w:type="paragraph" w:customStyle="1" w:styleId="LabelTop"><w:name w:val="Label first line"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
      '<w:pPr><w:spacing w:before="' + tw(f.topSpace || 0.4) + '"/></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/></w:style>' +
      '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/>' +
      '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
      '</w:styles>';
  }

  var SETTINGS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' +
    '</w:settings>';

  var WE_NS = 'http://schemas.microsoft.com/office/webextensions/webextension/2010/11';
  var WETP_NS = 'http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11';

  // The add-in reference Word wrote into a document where the panel was open
  // ({ id, version, store, storeType }). Copying it into a new sheet, with
  // Office.AutoShowTaskpaneWithDocument set, makes Word open the panel with it.
  function webExtensionXml(ref) {
    var g = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
      var r = Math.random() * 16 | 0; return (c === 'x' ? r : (r & 3 | 8)).toString(16);
    }).toUpperCase();
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<we:webextension xmlns:we="' + WE_NS + '" id="{' + g + '}">' +
      '<we:reference id="' + esc(ref.id) + '" version="' + esc(ref.version) + '" store="' + esc(ref.store) + '" storeType="' + esc(ref.storeType) + '"/>' +
      '<we:alternateReferences/>' +
      '<we:properties><we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/></we:properties>' +
      '<we:bindings/><we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>' +
      '</we:webextension>';
  }
  var TASKPANES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<wetp:taskpanes xmlns:wetp="' + WETP_NS + '">' +
    '<wetp:taskpane dockstate="right" visibility="1" width="350" row="0">' +
    '<wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>' +
    '</wetp:taskpane></wetp:taskpanes>';
  var TASKPANES_RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/>' +
    '</Relationships>';

  function contentTypes(we) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>' +
      (we ? '<Override PartName="/word/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>' +
        '<Override PartName="/word/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/>' : '') +
      '</Types>';
  }
  function rootRels(we) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>' +
      (we ? '<Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="word/webextensions/taskpanes.xml"/>' : '') +
      '</Relationships>';
  }
  function docRels() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
      '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' +
      '</Relationships>';
  }

  function coreXml(f, opts) {
    var title = (opts && opts.test ? 'Alignment test – ' : 'Label sheet – ') + f.name;
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:title>' + esc(title) + '</dc:title><dc:creator>Label Sheet Keeper</dc:creator></cp:coreProperties>';
  }

  // The measurements a sheet needs to work on its own.
  function specOf(f) {
    return {
      id: f.id, name: f.name, detail: f.detail, page: f.page, across: f.across, down: f.down,
      label: f.label, firstLabel: f.firstLabel, pitch: f.pitch, padding: f.padding, align: f.align,
      topSpace: f.topSpace, font: f.font, corner: f.corner, shipping: f.shipping
    };
  }
  function chunks(s, n) { var out = []; for (var i = 0; i < s.length; i += n) out.push(s.slice(i, i + n)); return out; }

  // A real sheet carries LabelSheetFormat so the panel recognizes it, plus
  // its margins and measurements. A test page gets a different name so the
  // panel doesn't track it as a sheet.
  function customXml(f, L, opts) {
    var props = [];
    if (opts && opts.test) props.push(['LabelSheetTest', f.id]);
    else {
      props.push([PROP, f.id]);
      props.push([LAYOUT_PROP, JSON.stringify({ v: 2, top: L.margins.top, left: L.margins.left })]);
      chunks(JSON.stringify(specOf(f)), 240).forEach(function (c, i) { props.push([SPEC_PROP + (i + 1), c]); });
    }
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      props.map(function (p, i) {
        return '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="' + (i + 2) + '" name="' + esc(p[0]) + '"><vt:lpwstr>' + esc(p[1]) + '</vt:lpwstr></property>';
      }).join('') +
      '</Properties>';
  }

  // Builds the .docx. Returns a Promise of base64 (default) or a Node buffer
  // (opts.type 'nodebuffer').
  // opts: { test, nudge, addinRef, feedArrow, prefill, outline, numbers, type }
  function buildDocx(f, JSZipLib, opts) {
    opts = opts || {};
    var Z = JSZipLib || root.JSZip;
    if (!Z) return Promise.reject(new Error('The file builder didn’t load. Close and reopen the panel.'));
    if (opts.test) { opts.outline = true; opts.numbers = true; }
    var L = layout(f, opts);
    var header = (opts.feedArrow || opts.test) && hasStrip(f) && L.top >= tw(STRIP_MARGIN) ? (opts.test ? TEST_ARROW_TEXT : ARROW_TEXT) : '';
    var docOpts = {}; Object.keys(opts).forEach(function (k) { docOpts[k] = opts[k]; });
    docOpts.parts = true;
    var we = opts.addinRef && opts.addinRef.id && !opts.test;
    var zip = new Z();
    zip.file('[Content_Types].xml', contentTypes(we));
    zip.file('_rels/.rels', rootRels(we));
    if (we) {
      zip.file('word/webextensions/taskpanes.xml', TASKPANES);
      zip.file('word/webextensions/_rels/taskpanes.xml.rels', TASKPANES_RELS);
      zip.file('word/webextensions/webextension1.xml', webExtensionXml(opts.addinRef));
    }
    zip.file('docProps/core.xml', coreXml(f, opts));
    zip.file('docProps/custom.xml', customXml(f, L, opts));
    zip.file('word/document.xml', documentXml(f, docOpts));
    zip.file('word/_rels/document.xml.rels', docRels());
    zip.file('word/styles.xml', stylesXml(f));
    zip.file('word/settings.xml', SETTINGS);
    zip.file('word/header1.xml', headerXml(header));
    zip.file('word/footer1.xml', footerXml());
    return zip.generateAsync({ type: opts.type || 'base64', compression: 'DEFLATE' });
  }

  // ---------- backup of the panel's settings ----------
  // A plain Word document that carries the settings (base64 JSON) in a custom
  // XML part. Opened in Word, its panel offers to restore them.
  // opts: { data (base64), lines: [text], addinRef, type }
  var BACKUP_NS = 'urn:lsk:backup', BACKUP_PROP = 'LabelSheetBackup';
  function backupDocumentXml(lines) {
    var p = function (t, size, b, after) {
      return '<w:p><w:pPr><w:spacing w:before="0" w:after="' + (after || 120) + '"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>' +
        (b ? '<w:b/>' : '') + '<w:sz w:val="' + size * 2 + '"/><w:szCs w:val="' + size * 2 + '"/></w:rPr><w:t xml:space="preserve">' + esc(t) + '</w:t></w:r></w:p>';
    };
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + '><w:body>' +
      p('Label Sheet Keeper backup', 20, true, 240) + lines.map(function (t) { return p(t, 11, false); }).join('') +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>' +
      '</w:body></w:document>';
  }
  function buildBackupDocx(JSZipLib, opts) {
    var Z = JSZipLib || root.JSZip;
    if (!Z) return Promise.reject(new Error('The file builder didn’t load. Close and reopen the panel.'));
    var we = opts.addinRef && opts.addinRef.id;
    var zip = new Z();
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>' +
      (we ? '<Override PartName="/word/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>' +
        '<Override PartName="/word/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/>' : '') +
      '</Types>');
    zip.file('_rels/.rels', rootRels(we));
    if (we) {
      zip.file('word/webextensions/taskpanes.xml', TASKPANES);
      zip.file('word/webextensions/_rels/taskpanes.xml.rels', TASKPANES_RELS);
      zip.file('word/webextensions/webextension1.xml', webExtensionXml(opts.addinRef));
    }
    zip.file('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">' +
      '<dc:title>Label Sheet Keeper backup</dc:title><dc:creator>Label Sheet Keeper</dc:creator></cp:coreProperties>');
    zip.file('docProps/custom.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="' + BACKUP_PROP + '"><vt:lpwstr>1</vt:lpwstr></property></Properties>');
    zip.file('word/document.xml', backupDocumentXml(opts.lines || []));
    zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>' +
      '</Relationships>');
    zip.file('customXml/item1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><lskBackup xmlns="' + BACKUP_NS + '" v="1">' + opts.data + '</lskBackup>');
    zip.file('customXml/itemProps1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="no"?>' +
      '<ds:datastoreItem ds:itemID="{6B1A4C2E-0F5D-4E8B-9C3A-2D7E1F0B5A94}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>');
    zip.file('customXml/_rels/item1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/>' +
      '</Relationships>');
    return zip.generateAsync({ type: opts.type || 'base64', compression: 'DEFLATE' });
  }

  var api = {
    PROP: PROP, LAYOUT_PROP: LAYOUT_PROP, SPEC_PROP: SPEC_PROP, SPACER_TAG: SPACER_TAG, WE_NS: WE_NS,
    ARROW_TEXT: ARROW_TEXT, ARROW_BEFORE_PT: ARROW_BEFORE_PT, ARROW_SIZE_PT: ARROW_SIZE_PT, STRIP_MARGIN: STRIP_MARGIN,
    hasStrip: hasStrip, layout: layout, count: count, cellOf: cellOf, labelAt: labelAt, tableShape: tableShape,
    buildDocx: buildDocx, documentXml: documentXml, tablePackage: tablePackage, paraXml: paraXml, specOf: specOf,
    buildBackupDocx: buildBackupDocx, BACKUP_NS: BACKUP_NS, BACKUP_PROP: BACKUP_PROP
  };
  root.LSK_SHEET = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
