/*
  Label sheet sizes for Label Sheet Keeper.

  HOW TO ADD A NEW SIZE
  1. Copy one of the entries below and paste it at the end of the list.
  2. Give it a new, unique id (lowercase letters, numbers and dashes).
  3. Fill in the measurements in inches. You can find them on the label
     package or in the maker's Word template (Layout > Margins, and the
     table's column widths and row heights):
       page        the whole sheet
       across/down how many labels in each row / in each column
       label       one label's width and height
       firstLabel  from the sheet's left edge and top edge to the first
                   (top-left) label's corner
       pitch       from one label's left edge to the next label's left edge
                   (across), and from one label's top edge to the next
                   label's top edge (down). If labels touch, pitch equals
                   the label size.
       padding     blank space kept inside each label so text stays clear
                   of the edges
       align       where text sits in the label: "top", "center" or "bottom"
       font        starting font and size for new sheets
  4. Save the file. GitHub publishes it in a minute or two. In Word, open
     the panel, choose "Make a new label sheet", and the new size is listed.
*/
(function (root) {
  'use strict';

  var FORMATS = [
    {
      id: 'avery-5160',
      name: 'Address labels, 30 per sheet',
      detail: '1″ × 2⅝″ · fits Avery 5160 / 8160',
      page: { width: 8.5, height: 11 },
      across: 3,
      down: 10,
      label: { width: 2.625, height: 1 },
      firstLabel: { left: 0.1875, top: 0.5 },
      pitch: { across: 2.75, down: 1 },
      padding: { top: 0.03, right: 0.1, bottom: 0.03, left: 0.1 },
      align: 'center',
      font: { name: 'Arial', size: 10 },
      corner: 0.0625
    },
    {
      id: 'half-sheet-2',
      name: 'Shipping labels, 2 per sheet',
      detail: '8½″ × 5½″ · fits Avery 5126 / 8126 and other 2-per-page shipping labels',
      page: { width: 8.5, height: 11 },
      across: 1,
      down: 2,
      label: { width: 8.5, height: 5.5 },
      firstLabel: { left: 0, top: 0 },
      pitch: { across: 8.5, down: 5.5 },
      padding: { top: 0.4, right: 0.4, bottom: 0.4, left: 0.4 },
      align: 'top',
      font: { name: 'Arial', size: 14 },
      corner: 0
    }
  ];

  // Basic sanity checks so a typo in a new entry shows a clear message.
  function problems(f) {
    var out = [];
    var need = ['id', 'name', 'page', 'across', 'down', 'label', 'firstLabel', 'pitch'];
    need.forEach(function (k) { if (f[k] === undefined) out.push('missing "' + k + '"'); });
    if (out.length) return out;
    if (f.pitch.across + 1e-6 < f.label.width && f.across > 1) out.push('pitch.across is smaller than the label width');
    if (f.pitch.down + 1e-6 < f.label.height && f.down > 1) out.push('pitch.down is smaller than the label height');
    var right = f.firstLabel.left + (f.across - 1) * f.pitch.across + f.label.width;
    var bottom = f.firstLabel.top + (f.down - 1) * f.pitch.down + f.label.height;
    if (right > f.page.width + 1e-6) out.push('labels run past the right edge of the page');
    if (bottom > f.page.height + 1e-6) out.push('labels run past the bottom edge of the page');
    return out;
  }

  function withDefaults(f) {
    f.padding = f.padding || { top: 0.05, right: 0.1, bottom: 0.05, left: 0.1 };
    f.align = f.align === 'top' || f.align === 'bottom' ? f.align : 'center';
    f.font = f.font || { name: 'Arial', size: 10 };
    f.corner = f.corner || 0;
    f.detail = f.detail || '';
    return f;
  }

  var VALID = [];
  var INVALID = [];
  FORMATS.forEach(function (f) {
    var p = problems(f);
    if (p.length) INVALID.push({ id: f.id || '(no id)', problems: p });
    else VALID.push(withDefaults(f));
  });

  function getFormat(id) {
    for (var i = 0; i < VALID.length; i++) if (VALID[i].id === id) return VALID[i];
    return null;
  }

  var api = { formats: VALID, invalid: INVALID, getFormat: getFormat, problems: problems };
  root.LSK_FORMATS = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
