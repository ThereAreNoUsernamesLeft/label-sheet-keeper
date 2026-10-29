/*
  Label sheet sizes for Label Sheet Keeper.

  HOW TO ADD A NEW SIZE
  1. Copy one of the entries below and paste it where you want it in the
     list (the panel shows sizes in this order).
  2. Give it a new, unique id (lowercase letters, numbers and dashes).
  3. Fill in the measurements in inches. You can find them on the label
     package or in the maker's Word template (Layout > Margins, and the
     table's column widths and row heights):
       page        the whole sheet as you want to see it in Word. Use
                   { width: 11, height: 8.5 } for a sideways page; the
                   physical sheet still feeds into the printer the same way.
       across/down how many labels in each row / in each column
       label       one label's width and height
       firstLabel  from the sheet's left edge and top edge to the first
                   (top-left) label's corner
       pitch       from one label's left edge to the next label's left edge
                   (across), and from one label's top edge to the next
                   label's top edge (down). If labels touch, pitch equals
                   the label size.
       padding     blank space kept inside the left and right sides of each
                   label so text stays clear of the edges
       align       where text sits from top to bottom: "center" (keeps text
                   clear of the top and bottom edges), "top" or "bottom"
       topSpace    with align "top": blank space above the first line, so
                   text clears the top edge (default 0.4)
       font        starting font and size for new sheets
       shipping    true to show the shipping tools in the panel: put a
                   label file (PDF or picture) on the sheet, or build an
                   address label
  4. Commit the change. GitHub publishes it in a minute or two. In Word,
     open the panel and click "Make another sheet": the new size is listed.
     Print a test page before using it on real labels.
*/
(function (root) {
  'use strict';

  // Listed in the panel in this order.
  var FORMATS = [
    {
      id: 'avery-5160',
      name: 'Address labels, 30 per sheet',
      detail: '1″ × 2⅝″ · Avery 5160, 8160',
      page: { width: 8.5, height: 11 },
      across: 3,
      down: 10,
      label: { width: 2.625, height: 1 },
      firstLabel: { left: 0.1875, top: 0.5 },
      pitch: { across: 2.75, down: 1 },
      padding: { left: 0.1, right: 0.1 },
      align: 'center',
      font: { name: 'Arial', size: 10 },
      corner: 0.0625
    },
    {
      id: 'avery-5167',
      name: 'Return address labels, 80 per sheet',
      detail: '½″ × 1¾″ · Avery 5167, 8167',
      page: { width: 8.5, height: 11 },
      across: 4,
      down: 20,
      label: { width: 1.75, height: 0.5 },
      firstLabel: { left: 0.3, top: 0.5 },
      pitch: { across: 2.05, down: 0.5 },
      padding: { left: 0.06, right: 0.06 },
      align: 'center',
      font: { name: 'Arial', size: 7 },
      corner: 0.05
    },
    {
      id: 'avery-5163',
      name: 'Shipping labels, 10 per sheet',
      detail: '2″ × 4″ · Avery 5163, 8163',
      page: { width: 8.5, height: 11 },
      across: 2,
      down: 5,
      label: { width: 4, height: 2 },
      firstLabel: { left: 0.15625, top: 0.5 },
      pitch: { across: 4.1875, down: 2 },
      padding: { left: 0.15, right: 0.15 },
      align: 'center',
      font: { name: 'Arial', size: 12 },
      corner: 0.0625
    },
    {
      id: 'avery-5164',
      name: 'Shipping labels, 6 per sheet',
      detail: '3⅓″ × 4″ · Avery 5164, 8164',
      page: { width: 8.5, height: 11 },
      across: 2,
      down: 3,
      label: { width: 4, height: 3.3333 },
      firstLabel: { left: 0.15625, top: 0.5 },
      pitch: { across: 4.1875, down: 3.3333 },
      padding: { left: 0.2, right: 0.2 },
      align: 'center',
      font: { name: 'Arial', size: 12 },
      corner: 0.0625
    },
    {
      // The physical sheet is 8½ × 11 with two 8½ × 5½ labels. Shown sideways
      // so each label stands upright, like a 4 × 6 shipping label.
      id: 'shipping-2-sideways',
      name: 'Shipping labels, 2 per sheet',
      detail: '5½″ × 8½″, shown sideways · Avery 5126, 8126',
      page: { width: 11, height: 8.5 },
      across: 2,
      down: 1,
      label: { width: 5.5, height: 8.5 },
      firstLabel: { left: 0, top: 0 },
      pitch: { across: 5.5, down: 8.5 },
      padding: { left: 0.4, right: 0.4 },
      align: 'top',
      topSpace: 0.45,
      font: { name: 'Arial', size: 12 },
      corner: 0,
      shipping: true
    },
    {
      id: 'full-sheet-1',
      name: 'Full-sheet labels, 1 per sheet',
      detail: '8½″ × 11″ · Avery 5165, 8165',
      page: { width: 8.5, height: 11 },
      across: 1,
      down: 1,
      label: { width: 8.5, height: 11 },
      firstLabel: { left: 0, top: 0 },
      pitch: { across: 8.5, down: 11 },
      padding: { left: 0.4, right: 0.4 },
      align: 'top',
      topSpace: 0.45,
      font: { name: 'Arial', size: 14 },
      corner: 0,
      shipping: true
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
    f.padding = f.padding || {};
    if (typeof f.padding.left !== 'number') f.padding.left = 0.1;
    if (typeof f.padding.right !== 'number') f.padding.right = 0.1;
    f.align = f.align === 'top' || f.align === 'bottom' ? f.align : 'center';
    f.topSpace = typeof f.topSpace === 'number' ? f.topSpace : 0.4;
    f.shipping = !!f.shipping;
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
