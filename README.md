# Label Sheet Keeper for Word

A side panel for Microsoft Word that keeps track of which labels on a sheet are already used, so partly used sheets don't go to waste.

- The Word document **is** the label sheet. Type and format labels right in it.
- The panel shows a map of the sheet: used, ready to print, and open labels.
- Fill labels by copying one you typed, by placing an address (with a design), or from a table of many addresses.
- The panel notices when Word prints the sheet, marks those labels used (with Undo), and counts the sheet's trips through the printer.
- Everything about a sheet is saved inside its Word document, so it goes wherever the file goes.

## Label sizes

| Size | Labels | Fits |
| --- | --- | --- |
| 1″ × 2⅝″ | 30 per sheet | Avery 5160 (laser), 8160 (inkjet) |
| ½″ × 1¾″ | 80 per sheet | Avery 5167 (laser), 8167 (inkjet) |
| 2″ × 4″ | 10 per sheet | Avery 5163 (laser), 8163 (inkjet) |
| 3⅓″ × 4″ | 6 per sheet | Avery 5164 (laser), 8164 (inkjet) |
| 5½″ × 8½″, sideways page | 2 per sheet | Avery 5126 (laser), 8126 (inkjet) |
| 8½″ × 11″ | 1 per sheet | Avery 5165 (laser), 8165 (inkjet) |

Use the laser version of a label on a laser printer. You can add your own sizes in the panel (**Add your own size** when making a sheet) or in `formats.js`.

Panel address: https://therearenousernamesleft.github.io/label-sheet-keeper/taskpane.html

## One-time setup on a Mac

1. Download `manifest.xml` from this repository.
2. Quit Word.
3. In Finder, press **⌘⇧G** and go to `~/Library/Containers/com.microsoft.Word/Data/Documents/`
4. If there's no folder named `wef` there, make one (File › New Folder, name it `wef`).
5. Put `manifest.xml` inside `wef`.
6. Open Word and any document. On the **Home** tab, click **Label Sheet**. If you don't see it, click **Add-ins** on the Home tab and choose **Label Sheet Keeper**.

## Everyday use

Point at an **ⓘ** in the panel for a short explanation of that tool.

1. Click **Label Sheet**, choose your labels, and click **Make the sheet**. A new document opens with the panel. Save it, for example as `Labels 30 per sheet.docx`.
2. **Mark used labels:** click each label on the map that's already peeled off your sheet.
3. **Fill labels:**
   - **Copy:** type a label in the document, leave the cursor in it, and click **Copy** or **Fill every open label**. **Fit text** shrinks text that's too big (never below a readable size); labels whose text may be cut off get a red outline on the map.
   - **Address:** type an address or pick a saved one, choose a design (Plain, Bold name, Return to, With logo), and put it on the next open label or every open label.
   - **Table:** opens a bigger window like a small spreadsheet, with one row for each open label. The first column shows which label a row goes on. Type addresses, paste cells from Excel or Numbers, or open an Excel, CSV or contacts file. Rows past your open labels are marked "New sheet"; **Fill labels** puts them on new sheets, each in its own window.
   - **Numbers:** numbers labels in order ("Box 1 of 12", "BIN-007"), with a QR code or a Code 128 barcode under each if you like. Numbering picks up where you left off on the next sheet.
4. **Print** with ⌘P at 100% scale. Feed the sheet top edge first, print side up, and choose **Labels** as the paper type. The panel notices the print and marks those labels used.

When you load a fresh sheet of labels, click **Reset for a fresh sheet**.

### Printing on partly used sheets

Avery says laser label sheets should go through a laser printer only once, because the heat softens the glue. Each extra trip raises the chance of a label peeling off inside the printer. The panel helps keep trips down:

- It counts each sheet's **trips** through the printer and warns before a second trip. You set the limit in **Settings** (3 to start).
- On a sheet's **last trip**, it offers to fill the remaining open labels (with your return address or copies of a label) so nothing goes to waste.
- Let several labels pile up as "ready to print," then print them in one trip.
- When a table or a set of copies doesn't fit on this sheet, and you have other part-used sheets of the same size, the panel offers to fill those first. It opens each one in turn; that sheet's panel fills its open labels and opens the next. Anything left goes on new sheets.
- Press down any lifted label edges before feeding a sheet, and retire a sheet with a peeling label.

### Feed arrow

On sheets with a blank strip above the labels, the panel prints a small "▲ TOP · feed this edge into the printer first" in that strip, so you can tell which way a partly used sheet goes in. Turn it off in **Settings**.

### Test page and printer adjustment

With a new printer or a new pack of labels, click **Make a test page** and print it on scrap paper. Hold it behind a label sheet against a light. The panel then asks whether it lined up. If the boxes are off the same way every time, enter the shift in **Settings › Printer adjustment**. New sheets use it, and **Apply to this sheet** moves the labels on the sheet you have open.

If your printer has a scanner, the panel can measure the test page for you: scan the printed test page (the whole page, at 100%, top edge against the top of the glass), then click **Use a scan** (after a test page, or in **Settings › Printer adjustment**) and choose the scan (PDF, PNG or JPEG). The panel finds the boxes, says how far off they printed, and offers the adjustment. This works for sizes that print upright.

### Shipping labels (2 per sheet and full sheet)

- **Label file:** choose an Amazon return label PDF or a picture of any label, drop it on the panel, or copy a label image and press ⌘V in the panel. The panel finds the label's printed border by itself; drag across the preview to change it. Then click **Put on the next open label**.
- **Several labels at once:** choose or drop several PDFs or pictures. Each label goes on a waiting list; **Put the next labels on this sheet** fills the open labels in order. After printing, load a fresh sheet and click it again. For one PDF with several labels, **Use every page** adds every page that has a label border.
- **Address:** fill in From and Ship to (or pick a saved address). Your return address is remembered.
- Print the test page once and tell the panel which half of your sheet Label 1 came out on; from then on it says "top label" and "bottom label."

### Address suggestions

In the table window, type a house number and street, and the panel suggests matching US and Canadian addresses. Pick one to fill in the street, city, state or province, and ZIP or postal code. A ZIP code or postal code on its own also fills in the city (using the free Zippopotam.us service).

Suggestions come from [Photon by Komoot](https://photon.komoot.io/), a free service with no sign-up, using [OpenStreetMap](https://www.openstreetmap.org/copyright) data. What you type is sent to Photon. Photon is shared by everyone, so it asks for light use, and some house numbers aren't in OpenStreetMap yet. Turn suggestions off in **Settings › Address suggestions**.

**Bring your own key (optional).** For your own daily allowance, use a free Geoapify key instead:

1. Sign up for a free account at [geoapify.com](https://www.geoapify.com/). No credit card is needed; the free plan includes 3,000 lookups a day.
2. Create a project and copy its API key.
3. In the key's settings, add `https://therearenousernamesleft.github.io` as an allowed origin (or HTTP referrer), so the key only works from this add-in.
4. In Word, open the panel, then **Settings › Address suggestions**. Paste the key and click **Save key**.

The key is kept by the panel on your Mac only, not in this repository or in your Word documents. With a key, suggestions go to Geoapify instead of Photon. On another Mac, paste the key there too.

### Canadian addresses

Canadian addresses follow Canada Post: `5-123 Main St` for a unit, and `Toronto ON  M5V 3L9` on the last line (two spaces before the postal code). Mail to Canada gets a `CANADA` line. If you mail from Canada, choose **Settings › You mail from › Canada**, and US addresses get a `USA` line instead.

### Label stock and reordering

Open **Label stock** on a sheet and enter how many sheets of those labels you have, how many come in a pack, and when to remind you. A sheet comes off the count the first time it goes through the printer (Undo puts it back). When you're running low, the Print step shows a reminder with **Reorder** and **I got a new pack**.

Paste a link to the labels you buy. **Reorder** then:

- **Amazon** (a product link with `/dp/` in it): adds the number of packs you chose to your Amazon cart, using Amazon's add-to-cart link (Amazon may ask you to confirm first).
- **Walmart** (a product link with `/ip/` in it): adds the packs to your Walmart cart.
- **Any other store** (Staples, Office Depot, Avery…): opens the product page; add the labels to your cart there.

Buy the kind made for your printer: laser labels (Avery 5160, 5126…) for a laser printer, inkjet labels (8160, 8126…) for an inkjet.

### Back up and restore settings

**Settings › Backup** (or **Back up or restore settings** in a document that isn't a sheet) makes a Word document holding your address book, return address, label sizes, printer adjustments, label stock and logo. Your Geoapify key is left out unless you tick **Include my Geoapify key**. Save the document somewhere safe; it holds your addresses, so keep it private.

To restore, open the backup in Word and click **Restore these settings**, or click **Restore from a file** in the panel and choose it. Restoring replaces the panel's settings on that Mac. Your label sheets are their own documents and aren't part of the backup.

### Your sheets

**Your sheets** lists every sheet with labels left, with open labels and trips used. When you make a new sheet in a size you already have, the panel asks whether to finish the old one first.

### Repair

If the label layout gets damaged (a row deleted, text typed above the labels, a table pasted inside a label), the panel says so. **Repair this sheet** rebuilds the layout and keeps used labels and labels ready to print.

## Adding a new label size in formats.js

1. Open `formats.js` in this repository on GitHub and click the pencil icon to edit.
2. Copy one of the entries between `{` and `},`, paste it, and change the values. The comment at the top of the file explains each measurement.
3. Click **Commit changes**. GitHub publishes the change in a minute or two.
4. In Word, open the panel and click **Make another sheet**. The new size is listed. Print a test page before using it on real labels.

## Updating the panel

All of the panel's code lives in this repository. Change a file here, commit, and Word picks up the new version the next time the panel opens (the panel loads its files with a fresh address each time). Changes to `taskpane.html` or `table.html` themselves can take up to 10 minutes to show, because GitHub lets Word keep a copy of those pages for that long.

You only need to touch `manifest.xml` on your Mac again if the add-in's name or web address changes.

## Files

| File | What it does |
| --- | --- |
| `manifest.xml` | Setup file for Word on your Mac |
| `taskpane.html`, `taskpane.css`, `taskpane.js` | The side panel |
| `table.html`, `table.css`, `table.js` | The table window |
| `formats.js` | The list of label sizes |
| `sheetdoc.js` | Builds Word documents laid out for a label size |
| `lsk-tools.js` | Reads address lists and contacts, formats US and Canadian addresses, looks up ZIP and postal codes, suggests addresses (Photon or Geoapify), fits text, finds label borders, makes reorder links, numbers, QR codes and Code 128 barcodes, and measures scanned test pages |
| `templates/` | Ready-made blank sheets for each size, in case Word can't open a new document from the panel |
| `vendor/pdfjs/` | Mozilla PDF.js 3.11.174 (Apache 2.0 license), used to read label PDFs |
| `vendor/jszip/` | JSZip 3.10.1 (MIT license), used to build and read Word and Excel files |
| `vendor/qrcode-generator/` | qrcode-generator 2.0.4 by Kazuhiko Arase (MIT license), used to make QR codes. "QR Code" is a registered trademark of DENSO WAVE INCORPORATED |
| `index.html` | This project's web page with the setup steps |
