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

1. Click **Label Sheet**, choose your labels, and click **Make the sheet**. A new document opens with the panel. Save it, for example as `Labels 30 per sheet.docx`.
2. **Mark used labels:** click each label on the map that's already peeled off your sheet.
3. **Fill labels:**
   - **Copy:** type a label in the document, leave the cursor in it, and click **Copy** or **Fill every open label**. **Fit text** shrinks text that's too big (never below a readable size); labels whose text may be cut off get a red outline on the map.
   - **Address:** type an address or pick a saved one, choose a design (Plain, Bold name, Return to, With logo), and put it on the next open label or every open label.
   - **Table:** opens a bigger window like a small spreadsheet, with one row for each open label. The first column shows which label a row goes on. Type addresses, paste cells from Excel or Numbers, or open an Excel, CSV or contacts file. Rows past your open labels are marked "New sheet"; **Fill labels** puts them on new sheets, each in its own window.
4. **Print** with ⌘P at 100% scale. Feed the sheet top edge first, print side up, and choose **Labels** as the paper type. The panel notices the print and marks those labels used.

When you load a fresh sheet of labels, click **Reset for a fresh sheet**.

### Printing on partly used sheets

Avery says laser label sheets should go through a laser printer only once, because the heat softens the glue. Each extra trip raises the chance of a label peeling off inside the printer. The panel helps keep trips down:

- It counts each sheet's **trips** through the printer and warns before a second trip. You set the limit in **Settings** (3 to start).
- On a sheet's **last trip**, it offers to fill the remaining open labels (with your return address or copies of a label) so nothing goes to waste.
- Let several labels pile up as "ready to print," then print them in one trip.
- Press down any lifted label edges before feeding a sheet, and retire a sheet with a peeling label.

### Feed arrow

On sheets with a blank strip above the labels, the panel prints a small "▲ TOP · feed this edge into the printer first" in that strip, so you can tell which way a partly used sheet goes in. Turn it off in **Settings**.

### Test page and printer adjustment

With a new printer or a new pack of labels, click **Make a test page** and print it on scrap paper. Hold it behind a label sheet against a light. The panel then asks whether it lined up. If the boxes are off the same way every time, enter the shift in **Settings › Printer adjustment**. New sheets use it, and **Apply to this sheet** moves the labels on the sheet you have open.

### Shipping labels (2 per sheet and full sheet)

- **Label file:** choose an Amazon return label PDF or a picture of any label, drop it on the panel, or copy a label image and press ⌘V in the panel. The panel finds the label's printed border by itself; drag across the preview to change it. Then click **Put on the next open label**.
- **Address:** fill in From and Ship to (or pick a saved address). Your return address is remembered.
- Print the test page once and tell the panel which half of your sheet Label 1 came out on; from then on it says "top label" and "bottom label."

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
| `lsk-tools.js` | Reads address lists and contacts, fits text, finds label borders |
| `templates/` | Ready-made blank sheets for each size, in case Word can't open a new document from the panel |
| `vendor/pdfjs/` | Mozilla PDF.js 3.11.174 (Apache 2.0 license), used to read label PDFs |
| `vendor/jszip/` | JSZip 3.10.1 (MIT license), used to build and read Word and Excel files |
| `index.html` | This project's web page with the setup steps |
