# Label Sheet Keeper for Word

A side panel for Microsoft Word that keeps track of which labels on a sheet are already used, so partly used sheets don't go to waste.

- The Word document **is** the label sheet. Type and format labels right in it.
- The panel shows a map of the sheet: used, filled, and open spots.
- Copy one label to the next open spots, print with ⌘P, then click **I printed these**. Those spots become used and locked, so you can't type in them by accident.
- The record of used spots is saved inside the Word document, so it goes wherever the file goes.
- The panel notices when Word prints the sheet and marks the filled labels used (with Undo).
- After any change, the panel saves the document for you (once it has been saved the first time).
- Label sheets open with the panel already showing. **About this panel**, at the bottom of the panel, shows whether that's set up.
- **Fill every open label** copies one label to every open label, for a sheet of return address labels.

Sizes included:

| Size | Labels | Fits |
| --- | --- | --- |
| 1″ × 2⅝″ | 30 per sheet | Avery 5160 / 8160 |
| 5½″ × 8½″, sideways page | 2 per sheet | Avery 5126 / 8126 and other 2-per-page shipping labels |
| ½″ × 1¾″ | 80 per sheet | Avery 5167 / 8167 |
| 2″ × 4″ | 10 per sheet | Avery 5163 / 8163 |
| 3⅓″ × 4″ | 6 per sheet | Avery 5164 / 8164 |
| 8½″ × 11″ | 1 per sheet | Avery 5165 / 8165 (full-sheet; good for a carrier label at full size) |

Print the test page once for any size you haven't used before, to confirm it lines up with your labels.

Panel address: https://therearenousernamesleft.github.io/label-sheet-keeper/taskpane.html

## One-time setup on a Mac

1. Download `manifest.xml` from this repository.
2. Quit Word.
3. In Finder, press **⌘⇧G** and go to:
   `~/Library/Containers/com.microsoft.Word/Data/Documents/`
4. If there's no folder named `wef` there, make one (File › New Folder, name it `wef`).
5. Put `manifest.xml` inside `wef`.
6. Open Word and open any document. On the **Home** tab, click **Label Sheet** at the right end. If you don't see it, click **Add-ins** on the Home tab and pick **Label Sheet Keeper** from the list.

## Everyday use

1. Click **Label Sheet**, choose your labels, and click **Make the sheet**. A new document opens with the panel. Save it, for example as `Labels 30 per sheet.docx`.
2. Click each label on the map that's already peeled off your real sheet.
3. Type a label in the document. Leave the cursor in it, set **Copies**, and click **Copy** (or **Fill every open label**).
4. Print with ⌘P. In the print dialog, keep Scale at 100%. If Word warns that margins are outside the printable area, click **Ignore**. The labels have their own blank space inside.
5. The panel notices the print and marks those labels used (click **Undo** if they didn't come out right). If it ever misses a print, click **I printed these**.

With a new printer or a new pack of labels, click **Make a test page** in the panel first. Print it on plain paper and hold it behind a label sheet against a light to check that the boxes line up. Word shows the label edges on screen if you turn on **View Gridlines** on the table **Layout** tab; gridlines never print.

### Shipping labels (2 per sheet)

The page shows sideways with the two labels standing upright, side by side. In the panel:

- **Label file** tab: choose an Amazon return label PDF or a picture of any label (or drop it on the panel). Drag across the label on the preview to leave out instructions or a packing slip, then click **Put on the next open label**. The label is sized to fit and turned upright if needed.
- **Address** tab: fill in From and Ship to. Your return address is remembered for next time.

Print the test page once and tell the panel whether Label 1 came out on the top or bottom half of your sheet. From then on the panel calls them the top label and bottom label.

Keep one saved document per physical sheet you have in progress. When you load a fresh sheet of labels, click **Reset for a fresh sheet**.

## Adding a new label size

1. Open `formats.js` in this repository on GitHub and click the pencil icon to edit.
2. Copy one of the entries between `{` and `},`, paste it at the end of the list, and change the values. The comment at the top of the file explains each measurement.
3. Click **Commit changes**. GitHub publishes the change in a minute or two.
4. In Word, open the panel and click **Make another sheet**. The new size is listed. Print a test page before using it on real labels.

If a measurement doesn't add up (for example labels running off the page), the panel skips that size and says why.

## Updating the panel

All of the panel's code lives in this repository. Change a file here, commit, and Word picks up the new version the next time the panel opens (the panel loads its files with a fresh address each time). Changes to `taskpane.html` itself can take up to 10 minutes to show, because GitHub lets Word keep a copy of that page for that long.

You only need to touch `manifest.xml` on your Mac again if the add-in's name or web address changes.

## Files

| File | What it does |
| --- | --- |
| `manifest.xml` | Setup file for Word on your Mac |
| `taskpane.html`, `taskpane.css`, `taskpane.js` | The side panel |
| `formats.js` | The list of label sizes |
| `sheetdoc.js` | Builds a Word document laid out for a label size |
| `templates/` | Ready-made blank sheets for each size, in case Word can't open a new document from the panel |
| `vendor/pdfjs/` | Mozilla PDF.js 3.11.174 (Apache 2.0 license), used to read label PDFs |
| `index.html` | This project's web page with the setup steps |
