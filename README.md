# Label Sheet Keeper for Word

A side panel for Microsoft Word that keeps track of which labels on a sheet are already used, so partly used sheets don't go to waste.

- The Word document **is** the label sheet. Type and format labels right in it.
- The panel shows a map of the sheet: used, filled, and open spots.
- Copy one label to the next open spots, print with ⌘P, then click **I printed these**. Those spots become used and locked, so you can't type in them by accident.
- The record of used spots is saved inside the Word document, so it goes wherever the file goes.

Sizes included:

| Size | Labels | Fits |
| --- | --- | --- |
| 1″ × 2⅝″ | 30 per sheet | Avery 5160 / 8160 |
| 5½″ × 8½″, sideways page | 2 per sheet | Avery 5126 / 8126 and other 2-per-page shipping labels |

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

1. Click **Label Sheet**, choose a size, and click **Make a new label sheet**. A new document opens. Save it, for example as `Labels 30 per sheet.docx`.
2. In that document, click **Label Sheet** again. Click spots on the map that are already peeled off your real sheet.
3. Type a label in the document. Leave the cursor in it, set **Copies**, and click **Copy to next open**.
4. Print with ⌘P. In the print dialog, keep Scale at 100%. If Word warns that margins are outside the printable area, click **Ignore**. The labels have their own blank space inside.
5. When the labels came out right, click **I printed these**. Save the document.

With a new printer or a new pack of labels, click **Make a test page** in the panel first. Print it on plain paper and hold it behind a label sheet against a light to check that the boxes line up. Word shows the label edges on screen if you turn on **View Gridlines** on the table **Layout** tab; gridlines never print.

### Shipping labels (2 per sheet)

The page shows sideways with the two labels standing upright, side by side. In the panel:

- **Use a label file:** choose an Amazon return label PDF or a picture of any label (or drop it on the panel). Drag across the label on the preview to leave out instructions or packing slips, then click **Put on the next open label**. The label is sized to fit and turned upright if needed.
- **Type an address label:** fill in From and Ship to. Your return address is remembered for next time.

Print the test page once to see which half of your physical sheet is label 1, so you know which spot to mark when you peel one off.

Keep one saved document per physical sheet you have in progress. When you load a fresh sheet of labels, click **Start a new sheet**.

## Adding a new label size

1. Open `formats.js` in this repository on GitHub and click the pencil icon to edit.
2. Copy one of the entries between `{` and `},`, paste it at the end of the list, and change the values. The comment at the top of the file explains each measurement.
3. Click **Commit changes**. GitHub publishes the change in a minute or two.
4. In Word, open the panel, click **Make a different size**, and the new size is listed.

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
