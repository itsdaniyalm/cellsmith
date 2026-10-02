# Installing Cellsmith

Cellsmith is free and runs inside Excel. There is nothing to install on your computer: you only tell Excel where to find it. It takes about 5 minutes, you do it once, and updates arrive automatically after that.

**You need:**

- Microsoft 365 Excel (the subscription version of Excel)
- An internet connection (Cellsmith loads from the web each time you open it; your workbook data stays on your computer)

**Jump to:** [Step 1: Download](#step-1-download-the-cellsmith-file) · [Windows](#excel-on-windows-desktop-app) · [Web](#excel-on-the-web-browser) · [Mac](#excel-on-mac) · [Using it](#step-3-use-it) · [Troubleshooting](#troubleshooting) · [Removing it](#removing-cellsmith)

## Step 1: Download the Cellsmith file

1. Go to **[cellsmith.itsdaniyalm.workers.dev](https://cellsmith.itsdaniyalm.workers.dev)**.
2. Click **Download manifest**. A small file called `manifest.xml` is saved to your **Downloads** folder.

This file only tells Excel where Cellsmith lives. It contains no program code.

## Step 2: Add it to Excel

Follow the section for the Excel you use.

### Excel on Windows (desktop app)

The Windows desktop app only loads add-ins like this one from a *shared folder*, so first you create one. It is shared only with you.

**A. Create the folder**

1. Open **File Explorer** and click **This PC**, then open your **C:** drive.
2. Right-click an empty area and choose **New → Folder**. Name it `CellsmithAddin`.
3. Move `manifest.xml` from your **Downloads** folder into `CellsmithAddin`.

**B. Share the folder (only with yourself)**

1. Right-click the `CellsmithAddin` folder and choose **Properties**.
2. Open the **Sharing** tab and click **Share…**.
3. Your own name should already be in the list. Click **Share**, then **Done**.
4. Back on the **Sharing** tab, look under **Network Path**. You will see something like `\\YOUR-PC-NAME\CellsmithAddin`. Select all of it and copy it (<kbd>Ctrl</kbd>+<kbd>C</kbd>).
5. Click **Close**.

**C. Tell Excel about the folder**

1. Open Excel and any workbook (a blank one is fine).
2. Click **File → Options**. (On the start screen, **Options** is at the bottom left.)
3. Click **Trust Center**, then the **Trust Center Settings…** button.
4. Click **Trusted Add-in Catalogs** on the left.
5. Paste the network path into the **Catalog Url** box and click **Add catalog**.
6. In the list below, tick **Show in Menu** next to your folder.
7. Click **OK**, then **OK** again.
8. **Close Excel completely** (every Excel window), then open it again.

**D. Turn on Cellsmith**

1. On the **Home** tab, click **Add-ins**, then **More Add-ins**. (In some versions this is **Insert → Get Add-ins**.)
2. At the top of the window that opens, click the **SHARED FOLDER** tab.
3. Click **Cellsmith**, then **Add**.

A **Formula Editor** button now appears at the right-hand end of the **Home** tab, and it stays there every time you open Excel.

> **On a work computer?** Your IT team may block folder sharing or the Trust Center setting. If a step is greyed out or fails, send your IT admin a link to this page. They can deploy Cellsmith to you (or your whole team) from the Microsoft 365 admin center under **Settings → Integrated apps**, using the same `manifest.xml`.

### Excel on the web (browser)

> Not tested yet on Excel for the web. If something doesn't work, please [report it](https://github.com/itsdaniyalm/cellsmith/issues).

1. Open any workbook at [office.com](https://www.office.com).
2. On the **Home** tab, click **Add-ins**, then **More Add-ins**. (Or **Insert → Add-ins**.)
3. Click **My Add-ins**, then **Upload My Add-in**.
4. Click **Browse**, choose the `manifest.xml` you downloaded, and click **Upload**.

### Excel on Mac

> Not tested yet on Mac. If something doesn't work, please [report it](https://github.com/itsdaniyalm/cellsmith/issues).

1. Quit Excel.
2. In **Finder**, click **Go → Go to Folder…** in the menu bar and paste this, then press **Return**:
   ```
   ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
   ```
   If Finder says the folder doesn't exist, go up one level to `Documents` and create a new folder named `wef`.
3. Copy `manifest.xml` into that folder.
4. Open Excel and any workbook, then click **Insert → Add-ins → My Add-ins** and choose **Cellsmith**.

## Step 3: Use it

1. Click any cell that has a formula.
2. Click **Home → Formula Editor**. A panel opens on the right showing that formula, neatly formatted.
3. Edit the formula in the panel. When it is valid, Cellsmith writes it back to the cell for you. Problems are listed under the editor while you type.
4. Click another cell to edit its formula. The panel follows your selection.

Handy keys while you are in the panel:

| Keys | What it does |
|---|---|
| <kbd>Shift</kbd>+<kbd>Alt</kbd>+<kbd>F</kbd> | Format the formula |
| <kbd>Ctrl</kbd>+<kbd>Space</kbd> | Show suggestions |
| <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>Space</kbd> | Show the function's arguments |
| <kbd>Ctrl</kbd>+<kbd>Enter</kbd> | Write to the cell right away |
| <kbd>Ctrl</kbd>+<kbd>.</kbd> | Quick fix (for example, correct a misspelled function) |
| <kbd>F2</kbd> | Rename a `LET` variable everywhere it is used |

## Troubleshooting

**There is no SHARED FOLDER tab (Windows).**
Check that **Show in Menu** is ticked in **Trusted Add-in Catalogs**, then close Excel completely and reopen it. If it is still missing, Excel may still be running in the background: open **Task Manager** (<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>Esc</kbd>), end any **Microsoft Excel** task, and start Excel again.

**The SHARED FOLDER tab is empty (Windows).**
Make sure `manifest.xml` is inside the `CellsmithAddin` folder (not in a subfolder), and that the path you pasted starts with `\\`, not `C:\`.

**The panel is blank or shows an error.**
Cellsmith loads from the internet. Check that you can open [cellsmith.itsdaniyalm.workers.dev](https://cellsmith.itsdaniyalm.workers.dev) in your browser, then close and reopen the panel.

**The Formula Editor button disappeared.**
Repeat [step D](#excel-on-windows-desktop-app) (Windows) or the last step for your version.

**My edit didn't reach the cell.**
Cellsmith only writes formulas that are valid, because Excel would reject them anyway. Look at the problems listed under the editor and fix them; the formula is written as soon as they are gone. If you are typing inside the cell in Excel at the same time, press <kbd>Esc</kbd> in Excel first.

Still stuck? [Open an issue](https://github.com/itsdaniyalm/cellsmith/issues) and describe what you see.

## Removing Cellsmith

- **Windows:** in **File → Options → Trust Center → Trust Center Settings… → Trusted Add-in Catalogs**, select your folder and click **Remove**. Then delete the `CellsmithAddin` folder.
- **Web:** **Home → Add-ins → More Add-ins → My Add-ins**, click the **…** on Cellsmith and choose **Remove**.
- **Mac:** delete `manifest.xml` from the `wef` folder.
