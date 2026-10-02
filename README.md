<p align="center"><img src="assets/icon-128.png" alt="Cellsmith icon" width="96" height="96"></p>

# Cellsmith

An Excel add-in that gives you a real code editor (Monaco, the engine behind VS Code) for writing formulas: auto-format, autocomplete, argument hints, and live error checking, in a side pane that stays in sync with the selected cell.

Targets **Microsoft 365 Excel** (desktop and web), A1 notation, US-English function names.

## What you get

| | |
|---|---|
| **Readable formulas** | Pretty-prints nested formulas with indentation that fits the pane width. `LET` is expanded one binding per line; `LAMBDA` keeps parameters together. Syntax coloring, rainbow brackets, indent guides, folding. |
| **No syntax lookups** | Autocomplete for ~480 functions with snippet placeholders (Tab through the arguments), signature help with the active parameter highlighted, hover docs. Also completes `LET` variables, `LAMBDA` parameters, defined names, **named LAMBDAs**, sheets, tables and table columns (`Sales[`). |
| **Errors caught early** | Live diagnostics: syntax errors, wrong argument counts, unknown functions (with "did you mean"), undefined names, unreplaced snippet placeholders, unused `LET` variables (faded), `#REF!`, divide by zero, array constant mistakes, and Excel's hard limits (8,192 characters, 64 nesting levels, 255 arguments). Quick fixes via <kbd>Ctrl</kbd>+<kbd>.</kbd>. |
| **Refactoring** | <kbd>F2</kbd> renames a `LET` variable or `LAMBDA` parameter everywhere; <kbd>Ctrl</kbd>+click jumps to its definition. |
| **Two-way sync** | Select a cell and its formula loads into the editor. Edits are written back automatically once the formula is valid (Excel rejects malformed formulas, so invalid edits are never sent). |

### Keyboard

| Key | Action |
|---|---|
| <kbd>Ctrl</kbd>+<kbd>Enter</kbd> | Apply to the cell now |
| <kbd>Shift</kbd>+<kbd>Alt</kbd>+<kbd>F</kbd> | Format |
| <kbd>Ctrl</kbd>+<kbd>Space</kbd> | Suggestions |
| <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>Space</kbd> | Argument hints |
| <kbd>F2</kbd> | Rename variable |
| <kbd>F1</kbd> | Command palette (includes *Minify Formula*) |

## Running it

Requirements: Node 20+, Microsoft 365 Excel on Windows (uses the WebView2 runtime).

```bash
npm install

# One-time: create and trust a localhost HTTPS certificate (Windows shows a trust prompt).
npm run certs

# Start the dev server, register the add-in, and open Excel with it loaded.
npm run start:excel
```

Then open the **Home** tab and click **Formula Editor**. `npm run stop:excel` unregisters it.

If you prefer to do it in two steps: `npm run dev` in one terminal, `npm run sideload` in another.

**Preview without Excel:** `npm run dev`, then open `http://localhost:3000/taskpane.html` in a browser (works over plain HTTP if you haven't created certificates). It loads a demo workbook, so completion for tables, names and sheets can be tried, but nothing is written anywhere.

The dev certificate is created for 365 days (`office-addin-dev-certs` defaults to 30). After it expires, rerun `npm run certs`.

### Settings (⚙ in the pane)

Line width (fit pane / fixed), indent, when `LET` expands, pretty-print on cell select, upper-case function names, and **write mode**: *as typed* keeps your layout in the cell's formula bar; *minified* strips whitespace on write.

## How the sync behaves

- Loading never writes. Pretty-printing a formula on select only changes the editor; the cell changes when **you** edit.
- Formulas you have already laid out over several lines are left alone.
- If you change selection with a pending valid edit, it is flushed to the **previous** cell first. A pending edit that still has errors is discarded, and the status line says so.
- Multi-cell selections edit the first cell.
- If Excel is in cell-edit mode, writes are retried automatically.

## Building and hosting

```bash
npm run build          # type-checks, then bundles to dist/
npm test               # 277 unit tests
```

`dist/` is a static site. To use a hosted copy, serve it over HTTPS and generate a matching manifest:

```bash
npm run manifest -- https://your-host.example/cellsmith dist/manifest.xml
```

Users sideload that manifest or you deploy it through the Microsoft 365 admin center. Publishing publicly on AppSource additionally needs Microsoft validation, a privacy policy and a support page. The manifest ID in `manifest.template.xml` should stay stable once you ship.

The bundle is about 3.6 MB (920 KB gzipped): Monaco is imported selectively (`src/editor/monaco.ts`) so none of its 80+ bundled languages are included. Those are deep imports into `monaco-editor`'s ESM sources, so the version is pinned exactly; when upgrading, compare against `monaco-editor/esm/vs/editor/editor.main.js`.

## Project layout

```
src/core/       Pure TypeScript, no browser or Excel dependencies (fully unit-tested)
  tokenizer.ts    Excel formula lexer: refs, sheet prefixes, table refs, spill (#), whole row/column ranges
  parser.ts       Error-tolerant Pratt parser with Excel's operator precedence
  formatter.ts    Wadler-style pretty-printer and minifier
  analysis.ts     LET/LAMBDA scoping and the lint rules
  functions.ts    Function catalog loader (arity rules, signature help mapping)
  functionData.ts The catalog itself: edit this to add or fix functions
  cursor.ts       "What is the caret in?" helpers for completion and signature help
src/editor/     Monaco integration: language, themes, completion/hover/rename/quick-fix providers
src/excel/      Office.js adapter, standalone (preview) adapter, and the sync session state machine
src/taskpane/   The pane UI
tests/          Vitest suites
```

## Known limitations

- **Not yet exercised inside a real Excel session.** The editor, linter, formatter and sync logic are tested (unit tests, plus the editor driven in a browser against a mock workbook), but the Office.js adapter (`src/excel/office.ts`) has only been type-checked against the Office typings. See the checklist below.
- Excel 365 only: no R1C1 notation, no localized function names or separators, and no compatibility warnings for older Excel versions.
- The function catalog is hand-written and covers functions through 2025; brand-new functions need adding to `functionData.ts`. Unknown functions are reported as warnings, never blocked.
- Lint is structural: it checks names, scoping and argument counts, not argument *types* or values.
- Cells inside a dynamic-array spill range show their values, not the anchor's formula; select the anchor cell to edit.
- Legacy Ctrl+Shift+Enter array formulas (`{=...}`) are not specially handled.

### First-run checklist (things to confirm in Excel)

1. Select a cell with a formula: it appears in the pane, pretty-printed.
2. Edit it: the status goes *Editing → Applied* and the cell recalculates; the **Result** strip shows the new value.
3. After writing a multi-line formula, check the cell's formula bar: line breaks and indentation should be preserved. If Excel flattens them, switch **Write to cell** to *minified*.
4. Select another cell mid-edit: the edit lands in the first cell, not the second.
5. Press F2 on a cell (Excel edit mode) while typing in the pane: the status shows *Excel busy* and recovers when you leave edit mode.
6. In a workbook with a table and a named LAMBDA, type `Table1[` and the LAMBDA's name: columns and signature help should appear.

## Contributing

Issues and pull requests are welcome; see [CONTRIBUTING.md](CONTRIBUTING.md). The function catalog (`src/core/functionData.ts`) is the easiest place to start.

## License

[MIT](LICENSE). Third-party attributions are in [NOTICE.md](NOTICE.md).

Cellsmith is an independent project and is not affiliated with or endorsed by Microsoft. Microsoft, Excel and Office are trademarks of the Microsoft group of companies.
