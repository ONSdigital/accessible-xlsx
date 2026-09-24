# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

`@onsvisual/accessible-xlsx` — a thin wrapper around the [Documonster Excel module](https://github.com/documonster/documonster/blob/main/src/modules/excel/README.md) that produces XLSX files meeting the [Analysis Function spreadsheet accessibility guidance](https://analysisfunction.civilservice.gov.uk/policy-store/releasing-statistics-in-spreadsheets/). Output should pass every check in the [XLSX Accessibility Checker](https://onsdigital.github.io/xlsx-accessibility-checker/). It runs both in the browser and in Node.js.

## Commands

- `npm run dev` — Vite dev server serving `index.html`, a demo page with buttons that generate single- and multi-sheet example spreadsheets. This is the main way to test changes by hand. Open the downloaded files in Excel or run them through the accessibility checker.
- `npm run build` — runs both builds below.
  - `npm run build:package` — library build from `lib/main.ts` to `dist/accessible-xlsx.{es,umd}.js` (UMD global name `accessibleXLSX`).
  - `npm run build:docs` — builds `index.html` into `docs/` (Vite `--mode docs`, relative `base`) for GitHub Pages.
- `npm run deploy` — publishes `docs/` to the `gh-pages` branch.
- Type check: `npx tsc` (`noEmit`, strict unused-locals/params checks; covers `lib/` only).

There is no test suite or linter. Formatting uses Prettier (`.prettierrc`: tabs, width 4, print width 100, no trailing commas).

`dist/` and `docs/` are gitignored. Both builds set `codeSplitting: false`, so Documonster is bundled into a single output file. That's why it is a devDependency and not a runtime dependency. Both builds also set `treeshake: false`. It works around a rolldown bug where tree-shaking splits documonster's `excel/surface/worksheet.js` into its own chunk, which breaks the UMD build. With `treeshake: false`, rolldown 1.1.x (Vite 8.1) produces a bundle that fails at runtime (`registerDuplexFrom is not defined`), so keep Vite ≥ 8.3. After changing the bundler config, check that the built bundle actually runs, e.g. by importing it in Node, not only that the build succeeds.

`Workbook.toBuffer` is called with `{ validate: false }`. documonster's self-check would otherwise flag every table ("Excel drops tables with a fully-hidden autoFilter"), because every column has `filterButton: false`. That is a false alarm: the output has been verified to open correctly in Excel.

## Architecture

All library code lives in `lib/main.ts`. It has one default export, `async accessibleXLSX(data)`, which returns a buffer from `Workbook.toBuffer`.

**Single sheet vs multi-sheet:** the mode depends on whether `data.sheets` is present.
- **Single-sheet** input is one `sheetData` object at the top level. Only the data sheet is produced.
- **Multi-sheet** input adds these sheets before the data sheets, in order: `Cover_sheet` (title plus `coverSheetContents`), `Table_of_contents` (a table of hyperlinks to each data sheet), and `Notes` (only when `notes` is non-empty).
- Data sheets are always named `"1"`, `"2"`, … and the Table of contents hyperlinks point to `#'N'!A1`, so changing sheet naming would break those links.

**Data sheet layout:** an H1 title, then the `sheetIntroText` rows (the last one gets a 40pt height and top alignment, which gives visual spacing), then one Documonster table directly underneath. The table name is `tableName` or the slugified sheet name. Filter buttons are off and the header row is bold and wrapped.

**Three input shapes for rows**, normalised to `dataRow[]` by `getDataRows`:
1. `rows` as an array of arrays (passed straight through; the fastest option).
2. `rows` as an array of objects, looked up by `column.key`, falling back to `column.heading`.
3. No `rows`, with values in each `column.values`.

**Mini text markup** in `addTextRow` (used for cover sheet contents, intro text and sheet titles). Each line is one row in column A:
- `# ` → the "Heading 1" named cell style (18pt bold)
- `## ` → "Heading 2" (14pt bold, row height 40)
- A line starting with `[` → a hyperlink parsed from `[text](url)`
- Anything else → plain text

Headings get the named cell style (registered with `Workbook.defineCellStyle`) and the matching font is also set directly on the cell.

**Column styles:** `text`, `number_with_commas`/`number_0dp` (`#,##0`), and `number_Ndp` (`#,##0.` followed by N zeros). Numeric columns get a fixed minimum width of 12. Text columns are sized to their longest value, clamped between 12 and 28.

**Notes:** a note reference like `[note_1]` inside headings or text is plain text. The library does not process it. It only renders the `notes` array as a table on the Notes sheet.

Several documonster calls have `// @ts-ignore` because documonster doesn't export a `WorksheetData` type.

## Releases

A version bump is its own commit (e.g. `v0.2.1 - ...`). The README's unpkg script URLs include a version number (currently `@0.2.1`), so update them when you release.
