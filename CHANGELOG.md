# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [0.3.0] — 2025-04-30

### Added
- **ExcelJS:** `setActiveSheet` — sets a default sheet on the workbook; all subsequent operations fall back to it when `sheet` is not explicitly provided
- **ExcelJS:** `getInfo` — returns workbook metadata (`sheetCount`, `sheets[]` with name, row counts, used range, state); workbook preserved in `msg._doc`
- **PptxGenJS:** `getSlideCount` — returns `{ slideCount, currentSlideIndex }` without modifying the presentation; preserved in `msg._doc`
- **All nodes:** `node.status()` visual indicator — green dot on success (shows operation name), red ring on error (shows truncated message)
- **All nodes:** `msg.info` on every success output with contextual metadata (`operation`, `sheet`, `rowsAdded`, `format`, `filePath`, `slideCount`, etc.)
- **All nodes:** improved error messages — `params.filePath`, `params.sheetName`, `params.tableName` now match the actual parameter names used at runtime
- **package.json:** `files` field — limits `npm publish` to `nodes/`, `examples/`, `README.md`, `CHANGELOG.md`, `LICENSE`
- **LICENSE:** third-party dependency table (ExcelJS, docx, PptxGenJS — all MIT)
- **CHANGELOG.md:** this file

### Fixed
- Error message for `addSheet` referred to `params.name`; corrected to `params.sheetName`
- Error message for `addTable` (ExcelJS) referred to `params.name`; corrected to `params.tableName`
- Error message for `write` (all nodes) referred to `params.path`; corrected to `params.filePath`

---

## [0.1.0] — 2025-04-30

### Added

#### ExcelJS node
- `create` — new workbook with optional creator / lastModifiedBy metadata
- `read` — load existing `.xlsx` from file path or Buffer
- `addSheet` — add worksheet with optional tab color and visibility state
- `setActiveSheet` — set a default sheet for subsequent operations (avoids repeating `sheet` on every node)
- `getInfo` — read workbook metadata: sheet names, row counts, used ranges; workbook preserved in `msg._doc`
- `addRow` / `addRows` — append rows (array or key/value object)
- `setCell` — set cell value (incl. formulas with `=` prefix) and style
- `styleRange` — apply font, fill, alignment, numFmt to a cell range
- `mergeCell` — merge a cell range
- `setColumnWidth` — set column character width and hidden flag
- `setRowHeight` — set row height (pt) and hidden flag
- `freezePanes` — freeze rows and/or columns
- `readCell` — read value, text, type, formula from a single cell; workbook preserved in `msg._doc`
- `readRange` — read a rectangular range into a 2-D array; workbook preserved in `msg._doc`
- `addTable` — insert a structured Excel table with style and filter buttons
- `conditionalFormat` — add conditional formatting rules to a range
- `protect` — password-protect a sheet (workbook-level falls back to all sheets)
- `write` — serialize to Buffer, Base64 string, or file

#### Docx node
- `create` — new document with title, author, page size (DXA), margins (DXA), default font
- `addHeading` — heading levels 1–6 with optional bold/color/size style
- `addParagraph` — plain text or mixed inline styles via `runs` array; spacing and alignment support
- `addList` — bullet or numbered list with nesting levels 0–4
- `addTable` — table with optional header row styling and per-column DXA widths
- `addImage` — inline image from file path, Buffer, or base64 data URI
- `addHeader` — page header (default / first / even page)
- `addFooter` — page footer with static text and optional automatic page numbers
- `pageBreak` — insert a hard page break
- `write` — serialize to Buffer, Base64 string, or file
- Paragraph `runs` support `hyperlink` field for inline `ExternalHyperlink` wrapping

#### PptxGenJS node
- `create` — new presentation with layout (16x9, 4x3, WIDE, or custom inches), author, company, theme fonts
- `addSlide` — add slide, optionally referencing a named master; solid or image background
- `addText` — text box with decomposed position (x/y/w/h inches) and text style fields
- `addShape` — geometric shape with decomposed position and shape style fields
- `addImage` — image from file path, URL, Buffer, or base64 data URI; sizing and hyperlink support
- `addChart` — bar, line, pie, area, scatter, bubble, donut charts with multi-series data
- `addTable` — table with per-cell options
- `addNotes` — plain-text speaker notes for a slide
- `defineMaster` — named slide master with background and static objects
- `getSlideCount` — returns `{slideCount, currentSlideIndex}`; presentation preserved in `msg._doc`
- `setLayout` — change layout after creation
- `write` — serialize to Buffer, Base64 string, or file

#### General
- Two-output pattern: port 1 = success (`msg.payload`), port 2 = error (`msg.error`)
- `node.status()` indicator on every node: green dot on success, red ring on error
- `msg.info` on every success output with contextual metadata (operation, sheet, rowsAdded, format, filePath, slideCount, …)
- `msg.params` runtime override for all panel fields; full style/position objects bypass individual fields
- `msg.operation` runtime override for the operation dropdown
- `msg.contextKey` for shared documents across parallel flow branches (flow context)
- `msg._doc` preserves the original document after `write`, `readCell`, `readRange`, `getInfo`, `getSlideCount`

