# node-red-contrib-officedocs

[![npm version](https://img.shields.io/npm/v/node-red-contrib-officedocs.svg)](https://www.npmjs.com/package/node-red-contrib-officedocs)
[![npm downloads](https://img.shields.io/npm/dm/node-red-contrib-officedocs.svg)](https://www.npmjs.com/package/node-red-contrib-officedocs)
[![Node.js ≥16](https://img.shields.io/badge/node-%3E%3D16-brightgreen.svg)](https://nodejs.org)
[![Node-RED ≥2](https://img.shields.io/badge/node--red-%3E%3D2.0-red.svg)](https://nodered.org)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

[![Ko-fi](https://img.shields.io/badge/Buy%20me%20a%20coffee-%23FF5E5B?logo=ko-fi&logoColor=white&style=flat)](https://ko-fi.com/davidetravaglini)

Node-RED nodes to create and manipulate Office documents: **Excel** (.xlsx), **Word** (.docx) and **PowerPoint** (.pptx) — without Microsoft Office.

| Node | Library | Format |
|------|---------|--------|
| `exceljs` | [ExcelJS](https://exceljs.readthedocs.io/) | `.xlsx` |
| `docx`    | [docx](https://docx.js.org/)              | `.docx` |
| `pptx`    | [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) | `.pptx` |

---

## Installation

```bash
# from the Node-RED user directory (usually ~/.node-red)
npm install node-red-contrib-officedocs
```

Or via the Node-RED **Manage Palette** UI → search `node-red-contrib-officedocs`.

### Requirements

- Node.js ≥ 16
- Node-RED ≥ 2.0

---

## Concepts

### Document travels in `msg.payload`

Nodes are designed to be **chained**: the output of one node is the input of the next.

```
[inject]
  → [exceljs: create]
  → [exceljs: addSheet]
  → [exceljs: addRows]
  → [exceljs: write]      ← msg.payload = Buffer / path / base64
  → [http response]
```

### Two outputs — success and error

Every node has **two output ports**:
- **Port 1 — success**: `msg.payload` contains the updated document.
- **Port 2 — error**: `msg.error = { message, operation, params, stack }`.

### Panel fields vs `msg.params`

Every parameter can be configured in two ways (evaluated in priority order):

1. **`msg.params.fieldName`** — passed at runtime, highest priority.
2. **Panel field** — fixed value, TypedInput (str / num / bool / msg / env / JSONata), configured in the node editor.

Style and position parameters that consist of multiple sub-fields (color, bold, size, x, y, w, h…) are **decomposed into individual panel fields**. You can still pass the full object at runtime via `msg.params` to bypass all individual fields at once:

```javascript
// panel fields are ignored when the full object is passed at runtime
msg.params = {
  style:     { bold: true, color: "2E75B6", size: 32 },   // docx heading
  position:  { x: 0.5, y: 0.5, w: 9, h: 1.5 },           // pptx
  textStyle: { fontSize: 36, bold: true, color: "FFFFFF" } // pptx
};
```

### Runtime operation override

```javascript
msg.operation = "addRow";          // overrides the operation dropdown
msg.params = { sheet: "Sales", rowValues: ["Q1", 48500] };
```

### Flow context (shared document between parallel branches)

```javascript
msg.contextKey = "myReport";  // all branches share the same workbook
```

The document is stored/retrieved from flow context rather than `msg.payload`. Clean up after use:

```javascript
flow.set("myReport", null);
```

### Write formats

| `format` | `msg.payload` after write |
|----------|--------------------------|
| `buffer` | Node.js `Buffer` |
| `base64` | Base64 string |
| `file`   | Absolute path of the written file |

---

## Unit reference

| Unit | Used in | Conversion |
|------|---------|-----------|
| **DXA (twips)** | docx page size, margins, column widths, spacing | 1 inch = 1440 DXA · 1 cm ≈ 567 DXA |
| **Half-points** | docx font size | 24 = 12 pt · 28 = 14 pt · 32 = 16 pt |
| **Points (pt)** | ExcelJS font size, pptx font size, pptx line width | — |
| **Inches** | pptx positions (x, y, w, h) | — |
| **Pixels** | docx image width/height (at 96 dpi) | — |
| **ARGB hex** | ExcelJS colors | 8 digits, no `#` — first 2 = alpha (`FF` = opaque) |
| **RGB hex** | docx and pptx colors | 6 digits, no `#` |

Common docx page sizes in DXA:

| Format | Width | Height |
|--------|-------|--------|
| US Letter (portrait) | 12240 | 15840 |
| A4 (portrait)        | 11906 | 16838 |
| A4 (landscape)       | 16838 | 11906 |

---

## ExcelJS Node

Creates and manipulates Excel workbooks.

### Operations

#### `create`

Creates a new empty workbook.

```javascript
msg.params = {
  creator:        "Node-RED",
  lastModifiedBy: "Node-RED"
};
```

#### `read`

Loads an existing `.xlsx` from disk or Buffer.

```javascript
msg.params = {
  source: "/path/to/file.xlsx"  // string path or Buffer
};
```

#### `addSheet`

Adds a new worksheet.

```javascript
msg.params = {
  sheetName: "Sales 2024",   // required — error if already exists
  tabColor:  "FF0000",       // optional — ARGB hex, no #
  sheetState: "visible"      // "visible" | "hidden" | "veryHidden"
};
```

#### `addRow`

Appends a single row.

```javascript
msg.params = {
  sheet:     "Sales 2024",
  rowValues: ["Q1", "Rossi", 48500]   // array or { colName: value } object
};
```

#### `addRows`

Appends multiple rows in one call (more efficient than repeated `addRow`).

```javascript
msg.params = {
  sheet:    "Sales 2024",
  rowsData: [
    ["Q1", "Rossi",   48500],
    ["Q2", "Bianchi", 52000]
  ]
};
```

#### `setCell`

Sets the value and/or style of a single cell.

```javascript
msg.params = {
  sheet:     "Sales 2024",
  cell:      "B3",       // A1 notation or { row: 3, col: 2 }
  cellValue: 48500,      // prefix "=" for formulas: "=SUM(B1:B2)"
  // individual panel style fields: styleFontBold, styleFontSize, styleFontColor,
  // styleFillColor, styleNumFmt, styleAlignment
  // OR pass the full object to override all of them at once:
  style: {
    font:      { bold: true, size: 12, color: { argb: "FF0000FF" } },
    fill:      { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } },
    alignment: { horizontal: "center" },
    numFmt:    "#,##0.00"
  }
};
```

> **Color format:** ExcelJS uses **ARGB** (8 hex digits, no `#`). `FF000000` = opaque black, `FFFF0000` = opaque red, `FFFFFF00` = opaque yellow.

#### `styleRange`

Applies a style to a cell range (same style fields as `setCell`).

```javascript
msg.params = {
  sheet: "Sales 2024",
  range: "A1:D1",        // A1:B2 notation
  style: { font: { bold: true, size: 12 } }
};
```

#### `addTable`

Inserts a structured Excel table.

```javascript
msg.params = {
  sheet:          "Sales 2024",
  tableName:      "SalesTable",   // unique within the workbook
  tableRef:       "A1",
  tableStyle:     "TableStyleMedium9",
  showRowStripes: true,
  columns: [
    { name: "Quarter", filterButton: true },
    { name: "Agent",   filterButton: true },
    { name: "Amount",  filterButton: true, totalsRowFunction: "sum" }
  ],
  tableRows: [
    ["Q1", "Rossi",   48500],
    ["Q2", "Bianchi", 52000]
  ]
};
```

#### `conditionalFormat`

Adds conditional formatting rules to a range.

```javascript
msg.params = {
  sheet:   "Sales 2024",
  range:   "C2:C100",
  cfRules: [
    {
      type: "cellIs", operator: "greaterThan", formulae: [50000],
      style: { fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } } }
    }
  ]
};
```

Common rule types: `cellIs`, `expression`, `colorScale`, `dataBar`, `iconSet`, `top10`.

#### `mergeCell`

Merges a range of cells. Merged cells display only the top-left cell's value.

```javascript
msg.params = {
  sheet: "Sales 2024",
  range: "A1:D1"   // A1 notation range
};
```

> Merging a range that already contains data preserves only the top-left value; all other cell values in the range are discarded.

#### `setColumnWidth`

Sets width and optionally hides a column.

```javascript
msg.params = {
  sheet:     "Sales 2024",
  colRef:    "B",    // column letter or 1-based number
  colWidth:  20,     // character-width units (ExcelJS default ≈ 8.38)
  colHidden: false
};
```

#### `setRowHeight`

Sets height and optionally hides a row.

```javascript
msg.params = {
  sheet:     "Sales 2024",
  rowNumber: 1,     // 1-based row number
  rowHeight: 30,    // points
  rowHidden: false
};
```

#### `freezePanes`

Freezes rows and/or columns so they stay visible when scrolling.

```javascript
msg.params = {
  sheet:     "Sales 2024",
  freezeRow: 1,   // number of rows to freeze from the top (0 = none)
  freezeCol: 0    // number of columns to freeze from the left (0 = none)
};
// Freeze first row only: freezeRow=1, freezeCol=0
// Freeze first column only: freezeRow=0, freezeCol=1
// Freeze both: freezeRow=1, freezeCol=1
```

#### `readCell`

Reads the value and metadata of a single cell. The workbook is preserved in `msg._doc`.

```javascript
msg.params = {
  sheet: "Sales 2024",
  cell:  "B3"      // A1 notation
};
// msg.payload → { address, value, text, type, formula }
// msg._doc    → original Workbook (chain continues via msg._doc)
```

#### `readRange`

Reads a rectangular range into a 2-D array of `{ value, text, type, formula }` objects.

```javascript
msg.params = {
  sheet: "Sales 2024",
  range: "A1:D10"
};
// msg.payload → 2-D array [ [ { value, text, type, formula }, … ], … ]
// msg._doc    → original Workbook
```

> After `readCell` / `readRange`, pass `msg._doc` back into `msg.payload` in a function node if you want to continue modifying the workbook.

#### `protect`

Protects a worksheet with an optional password.

```javascript
msg.params = {
  protectTarget: "sheet",          // "sheet" | "workbook"*
  sheet:         "Sales 2024",
  password:      "secret",         // optional — leave empty for no password
  // individual panel protection fields: protectSelectLocked, protectSelectUnlocked,
  // protectFormatCells, protectFormatColumns, protectFormatRows,
  // protectInsertRows, protectDeleteRows, protectSort, protectAutoFilter
  // OR pass the full object:
  options: {
    selectLockedCells:   true,
    selectUnlockedCells: true,
    formatCells:         false
  }
};
```

> \* `target: "workbook"` logs a warning and falls back to protecting all sheets — ExcelJS has limited workbook-level protection support.

#### `write`

Serializes the workbook.

```javascript
msg.params = {
  format:   "buffer",             // "file" | "buffer" | "base64"
  filePath: "/reports/file.xlsx"  // required only when format = "file"
};
// msg._doc     → original Workbook (available after write)
// msg.payload  → serialized output
```

---

## Docx Node

Creates Word documents from scratch.

> **Note:** The document is assembled as an in-memory tree throughout the chain. The actual `.docx` file is built only at `write` time. Editing **existing** `.docx` files is not supported by the underlying library.

### Operations

#### `create`

```javascript
msg.params = {
  title:       "Q1 Report",
  author:      "Node-RED",
  pageSize:    { width: 12240, height: 15840 },           // DXA — US Letter
  margins:     { top: 1440, right: 1440, bottom: 1440, left: 1440 },  // DXA
  defaultFont: { name: "Arial", size: 24 }                // size in half-points
};
```

#### `addHeading`

```javascript
msg.params = {
  headingText:  "Quarterly Results",
  headingLevel: 1,          // 1–6
  // individual panel style fields: headingBold, headingColor, headingFontSize
  // OR pass the full object to override all:
  style: { color: "2E75B6", bold: true, size: 32 }
};
```

Color is a 6-digit RGB hex without `#`. Font size is in half-points.

#### `addParagraph`

```javascript
// Simple text with style
msg.params = {
  paraText:      "Hello world.",
  paraAlignment: "left",  // "left" | "center" | "right" | "justified" | "both"
  // individual panel style fields: paraBold, paraItalic, paraUnderline,
  // paraColor, paraFontSize, paraFont
  // OR pass the full object:
  style: { bold: false, italic: false, color: "333333", size: 24 }
};

// Mixed inline styles via runs (takes precedence over text and style fields)
msg.params = {
  paraRuns: [
    { text: "Normal " },
    { text: "bold",   bold: true },
    { text: " and " },
    { text: "italic", italic: true, color: "CC0000" }
  ],
  paraAlignment: "justified",
  paraSpacing:   { before: 120, after: 120 }  // twips
};
```

Run object shape: `{ text, bold, italic, underline, color, size, font, hyperlink }`.

To add a clickable hyperlink inside a paragraph, include a `hyperlink` URL on any run:

```javascript
msg.params = {
  paraRuns: [
    { text: "See the " },
    { text: "full report", hyperlink: "https://example.com/report", color: "1657F4" },
    { text: " for details." }
  ]
};
```

#### `addList`

```javascript
msg.params = {
  listType:  "bullet",   // "bullet" | "number"
  listItems: [
    "First item",
    "Second item",
    { text: "Nested item", level: 1 },   // level 0–4
    "Third item"
  ]
};
```

#### `addTable`

```javascript
msg.params = {
  tableRows:      [
    ["Quarter", "Agent",   "Amount"],
    ["Q1",      "Rossi",   "€48,500"],
    ["Q2",      "Bianchi", "€52,000"]
  ],
  tableHeaderRow: true,
  columnWidths:   [2500, 4000, 2500],   // DXA per column
  tableBorders:   true,
  // individual header style fields: tableHeaderBold, tableHeaderFill, tableHeaderColor
  // OR pass the full object:
  headerStyle: { bold: true, fill: "2E75B6", color: "FFFFFF" }
};
```

#### `addImage`

Width and height are in **pixels** (rendered at 96 dpi).

```javascript
msg.params = {
  imageSource:  "/path/to/chart.png",  // path | Buffer | base64 data URI
  imageWidth:   400,
  imageHeight:  250,
  imageAltText: "Sales chart"
};
```

Supported image formats: PNG, JPG, GIF, BMP, SVG.

#### `addHeader`

Registers a page header. Call once per header type before `write`.

```javascript
msg.params = {
  headerText:      "Confidential — Q1 Report",
  headerType:      "default",   // "default" | "first" | "even"
  headerAlignment: "right"      // "left" | "center" | "right"
};
```

| `headerType` | When it appears |
|-------------|----------------|
| `default` | All pages (unless overridden by `first` / `even`) |
| `first`   | First page only |
| `even`    | Even-numbered pages (requires *Even & Odd Headers* in Word) |

#### `addFooter`

Registers a page footer. Supports static text and automatic page numbers.

```javascript
// Static text footer
msg.params = {
  footerText:        "Generated by Node-RED",
  footerType:        "default",
  footerAlignment:   "center",
  footerShowPageNum: false
};

// Footer with automatic page numbers ("Generated by Node-RED   Page 1 of 5")
msg.params = {
  footerText:        "Generated by Node-RED   ",
  footerAlignment:   "center",
  footerShowPageNum: true
};
```

When `footerShowPageNum` is `true`, Word field codes for *Page X of Y* are appended after the static text.

#### `pageBreak`

Inserts a page break. No `params` needed.

#### `write`

```javascript
msg.params = {
  format:   "file",
  filePath: "/reports/report.docx"
};
```

---

## PptxGenJS Node

Creates PowerPoint presentations.

> The **current slide** is tracked internally (last added slide). Content operations target the current slide unless `slideIndex` (0-based) is specified.

### Operations

#### `create`

```javascript
msg.params = {
  layout:  "LAYOUT_16x9",   // "LAYOUT_16x9" | "LAYOUT_4x3" | "LAYOUT_WIDE"
                             // or custom via: { width: 12, height: 6.75 } (inches)
  author:  "Node-RED",
  company: "Acme Corp",
  theme:   { headFontFace: "Calibri", bodyFontFace: "Calibri" }
};
```

| Layout | Dimensions |
|--------|-----------|
| `LAYOUT_16x9` | 10″ × 5.625″ |
| `LAYOUT_4x3`  | 10″ × 7.5″  |
| `LAYOUT_WIDE` | 13.33″ × 7.5″ |

#### `addSlide`

```javascript
msg.params = {
  masterName: null,      // optional — registered master slide name
  bgColor:    "003366"   // optional — 6-digit hex background color
};
// The new slide becomes the current slide
```

> For image backgrounds pass `msg.params.bgColor = { path: "/img/bg.png" }` — an object is forwarded directly to PptxGenJS.

#### `addText`

Position and style can be set as individual panel fields or passed as full objects via `msg.params`:

```javascript
// Individual fields approach (can be set in the panel)
msg.params = {
  text:        "Slide title",
  slideIndex:  null,         // null = current slide
  posX: 0.5,  posY: 0.5,   // inches
  posW: 9,    posH: 1.5,
  textFontSize: 36,
  textBold:     true,
  textColor:    "FFFFFF",    // 6-digit hex, no #
  textAlign:    "left",      // "left" | "center" | "right" | "justify"
  textFontFace: "Calibri",
  textValign:   "middle",    // "top" | "middle" | "bottom"
  textWrap:     true
};

// Full-object approach (overrides all individual fields)
msg.params = {
  text:      "Slide title",
  position:  { x: 0.5, y: 0.5, w: 9, h: 1.5 },
  textStyle: { fontSize: 36, bold: true, color: "FFFFFF", align: "left", fontFace: "Calibri" }
};
```

#### `addShape`

```javascript
// Individual fields approach
msg.params = {
  shapeType:      "rect",   // rect | ellipse | triangle | roundRect | …
  posX: 1, posY: 1, posW: 4, posH: 2,
  shapeFillColor: "4472C4",
  shapeLineColor: "002060",
  shapeLineWidth: 1
};

// Full-object approach
msg.params = {
  shapeType:  "rect",
  position:   { x: 1, y: 1, w: 4, h: 2 },
  shapeStyle: { fill: { color: "4472C4" }, line: { color: "002060", width: 1 } }
};
```

#### `addImage`

```javascript
msg.params = {
  imageSource: "/path/to/logo.png",   // path | HTTP(S) URL | Buffer | base64 data URI
  imageType:   "png",                 // required when source is a Buffer
  posX: 0.5, posY: 1.5, posW: 4, posH: 3,
  sizing:    { type: "contain" },
  hyperlink: { url: "https://example.com", tooltip: "Visit site" }
};
```

#### `addChart`

```javascript
msg.params = {
  chartType: "bar",   // "bar" | "line" | "pie" | "area" | "scatter" | "bubble" | "donut"
  posX: 0.5, posY: 1.5, posW: 8, posH: 4.5,
  chartData: [
    {
      name:   "Sales 2024",
      labels: ["Q1", "Q2", "Q3", "Q4"],
      values: [48500, 52000, 61000, 58000]
    }
  ],
  chartOptions: {
    showLegend: true, legendPos: "b",
    showTitle:  true, title: "Sales 2024"
  }
};
```

Multiple series: add more objects to the `chartData` array.

#### `addTable`

```javascript
msg.params = {
  tableRows: [
    [
      { text: "Quarter", options: { bold: true, fill: "003366", color: "FFFFFF" } },
      { text: "Amount",  options: { bold: true, fill: "003366", color: "FFFFFF" } }
    ],
    [{ text: "Q1" }, { text: "€48,500" }]
  ],
  posX: 0.5, posY: 2, posW: 8,
  tableOptions: { colW: [3, 3], fontSize: 14 }
};
```

#### `addNotes`

Adds plain-text speaker notes to a slide. Notes appear in PowerPoint Presenter View but are not rendered on the slide itself.

```javascript
msg.params = {
  notes:      "Remind the audience about the Q3 dip — supply chain disruption.",
  slideIndex: null   // null = current slide; 0-based index to target a specific slide
};
```

#### `defineMaster`

Defines a named slide master (background color or image, fixed objects on every slide). **Must be called before any `addSlide` that references the master by name.**

```javascript
msg.params = {
  masterTitle:   "CORP_MASTER",   // unique name — referenced in addSlide.masterName
  masterBkg:     "1F3864",        // 6-digit hex solid background
  // OR an image background via msg.params at runtime:
  // masterBkg: { path: "/assets/bg.png" }
  masterObjects: [
    {
      type: "text",
      text: "CONFIDENTIAL",
      options: { x: 0.1, y: 5.3, w: 2, h: 0.3, fontSize: 10, color: "AAAAAA" }
    }
  ]
};
```

After defining the master, reference it in `addSlide`:

```javascript
msg.params = { masterName: "CORP_MASTER" };
```

#### `setLayout`

```javascript
msg.params = { layout: "LAYOUT_4x3" };
```

#### `write`

```javascript
msg.params = {
  format:   "base64",              // "file" | "buffer" | "base64"
  filePath: "/reports/deck.pptx"   // required only when format = "file"
};
```

---

## msg contract

### Input

| Property | Type | Required | Description |
|----------|------|----------|-------------|
| `msg.payload` | Workbook / doc object / null | Yes (except `create`) | Document to operate on. |
| `msg.operation` | string | No | Runtime override of the panel operation. |
| `msg.params` | object | Depends | Operation parameters — override individual panel fields. Full style/position objects bypass individual fields entirely. |
| `msg.contextKey` | string | No | Flow-context key for shared documents across parallel branches. |

### Output — port 1 (success)

| Property | Type | Description |
|----------|------|-------------|
| `msg.payload` | document / Buffer / string | Updated document, or serialized output after `write`. |
| `msg._doc` | document | Original document before `write` — useful for post-save processing. |

### Output — port 2 (error)

| Property | Type | Description |
|----------|------|-------------|
| `msg.error` | object | `{ message, operation, params, stack }` |
| `msg.payload` | any | Original payload preserved for debugging. |

---

## Flow examples

### Excel — HTTP endpoint returning a .xlsx

```
[http in: GET /report]
  → [exceljs: create]
  → [function: set params]
  → [exceljs: addSheet]
  → [exceljs: addRows]
  → [exceljs: write]         format: buffer
  → [function: set headers]  Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
  → [http response]
```

### Word — generate from JSON body

```
[http in: POST /generate-doc]
  → [docx: create]
  → [docx: addHeading]    headingText = msg.payload.title
  → [docx: addParagraph]  paraText = msg.payload.body
  → [docx: addTable]      msg.params.tableRows = msg.payload.tableData
  → [docx: write]         format: base64
  → [http response]
```

### PowerPoint — live KPI presentation

```
[inject / http in]
  → [http request: fetch KPIs]
  → [pptx: create]     layout: LAYOUT_16x9
  → [pptx: addSlide]   bgColor: 003366
  → [pptx: addText]    title with data
  → [pptx: addSlide]
  → [pptx: addChart]   time series data
  → [pptx: write]      format: file, filePath: /reports/kpi.pptx
  → [email: send attachment]
```

Import ready-to-use example flows from the [`examples/`](./examples/) folder in this package.

---

## Compatibility

| Component | Version |
|-----------|---------|
| Node.js   | ≥ 16.x  |
| Node-RED  | ≥ 2.x   |
| ExcelJS   | ^4.4.0  |
| docx      | ^8.5.0  |
| PptxGenJS | ^3.12.0 |

### Generated file compatibility

| Node | Compatible with |
|------|----------------|
| exceljs | Excel 2010+, LibreOffice Calc 6+, Google Sheets |
| docx    | Word 2013+, LibreOffice Writer 6+, Google Docs  |
| pptx    | PowerPoint 2016+, LibreOffice Impress 6+, Google Slides |

---

## Known limitations

| Node | Limitation |
|------|-----------|
| `docx` | Reading / editing existing `.docx` files is **not supported** — the `docx` library is write-only. |
| `pptx` | Reading existing `.pptx` files is **not supported** — PptxGenJS is write-only. |
| `exceljs` | `protect: workbook` has limited support in ExcelJS; the node warns and falls back to protecting all sheets individually. |
| `exceljs` | Auto-fit column widths is not natively supported by ExcelJS — column widths must be set explicitly. |
| `pptx` | Animated slide transitions are not supported by PptxGenJS. |

### Known issues

- **`pptx` / `addSlide` background image:** the `bgColor` panel field only accepts a hex color string. To use an image background, pass the full object at runtime: `msg.params.bgColor = { path: "/img/bg.png" }`.
- **`exceljs` / `setCell` dates:** pass `Date` objects via `msg.params.cellValue` — the panel's TypedInput field returns a string, which ExcelJS stores as text rather than a date.

---

## Roadmap

The following features are candidates for future releases, roughly ordered by implementation effort.

### Completed ✅
- **ExcelJS:** `mergeCell` — merge a range of cells
- **ExcelJS:** `setColumnWidth` / `setRowHeight` — column/row sizing and visibility
- **ExcelJS:** `freezePanes` — freeze rows and columns
- **ExcelJS:** `readCell` / `readRange` — extract cell values from a loaded workbook
- **ExcelJS:** `setActiveSheet` — set a default sheet, avoids repeating `sheet` on every node
- **ExcelJS:** `getInfo` — workbook metadata (sheet names, row counts, used ranges)
- **PptxGenJS:** `addNotes` — speaker notes per slide
- **PptxGenJS:** `defineMaster` — named slide master with background and objects
- **PptxGenJS:** `getSlideCount` — slide count without modifying the presentation
- **Docx:** `addHeader` / `addFooter` — page headers and footers with optional page numbers
- **Docx:** hyperlinks inside paragraph runs (`ExternalHyperlink`)
- **All nodes:** `node.status()` visual indicator and `msg.info` contextual metadata on every output

### Low effort
- **ExcelJS:** `deleteSheet` / `renameSheet` — remove or rename a worksheet; useful when working from `.xlsx` templates
- **ExcelJS:** `autoFilter` — add dropdown filter buttons to a range
- **ExcelJS:** `addImage` — embed an image into a worksheet (logos, charts exported as PNG)
- **PptxGenJS:** `deleteSlide` / `reorderSlide` — remove or move slides; needed for template-based generation
- **Docx:** `addSection` — section break with independent properties (e.g. a landscape page inside a portrait document)

### Medium effort
- **All nodes:** batch mode — accept an array of operations in `msg.payload` and execute them in sequence inside a single node, reducing flow complexity for straightforward pipelines:
  ```javascript
  msg.payload = [
    { op: 'addSheet', sheetName: 'Sales' },
    { op: 'addRows',  rowsData: [...] },
    { op: 'write',    format: 'buffer' }
  ];
  ```
- **ExcelJS:** template mode — open an existing `.xlsx` as base and only fill in data, preserving branding (borders, logos, styles)
- **Docx:** `addTOC` — automatic table of contents built from heading levels

### Higher effort / library limitations
- **ExcelJS:** chart support (complex API, many chart types)
- **Docx:** table of contents with page numbers (requires Word field codes)
- **All nodes:** automated test suite (Jest/Mocha) covering the full create → add → write chain — prerequisite for `1.0.0`
- **All nodes:** GitHub Actions CI running tests on Node.js 18 / 20 / 22

---

## License

MIT © [Davide Travaglini](mailto:davide.travaglini@cyberevolution.it)
