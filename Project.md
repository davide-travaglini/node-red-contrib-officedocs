# node-red-contrib-officedocs — Specifica di progetto

---

## Indice

1. [Panoramica](#1-panoramica)
2. [Architettura generale](#2-architettura-generale)
3. [Contratto msg](#3-contratto-msg)
4. [Gestione errori](#4-gestione-errori)
5. [Flow context](#5-flow-context)
6. [Formati di output (write)](#6-formati-di-output-write)
7. [Nodo ExcelJS](#7-nodo-exceljs)
8. [Nodo Docx](#8-nodo-docx)
9. [Nodo PptxGenJS](#9-nodo-pptxgenjs)
10. [Pannelli di configurazione UI](#10-pannelli-di-configurazione-ui)
11. [Struttura del package](#11-struttura-del-package)
12. [Dipendenze e compatibilità](#12-dipendenze-e-compatibilità)
13. [Comportamenti edge case](#13-comportamenti-edge-case)
14. [Pattern di flusso consigliati](#14-pattern-di-flusso-consigliati)

---

## 1. Panoramica

`node-red-contrib-officedocs` è un package Node-RED che espone tre nodi per la creazione e manipolazione di documenti Office:

| Nodo | Libreria sottostante | Formato |
|------|----------------------|---------|
| `exceljs` | ExcelJS `^4.x` | `.xlsx` |
| `docx` | docx `^8.x` | `.docx` |
| `pptx` | PptxGenJS `^3.x` | `.pptx` |

**Principi di design:**

- Il documento viaggia in `msg.payload` tra i nodi (stateless, flow-oriented).
- Un nodo per libreria con operazione selezionabile nel pannello o sovrascrivibile a runtime.
- Porta 2 dedicata agli errori, stile Node-RED standard.
- Supporto opzionale al flow context per documenti condivisi tra branch paralleli.
- Operazione `write` con output in quattro formati: file su disco, Buffer, Base64, Stream.

---

## 2. Architettura generale

```
[inject / http-in / altro]
        │
        │  msg.payload = null (primo nodo) o Workbook/Document/Presentation
        │  msg.operation = "create" | "addRow" | "write" | …
        │  msg.params   = { …parametri specifici dell'operazione… }
        ▼
┌─────────────────────┐
│   Nodo officedocs   │  ← ExcelJS | Docx | Pptx
│  (operazione scelta)│
└──────┬──────────────┘
       │ porta 1 (success)     msg.payload = documento aggiornato
       │                       msg._doc    = copia prima di write
       ├─────────────────────► nodo successivo
       │
       │ porta 2 (error)       msg.error = { message, operation, params, stack }
       └─────────────────────► handler errori
```

I nodi sono progettati per essere **concatenati**: l'output di un nodo è l'input del successivo senza nodi `function` intermedi per la gestione del documento.

---

## 3. Contratto msg

### Proprietà in ingresso

| Proprietà | Tipo | Obbligatorio | Descrizione |
|-----------|------|--------------|-------------|
| `msg.payload` | `Workbook \| Document \| Presentation \| null` | Sì (eccetto `create`) | Il documento su cui operare. `null` o assente per l'operazione `create`. |
| `msg.operation` | `string` | No | Override runtime dell'operazione impostata nel pannello. Se assente usa il valore del pannello. |
| `msg.params` | `object` | Dipende | Parametri specifici dell'operazione (vedi sezioni per nodo). |
| `msg.topic` | `string` | No | Metadati liberi propagati inalterati (es. nome file, sheet corrente). |
| `msg.contextKey` | `string` | No | Chiave nel flow context. Se valorizzata, il documento viene caricato/salvato nel contesto invece che solo in `msg.payload`. |

### Proprietà in uscita (porta 1 — success)

| Proprietà | Tipo | Descrizione |
|-----------|------|-------------|
| `msg.payload` | `Workbook \| Document \| Presentation \| string \| Buffer \| Stream` | Documento aggiornato. Dopo `write` contiene il file di output nel formato richiesto. |
| `msg._doc` | `Workbook \| Document \| Presentation` | Copia del documento originale prima di `write`. Permette operazioni post-salvataggio senza rigenerare. |
| `msg.topic` | `string` | Propagato inalterato. |

### Proprietà in uscita (porta 2 — error)

| Proprietà | Tipo | Descrizione |
|-----------|------|-------------|
| `msg.error` | `object` | Oggetto errore strutturato (vedi sezione 4). |
| `msg.payload` | `any` | Payload originale preservato per debug. |

### Esempi pratici

**Creare un workbook e aggiungere una riga:**

```javascript
// Nodo 1 — create
msg.operation = "create";
msg.params = { creator: "Node-RED", created: new Date() };

// Nodo 2 — addSheet
msg.operation = "addSheet";
msg.params = { name: "Vendite 2024" };

// Nodo 3 — addRow
msg.operation = "addRow";
msg.params = {
  sheet: "Vendite 2024",
  values: ["2024-Q1", "Luca Rossi", 48500]
};

// Nodo 4 — write
msg.operation = "write";
msg.params = {
  format: "buffer"   // restituisce Buffer in msg.payload
};
```

**Documento condiviso tra branch (flow context):**

```javascript
// Entrambi i branch usano la stessa chiave
msg.contextKey = "reportMensile";

// Branch A: msg.operation = "addRow"; msg.params = { sheet: "Entrate", ... }
// Branch B: msg.operation = "addRow"; msg.params = { sheet: "Uscite",  ... }
// Join: msg.operation = "write"; msg.params = { format: "file", path: "/reports/report.xlsx" }
```

---

## 4. Gestione errori

### Comportamento

Ogni nodo ha **due uscite**:

- **Porta 1** — successo: il messaggio prosegue con `msg.payload` aggiornato.
- **Porta 2** — errore: il messaggio esce con `msg.error` valorizzato; `msg.payload` originale è preservato.

In caso di errore il nodo **non emette nulla sulla porta 1**.

### Struttura di `msg.error`

```javascript
msg.error = {
  message:   "Sheet 'Vendite' non trovato",   // stringa leggibile
  operation: "addRow",                          // operazione che ha fallito
  params:    { sheet: "Vendite", values: [] },  // parametri usati
  stack:     "Error: ...\n  at ..."             // stack trace completo
};
```

### Implementazione Node-RED

```javascript
node.send([
  [successMsg],  // porta 1 — success
  null           // porta 2 — nessun errore
]);

// oppure in caso di errore:
node.send([
  null,
  [errorMsg]     // porta 2 — errore
]);
```

### Errori attesi per tipo

| Condizione | Operazione | Messaggio tipo |
|------------|------------|----------------|
| `msg.payload` non è un documento valido | qualsiasi (tranne `create`) | `"msg.payload non è un Workbook valido"` |
| Sheet non trovato | `addRow`, `setCell`, ecc. | `"Sheet 'X' non trovato nel workbook"` |
| Path non scrivibile | `write` (format: file) | `"Impossibile scrivere su '/path/to/file': EACCES"` |
| Parametri mancanti | qualsiasi | `"params.sheet è obbligatorio per addRow"` |
| Operazione sconosciuta | qualsiasi | `"Operazione 'fooBar' non supportata"` |

---

## 5. Flow context

### Scopo

Permette a **branch paralleli** di operare sullo stesso documento. Senza flow context, un join di due branch che modificano `msg.payload` causerebbe conflitti perché ogni branch porta una copia diversa del documento.

### Logica di caricamento (priorità)

```javascript
// All'ingresso del nodo:
const doc = msg.contextKey
  ? flow.get(msg.contextKey) ?? msg.payload
  : msg.payload;
```

Se `msg.contextKey` è valorizzato ma il context è vuoto (prima esecuzione), il nodo fa fallback su `msg.payload`.

### Logica di salvataggio

```javascript
// Dopo ogni operazione andata a buon fine:
if (msg.contextKey) {
  flow.set(msg.contextKey, updatedDoc);
}
msg.payload = updatedDoc;
```

### Pulizia del context

Il flow context **non viene mai pulito automaticamente** dal nodo. È responsabilità del flusso chiamare esplicitamente:

```javascript
// In un nodo function dopo la fine del lavoro:
flow.set("reportMensile", null);
```

### Limitazioni

- Il flow context di Node-RED è in-memory per default. Se si usa un context store persistente (es. `localfilesystem`), oggetti complessi come Workbook potrebbero non serializzarsi correttamente. In quel caso usare il context store `memory`.
- Non thread-safe per scritture concorrenti reali: Node-RED è single-thread, quindi i branch paralleli in realtà si alternano nell'event loop. Il rischio di race condition è basso ma presente con subflow asincroni multipli.

---

## 6. Formati di output (`write`)

L'operazione `write` è condivisa da tutti e tre i nodi. Il formato si specifica in `msg.params.format`.

### `file`

```javascript
msg.params = {
  format: "file",
  path: "/var/reports/report_2024.xlsx"  // assoluto, relativo, UNC (\\server\share\file.xlsx), SMB montato
};
// msg.payload dopo write = stringa con il path assoluto del file scritto
```

Supporto path:
- Path assoluti Unix/Windows: `/home/user/report.xlsx`, `C:\Reports\report.xlsx`
- Path relativi: risolti rispetto a `process.cwd()` del processo Node-RED
- UNC: `\\\\server\\share\\report.xlsx` (su Windows o con Samba montato)
- Mount SMB su Linux: `/mnt/nas/report.xlsx`

### `buffer`

```javascript
msg.params = { format: "buffer" };
// msg.payload dopo write = Buffer Node.js
// Pronto per: HTTP response, upload S3/GCS, pipe a WriteStream
```

### `base64`

```javascript
msg.params = { format: "base64" };
// msg.payload dopo write = stringa base64
// Utile per: data URI, invio via MQTT, WebSocket, email attachment
```

### `stream`

```javascript
msg.params = { format: "stream" };
// msg.payload dopo write = PassThrough stream
// Pipeable direttamente a fs.createWriteStream() o a http.ServerResponse
```

---

## 7. Nodo ExcelJS

### Dipendenza

```
exceljs ^4.4.0
```

### Operazioni

---

#### `create`

Crea un nuovo Workbook vuoto.

```javascript
msg.operation = "create";
msg.params = {
  creator:      "string",   // opzionale — autore del file
  lastModified: Date,       // opzionale — default: now
  useSharedStrings: true,   // opzionale — default: true (ottimizza dimensioni)
  useStyles:    true        // opzionale — default: true
};
// msg.payload → nuovo Workbook
```

---

#### `read`

Legge un file `.xlsx` esistente da disco o da Buffer.

```javascript
msg.operation = "read";
msg.params = {
  source: "/path/to/file.xlsx"  // stringa (path) | Buffer | Stream
};
// msg.payload → Workbook caricato
```

---

#### `addSheet`

Aggiunge un nuovo worksheet al workbook.

```javascript
msg.operation = "addSheet";
msg.params = {
  name:       "Vendite 2024",  // obbligatorio
  tabColor:   "FF0000",        // opzionale — colore tab (hex senza #)
  state:      "visible",       // opzionale — "visible" | "hidden" | "veryHidden"
  properties: {}               // opzionale — ExcelJS WorksheetProperties
};
// msg.payload → Workbook con il nuovo sheet
```

**Edge case:** se esiste già un sheet con lo stesso nome, il nodo lancia errore sulla porta 2. Non viene rinominato silenziosamente.

---

#### `addRow`

Aggiunge una singola riga a un worksheet.

```javascript
msg.operation = "addRow";
msg.params = {
  sheet:  "Vendite 2024",           // obbligatorio — nome del sheet
  values: ["2024-Q1", "Rossi", 48500]  // array o oggetto { colName: value }
};
// msg.payload → Workbook aggiornato
```

**Edge case:** se `params.sheet` non esiste, il nodo crea automaticamente il sheet (comportamento permissivo, configurabile nel pannello con un toggle "Auto-crea sheet mancante").

---

#### `addRows`

Aggiunge multiple righe in una sola operazione (più efficiente di chiamate `addRow` ripetute).

```javascript
msg.operation = "addRows";
msg.params = {
  sheet:  "Vendite 2024",
  rows: [
    ["2024-Q1", "Rossi", 48500],
    ["2024-Q2", "Bianchi", 52000],
    ["2024-Q3", "Verdi",  61000]
  ]
};
```

---

#### `setCell`

Imposta il valore (e opzionalmente lo stile) di una singola cella.

```javascript
msg.operation = "setCell";
msg.params = {
  sheet:   "Vendite 2024",
  cell:    "B3",               // notazione A1 o { row: 3, col: 2 }
  value:   48500,              // valore o formula es. "=SUM(B1:B2)"
  style: {                     // opzionale — ExcelJS CellStyle
    font:      { bold: true, color: { argb: "FF0000FF" } },
    fill:      { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } },
    alignment: { horizontal: "center" },
    numFmt:    "#,##0.00"
  }
};
```

---

#### `styleRange`

Applica uno stile a un range di celle.

```javascript
msg.operation = "styleRange";
msg.params = {
  sheet:  "Vendite 2024",
  range:  "A1:D1",           // notazione A1:B2
  style: {
    font: { bold: true, size: 12 },
    fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF4472C4" } }
  }
};
```

---

#### `addTable`

Aggiunge una tabella Excel strutturata (con intestazioni e stile nativo Excel).

```javascript
msg.operation = "addTable";
msg.params = {
  sheet:    "Vendite 2024",
  name:     "TabellaVendite",   // nome univoco della tabella
  ref:      "A1",               // cella di ancoraggio top-left
  style:    "TableStyleMedium9", // stile Excel nativo
  showRowStripes: true,
  columns: [
    { name: "Trimestre", filterButton: true },
    { name: "Agente",    filterButton: true },
    { name: "Importo",   filterButton: true, totalsRowFunction: "sum" }
  ],
  rows: [
    ["2024-Q1", "Rossi",  48500],
    ["2024-Q2", "Bianchi",52000]
  ]
};
```

---

#### `addChart`

Aggiunge un grafico a un worksheet.

```javascript
msg.operation = "addChart";
msg.params = {
  sheet:    "Vendite 2024",
  type:     "bar",           // "bar" | "line" | "pie" | "area" | "scatter" | "donut"
  title:    "Vendite per trimestre",
  dataRange: "A1:C5",       // range dati sorgente (incluse intestazioni)
  position: {
    tl: { col: 5, row: 0 },   // top-left in coordinate col/row (0-based)
    br: { col: 12, row: 15 }  // bottom-right
  }
};
```

---

#### `conditionalFormat`

Aggiunge formattazione condizionale a un range.

```javascript
msg.operation = "conditionalFormat";
msg.params = {
  sheet: "Vendite 2024",
  range: "C2:C100",
  rules: [
    {
      type:     "cellIs",
      operator: "greaterThan",
      formulae: [50000],
      style: { fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } } }
    },
    {
      type:     "colorScale",
      cfvo:     [{ type: "min" }, { type: "max" }],
      color:    [{ argb: "FFFF0000" }, { argb: "FF00FF00" }]
    }
  ]
};
```

---

#### `protect`

Protegge un worksheet o l'intero workbook con password.

```javascript
msg.operation = "protect";
msg.params = {
  target:   "sheet",           // "sheet" | "workbook"
  sheet:    "Vendite 2024",    // obbligatorio se target = "sheet"
  password: "secretPass123",   // opzionale — stringa vuota = protezione senza password
  options: {                   // opzionale — solo per target: "sheet"
    selectLockedCells:   true,
    selectUnlockedCells: true,
    formatCells:         false,
    formatRows:          false,
    insertRows:          false,
    deleteRows:          false
  }
};
```

---

#### `write`

Serializza e salva il Workbook. Vedi [sezione 6](#6-formati-di-output-write) per i dettagli sui formati.

```javascript
msg.operation = "write";
msg.params = {
  format: "buffer",    // "file" | "buffer" | "base64" | "stream"
  path:   null         // obbligatorio solo se format = "file"
};
// msg._doc    → Workbook originale (copia pre-write)
// msg.payload → output nel formato richiesto
```

---

## 8. Nodo Docx

### Dipendenza

```
docx ^8.5.0
```

### Note implementative

La libreria `docx` è costruttiva: il documento viene assemblato come albero di oggetti e serializzato in un solo passaggio con `Packer.toBuffer()`. Non è possibile modificare un `.docx` esistente con questa libreria (per quello serve un approccio XML diretto). L'operazione `read` quindi non è supportata — usare il nodo per creare documenti da zero.

### Operazioni

---

#### `create`

Crea un nuovo documento Word vuoto con opzioni di layout.

```javascript
msg.operation = "create";
msg.params = {
  title:       "Report Q1 2024",   // opzionale — metadato documento
  author:      "Node-RED",          // opzionale
  pageSize: {
    width:  12240,   // DXA — default: US Letter (8.5")
    height: 15840    // DXA — default: US Letter (11")
  },
  margins: {
    top: 1440, right: 1440, bottom: 1440, left: 1440  // DXA — default: 1 inch
  },
  defaultFont: {
    name: "Arial",
    size: 24         // half-points (24 = 12pt)
  },
  numbering: []      // opzionale — config liste (vedi addList)
};
// msg.payload → { sections: [], styles: {}, numbering: {}, meta: {} }
// (oggetto interno del nodo, non un'istanza Document diretta)
```

**Nota implementativa:** il documento viene tenuto in memoria come struttura dati mutabile (array di `sections` con `children`). L'istanza `Document` di docx viene creata solo al momento del `write`. Questo permette operazioni successive di aggiunta.

---

#### `addHeading`

Aggiunge un'intestazione.

```javascript
msg.operation = "addHeading";
msg.params = {
  text:  "Risultati del trimestre",
  level: 1,              // 1–6 (HeadingLevel.HEADING_1 … HEADING_6)
  style: {               // opzionale — override stile
    color:  "2E75B6",    // hex senza #
    bold:   true,
    size:   32           // half-points
  }
};
```

---

#### `addParagraph`

Aggiunge un paragrafo di testo.

```javascript
msg.operation = "addParagraph";
msg.params = {
  text:      "Testo del paragrafo.",
  alignment: "left",     // "left" | "center" | "right" | "justified"
  spacing: {
    before: 120,          // DXA twips prima del paragrafo
    after:  120
  },
  style: {               // opzionale
    bold:      false,
    italic:    false,
    underline: false,
    color:     "000000",
    size:      24         // 12pt
  },
  // Per testo misto (bold inline, link, ecc.) usare runs:
  runs: [
    { text: "Testo normale " },
    { text: "grassetto", bold: true },
    { text: " e " },
    { text: "corsivo", italic: true }
  ]
  // "text" e "runs" sono mutuamente esclusivi; "runs" ha precedenza
};
```

---

#### `addList`

Aggiunge un elenco puntato o numerato.

```javascript
msg.operation = "addList";
msg.params = {
  type:  "bullet",    // "bullet" | "number"
  items: [
    "Primo elemento",
    "Secondo elemento",
    { text: "Elemento annidato", level: 1 },  // livello 0 = top
    "Terzo elemento"
  ]
};
```

---

#### `addTable`

Aggiunge una tabella.

```javascript
msg.operation = "addTable";
msg.params = {
  rows: [
    ["Trimestre", "Agente",   "Importo"],   // prima riga = intestazione se headerRow: true
    ["Q1",        "Rossi",    "€48.500"],
    ["Q2",        "Bianchi",  "€52.000"]
  ],
  headerRow:    true,            // opzionale — default: false
  columnWidths: [2500, 4000, 2500],  // DXA — somma deve essere ≤ larghezza contenuto
  borders:      true,            // opzionale — default: true
  headerStyle: {                 // opzionale — stile celle intestazione
    bold: true,
    fill: "2E75B6",              // hex senza #
    color: "FFFFFF"
  }
};
```

---

#### `addImage`

Aggiunge un'immagine inline.

```javascript
msg.operation = "addImage";
msg.params = {
  source: "/path/to/image.png",  // path | Buffer | base64 string
  type:   "png",                  // "png" | "jpg" | "gif" | "bmp" | "svg"
  width:  400,                    // opzionale — punti EMU (914400 = 1 inch)
  height: 300,                    // opzionale
  altText: "Descrizione immagine" // opzionale — accessibilità
};
```

---

#### `pageBreak`

Inserisce un'interruzione di pagina.

```javascript
msg.operation = "pageBreak";
// nessun params necessario
```

---

#### `write`

Serializza il documento. Vedi [sezione 6](#6-formati-di-output-write).

```javascript
msg.operation = "write";
msg.params = {
  format: "file",
  path:   "/var/reports/report.docx"
};
```

---

## 9. Nodo PptxGenJS

### Dipendenza

```
pptxgenjs ^3.12.0
```

### Note implementative

PptxGenJS mantiene la presentazione in un'istanza `PptxGenJs`. Il nodo tiene traccia dello **slide corrente** internamente: `addText`, `addShape`, `addImage`, `addTable`, `addChart` operano sempre sull'ultimo slide aggiunto con `addSlide`. Per operare su uno slide diverso da quello corrente, è necessario passare `params.slideIndex` (0-based).

### Operazioni

---

#### `create`

Crea una nuova presentazione vuota.

```javascript
msg.operation = "create";
msg.params = {
  layout:   "LAYOUT_16x9",      // "LAYOUT_16x9" | "LAYOUT_4x3" | "LAYOUT_WIDE" | { width, height }
  author:   "Node-RED",          // opzionale
  company:  "Acme Corp",         // opzionale
  theme: {                       // opzionale
    headFontFace: "Calibri",
    bodyFontFace: "Calibri"
  }
};
// msg.payload → istanza PptxGenJs
```

---

#### `addSlide`

Aggiunge un nuovo slide e lo imposta come slide corrente.

```javascript
msg.operation = "addSlide";
msg.params = {
  masterName:  null,           // opzionale — nome master slide registrato
  background: {
    color: "FFFFFF"            // hex senza # oppure { path } per immagine
  }
};
// msg.payload → istanza PptxGenJs aggiornata
// Il nuovo slide diventa il "slide corrente" per le operazioni successive
```

---

#### `addText`

Aggiunge un elemento testo allo slide corrente (o a `slideIndex`).

```javascript
msg.operation = "addText";
msg.params = {
  text:  "Titolo della slide",   // stringa o array di TextProps per testo misto
  slideIndex: null,               // opzionale — se null usa slide corrente (0-based)
  position: {
    x: 0.5,   // pollici dal bordo sinistro
    y: 0.5,   // pollici dal bordo superiore
    w: 9,     // larghezza in pollici
    h: 1.5    // altezza in pollici
  },
  style: {
    fontSize:  36,
    bold:      true,
    color:     "003366",          // hex senza #
    align:     "left",            // "left" | "center" | "right"
    fontFace:  "Calibri",
    wrap:      true,
    valign:    "top"              // "top" | "middle" | "bottom"
  }
};
```

---

#### `addShape`

Aggiunge una forma geometrica allo slide corrente.

```javascript
msg.operation = "addShape";
msg.params = {
  type:       "rect",            // qualsiasi ShapeType di PptxGenJS
  slideIndex: null,
  position:   { x: 1, y: 1, w: 4, h: 2 },
  style: {
    fill:       { color: "4472C4" },
    line:       { color: "002060", width: 1 },
    shadow:     null              // opzionale
  }
};
```

---

#### `addImage`

Aggiunge un'immagine allo slide corrente.

```javascript
msg.operation = "addImage";
msg.params = {
  source:     "/path/to/image.png",  // path | URL | Buffer | base64 data URI
  slideIndex: null,
  position:   { x: 0.5, y: 1.5, w: 4, h: 3 },
  sizing: {
    type:   "contain",    // "contain" | "cover" | "crop"
    x:      0,
    y:      0
  },
  hyperlink:  null         // opzionale — { url: "https://…", tooltip: "…" }
};
```

---

#### `addChart`

Aggiunge un grafico allo slide corrente.

```javascript
msg.operation = "addChart";
msg.params = {
  type:       "bar",          // "bar" | "line" | "pie" | "area" | "scatter" | "bubble" | "donut"
  slideIndex: null,
  position:   { x: 0.5, y: 1.5, w: 8, h: 4.5 },
  data: [
    {
      name: "Serie 1",
      labels: ["Q1", "Q2", "Q3", "Q4"],
      values: [48500, 52000, 61000, 58000]
    }
  ],
  options: {
    showLegend:    true,
    legendPos:     "b",         // "b" | "t" | "l" | "r"
    showTitle:     true,
    title:         "Vendite 2024",
    dataLabelSize: 11
  }
};
```

---

#### `addTable`

Aggiunge una tabella allo slide corrente.

```javascript
msg.operation = "addTable";
msg.params = {
  rows: [
    [
      { text: "Trimestre", options: { bold: true, fill: "003366", color: "FFFFFF" } },
      { text: "Importo",   options: { bold: true, fill: "003366", color: "FFFFFF" } }
    ],
    [{ text: "Q1" }, { text: "€48.500" }],
    [{ text: "Q2" }, { text: "€52.000" }]
  ],
  slideIndex: null,
  position:   { x: 0.5, y: 2, w: 8, h: null },  // h: null → auto-height
  options: {
    colW:    [3, 3],        // larghezze colonne in pollici
    border:  { type: "solid", pt: 1, color: "CCCCCC" },
    fontSize: 14
  }
};
```

---

#### `setLayout`

Imposta il layout della presentazione dopo la creazione (sovrascrive il layout di `create`).

```javascript
msg.operation = "setLayout";
msg.params = {
  layout: "LAYOUT_4x3"   // "LAYOUT_16x9" | "LAYOUT_4x3" | "LAYOUT_WIDE" | { width, height }
};
```

---

#### `write`

Serializza la presentazione. Vedi [sezione 6](#6-formati-di-output-write).

```javascript
msg.operation = "write";
msg.params = {
  format: "base64"
};
```

---

## 10. Pannelli di configurazione UI

Ogni nodo espone un pannello `.html` nel Node-RED editor con i seguenti campi.

### Campi comuni a tutti i nodi

| Campo | Tipo | Descrizione |
|-------|------|-------------|
| **Nome** | text | Nome del nodo nel flow (opzionale). |
| **Operazione** | select | Operazione di default. Può essere sovrascritta da `msg.operation`. |

### Indicatore override runtime

Sotto il select dell'operazione, un testo muted: *"L'operazione può essere sovrascritta via `msg.operation` a runtime."*

### Campi specifici per operazione

I campi del pannello cambiano dinamicamente in base all'operazione selezionata, mostrando solo i parametri rilevanti con i relativi default.

**Esempio — pannello ExcelJS con operazione `addRow` selezionata:**

```
┌─────────────────────────────────────────────┐
│ Nome:       [addRow vendite            ]     │
│ Operazione: [addRow                  ▼]     │
│             L'operazione può essere          │
│             sovrascritta via msg.operation   │
│ ─────────────────────────────────────────── │
│ Sheet:      [                          ]     │
│             (usa msg.params.sheet se vuoto)  │
│ ─────────────────────────────────────────── │
│ [✓] Auto-crea sheet se non esiste            │
└─────────────────────────────────────────────┘
```

### Toggle "Auto-crea sheet mancante" (ExcelJS)

Presente nelle operazioni `addRow`, `addRows`, `setCell`, `styleRange`, `addTable`. Se abilitato (default: abilitato), il nodo crea automaticamente il sheet se non esiste invece di inviare un errore sulla porta 2.

---

## 11. Struttura del package

```
node-red-contrib-officedocs/
│
├── package.json
├── README.md
│
└── nodes/
    ├── exceljs/
    │   ├── exceljs.js          ← runtime del nodo (registrazione + handler)
    │   └── exceljs.html        ← pannello editor Node-RED
    │
    ├── docx/
    │   ├── docx.js
    │   └── docx.html
    │
    └── pptx/
        ├── pptx.js
        └── pptx.html
```

### `package.json` (chiavi rilevanti)

```json
{
  "name": "node-red-contrib-officedocs",
  "version": "1.0.0",
  "node-red": {
    "version": ">=2.0.0",
    "nodes": {
      "exceljs": "nodes/exceljs/exceljs.js",
      "docx":    "nodes/docx/docx.js",
      "pptx":    "nodes/pptx/pptx.js"
    }
  },
  "dependencies": {
    "exceljs":    "^4.4.0",
    "docx":       "^8.5.0",
    "pptxgenjs":  "^3.12.0"
  },
  "engines": {
    "node": ">=16.0.0"
  }
}
```

### Struttura interna di ogni `nodo.js`

```javascript
module.exports = function(RED) {

  function OfficeDocsExcelNode(config) {
    RED.nodes.createNode(this, config);
    const node = this;

    node.on("input", async function(msg, send, done) {
      // 1. Risolvi operazione (msg.operation ha priorità sul pannello)
      const operation = msg.operation ?? config.operation;

      // 2. Carica documento da flow context o msg.payload
      let doc = msg.contextKey
        ? node.context().flow.get(msg.contextKey) ?? msg.payload
        : msg.payload;

      try {
        // 3. Esegui operazione
        doc = await handleOperation(operation, doc, msg.params ?? {}, config);

        // 4. Salva nel flow context se richiesto
        if (msg.contextKey) {
          node.context().flow.set(msg.contextKey, doc);
        }

        // 5. Emetti sulla porta 1
        msg.payload = doc;
        send([msg, null]);
        done();

      } catch (err) {
        // 6. Emetti sulla porta 2
        msg.error = {
          message:   err.message,
          operation: operation,
          params:    msg.params,
          stack:     err.stack
        };
        send([null, msg]);
        done();
      }
    });
  }

  RED.nodes.registerType("exceljs", OfficeDocsExcelNode);
};
```

---

## 12. Dipendenze e compatibilità

### Versioni Node.js

| Node.js | Supportato |
|---------|------------|
| 16.x LTS | Sì |
| 18.x LTS | Sì (raccomandato) |
| 20.x LTS | Sì |
| < 16 | No |

### Versioni Node-RED

| Node-RED | Supportato |
|----------|------------|
| 2.x | Sì |
| 3.x | Sì (raccomandato) |
| 4.x | Sì |

### Librerie

| Libreria | Versione | Licenza | Note |
|----------|----------|---------|------|
| `exceljs` | `^4.4.0` | MIT | Supporta .xlsx; non .xls legacy |
| `docx` | `^8.5.0` | MIT | Solo creazione, non editing di .docx esistenti |
| `pptxgenjs` | `^3.12.0` | MIT | Genera .pptx compatibili con PowerPoint 2016+ |

### Compatibilità file generati

| Nodo | Software compatibile |
|------|---------------------|
| ExcelJS | Excel 2010+, LibreOffice Calc 6+, Google Sheets |
| Docx | Word 2013+, LibreOffice Writer 6+, Google Docs |
| PptxGenJS | PowerPoint 2016+, LibreOffice Impress 6+, Google Slides |

---

## 13. Comportamenti edge case

### ExcelJS

| Situazione | Comportamento |
|------------|---------------|
| `addRow` su sheet inesistente | Crea sheet automaticamente se "Auto-crea" attivo; errore porta 2 altrimenti |
| `create` senza `msg.payload` | Comportamento corretto — crea workbook nuovo |
| `read` con Buffer non valido | Errore porta 2: `"Buffer non è un file xlsx valido"` |
| `write` con format `file` e directory inesistente | Errore porta 2: `"Directory non trovata: /path/..."` — il nodo non crea directory |
| `setCell` con formula | Accettato: `value: "=SUM(A1:A10)"`. La formula viene scritta ma non calcolata — aprire in Excel per il calcolo |
| `protect` senza password | Applica protezione senza password (sheet è bloccato ma senza PIN) |

### Docx

| Situazione | Comportamento |
|------------|---------------|
| `addImage` con path inesistente | Errore porta 2 |
| `addImage` con Buffer | Supportato; `params.type` diventa obbligatorio |
| `addTable` con `columnWidths` la cui somma supera la larghezza pagina | Il nodo non valida; docx-js potrebbe generare un file con tabella troncata. Warning su `node.warn()` ma operazione continua |
| `write` senza `create` precedente | Errore porta 2: `"msg.payload non è un documento Docx valido"` |

### PptxGenJS

| Situazione | Comportamento |
|------------|---------------|
| `addText` prima di `addSlide` | Errore porta 2: `"Nessuno slide presente. Eseguire addSlide prima."` |
| `addText` con `slideIndex` out of range | Errore porta 2: `"slideIndex 5 fuori range (presentazione ha 3 slide)"` |
| `addImage` con URL remoto | Supportato da PptxGenJS nativamente (fetch HTTP interno) |
| `write` con format `stream` | PptxGenJS genera prima un Buffer internamente, poi lo wrappa in un PassThrough — non è un vero stream lazy |

---

## 14. Pattern di flusso consigliati

### Pattern 1 — Report Excel da array di dati (lineare)

```
[inject data]
   → [exceljs: create]
   → [exceljs: addSheet]  msg.params.name = "Dati"
   → [exceljs: addRows]   msg.params.rows = msg.payload (array dal DB)
   → [exceljs: styleRange] header bold
   → [exceljs: write]     format: "buffer"
   → [http response]      headers: Content-Type: application/vnd.openxmlformats...
```

### Pattern 2 — Documento Word da template JSON (lineare)

```
[http-in POST con JSON body]
   → [docx: create]
   → [docx: addHeading]    msg.params.text = msg.payload.title
   → [docx: addParagraph]  msg.params.text = msg.payload.body
   → [docx: addTable]      msg.params.rows = msg.payload.tableData
   → [docx: write]         format: "base64"
   → [http response]
```

### Pattern 3 — Workbook multi-sheet con branch paralleli (flow context)

```
[inject]
   → [exceljs: create]    msg.contextKey = "wb_mensile"
   → [split in 3 branch]

Branch A:
   → [query DB entrate]
   → [exceljs: addSheet]  sheet: "Entrate"  msg.contextKey = "wb_mensile"
   → [exceljs: addRows]   msg.contextKey = "wb_mensile"

Branch B:
   → [query DB uscite]
   → [exceljs: addSheet]  sheet: "Uscite"   msg.contextKey = "wb_mensile"
   → [exceljs: addRows]   msg.contextKey = "wb_mensile"

Branch C:
   → [query DB riepilogo]
   → [exceljs: addSheet]  sheet: "Riepilogo" msg.contextKey = "wb_mensile"
   → [exceljs: addRows]   msg.contextKey = "wb_mensile"

[join — attende tutti e 3]
   → [exceljs: write]     msg.contextKey = "wb_mensile", format: "file"
   → [function: flow.set("wb_mensile", null)]  ← pulizia context
   → [notify]
```

### Pattern 4 — Presentazione con dati live (lineare)

```
[inject]
   → [http request: fetch KPI API]
   → [pptx: create]      layout: "LAYOUT_16x9"
   → [pptx: addSlide]    background: "003366"
   → [pptx: addText]     titolo con dati da msg.payload
   → [pptx: addSlide]
   → [pptx: addChart]    dati serie temporale
   → [pptx: addSlide]
   → [pptx: addTable]    tabella KPI dettaglio
   → [pptx: write]       format: "file", path: "/var/slides/kpi_live.pptx"
   → [email: invia allegato]
```

---
