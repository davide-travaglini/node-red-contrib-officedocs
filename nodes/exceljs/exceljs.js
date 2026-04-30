'use strict';

const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

module.exports = function (RED) {

    // ── helpers ────────────────────────────────────────────────────────────

    function toBool(val, defaultVal) {
        if (val === null || val === undefined || val === '') return defaultVal;
        return val === true || val === 'true';
    }

    function cellRefToRowCol(ref) {
        const match = /^([A-Z]+)(\d+)$/.exec(ref.toUpperCase());
        if (!match) throw new Error(`Invalid cell reference: ${ref}`);
        let col = 0;
        for (const ch of match[1]) col = col * 26 + ch.charCodeAt(0) - 64;
        return { row: parseInt(match[2], 10), col };
    }

    function resolveConfigParam(config, name, node, msg) {
        const val = config[name];
        const type = config[name + 'Type'] || 'str';
        if (val === undefined || val === null || val === '') return undefined;
        try { return RED.util.evaluateNodeProperty(val, type, node, msg); }
        catch (_) { return undefined; }
    }

    function getParam(msgParams, config, name, node, msg, defaultVal) {
        if (msgParams && msgParams[name] !== undefined) return msgParams[name];
        const v = resolveConfigParam(config, name, node, msg);
        return v !== undefined ? v : defaultVal;
    }

    function requireWorkbook(doc) {
        if (!doc || typeof doc.addWorksheet !== 'function')
            throw new Error('msg.payload is not a valid ExcelJS Workbook');
    }

    function getOrCreateSheet(doc, sheetName, autoCreate) {
        let ws = doc.getWorksheet(sheetName);
        if (!ws) {
            if (autoCreate) ws = doc.addWorksheet(sheetName);
            else throw new Error(`Sheet '${sheetName}' not found in workbook`);
        }
        return ws;
    }

    // ── style builder ──────────────────────────────────────────────────────

    function buildCellStyle(p, gp) {
        if (p.style && typeof p.style === 'object') return p.style;
        const s = {};
        const font = {};
        const bold = gp('styleFontBold', null);
        const fontSize = gp('styleFontSize', null);
        const fontColor = gp('styleFontColor', null);
        if (bold !== null && bold !== '') font.bold = toBool(bold, false);
        if (fontSize !== null && fontSize !== '') {
            const n = parseFloat(fontSize);
            if (!isNaN(n)) font.size = n;
        }
        if (fontColor) font.color = { argb: fontColor };
        if (Object.keys(font).length) s.font = font;

        const fillColor = gp('styleFillColor', null);
        if (fillColor) s.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };

        const numFmt = gp('styleNumFmt', null);
        if (numFmt) s.numFmt = numFmt;

        const alignment = gp('styleAlignment', null);
        if (alignment) s.alignment = { horizontal: alignment };

        return Object.keys(s).length > 0 ? s : null;
    }

    // ── protect options builder ────────────────────────────────────────────

    function buildProtectOptions(p, gp) {
        if (p.options && typeof p.options === 'object') return p.options;
        const opts = {};
        const pairs = [
            ['protectSelectLocked',   'selectLockedCells',   true],
            ['protectSelectUnlocked', 'selectUnlockedCells', true],
            ['protectFormatCells',    'formatCells',         false],
            ['protectFormatColumns',  'formatColumns',       false],
            ['protectFormatRows',     'formatRows',          false],
            ['protectInsertRows',     'insertRows',          false],
            ['protectDeleteRows',     'deleteRows',          false],
            ['protectSort',           'sort',                false],
            ['protectAutoFilter',     'autoFilter',          false],
        ];
        for (const [configKey, optKey, def] of pairs) {
            const v = gp(configKey, null);
            if (v !== null && v !== undefined && v !== '') opts[optKey] = toBool(v, def);
        }
        return opts;
    }

    // ── node ───────────────────────────────────────────────────────────────

    function ExcelJSNode(config) {
        RED.nodes.createNode(this, config);
        const node = this;

        node.on('input', async function (msg, send, done) {
            send = send || function () { node.send.apply(node, arguments); };

            const operation = msg.operation
                ?? resolveConfigParam(config, 'operation', node, msg)
                ?? 'create';

            const p = (typeof msg.params === 'object' && msg.params !== null) ? msg.params : {};

            let doc = msg.contextKey
                ? (node.context().flow.get(msg.contextKey) ?? msg.payload)
                : msg.payload;

            const autoCreate = config.autoCreateSheet !== false;

            function gp(name, defaultVal) {
                return getParam(p, config, name, node, msg, defaultVal);
            }

            // Resolves sheet param, falling back to doc._activeSheet if set
            function gs(opName) {
                const s = gp('sheet', (doc && doc._activeSheet) || null);
                if (!s) throw new Error(`params.sheet is required for ${opName}. Set it in the panel, via msg.params.sheet, or use setActiveSheet first.`);
                return s;
            }

            try {
                let result;
                const info = { operation };

                switch (operation) {

                    case 'create': {
                        const wb = new ExcelJS.Workbook();
                        const creator = gp('creator', '');
                        const lastModifiedBy = gp('lastModifiedBy', '');
                        if (creator) wb.creator = creator;
                        if (lastModifiedBy) wb.lastModifiedBy = lastModifiedBy;
                        wb.created = new Date();
                        wb.modified = new Date();
                        result = wb;
                        break;
                    }

                    case 'read': {
                        const source = gp('source', null);
                        if (!source) throw new Error('params.source is required for read');
                        const wb = new ExcelJS.Workbook();
                        if (typeof source === 'string') {
                            await wb.xlsx.readFile(source);
                        } else if (Buffer.isBuffer(source)) {
                            await wb.xlsx.load(source);
                        } else {
                            throw new Error('params.source must be a file path string or a Buffer');
                        }
                        result = wb;
                        info.sheetCount = wb.worksheets.length;
                        break;
                    }

                    case 'addSheet': {
                        requireWorkbook(doc);
                        const sheetName = gp('sheetName', null);
                        if (!sheetName) throw new Error('params.sheetName is required for addSheet');
                        if (doc.getWorksheet(sheetName))
                            throw new Error(`Sheet '${sheetName}' already exists in workbook`);
                        const tabColor = gp('tabColor', null);
                        const state = gp('sheetState', 'visible');
                        const wsOpts = { properties: {}, state: state || 'visible' };
                        if (tabColor) wsOpts.properties.tabColor = { argb: tabColor };
                        doc.addWorksheet(sheetName, wsOpts);
                        result = doc;
                        info.sheet = sheetName;
                        break;
                    }

                    case 'setActiveSheet': {
                        requireWorkbook(doc);
                        const sheetName = gp('sheetName', null);
                        if (!sheetName) throw new Error('params.sheetName is required for setActiveSheet');
                        const ws = doc.getWorksheet(sheetName);
                        if (!ws) throw new Error(`Sheet '${sheetName}' not found in workbook`);
                        doc._activeSheet = sheetName;
                        result = doc;
                        info.sheet = sheetName;
                        break;
                    }

                    case 'getInfo': {
                        requireWorkbook(doc);
                        const sheets = [];
                        doc.eachSheet(ws => {
                            sheets.push({
                                name:           ws.name,
                                state:          ws.state,
                                rowCount:       ws.rowCount,
                                actualRowCount: ws.actualRowCount,
                                columnCount:    ws.columnCount,
                                usedRange:      ws.dimensions ? ws.dimensions.model : null
                            });
                        });
                        msg._doc = doc;
                        result = { sheetCount: sheets.length, sheets };
                        break;
                    }

                    case 'addRow': {
                        requireWorkbook(doc);
                        const sheet = gs('addRow');
                        const values = gp('rowValues', null);
                        if (values === null || values === undefined)
                            throw new Error('params.rowValues is required for addRow');
                        const ws = getOrCreateSheet(doc, sheet, autoCreate);
                        ws.addRow(values);
                        result = doc;
                        info.sheet = sheet;
                        break;
                    }

                    case 'addRows': {
                        requireWorkbook(doc);
                        const sheet = gs('addRows');
                        const rows = gp('rowsData', null);
                        if (!Array.isArray(rows))
                            throw new Error('params.rowsData must be an array for addRows');
                        const ws = getOrCreateSheet(doc, sheet, autoCreate);
                        ws.addRows(rows);
                        result = doc;
                        info.sheet = sheet;
                        info.rowsAdded = rows.length;
                        break;
                    }

                    case 'setCell': {
                        requireWorkbook(doc);
                        const sheet = gs('setCell');
                        const cellRef = gp('cell', null);
                        if (!cellRef) throw new Error('params.cell is required for setCell');
                        const cellValue = gp('cellValue', undefined);
                        const ws = getOrCreateSheet(doc, sheet, autoCreate);

                        let cellObj;
                        if (typeof cellRef === 'string') {
                            cellObj = ws.getCell(cellRef);
                        } else if (typeof cellRef === 'object' && cellRef.row && cellRef.col) {
                            cellObj = ws.getCell(cellRef.row, cellRef.col);
                        } else {
                            throw new Error('params.cell must be an A1 string or {row, col} object (1-based col number)');
                        }

                        if (cellValue !== undefined) {
                            if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
                                cellObj.value = { formula: cellValue.slice(1) };
                            } else {
                                cellObj.value = cellValue;
                            }
                        }

                        const style = buildCellStyle(p, gp);
                        if (style) {
                            if (style.font)      cellObj.font      = style.font;
                            if (style.fill)      cellObj.fill      = style.fill;
                            if (style.alignment) cellObj.alignment = style.alignment;
                            if (style.border)    cellObj.border    = style.border;
                            if (style.numFmt)    cellObj.numFmt    = style.numFmt;
                        }
                        result = doc;
                        info.sheet = sheet;
                        info.cell = typeof cellRef === 'string' ? cellRef : `R${cellRef.row}C${cellRef.col}`;
                        break;
                    }

                    case 'styleRange': {
                        requireWorkbook(doc);
                        const sheet = gs('styleRange');
                        const range = gp('range', null);
                        if (!range) throw new Error('params.range is required for styleRange');
                        const style = buildCellStyle(p, gp);
                        if (!style) throw new Error('At least one style property is required for styleRange');
                        const ws = getOrCreateSheet(doc, sheet, autoCreate);
                        const parts = range.split(':');
                        if (parts.length !== 2)
                            throw new Error(`Invalid range '${range}'. Use A1:B2 notation.`);
                        const start = cellRefToRowCol(parts[0]);
                        const end   = cellRefToRowCol(parts[1]);
                        for (let r = start.row; r <= end.row; r++) {
                            for (let c = start.col; c <= end.col; c++) {
                                const cell = ws.getCell(r, c);
                                if (style.font)      cell.font      = style.font;
                                if (style.fill)      cell.fill      = style.fill;
                                if (style.alignment) cell.alignment = style.alignment;
                                if (style.border)    cell.border    = style.border;
                                if (style.numFmt)    cell.numFmt    = style.numFmt;
                            }
                        }
                        result = doc;
                        info.sheet = sheet;
                        info.range = range;
                        break;
                    }

                    case 'addTable': {
                        requireWorkbook(doc);
                        const sheet = gs('addTable');
                        const tableName = gp('tableName', null);
                        if (!tableName) throw new Error('params.tableName is required for addTable');
                        const ref            = gp('tableRef', 'A1');
                        const tableStyle     = gp('tableStyle', 'TableStyleMedium9');
                        const showRowStripes = toBool(gp('showRowStripes', null), true);
                        const columns        = gp('columns', []);
                        const tableRows      = gp('tableRows', []);
                        if (!Array.isArray(columns))
                            throw new Error('params.columns must be an array for addTable');
                        if (!Array.isArray(tableRows))
                            throw new Error('params.tableRows must be an array for addTable');
                        const ws = getOrCreateSheet(doc, sheet, autoCreate);
                        ws.addTable({
                            name: tableName, ref,
                            headerRow: true,
                            style: { theme: tableStyle, showRowStripes },
                            columns, rows: tableRows
                        });
                        result = doc;
                        info.sheet = sheet;
                        info.tableName = tableName;
                        break;
                    }

                    case 'conditionalFormat': {
                        requireWorkbook(doc);
                        const sheet = gs('conditionalFormat');
                        const range = gp('range', null);
                        if (!range) throw new Error('params.range is required for conditionalFormat');
                        const rules = gp('cfRules', null);
                        if (!Array.isArray(rules))
                            throw new Error('params.cfRules must be an array for conditionalFormat');
                        const ws = doc.getWorksheet(sheet);
                        if (!ws) throw new Error(`Sheet '${sheet}' not found in workbook`);
                        ws.addConditionalFormatting({ ref: range, rules });
                        result = doc;
                        info.sheet = sheet;
                        info.range = range;
                        break;
                    }

                    case 'protect': {
                        requireWorkbook(doc);
                        const target   = gp('protectTarget', 'sheet');
                        const password = gp('password', '');
                        const options  = buildProtectOptions(p, gp);
                        if (target === 'workbook') {
                            node.warn('ExcelJS has limited workbook-level protection support. Applying sheet protection to all sheets instead.');
                            const promises = [];
                            doc.eachSheet(ws => promises.push(ws.protect(password, options)));
                            await Promise.all(promises);
                        } else {
                            const sheet = gs('protect');
                            const ws = doc.getWorksheet(sheet);
                            if (!ws) throw new Error(`Sheet '${sheet}' not found in workbook`);
                            await ws.protect(password, options);
                            info.sheet = sheet;
                        }
                        result = doc;
                        break;
                    }

                    case 'mergeCell': {
                        requireWorkbook(doc);
                        const sheet = gs('mergeCell');
                        const range = gp('range', null);
                        if (!range) throw new Error('params.range is required for mergeCell');
                        const ws = getOrCreateSheet(doc, sheet, autoCreate);
                        ws.mergeCells(range);
                        result = doc;
                        info.sheet = sheet;
                        info.range = range;
                        break;
                    }

                    case 'setColumnWidth': {
                        requireWorkbook(doc);
                        const sheet = gs('setColumnWidth');
                        const colRef    = gp('colRef', null);
                        if (!colRef) throw new Error('params.colRef is required for setColumnWidth');
                        const colWidth  = gp('colWidth', null);
                        const colHidden = gp('colHidden', null);
                        const ws = getOrCreateSheet(doc, sheet, autoCreate);
                        const col = ws.getColumn(colRef);
                        if (colWidth  !== null && colWidth  !== '') col.width  = parseFloat(colWidth);
                        if (colHidden !== null && colHidden !== '') col.hidden = toBool(colHidden, false);
                        result = doc;
                        info.sheet = sheet;
                        break;
                    }

                    case 'setRowHeight': {
                        requireWorkbook(doc);
                        const sheet = gs('setRowHeight');
                        const rowNum = parseInt(gp('rowNumber', null), 10);
                        if (isNaN(rowNum)) throw new Error('params.rowNumber is required for setRowHeight');
                        const rowHeight = gp('rowHeight', null);
                        const rowHidden = gp('rowHidden', null);
                        const ws = getOrCreateSheet(doc, sheet, autoCreate);
                        const row = ws.getRow(rowNum);
                        if (rowHeight !== null && rowHeight !== '') row.height = parseFloat(rowHeight);
                        if (rowHidden !== null && rowHidden !== '') row.hidden = toBool(rowHidden, false);
                        result = doc;
                        info.sheet = sheet;
                        break;
                    }

                    case 'freezePanes': {
                        requireWorkbook(doc);
                        const sheet = gs('freezePanes');
                        const freezeRow = parseInt(gp('freezeRow', 0), 10) || 0;
                        const freezeCol = parseInt(gp('freezeCol', 0), 10) || 0;
                        if (freezeRow === 0 && freezeCol === 0)
                            node.warn('freezePanes: both freezeRow and freezeCol are 0 — no panes will be frozen');
                        const ws = getOrCreateSheet(doc, sheet, autoCreate);
                        ws.views = [{ state: 'frozen', ySplit: freezeRow, xSplit: freezeCol }];
                        result = doc;
                        info.sheet = sheet;
                        info.freezeRow = freezeRow;
                        info.freezeCol = freezeCol;
                        break;
                    }

                    case 'readCell': {
                        requireWorkbook(doc);
                        const sheet = gs('readCell');
                        const cellRef = gp('cell', null);
                        if (!cellRef) throw new Error('params.cell is required for readCell');
                        const ws = doc.getWorksheet(sheet);
                        if (!ws) throw new Error(`Sheet '${sheet}' not found in workbook`);
                        let cellObj;
                        if (typeof cellRef === 'string') {
                            cellObj = ws.getCell(cellRef);
                        } else if (typeof cellRef === 'object' && cellRef.row && cellRef.col) {
                            cellObj = ws.getCell(cellRef.row, cellRef.col);
                        } else {
                            throw new Error('params.cell must be an A1 string or {row, col} object (1-based col number)');
                        }
                        msg._doc = doc;
                        result = {
                            address: cellObj.address,
                            value:   cellObj.value,
                            text:    cellObj.text,
                            type:    cellObj.type,
                            formula: cellObj.formula || null
                        };
                        info.sheet = sheet;
                        info.cell = cellObj.address;
                        break;
                    }

                    case 'readRange': {
                        requireWorkbook(doc);
                        const sheet = gs('readRange');
                        const range = gp('range', null);
                        if (!range) throw new Error('params.range is required for readRange');
                        const ws = doc.getWorksheet(sheet);
                        if (!ws) throw new Error(`Sheet '${sheet}' not found in workbook`);
                        const parts = range.split(':');
                        if (parts.length !== 2) throw new Error(`Invalid range '${range}'. Use A1:B2 notation.`);
                        const start = cellRefToRowCol(parts[0]);
                        const end   = cellRefToRowCol(parts[1]);
                        const grid = [];
                        for (let r = start.row; r <= end.row; r++) {
                            const rowArr = [];
                            for (let c = start.col; c <= end.col; c++) {
                                rowArr.push(ws.getCell(r, c).value);
                            }
                            grid.push(rowArr);
                        }
                        msg._doc = doc;
                        result = grid;
                        info.sheet = sheet;
                        info.range = range;
                        info.rows = grid.length;
                        info.cols = grid[0] ? grid[0].length : 0;
                        break;
                    }

                    case 'write': {
                        requireWorkbook(doc);
                        const format = gp('format', 'buffer');
                        msg._doc = doc;
                        switch (format) {
                            case 'buffer':
                                result = await doc.xlsx.writeBuffer();
                                break;
                            case 'base64': {
                                const buf = await doc.xlsx.writeBuffer();
                                result = buf.toString('base64');
                                break;
                            }
                            case 'file': {
                                const filePath = gp('filePath', null);
                                if (!filePath)
                                    throw new Error('params.filePath is required for write with format "file"');
                                const absPath = path.resolve(filePath);
                                if (!fs.existsSync(path.dirname(absPath)))
                                    throw new Error(`Directory not found: ${path.dirname(absPath)}`);
                                await doc.xlsx.writeFile(absPath);
                                result = absPath;
                                info.filePath = absPath;
                                break;
                            }
                            default:
                                throw new Error(`Unsupported write format '${format}'. Use "file", "buffer" or "base64"`);
                        }
                        info.format = format;
                        break;
                    }

                    default:
                        throw new Error(`Operation '${operation}' is not supported`);
                }

                if (msg.contextKey && result && typeof result.addWorksheet === 'function')
                    node.context().flow.set(msg.contextKey, result);

                msg.payload = result;
                msg.info = info;
                node.status({ fill: 'green', shape: 'dot', text: operation });
                send([msg, null]);
                done();

            } catch (err) {
                msg.error = { message: err.message, operation, params: p, stack: err.stack };
                node.status({ fill: 'red', shape: 'ring', text: err.message.substring(0, 50) });
                send([null, msg]);
                done();
            }
        });
    }

    RED.nodes.registerType('exceljs', ExcelJSNode);
};
