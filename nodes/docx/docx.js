'use strict';

const path = require('path');
const fs = require('fs');

module.exports = function (RED) {

    // ── helpers ────────────────────────────────────────────────────────────

    function toBool(val, defaultVal) {
        if (val === null || val === undefined || val === '') return defaultVal;
        return val === true || val === 'true';
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

    function isDocxDoc(doc) {
        return doc && doc._type === 'docx' && Array.isArray(doc.children);
    }

    function requireDoc(doc) {
        if (!isDocxDoc(doc)) throw new Error('msg.payload is not a valid docx document object');
    }

    // ── style builders ─────────────────────────────────────────────────────

    function buildHeadingStyle(p, gp) {
        if (p.style && typeof p.style === 'object') return p.style;
        const s = {};
        const bold = gp('headingBold', null);
        const color = gp('headingColor', null);
        const size = gp('headingFontSize', null);
        if (bold !== null && bold !== '') s.bold = toBool(bold, false);
        if (color) s.color = color;
        if (size !== null && size !== '') { const n = parseFloat(size); if (!isNaN(n)) s.size = n; }
        return Object.keys(s).length ? s : null;
    }

    function buildParaStyle(p, gp) {
        if (p.style && typeof p.style === 'object') return p.style;
        const s = {};
        const bold      = gp('paraBold', null);
        const italic    = gp('paraItalic', null);
        const underline = gp('paraUnderline', null);
        const color     = gp('paraColor', null);
        const size      = gp('paraFontSize', null);
        const font      = gp('paraFont', null);
        if (bold      !== null && bold !== '')      s.bold      = toBool(bold, false);
        if (italic    !== null && italic !== '')    s.italic    = toBool(italic, false);
        if (underline !== null && underline !== '') s.underline = toBool(underline, false);
        if (color) s.color = color;
        if (size !== null && size !== '') { const n = parseFloat(size); if (!isNaN(n)) s.size = n; }
        if (font) s.font = font;
        return Object.keys(s).length ? s : null;
    }

    function buildTableHeaderStyle(p, gp) {
        if (p.headerStyle && typeof p.headerStyle === 'object') return p.headerStyle;
        const s = {};
        const bold  = gp('tableHeaderBold', null);
        const fill  = gp('tableHeaderFill', null);
        const color = gp('tableHeaderColor', null);
        if (bold  !== null && bold !== '') s.bold  = toBool(bold, true);
        if (fill)  s.fill  = fill;
        if (color) s.color = color;
        return Object.keys(s).length ? s : null;
    }

    // ── document builder (called at write time) ────────────────────────────

    async function buildDocument(internalDoc) {
        const {
            Document, Paragraph, TextRun, HeadingLevel, AlignmentType,
            Table, TableRow, TableCell, WidthType, BorderStyle,
            ImageRun, PageBreak, ShadingType, LevelFormat, UnderlineType,
            ExternalHyperlink, Header, Footer, PageNumber
        } = require('docx');

        const HEADING_MAP = {
            1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2,
            3: HeadingLevel.HEADING_3, 4: HeadingLevel.HEADING_4,
            5: HeadingLevel.HEADING_5, 6: HeadingLevel.HEADING_6
        };
        const ALIGN_MAP = {
            left: AlignmentType.LEFT, center: AlignmentType.CENTER,
            right: AlignmentType.RIGHT, justified: AlignmentType.JUSTIFIED,
            both: AlignmentType.BOTH
        };

        const props = internalDoc.properties || {};

        function buildTextRunStyle(style) {
            if (!style) return {};
            const opts = {};
            if (style.bold)      opts.bold      = true;
            if (style.italic)    opts.italics   = true;
            if (style.underline) opts.underline = { type: UnderlineType.SINGLE };
            if (style.color)     opts.color     = style.color;
            if (style.size)      opts.size      = style.size;
            if (style.font)      opts.font      = style.font;
            return opts;
        }

        function buildParagraphEl(item) {
            const paraOpts = {};
            if (item.alignment) paraOpts.alignment = ALIGN_MAP[item.alignment] || AlignmentType.LEFT;
            if (item.spacing)   paraOpts.spacing   = item.spacing;
            let children;
            if (Array.isArray(item.runs) && item.runs.length > 0) {
                children = item.runs.map(r => {
                    const run = new TextRun({ text: r.text || '', ...buildTextRunStyle(r) });
                    return r.hyperlink
                        ? new ExternalHyperlink({ link: r.hyperlink, children: [run] })
                        : run;
                });
            } else {
                children = [new TextRun({ text: item.text || '', ...buildTextRunStyle(item.style) })];
            }
            paraOpts.children = children;
            return new Paragraph(paraOpts);
        }

        function buildHeadingEl(item) {
            const style = item.style || {};
            return new Paragraph({
                heading: HEADING_MAP[item.level] || HeadingLevel.HEADING_1,
                children: [new TextRun({
                    text:  item.text || '',
                    bold:  style.bold  !== undefined ? style.bold  : undefined,
                    color: style.color || undefined,
                    size:  style.size  || undefined
                })]
            });
        }

        function buildListEl(item) {
            return (item.items || []).map(li => {
                const isObj = typeof li === 'object' && li !== null;
                const text  = isObj ? li.text  : li;
                const level = isObj ? (li.level || 0) : 0;
                if (item.type === 'number') {
                    return new Paragraph({
                        numbering: { reference: 'officedocs-numbering', level },
                        children:  [new TextRun(text)]
                    });
                }
                return new Paragraph({ bullet: { level }, children: [new TextRun(text)] });
            });
        }

        function buildTableEl(item) {
            const headerStyle = item.headerStyle || {};
            const rows = (item.rows || []).map((rowData, rowIdx) => {
                const isHeader = item.headerRow && rowIdx === 0;
                const cells = rowData.map((cellText, colIdx) => {
                    const width = item.columnWidths && item.columnWidths[colIdx]
                        ? { size: item.columnWidths[colIdx], type: WidthType.DXA }
                        : { size: 20, type: WidthType.PERCENTAGE };
                    const runOpts = { text: String(cellText) };
                    if (isHeader && headerStyle.bold !== false) runOpts.bold = true;
                    if (isHeader && headerStyle.color) runOpts.color = headerStyle.color;
                    const cellOpts = {
                        children: [new Paragraph({ children: [new TextRun(runOpts)] })],
                        width
                    };
                    if (isHeader && headerStyle.fill) {
                        cellOpts.shading = { fill: headerStyle.fill, type: ShadingType.CLEAR, color: 'auto' };
                    }
                    if (item.borders !== false) {
                        cellOpts.borders = {
                            top:    { style: BorderStyle.SINGLE, size: 1 },
                            bottom: { style: BorderStyle.SINGLE, size: 1 },
                            left:   { style: BorderStyle.SINGLE, size: 1 },
                            right:  { style: BorderStyle.SINGLE, size: 1 }
                        };
                    }
                    return new TableCell(cellOpts);
                });
                return new TableRow({ children: cells, tableHeader: isHeader });
            });
            return new Table({ rows });
        }

        function buildImageEl(item) {
            let data;
            if (Buffer.isBuffer(item.source)) {
                data = item.source;
            } else if (typeof item.source === 'string') {
                data = item.source.startsWith('data:')
                    ? Buffer.from(item.source.split(',')[1], 'base64')
                    : fs.readFileSync(item.source);
            } else {
                throw new Error('addImage: source must be a file path, base64 data URI, or Buffer');
            }
            return new Paragraph({
                children: [new ImageRun({
                    data,
                    transformation: { width: item.width || 200, height: item.height || 150 },
                    altText: item.altText ? { description: item.altText } : undefined
                })]
            });
        }

        const headerItems = internalDoc.children.filter(c => c._kind === 'header');
        const footerItems = internalDoc.children.filter(c => c._kind === 'footer');
        const bodyItems   = internalDoc.children.filter(c => c._kind !== 'header' && c._kind !== 'footer');

        const hasNumbered = bodyItems.some(c => c._kind === 'list' && c.type === 'number');
        const numbering = hasNumbered ? {
            config: [{
                reference: 'officedocs-numbering',
                levels: [0, 1, 2, 3, 4].map(l => ({
                    level: l, format: LevelFormat.DECIMAL, text: `%${l + 1}.`,
                    alignment: AlignmentType.LEFT,
                    style: { paragraph: { indent: { left: 720 * (l + 1), hanging: 360 } } }
                }))
            }]
        } : undefined;

        const docChildren = [];
        for (const item of bodyItems) {
            switch (item._kind) {
                case 'heading':   docChildren.push(buildHeadingEl(item)); break;
                case 'paragraph': docChildren.push(buildParagraphEl(item)); break;
                case 'list':      docChildren.push(...buildListEl(item)); break;
                case 'table':     docChildren.push(buildTableEl(item)); break;
                case 'image':     docChildren.push(buildImageEl(item)); break;
                case 'pageBreak': docChildren.push(new Paragraph({ children: [new PageBreak()] })); break;
            }
        }

        const docOpts = { sections: [{ properties: {}, children: docChildren }] };
        if (props.title || props.author) {
            docOpts.creator = props.author || '';
            docOpts.title   = props.title  || '';
        }
        if (props.pageSize) {
            docOpts.sections[0].properties.page = { size: props.pageSize };
        }
        if (props.margins) {
            docOpts.sections[0].properties.page = {
                ...(docOpts.sections[0].properties.page || {}),
                margin: props.margins
            };
        }
        if (props.defaultFont) {
            docOpts.styles = {
                default: { document: { run: { font: props.defaultFont.name, size: props.defaultFont.size } } }
            };
        }
        if (numbering) docOpts.numbering = numbering;

        if (headerItems.length > 0) {
            const docHeaders = {};
            for (const h of headerItems) {
                docHeaders[h.headerType] = new Header({
                    children: [new Paragraph({
                        alignment: ALIGN_MAP[h.alignment] || AlignmentType.LEFT,
                        children:  [new TextRun(h.text || '')]
                    })]
                });
            }
            docOpts.sections[0].headers = docHeaders;
        }

        if (footerItems.length > 0) {
            const docFooters = {};
            for (const f of footerItems) {
                const runs = [];
                if (f.text) runs.push(new TextRun(f.text));
                if (f.showPageNum) {
                    if (f.text) runs.push(new TextRun(' — '));
                    runs.push(new TextRun({ children: [PageNumber.CURRENT] }));
                    runs.push(new TextRun(' / '));
                    runs.push(new TextRun({ children: [PageNumber.TOTAL_PAGES] }));
                }
                if (runs.length === 0) runs.push(new TextRun(''));
                docFooters[f.footerType] = new Footer({
                    children: [new Paragraph({
                        alignment: ALIGN_MAP[f.alignment] || AlignmentType.CENTER,
                        children:  runs
                    })]
                });
            }
            docOpts.sections[0].footers = docFooters;
        }

        const { Document: Doc } = require('docx');
        return new Doc(docOpts);
    }

    // ── node ───────────────────────────────────────────────────────────────

    function DocxNode(config) {
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

            function gp(name, defaultVal) {
                return getParam(p, config, name, node, msg, defaultVal);
            }

            try {
                let result;
                const info = { operation };

                switch (operation) {

                    case 'create': {
                        result = {
                            _type: 'docx',
                            properties: {
                                title:       gp('title', ''),
                                author:      gp('author', ''),
                                pageSize:    gp('pageSize', null),
                                margins:     gp('margins', null),
                                defaultFont: gp('defaultFont', null)
                            },
                            children: []
                        };
                        break;
                    }

                    case 'addHeading': {
                        requireDoc(doc);
                        const text  = gp('headingText', '');
                        const level = parseInt(gp('headingLevel', 1), 10);
                        const style = buildHeadingStyle(p, gp);
                        doc.children.push({ _kind: 'heading', text, level, style });
                        result = doc;
                        info.level = level;
                        info.text = text.substring(0, 40);
                        break;
                    }

                    case 'addParagraph': {
                        requireDoc(doc);
                        const text      = gp('paraText', '');
                        const runs      = gp('paraRuns', null);
                        const alignment = gp('paraAlignment', 'left');
                        const spacing   = gp('paraSpacing', null);
                        const style     = buildParaStyle(p, gp);
                        doc.children.push({ _kind: 'paragraph', text, runs, alignment, spacing, style });
                        result = doc;
                        break;
                    }

                    case 'addList': {
                        requireDoc(doc);
                        const type  = gp('listType', 'bullet');
                        const items = gp('listItems', []);
                        if (!Array.isArray(items)) throw new Error('params.items must be an array for addList');
                        doc.children.push({ _kind: 'list', type, items });
                        result = doc;
                        break;
                    }

                    case 'addTable': {
                        requireDoc(doc);
                        const rows         = gp('tableRows', []);
                        const headerRow    = toBool(gp('tableHeaderRow', null), false);
                        const columnWidths = gp('columnWidths', null);
                        const borders      = toBool(gp('tableBorders', null), true);
                        const headerStyle  = buildTableHeaderStyle(p, gp);
                        if (!Array.isArray(rows)) throw new Error('params.rows must be an array for addTable');
                        if (columnWidths) {
                            const pageW = (doc.properties.pageSize && doc.properties.pageSize.width)
                                ? doc.properties.pageSize.width
                                    - ((doc.properties.margins && doc.properties.margins.left) || 1440)
                                    - ((doc.properties.margins && doc.properties.margins.right) || 1440)
                                : 9360;
                            if (columnWidths.reduce((a, b) => a + b, 0) > pageW)
                                node.warn('addTable: columnWidths sum may exceed page content width');
                        }
                        doc.children.push({ _kind: 'table', rows, headerRow, columnWidths, borders, headerStyle });
                        result = doc;
                        break;
                    }

                    case 'addImage': {
                        requireDoc(doc);
                        const source  = gp('imageSource', null);
                        if (!source) throw new Error('params.source is required for addImage');
                        const width   = parseInt(gp('imageWidth', 200), 10);
                        const height  = parseInt(gp('imageHeight', 150), 10);
                        const altText = gp('imageAltText', '');
                        doc.children.push({ _kind: 'image', source, width, height, altText });
                        result = doc;
                        break;
                    }

                    case 'pageBreak': {
                        requireDoc(doc);
                        doc.children.push({ _kind: 'pageBreak' });
                        result = doc;
                        break;
                    }

                    case 'addHeader': {
                        requireDoc(doc);
                        const text       = gp('headerText', '');
                        const headerType = gp('headerType', 'default');
                        const alignment  = gp('headerAlignment', 'left');
                        doc.children.push({ _kind: 'header', text, headerType, alignment });
                        result = doc;
                        break;
                    }

                    case 'addFooter': {
                        requireDoc(doc);
                        const text       = gp('footerText', '');
                        const footerType = gp('footerType', 'default');
                        const alignment  = gp('footerAlignment', 'center');
                        const showPageNum = toBool(gp('footerShowPageNum', null), false);
                        doc.children.push({ _kind: 'footer', text, footerType, alignment, showPageNum });
                        result = doc;
                        break;
                    }

                    case 'write': {
                        requireDoc(doc);
                        const { Packer } = require('docx');
                        const format = gp('format', 'buffer');
                        msg._doc = doc;
                        const document = await buildDocument(doc);
                        switch (format) {
                            case 'buffer':
                                result = await Packer.toBuffer(document);
                                break;
                            case 'base64':
                                result = await Packer.toBase64String(document);
                                break;
                            case 'file': {
                                const filePath = gp('filePath', null);
                                if (!filePath) throw new Error('params.filePath is required for write with format "file"');
                                const absPath = path.resolve(filePath);
                                if (!fs.existsSync(path.dirname(absPath)))
                                    throw new Error(`Directory not found: ${path.dirname(absPath)}`);
                                const buf = await Packer.toBuffer(document);
                                fs.writeFileSync(absPath, buf);
                                result = absPath;
                                info.filePath = absPath;
                                break;
                            }
                            default:
                                throw new Error(`Unsupported write format '${format}'. Use "file", "buffer" or "base64"`);
                        }
                        info.format = format;
                        info.childrenCount = doc.children.length;
                        break;
                    }

                    default:
                        throw new Error(`Operation '${operation}' is not supported`);
                }

                if (msg.contextKey && result && result._type === 'docx')
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

    RED.nodes.registerType('docx', DocxNode);
};
