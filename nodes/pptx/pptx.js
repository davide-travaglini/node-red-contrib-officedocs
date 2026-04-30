'use strict';

const path = require('path');
const fs = require('fs');

module.exports = function (RED) {

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

    function isPptxDoc(doc) {
        return doc && doc._type === 'pptx' && doc.pptx && Array.isArray(doc.slides);
    }

    function requirePptx(doc) {
        if (!isPptxDoc(doc)) throw new Error('msg.payload is not a valid pptx presentation object');
    }

    function getSlide(doc, slideIndex) {
        const idx = slideIndex !== null && slideIndex !== undefined
            ? slideIndex
            : doc.currentSlideIndex;
        if (idx < 0 || idx >= doc.slides.length) {
            throw new Error(`slideIndex ${idx} is out of range (presentation has ${doc.slides.length} slide(s))`);
        }
        return doc.slides[idx];
    }

    function resolveImageSource(source, type) {
        if (Buffer.isBuffer(source)) {
            const ext = (type || 'png').toLowerCase();
            return { data: `image/${ext};base64,${source.toString('base64')}` };
        }
        if (typeof source === 'string') {
            if (source.startsWith('http://') || source.startsWith('https://')) return { path: source };
            if (source.startsWith('data:')) return { data: source };
            return { path: source };
        }
        throw new Error('addImage: source must be a URL, file path, base64 data URI, or Buffer');
    }

    // ── position builder ───────────────────────────────────────────────────
    // Priority: msg.params.position (whole object) → individual posX/posY/posW/posH fields

    function buildPosition(p, gp, defX, defY, defW, defH) {
        if (p.position && typeof p.position === 'object') return p.position;
        const xv = gp('posX', defX); const yv = gp('posY', defY);
        const wv = gp('posW', defW); const hv = gp('posH', defH);
        const x = parseFloat(xv); const y = parseFloat(yv);
        const w = parseFloat(wv); const h = parseFloat(hv);
        const pos = {
            x: isNaN(x) ? defX : x,
            y: isNaN(y) ? defY : y,
            w: isNaN(w) ? defW : w
        };
        if (!isNaN(h)) pos.h = h;
        else if (defH !== undefined && defH !== null) pos.h = defH;
        return pos;
    }

    // ── text style builder ─────────────────────────────────────────────────

    function buildTextStyle(p, gp) {
        if (p.textStyle && typeof p.textStyle === 'object') return p.textStyle;
        const s = {};
        const fontSize = gp('textFontSize', null);
        const bold     = gp('textBold', null);
        const color    = gp('textColor', null);
        const align    = gp('textAlign', null);
        const fontFace = gp('textFontFace', null);
        const valign   = gp('textValign', null);
        const wrap     = gp('textWrap', null);
        if (fontSize !== null && fontSize !== '') { const n = parseFloat(fontSize); if (!isNaN(n)) s.fontSize = n; }
        if (bold     !== null && bold !== '')     s.bold     = toBool(bold, false);
        if (color)    s.color    = color;
        if (align)    s.align    = align;
        if (fontFace) s.fontFace = fontFace;
        if (valign)   s.valign   = valign;
        if (wrap      !== null && wrap !== '')    s.wrap     = toBool(wrap, true);
        return s;
    }

    // ── shape style builder ────────────────────────────────────────────────

    function buildShapeStyle(p, gp) {
        if (p.shapeStyle && typeof p.shapeStyle === 'object') return p.shapeStyle;
        const s = {};
        const fillColor  = gp('shapeFillColor', null);
        const lineColor  = gp('shapeLineColor', null);
        const lineWidth  = gp('shapeLineWidth', null);
        if (fillColor) s.fill = { color: fillColor };
        if (lineColor || (lineWidth !== null && lineWidth !== '')) {
            s.line = {};
            if (lineColor) s.line.color = lineColor;
            if (lineWidth !== null && lineWidth !== '') {
                const n = parseFloat(lineWidth);
                if (!isNaN(n)) s.line.width = n;
            }
        }
        return s;
    }

    // ── node ───────────────────────────────────────────────────────────────

    function PptxNode(config) {
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
                const PptxGenJS = require('pptxgenjs');

                switch (operation) {

                    case 'create': {
                        const pptx = new PptxGenJS();
                        const layout = gp('layout', 'LAYOUT_16x9');
                        if (typeof layout === 'object' && layout.width && layout.height) {
                            pptx.defineLayout({ name: 'CUSTOM', width: layout.width, height: layout.height });
                            pptx.layout = 'CUSTOM';
                        } else {
                            pptx.layout = layout;
                        }
                        const author  = gp('author', '');
                        const company = gp('company', '');
                        const theme   = gp('theme', null);
                        if (author)  pptx.author  = author;
                        if (company) pptx.company = company;
                        if (theme && theme.headFontFace) pptx.theme = { headFontFace: theme.headFontFace };
                        if (theme && theme.bodyFontFace) pptx.theme = { ...(pptx.theme || {}), bodyFontFace: theme.bodyFontFace };
                        result = { _type: 'pptx', pptx, slides: [], currentSlideIndex: -1 };
                        break;
                    }

                    case 'addSlide': {
                        requirePptx(doc);
                        const slideOpts = {};
                        const masterName = gp('masterName', null);
                        const bgColor    = gp('bgColor', null);
                        if (masterName) slideOpts.masterName = masterName;
                        const slide = doc.pptx.addSlide(slideOpts);
                        if (bgColor) {
                            slide.background = { color: bgColor };
                        }
                        doc.slides.push(slide);
                        doc.currentSlideIndex = doc.slides.length - 1;
                        result = doc;
                        info.slideCount = doc.slides.length;
                        break;
                    }

                    case 'getSlideCount': {
                        requirePptx(doc);
                        msg._doc = doc;
                        result = { slideCount: doc.slides.length, currentSlideIndex: doc.currentSlideIndex };
                        info.slideCount = doc.slides.length;
                        break;
                    }

                    case 'addText': {
                        requirePptx(doc);
                        if (doc.slides.length === 0) throw new Error('No slides present. Run addSlide first.');
                        const text     = gp('text', '');
                        const slideIdx = gp('slideIndex', null);
                        const position = buildPosition(p, gp, 0.5, 0.5, 9, 1.5);
                        const style    = buildTextStyle(p, gp);
                        const slide    = getSlide(doc, slideIdx);
                        slide.addText(text, { ...position, ...style });
                        result = doc;
                        break;
                    }

                    case 'addShape': {
                        requirePptx(doc);
                        if (doc.slides.length === 0) throw new Error('No slides present. Run addSlide first.');
                        const shapeType = gp('shapeType', 'rect');
                        const slideIdx  = gp('slideIndex', null);
                        const position  = buildPosition(p, gp, 1, 1, 4, 2);
                        const style     = buildShapeStyle(p, gp);
                        const slide     = getSlide(doc, slideIdx);
                        slide.addShape(doc.pptx.ShapeType[shapeType] || shapeType, { ...position, ...style });
                        result = doc;
                        break;
                    }

                    case 'addImage': {
                        requirePptx(doc);
                        if (doc.slides.length === 0) throw new Error('No slides present. Run addSlide first.');
                        const source   = gp('imageSource', null);
                        if (!source) throw new Error('params.source is required for addImage');
                        const imgType  = gp('imageType', 'png');
                        const slideIdx = gp('slideIndex', null);
                        const position = buildPosition(p, gp, 0.5, 1.5, 4, 3);
                        const sizing   = gp('sizing', null);
                        const hyperlink = gp('hyperlink', null);
                        const slide    = getSlide(doc, slideIdx);
                        const imgOpts  = { ...resolveImageSource(source, imgType), ...position };
                        if (sizing)    imgOpts.sizing    = sizing;
                        if (hyperlink) imgOpts.hyperlink = hyperlink;
                        slide.addImage(imgOpts);
                        result = doc;
                        break;
                    }

                    case 'addChart': {
                        requirePptx(doc);
                        if (doc.slides.length === 0) throw new Error('No slides present. Run addSlide first.');
                        const chartType = gp('chartType', 'bar');
                        const slideIdx  = gp('slideIndex', null);
                        const position  = buildPosition(p, gp, 0.5, 1.5, 8, 4.5);
                        const data      = gp('chartData', []);
                        const options   = gp('chartOptions', {}) || {};
                        if (!Array.isArray(data)) throw new Error('params.data must be an array for addChart');
                        const slide     = getSlide(doc, slideIdx);
                        const pptxChartType = doc.pptx.ChartType
                            ? (doc.pptx.ChartType[chartType.toUpperCase()] || chartType)
                            : chartType;
                        slide.addChart(pptxChartType, data, { ...position, ...options });
                        result = doc;
                        break;
                    }

                    case 'addTable': {
                        requirePptx(doc);
                        if (doc.slides.length === 0) throw new Error('No slides present. Run addSlide first.');
                        const rows     = gp('tableRows', []);
                        const slideIdx = gp('slideIndex', null);
                        const position = buildPosition(p, gp, 0.5, 2, 8, null);
                        const options  = gp('tableOptions', {}) || {};
                        if (!Array.isArray(rows)) throw new Error('params.rows must be an array for addTable');
                        const slide    = getSlide(doc, slideIdx);
                        slide.addTable(rows, { ...position, ...options });
                        result = doc;
                        break;
                    }

                    case 'setLayout': {
                        requirePptx(doc);
                        const layout = gp('layout', 'LAYOUT_16x9');
                        if (typeof layout === 'object' && layout.width && layout.height) {
                            doc.pptx.defineLayout({ name: 'CUSTOM', width: layout.width, height: layout.height });
                            doc.pptx.layout = 'CUSTOM';
                        } else {
                            doc.pptx.layout = layout;
                        }
                        result = doc;
                        break;
                    }

                    case 'addNotes': {
                        requirePptx(doc);
                        if (doc.slides.length === 0) throw new Error('No slides present. Run addSlide first.');
                        const notes    = gp('notes', '');
                        if (typeof notes !== 'string') throw new Error('params.notes must be a string');
                        const slideIdx = gp('slideIndex', null);
                        const slide    = getSlide(doc, slideIdx);
                        slide.addNotes(notes);
                        result = doc;
                        break;
                    }

                    case 'defineMaster': {
                        requirePptx(doc);
                        if (doc.slides.length > 0)
                            node.warn('defineMaster: master slides should be defined before addSlide for reliable layout resolution');
                        const masterTitle   = gp('masterTitle', 'MASTER_SLIDE');
                        if (!masterTitle) throw new Error('params.masterTitle is required for defineMaster');
                        const masterBkg     = gp('masterBkg', null);
                        const masterObjects = gp('masterObjects', null);
                        const masterDef = { title: masterTitle };
                        if (masterBkg) {
                            masterDef.background = typeof masterBkg === 'object'
                                ? masterBkg
                                : { color: masterBkg };
                        }
                        if (Array.isArray(masterObjects) && masterObjects.length > 0)
                            masterDef.objects = masterObjects;
                        doc.pptx.defineSlideMaster(masterDef);
                        result = doc;
                        info.masterTitle = masterTitle;
                        break;
                    }

                    case 'write': {
                        requirePptx(doc);
                        const format = gp('format', 'buffer');
                        msg._doc = doc;
                        switch (format) {
                            case 'buffer':
                                result = await doc.pptx.write({ outputType: 'nodebuffer' });
                                break;
                            case 'base64':
                                result = await doc.pptx.write({ outputType: 'base64' });
                                break;
                            case 'file': {
                                const filePath = gp('filePath', null);
                                if (!filePath) throw new Error('params.filePath is required for write with format "file"');
                                const absPath = path.resolve(filePath);
                                if (!fs.existsSync(path.dirname(absPath)))
                                    throw new Error(`Directory not found: ${path.dirname(absPath)}`);
                                await doc.pptx.writeFile({ fileName: absPath });
                                result = absPath;
                                info.filePath = absPath;
                                break;
                            }
                            default:
                                throw new Error(`Unsupported write format '${format}'. Use "file", "buffer", or "base64"`);
                        }
                        info.format = format;
                        info.slideCount = doc.slides.length;
                        break;
                    }

                    default:
                        throw new Error(`Operation '${operation}' is not supported`);
                }

                if (msg.contextKey && result && result._type === 'pptx')
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

    RED.nodes.registerType('pptx', PptxNode);
};
