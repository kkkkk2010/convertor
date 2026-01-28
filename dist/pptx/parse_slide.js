"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.emuToPx = emuToPx;
exports.ooxmlRotToDeg = ooxmlRotToDeg;
exports.parseSlide = parseSlide;
const node_path_1 = __importDefault(require("node:path"));
const promises_1 = __importDefault(require("node:fs/promises"));
const theme_1 = require("./theme");
const EMU_PER_INCH = 914400;
function emuToPx(emu) {
    return (emu / EMU_PER_INCH) * 96;
}
function ooxmlRotToDeg(rot) {
    if (!rot) {
        return 0;
    }
    return rot / 60000;
}
async function parseSlide(slideXml, options) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o;
    const spTree = (_b = (_a = slideXml["p:sld"]) === null || _a === void 0 ? void 0 : _a["p:cSld"]) === null || _b === void 0 ? void 0 : _b["p:spTree"];
    const shapes = ensureArray(spTree === null || spTree === void 0 ? void 0 : spTree["p:sp"]);
    const pics = ensureArray(spTree === null || spTree === void 0 ? void 0 : spTree["p:pic"]);
    const elements = [];
    let textCount = 0;
    let schemeClrCount = 0;
    let multistyleCount = 0;
    for (const shape of shapes) {
        const textBody = shape === null || shape === void 0 ? void 0 : shape["p:txBody"];
        if (!textBody) {
            continue;
        }
        textCount += 1;
        const xfrm = (_c = shape === null || shape === void 0 ? void 0 : shape["p:spPr"]) === null || _c === void 0 ? void 0 : _c["a:xfrm"];
        const off = xfrm === null || xfrm === void 0 ? void 0 : xfrm["a:off"];
        const ext = xfrm === null || xfrm === void 0 ? void 0 : xfrm["a:ext"];
        const rotation = ooxmlRotToDeg(Number(xfrm === null || xfrm === void 0 ? void 0 : xfrm["@_rot"]));
        const { text, style, schemeClrUsed, multistyle } = extractTextAndStyle(textBody, options.theme);
        if (schemeClrUsed) {
            schemeClrCount += 1;
        }
        if (multistyle) {
            multistyleCount += 1;
            console.warn(`MULTISTYLE slide ${options.slideIndex} text t${textCount} using base`);
        }
        const element = {
            id: `t${textCount}`,
            type: "text",
            x: emuToPx(Number((_d = off === null || off === void 0 ? void 0 : off["@_x"]) !== null && _d !== void 0 ? _d : 0)),
            y: emuToPx(Number((_e = off === null || off === void 0 ? void 0 : off["@_y"]) !== null && _e !== void 0 ? _e : 0)),
            width: emuToPx(Number((_f = ext === null || ext === void 0 ? void 0 : ext["@_cx"]) !== null && _f !== void 0 ? _f : 0)),
            height: emuToPx(Number((_g = ext === null || ext === void 0 ? void 0 : ext["@_cy"]) !== null && _g !== void 0 ? _g : 0)),
            rotation,
            text,
            style,
        };
        elements.push(element);
    }
    let imageCount = 0;
    for (const pic of pics) {
        const blip = (_h = pic === null || pic === void 0 ? void 0 : pic["p:blipFill"]) === null || _h === void 0 ? void 0 : _h["a:blip"];
        const embed = blip === null || blip === void 0 ? void 0 : blip["@_r:embed"];
        if (!embed) {
            continue;
        }
        imageCount += 1;
        const xfrm = (_j = pic === null || pic === void 0 ? void 0 : pic["p:spPr"]) === null || _j === void 0 ? void 0 : _j["a:xfrm"];
        const off = xfrm === null || xfrm === void 0 ? void 0 : xfrm["a:off"];
        const ext = xfrm === null || xfrm === void 0 ? void 0 : xfrm["a:ext"];
        const rotation = ooxmlRotToDeg(Number(xfrm === null || xfrm === void 0 ? void 0 : xfrm["@_rot"]));
        const target = resolveRelationship(options.rels, embed);
        if (!target) {
            continue;
        }
        let normalizedTarget = normalizeTargetPath(target);
        let extension = node_path_1.default.posix.extname(normalizedTarget).toLowerCase() || ".png";
        const imageBaseName = `slide-${options.slideIndex}-img-${imageCount}`;
        let data = await options.zipReadFile(normalizedTarget);
        if (extension === ".png" && data.length < 2000) {
            const svgTarget = findSvgAlternative(normalizedTarget, options.zipFileExists);
            if (svgTarget) {
                normalizedTarget = svgTarget;
                extension = ".svg";
                data = await options.zipReadFile(normalizedTarget);
            }
        }
        const { src } = await saveImageAsset({
            extension,
            data,
            imageBaseName,
            imagesDir: options.imagesDir,
        });
        const element = {
            id: `i${imageCount}`,
            type: "image",
            x: emuToPx(Number((_k = off === null || off === void 0 ? void 0 : off["@_x"]) !== null && _k !== void 0 ? _k : 0)),
            y: emuToPx(Number((_l = off === null || off === void 0 ? void 0 : off["@_y"]) !== null && _l !== void 0 ? _l : 0)),
            width: emuToPx(Number((_m = ext === null || ext === void 0 ? void 0 : ext["@_cx"]) !== null && _m !== void 0 ? _m : 0)),
            height: emuToPx(Number((_o = ext === null || ext === void 0 ? void 0 : ext["@_cy"]) !== null && _o !== void 0 ? _o : 0)),
            rotation,
            src,
            objectFit: "cover",
        };
        elements.push(element);
    }
    return {
        elements,
        stats: {
            textElements: textCount,
            schemeClrElements: schemeClrCount,
            multistyleElements: multistyleCount,
            imageElements: imageCount,
        },
    };
}
function ensureArray(value) {
    if (!value) {
        return [];
    }
    return Array.isArray(value) ? value : [value];
}
function extractTextAndStyle(textBody, theme) {
    const body = textBody;
    const paragraphs = ensureArray(body["a:p"]);
    const chunks = paragraphs.map((paragraph) => {
        const p = paragraph;
        const runs = ensureArray(p === null || p === void 0 ? void 0 : p["a:r"]);
        const textRuns = runs
            .map((run) => run === null || run === void 0 ? void 0 : run["a:t"])
            .filter((value) => typeof value === "string");
        return textRuns.join("");
    });
    const text = chunks.join("\n").trim();
    let baseRun = null;
    let firstRun = null;
    const runStyles = [];
    for (const paragraph of paragraphs) {
        const p = paragraph;
        const pPr = p === null || p === void 0 ? void 0 : p["a:pPr"];
        const defRPr = pPr === null || pPr === void 0 ? void 0 : pPr["a:defRPr"];
        const runs = ensureArray(p === null || p === void 0 ? void 0 : p["a:r"]);
        for (const run of runs) {
            const r = run;
            const rPr = r === null || r === void 0 ? void 0 : r["a:rPr"];
            const hasRPr = Boolean(rPr && Object.keys(rPr).length > 0);
            if (!firstRun) {
                firstRun = { rPr, defRPr, pPr };
            }
            if (!baseRun && hasRPr) {
                baseRun = { rPr, defRPr, pPr };
            }
            const style = resolveRunStyle(rPr, defRPr, theme);
            runStyles.push(style);
        }
    }
    const resolvedBase = baseRun !== null && baseRun !== void 0 ? baseRun : firstRun;
    const baseStyle = resolveRunStyle(resolvedBase === null || resolvedBase === void 0 ? void 0 : resolvedBase.rPr, resolvedBase === null || resolvedBase === void 0 ? void 0 : resolvedBase.defRPr, theme);
    const { align, lineHeight } = (0, theme_1.resolveParagraphProps)(resolvedBase === null || resolvedBase === void 0 ? void 0 : resolvedBase.pPr, baseStyle.fontSizePt);
    const style = {
        fontFamily: baseStyle.fontFamily,
        fontSizePt: baseStyle.fontSizePt,
        color: baseStyle.color,
        bold: baseStyle.bold,
        italic: baseStyle.italic,
        underline: baseStyle.underline,
        align,
        lineHeight,
    };
    const multistyle = runStyles.some((runStyle) => isStyleDifferent(runStyle, baseStyle));
    return {
        text,
        style,
        schemeClrUsed: baseStyle.usedScheme,
        multistyle,
    };
}
function resolveRunStyle(rPr, defRPr, theme) {
    var _a, _b;
    const fontSizePt = (_a = (0, theme_1.resolveFontSize)(rPr, defRPr)) !== null && _a !== void 0 ? _a : 28;
    const fontFamily = (_b = (0, theme_1.resolveFontFamily)(rPr, defRPr)) !== null && _b !== void 0 ? _b : "Arial";
    const bold = (0, theme_1.resolveBooleanAttr)(rPr, defRPr, "@_b");
    const italic = (0, theme_1.resolveBooleanAttr)(rPr, defRPr, "@_i");
    const underline = (0, theme_1.resolveUnderline)(rPr, defRPr);
    const { color, usedScheme } = (0, theme_1.resolveColor)(rPr, defRPr, theme);
    return {
        fontFamily,
        fontSizePt,
        color,
        bold,
        italic,
        underline,
        usedScheme,
    };
}
function isStyleDifferent(a, b) {
    return (a.fontFamily !== b.fontFamily ||
        a.fontSizePt !== b.fontSizePt ||
        a.color !== b.color ||
        a.bold !== b.bold ||
        a.italic !== b.italic ||
        a.underline !== b.underline);
}
function resolveRelationship(rels, id) {
    var _a, _b;
    if (!((_a = rels === null || rels === void 0 ? void 0 : rels.Relationships) === null || _a === void 0 ? void 0 : _a.Relationship)) {
        return null;
    }
    const relationships = ensureArray(rels.Relationships.Relationship);
    const match = relationships.find((rel) => rel["@_Id"] === id);
    return (_b = match === null || match === void 0 ? void 0 : match["@_Target"]) !== null && _b !== void 0 ? _b : null;
}
function normalizeTargetPath(target) {
    const normalized = target.replace(/^..\//, "");
    if (normalized.startsWith("ppt/")) {
        return normalized;
    }
    return node_path_1.default.posix.join("ppt", normalized);
}
function findSvgAlternative(pngTarget, fileExists) {
    const ext = node_path_1.default.posix.extname(pngTarget).toLowerCase();
    if (ext !== ".png") {
        return null;
    }
    const base = pngTarget.slice(0, -ext.length);
    const match = base.match(/^(.*)-\d+$/);
    if (match) {
        const candidate = `${match[1]}-3.svg`;
        if (fileExists(candidate)) {
            return candidate;
        }
    }
    const direct = `${base}.svg`;
    if (fileExists(direct)) {
        return direct;
    }
    return null;
}
async function saveImageAsset(options) {
    const { extension, data, imageBaseName, imagesDir } = options;
    if (extension === ".svg") {
        const imageName = `${imageBaseName}${extension}`;
        const imagePath = node_path_1.default.join(imagesDir, imageName);
        await promises_1.default.writeFile(imagePath, data);
        return { src: node_path_1.default.posix.join("assets/images", imageName) };
    }
    const imageName = `${imageBaseName}${extension}`;
    const imagePath = node_path_1.default.join(imagesDir, imageName);
    await promises_1.default.writeFile(imagePath, data);
    return { src: node_path_1.default.posix.join("assets/images", imageName) };
}
