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
        const text = extractText(textBody);
        const element = {
            id: `t${textCount}`,
            type: "text",
            x: emuToPx(Number((_d = off === null || off === void 0 ? void 0 : off["@_x"]) !== null && _d !== void 0 ? _d : 0)),
            y: emuToPx(Number((_e = off === null || off === void 0 ? void 0 : off["@_y"]) !== null && _e !== void 0 ? _e : 0)),
            width: emuToPx(Number((_f = ext === null || ext === void 0 ? void 0 : ext["@_cx"]) !== null && _f !== void 0 ? _f : 0)),
            height: emuToPx(Number((_g = ext === null || ext === void 0 ? void 0 : ext["@_cy"]) !== null && _g !== void 0 ? _g : 0)),
            rotation,
            text,
            style: {
                fontFamily: "Arial",
                fontSize: 28,
                color: "#111111",
                bold: false,
                italic: false,
                underline: false,
                align: "left",
            },
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
        const normalizedTarget = normalizeTargetPath(target);
        const extension = node_path_1.default.posix.extname(normalizedTarget) || ".png";
        const imageName = `slide-${options.slideIndex}-img-${imageCount}${extension}`;
        const imagePath = node_path_1.default.join(options.imagesDir, imageName);
        const data = await options.zipReadFile(normalizedTarget);
        await promises_1.default.writeFile(imagePath, data);
        const element = {
            id: `i${imageCount}`,
            type: "image",
            x: emuToPx(Number((_k = off === null || off === void 0 ? void 0 : off["@_x"]) !== null && _k !== void 0 ? _k : 0)),
            y: emuToPx(Number((_l = off === null || off === void 0 ? void 0 : off["@_y"]) !== null && _l !== void 0 ? _l : 0)),
            width: emuToPx(Number((_m = ext === null || ext === void 0 ? void 0 : ext["@_cx"]) !== null && _m !== void 0 ? _m : 0)),
            height: emuToPx(Number((_o = ext === null || ext === void 0 ? void 0 : ext["@_cy"]) !== null && _o !== void 0 ? _o : 0)),
            rotation,
            src: node_path_1.default.posix.join("assets/images", imageName),
            objectFit: "cover",
        };
        elements.push(element);
    }
    return elements;
}
function ensureArray(value) {
    if (!value) {
        return [];
    }
    return Array.isArray(value) ? value : [value];
}
function extractText(textBody) {
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
    return chunks.join("\n").trim();
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
