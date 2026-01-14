"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.createCleanPptx = createCleanPptx;
const promises_1 = __importDefault(require("node:fs/promises"));
const node_path_1 = __importDefault(require("node:path"));
const xmldom_1 = require("@xmldom/xmldom");
async function createCleanPptx(zip, slidePaths, outPath) {
    for (const slidePath of slidePaths) {
        const file = zip.file(slidePath);
        if (!file) {
            continue;
        }
        const xml = await file.async("string");
        const cleaned = cleanSlideXml(xml);
        zip.file(slidePath, cleaned);
    }
    await promises_1.default.mkdir(node_path_1.default.dirname(outPath), { recursive: true });
    const buffer = await zip.generateAsync({ type: "nodebuffer" });
    await promises_1.default.writeFile(outPath, buffer);
}
function cleanSlideXml(xml) {
    var _a, _b;
    const doc = new xmldom_1.DOMParser().parseFromString(xml, "text/xml");
    const pics = Array.from(doc.getElementsByTagName("p:pic"));
    for (const node of pics) {
        (_a = node.parentNode) === null || _a === void 0 ? void 0 : _a.removeChild(node);
    }
    const shapes = Array.from(doc.getElementsByTagName("p:sp"));
    for (const sp of shapes) {
        const hasTextBody = sp.getElementsByTagName("p:txBody").length > 0;
        if (hasTextBody) {
            (_b = sp.parentNode) === null || _b === void 0 ? void 0 : _b.removeChild(sp);
        }
    }
    return new xmldom_1.XMLSerializer().serializeToString(doc);
}
