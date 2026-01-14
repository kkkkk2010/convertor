"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.readPptx = readPptx;
exports.readXml = readXml;
exports.listSlidePaths = listSlidePaths;
exports.getSlideRelsPath = getSlideRelsPath;
const promises_1 = __importDefault(require("node:fs/promises"));
const node_path_1 = __importDefault(require("node:path"));
const jszip_1 = __importDefault(require("jszip"));
const fast_xml_parser_1 = require("fast-xml-parser");
const parser = new fast_xml_parser_1.XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
});
async function readPptx(pptxPath) {
    const buffer = await promises_1.default.readFile(pptxPath);
    const zip = await jszip_1.default.loadAsync(buffer);
    return { zip };
}
async function readXml(zip, zipPath) {
    const file = zip.file(zipPath);
    if (!file) {
        throw new Error(`Missing file in pptx: ${zipPath}`);
    }
    const xml = await file.async("string");
    return parser.parse(xml);
}
function listSlidePaths(zip) {
    const slides = zip
        .file(/^ppt\/slides\/slide\d+\.xml$/)
        .map((file) => file.name);
    return slides.sort((a, b) => extractSlideNumber(a) - extractSlideNumber(b));
}
function getSlideRelsPath(slidePath) {
    const fileName = node_path_1.default.posix.basename(slidePath);
    return `ppt/slides/_rels/${fileName}.rels`;
}
function extractSlideNumber(slidePath) {
    const match = slidePath.match(/slide(\d+)\.xml$/);
    if (!match) {
        return Number.MAX_SAFE_INTEGER;
    }
    return Number(match[1]);
}
