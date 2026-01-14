"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const promises_1 = __importDefault(require("node:fs/promises"));
const node_path_1 = __importDefault(require("node:path"));
const node_os_1 = __importDefault(require("node:os"));
const read_1 = require("./pptx/read");
const parse_slide_1 = require("./pptx/parse_slide");
const backgrounds_1 = require("./render/backgrounds");
const clean_pptx_1 = require("./pptx/clean_pptx");
async function main() {
    try {
        const { input, outDir } = parseArgs(process.argv.slice(2));
        await ensureInput(input);
        await promises_1.default.mkdir(outDir, { recursive: true });
        const archive = await (0, read_1.readPptx)(input);
        const slidePaths = (0, read_1.listSlidePaths)(archive.zip);
        if (slidePaths.length === 0) {
            throw new Error("No slides found in PPTX.");
        }
        const presentation = await (0, read_1.readXml)(archive.zip, "ppt/presentation.xml");
        const slideSize = extractSlideSize(presentation);
        const assetsDir = node_path_1.default.join(outDir, "assets/images");
        await promises_1.default.mkdir(assetsDir, { recursive: true });
        console.log(`Slides found: ${slidePaths.length}`);
        const slides = [];
        for (let i = 0; i < slidePaths.length; i += 1) {
            const slidePath = slidePaths[i];
            const slideIndex = i + 1;
            const slideXml = await (0, read_1.readXml)(archive.zip, slidePath);
            const relsPath = (0, read_1.getSlideRelsPath)(slidePath);
            let relsXml = null;
            try {
                relsXml = await (0, read_1.readXml)(archive.zip, relsPath);
            }
            catch {
                relsXml = null;
            }
            const elements = await (0, parse_slide_1.parseSlide)(slideXml, {
                slideIndex,
                rels: relsXml,
                zipReadFile: async (zipPath) => {
                    const file = archive.zip.file(zipPath);
                    if (!file) {
                        throw new Error(`Missing media file: ${zipPath}`);
                    }
                    const data = await file.async("nodebuffer");
                    return Buffer.from(data);
                },
                imagesDir: assetsDir,
            });
            const textCount = elements.filter((el) => el.type === "text").length;
            const imageCount = elements.filter((el) => el.type === "image").length;
            console.log(`Slide ${slideIndex}: ${textCount} text, ${imageCount} images`);
            slides.push({
                id: `s${slideIndex}`,
                background: {
                    type: "image",
                    src: node_path_1.default.posix.join("backgrounds", `slide-${slideIndex}.png`),
                },
                elements,
            });
        }
        const doc = {
            schemaVersion: 1,
            slideSize,
            slides,
        };
        const tmpDir = await promises_1.default.mkdtemp(node_path_1.default.join(node_os_1.default.tmpdir(), "pptx-import-clean-"));
        const cleanedPptxPath = node_path_1.default.join(tmpDir, "cleaned.pptx");
        let backgroundPptxPath = input;
        try {
            await (0, clean_pptx_1.createCleanPptx)(archive.zip, slidePaths, cleanedPptxPath);
            backgroundPptxPath = cleanedPptxPath;
        }
        catch (cleanError) {
            const message = cleanError instanceof Error
                ? cleanError.message
                : "Unknown error";
            console.warn(`Warning: failed to create cleaned PPTX, falling back to original. ${message}`);
        }
        await (0, backgrounds_1.renderBackgrounds)(backgroundPptxPath, outDir);
        await promises_1.default.writeFile(node_path_1.default.join(outDir, "doc.json"), JSON.stringify(doc, null, 2), "utf-8");
    }
    catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error occurred.";
        console.error(message);
        process.exit(1);
    }
}
function parseArgs(argv) {
    const args = new Map();
    for (let i = 0; i < argv.length; i += 1) {
        const arg = argv[i];
        if (arg.startsWith("--")) {
            args.set(arg, argv[i + 1]);
            i += 1;
        }
    }
    const input = args.get("--input");
    const outDir = args.get("--out");
    if (!input || !outDir) {
        throw new Error("Usage: node dist/importer.js --input <path/to/input.pptx> --out <path/to/outDir>");
    }
    return { input, outDir };
}
async function ensureInput(inputPath) {
    try {
        const stat = await promises_1.default.stat(inputPath);
        if (!stat.isFile()) {
            throw new Error("Input path is not a file.");
        }
    }
    catch {
        throw new Error(`Input file not found: ${inputPath}`);
    }
}
function extractSlideSize(presentation) {
    var _a;
    const size = (_a = presentation["p:presentation"]) === null || _a === void 0 ? void 0 : _a["p:sldSz"];
    if (!(size === null || size === void 0 ? void 0 : size["@_cx"]) || !(size === null || size === void 0 ? void 0 : size["@_cy"])) {
        return {
            width: 13.333,
            height: 7.5,
            unit: "in",
        };
    }
    const widthPx = (0, parse_slide_1.emuToPx)(Number(size["@_cx"]));
    const heightPx = (0, parse_slide_1.emuToPx)(Number(size["@_cy"]));
    return {
        width: Number((widthPx / 96).toFixed(3)),
        height: Number((heightPx / 96).toFixed(3)),
        unit: "in",
    };
}
void main();
