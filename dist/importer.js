"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const promises_1 = __importDefault(require("node:fs/promises"));
const node_fs_1 = __importDefault(require("node:fs"));
const node_path_1 = __importDefault(require("node:path"));
const node_os_1 = __importDefault(require("node:os"));
const read_1 = require("./pptx/read");
const parse_slide_1 = require("./pptx/parse_slide");
const backgrounds_1 = require("./render/backgrounds");
const clean_pptx_1 = require("./pptx/clean_pptx");
const theme_1 = require("./pptx/theme");
const archiver_1 = __importDefault(require("archiver"));
async function main() {
    try {
        const { input, outRaw } = parseArgs(process.argv.slice(2));
        await ensureInput(input);
        const outIsZip = outRaw.toLowerCase().endsWith(".zip");
        const outZipPath = outIsZip
            ? node_path_1.default.resolve(outRaw)
            : node_path_1.default.join(node_path_1.default.resolve(outRaw), "out.zip");
        const outDir = outIsZip
            ? node_path_1.default.dirname(outZipPath)
            : node_path_1.default.resolve(outRaw);
        console.log(`outDir: ${outDir}`);
        console.log(`outZipPath: ${outZipPath}`);
        console.log(`outIsZip: ${outIsZip}`);
        await promises_1.default.mkdir(outDir, { recursive: true });
        if (node_fs_1.default.existsSync(outZipPath)) {
            const stat = node_fs_1.default.statSync(outZipPath);
            if (stat.isDirectory()) {
                throw new Error(`Expected a .zip FILE but path is a DIRECTORY: ${outZipPath}`);
            }
            if (stat.isFile()) {
                node_fs_1.default.unlinkSync(outZipPath);
            }
        }
        const tempOutDir = node_fs_1.default.mkdtempSync(node_path_1.default.join(node_os_1.default.tmpdir(), "pptx-import-"));
        const archive = await (0, read_1.readPptx)(input);
        const slidePaths = (0, read_1.listSlidePaths)(archive.zip);
        if (slidePaths.length === 0) {
            throw new Error("No slides found in PPTX.");
        }
        let themeColors = {};
        try {
            const themeXml = await (0, read_1.readXml)(archive.zip, "ppt/theme/theme1.xml");
            themeColors = (0, theme_1.parseThemeColors)(themeXml);
        }
        catch {
            themeColors = {};
        }
        const presentation = await (0, read_1.readXml)(archive.zip, "ppt/presentation.xml");
        const slideSize = extractSlideSize(presentation);
        const assetsDir = node_path_1.default.join(tempOutDir, "assets/images");
        await promises_1.default.mkdir(assetsDir, { recursive: true });
        console.log(`Slides found: ${slidePaths.length}`);
        let totalTextElements = 0;
        let totalSchemeClr = 0;
        let totalMultistyle = 0;
        let totalImages = 0;
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
            const { elements, stats } = await (0, parse_slide_1.parseSlide)(slideXml, {
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
                zipFileExists: (zipPath) => Boolean(archive.zip.file(zipPath)),
                imagesDir: assetsDir,
                theme: themeColors,
            });
            const textCount = elements.filter((el) => el.type === "text").length;
            const imageCount = elements.filter((el) => el.type === "image").length;
            console.log(`Slide ${slideIndex}: ${textCount} text, ${imageCount} images`);
            totalTextElements += stats.textElements;
            totalSchemeClr += stats.schemeClrElements;
            totalMultistyle += stats.multistyleElements;
            totalImages += stats.imageElements;
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
        console.log(`Text elements: ${totalTextElements}`);
        console.log(`Text elements with schemeClr: ${totalSchemeClr}`);
        console.log(`MULTISTYLE: ${totalMultistyle}`);
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
        await (0, backgrounds_1.renderBackgrounds)(backgroundPptxPath, tempOutDir);
        await promises_1.default.writeFile(node_path_1.default.join(tempOutDir, "doc.json"), JSON.stringify(doc, null, 2), "utf-8");
        const zipSize = await buildZip(tempOutDir, outZipPath);
        const zipStat = node_fs_1.default.statSync(outZipPath);
        if (!zipStat.isFile() || zipStat.size <= 1024) {
            throw new Error(`Expected zip file larger than 1KB, got ${zipStat.size} bytes at ${outZipPath}`);
        }
        await promises_1.default.rm(tempOutDir, { recursive: true, force: true });
        await promises_1.default.rm(tmpDir, { recursive: true, force: true });
        console.log(`wrote zip file size=${zipSize}`);
        console.log(`✔ Exported presentation to ${node_path_1.default.basename(outZipPath)} (${slidePaths.length} slides, ${totalImages} images)`);
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
    const outRaw = args.get("--out");
    if (!input || !outRaw) {
        throw new Error("Usage: node dist/importer.js --input <path/to/input.pptx> --out <path/to/out.zip>");
    }
    return { input, outRaw };
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
async function buildZip(sourceDir, outZipPath) {
    const output = node_fs_1.default.createWriteStream(outZipPath);
    const archive = (0, archiver_1.default)("zip", { zlib: { level: 9 } });
    archive.directory(sourceDir, false);
    archive.pipe(output);
    await new Promise((resolve, reject) => {
        output.on("close", () => resolve());
        output.on("error", (error) => reject(error));
        archive.on("error", (error) => reject(error));
        void archive.finalize();
    });
    return archive.pointer();
}
void main();
