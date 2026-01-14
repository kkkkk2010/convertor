"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.renderBackgrounds = renderBackgrounds;
const promises_1 = __importDefault(require("node:fs/promises"));
const node_path_1 = __importDefault(require("node:path"));
const node_os_1 = __importDefault(require("node:os"));
const node_child_process_1 = require("node:child_process");
const node_util_1 = require("node:util");
const execFileAsync = (0, node_util_1.promisify)(node_child_process_1.execFile);
async function renderBackgrounds(inputPptx, outDir) {
    const libreOfficeBinary = resolveLibreOfficeBinary();
    if (!libreOfficeBinary) {
        const candidates = getLibreOfficeCandidates();
        const hintText = [
            "Install LibreOffice and ensure it is on PATH.",
            `Tried: ${candidates.join(", ")}`,
        ];
        throw new Error(`Missing dependency: LibreOffice.\n${hintText
            .map((hint) => `- ${hint}`)
            .join("\n")}`);
    }
    await ensureBinary("pdftoppm", [
        "Install poppler-utils (pdftoppm) and ensure it is on PATH.",
    ]);
    const backgroundsDir = node_path_1.default.join(outDir, "backgrounds");
    await promises_1.default.mkdir(backgroundsDir, { recursive: true });
    const tmpDir = await promises_1.default.mkdtemp(node_path_1.default.join(node_os_1.default.tmpdir(), "pptx-import-"));
    const pptxName = node_path_1.default.basename(inputPptx);
    const pdfName = pptxName.replace(/\.pptx$/i, ".pdf");
    const pdfPath = node_path_1.default.join(tmpDir, pdfName);
    await execFileAsync(libreOfficeBinary, [
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        tmpDir,
        inputPptx,
    ]);
    await execFileAsync("pdftoppm", [
        "-png",
        "-r",
        "144",
        pdfPath,
        node_path_1.default.join(backgroundsDir, "slide"),
    ]);
    await normalizeBackgroundNames(backgroundsDir);
}
function resolveLibreOfficeBinary() {
    const candidates = getLibreOfficeCandidates();
    for (const candidate of candidates) {
        if (isBinaryAvailable(candidate)) {
            return candidate;
        }
    }
    return null;
}
function getLibreOfficeCandidates() {
    if (node_os_1.default.platform() !== "win32") {
        return ["libreoffice"];
    }
    return [
        "soffice",
        "soffice.com",
        "soffice.exe",
        "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
        "C:\\Program Files\\LibreOffice\\program\\soffice.com",
        "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
        "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.com",
    ];
}
async function ensureBinary(command, hints) {
    if (isBinaryAvailable(command)) {
        return;
    }
    const hintText = hints.map((hint) => `- ${hint}`).join("\n");
    throw new Error(`Missing dependency: ${command}.\n${hintText}`);
}
function isBinaryAvailable(command) {
    const result = (0, node_child_process_1.spawnSync)(command, ["--version"], { stdio: "ignore" });
    return result.error == null;
}
async function normalizeBackgroundNames(backgroundsDir) {
    const files = await promises_1.default.readdir(backgroundsDir);
    const slideFiles = files
        .filter((file) => file.startsWith("slide-") && file.endsWith(".png"))
        .sort((a, b) => extractNumber(a) - extractNumber(b));
    await Promise.all(slideFiles.map(async (file, index) => {
        const target = `slide-${index + 1}.png`;
        if (file === target) {
            return;
        }
        await promises_1.default.rename(node_path_1.default.join(backgroundsDir, file), node_path_1.default.join(backgroundsDir, target));
    }));
}
function extractNumber(fileName) {
    const match = fileName.match(/slide-(\d+)/);
    if (!match) {
        return Number.MAX_SAFE_INTEGER;
    }
    return Number(match[1]);
}
