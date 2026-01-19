import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { readPptx, readXml, listSlidePaths, getSlideRelsPath } from "./pptx/read";
import { parseSlide, emuToPx } from "./pptx/parse_slide";
import { renderBackgrounds } from "./render/backgrounds";
import { DocJson } from "./types";
import { createCleanPptx } from "./pptx/clean_pptx";
import { ThemeColorMap, parseThemeColors } from "./pptx/theme";
import JSZip from "jszip";

type PresentationXml = {
  "p:presentation"?: {
    "p:sldSz"?: {
      "@_cx"?: string;
      "@_cy"?: string;
    };
  };
};

type SlideXml = Record<string, unknown>;
type RelsXml = Record<string, unknown>;

async function main() {
  try {
    const { input, outDir } = parseArgs(process.argv.slice(2));
    await ensureInput(input);
    const outZipPath = await resolveOutZipPath(outDir);
    const tempOutDir = await fs.mkdtemp(path.join(os.tmpdir(), "pptx-import-out-"));

    const archive = await readPptx(input);
    const slidePaths = listSlidePaths(archive.zip);
    if (slidePaths.length === 0) {
      throw new Error("No slides found in PPTX.");
    }

    let themeColors: ThemeColorMap = {};
    try {
      const themeXml = await readXml<Record<string, unknown>>(
        archive.zip,
        "ppt/theme/theme1.xml",
      );
      themeColors = parseThemeColors(themeXml);
    } catch {
      themeColors = {};
    }

    const presentation = await readXml<PresentationXml>(
      archive.zip,
      "ppt/presentation.xml",
    );
    const slideSize = extractSlideSize(presentation);

    const assetsDir = path.join(tempOutDir, "assets/images");
    await fs.mkdir(assetsDir, { recursive: true });

    console.log(`Slides found: ${slidePaths.length}`);

    let totalTextElements = 0;
    let totalSchemeClr = 0;
    let totalMultistyle = 0;
    let totalImages = 0;
    const slides = [];
    for (let i = 0; i < slidePaths.length; i += 1) {
      const slidePath = slidePaths[i];
      const slideIndex = i + 1;
      const slideXml = await readXml<SlideXml>(archive.zip, slidePath);
      const relsPath = getSlideRelsPath(slidePath);
      let relsXml: RelsXml | null = null;
      try {
        relsXml = await readXml<RelsXml>(archive.zip, relsPath);
      } catch {
        relsXml = null;
      }

      const { elements, stats } = await parseSlide(slideXml, {
        slideIndex,
        rels: relsXml as RelsXml,
        zipReadFile: async (zipPath: string) => {
          const file = archive.zip.file(zipPath);
          if (!file) {
            throw new Error(`Missing media file: ${zipPath}`);
          }
          const data = await file.async("nodebuffer");
          return Buffer.from(data);
        },
        zipFileExists: (zipPath: string) => Boolean(archive.zip.file(zipPath)),
        imagesDir: assetsDir,
        theme: themeColors,
      });

      const textCount = elements.filter((el) => el.type === "text").length;
      const imageCount = elements.filter((el) => el.type === "image").length;
      console.log(
        `Slide ${slideIndex}: ${textCount} text, ${imageCount} images`,
      );

      totalTextElements += stats.textElements;
      totalSchemeClr += stats.schemeClrElements;
      totalMultistyle += stats.multistyleElements;
      totalImages += stats.imageElements;

      slides.push({
        id: `s${slideIndex}`,
        background: {
          type: "image" as const,
          src: path.posix.join("backgrounds", `slide-${slideIndex}.png`),
        },
        elements,
      });
    }

    const doc: DocJson = {
      schemaVersion: 1,
      slideSize,
      slides,
    };

    console.log(`Text elements: ${totalTextElements}`);
    console.log(`Text elements with schemeClr: ${totalSchemeClr}`);
    console.log(`MULTISTYLE: ${totalMultistyle}`);

    const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "pptx-import-clean-"));
    const cleanedPptxPath = path.join(tmpDir, "cleaned.pptx");
    let backgroundPptxPath = input;
    try {
      await createCleanPptx(archive.zip, slidePaths, cleanedPptxPath);
      backgroundPptxPath = cleanedPptxPath;
    } catch (cleanError) {
      const message =
        cleanError instanceof Error
          ? cleanError.message
          : "Unknown error";
      console.warn(
        `Warning: failed to create cleaned PPTX, falling back to original. ${message}`,
      );
    }

    await renderBackgrounds(backgroundPptxPath, tempOutDir);

    await fs.writeFile(
      path.join(tempOutDir, "doc.json"),
      JSON.stringify(doc, null, 2),
      "utf-8",
    );

    await buildZip(tempOutDir, outZipPath);
    await fs.rm(tempOutDir, { recursive: true, force: true });
    await fs.rm(tmpDir, { recursive: true, force: true });
    console.log(
      `✔ Exported presentation to ${path.basename(outZipPath)} (${slidePaths.length} slides, ${totalImages} images)`,
    );
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Unknown error occurred.";
    console.error(message);
    process.exit(1);
  }
}

function parseArgs(argv: string[]): { input: string; outDir: string } {
  const args = new Map<string, string>();
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
    throw new Error(
      "Usage: node dist/importer.js --input <path/to/input.pptx> --out <path/to/out.zip>",
    );
  }
  return { input, outDir };
}

async function ensureInput(inputPath: string): Promise<void> {
  try {
    const stat = await fs.stat(inputPath);
    if (!stat.isFile()) {
      throw new Error("Input path is not a file.");
    }
  } catch {
    throw new Error(`Input file not found: ${inputPath}`);
  }
}

async function resolveOutZipPath(outPath: string): Promise<string> {
  const hasZipExt = path.extname(outPath).toLowerCase() === ".zip";
  if (hasZipExt) {
    await fs.mkdir(path.dirname(outPath), { recursive: true });
    return outPath;
  }
  try {
    const stat = await fs.stat(outPath);
    if (stat.isDirectory()) {
      return path.join(outPath, "out.zip");
    }
  } catch {
    await fs.mkdir(outPath, { recursive: true });
    return path.join(outPath, "out.zip");
  }
  return `${outPath}.zip`;
}

function extractSlideSize(presentation: PresentationXml): DocJson["slideSize"] {
  const size = presentation["p:presentation"]?.["p:sldSz"];
  if (!size?.["@_cx"] || !size?.["@_cy"]) {
    return {
      width: 13.333,
      height: 7.5,
      unit: "in",
    };
  }
  const widthPx = emuToPx(Number(size["@_cx"]));
  const heightPx = emuToPx(Number(size["@_cy"]));
  return {
    width: Number((widthPx / 96).toFixed(3)),
    height: Number((heightPx / 96).toFixed(3)),
    unit: "in",
  };
}

async function buildZip(sourceDir: string, outZipPath: string): Promise<void> {
  const zip = new JSZip();
  await addDirectoryToZip(zip, sourceDir, "");
  const content = await zip.generateAsync({ type: "nodebuffer" });
  await fs.writeFile(outZipPath, content);
}

async function addDirectoryToZip(
  zip: JSZip,
  dirPath: string,
  prefix: string,
): Promise<void> {
  const entries = await fs.readdir(dirPath, { withFileTypes: true });
  for (const entry of entries) {
    const fullPath = path.join(dirPath, entry.name);
    const zipPath = path.posix.join(prefix, entry.name);
    if (entry.isDirectory()) {
      await addDirectoryToZip(zip, fullPath, zipPath);
    } else if (entry.isFile()) {
      const data = await fs.readFile(fullPath);
      zip.file(zipPath, data);
    }
  }
}

void main();
