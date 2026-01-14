import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { readPptx, readXml, listSlidePaths, getSlideRelsPath } from "./pptx/read";
import { parseSlide, emuToPx } from "./pptx/parse_slide";
import { renderBackgrounds } from "./render/backgrounds";
import { DocJson } from "./types";
import { createCleanPptx } from "./pptx/clean_pptx";

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
    await fs.mkdir(outDir, { recursive: true });

    const archive = await readPptx(input);
    const slidePaths = listSlidePaths(archive.zip);
    if (slidePaths.length === 0) {
      throw new Error("No slides found in PPTX.");
    }

    const presentation = await readXml<PresentationXml>(
      archive.zip,
      "ppt/presentation.xml",
    );
    const slideSize = extractSlideSize(presentation);

    const assetsDir = path.join(outDir, "assets/images");
    const originalsDir = path.join(outDir, "assets/original");
    await fs.mkdir(assetsDir, { recursive: true });
    await fs.mkdir(originalsDir, { recursive: true });

    console.log(`Slides found: ${slidePaths.length}`);

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

      const elements = await parseSlide(slideXml, {
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
        originalsDir,
      });

      const textCount = elements.filter((el) => el.type === "text").length;
      const imageCount = elements.filter((el) => el.type === "image").length;
      console.log(
        `Slide ${slideIndex}: ${textCount} text, ${imageCount} images`,
      );

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

    await renderBackgrounds(backgroundPptxPath, outDir);

    await fs.writeFile(
      path.join(outDir, "doc.json"),
      JSON.stringify(doc, null, 2),
      "utf-8",
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
      "Usage: node dist/importer.js --input <path/to/input.pptx> --out <path/to/outDir>",
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

void main();
