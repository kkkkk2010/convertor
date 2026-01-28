import fs from "node:fs/promises";
import fsSync from "node:fs";
import os from "node:os";
import path from "node:path";
import archiver from "archiver";
import {
  getSlideRelsPath,
  listSlidePaths,
  readBinary,
  readPptxBuffer,
  readXml,
} from "./pptx/read";
import { parseSlide, emuToPx } from "./pptx/parse_slide";
import { renderBackgrounds } from "./render/backgrounds";
import { DocJson } from "./types";
import { createCleanPptx } from "./pptx/clean_pptx";
import { ThemeColorMap, parseThemeColors } from "./pptx/theme";
import { getLimits } from "./limits";

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

type ConvertDependencies = {
  renderBackgrounds: typeof renderBackgrounds;
};

const defaultDependencies: ConvertDependencies = {
  renderBackgrounds,
};

export type ConversionStats = {
  slideCount: number;
  imageCount: number;
};

export type ConversionResult = Buffer & {
  conversionStats?: ConversionStats;
};

export async function convertPptxToOutZip(
  inputBuffer: Buffer,
): Promise<ConversionResult> {
  return convertPptxToOutZipInternal(inputBuffer, defaultDependencies);
}

export async function convertPptxToOutZipInternal(
  inputBuffer: Buffer,
  dependencies: ConvertDependencies,
): Promise<ConversionResult> {
  const limits = getLimits();
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "pptx-import-"));
  const tempOutDir = path.join(tempRoot, "out");
  const tempInputPath = path.join(tempRoot, "input.pptx");
  const tempZipPath = path.join(tempRoot, "out.zip");
  const cleanDir = await fs.mkdtemp(path.join(os.tmpdir(), "pptx-import-clean-"));

  try {
    await fs.mkdir(tempOutDir, { recursive: true });
    await fs.writeFile(tempInputPath, inputBuffer);

    const archive = await readPptxBuffer(inputBuffer, limits);
    const slidePaths = listSlidePaths(archive.zip);
    if (slidePaths.length === 0) {
      throw new Error("No slides found in PPTX.");
    }

    let themeColors: ThemeColorMap = {};
    try {
      const themeXml = await readXml<Record<string, unknown>>(
        archive.zip,
        "ppt/theme/theme1.xml",
        limits,
      );
      themeColors = parseThemeColors(themeXml);
    } catch {
      themeColors = {};
    }

    const presentation = await readXml<PresentationXml>(
      archive.zip,
      "ppt/presentation.xml",
      limits,
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
      const slideXml = await readXml<SlideXml>(archive.zip, slidePath, limits);
      const relsPath = getSlideRelsPath(slidePath);
      let relsXml: RelsXml | null = null;
      try {
        relsXml = await readXml<RelsXml>(archive.zip, relsPath, limits);
      } catch {
        relsXml = null;
      }

      const { elements, stats } = await parseSlide(slideXml, {
        slideIndex,
        rels: relsXml as RelsXml,
        zipReadFile: async (zipPath: string) =>
          readBinary(archive.zip, zipPath, limits),
        zipFileExists: (zipPath: string) => Boolean(archive.zip.file(zipPath)),
        imagesDir: assetsDir,
        theme: themeColors,
        debugSvg: isDebugSvgEnabled(),
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

    const cleanedPptxPath = path.join(cleanDir, "cleaned.pptx");
    let backgroundPptxPath = tempInputPath;
    try {
      await createCleanPptx(archive.zip, slidePaths, cleanedPptxPath);
      backgroundPptxPath = cleanedPptxPath;
    } catch (cleanError) {
      const message =
        cleanError instanceof Error ? cleanError.message : "Unknown error";
      console.warn(
        `Warning: failed to create cleaned PPTX, falling back to original. ${message}`,
      );
    }

    await dependencies.renderBackgrounds(backgroundPptxPath, tempOutDir);

    await fs.writeFile(
      path.join(tempOutDir, "doc.json"),
      JSON.stringify(doc, null, 2),
      "utf-8",
    );

    const zipSize = await buildZip(tempOutDir, tempZipPath);
    const zipStat = fsSync.statSync(tempZipPath);
    if (!zipStat.isFile() || zipStat.size <= 1024) {
      throw new Error(
        `Expected zip file larger than 1KB, got ${zipStat.size} bytes at ${tempZipPath}`,
      );
    }
    console.log(`wrote zip file size=${zipSize}`);

    const buffer = (await fs.readFile(tempZipPath)) as ConversionResult;
    buffer.conversionStats = {
      slideCount: slidePaths.length,
      imageCount: totalImages,
    };
    return buffer;
  } finally {
    await fs.rm(tempRoot, { recursive: true, force: true });
    await fs.rm(cleanDir, { recursive: true, force: true });
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

async function buildZip(sourceDir: string, outZipPath: string): Promise<number> {
  const output = fsSync.createWriteStream(outZipPath);
  const archive = archiver("zip", { zlib: { level: 9 } });
  archive.directory(sourceDir, false);
  archive.pipe(output);

  await new Promise<void>((resolve, reject) => {
    output.on("close", () => resolve());
    output.on("error", (error: Error) => reject(error));
    archive.on("error", (error: Error) => reject(error));
    void archive.finalize();
  });

  return archive.pointer();
}

function isDebugSvgEnabled(): boolean {
  const value = process.env.PPTX_IMPORTER_DEBUG_SVG;
  if (!value) {
    return false;
  }
  return value.toLowerCase() === "1" || value.toLowerCase() === "true";
}
