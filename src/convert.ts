import fs from "node:fs/promises";
import fsSync from "node:fs";
import path from "node:path";
import os from "node:os";
import { PassThrough } from "node:stream";
import archiver from "archiver";
import { readPptxBuffer, readXml, listSlidePaths, getSlideRelsPath } from "./pptx/read";
import { parseSlide, emuToPx } from "./pptx/parse_slide";
import { renderBackgrounds } from "./render/backgrounds";
import { DocJson } from "./types";
import { createCleanPptx } from "./pptx/clean_pptx";
import { ThemeColorMap, parseThemeColors } from "./pptx/theme";
import { ConversionLimits, getConversionLimitsFromEnv } from "./limits";

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

export type ConversionDependencies = {
  renderBackgrounds: typeof renderBackgrounds;
};

export type ConversionResult = {
  zipBuffer: Buffer;
  slideCount: number;
  totalImages: number;
};

export async function convertPptxToOutZip(
  inputBuffer: Buffer,
): Promise<Buffer> {
  const result = await convertPptxToOutZipWithDependencies(inputBuffer, {
    renderBackgrounds,
  });
  return result.zipBuffer;
}

export async function convertPptxToOutZipWithDependencies(
  inputBuffer: Buffer,
  dependencies: ConversionDependencies,
  limits: ConversionLimits = getConversionLimitsFromEnv(),
): Promise<ConversionResult> {
  if (inputBuffer.length === 0) {
    throw new Error("Input buffer is empty.");
  }
  if (inputBuffer.length > limits.maxInputBytes) {
    throw new Error(
      `PPTX exceeds max input size (${inputBuffer.length} bytes > ${limits.maxInputBytes} bytes).`,
    );
  }

  const tempOutDir = await fs.mkdtemp(path.join(os.tmpdir(), "pptx-import-"));
  const tempWorkDir = await fs.mkdtemp(
    path.join(os.tmpdir(), "pptx-import-work-"),
  );
  const tempInputPath = path.join(tempWorkDir, "input.pptx");
  await fs.writeFile(tempInputPath, inputBuffer);

  try {
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

    const cleanedPptxPath = path.join(tempWorkDir, "cleaned.pptx");
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

    const zipBuffer = await buildZipBuffer(tempOutDir);
    if (zipBuffer.length <= 1024) {
      throw new Error(
        `Expected zip file larger than 1KB, got ${zipBuffer.length} bytes.`,
      );
    }
    console.log(`wrote zip file size=${zipBuffer.length}`);
    return {
      zipBuffer,
      slideCount: slidePaths.length,
      totalImages,
    };
  } finally {
    await fs.rm(tempOutDir, { recursive: true, force: true });
    await fs.rm(tempWorkDir, { recursive: true, force: true });
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

async function buildZipBuffer(sourceDir: string): Promise<Buffer> {
  const archive = archiver("zip", { zlib: { level: 9 } });
  const output = new PassThrough();
  const chunks: Buffer[] = [];

  output.on("data", (chunk: Buffer) => chunks.push(chunk));

  const completion = new Promise<Buffer>((resolve, reject) => {
    output.on("end", () => resolve(Buffer.concat(chunks)));
    output.on("error", (error: Error) => reject(error));
    archive.on("error", (error: Error) => reject(error));
  });

  archive.directory(sourceDir, false);
  archive.pipe(output);
  void archive.finalize();

  const buffer = await completion;
  if (buffer.length === 0) {
    throw new Error("Failed to build zip buffer.");
  }

  const sizeCheck = fsSync.existsSync(sourceDir);
  if (!sizeCheck) {
    throw new Error("Source directory missing while building zip buffer.");
  }

  return buffer;
}
