import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { ConversionLimits, getConversionLimitsFromEnv } from "../limits";

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
});

export type PptxArchive = {
  zip: JSZip;
};

export async function readPptx(
  pptxPath: string,
  limits: ConversionLimits = getConversionLimitsFromEnv(),
): Promise<PptxArchive> {
  const stat = await fs.stat(pptxPath);
  if (!stat.isFile()) {
    throw new Error("Input path is not a file.");
  }
  if (stat.size > limits.maxInputBytes) {
    throw new Error(
      `PPTX exceeds max input size (${stat.size} bytes > ${limits.maxInputBytes} bytes).`,
    );
  }
  const buffer = await fs.readFile(pptxPath);
  return readPptxBuffer(buffer, limits);
}

export async function readPptxBuffer(
  buffer: Buffer,
  limits: ConversionLimits = getConversionLimitsFromEnv(),
): Promise<PptxArchive> {
  if (buffer.length > limits.maxInputBytes) {
    throw new Error(
      `PPTX exceeds max input size (${buffer.length} bytes > ${limits.maxInputBytes} bytes).`,
    );
  }
  const zip = await JSZip.loadAsync(buffer);
  await validateZipLimits(zip, limits);
  return { zip };
}

export async function readXml<T>(zip: JSZip, zipPath: string): Promise<T> {
  const file = zip.file(zipPath);
  if (!file) {
    throw new Error(`Missing file in pptx: ${zipPath}`);
  }
  const xml = await file.async("string");
  return parser.parse(xml) as T;
}

export function listSlidePaths(zip: JSZip): string[] {
  const slides = zip
    .file(/^ppt\/slides\/slide\d+\.xml$/)
    .map((file: JSZip.JSZipObject) => file.name);
  return slides.sort(
    (a: string, b: string) => extractSlideNumber(a) - extractSlideNumber(b),
  );
}

export function getSlideRelsPath(slidePath: string): string {
  const fileName = path.posix.basename(slidePath);
  return `ppt/slides/_rels/${fileName}.rels`;
}

function extractSlideNumber(slidePath: string): number {
  const match = slidePath.match(/slide(\d+)\.xml$/);
  if (!match) {
    return Number.MAX_SAFE_INTEGER;
  }
  return Number(match[1]);
}

async function validateZipLimits(
  zip: JSZip,
  limits: ConversionLimits,
): Promise<void> {
  const files = Object.values(zip.files).filter(
    (file): file is JSZip.JSZipObject => !file.dir,
  );
  if (files.length > limits.maxZipEntries) {
    throw new Error(
      `PPTX exceeds max zip entries (${files.length} > ${limits.maxZipEntries}).`,
    );
  }

  let totalUncompressed = 0;
  for (const file of files) {
    const size = await getUncompressedSize(file);
    if (size > limits.maxFileUncompressedBytes) {
      throw new Error(
        `PPTX entry too large (${file.name}: ${size} bytes > ${limits.maxFileUncompressedBytes} bytes).`,
      );
    }
    totalUncompressed += size;
    if (totalUncompressed > limits.maxTotalUncompressedBytes) {
      throw new Error(
        `PPTX exceeds max uncompressed size (${totalUncompressed} bytes > ${limits.maxTotalUncompressedBytes} bytes).`,
      );
    }
  }
}

async function getUncompressedSize(file: JSZip.JSZipObject): Promise<number> {
  const data = (file as { _data?: { uncompressedSize?: number | string } })
    ._data;
  if (data?.uncompressedSize != null) {
    return Number(data.uncompressedSize);
  }
  const content = await file.async("nodebuffer");
  return content.length;
}
