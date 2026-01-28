import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { Limits } from "../limits";

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
});

export type PptxArchive = {
  zip: JSZip;
};

export async function readPptx(pptxPath: string): Promise<PptxArchive> {
  const buffer = await fs.readFile(pptxPath);
  const zip = await JSZip.loadAsync(buffer);
  return { zip };
}

export async function readPptxBuffer(
  buffer: Buffer,
  limits: Limits,
): Promise<PptxArchive> {
  if (buffer.length > limits.maxPptxSizeBytes) {
    throw new Error(
      `PPTX size ${buffer.length} exceeds limit ${limits.maxPptxSizeBytes} bytes.`,
    );
  }
  const zip = await JSZip.loadAsync(buffer);
  enforceZipLimits(zip, limits);
  return { zip };
}

export async function readXml<T>(
  zip: JSZip,
  zipPath: string,
  limits?: Limits,
): Promise<T> {
  const file = zip.file(zipPath);
  if (!file) {
    throw new Error(`Missing file in pptx: ${zipPath}`);
  }
  if (limits) {
    assertEntryWithinLimits(file, zipPath, limits);
  }
  const xml = await file.async("string");
  if (limits && Buffer.byteLength(xml) > limits.maxEntryBytes) {
    throw new Error(
      `PPTX entry ${zipPath} size exceeds limit ${limits.maxEntryBytes} bytes.`,
    );
  }
  return parser.parse(xml) as T;
}

export async function readBinary(
  zip: JSZip,
  zipPath: string,
  limits?: Limits,
): Promise<Buffer> {
  const file = zip.file(zipPath);
  if (!file) {
    throw new Error(`Missing file in pptx: ${zipPath}`);
  }
  if (limits) {
    assertEntryWithinLimits(file, zipPath, limits);
  }
  const data = await file.async("nodebuffer");
  if (limits && data.length > limits.maxEntryBytes) {
    throw new Error(
      `PPTX entry ${zipPath} size exceeds limit ${limits.maxEntryBytes} bytes.`,
    );
  }
  return Buffer.from(data);
}

export function listSlidePaths(zip: JSZip): string[] {
  const slides = (zip.file(
    /^ppt\/slides\/slide\d+\.xml$/,
  ) as JSZip.JSZipObject[]).map((file) => file.name);
  return slides.sort((a, b) => extractSlideNumber(a) - extractSlideNumber(b));
}

export function getSlideRelsPath(slidePath: string): string {
  const fileName = path.posix.basename(slidePath);
  return `ppt/slides/_rels/${fileName}.rels`;
}

function enforceZipLimits(zip: JSZip, limits: Limits): void {
  const entries = (Object.values(zip.files) as JSZip.JSZipObject[]).filter(
    (file) => !file.dir,
  );
  if (entries.length > limits.maxZipEntries) {
    throw new Error(
      `PPTX has ${entries.length} entries, exceeds limit ${limits.maxZipEntries}.`,
    );
  }
  let totalSize = 0;
  for (const entry of entries) {
    const size = getEntrySize(entry);
    if (size == null) {
      throw new Error(`Unable to determine size for entry: ${entry.name}`);
    }
    if (size > limits.maxEntryBytes) {
      throw new Error(
        `PPTX entry ${entry.name} size exceeds limit ${limits.maxEntryBytes} bytes.`,
      );
    }
    totalSize += size;
    if (totalSize > limits.maxTotalUncompressedBytes) {
      throw new Error(
        `PPTX total uncompressed size exceeds limit ${limits.maxTotalUncompressedBytes} bytes.`,
      );
    }
  }
}

function assertEntryWithinLimits(
  file: JSZip.JSZipObject,
  zipPath: string,
  limits: Limits,
): void {
  const size = getEntrySize(file);
  if (size == null) {
    throw new Error(`Unable to determine size for entry: ${zipPath}`);
  }
  if (size > limits.maxEntryBytes) {
    throw new Error(
      `PPTX entry ${zipPath} size exceeds limit ${limits.maxEntryBytes} bytes.`,
    );
  }
}

function getEntrySize(file: JSZip.JSZipObject): number | null {
  const data = (file as { _data?: { uncompressedSize?: number; length?: number } })
    ._data;
  const size = data?.uncompressedSize ?? data?.length;
  return typeof size === "number" ? size : null;
}

function extractSlideNumber(slidePath: string): number {
  const match = slidePath.match(/slide(\d+)\.xml$/);
  if (!match) {
    return Number.MAX_SAFE_INTEGER;
  }
  return Number(match[1]);
}
