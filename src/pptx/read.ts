import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

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
    .map((file) => file.name);
  return slides.sort((a, b) => extractSlideNumber(a) - extractSlideNumber(b));
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
