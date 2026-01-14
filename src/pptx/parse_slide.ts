import path from "node:path";
import fs from "node:fs/promises";
import os from "node:os";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { ImageElement, SlideElement, TextElement } from "../types";
import {
  getLibreOfficeCandidates,
  resolveLibreOfficeBinary,
} from "../render/backgrounds";

const execFileAsync = promisify(execFile);

const EMU_PER_INCH = 914400;

export function emuToPx(emu: number): number {
  return (emu / EMU_PER_INCH) * 96;
}

export function ooxmlRotToDeg(rot?: number): number {
  if (!rot) {
    return 0;
  }
  return rot / 60000;
}

type SlideXml = {
  "p:sld"?: {
    "p:cSld"?: {
      "p:spTree"?: {
        "p:sp"?: unknown;
        "p:pic"?: unknown;
      };
    };
  };
};

type RelsXml = {
  Relationships?: {
    Relationship?: Array<{
      "@_Id": string;
      "@_Target": string;
    }> | {
      "@_Id": string;
      "@_Target": string;
    };
  };
};

type ParseSlideOptions = {
  slideIndex: number;
  rels: RelsXml | null;
  zipReadFile: (zipPath: string) => Promise<Buffer>;
  imagesDir: string;
  originalsDir: string;
};

type AnyRecord = Record<string, any>;

export async function parseSlide(
  slideXml: SlideXml,
  options: ParseSlideOptions,
): Promise<SlideElement[]> {
  const spTree = slideXml["p:sld"]?.["p:cSld"]?.["p:spTree"] as AnyRecord | undefined;
  const shapes = ensureArray(spTree?.["p:sp"]) as AnyRecord[];
  const pics = ensureArray(spTree?.["p:pic"]) as AnyRecord[];
  const elements: SlideElement[] = [];

  let textCount = 0;
  for (const shape of shapes) {
    const textBody = shape?.["p:txBody"];
    if (!textBody) {
      continue;
    }
    textCount += 1;
    const xfrm = shape?.["p:spPr"]?.["a:xfrm"];
    const off = xfrm?.["a:off"];
    const ext = xfrm?.["a:ext"];
    const rotation = ooxmlRotToDeg(Number(xfrm?.["@_rot"]));
    const text = extractText(textBody);
    const element: TextElement = {
      id: `t${textCount}`,
      type: "text",
      x: emuToPx(Number(off?.["@_x"] ?? 0)),
      y: emuToPx(Number(off?.["@_y"] ?? 0)),
      width: emuToPx(Number(ext?.["@_cx"] ?? 0)),
      height: emuToPx(Number(ext?.["@_cy"] ?? 0)),
      rotation,
      text,
      style: {
        fontFamily: "Arial",
        fontSize: 28,
        color: "#111111",
        bold: false,
        italic: false,
        underline: false,
        align: "left",
      },
    };
    elements.push(element);
  }

  let imageCount = 0;
  for (const pic of pics) {
    const blip = pic?.["p:blipFill"]?.["a:blip"];
    const embed = blip?.["@_r:embed"];
    if (!embed) {
      continue;
    }
    imageCount += 1;
    const xfrm = pic?.["p:spPr"]?.["a:xfrm"];
    const off = xfrm?.["a:off"];
    const ext = xfrm?.["a:ext"];
    const rotation = ooxmlRotToDeg(Number(xfrm?.["@_rot"]));
    const target = resolveRelationship(options.rels, embed);
    if (!target) {
      continue;
    }
    const normalizedTarget = normalizeTargetPath(target);
    const extension =
      path.posix.extname(normalizedTarget).toLowerCase() || ".png";
    const imageBaseName = `slide-${options.slideIndex}-img-${imageCount}`;
    const data = await options.zipReadFile(normalizedTarget);
    const { src } = await saveImageAsset({
      extension,
      data,
      imageBaseName,
      imagesDir: options.imagesDir,
      originalsDir: options.originalsDir,
    });
    const element: ImageElement = {
      id: `i${imageCount}`,
      type: "image",
      x: emuToPx(Number(off?.["@_x"] ?? 0)),
      y: emuToPx(Number(off?.["@_y"] ?? 0)),
      width: emuToPx(Number(ext?.["@_cx"] ?? 0)),
      height: emuToPx(Number(ext?.["@_cy"] ?? 0)),
      rotation,
      src,
      objectFit: "cover",
    };
    elements.push(element);
  }

  return elements;
}

function ensureArray<T>(value: T | T[] | undefined): T[] {
  if (!value) {
    return [];
  }
  return Array.isArray(value) ? value : [value];
}

function extractText(textBody: Record<string, unknown>): string {
  const body = textBody as Record<string, any>;
  const paragraphs = ensureArray(body["a:p"]);
  const chunks = paragraphs.map((paragraph) => {
    const p = paragraph as Record<string, any> | undefined;
    const runs = ensureArray(p?.["a:r"]);
    const textRuns = runs
      .map((run) => (run as Record<string, any> | undefined)?.["a:t"])
      .filter((value) => typeof value === "string") as string[];
    return textRuns.join("");
  });
  return chunks.join("\n").trim();
}

function resolveRelationship(rels: RelsXml | null, id: string): string | null {
  if (!rels?.Relationships?.Relationship) {
    return null;
  }
  const relationships = ensureArray(rels.Relationships.Relationship);
  const match = relationships.find((rel) => rel["@_Id"] === id);
  return match?.["@_Target"] ?? null;
}

function normalizeTargetPath(target: string): string {
  const normalized = target.replace(/^..\//, "");
  if (normalized.startsWith("ppt/")) {
    return normalized;
  }
  return path.posix.join("ppt", normalized);
}

type SaveImageOptions = {
  extension: string;
  data: Buffer;
  imageBaseName: string;
  imagesDir: string;
  originalsDir: string;
};

async function saveImageAsset(options: SaveImageOptions): Promise<{ src: string }> {
  const { extension, data, imageBaseName, imagesDir, originalsDir } = options;

  if (extension === ".svg") {
    const imageName = `${imageBaseName}${extension}`;
    const imagePath = path.join(imagesDir, imageName);
    await fs.writeFile(imagePath, data);
    return { src: path.posix.join("assets/images", imageName) };
  }

  if (extension === ".emf" || extension === ".wmf") {
    const originalName = `${imageBaseName}${extension}`;
    const originalPath = path.join(originalsDir, originalName);
    await fs.writeFile(originalPath, data);
    const rasterName = `${imageBaseName}.png`;
    await rasterizeVector(originalPath, imagesDir);
    return { src: path.posix.join("assets/images", rasterName) };
  }

  const imageName = `${imageBaseName}${extension}`;
  const imagePath = path.join(imagesDir, imageName);
  await fs.writeFile(imagePath, data);
  return { src: path.posix.join("assets/images", imageName) };
}

async function rasterizeVector(inputPath: string, outDir: string): Promise<void> {
  const libreOfficeBinary = resolveLibreOfficeBinary();
  if (!libreOfficeBinary) {
    const candidates = getLibreOfficeCandidates();
    throw new Error(
      `Missing dependency: LibreOffice.\n- Install LibreOffice and ensure it is on PATH.\n- Tried: ${candidates.join(", ")}`,
    );
  }

  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "pptx-vector-"));
  await execFileAsync(libreOfficeBinary, [
    "--headless",
    "--convert-to",
    "png",
    "--outdir",
    tmpDir,
    inputPath,
  ]);

  const baseName = path.basename(inputPath, path.extname(inputPath));
  const renderedPath = path.join(tmpDir, `${baseName}.png`);
  await fs.rename(renderedPath, path.join(outDir, `${baseName}.png`));
}
