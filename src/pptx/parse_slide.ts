import path from "node:path";
import fs from "node:fs/promises";
import { ImageElement, SlideElement, TextElement } from "../types";

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
    const extension = path.posix.extname(normalizedTarget) || ".png";
    const imageName = `slide-${options.slideIndex}-img-${imageCount}${extension}`;
    const imagePath = path.join(options.imagesDir, imageName);
    const data = await options.zipReadFile(normalizedTarget);
    await fs.writeFile(imagePath, data);
    const element: ImageElement = {
      id: `i${imageCount}`,
      type: "image",
      x: emuToPx(Number(off?.["@_x"] ?? 0)),
      y: emuToPx(Number(off?.["@_y"] ?? 0)),
      width: emuToPx(Number(ext?.["@_cx"] ?? 0)),
      height: emuToPx(Number(ext?.["@_cy"] ?? 0)),
      rotation,
      src: path.posix.join("assets/images", imageName),
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
