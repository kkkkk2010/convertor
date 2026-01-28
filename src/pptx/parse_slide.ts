import path from "node:path";
import fs from "node:fs/promises";
import { ImageElement, SlideElement, TextElement } from "../types";
import {
  ThemeColorMap,
  resolveBooleanAttr,
  resolveColor,
  resolveFontFamily,
  resolveFontSize,
  resolveParagraphProps,
  resolveUnderline,
} from "./theme";

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
  zipFileExists: (zipPath: string) => boolean;
  imagesDir: string;
  theme: ThemeColorMap;
  debugSvg: boolean;
};

type AnyRecord = Record<string, any>;

export async function parseSlide(
  slideXml: SlideXml,
  options: ParseSlideOptions,
): Promise<{
  elements: SlideElement[];
  stats: {
    textElements: number;
    schemeClrElements: number;
    multistyleElements: number;
    imageElements: number;
  };
}> {
  const spTree = slideXml["p:sld"]?.["p:cSld"]?.["p:spTree"] as AnyRecord | undefined;
  const shapes = ensureArray(spTree?.["p:sp"]) as AnyRecord[];
  const pics = ensureArray(spTree?.["p:pic"]) as AnyRecord[];
  const elements: SlideElement[] = [];

  let textCount = 0;
  let schemeClrCount = 0;
  let multistyleCount = 0;
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
    const { text, style, schemeClrUsed, multistyle } = extractTextAndStyle(
      textBody,
      options.theme,
    );
    if (schemeClrUsed) {
      schemeClrCount += 1;
    }
    if (multistyle) {
      multistyleCount += 1;
      console.warn(
        `MULTISTYLE slide ${options.slideIndex} text t${textCount} using base`,
      );
    }
    const element: TextElement = {
      id: `t${textCount}`,
      type: "text",
      x: emuToPx(Number(off?.["@_x"] ?? 0)),
      y: emuToPx(Number(off?.["@_y"] ?? 0)),
      width: emuToPx(Number(ext?.["@_cx"] ?? 0)),
      height: emuToPx(Number(ext?.["@_cy"] ?? 0)),
      rotation,
      text,
      style,
    };
    elements.push(element);
  }

  let imageCount = 0;
  for (const pic of pics) {
    const blip = pic?.["p:blipFill"]?.["a:blip"] as AnyRecord | undefined;
    const rasterRid = blip?.["@_r:embed"] as string | undefined;
    const svgRid = findSvgBlipEmbed(blip);
    const chosenRid = svgRid ?? rasterRid;
    if (!chosenRid) {
      continue;
    }
    imageCount += 1;
    const xfrm = pic?.["p:spPr"]?.["a:xfrm"];
    const off = xfrm?.["a:off"];
    const ext = xfrm?.["a:ext"];
    const rotation = ooxmlRotToDeg(Number(xfrm?.["@_rot"]));
    const relationship = resolveRelationship(options.rels, chosenRid);
    if (!relationship) {
      if (svgRid) {
        throw new Error(
          `Missing SVG relationship for rId ${svgRid} on slide ${options.slideIndex}.`,
        );
      }
      continue;
    }
    let normalizedTarget = normalizeTargetPath(relationship.target);
    const originalTarget = normalizedTarget;
    const baseExtension =
      path.posix.extname(normalizedTarget).toLowerCase() || ".png";
    const imageBaseName = `slide-${options.slideIndex}-img-${imageCount}`;
    let extension = baseExtension;
    let detectedType: ImageType | null = null;
    let data = await options.zipReadFile(normalizedTarget);
    detectedType = detectImageType(data);
    extension = resolveExtension(baseExtension, detectedType);
    if (extension === ".png" && data.length < 2000) {
      detectedType = detectImageType(data);
      extension = resolveExtension(baseExtension, detectedType);
    }
    if (extension === ".png" && detectedType !== "png") {
      extension = fallbackNonPngExtension(baseExtension, detectedType);
    }
    const { src } = await saveImageAsset({
      extension,
      data,
      imageBaseName,
      imagesDir: options.imagesDir,
      debugInfo: {
        slideIndex: options.slideIndex,
        relTarget: originalTarget,
        chosenTarget: normalizedTarget,
      },
    });
    if (options.debugSvg) {
      console.log(
        `SVG_DEBUG slide=${options.slideIndex} rasterRid=${rasterRid ?? "none"} svgRid=${svgRid ?? "none"} chosenRid=${relationship.id} target=${relationship.target} out=${src}`,
      );
    }
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

  return {
    elements,
    stats: {
      textElements: textCount,
      schemeClrElements: schemeClrCount,
      multistyleElements: multistyleCount,
      imageElements: imageCount,
    },
  };
}

function ensureArray<T>(value: T | T[] | undefined): T[] {
  if (!value) {
    return [];
  }
  return Array.isArray(value) ? value : [value];
}

function extractTextAndStyle(
  textBody: Record<string, unknown>,
  theme: ThemeColorMap,
): {
  text: string;
  style: TextElement["style"];
  schemeClrUsed: boolean;
  multistyle: boolean;
} {
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
  const text = chunks.join("\n").trim();

  let baseRun: {
    rPr: AnyRecord | undefined;
    defRPr: AnyRecord | undefined;
    pPr: AnyRecord | undefined;
  } | null = null;
  let firstRun: {
    rPr: AnyRecord | undefined;
    defRPr: AnyRecord | undefined;
    pPr: AnyRecord | undefined;
  } | null = null;
  const runStyles: RunStyle[] = [];

  for (const paragraph of paragraphs) {
    const p = paragraph as Record<string, any> | undefined;
    const pPr = p?.["a:pPr"] as AnyRecord | undefined;
    const defRPr = pPr?.["a:defRPr"] as AnyRecord | undefined;
    const runs = ensureArray(p?.["a:r"]);
    for (const run of runs) {
      const r = run as AnyRecord | undefined;
      const rPr = r?.["a:rPr"] as AnyRecord | undefined;
      const hasRPr = Boolean(rPr && Object.keys(rPr).length > 0);
      if (!firstRun) {
        firstRun = { rPr, defRPr, pPr };
      }
      if (!baseRun && hasRPr) {
        baseRun = { rPr, defRPr, pPr };
      }
      const style = resolveRunStyle(rPr, defRPr, theme);
      runStyles.push(style);
    }
  }

  const resolvedBase = baseRun ?? firstRun;
  const baseStyle = resolveRunStyle(
    resolvedBase?.rPr,
    resolvedBase?.defRPr,
    theme,
  );
  const { align, lineHeight } = resolveParagraphProps(
    resolvedBase?.pPr,
    baseStyle.fontSizePt,
  );
  const style: TextElement["style"] = {
    fontFamily: baseStyle.fontFamily,
    fontSizePt: baseStyle.fontSizePt,
    color: baseStyle.color,
    bold: baseStyle.bold,
    italic: baseStyle.italic,
    underline: baseStyle.underline,
    align,
    lineHeight,
  };

  const multistyle = runStyles.some((runStyle) =>
    isStyleDifferent(runStyle, baseStyle),
  );

  return {
    text,
    style,
    schemeClrUsed: baseStyle.usedScheme,
    multistyle,
  };
}

type RunStyle = {
  fontFamily: string;
  fontSizePt: number;
  color: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  usedScheme: boolean;
};

function resolveRunStyle(
  rPr: AnyRecord | undefined,
  defRPr: AnyRecord | undefined,
  theme: ThemeColorMap,
): RunStyle {
  const fontSizePt = resolveFontSize(rPr, defRPr) ?? 28;
  const fontFamily = resolveFontFamily(rPr, defRPr) ?? "Arial";
  const bold = resolveBooleanAttr(rPr, defRPr, "@_b");
  const italic = resolveBooleanAttr(rPr, defRPr, "@_i");
  const underline = resolveUnderline(rPr, defRPr);
  const { color, usedScheme } = resolveColor(rPr, defRPr, theme);
  return {
    fontFamily,
    fontSizePt,
    color,
    bold,
    italic,
    underline,
    usedScheme,
  };
}

function isStyleDifferent(a: RunStyle, b: RunStyle): boolean {
  return (
    a.fontFamily !== b.fontFamily ||
    a.fontSizePt !== b.fontSizePt ||
    a.color !== b.color ||
    a.bold !== b.bold ||
    a.italic !== b.italic ||
    a.underline !== b.underline
  );
}

function resolveRelationship(
  rels: RelsXml | null,
  id: string,
): { id: string; target: string } | null {
  if (!rels?.Relationships?.Relationship) {
    return null;
  }
  const relationships = ensureArray(rels.Relationships.Relationship);
  const match = relationships.find((rel) => rel["@_Id"] === id);
  if (!match?.["@_Target"]) {
    return null;
  }
  return { id, target: match["@_Target"] };
}

function normalizeTargetPath(target: string): string {
  const normalized = target.replace(/^..\//, "");
  if (normalized.startsWith("ppt/")) {
    return normalized;
  }
  return path.posix.join("ppt", normalized);
}

function findSvgBlipEmbed(blip: AnyRecord | undefined): string | undefined {
  if (!blip || typeof blip !== "object") {
    return undefined;
  }
  const stack: unknown[] = [blip];
  while (stack.length > 0) {
    const current = stack.pop();
    if (!current || typeof current !== "object") {
      continue;
    }
    const record = current as Record<string, unknown>;
    for (const [key, value] of Object.entries(record)) {
      if (key.endsWith("svgBlip") && value && typeof value === "object") {
        const embed = extractSvgRid(value);
        if (embed) {
          return embed;
        }
      }
      if (Array.isArray(value)) {
        stack.push(...value);
      } else if (value && typeof value === "object") {
        stack.push(value);
      }
    }
  }
  return undefined;
}

function extractSvgRid(value: unknown): string | undefined {
  if (!value) {
    return undefined;
  }
  if (Array.isArray(value)) {
    for (const entry of value) {
      const extracted = extractSvgRid(entry);
      if (extracted) {
        return extracted;
      }
    }
    return undefined;
  }
  if (typeof value !== "object") {
    return undefined;
  }
  const record = value as Record<string, unknown>;
  const embed = record["@_r:embed"];
  if (typeof embed === "string" && embed.length > 0) {
    return embed;
  }
  const link = record["@_r:link"];
  if (typeof link === "string" && link.length > 0) {
    return link;
  }
  return undefined;
}

type ImageType = "svg" | "png" | "jpeg";

function detectImageType(data: Buffer): ImageType | null {
  if (isSvgData(data)) {
    return "svg";
  }
  if (isPngSignature(data)) {
    return "png";
  }
  if (isJpegSignature(data)) {
    return "jpeg";
  }
  return null;
}

function resolveExtension(
  baseExtension: string,
  detectedType: ImageType | null,
): string {
  if (detectedType === "svg") {
    return ".svg";
  }
  if (detectedType === "png") {
    return ".png";
  }
  if (detectedType === "jpeg") {
    return ".jpg";
  }
  return baseExtension || ".bin";
}

function fallbackNonPngExtension(
  baseExtension: string,
  detectedType: ImageType | null,
): string {
  if (detectedType === "svg") {
    return ".svg";
  }
  if (detectedType === "jpeg") {
    return ".jpg";
  }
  if (baseExtension && baseExtension !== ".png") {
    return baseExtension;
  }
  return ".bin";
}

function isSvgData(data: Buffer): boolean {
  const snippet = data
    .toString("utf8", 0, 2048)
    .replace(/^\uFEFF/, "")
    .trimStart();
  if (snippet.startsWith("<svg")) {
    return true;
  }
  if (snippet.startsWith("<?xml") && snippet.includes("<svg")) {
    return true;
  }
  return false;
}

function isPngSignature(data: Buffer): boolean {
  if (data.length < 8) {
    return false;
  }
  return (
    data[0] === 0x89 &&
    data[1] === 0x50 &&
    data[2] === 0x4e &&
    data[3] === 0x47 &&
    data[4] === 0x0d &&
    data[5] === 0x0a &&
    data[6] === 0x1a &&
    data[7] === 0x0a
  );
}

function isJpegSignature(data: Buffer): boolean {
  if (data.length < 3) {
    return false;
  }
  return data[0] === 0xff && data[1] === 0xd8 && data[2] === 0xff;
}

type SaveImageOptions = {
  extension: string;
  data: Buffer;
  imageBaseName: string;
  imagesDir: string;
  debugInfo: {
    slideIndex: number;
    relTarget: string;
    chosenTarget: string;
  };
};

async function saveImageAsset(options: SaveImageOptions): Promise<{ src: string }> {
  const { extension, data, imageBaseName, imagesDir, debugInfo } = options;

  if (extension === ".svg") {
    if (!isSvgData(data)) {
      throw new Error(
        `Attempted to write non-SVG buffer to .svg file (slide ${debugInfo.slideIndex}, rel ${debugInfo.relTarget}, chosen ${debugInfo.chosenTarget}).`,
      );
    }
    const imageName = `${imageBaseName}${extension}`;
    const imagePath = path.join(imagesDir, imageName);
    await fs.writeFile(imagePath, data);
    return { src: path.posix.join("assets/images", imageName) };
  }

  if (extension === ".png" && !isPngSignature(data)) {
    throw new Error(
      `Attempted to write non-PNG buffer to .png file (slide ${debugInfo.slideIndex}, rel ${debugInfo.relTarget}, chosen ${debugInfo.chosenTarget}).`,
    );
  }

  const imageName = `${imageBaseName}${extension}`;
  const imagePath = path.join(imagesDir, imageName);
  await fs.writeFile(imagePath, data);
  return { src: path.posix.join("assets/images", imageName) };
}
