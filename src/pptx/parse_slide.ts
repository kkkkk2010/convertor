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
  svgTargets: Set<string>;
  debugImages: boolean;
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
    let normalizedTarget = normalizeTargetPath(target);
    const originalTarget = normalizedTarget;
    const baseExtension =
      path.posix.extname(normalizedTarget).toLowerCase() || ".png";
    const imageBaseName = `slide-${options.slideIndex}-img-${imageCount}`;
    let extension = baseExtension;
    let detectedType: ImageType | null = null;
    const preferredSvg = findPreferredSvgTarget(
      normalizedTarget,
      options.zipFileExists,
      options.svgTargets,
    );
    if (preferredSvg) {
      normalizedTarget = preferredSvg;
      extension = ".svg";
    }
    let data = await options.zipReadFile(normalizedTarget);
    detectedType = detectImageType(data, normalizedTarget, options.svgTargets);
    extension = resolveExtension(baseExtension, detectedType);
    if (extension === ".png" && data.length < 2000) {
      const svgTarget = findSvgAlternative(
        normalizedTarget,
        options.zipFileExists,
        options.svgTargets,
      );
      if (svgTarget) {
        normalizedTarget = svgTarget;
        data = await options.zipReadFile(normalizedTarget);
        detectedType = detectImageType(data, normalizedTarget, options.svgTargets);
        extension = resolveExtension(baseExtension, detectedType);
      }
    }
    if (extension === ".png" && detectedType !== "png") {
      extension = fallbackNonPngExtension(baseExtension, detectedType);
    }
    const { src } = await saveImageAsset({
      extension,
      data,
      imageBaseName,
      imagesDir: options.imagesDir,
    });
    if (options.debugImages) {
      console.log(
        `[images] target=${originalTarget} detected=${detectedType ?? "unknown"} output=${src}`,
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

function findSvgAlternative(
  pngTarget: string,
  fileExists: (zipPath: string) => boolean,
  svgTargets: Set<string>,
): string | null {
  const ext = path.posix.extname(pngTarget).toLowerCase();
  if (ext !== ".png") {
    return null;
  }
  const base = pngTarget.slice(0, -ext.length);
  const match = base.match(/^(.*)-\d+$/);
  if (match) {
    const candidate = `${match[1]}-3.svg`;
    if (fileExists(candidate) || svgTargets.has(candidate)) {
      return candidate;
    }
  }
  const direct = `${base}.svg`;
  if (fileExists(direct) || svgTargets.has(direct)) {
    return direct;
  }
  return null;
}

function findPreferredSvgTarget(
  target: string,
  fileExists: (zipPath: string) => boolean,
  svgTargets: Set<string>,
): string | null {
  if (svgTargets.has(target)) {
    return target;
  }
  const ext = path.posix.extname(target).toLowerCase();
  const base = ext ? target.slice(0, -ext.length) : target;
  const direct = `${base}.svg`;
  if (fileExists(direct) || svgTargets.has(direct)) {
    return direct;
  }
  if (ext === ".png") {
    return findSvgAlternative(target, fileExists, svgTargets);
  }
  return null;
}

type ImageType = "svg" | "png" | "jpeg";

function detectImageType(
  data: Buffer,
  target: string,
  svgTargets: Set<string>,
): ImageType | null {
  if (svgTargets.has(target)) {
    return "svg";
  }
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
};

async function saveImageAsset(options: SaveImageOptions): Promise<{ src: string }> {
  const { extension, data, imageBaseName, imagesDir } = options;

  if (extension === ".svg") {
    const imageName = `${imageBaseName}${extension}`;
    const imagePath = path.join(imagesDir, imageName);
    await fs.writeFile(imagePath, data);
    return { src: path.posix.join("assets/images", imageName) };
  }

  const imageName = `${imageBaseName}${extension}`;
  const imagePath = path.join(imagesDir, imageName);
  await fs.writeFile(imagePath, data);
  return { src: path.posix.join("assets/images", imageName) };
}
