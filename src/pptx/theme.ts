type AnyRecord = Record<string, any>;

export type ThemeColorMap = Record<string, string>;

export type ResolvedColor = {
  color: string;
  usedScheme: boolean;
};

export function parseThemeColors(themeXml: AnyRecord | null): ThemeColorMap {
  const scheme =
    themeXml?.["a:theme"]?.["a:themeElements"]?.["a:clrScheme"] ?? null;
  if (!scheme || typeof scheme !== "object") {
    return {};
  }

  const entries = Object.entries(scheme).filter(([key]) => key !== "@_name");
  const map: ThemeColorMap = {};
  for (const [key, value] of entries) {
    const node = value as AnyRecord;
    const hex = extractColorFromNode(node);
    if (hex) {
      map[key] = hex;
    }
  }
  return map;
}

export function resolveColor(
  rPr: AnyRecord | undefined,
  defRPr: AnyRecord | undefined,
  theme: ThemeColorMap,
): ResolvedColor {
  const rPrFill = rPr?.["a:solidFill"];
  const defFill = defRPr?.["a:solidFill"];
  const resolved =
    resolveSolidFill(rPrFill, theme) ?? resolveSolidFill(defFill, theme);
  return resolved ?? { color: "#111111", usedScheme: false };
}

export function resolveFontSize(
  rPr: AnyRecord | undefined,
  defRPr: AnyRecord | undefined,
): number | null {
  const raw = rPr?.["@_sz"] ?? defRPr?.["@_sz"];
  if (raw === undefined || raw === null) {
    return null;
  }
  const size = Number(raw);
  if (Number.isNaN(size)) {
    return null;
  }
  return size / 100;
}

export function resolveFontFamily(
  rPr: AnyRecord | undefined,
  defRPr: AnyRecord | undefined,
): string | null {
  const typeface =
    rPr?.["a:latin"]?.["@_typeface"] ??
    defRPr?.["a:latin"]?.["@_typeface"];
  if (typeof typeface === "string" && typeface.trim().length > 0) {
    return typeface;
  }
  return null;
}

export function resolveBooleanAttr(
  rPr: AnyRecord | undefined,
  defRPr: AnyRecord | undefined,
  attr: string,
): boolean {
  const value = rPr?.[attr] ?? defRPr?.[attr];
  if (value === undefined || value === null) {
    return false;
  }
  if (typeof value === "string") {
    return value === "1" || value.toLowerCase() === "true";
  }
  return Boolean(value);
}

export function resolveUnderline(
  rPr: AnyRecord | undefined,
  defRPr: AnyRecord | undefined,
): boolean {
  const value = rPr?.["@_u"] ?? defRPr?.["@_u"];
  if (!value) {
    return false;
  }
  if (typeof value === "string") {
    return value.toLowerCase() !== "none";
  }
  return Boolean(value);
}

export function resolveParagraphProps(
  pPr: AnyRecord | undefined,
  fontSizePt: number,
): { align: "left" | "center" | "right" | "justify"; lineHeight: number | null } {
  const alignRaw = pPr?.["@_algn"];
  const align = mapAlign(alignRaw);
  const lineHeight = resolveLineHeight(pPr?.["a:lnSpc"], fontSizePt);
  return { align, lineHeight };
}

function extractColorFromNode(node: AnyRecord): string | null {
  const srgb = node?.["a:srgbClr"];
  const sys = node?.["a:sysClr"];
  const srgbVal = srgb?.["@_val"];
  if (typeof srgbVal === "string") {
    return formatHex(srgbVal);
  }
  const sysVal = sys?.["@_lastClr"] ?? sys?.["@_val"];
  if (typeof sysVal === "string") {
    return formatHex(sysVal);
  }
  return null;
}

function resolveSolidFill(
  solidFill: AnyRecord | undefined,
  theme: ThemeColorMap,
): ResolvedColor | null {
  if (!solidFill || typeof solidFill !== "object") {
    return null;
  }
  const srgb = solidFill["a:srgbClr"];
  if (srgb?.["@_val"]) {
    const base = formatHex(srgb["@_val"]);
    const adjusted = applyLumAdjust(base, srgb);
    return { color: adjusted, usedScheme: false };
  }
  const scheme = solidFill["a:schemeClr"];
  if (scheme?.["@_val"]) {
    const key = scheme["@_val"];
    const base = theme[key] ?? "#111111";
    const adjusted = applyLumAdjust(base, scheme);
    return { color: adjusted, usedScheme: true };
  }
  return null;
}

function applyLumAdjust(hex: string, node: AnyRecord): string {
  const lumMod = Number(node?.["a:lumMod"]?.["@_val"] ?? 100000);
  const lumOff = Number(node?.["a:lumOff"]?.["@_val"] ?? 0);
  const mod = Number.isNaN(lumMod) ? 100000 : lumMod;
  const off = Number.isNaN(lumOff) ? 0 : lumOff;
  const rgb = hexToRgb(hex);
  if (!rgb) {
    return hex;
  }
  const adjust = (channel: number) => {
    const scaled = channel * (mod / 100000);
    const offset = 255 * (off / 100000);
    return clamp(Math.round(scaled + offset));
  };
  const result = {
    r: adjust(rgb.r),
    g: adjust(rgb.g),
    b: adjust(rgb.b),
  };
  return rgbToHex(result);
}

function resolveLineHeight(
  lnSpc: AnyRecord | undefined,
  fontSizePt: number,
): number | null {
  if (!lnSpc || typeof lnSpc !== "object") {
    return null;
  }
  if (lnSpc["a:spcPct"]?.["@_val"]) {
    const val = Number(lnSpc["a:spcPct"]["@_val"]);
    if (!Number.isNaN(val)) {
      return val / 100000;
    }
  }
  if (lnSpc["a:spcPts"]?.["@_val"]) {
    const val = Number(lnSpc["a:spcPts"]["@_val"]);
    if (!Number.isNaN(val) && fontSizePt > 0) {
      return val / 100 / fontSizePt;
    }
  }
  return null;
}

function mapAlign(raw: string | undefined): "left" | "center" | "right" | "justify" {
  switch (raw) {
    case "ctr":
      return "center";
    case "r":
      return "right";
    case "just":
    case "dist":
      return "justify";
    case "l":
    default:
      return "left";
  }
}

function formatHex(hex: string): string {
  const clean = hex.replace("#", "").trim().toUpperCase();
  if (clean.length === 6) {
    return `#${clean}`;
  }
  return `#${clean.padStart(6, "0")}`;
}

function hexToRgb(hex: string): { r: number; g: number; b: number } | null {
  const clean = hex.replace("#", "");
  if (clean.length !== 6) {
    return null;
  }
  const r = Number.parseInt(clean.slice(0, 2), 16);
  const g = Number.parseInt(clean.slice(2, 4), 16);
  const b = Number.parseInt(clean.slice(4, 6), 16);
  if ([r, g, b].some((val) => Number.isNaN(val))) {
    return null;
  }
  return { r, g, b };
}

function rgbToHex(rgb: { r: number; g: number; b: number }): string {
  const toHex = (value: number) =>
    clamp(value).toString(16).padStart(2, "0").toUpperCase();
  return `#${toHex(rgb.r)}${toHex(rgb.g)}${toHex(rgb.b)}`;
}

function clamp(value: number): number {
  return Math.max(0, Math.min(255, value));
}
