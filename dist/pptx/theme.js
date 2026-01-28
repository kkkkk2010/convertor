"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.parseThemeColors = parseThemeColors;
exports.resolveColor = resolveColor;
exports.resolveFontSize = resolveFontSize;
exports.resolveFontFamily = resolveFontFamily;
exports.resolveBooleanAttr = resolveBooleanAttr;
exports.resolveUnderline = resolveUnderline;
exports.resolveParagraphProps = resolveParagraphProps;
function parseThemeColors(themeXml) {
    var _a, _b, _c;
    const scheme = (_c = (_b = (_a = themeXml === null || themeXml === void 0 ? void 0 : themeXml["a:theme"]) === null || _a === void 0 ? void 0 : _a["a:themeElements"]) === null || _b === void 0 ? void 0 : _b["a:clrScheme"]) !== null && _c !== void 0 ? _c : null;
    if (!scheme || typeof scheme !== "object") {
        return {};
    }
    const entries = Object.entries(scheme).filter(([key]) => key !== "@_name");
    const map = {};
    for (const [key, value] of entries) {
        const node = value;
        const hex = extractColorFromNode(node);
        if (hex) {
            map[key] = hex;
        }
    }
    return map;
}
function resolveColor(rPr, defRPr, theme) {
    var _a;
    const rPrFill = rPr === null || rPr === void 0 ? void 0 : rPr["a:solidFill"];
    const defFill = defRPr === null || defRPr === void 0 ? void 0 : defRPr["a:solidFill"];
    const resolved = (_a = resolveSolidFill(rPrFill, theme)) !== null && _a !== void 0 ? _a : resolveSolidFill(defFill, theme);
    return resolved !== null && resolved !== void 0 ? resolved : { color: "#111111", usedScheme: false };
}
function resolveFontSize(rPr, defRPr) {
    var _a;
    const raw = (_a = rPr === null || rPr === void 0 ? void 0 : rPr["@_sz"]) !== null && _a !== void 0 ? _a : defRPr === null || defRPr === void 0 ? void 0 : defRPr["@_sz"];
    if (raw === undefined || raw === null) {
        return null;
    }
    const size = Number(raw);
    if (Number.isNaN(size)) {
        return null;
    }
    return size / 100;
}
function resolveFontFamily(rPr, defRPr) {
    var _a, _b, _c;
    const typeface = (_b = (_a = rPr === null || rPr === void 0 ? void 0 : rPr["a:latin"]) === null || _a === void 0 ? void 0 : _a["@_typeface"]) !== null && _b !== void 0 ? _b : (_c = defRPr === null || defRPr === void 0 ? void 0 : defRPr["a:latin"]) === null || _c === void 0 ? void 0 : _c["@_typeface"];
    if (typeof typeface === "string" && typeface.trim().length > 0) {
        return typeface;
    }
    return null;
}
function resolveBooleanAttr(rPr, defRPr, attr) {
    var _a;
    const value = (_a = rPr === null || rPr === void 0 ? void 0 : rPr[attr]) !== null && _a !== void 0 ? _a : defRPr === null || defRPr === void 0 ? void 0 : defRPr[attr];
    if (value === undefined || value === null) {
        return false;
    }
    if (typeof value === "string") {
        return value === "1" || value.toLowerCase() === "true";
    }
    return Boolean(value);
}
function resolveUnderline(rPr, defRPr) {
    var _a;
    const value = (_a = rPr === null || rPr === void 0 ? void 0 : rPr["@_u"]) !== null && _a !== void 0 ? _a : defRPr === null || defRPr === void 0 ? void 0 : defRPr["@_u"];
    if (!value) {
        return false;
    }
    if (typeof value === "string") {
        return value.toLowerCase() !== "none";
    }
    return Boolean(value);
}
function resolveParagraphProps(pPr, fontSizePt) {
    const alignRaw = pPr === null || pPr === void 0 ? void 0 : pPr["@_algn"];
    const align = mapAlign(alignRaw);
    const lineHeight = resolveLineHeight(pPr === null || pPr === void 0 ? void 0 : pPr["a:lnSpc"], fontSizePt);
    return { align, lineHeight };
}
function extractColorFromNode(node) {
    var _a;
    const srgb = node === null || node === void 0 ? void 0 : node["a:srgbClr"];
    const sys = node === null || node === void 0 ? void 0 : node["a:sysClr"];
    const srgbVal = srgb === null || srgb === void 0 ? void 0 : srgb["@_val"];
    if (typeof srgbVal === "string") {
        return formatHex(srgbVal);
    }
    const sysVal = (_a = sys === null || sys === void 0 ? void 0 : sys["@_lastClr"]) !== null && _a !== void 0 ? _a : sys === null || sys === void 0 ? void 0 : sys["@_val"];
    if (typeof sysVal === "string") {
        return formatHex(sysVal);
    }
    return null;
}
function resolveSolidFill(solidFill, theme) {
    var _a;
    if (!solidFill || typeof solidFill !== "object") {
        return null;
    }
    const srgb = solidFill["a:srgbClr"];
    if (srgb === null || srgb === void 0 ? void 0 : srgb["@_val"]) {
        const base = formatHex(srgb["@_val"]);
        const adjusted = applyLumAdjust(base, srgb);
        return { color: adjusted, usedScheme: false };
    }
    const scheme = solidFill["a:schemeClr"];
    if (scheme === null || scheme === void 0 ? void 0 : scheme["@_val"]) {
        const key = scheme["@_val"];
        const base = (_a = theme[key]) !== null && _a !== void 0 ? _a : "#111111";
        const adjusted = applyLumAdjust(base, scheme);
        return { color: adjusted, usedScheme: true };
    }
    return null;
}
function applyLumAdjust(hex, node) {
    var _a, _b, _c, _d;
    const lumMod = Number((_b = (_a = node === null || node === void 0 ? void 0 : node["a:lumMod"]) === null || _a === void 0 ? void 0 : _a["@_val"]) !== null && _b !== void 0 ? _b : 100000);
    const lumOff = Number((_d = (_c = node === null || node === void 0 ? void 0 : node["a:lumOff"]) === null || _c === void 0 ? void 0 : _c["@_val"]) !== null && _d !== void 0 ? _d : 0);
    const mod = Number.isNaN(lumMod) ? 100000 : lumMod;
    const off = Number.isNaN(lumOff) ? 0 : lumOff;
    const rgb = hexToRgb(hex);
    if (!rgb) {
        return hex;
    }
    const adjust = (channel) => {
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
function resolveLineHeight(lnSpc, fontSizePt) {
    var _a, _b;
    if (!lnSpc || typeof lnSpc !== "object") {
        return null;
    }
    if ((_a = lnSpc["a:spcPct"]) === null || _a === void 0 ? void 0 : _a["@_val"]) {
        const val = Number(lnSpc["a:spcPct"]["@_val"]);
        if (!Number.isNaN(val)) {
            return val / 100000;
        }
    }
    if ((_b = lnSpc["a:spcPts"]) === null || _b === void 0 ? void 0 : _b["@_val"]) {
        const val = Number(lnSpc["a:spcPts"]["@_val"]);
        if (!Number.isNaN(val) && fontSizePt > 0) {
            return val / 100 / fontSizePt;
        }
    }
    return null;
}
function mapAlign(raw) {
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
function formatHex(hex) {
    const clean = hex.replace("#", "").trim().toUpperCase();
    if (clean.length === 6) {
        return `#${clean}`;
    }
    return `#${clean.padStart(6, "0")}`;
}
function hexToRgb(hex) {
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
function rgbToHex(rgb) {
    const toHex = (value) => clamp(value).toString(16).padStart(2, "0").toUpperCase();
    return `#${toHex(rgb.r)}${toHex(rgb.g)}${toHex(rgb.b)}`;
}
function clamp(value) {
    return Math.max(0, Math.min(255, value));
}
