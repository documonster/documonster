/**
 * Shape property (spPr) and text property (txPr) utilities.
 *
 * During chart parsing, spPr and txPr elements are captured as raw XML strings
 * for perfect round-trip fidelity. These utilities provide structured read/write
 * access to the most commonly used properties:
 *
 * - Fill color (solid sRGB / theme)
 * - Line/outline color and width
 * - Font size, bold, italic, color, family
 *
 * The approach: parse raw XML on demand using regex extraction (no full DOM),
 * and generate raw XML from structured models when creating new properties.
 */

import type {
  ShapeProperties,
  ChartTextProperties,
  ChartColor,
  ChartFill,
  ChartLine,
  EffectList,
  Shadow,
  Scene3D,
  ShapeProperties3D,
  Bevel
} from "./types";

// ============================================================================
// Raw XML access helpers
// ============================================================================

/** Get the raw XML string if the object was captured as raw XML */
function getRawXml(obj: any): string | undefined {
  return obj?._rawXml;
}

/** Check if the object is a raw XML capture (not yet structured) */
function isRawXml(obj: any): boolean {
  return typeof obj?._rawXml === "string";
}

// ============================================================================
// Color parsing from raw XML
// ============================================================================

// Pre-compiled regexes for parseColorFromXml (always called with "a" prefix)
const SRGB_CLR_RE = /<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/i;
const SCHEME_CLR_RE = /<a:schemeClr\s+val="([^"]+)"/i;

// Pre-compiled regexes for parseShadowElement
const OUTER_SHDW_RE = /<a:outerShdw\s+([^>]*)>([\s\S]*?)<\/a:outerShdw>/;
const INNER_SHDW_RE = /<a:innerShdw\s+([^>]*)>([\s\S]*?)<\/a:innerShdw>/;
const SHADOW_RES: Record<string, RegExp> = {
  "a:outerShdw": OUTER_SHDW_RE,
  "a:innerShdw": INNER_SHDW_RE
};

/**
 * Extract a region from `xml` between the opening tag match and the corresponding
 * closing tag (or a reasonable fallback for self-closing elements).
 */
function extractColorRegion(xml: string, openMatch: RegExpExecArray, closeTag: string): string {
  const endIdx = xml.indexOf(closeTag, openMatch.index);
  return xml.slice(
    openMatch.index,
    endIdx >= 0 ? endIdx + closeTag.length + 1 : openMatch.index + openMatch[0].length + 200 // Generous fallback for self-closing tags or short fragments
  );
}

/**
 * Parse all DrawingML color modifiers from a region of XML.
 * Handles: alpha, tint, shade, satMod, lumMod, lumOff.
 */
function parseColorModifiers(region: string, color: ChartColor): void {
  const alphaMatch = /<a:alpha\s+val="(\d+)"/.exec(region);
  if (alphaMatch) {
    color.alpha = parseInt(alphaMatch[1], 10);
  }
  const tintMatch = /<a:tint\s+val="(\d+)"/.exec(region);
  if (tintMatch) {
    color.tint = parseInt(tintMatch[1], 10) / 100000;
  }
  const shadeMatch = /<a:shade\s+val="(\d+)"/.exec(region);
  if (shadeMatch) {
    color.shade = parseInt(shadeMatch[1], 10);
  }
  const satModMatch = /<a:satMod\s+val="(\d+)"/.exec(region);
  if (satModMatch) {
    color.satMod = parseInt(satModMatch[1], 10);
  }
  const lumModMatch = /<a:lumMod\s+val="(\d+)"/.exec(region);
  if (lumModMatch) {
    color.lumMod = parseInt(lumModMatch[1], 10);
  }
  const lumOffMatch = /<a:lumOff\s+val="(\d+)"/.exec(region);
  if (lumOffMatch) {
    color.lumOff = parseInt(lumOffMatch[1], 10);
  }
}

// Pre-compiled regexes for sysClr / prstClr
const SYS_CLR_RE = /<a:sysClr\s+val="([^"]+)"/i;
const PRST_CLR_RE = /<a:prstClr\s+val="([^"]+)"/i;

function parseColorFromXml(xml: string): ChartColor | undefined {
  // Try srgbClr
  const srgbMatch = SRGB_CLR_RE.exec(xml);
  if (srgbMatch) {
    const color: ChartColor = { srgb: srgbMatch[1].toUpperCase() };
    const region = extractColorRegion(xml, srgbMatch, `</a:srgbClr`);
    parseColorModifiers(region, color);
    return color;
  }
  // Try schemeClr (theme)
  const schemeMatch = SCHEME_CLR_RE.exec(xml);
  if (schemeMatch) {
    const themeMap: Record<string, number> = {
      dk1: 0,
      lt1: 1,
      dk2: 2,
      lt2: 3,
      accent1: 4,
      accent2: 5,
      accent3: 6,
      accent4: 7,
      accent5: 8,
      accent6: 9,
      hlink: 10,
      folHlink: 11
    };
    const color: ChartColor = { theme: themeMap[schemeMatch[1]] ?? 0 };
    const region = extractColorRegion(xml, schemeMatch, `</a:schemeClr`);
    parseColorModifiers(region, color);
    return color;
  }
  // Try sysClr (system color)
  const sysMatch = SYS_CLR_RE.exec(xml);
  if (sysMatch) {
    const color: ChartColor = { sysClr: sysMatch[1] };
    const region = extractColorRegion(xml, sysMatch, `</a:sysClr`);
    parseColorModifiers(region, color);
    return color;
  }
  // Try prstClr (preset color)
  const prstMatch = PRST_CLR_RE.exec(xml);
  if (prstMatch) {
    const color: ChartColor = { prstClr: prstMatch[1] };
    const region = extractColorRegion(xml, prstMatch, `</a:prstClr`);
    parseColorModifiers(region, color);
    return color;
  }
  return undefined;
}

// ============================================================================
// Gradient / pattern fill parsing from raw XML
// ============================================================================

function parseGradientFill(xml: string): ChartFill | undefined {
  const gradStart = xml.indexOf("<a:gradFill");
  const gradEnd = xml.indexOf("</a:gradFill");
  if (gradStart < 0 || gradEnd < 0) {
    return undefined;
  }
  const region = xml.slice(gradStart, gradEnd + 20);

  // Parse gradient stops
  const stops: Array<{ position: number; color: ChartColor }> = [];
  // Find each <a:gs pos="..."> ... </a:gs>
  const gsListStart = region.indexOf("<a:gsLst");
  const gsListEnd = region.indexOf("</a:gsLst");
  if (gsListStart >= 0 && gsListEnd >= 0) {
    const gsListRegion = region.slice(gsListStart, gsListEnd + 12);
    const gsPosRegex = /<a:gs\s+pos="(\d+)"/g;
    let match: RegExpExecArray | null;
    const positions: Array<{ pos: number; startIdx: number }> = [];
    while ((match = gsPosRegex.exec(gsListRegion)) !== null) {
      positions.push({ pos: parseInt(match[1], 10), startIdx: match.index });
    }
    for (let i = 0; i < positions.length; i++) {
      const start = positions[i].startIdx;
      const end = i + 1 < positions.length ? positions[i + 1].startIdx : gsListRegion.length;
      const gsRegion = gsListRegion.slice(start, end);
      const color = parseColorFromXml(gsRegion);
      if (color) {
        stops.push({ position: positions[i].pos / 1000, color });
      }
    }
  }

  if (stops.length === 0) {
    return undefined;
  }

  // Parse angle from <a:lin ang="...">
  let angle: number | undefined;
  let type: "linear" | "circle" | "rect" | "shape" | undefined;
  const linMatch = /<a:lin\s+ang="(\d+)"/.exec(region);
  if (linMatch) {
    angle = parseInt(linMatch[1], 10) / 60000;
    type = "linear";
  }
  // Check for path gradient
  const pathMatch = /<a:path\s+path="([^"]+)"/.exec(region);
  if (pathMatch) {
    type = pathMatch[1] as typeof type;
  }

  return { gradient: { stops, angle, type } };
}

function parsePatternFill(xml: string): ChartFill | undefined {
  const pattMatch = /<a:pattFill\s+prst="([^"]+)"/.exec(xml);
  if (!pattMatch) {
    return undefined;
  }
  const preset = pattMatch[1];
  const pattStart = pattMatch.index;
  const pattEnd = xml.indexOf("</a:pattFill", pattStart);
  const region = xml.slice(pattStart, pattEnd > 0 ? pattEnd + 15 : undefined);

  let foreground: ChartColor | undefined;
  let background: ChartColor | undefined;

  const fgStart = region.indexOf("<a:fgClr");
  if (fgStart >= 0) {
    const fgEnd = region.indexOf("</a:fgClr", fgStart);
    const fgRegion = region.slice(fgStart, fgEnd > 0 ? fgEnd + 12 : undefined);
    foreground = parseColorFromXml(fgRegion);
  }

  const bgStart = region.indexOf("<a:bgClr");
  if (bgStart >= 0) {
    const bgEnd = region.indexOf("</a:bgClr", bgStart);
    const bgRegion = region.slice(bgStart, bgEnd > 0 ? bgEnd + 12 : undefined);
    background = parseColorFromXml(bgRegion);
  }

  return { pattern: { preset, foreground, background } };
}

// ============================================================================
// spPr: Read structured properties from raw XML
// ============================================================================

/**
 * Extract structured fill and line properties from a raw spPr XML string.
 * Returns the structured ShapeProperties if extraction succeeds.
 */
export function parseSpPr(spPr: ShapeProperties): ShapeProperties {
  const rawXml = getRawXml(spPr);
  if (!rawXml) {
    return spPr; // already structured
  }

  const result: ShapeProperties = {};

  // Parse fill
  if (rawXml.includes("<a:solidFill")) {
    const color = parseColorFromXml(
      rawXml.slice(rawXml.indexOf("<a:solidFill"), rawXml.indexOf("</a:solidFill") + 20)
    );
    if (color) {
      result.fill = { solid: color };
    }
  } else if (rawXml.includes("<a:noFill")) {
    result.fill = { noFill: true };
  } else if (rawXml.includes("<a:gradFill")) {
    result.fill = parseGradientFill(rawXml);
  } else if (rawXml.includes("<a:pattFill")) {
    result.fill = parsePatternFill(rawXml);
  }

  // Parse line
  const lnMatch = /<a:ln(?:\s+w="(\d+)")?/.exec(rawXml);
  if (lnMatch) {
    const lnEnd = rawXml.indexOf("</a:ln", lnMatch.index);
    const lnRegion = rawXml.slice(lnMatch.index, lnEnd > 0 ? lnEnd + 10 : undefined);

    const line: ChartLine = {};
    if (lnMatch[1]) {
      line.width = parseInt(lnMatch[1], 10);
    }
    if (lnRegion.includes("<a:noFill")) {
      line.noFill = true;
    } else if (lnRegion.includes("<a:solidFill")) {
      line.color = parseColorFromXml(lnRegion);
    }
    const dashMatch = /<a:prstDash\s+val="([^"]+)"/.exec(lnRegion);
    if (dashMatch) {
      line.dash = dashMatch[1] as ChartLine["dash"];
    }
    result.line = line;
  }

  // Parse effect list
  if (rawXml.includes("<a:effectLst") || rawXml.includes("<a:effectDag")) {
    const effects = parseEffectList(rawXml);
    if (effects) {
      result.effectList = effects;
    }
  }

  // Parse 3D scene and shape properties
  if (rawXml.includes("<a:scene3d")) {
    const scene = parseScene3D(rawXml);
    if (scene) {
      result.scene3d = scene;
    }
  }
  if (rawXml.includes("<a:sp3d")) {
    const sp3d = parseSp3D(rawXml);
    if (sp3d) {
      result.sp3d = sp3d;
    }
  }

  return result;
}

// ============================================================================
// Effect list parsing
// ============================================================================

function parseEffectList(xml: string): EffectList | undefined {
  const effStart = xml.indexOf("<a:effectLst");
  const effEnd = xml.indexOf("</a:effectLst");
  if (effStart < 0 || effEnd < 0) {
    return undefined;
  }
  const region = xml.slice(effStart, effEnd + 14);
  const result: EffectList = {};

  // Blur
  const blurMatch = /<a:blur(\s+[^/>]*)?\s*\/>/.exec(region);
  if (blurMatch) {
    const attrs = parseAttrs(blurMatch[0]);
    result.blur = {};
    if (attrs.rad) {
      result.blur.radius = parseInt(attrs.rad, 10);
    }
    if (attrs.grow === "1") {
      result.blur.grow = true;
    }
  }

  // Outer shadow
  const outerShadow = parseShadowElement(region, "a:outerShdw");
  if (outerShadow) {
    result.outerShadow = outerShadow;
  }

  // Inner shadow
  const innerShadow = parseShadowElement(region, "a:innerShdw");
  if (innerShadow) {
    result.innerShadow = innerShadow;
  }

  // Preset shadow
  const prstMatch = /<a:prstShdw\s+([^>]*)>([\s\S]*?)<\/a:prstShdw>/.exec(region);
  if (prstMatch) {
    const attrs = parseAttrs(prstMatch[0]);
    const color = parseColorFromXml(prstMatch[2]);
    if (attrs.prst && color) {
      result.presetShadow = { preset: attrs.prst, color };
      if (attrs.dist) {
        result.presetShadow.distance = parseInt(attrs.dist, 10);
      }
      if (attrs.dir) {
        result.presetShadow.direction = parseInt(attrs.dir, 10);
      }
    }
  }

  // Glow
  const glowMatch = /<a:glow\s+([^>]*)>([\s\S]*?)<\/a:glow>/.exec(region);
  if (glowMatch) {
    const attrs = parseAttrs(glowMatch[0]);
    const color = parseColorFromXml(glowMatch[2]);
    if (attrs.rad && color) {
      result.glow = { radius: parseInt(attrs.rad, 10), color };
    }
  }

  // Soft edge
  const softEdgeMatch = /<a:softEdge\s+([^/>]*)\s*\/>/.exec(region);
  if (softEdgeMatch) {
    const attrs = parseAttrs(softEdgeMatch[0]);
    if (attrs.rad) {
      result.softEdge = { radius: parseInt(attrs.rad, 10) };
    }
  }

  // Reflection
  const reflMatch = /<a:reflection\s+([^/>]*)\s*\/>/.exec(region);
  if (reflMatch) {
    const attrs = parseAttrs(reflMatch[0]);
    const r: NonNullable<EffectList["reflection"]> = {};
    if (attrs.blurRad) {
      r.blurRadius = parseInt(attrs.blurRad, 10);
    }
    if (attrs.stA) {
      r.startOpacity = parseInt(attrs.stA, 10);
    }
    if (attrs.stPos) {
      r.startPosition = parseInt(attrs.stPos, 10);
    }
    if (attrs.endA) {
      r.endOpacity = parseInt(attrs.endA, 10);
    }
    if (attrs.endPos) {
      r.endPosition = parseInt(attrs.endPos, 10);
    }
    if (attrs.dist) {
      r.distance = parseInt(attrs.dist, 10);
    }
    if (attrs.dir) {
      r.direction = parseInt(attrs.dir, 10);
    }
    if (attrs.fadeDir) {
      r.fadeDirection = parseInt(attrs.fadeDir, 10);
    }
    if (attrs.sx) {
      r.scaleHorizontal = parseInt(attrs.sx, 10);
    }
    if (attrs.sy) {
      r.scaleVertical = parseInt(attrs.sy, 10);
    }
    if (attrs.kx) {
      r.skewHorizontal = parseInt(attrs.kx, 10);
    }
    if (attrs.ky) {
      r.skewVertical = parseInt(attrs.ky, 10);
    }
    if (attrs.algn) {
      r.alignment = attrs.algn as NonNullable<typeof r.alignment>;
    }
    if (attrs.rotWithShape === "1") {
      r.rotateWithShape = true;
    }
    result.reflection = r;
  }

  if (Object.keys(result).length === 0) {
    return undefined;
  }
  return result;
}

function parseShadowElement(xml: string, tag: string): Shadow | undefined {
  const re = SHADOW_RES[tag] ?? new RegExp(`<${tag}\\s+([^>]*)>([\\s\\S]*?)<\\/${tag}>`);
  const m = re.exec(xml);
  if (!m) {
    return undefined;
  }
  const attrs = parseAttrs(m[0]);
  const color = parseColorFromXml(m[2]);
  if (!color) {
    return undefined;
  }
  const shadow: Shadow = { color };
  if (attrs.blurRad) {
    shadow.blurRadius = parseInt(attrs.blurRad, 10);
  }
  if (attrs.dist) {
    shadow.distance = parseInt(attrs.dist, 10);
  }
  if (attrs.dir) {
    shadow.direction = parseInt(attrs.dir, 10);
  }
  if (attrs.algn) {
    shadow.alignment = attrs.algn as NonNullable<typeof shadow.alignment>;
  }
  if (attrs.rotWithShape === "1") {
    shadow.rotateWithShape = true;
  }
  if (attrs.sx) {
    shadow.scaleHorizontal = parseInt(attrs.sx, 10);
  }
  if (attrs.sy) {
    shadow.scaleVertical = parseInt(attrs.sy, 10);
  }
  if (attrs.kx) {
    shadow.skewHorizontal = parseInt(attrs.kx, 10);
  }
  if (attrs.ky) {
    shadow.skewVertical = parseInt(attrs.ky, 10);
  }
  return shadow;
}

function parseScene3D(xml: string): Scene3D | undefined {
  const sceneStart = xml.indexOf("<a:scene3d");
  const sceneEnd = xml.indexOf("</a:scene3d");
  if (sceneStart < 0 || sceneEnd < 0) {
    return undefined;
  }
  const region = xml.slice(sceneStart, sceneEnd + 13);
  const result: Scene3D = {};

  const cameraMatch = /<a:camera\s+([^>]*)(?:\/|>[\s\S]*?<\/a:camera)>/.exec(region);
  if (cameraMatch) {
    const attrs = parseAttrs(cameraMatch[0]);
    if (attrs.prst) {
      result.camera = { preset: attrs.prst };
      if (attrs.fov) {
        result.camera.fov = parseInt(attrs.fov, 10);
      }
      if (attrs.zoom) {
        result.camera.zoom = parseInt(attrs.zoom, 10);
      }
      const rotMatch = /<a:rot\s+([^/>]*)\s*\/>/.exec(cameraMatch[0]);
      if (rotMatch) {
        const ra = parseAttrs(rotMatch[0]);
        if (ra.lat && ra.lon && ra.rev) {
          result.camera.rotation = {
            lat: parseInt(ra.lat, 10),
            lon: parseInt(ra.lon, 10),
            rev: parseInt(ra.rev, 10)
          };
        }
      }
    }
  }
  const lightRigMatch = /<a:lightRig\s+([^>]*)(?:\/|>[\s\S]*?<\/a:lightRig)>/.exec(region);
  if (lightRigMatch) {
    const attrs = parseAttrs(lightRigMatch[0]);
    if (attrs.rig && attrs.dir) {
      result.lightRig = { rig: attrs.rig, direction: attrs.dir };
    }
  }
  if (!result.camera && !result.lightRig) {
    return undefined;
  }
  return result;
}

function parseSp3D(xml: string): ShapeProperties3D | undefined {
  const spStart = xml.indexOf("<a:sp3d");
  if (spStart < 0) {
    return undefined;
  }
  // sp3d can be self-closing or have children (bevels)
  const selfClose = /<a:sp3d\s+[^>]*\/>/.exec(xml);
  const openClose = /<a:sp3d\s+([^>]*)>([\s\S]*?)<\/a:sp3d>/.exec(xml);
  const region = openClose ? openClose[0] : selfClose ? selfClose[0] : xml.slice(spStart);
  const attrs = parseAttrs(region);
  const result: ShapeProperties3D = {};
  if (attrs.z) {
    result.z = parseInt(attrs.z, 10);
  }
  if (attrs.extrusionH) {
    result.extrusionHeight = parseInt(attrs.extrusionH, 10);
  }
  if (attrs.contourW) {
    result.contourWidth = parseInt(attrs.contourW, 10);
  }
  if (attrs.prstMaterial) {
    result.material = attrs.prstMaterial;
  }
  const bevelT = /<a:bevelT\s+([^/>]*)\s*\/>/.exec(region);
  if (bevelT) {
    result.bevelTop = parseBevelAttrs(bevelT[0]);
  }
  const bevelB = /<a:bevelB\s+([^/>]*)\s*\/>/.exec(region);
  if (bevelB) {
    result.bevelBottom = parseBevelAttrs(bevelB[0]);
  }
  if (openClose) {
    const extClrMatch = /<a:extrusionClr>([\s\S]*?)<\/a:extrusionClr>/.exec(openClose[2]);
    if (extClrMatch) {
      const c = parseColorFromXml(extClrMatch[1]);
      if (c) {
        result.extrusionColor = c;
      }
    }
    const contClrMatch = /<a:contourClr>([\s\S]*?)<\/a:contourClr>/.exec(openClose[2]);
    if (contClrMatch) {
      const c = parseColorFromXml(contClrMatch[1]);
      if (c) {
        result.contourColor = c;
      }
    }
  }
  return result;
}

function parseBevelAttrs(tag: string): Bevel {
  const attrs = parseAttrs(tag);
  const bevel: Bevel = {};
  if (attrs.w) {
    bevel.width = parseInt(attrs.w, 10);
  }
  if (attrs.h) {
    bevel.height = parseInt(attrs.h, 10);
  }
  if (attrs.prst) {
    bevel.preset = attrs.prst;
  }
  return bevel;
}

function parseAttrs(tag: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  const re = /(\w+)="([^"]*)"/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(tag)) !== null) {
    attrs[m[1]] = m[2];
  }
  return attrs;
}

/**
 * Get the solid fill color from a ShapeProperties object.
 * Works for both raw XML and structured models.
 */
export function getSpPrFillColor(spPr: ShapeProperties): ChartColor | undefined {
  const parsed = isRawXml(spPr) ? parseSpPr(spPr) : spPr;
  return parsed.fill?.solid;
}

/**
 * Get the line/outline properties from a ShapeProperties object.
 */
export function getSpPrLine(spPr: ShapeProperties): ChartLine | undefined {
  const parsed = isRawXml(spPr) ? parseSpPr(spPr) : spPr;
  return parsed.line;
}

/**
 * Get the gradient fill from a ShapeProperties object.
 */
export function getSpPrGradient(spPr: ShapeProperties): ChartFill["gradient"] | undefined {
  const parsed = isRawXml(spPr) ? parseSpPr(spPr) : spPr;
  return parsed.fill?.gradient;
}

/**
 * Get the pattern fill from a ShapeProperties object.
 */
export function getSpPrPattern(spPr: ShapeProperties): ChartFill["pattern"] | undefined {
  const parsed = isRawXml(spPr) ? parseSpPr(spPr) : spPr;
  return parsed.fill?.pattern;
}

// ============================================================================
// txPr: Read structured properties from raw XML
// ============================================================================

/**
 * Extract structured text properties from a raw txPr XML string.
 */
export function parseTxPr(txPr: ChartTextProperties): ChartTextProperties {
  const rawXml = getRawXml(txPr);
  if (!rawXml) {
    return txPr; // already structured
  }

  const result: ChartTextProperties = {};

  // Font size (a:sz)
  const szMatch = /<a:(?:defRPr|rPr)[^>]*\s+sz="(\d+)"/.exec(rawXml);
  if (szMatch) {
    result.size = parseInt(szMatch[1], 10);
  }

  // Bold
  const bMatch = /<a:(?:defRPr|rPr)[^>]*\s+b="1"/.exec(rawXml);
  if (bMatch) {
    result.bold = true;
  }

  // Italic
  const iMatch = /<a:(?:defRPr|rPr)[^>]*\s+i="1"/.exec(rawXml);
  if (iMatch) {
    result.italic = true;
  }

  // Font color
  const rPrStart = rawXml.indexOf("<a:defRPr");
  if (rPrStart >= 0) {
    const rPrEnd = rawXml.indexOf("</a:defRPr", rPrStart);
    const rPrRegion = rawXml.slice(rPrStart, rPrEnd > 0 ? rPrEnd + 15 : undefined);
    if (rPrRegion.includes("<a:solidFill")) {
      result.color = parseColorFromXml(rPrRegion);
    }
  }

  // Font family
  const latinMatch = /<a:latin\s+typeface="([^"]+)"/.exec(rawXml);
  if (latinMatch) {
    result.fontFamily = latinMatch[1];
  }

  // Rotation (on bodyPr)
  const rotMatch = /<a:bodyPr[^>]*\s+rot="(-?\d+)"/.exec(rawXml);
  if (rotMatch) {
    result.rotation = parseInt(rotMatch[1], 10);
  }

  return result;
}

/**
 * Get font size from a ChartTextProperties object (returns points, e.g. 10).
 */
export function getTxPrFontSize(txPr: ChartTextProperties): number | undefined {
  const parsed = isRawXml(txPr) ? parseTxPr(txPr) : txPr;
  return parsed.size !== undefined ? parsed.size / 100 : undefined;
}

/**
 * Get font color from a ChartTextProperties object.
 */
export function getTxPrColor(txPr: ChartTextProperties): ChartColor | undefined {
  const parsed = isRawXml(txPr) ? parseTxPr(txPr) : txPr;
  return parsed.color;
}

// ============================================================================
// Generate spPr/txPr XML from structured models
// ============================================================================

function colorToXml(color: ChartColor, indent: string): string {
  if (color.srgb) {
    const children: string[] = [];
    if (color.alpha !== undefined) {
      children.push(`<a:alpha val="${color.alpha}"/>`);
    }
    if (color.tint !== undefined) {
      children.push(`<a:tint val="${Math.round(color.tint * 100000)}"/>`);
    }
    if (color.lumMod !== undefined) {
      children.push(`<a:lumMod val="${color.lumMod}"/>`);
    }
    if (color.lumOff !== undefined) {
      children.push(`<a:lumOff val="${color.lumOff}"/>`);
    }
    if (color.shade !== undefined) {
      children.push(`<a:shade val="${color.shade}"/>`);
    }
    if (color.satMod !== undefined) {
      children.push(`<a:satMod val="${color.satMod}"/>`);
    }
    if (children.length > 0) {
      return `${indent}<a:srgbClr val="${color.srgb}">${children.join("")}</a:srgbClr>`;
    }
    return `${indent}<a:srgbClr val="${color.srgb}"/>`;
  }
  if (color.theme !== undefined) {
    const themeNames = [
      "dk1",
      "lt1",
      "dk2",
      "lt2",
      "accent1",
      "accent2",
      "accent3",
      "accent4",
      "accent5",
      "accent6",
      "hlink",
      "folHlink"
    ];
    const themeName = themeNames[color.theme] ?? "dk1";
    const children: string[] = [];
    if (color.alpha !== undefined) {
      children.push(`<a:alpha val="${color.alpha}"/>`);
    }
    if (color.lumMod !== undefined) {
      children.push(`<a:lumMod val="${color.lumMod}"/>`);
    }
    if (color.lumOff !== undefined) {
      children.push(`<a:lumOff val="${color.lumOff}"/>`);
    }
    if (color.tint !== undefined) {
      children.push(`<a:tint val="${Math.round(color.tint * 100000)}"/>`);
    }
    if (color.shade !== undefined) {
      children.push(`<a:shade val="${color.shade}"/>`);
    }
    if (color.satMod !== undefined) {
      children.push(`<a:satMod val="${color.satMod}"/>`);
    }
    if (children.length > 0) {
      return `${indent}<a:schemeClr val="${themeName}">${children.join("")}</a:schemeClr>`;
    }
    return `${indent}<a:schemeClr val="${themeName}"/>`;
  }
  if (color.sysClr) {
    const children: string[] = [];
    if (color.alpha !== undefined) {
      children.push(`<a:alpha val="${color.alpha}"/>`);
    }
    if (color.tint !== undefined) {
      children.push(`<a:tint val="${Math.round(color.tint * 100000)}"/>`);
    }
    if (color.lumMod !== undefined) {
      children.push(`<a:lumMod val="${color.lumMod}"/>`);
    }
    if (color.lumOff !== undefined) {
      children.push(`<a:lumOff val="${color.lumOff}"/>`);
    }
    if (color.shade !== undefined) {
      children.push(`<a:shade val="${color.shade}"/>`);
    }
    if (color.satMod !== undefined) {
      children.push(`<a:satMod val="${color.satMod}"/>`);
    }
    if (children.length > 0) {
      return `${indent}<a:sysClr val="${color.sysClr}">${children.join("")}</a:sysClr>`;
    }
    return `${indent}<a:sysClr val="${color.sysClr}"/>`;
  }
  if (color.prstClr) {
    const children: string[] = [];
    if (color.alpha !== undefined) {
      children.push(`<a:alpha val="${color.alpha}"/>`);
    }
    if (color.tint !== undefined) {
      children.push(`<a:tint val="${Math.round(color.tint * 100000)}"/>`);
    }
    if (color.lumMod !== undefined) {
      children.push(`<a:lumMod val="${color.lumMod}"/>`);
    }
    if (color.lumOff !== undefined) {
      children.push(`<a:lumOff val="${color.lumOff}"/>`);
    }
    if (color.shade !== undefined) {
      children.push(`<a:shade val="${color.shade}"/>`);
    }
    if (color.satMod !== undefined) {
      children.push(`<a:satMod val="${color.satMod}"/>`);
    }
    if (children.length > 0) {
      return `${indent}<a:prstClr val="${color.prstClr}">${children.join("")}</a:prstClr>`;
    }
    return `${indent}<a:prstClr val="${color.prstClr}"/>`;
  }
  return "";
}

/**
 * Build a structured ShapeProperties object from the given properties.
 *
 * Previous versions generated `_rawXml` but this caused data loss for
 * effectList/scene3d/sp3d. The returned object is now purely structured
 * and rendered correctly by `_renderSpPr` in chart-space-xform.ts.
 */
export function buildSpPr(props: ShapeProperties): ShapeProperties {
  // Return a purely structured object — no _rawXml.
  // _renderSpPr handles fill, line, effectList, scene3d, sp3d structurally.
  const result: ShapeProperties = {};
  if (props.fill) {
    result.fill = props.fill;
  }
  if (props.line) {
    result.line = props.line;
  }
  if (props.effectList) {
    result.effectList = props.effectList;
  }
  if (props.scene3d) {
    result.scene3d = props.scene3d;
  }
  if (props.sp3d) {
    result.sp3d = props.sp3d;
  }
  return result;
}

/**
 * Build a raw XML txPr string from structured ChartTextProperties.
 * Returns an object with `_rawXml` that can be assigned to `txPr`.
 *
 * Unlike `buildSpPr` which returns purely structured data, `buildTxPr`
 * generates `_rawXml` because `_renderTxPr` in chart-space-xform relies
 * on raw XML passthrough for text properties.
 */
export function buildTxPr(props: ChartTextProperties): ChartTextProperties {
  const rPrAttrs: string[] = [];
  if (props.size !== undefined) {
    rPrAttrs.push(` sz="${props.size}"`);
  }
  if (props.bold) {
    rPrAttrs.push(' b="1"');
  }
  if (props.italic) {
    rPrAttrs.push(' i="1"');
  }

  const rPrChildren: string[] = [];
  if (props.color) {
    rPrChildren.push(`<a:solidFill>${colorToXml(props.color, "")}</a:solidFill>`);
  }
  if (props.fontFamily) {
    rPrChildren.push(`<a:latin typeface="${props.fontFamily}"/>`);
    rPrChildren.push(`<a:cs typeface="${props.fontFamily}"/>`);
  }

  const rPrContent =
    rPrChildren.length > 0
      ? `<a:defRPr${rPrAttrs.join("")}>${rPrChildren.join("")}</a:defRPr>`
      : `<a:defRPr${rPrAttrs.join("")}/>`;

  const bodyPrAttrs = props.rotation !== undefined ? ` rot="${props.rotation}"` : "";

  const xml = [
    "<c:txPr>",
    `  <a:bodyPr${bodyPrAttrs}/>`,
    "  <a:lstStyle/>",
    "  <a:p>",
    `    <a:pPr>${rPrContent}</a:pPr>`,
    "    <a:endParaRPr/>",
    "  </a:p>",
    "</c:txPr>"
  ].join("\n");

  return { ...props, _rawXml: xml };
}

/**
 * Modify a property on an existing spPr (raw or structured).
 * Returns a new ShapeProperties object with the modification applied.
 */
export function setSpPrFill(spPr: ShapeProperties | undefined, fill: ChartFill): ShapeProperties {
  const parsed = spPr && isRawXml(spPr) ? parseSpPr(spPr) : (spPr ?? {});
  return buildSpPr({ ...parsed, fill });
}

/**
 * Modify line properties on an existing spPr.
 */
export function setSpPrLine(spPr: ShapeProperties | undefined, line: ChartLine): ShapeProperties {
  const parsed = spPr && isRawXml(spPr) ? parseSpPr(spPr) : (spPr ?? {});
  return buildSpPr({ ...parsed, line });
}
