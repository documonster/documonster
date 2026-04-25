/**
 * Parser and builder for chart sidecar files:
 * - `styles{N}.xml` (cs:chartStyle) — chart-wide style id
 * - `colors{N}.xml` (cs:colorStyle) — chart color palette
 *
 * These files are normally written and read as opaque XML; this module offers
 * structured access for users who want to inspect or rewrite the palette.
 */
import type { ChartColorsEntry, ChartColorsModel, ChartStyleModel } from "./types";

/**
 * Parse a `colors{N}.xml` raw XML string into a structured model.
 * Returns a best-effort representation — unknown children are preserved in rawXml.
 */
export function parseChartColors(rawXml: string): ChartColorsModel {
  const result: ChartColorsModel = { rawXml };

  const methodMatch = /<cs:colorStyle\s+[^>]*\bmeth="([^"]+)"/.exec(rawXml);
  if (methodMatch) {
    result.method = methodMatch[1];
  }
  const idMatch = /<cs:colorStyle\s+[^>]*\bid="(\d+)"/.exec(rawXml);
  if (idMatch) {
    result.id = parseInt(idMatch[1], 10);
  }

  // Colors come as a sequence of <a:schemeClr> and <a:srgbClr> at the top
  // level of <cs:colorStyle>, each optionally wrapped with children.
  const colors: ChartColorsEntry[] = [];
  // Match a group starting with schemeClr or srgbClr — may self-close or have nested modifiers.
  const colorBlockRe = /<a:(schemeClr|srgbClr)\b[^>]*(?:\/>|>([\s\S]*?)<\/a:\1>)/g;
  let m: RegExpExecArray | null;
  while ((m = colorBlockRe.exec(rawXml)) !== null) {
    // Skip if this block is inside a <cs:variation> — variations are fancier
    // per-color modifiers; rawXml round-trip keeps them.
    const before = rawXml.slice(0, m.index);
    const lastVar = before.lastIndexOf("<cs:variation");
    const lastVarEnd = before.lastIndexOf("</cs:variation");
    if (lastVar > lastVarEnd) {
      continue;
    }
    const entry: ChartColorsEntry = {};
    const fullMatch = m[0];
    const valMatch = /\bval="([^"]+)"/.exec(fullMatch);
    if (m[1] === "schemeClr" && valMatch) {
      entry.theme = valMatch[1];
    } else if (m[1] === "srgbClr" && valMatch) {
      entry.srgb = valMatch[1];
    }
    // Extract optional child modifiers
    const inner = m[2] ?? "";
    const lumModM = /<a:lumMod\s+val="(\d+)"/.exec(inner);
    if (lumModM) {
      entry.lumMod = parseInt(lumModM[1], 10);
    }
    const lumOffM = /<a:lumOff\s+val="(\d+)"/.exec(inner);
    if (lumOffM) {
      entry.lumOff = parseInt(lumOffM[1], 10);
    }
    const tintM = /<a:tint\s+val="(\d+)"/.exec(inner);
    if (tintM) {
      entry.tint = parseInt(tintM[1], 10);
    }
    const shadeM = /<a:shade\s+val="(\d+)"/.exec(inner);
    if (shadeM) {
      entry.shade = parseInt(shadeM[1], 10);
    }
    const satModM = /<a:satMod\s+val="(\d+)"/.exec(inner);
    if (satModM) {
      entry.satMod = parseInt(satModM[1], 10);
    }
    const alphaM = /<a:alpha\s+val="(\d+)"/.exec(inner);
    if (alphaM) {
      entry.alpha = parseInt(alphaM[1], 10);
    }
    if (entry.theme || entry.srgb) {
      colors.push(entry);
    }
  }

  if (colors.length > 0) {
    result.colors = colors;
  }
  return result;
}

/**
 * Build a `colors{N}.xml` string from a structured ChartColorsModel.
 * If `model.colors` is provided, it takes precedence over `rawXml`.
 */
export function buildChartColors(model: ChartColorsModel): string {
  if (!model.colors || model.colors.length === 0) {
    // Fall back to round-trip rawXml if available
    return model.rawXml ?? buildDefaultChartColors();
  }
  const parts: string[] = [];
  parts.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  const attrs: string[] = [
    'xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"',
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
  ];
  if (model.method !== undefined) {
    attrs.push(`meth="${model.method}"`);
  }
  if (model.id !== undefined) {
    attrs.push(`id="${model.id}"`);
  }
  parts.push(`<cs:colorStyle ${attrs.join(" ")}>`);
  for (const c of model.colors) {
    parts.push(colorEntryToXml(c));
  }
  parts.push("</cs:colorStyle>");
  return parts.join("\n");
}

function colorEntryToXml(c: ChartColorsEntry): string {
  const mods: string[] = [];
  if (c.lumMod !== undefined) {
    mods.push(`<a:lumMod val="${c.lumMod}"/>`);
  }
  if (c.lumOff !== undefined) {
    mods.push(`<a:lumOff val="${c.lumOff}"/>`);
  }
  if (c.tint !== undefined) {
    mods.push(`<a:tint val="${c.tint}"/>`);
  }
  if (c.shade !== undefined) {
    mods.push(`<a:shade val="${c.shade}"/>`);
  }
  if (c.satMod !== undefined) {
    mods.push(`<a:satMod val="${c.satMod}"/>`);
  }
  if (c.alpha !== undefined) {
    mods.push(`<a:alpha val="${c.alpha}"/>`);
  }

  let tag: string;
  let val: string;
  if (c.theme) {
    tag = "a:schemeClr";
    val = c.theme;
  } else if (c.srgb) {
    tag = "a:srgbClr";
    val = c.srgb;
  } else {
    return "";
  }
  if (mods.length === 0) {
    return `  <${tag} val="${val}"/>`;
  }
  return `  <${tag} val="${val}">${mods.join("")}</${tag}>`;
}

function buildDefaultChartColors(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" ',
    '  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10">',
    '  <a:schemeClr val="accent1"/>',
    '  <a:schemeClr val="accent2"/>',
    '  <a:schemeClr val="accent3"/>',
    '  <a:schemeClr val="accent4"/>',
    '  <a:schemeClr val="accent5"/>',
    '  <a:schemeClr val="accent6"/>',
    "</cs:colorStyle>"
  ].join("\n");
}

/**
 * Parse a `styles{N}.xml` file to extract the style ID.
 */
export function parseChartStyle(rawXml: string): ChartStyleModel {
  const result: ChartStyleModel = { rawXml };
  const idMatch = /<cs:chartStyle\s+[^>]*\bid="(\d+)"/.exec(rawXml);
  if (idMatch) {
    result.id = parseInt(idMatch[1], 10);
  }
  return result;
}
