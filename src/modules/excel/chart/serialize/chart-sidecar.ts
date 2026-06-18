/**
 * Parser and builder for chart sidecar files:
 * - `styles{N}.xml` (cs:chartStyle) — chart-wide style id
 * - `colors{N}.xml` (cs:colorStyle) — chart color palette
 *
 * These files are normally written and read as opaque XML; this module offers
 * structured access for users who want to inspect or rewrite the palette.
 */
import { escapeXmlAttr } from "@excel/chart/render/chart-utils";

/** indexOf-based attribute extraction — no backtracking risk. */
function extractAttr(xml: string, tagName: string, attrName: string): string | undefined {
  const tagStart = xml.indexOf(`<${tagName}`);
  if (tagStart < 0) {
    return undefined;
  }
  const tagEnd = xml.indexOf(">", tagStart);
  if (tagEnd < 0) {
    return undefined;
  }
  const tag = xml.slice(tagStart, tagEnd + 1);
  const attrStart = tag.indexOf(`${attrName}="`);
  if (attrStart < 0) {
    return undefined;
  }
  const valStart = attrStart + attrName.length + 2;
  const valEnd = tag.indexOf('"', valStart);
  return valEnd >= 0 ? tag.slice(valStart, valEnd) : undefined;
}

import type {
  ChartColorVariation,
  ChartColorsEntry,
  ChartColorsModel,
  ChartStyleElement,
  ChartStyleModel
} from "@excel/chart/model/types";

/**
 * Parse a `colors{N}.xml` raw XML string into a structured model.
 * Returns a best-effort representation — unknown children are preserved in rawXml.
 */
export function parseChartColors(rawXml: string): ChartColorsModel {
  const result: ChartColorsModel = { rawXml };

  const methodMatch = extractAttr(rawXml, "cs:colorStyle", "meth");
  if (methodMatch) {
    result.method = methodMatch;
  }
  const idVal = extractAttr(rawXml, "cs:colorStyle", "id");
  if (idVal) {
    result.id = parseInt(idVal, 10);
  }

  // Colors come as a sequence of <a:schemeClr> and <a:srgbClr> at the top
  // level of <cs:colorStyle>, each optionally wrapped with children.
  //
  // We locate every `<cs:variation>...</cs:variation>` span first so colour
  // blocks sitting *inside* a variation can be excluded in O(1) per match,
  // independent of document order (the previous `lastIndexOf` strategy was
  // order-sensitive — colours declared after a `variation` close tag were
  // classified based on the last variation in the file, not the enclosing
  // element).
  const variationRanges: Array<{ start: number; end: number }> = [];
  const varRangeRe = /<cs:variation\b[^>]*(?:\/>|>[\s\S]*?<\/cs:variation>)/g;
  let vrm: RegExpExecArray | null;
  while ((vrm = varRangeRe.exec(rawXml)) !== null) {
    variationRanges.push({ start: vrm.index, end: vrm.index + vrm[0].length });
  }
  const inVariation = (index: number): boolean =>
    variationRanges.some(r => index >= r.start && index < r.end);

  const colors: ChartColorsEntry[] = [];
  // Match a group starting with schemeClr or srgbClr — may self-close or have nested modifiers.
  const colorBlockRe = /<a:(schemeClr|srgbClr)\b[^>]*(?:\/>|>([\s\S]*?)<\/a:\1>)/g;
  let m: RegExpExecArray | null;
  while ((m = colorBlockRe.exec(rawXml)) !== null) {
    if (inVariation(m.index)) {
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
    // Extract optional child modifiers. Accept fractional / negative
    // values per `shape-properties.ts:parseColorModifiers` — third-party
    // exporters (e.g. LibreOffice, custom palettes) emit `val="-5000"`
    // or `val="12345.67"`; the previous `\d+`-only regex silently
    // dropped them. Round after parsing so wire values stay integers.
    const parseModAttr = (xml: string, tag: string): number | undefined => {
      const m = new RegExp(`<a:${tag}\\s+val="(-?\\d+(?:\\.\\d+)?)"`).exec(xml);
      return m ? Math.round(parseFloat(m[1])) : undefined;
    };
    const inner = m[2] ?? "";
    const lumModV = parseModAttr(inner, "lumMod");
    if (lumModV !== undefined) {
      entry.lumMod = lumModV;
    }
    const lumOffV = parseModAttr(inner, "lumOff");
    if (lumOffV !== undefined) {
      entry.lumOff = lumOffV;
    }
    const tintV = parseModAttr(inner, "tint");
    if (tintV !== undefined) {
      entry.tint = tintV;
    }
    const shadeV = parseModAttr(inner, "shade");
    if (shadeV !== undefined) {
      entry.shade = shadeV;
    }
    const satModV = parseModAttr(inner, "satMod");
    if (satModV !== undefined) {
      entry.satMod = satModV;
    }
    const alphaV = parseModAttr(inner, "alpha");
    if (alphaV !== undefined) {
      entry.alpha = alphaV;
    }
    if (entry.theme || entry.srgb) {
      colors.push(entry);
    }
  }

  if (colors.length > 0) {
    result.colors = colors;
  }

  // `<cs:variation>` blocks — each block's body is a list of
  // DrawingML colour modifiers. We parse them into structured
  // `ChartColorVariation` entries so editors can mutate the gradient
  // palette without string surgery.
  const variations: ChartColorVariation[] = [];
  const variationRe = /<cs:variation\b[^>]*>([\s\S]*?)<\/cs:variation>|<cs:variation\b[^>]*\/>/g;
  let vm: RegExpExecArray | null;
  // Reuse the sibling `parseModAttr` semantics from the colour parser
  // above: accept signed / fractional values. Third-party exports
  // (LibreOffice palette packs, custom theme builders) emit these
  // forms for colour variations; the previous `\d+`-only regex
  // silently dropped them.
  const parseModAttr = (xml: string, tag: string): number | undefined => {
    const m = new RegExp(`<a:${tag}\\b[^>]*\\bval="(-?\\d+(?:\\.\\d+)?)"`).exec(xml);
    return m ? Math.round(parseFloat(m[1])) : undefined;
  };
  while ((vm = variationRe.exec(rawXml)) !== null) {
    const body = vm[1] ?? "";
    const entry: ChartColorVariation = {};
    const lumModV = parseModAttr(body, "lumMod");
    if (lumModV !== undefined) {
      entry.lumMod = lumModV;
    }
    const lumOffV = parseModAttr(body, "lumOff");
    if (lumOffV !== undefined) {
      entry.lumOff = lumOffV;
    }
    const tintV = parseModAttr(body, "tint");
    if (tintV !== undefined) {
      entry.tint = tintV;
    }
    const shadeV = parseModAttr(body, "shade");
    if (shadeV !== undefined) {
      entry.shade = shadeV;
    }
    const satModV = parseModAttr(body, "satMod");
    if (satModV !== undefined) {
      entry.satMod = satModV;
    }
    const alphaV = parseModAttr(body, "alpha");
    if (alphaV !== undefined) {
      entry.alpha = alphaV;
    }
    // Push even empty variations so the count survives round-trip —
    // Excel uses the presence of a `<cs:variation/>` self-closing tag
    // as a gradient-index placeholder.
    variations.push(entry);
  }
  if (variations.length > 0) {
    result.variations = variations;
  }
  return result;
}

/**
 * Build a `colors{N}.xml` string from a structured ChartColorsModel.
 *
 * Structured data (`colors` / `variations`) takes precedence over `rawXml`.
 * If neither `colors` nor `variations` is present we fall back to `rawXml`
 * (round-trip) or the built-in default palette. Mutations to `variations`
 * alone therefore still take effect without requiring the caller to also
 * populate `colors`.
 */
export function buildChartColors(model: ChartColorsModel): string {
  const hasColors = !!model.colors && model.colors.length > 0;
  const hasVariations = !!model.variations && model.variations.length > 0;
  if (!hasColors && !hasVariations) {
    return model.rawXml ?? buildDefaultChartColors();
  }
  const parts: string[] = [];
  parts.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  const attrs: string[] = [
    'xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"',
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
  ];
  if (model.method !== undefined) {
    attrs.push(`meth="${escapeXmlAttr(model.method)}"`);
  }
  if (model.id !== undefined) {
    attrs.push(`id="${model.id}"`);
  }
  parts.push(`<cs:colorStyle ${attrs.join(" ")}>`);
  if (hasColors) {
    for (const c of model.colors!) {
      parts.push(colorEntryToXml(c));
    }
  }
  if (hasVariations) {
    for (const v of model.variations!) {
      parts.push(variationEntryToXml(v));
    }
  }
  parts.push("</cs:colorStyle>");
  return parts.join("\n");
}

// `escapeXmlAttr` is imported from `chart-utils`; it delegates to the
// canonical `xmlEncodeAttr` from `@xml/encode`. Sidecar output now
// shares byte-for-byte escaping semantics with the rest of the chart
// pipeline — previously this module carried a local copy that omitted
// the CR/LF/Tab → numeric-ref encoding, letting whitespace-laden
// attribute values lose data on round-trip.

function variationEntryToXml(v: ChartColorVariation): string {
  // Modifier values must be valid `xsd:int` on the wire. Interpolating
  // raw model numbers directly let `NaN` / `Infinity` / unrounded
  // floats through to the attribute, which Excel's strict reader
  // rejects as "invalid xs:int". Guard each slot with
  // `Number.isFinite` + `Math.round` so the sidecar consistently
  // produces well-formed XML even when a public API consumer hands in
  // an out-of-range number.
  const children: string[] = [];
  const emitInt = (tag: string, value: number | undefined): void => {
    if (value === undefined || !Number.isFinite(value)) {
      return;
    }
    children.push(`<a:${tag} val="${Math.round(value)}"/>`);
  };
  emitInt("lumMod", v.lumMod);
  emitInt("lumOff", v.lumOff);
  emitInt("tint", v.tint);
  emitInt("shade", v.shade);
  emitInt("satMod", v.satMod);
  emitInt("alpha", v.alpha);
  if (children.length === 0) {
    return "  <cs:variation/>";
  }
  return `  <cs:variation>${children.join("")}</cs:variation>`;
}

// =============================================================================
// Default chartStyle / chartColors XML — mirrored byte-for-byte from Microsoft
// Excel 2021's own `style1.xml` / `colors1.xml` output for a freshly-inserted
// waterfall chart. See the rationale on `buildChartStyle` /
// `buildChartColors` for why we embed the full content rather than emitting
// an id-only stub (Excel 2016+ rejects the stub form and drops the chartEx
// part + parent drawing with "Removed Part: drawing${N}.xml (Drawing shape)").
// These constants are the single source of truth for the Office built-in
// style id=395 + colors id=10 defaults.
// =============================================================================

const DEFAULT_CHART_STYLE_XML =
  '<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="395"><cs:axisTitle><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="900"/></cs:axisTitle><cs:categoryAxis><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz="900"/></cs:categoryAxis><cs:chartArea mods="allowNoFillOverride allowNoLineOverride"><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz="1000"/></cs:chartArea><cs:dataLabel><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="900"/></cs:dataLabel><cs:dataLabelCallout><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="dk1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:ln><a:solidFill><a:schemeClr val="dk1"><a:lumMod val="25000"/><a:lumOff val="75000"/></a:schemeClr></a:solidFill></a:ln></cs:spPr><cs:defRPr sz="900"/><cs:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="clip" horzOverflow="clip" vert="horz" wrap="square" lIns="36576" tIns="18288" rIns="36576" bIns="18288" anchor="ctr" anchorCtr="1"><a:spAutoFit/></cs:bodyPr></cs:dataLabelCallout><cs:dataPoint><cs:lnRef idx="0"/><cs:fillRef idx="0"><cs:styleClr val="auto"/></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></cs:spPr></cs:dataPoint><cs:dataPoint3D><cs:lnRef idx="0"/><cs:fillRef idx="0"><cs:styleClr val="auto"/></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></cs:spPr></cs:dataPoint3D><cs:dataPointLine><cs:lnRef idx="0"><cs:styleClr val="auto"/></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:round/></a:ln></cs:spPr></cs:dataPointLine><cs:dataPointMarker><cs:lnRef idx="0"/><cs:fillRef idx="0"><cs:styleClr val="auto"/></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:ln w="9525"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></cs:spPr></cs:dataPointMarker><cs:dataPointMarkerLayout symbol="circle" size="5"/><cs:dataPointWireframe><cs:lnRef idx="0"><cs:styleClr val="auto"/></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:round/></a:ln></cs:spPr></cs:dataPointWireframe><cs:dataTable><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:spPr><a:ln w="9525"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill></a:ln></cs:spPr><cs:defRPr sz="900"/></cs:dataTable><cs:downBar><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="dk1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="dk1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:ln w="9525"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill></a:ln></cs:spPr></cs:downBar><cs:dropLine><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="35000"/><a:lumOff val="65000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:dropLine><cs:errorBar><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:errorBar><cs:floor><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:floor><cs:gridlineMajor><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:gridlineMajor><cs:gridlineMinor><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:gridlineMinor><cs:hiLoLine><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:hiLoLine><cs:leaderLine><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="35000"/><a:lumOff val="65000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:leaderLine><cs:legend><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="900"/></cs:legend><cs:plotArea mods="allowNoFillOverride allowNoLineOverride"><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:plotArea><cs:plotArea3D mods="allowNoFillOverride allowNoLineOverride"><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:plotArea3D><cs:seriesAxis><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz="900"/></cs:seriesAxis><cs:seriesLine><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill><a:round/></a:ln></cs:spPr></cs:seriesLine><cs:title><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="1400"/></cs:title><cs:trendline><cs:lnRef idx="0"><cs:styleClr val="auto"/></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="19050" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="sysDash"/></a:ln></cs:spPr></cs:trendline><cs:trendlineLabel><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="900"/></cs:trendlineLabel><cs:upBar><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="dk1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:ln w="9525"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill></a:ln></cs:spPr></cs:upBar><cs:valueAxis><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="900"/></cs:valueAxis><cs:wall><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:wall></cs:chartStyle>';

const DEFAULT_CHART_COLORS_XML =
  '<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10"><a:schemeClr val="accent1"/><a:schemeClr val="accent2"/><a:schemeClr val="accent3"/><a:schemeClr val="accent4"/><a:schemeClr val="accent5"/><a:schemeClr val="accent6"/><cs:variation/><cs:variation><a:lumMod val="60000"/></cs:variation><cs:variation><a:lumMod val="80000"/><a:lumOff val="20000"/></cs:variation><cs:variation><a:lumMod val="80000"/></cs:variation><cs:variation><a:lumMod val="60000"/><a:lumOff val="40000"/></cs:variation><cs:variation><a:lumMod val="50000"/></cs:variation><cs:variation><a:lumMod val="70000"/><a:lumOff val="30000"/></cs:variation><cs:variation><a:lumMod val="70000"/></cs:variation><cs:variation><a:lumMod val="50000"/><a:lumOff val="50000"/></cs:variation></cs:colorStyle>';

export function buildChartStyle(model: ChartStyleModel): string {
  // Prefer raw XML for byte-preserving round-trip of files that
  // haven't been mutated. Callers who want their `elements`
  // modifications to take effect drop `rawXml` first (the same
  // convention as `renderChartEx`).
  if (model.rawXml) {
    return model.rawXml;
  }
  const id = model.id ?? 395;
  // Always emit the full Microsoft Excel default `<cs:chartStyle>`
  // template with the caller's requested `id`. Partial element
  // overrides declared via `ChartStyleModel.elements` are NOT safe
  // to splice onto the template: the public `ChartStyleElement` type
  // only exposes the four `*RefIdx` slots + a handful of raw-string
  // escape hatches, so a merged element inevitably misses the
  // `<cs:spPr>`, `<cs:defRPr>`, and fully-specified `<cs:fontRef>`
  // children that Excel's `CT_StyleEntry` treats as required. The
  // result (verified against Excel 2021 via
  // `tmp/reference-styled.xlsx` vs the library's previous output)
  // is a schema-shaped sidecar that Excel's strict loader still
  // drops with "Removed Part: /xl/drawings/drawingN.xml part.
  // (Drawing shape)".
  //
  // Authors who need genuine style customisation should:
  //   1. Pre-build the sidecar XML themselves and pass it via
  //      `rawXml` (honoured by the early return above), OR
  //   2. Call `setBuiltInStyle(N)` on the chart instead — the
  //      classic `<c:style val="N"/>` element is schema-legal
  //      standalone and doesn't require a sidecar.
  //
  // Dropping `elements` silently would be a bigger regression, so
  // log a one-time warning in non-production builds when an author
  // passes partial elements we're about to discard.
  if (
    model.elements &&
    Object.keys(model.elements).length > 0 &&
    typeof process !== "undefined" &&
    process.env?.NODE_ENV !== "production"
  ) {
    console.warn(
      "[excelts] buildChartStyle: `model.elements` overrides are not yet supported — " +
        "the default Excel style template will be emitted (customisation ignored). " +
        "Pass a pre-built sidecar via `rawXml` if you need genuine overrides."
    );
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${DEFAULT_CHART_STYLE_XML.replace(
    /id="\d+"/,
    `id="${id}"`
  )}`;
}

function colorEntryToXml(c: ChartColorsEntry): string {
  // Same guard rationale as `variationEntryToXml` — reject non-finite
  // modifier values so the serialiser can never produce
  // `<a:tint val="NaN"/>` etc.
  const mods: string[] = [];
  const emitInt = (tag: string, value: number | undefined): void => {
    if (value === undefined || !Number.isFinite(value)) {
      return;
    }
    mods.push(`<a:${tag} val="${Math.round(value)}"/>`);
  };
  emitInt("lumMod", c.lumMod);
  emitInt("lumOff", c.lumOff);
  emitInt("tint", c.tint);
  emitInt("shade", c.shade);
  emitInt("satMod", c.satMod);
  emitInt("alpha", c.alpha);

  let tag: string;
  let val: string;
  if (c.theme) {
    tag = "a:schemeClr";
    // Theme / srgb tokens live in attribute context — escape in case
    // a consumer passes an unvalidated string. `escapeXmlAttr`
    // strips illegal XML chars and encodes reserved entities.
    val = escapeXmlAttr(c.theme);
  } else if (c.srgb) {
    tag = "a:srgbClr";
    val = escapeXmlAttr(c.srgb);
  } else {
    return "";
  }
  if (mods.length === 0) {
    return `  <${tag} val="${val}"/>`;
  }
  return `  <${tag} val="${val}">${mods.join("")}</${tag}>`;
}

function buildDefaultChartColors(): string {
  // Excel 2016+ REJECTS stub `colorStyle` files that list only the
  // 6 accent `schemeClr` entries without the 9 follow-on `variation`
  // siblings — the chartEx is silently dropped along with its parent
  // drawing ("Removed Part: drawingN.xml (Drawing shape)"). Emit the
  // full Microsoft-authored default (id=10, 6 accents + 9 variations)
  // byte-for-byte; the template is mirrored from Excel 2021's own
  // `colors1.xml` output for a freshly-inserted chartEx. See the
  // matching rationale on `DEFAULT_CHART_STYLE_XML`.
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${DEFAULT_CHART_COLORS_XML}`;
}

/**
 * Parse a `styles{N}.xml` file to extract the style ID.
 */
export function parseChartStyle(rawXml: string): ChartStyleModel {
  const result: ChartStyleModel = { rawXml };
  const idVal = extractAttr(rawXml, "cs:chartStyle", "id");
  if (idVal) {
    result.id = parseInt(idVal, 10);
  }

  // Walk every top-level `<cs:…>` child of `<cs:chartStyle>` and
  // capture it as a {@link ChartStyleElement}. The structured slots
  // (`lnRefIdx`, `fillRefIdx`, `effectRefIdx`, `fontRefIdx`) are
  // extracted directly; `spPr`, `defRPr`, `bodyPr`, `fontRef` inner
  // XML are held verbatim so the DrawingML sub-tree round-trips
  // without this module enumerating every child type.
  const elements: Record<string, ChartStyleElement> = {};
  // Strip the outer <cs:chartStyle> wrapper so the outer regex doesn't
  // match itself.
  const outerOpen = /<cs:chartStyle\b[^>]*>/.exec(rawXml);
  const outerClose = rawXml.lastIndexOf("</cs:chartStyle>");
  if (outerOpen && outerClose >= 0) {
    const body = rawXml.slice(outerOpen.index + outerOpen[0].length, outerClose);
    // Match either <cs:name …/> self-closing or <cs:name …>…</cs:name>.
    const childRe = /<cs:([A-Za-z][A-Za-z0-9]*)\b([^>]*?)(?:\/>|>([\s\S]*?)<\/cs:\1>)/g;
    let m: RegExpExecArray | null;
    while ((m = childRe.exec(body)) !== null) {
      const name = m[1];
      const attrStr = m[2];
      const inner = m[3] ?? "";
      const entry: ChartStyleElement = {};

      const lnRef = /<cs:lnRef\b[^>]*\bidx="(\d+)"/.exec(inner);
      if (lnRef) {
        entry.lnRefIdx = parseInt(lnRef[1], 10);
      }
      const fillRef = /<cs:fillRef\b[^>]*\bidx="(\d+)"/.exec(inner);
      if (fillRef) {
        entry.fillRefIdx = parseInt(fillRef[1], 10);
      }
      const effectRef = /<cs:effectRef\b[^>]*\bidx="(\d+)"/.exec(inner);
      if (effectRef) {
        entry.effectRefIdx = parseInt(effectRef[1], 10);
      }

      const fontRefMatch =
        /<cs:fontRef\b[^>]*\bidx="([^"]+)"[^>]*(?:\/>|>([\s\S]*?)<\/cs:fontRef>)/.exec(inner);
      if (fontRefMatch) {
        const idxVal = fontRefMatch[1];
        if (idxVal === "minor" || idxVal === "major" || idxVal === "none") {
          entry.fontRefIdx = idxVal;
        }
        if (fontRefMatch[2] !== undefined) {
          entry.fontRefBody = fontRefMatch[2];
        }
      }

      const spPr = /<cs:spPr\b[^>]*>[\s\S]*?<\/cs:spPr>|<cs:spPr\b[^>]*\/>/.exec(inner);
      if (spPr) {
        entry.spPrXml = spPr[0];
      }
      const defRPr = /<cs:defRPr\b[^>]*(?:\/>|>[\s\S]*?<\/cs:defRPr>)/.exec(inner);
      if (defRPr) {
        entry.defRPrXml = defRPr[0];
      }
      const bodyPr = /<cs:bodyPr\b[^>]*(?:\/>|>[\s\S]*?<\/cs:bodyPr>)/.exec(inner);
      if (bodyPr) {
        entry.bodyPrXml = bodyPr[0];
      }

      const attrs: Record<string, string> = {};
      const attrRe = /\b([A-Za-z][A-Za-z0-9:_-]*)\s*=\s*"([^"]*)"/g;
      let am: RegExpExecArray | null;
      while ((am = attrRe.exec(attrStr)) !== null) {
        attrs[am[1]] = am[2];
      }
      if (Object.keys(attrs).length > 0) {
        entry.attributes = attrs;
      }
      elements[name] = entry;
    }
  }
  if (Object.keys(elements).length > 0) {
    result.elements = elements;
  }
  return result;
}
