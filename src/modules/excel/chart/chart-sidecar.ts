/**
 * Parser and builder for chart sidecar files:
 * - `styles{N}.xml` (cs:chartStyle) — chart-wide style id
 * - `colors{N}.xml` (cs:colorStyle) — chart color palette
 *
 * These files are normally written and read as opaque XML; this module offers
 * structured access for users who want to inspect or rewrite the palette.
 */
import type {
  ChartColorVariation,
  ChartColorsEntry,
  ChartColorsModel,
  ChartStyleElement,
  ChartStyleModel
} from "./types";

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

  // `<cs:variation>` blocks — each block's body is a list of
  // DrawingML colour modifiers. We parse them into structured
  // `ChartColorVariation` entries so editors can mutate the gradient
  // palette without string surgery.
  const variations: ChartColorVariation[] = [];
  const variationRe = /<cs:variation\b[^>]*>([\s\S]*?)<\/cs:variation>|<cs:variation\b[^>]*\/>/g;
  let vm: RegExpExecArray | null;
  while ((vm = variationRe.exec(rawXml)) !== null) {
    const body = vm[1] ?? "";
    const entry: ChartColorVariation = {};
    const lumMod = /<a:lumMod\b[^>]*\bval="(\d+)"/.exec(body);
    if (lumMod) {
      entry.lumMod = parseInt(lumMod[1], 10);
    }
    const lumOff = /<a:lumOff\b[^>]*\bval="(\d+)"/.exec(body);
    if (lumOff) {
      entry.lumOff = parseInt(lumOff[1], 10);
    }
    const tint = /<a:tint\b[^>]*\bval="(\d+)"/.exec(body);
    if (tint) {
      entry.tint = parseInt(tint[1], 10);
    }
    const shade = /<a:shade\b[^>]*\bval="(\d+)"/.exec(body);
    if (shade) {
      entry.shade = parseInt(shade[1], 10);
    }
    const satMod = /<a:satMod\b[^>]*\bval="(\d+)"/.exec(body);
    if (satMod) {
      entry.satMod = parseInt(satMod[1], 10);
    }
    const alpha = /<a:alpha\b[^>]*\bval="(\d+)"/.exec(body);
    if (alpha) {
      entry.alpha = parseInt(alpha[1], 10);
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
  if (model.variations) {
    for (const v of model.variations) {
      parts.push(variationEntryToXml(v));
    }
  }
  parts.push("</cs:colorStyle>");
  return parts.join("\n");
}

function variationEntryToXml(v: ChartColorVariation): string {
  const children: string[] = [];
  if (v.lumMod !== undefined) {
    children.push(`<a:lumMod val="${v.lumMod}"/>`);
  }
  if (v.lumOff !== undefined) {
    children.push(`<a:lumOff val="${v.lumOff}"/>`);
  }
  if (v.tint !== undefined) {
    children.push(`<a:tint val="${v.tint}"/>`);
  }
  if (v.shade !== undefined) {
    children.push(`<a:shade val="${v.shade}"/>`);
  }
  if (v.satMod !== undefined) {
    children.push(`<a:satMod val="${v.satMod}"/>`);
  }
  if (v.alpha !== undefined) {
    children.push(`<a:alpha val="${v.alpha}"/>`);
  }
  if (children.length === 0) {
    return "  <cs:variation/>";
  }
  return `  <cs:variation>${children.join("")}</cs:variation>`;
}

export function buildChartStyle(model: ChartStyleModel): string {
  // Prefer raw XML for byte-preserving round-trip of files that
  // haven't been mutated. Callers who want their `elements`
  // modifications to take effect drop `rawXml` first (the same
  // convention as `renderChartEx`).
  if (model.rawXml) {
    return model.rawXml;
  }
  if (model.elements && Object.keys(model.elements).length > 0) {
    const id = model.id ?? 227;
    const parts: string[] = [];
    parts.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    parts.push(
      `<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="${id}">`
    );
    for (const [name, element] of Object.entries(model.elements)) {
      parts.push(chartStyleElementToXml(name, element));
    }
    parts.push(`</cs:chartStyle>`);
    return parts.join("");
  }
  const id = model.id ?? 227;
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="${id}"/>`
  ].join("\n");
}

function chartStyleElementToXml(name: string, entry: ChartStyleElement): string {
  const attrParts: string[] = [];
  if (entry.attributes) {
    for (const [k, v] of Object.entries(entry.attributes)) {
      attrParts.push(`${k}="${v.replace(/"/g, "&quot;")}"`);
    }
  }
  const body: string[] = [];
  if (entry.lnRefIdx !== undefined) {
    body.push(`<cs:lnRef idx="${entry.lnRefIdx}"/>`);
  }
  if (entry.fillRefIdx !== undefined) {
    body.push(`<cs:fillRef idx="${entry.fillRefIdx}"/>`);
  }
  if (entry.effectRefIdx !== undefined) {
    body.push(`<cs:effectRef idx="${entry.effectRefIdx}"/>`);
  }
  if (entry.fontRefIdx !== undefined) {
    if (entry.fontRefBody) {
      body.push(`<cs:fontRef idx="${entry.fontRefIdx}">${entry.fontRefBody}</cs:fontRef>`);
    } else {
      body.push(`<cs:fontRef idx="${entry.fontRefIdx}"/>`);
    }
  }
  if (entry.spPrXml) {
    body.push(entry.spPrXml);
  }
  if (entry.defRPrXml) {
    body.push(entry.defRPrXml);
  }
  if (entry.bodyPrXml) {
    body.push(entry.bodyPrXml);
  }
  const attrs = attrParts.length > 0 ? " " + attrParts.join(" ") : "";
  if (body.length === 0) {
    return `<cs:${name}${attrs}/>`;
  }
  return `<cs:${name}${attrs}>${body.join("")}</cs:${name}>`;
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
