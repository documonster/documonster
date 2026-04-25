/**
 * DOCX Module - Reader / Parser
 *
 * Reads a DOCX ZIP file and parses it into a DocxDocument model.
 * Uses the archive module for ZIP reading and XML module for parsing.
 */

import { unzip } from "@archive/read-archive";
import { parseXml, findChild, findChildren, textContent } from "@xml/dom";
import type { XmlElement } from "@xml/types";

import { RelType } from "./constants";
import { DocxError, DocxParseError, DocxMissingPartError } from "./errors";
import type {
  DocxDocument,
  BodyContent,
  NoteType,
  Paragraph,
  ParagraphChild,
  ParagraphProperties,
  Run,
  RunProperties,
  RunContent,
  Table,
  TableRow,
  TableCell,
  TableProperties,
  TableRowProperties,
  TableCellProperties,
  TableBorders,
  TableWidth,
  TableCellMargins,
  TableFloat,
  Border,
  Shading,
  TabStop,
  SectionProperties,
  PageBorders,
  HeaderFooterRef,
  StyleDef,
  DocDefaults,
  AbstractNumbering,
  NumPicBullet,
  NumberingInstance,
  NumberingLevel,
  LevelOverride,
  HeaderDef,
  FooterDef,
  HeaderFooterContent,
  CoreProperties,
  AppProperties,
  FontDef,
  EmbeddedFont,
  CustomXmlPart,
  WebSettings,
  PersonInfo,
  DocumentSettings,
  ImageDef,
  ImageMediaType,
  FontSpec,
  ParagraphFrame,
  CommentDef,
  InsertedRun,
  DeletedRun,
  RevisionInfo,
  FloatingImage,
  MathContent,
  MathBlock,
  TextBox,
  StructuredDocumentTag,
  SdtProperties,
  SdtListItem,
  SdtDateProperties,
  TableOfContents,
  CheckBox,
  DocumentBackground,
  CustomProperty,
  CustomPropertyValue,
  ColorSpec,
  RunPropertyChange,
  SectionPropertyChange,
  FieldContent,
  FormField,
  TableStyleConditionalFormat,
  DocumentTheme,
  ThemeColorName,
  Watermark,
  TextWatermark,
  ImageWatermark,
  DrawingShape,
  OpaquePart,
  OpaqueRelationship,
  OpaqueDrawing
} from "./types";

// Module-level parsing context: set at the start of readDocx, used by parseParagraph etc.
let _parseRelMap: Map<string, ParsedRelationship> = new Map();

// =============================================================================
// Helper Functions
// =============================================================================

function attrVal(el: XmlElement, name: string): string | undefined {
  // Try with w: prefix and without
  return el.attributes[`w:${name}`] ?? el.attributes[name];
}

function attrInt(el: XmlElement, name: string): number | undefined {
  const v = attrVal(el, name);
  if (v === undefined) {
    return undefined;
  }
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? n : undefined;
}

/**
 * Read an attribute as a strict boolean:
 *   "1"|"true"  → true
 *   "0"|"false" → false
 *   otherwise   → undefined
 * Useful for attributes where default-false vs explicit-false matters.
 */
function _attrBool(el: XmlElement, name: string): boolean | undefined {
  const v = attrVal(el, name);
  if (v === undefined) {
    return undefined;
  }
  if (v === "1" || v === "true") {
    return true;
  }
  if (v === "0" || v === "false") {
    return false;
  }
  return undefined;
}

function findChildNs(el: XmlElement, localName: string): XmlElement | undefined {
  // Match either w:localName or just localName
  return findChild(el, `w:${localName}`) ?? findChild(el, localName);
}

function findChildrenNs(el: XmlElement, localName: string): XmlElement[] {
  const a = findChildren(el, `w:${localName}`);
  return a.length > 0 ? a : findChildren(el, localName);
}

/** Check for a boolean toggle element (present = true, w:val="0" or "false" = false). */
function boolToggle(parent: XmlElement, name: string): boolean | undefined {
  const el = findChildNs(parent, name);
  if (!el) {
    return undefined;
  }
  const v = attrVal(el, "val");
  if (v === "0" || v === "false") {
    return false;
  }
  return true;
}

/** Escape special XML characters in text content. */
function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Serialize an XmlElement back to an XML string (for opaque preservation). */
function serializeElement(el: XmlElement): string {
  let s = `<${el.name}`;
  for (const [k, v] of Object.entries(el.attributes)) {
    s += ` ${k}="${escapeXml(v)}"`;
  }
  if (el.children.length === 0) {
    return s + "/>";
  }
  s += ">";
  for (const child of el.children) {
    if (child.type === "element") {
      s += serializeElement(child);
    } else if (child.type === "text") {
      s += escapeXml(child.value);
    }
  }
  s += `</${el.name}>`;
  return s;
}

/** Extract all r:xxx attribute values (relationship IDs) from an element tree. */
function collectRIds(el: XmlElement, out: Set<string>): void {
  for (const [k, v] of Object.entries(el.attributes)) {
    if (k.startsWith("r:") || k === "r:id" || k === "r:embed" || k === "r:link") {
      out.add(v);
    }
  }
  for (const child of el.children) {
    if (child.type === "element") {
      collectRIds(child, out);
    }
  }
}

/** Get the .rels path for a given part path. */
function getPartRelsPath(partPath: string): string {
  const lastSlash = partPath.lastIndexOf("/");
  const dir = lastSlash >= 0 ? partPath.substring(0, lastSlash) : "";
  const name = lastSlash >= 0 ? partPath.substring(lastSlash + 1) : partPath;
  return dir ? `${dir}/_rels/${name}.rels` : `_rels/${name}.rels`;
}

/**
 * Resolve a relationship target path to an absolute package-root path.
 *
 * - Leading "/" → package root absolute
 * - "../" / "./" → resolved relative to the source part's directory
 * - Plain paths → resolved relative to the source part's directory
 */
function resolvePartPath(sourcePart: string, target: string): string {
  if (!target) {
    return "";
  }
  if (target.startsWith("/")) {
    return target.slice(1);
  }
  const lastSlash = sourcePart.lastIndexOf("/");
  const baseDir = lastSlash >= 0 ? sourcePart.substring(0, lastSlash).split("/") : [];
  const segs = target.split("/");
  for (const seg of segs) {
    if (seg === "..") {
      baseDir.pop();
    } else if (seg !== "." && seg !== "") {
      baseDir.push(seg);
    }
  }
  return baseDir.join("/");
}

/** Parse footnote/endnote properties element. */
function parseNoteProperties(el: XmlElement): any {
  const props: any = {};
  const numFmtEl = findChildNs(el, "numFmt");
  if (numFmtEl) {
    props.numFmt = attrVal(numFmtEl, "val");
  }
  const numStartEl = findChildNs(el, "numStart");
  if (numStartEl) {
    props.numStart = attrInt(numStartEl, "val");
  }
  const numRestartEl = findChildNs(el, "numRestart");
  if (numRestartEl) {
    props.numRestart = attrVal(numRestartEl, "val");
  }
  const posEl = findChildNs(el, "pos");
  if (posEl) {
    props.position = attrVal(posEl, "val");
  }
  return Object.keys(props).length > 0 ? props : undefined;
}

/** Parse w:ffData element into a FormField. */
function parseFfData(el: XmlElement): FormField | undefined {
  const nameEl = findChildNs(el, "name");
  const name = nameEl ? attrVal(nameEl, "val") : undefined;
  const enabledEl = findChildNs(el, "enabled");
  const enabled = enabledEl ? attrVal(enabledEl, "val") !== "0" : undefined;
  const helpTextEl = findChildNs(el, "helpText");
  const helpText = helpTextEl ? attrVal(helpTextEl, "val") : undefined;
  const statusTextEl = findChildNs(el, "statusText");
  const statusText = statusTextEl ? attrVal(statusTextEl, "val") : undefined;

  // Text input
  const textInputEl = findChildNs(el, "textInput");
  if (textInputEl) {
    const defEl = findChildNs(textInputEl, "default");
    const maxLenEl = findChildNs(textInputEl, "maxLength");
    const fmtEl = findChildNs(textInputEl, "format");
    return {
      type: "text",
      name,
      default: defEl ? attrVal(defEl, "val") : undefined,
      maxLength: maxLenEl ? attrInt(maxLenEl, "val") : undefined,
      format: fmtEl ? attrVal(fmtEl, "val") : undefined,
      helpText,
      statusText,
      enabled
    };
  }

  // CheckBox
  const cbEl = findChildNs(el, "checkBox");
  if (cbEl) {
    const checkedEl = findChildNs(cbEl, "checked");
    const defEl = findChildNs(cbEl, "default");
    const sizeEl = findChildNs(cbEl, "size");
    return {
      type: "checkBox",
      name,
      checked: checkedEl ? attrVal(checkedEl, "val") !== "0" : undefined,
      default: defEl ? attrVal(defEl, "val") !== "0" : undefined,
      size: sizeEl ? attrInt(sizeEl, "val") : undefined
    };
  }

  // Drop-down list
  const ddlEl = findChildNs(el, "ddList");
  if (ddlEl) {
    const defEl = findChildNs(ddlEl, "default");
    const entries: string[] = [];
    for (const le of findChildrenNs(ddlEl, "listEntry")) {
      const v = attrVal(le, "val");
      if (v !== undefined) {
        entries.push(v);
      }
    }
    return {
      type: "dropDown",
      name,
      entries: entries.length > 0 ? entries : undefined,
      default: defEl ? attrInt(defEl, "val") : undefined,
      helpText,
      statusText,
      enabled
    };
  }

  return undefined;
}

// =============================================================================
// Run Properties Parser
// =============================================================================

function parseRunProperties(rPrEl: XmlElement): RunProperties {
  const rPr: any = {};

  const rStyleEl = findChildNs(rPrEl, "rStyle");
  if (rStyleEl) {
    rPr.style = attrVal(rStyleEl, "val");
  }

  const fontsEl = findChildNs(rPrEl, "rFonts");
  if (fontsEl) {
    const f: FontSpec = {};
    const ascii = attrVal(fontsEl, "ascii");
    const hAnsi = attrVal(fontsEl, "hAnsi");
    const eastAsia = attrVal(fontsEl, "eastAsia");
    const cs = attrVal(fontsEl, "cs");
    const hint = attrVal(fontsEl, "hint");
    if (ascii) {
      (f as any).ascii = ascii;
    }
    if (hAnsi) {
      (f as any).hAnsi = hAnsi;
    }
    if (eastAsia) {
      (f as any).eastAsia = eastAsia;
    }
    if (cs) {
      (f as any).cs = cs;
    }
    if (hint) {
      (f as any).hint = hint;
    }
    const asciiTheme = attrVal(fontsEl, "asciiTheme");
    if (asciiTheme) {
      (f as any).asciiTheme = asciiTheme;
    }
    const hAnsiTheme = attrVal(fontsEl, "hAnsiTheme");
    if (hAnsiTheme) {
      (f as any).hAnsiTheme = hAnsiTheme;
    }
    const eastAsiaTheme = attrVal(fontsEl, "eastAsiaTheme");
    if (eastAsiaTheme) {
      (f as any).eastAsiaTheme = eastAsiaTheme;
    }
    const cstheme = attrVal(fontsEl, "cstheme");
    if (cstheme) {
      (f as any).cstheme = cstheme;
    }
    rPr.font = f;
  }

  // Boolean toggles: true-only (element presence = true; absence = undefined)
  // Format: [elementName, propertyKey]
  const RUN_ONCE_TOGGLES: ReadonlyArray<[string, keyof RunProperties]> = [
    ["strike", "strike"],
    ["dstrike", "doubleStrike"],
    ["caps", "caps"],
    ["smallCaps", "smallCaps"],
    ["vanish", "vanish"],
    ["emboss", "emboss"],
    ["imprint", "imprint"],
    ["noProof", "noProof"],
    ["specVanish", "specVanish"],
    ["outline", "outline"],
    ["shadow", "shadow"],
    ["cs", "complexScript"],
    ["oMath", "math"],
    ["webHidden", "webHidden"]
  ];
  for (const [tag, key] of RUN_ONCE_TOGGLES) {
    if (findChildNs(rPrEl, tag) && boolToggle(rPrEl, tag) !== false) {
      rPr[key] = true;
    }
  }

  // Boolean toggles: tri-state (can be explicitly true or false)
  const RUN_TRISTATE_TOGGLES: ReadonlyArray<[string, keyof RunProperties]> = [
    ["b", "bold"],
    ["bCs", "boldCs"],
    ["i", "italic"],
    ["iCs", "italicCs"],
    ["snapToGrid", "snapToGrid"],
    ["rtl", "rightToLeft"]
  ];
  for (const [tag, key] of RUN_TRISTATE_TOGGLES) {
    const v = boolToggle(rPrEl, tag);
    if (v !== undefined) {
      rPr[key] = v;
    }
  }

  // fitText
  const fitTextEl = findChildNs(rPrEl, "fitText");
  if (fitTextEl) {
    const val = attrInt(fitTextEl, "val");
    if (val !== undefined) {
      const fitText: any = { val };
      const id = attrInt(fitTextEl, "id");
      if (id !== undefined) {
        fitText.id = id;
      }
      rPr.fitText = fitText;
    }
  }

  const colorEl = findChildNs(rPrEl, "color");
  if (colorEl) {
    const val = attrVal(colorEl, "val");
    const themeColor = attrVal(colorEl, "themeColor");
    if (themeColor) {
      const spec: ColorSpec = { val, themeColor } as any;
      const themeTint = attrVal(colorEl, "themeTint");
      const themeShade = attrVal(colorEl, "themeShade");
      if (themeTint) {
        (spec as any).themeTint = themeTint;
      }
      if (themeShade) {
        (spec as any).themeShade = themeShade;
      }
      rPr.color = spec;
    } else {
      rPr.color = val;
    }
  }

  const szEl = findChildNs(rPrEl, "sz");
  if (szEl) {
    rPr.size = attrInt(szEl, "val");
  }

  const szCsEl = findChildNs(rPrEl, "szCs");
  if (szCsEl) {
    rPr.sizeCs = attrInt(szCsEl, "val");
  }

  const uEl = findChildNs(rPrEl, "u");
  if (uEl) {
    const uStyle = attrVal(uEl, "val") ?? "single";
    const uColor = attrVal(uEl, "color");
    if (uColor) {
      rPr.underline = { style: uStyle, color: uColor };
    } else {
      rPr.underline = uStyle;
    }
  }

  const highlightEl = findChildNs(rPrEl, "highlight");
  if (highlightEl) {
    rPr.highlight = attrVal(highlightEl, "val");
  }

  const vertAlignEl = findChildNs(rPrEl, "vertAlign");
  if (vertAlignEl) {
    rPr.vertAlign = attrVal(vertAlignEl, "val");
  }

  const spacingEl = findChildNs(rPrEl, "spacing");
  if (spacingEl) {
    rPr.spacing = attrInt(spacingEl, "val");
  }

  const shdEl = findChildNs(rPrEl, "shd");
  if (shdEl) {
    rPr.shading = parseShading(shdEl);
  }

  const langEl = findChildNs(rPrEl, "lang");
  if (langEl) {
    rPr.language = {
      val: attrVal(langEl, "val"),
      eastAsia: attrVal(langEl, "eastAsia"),
      bidi: attrVal(langEl, "bidi")
    };
  }

  // New valued properties
  const kernEl = findChildNs(rPrEl, "kern");
  if (kernEl) {
    rPr.kern = attrInt(kernEl, "val");
  }

  const positionEl = findChildNs(rPrEl, "position");
  if (positionEl) {
    rPr.position = attrInt(positionEl, "val");
  }

  const wEl = findChildNs(rPrEl, "w");
  if (wEl) {
    rPr.scale = attrInt(wEl, "val");
  }

  const effectEl = findChildNs(rPrEl, "effect");
  if (effectEl) {
    rPr.effect = attrVal(effectEl, "val");
  }

  const emEl = findChildNs(rPrEl, "em");
  if (emEl) {
    rPr.emphasisMark = attrVal(emEl, "val");
  }

  const bdrEl = findChildNs(rPrEl, "bdr");
  if (bdrEl) {
    rPr.border = parseBorder(bdrEl);
  }

  // rPrChange (track changes for run properties)
  const rPrChangeEl = findChildNs(rPrEl, "rPrChange");
  if (rPrChangeEl) {
    const rev = parseRevisionInfo(rPrChangeEl);
    if (rev) {
      const prevRPrEl = findChildNs(rPrChangeEl, "rPr");
      const change: RunPropertyChange = {
        revision: rev,
        previousProperties: prevRPrEl ? parseRunProperties(prevRPrEl) : undefined
      };
      rPr.propertyChange = change;
    }
  }

  return rPr;
}

function parseShading(el: XmlElement): Shading {
  return {
    pattern: attrVal(el, "val") as Shading["pattern"],
    color: attrVal(el, "color"),
    fill: attrVal(el, "fill") ?? "auto"
  };
}

function parseBorder(el: XmlElement): Border {
  const b: any = {
    style: (attrVal(el, "val") ?? "single") as any,
    size: attrInt(el, "sz"),
    space: attrInt(el, "space"),
    color: attrVal(el, "color")
  };
  const tc = attrVal(el, "themeColor");
  if (tc) {
    b.themeColor = tc;
  }
  const shadow = attrVal(el, "shadow");
  if (shadow === "1" || shadow === "true") {
    b.shadow = true;
  }
  const frame = attrVal(el, "frame");
  if (frame === "1" || frame === "true") {
    b.frame = true;
  }
  const art = attrVal(el, "art");
  if (art) {
    b.art = art;
  }
  return b;
}

function parseTableWidth(el: XmlElement): TableWidth {
  return {
    value: parseInt(el.attributes["w:w"] ?? el.attributes["w"] ?? "0", 10),
    type: (el.attributes["w:type"] ?? el.attributes["type"] ?? "dxa") as any
  };
}

// =============================================================================
// Paragraph Properties Parser
// =============================================================================

function parseParagraphProperties(pPrEl: XmlElement): ParagraphProperties {
  const pPr: any = {};

  const pStyleEl = findChildNs(pPrEl, "pStyle");
  if (pStyleEl) {
    pPr.style = attrVal(pStyleEl, "val");
  }

  const jcEl = findChildNs(pPrEl, "jc");
  if (jcEl) {
    pPr.alignment = attrVal(jcEl, "val");
  }

  if (findChildNs(pPrEl, "keepNext")) {
    pPr.keepNext = true;
  }
  if (findChildNs(pPrEl, "keepLines")) {
    pPr.keepLines = true;
  }
  if (findChildNs(pPrEl, "pageBreakBefore")) {
    pPr.pageBreakBefore = true;
  }
  if (findChildNs(pPrEl, "bidi")) {
    pPr.bidi = true;
  }

  // New boolean toggles
  const ctxSp = boolToggle(pPrEl, "contextualSpacing");
  if (ctxSp !== undefined) {
    pPr.contextualSpacing = ctxSp;
  }
  const suppLn = boolToggle(pPrEl, "suppressLineNumbers");
  if (suppLn !== undefined) {
    pPr.suppressLineNumbers = suppLn;
  }
  const suppHyph = boolToggle(pPrEl, "suppressAutoHyphens");
  if (suppHyph !== undefined) {
    pPr.suppressAutoHyphens = suppHyph;
  }
  const mirr = boolToggle(pPrEl, "mirrorIndents");
  if (mirr !== undefined) {
    pPr.mirrorIndents = mirr;
  }
  const wc = boolToggle(pPrEl, "widowControl");
  if (wc !== undefined) {
    pPr.widowControl = wc;
  }
  const ww = boolToggle(pPrEl, "wordWrap");
  if (ww !== undefined) {
    pPr.wordWrap = ww;
  }
  const stg = boolToggle(pPrEl, "snapToGrid");
  if (stg !== undefined) {
    pPr.snapToGrid = stg;
  }
  const ofp = boolToggle(pPrEl, "overflowPunct");
  if (ofp !== undefined) {
    pPr.overflowPunctuation = ofp;
  }
  const topLinePunct = boolToggle(pPrEl, "topLinePunct");
  if (topLinePunct !== undefined) {
    pPr.topLinePunctuation = topLinePunct;
  }
  const kinsoku = boolToggle(pPrEl, "kinsoku");
  if (kinsoku !== undefined) {
    pPr.kinsoku = kinsoku;
  }
  const asd = boolToggle(pPrEl, "autoSpaceDE");
  if (asd !== undefined) {
    pPr.autoSpaceEastAsianText = asd;
  }
  const asdn = boolToggle(pPrEl, "autoSpaceDN");
  if (asdn !== undefined) {
    pPr.autoSpaceEastAsianDigit = asdn;
  }

  const textAlignEl = findChildNs(pPrEl, "textAlignment");
  if (textAlignEl) {
    pPr.textAlignment = attrVal(textAlignEl, "val");
  }

  const outlineLvlEl = findChildNs(pPrEl, "outlineLvl");
  if (outlineLvlEl) {
    pPr.outlineLevel = attrInt(outlineLvlEl, "val");
  }

  const textDirEl = findChildNs(pPrEl, "textDirection");
  if (textDirEl) {
    pPr.textDirection = attrVal(textDirEl, "val");
  }

  // Paragraph frame
  const framePrEl = findChildNs(pPrEl, "framePr");
  if (framePrEl) {
    const frame: ParagraphFrame = {} as any;
    const f: any = frame;
    const dropCap = attrVal(framePrEl, "dropCap");
    if (dropCap) {
      f.dropCap = dropCap;
    }
    const lines = attrInt(framePrEl, "lines");
    if (lines !== undefined) {
      f.lines = lines;
    }
    const fw = attrInt(framePrEl, "w");
    if (fw !== undefined) {
      f.width = fw;
    }
    const fh = attrInt(framePrEl, "h");
    if (fh !== undefined) {
      f.height = fh;
    }
    const hSpace = attrInt(framePrEl, "hSpace");
    if (hSpace !== undefined) {
      f.hSpace = hSpace;
    }
    const vSpace = attrInt(framePrEl, "vSpace");
    if (vSpace !== undefined) {
      f.vSpace = vSpace;
    }
    const wrap = attrVal(framePrEl, "wrap");
    if (wrap) {
      f.wrap = wrap;
    }
    const hAnchor = attrVal(framePrEl, "hAnchor");
    if (hAnchor) {
      f.hAnchor = hAnchor;
    }
    const vAnchor = attrVal(framePrEl, "vAnchor");
    if (vAnchor) {
      f.vAnchor = vAnchor;
    }
    const x = attrInt(framePrEl, "x");
    if (x !== undefined) {
      f.x = x;
    }
    const xAlign = attrVal(framePrEl, "xAlign");
    if (xAlign) {
      f.xAlign = xAlign;
    }
    const y = attrInt(framePrEl, "y");
    if (y !== undefined) {
      f.y = y;
    }
    const yAlign = attrVal(framePrEl, "yAlign");
    if (yAlign) {
      f.yAlign = yAlign;
    }
    pPr.frame = frame;
  }

  // Thematic break: check for bottom border with special pattern
  const pBdrEl = findChildNs(pPrEl, "pBdr");
  if (pBdrEl) {
    const borders: any = {};
    for (const side of ["top", "bottom", "left", "right", "between", "bar"] as const) {
      const sideEl = findChildNs(pBdrEl, side);
      if (sideEl) {
        borders[side] = parseBorder(sideEl);
      }
    }
    pPr.borders = borders;
  }

  const spacingEl = findChildNs(pPrEl, "spacing");
  if (spacingEl) {
    const spacing: any = {};
    const before = attrInt(spacingEl, "before");
    const after = attrInt(spacingEl, "after");
    const line = attrInt(spacingEl, "line");
    const lineRule = attrVal(spacingEl, "lineRule");
    if (before !== undefined) {
      spacing.before = before;
    }
    if (after !== undefined) {
      spacing.after = after;
    }
    if (line !== undefined) {
      spacing.line = line;
      // Per ECMA-376, w:lineRule defaults to "auto" when line is set
      spacing.lineRule = (lineRule ?? "auto") as any;
    } else if (lineRule) {
      spacing.lineRule = lineRule;
    }
    const beforeAuto = attrVal(spacingEl, "beforeAutospacing");
    if (beforeAuto === "1" || beforeAuto === "true") {
      spacing.beforeAutoSpacing = true;
    }
    const afterAuto = attrVal(spacingEl, "afterAutospacing");
    if (afterAuto === "1" || afterAuto === "true") {
      spacing.afterAutoSpacing = true;
    }
    pPr.spacing = spacing;
  }

  const indEl = findChildNs(pPrEl, "ind");
  if (indEl) {
    const indent: any = {};
    const left = attrInt(indEl, "left");
    const right = attrInt(indEl, "right");
    const hanging = attrInt(indEl, "hanging");
    const firstLine = attrInt(indEl, "firstLine");
    const start = attrInt(indEl, "start");
    const end = attrInt(indEl, "end");
    if (left !== undefined) {
      indent.left = left;
    }
    if (right !== undefined) {
      indent.right = right;
    }
    if (hanging !== undefined) {
      indent.hanging = hanging;
    }
    if (firstLine !== undefined) {
      indent.firstLine = firstLine;
    }
    if (start !== undefined) {
      indent.start = start;
    }
    if (end !== undefined) {
      indent.end = end;
    }
    pPr.indent = indent;
  }

  const numPrEl = findChildNs(pPrEl, "numPr");
  if (numPrEl) {
    const ilvlEl = findChildNs(numPrEl, "ilvl");
    const numIdEl = findChildNs(numPrEl, "numId");
    // Per OOXML schema, numId is required but ilvl is optional (defaults to 0).
    if (numIdEl) {
      pPr.numbering = {
        level: ilvlEl ? (attrInt(ilvlEl, "val") ?? 0) : 0,
        numId: attrInt(numIdEl, "val") ?? 0
      };
    }
  }

  const tabsEl = findChildNs(pPrEl, "tabs");
  if (tabsEl) {
    const tabs: TabStop[] = [];
    for (const tabEl of findChildrenNs(tabsEl, "tab")) {
      tabs.push({
        type: (attrVal(tabEl, "val") ?? "left") as any,
        position: attrInt(tabEl, "pos") ?? 0,
        leader: attrVal(tabEl, "leader") as any
      });
    }
    if (tabs.length > 0) {
      pPr.tabs = tabs;
    }
  }

  const shdEl = findChildNs(pPrEl, "shd");
  if (shdEl) {
    pPr.shading = parseShading(shdEl);
  }

  const rPrEl = findChildNs(pPrEl, "rPr");
  if (rPrEl) {
    pPr.markRunProperties = parseRunProperties(rPrEl);
  }

  const sectPrEl = findChildNs(pPrEl, "sectPr");
  if (sectPrEl) {
    pPr.sectionProperties = parseSectionProperties(sectPrEl);
  }

  // Conditional formatting style mask
  const cnfStyleEl = findChildNs(pPrEl, "cnfStyle");
  if (cnfStyleEl) {
    pPr.cnfStyle = attrVal(cnfStyleEl, "val");
  }

  // Paragraph property change
  const pPrChangeEl = findChildNs(pPrEl, "pPrChange");
  if (pPrChangeEl) {
    const rev = parseRevisionInfo(pPrChangeEl);
    if (rev) {
      const prevPPrEl = findChildNs(pPrChangeEl, "pPr");
      pPr.propertyChange = {
        revision: rev,
        previousProperties: prevPPrEl ? parseParagraphProperties(prevPPrEl) : undefined
      };
    }
  }

  // Paragraph mark insertion/deletion (w:pPr > w:rPr > w:ins/w:del)
  const rPrInPPr = findChildNs(pPrEl, "rPr");
  if (rPrInPPr) {
    const insEl = findChildNs(rPrInPPr, "ins");
    if (insEl) {
      const rev = parseRevisionInfo(insEl);
      if (rev) {
        pPr.paragraphInsertion = rev;
      }
    }
    const delEl = findChildNs(rPrInPPr, "del");
    if (delEl) {
      const rev = parseRevisionInfo(delEl);
      if (rev) {
        pPr.paragraphDeletion = rev;
      }
    }
  }

  return pPr;
}

// =============================================================================
// Revision Info Parser
// =============================================================================

function parseRevisionInfo(el: XmlElement): RevisionInfo | undefined {
  const author = attrVal(el, "author");
  const id = attrInt(el, "id");
  if (author === undefined || id === undefined) {
    return undefined;
  }
  return {
    author,
    id,
    date: attrVal(el, "date")
  };
}

// =============================================================================
// Run Content Parser
// =============================================================================

function parseRunContent(el: XmlElement): RunContent[] {
  const content: RunContent[] = [];
  for (const child of el.children) {
    if (child.type !== "element") {
      continue;
    }
    const name = child.name.replace(/^w:/, "");
    switch (name) {
      case "t":
        content.push({ type: "text", text: textContent(child) });
        break;
      case "br": {
        const brType = attrVal(child, "type");
        content.push({ type: "break", breakType: brType as any });
        break;
      }
      case "tab":
        content.push({ type: "tab" });
        break;
      case "ptab": {
        const alignment = attrVal(child, "alignment") ?? "left";
        const relativeTo = attrVal(child, "relativeTo") ?? "margin";
        const leader = attrVal(child, "leader");
        const ptab: any = {
          type: "ptab",
          alignment,
          relativeTo
        };
        if (leader) {
          ptab.leader = leader;
        }
        content.push(ptab);
        break;
      }
      case "ruby": {
        const ruby: any = { type: "ruby" };
        const rubyPrEl = findChildNs(child, "rubyPr");
        if (rubyPrEl) {
          const props: any = {};
          const alignEl = findChildNs(rubyPrEl, "rubyAlign");
          if (alignEl) {
            props.align = attrVal(alignEl, "val");
          }
          const hpsEl = findChildNs(rubyPrEl, "hps");
          if (hpsEl) {
            props.fontSize = attrInt(hpsEl, "val");
          }
          const hpsRaiseEl = findChildNs(rubyPrEl, "hpsRaise");
          if (hpsRaiseEl) {
            props.raise = attrInt(hpsRaiseEl, "val");
          }
          const hpsBaseTextEl = findChildNs(rubyPrEl, "hpsBaseText");
          if (hpsBaseTextEl) {
            props.baseFontSize = attrInt(hpsBaseTextEl, "val");
          }
          const lidEl = findChildNs(rubyPrEl, "lid");
          if (lidEl) {
            props.language = attrVal(lidEl, "val");
          }
          if (Object.keys(props).length > 0) {
            ruby.properties = props;
          }
        }
        // Parse w:rt (ruby text)
        const rtEl = findChildNs(child, "rt");
        const rubyText: Run[] = [];
        if (rtEl) {
          for (const rtChild of rtEl.children) {
            if (rtChild.type === "element" && rtChild.name.replace(/^w:/, "") === "r") {
              rubyText.push(parseRun(rtChild));
            }
          }
        }
        ruby.rubyText = rubyText;
        // Parse w:rubyBase
        const baseEl = findChildNs(child, "rubyBase");
        const baseText: Run[] = [];
        if (baseEl) {
          for (const bChild of baseEl.children) {
            if (bChild.type === "element" && bChild.name.replace(/^w:/, "") === "r") {
              baseText.push(parseRun(bChild));
            }
          }
        }
        ruby.baseText = baseText;
        content.push(ruby);
        break;
      }
      case "sym":
        content.push({
          type: "symbol",
          font: attrVal(child, "font") ?? "",
          char: attrVal(child, "char") ?? ""
        });
        break;
      case "footnoteReference": {
        const fr: any = { type: "footnoteRef", id: attrInt(child, "id") ?? 0 };
        const cmf = attrVal(child, "customMarkFollows");
        if (cmf === "1" || cmf === "true") {
          fr.customMarkFollows = true;
        }
        content.push(fr);
        break;
      }
      case "endnoteReference": {
        const er: any = { type: "endnoteRef", id: attrInt(child, "id") ?? 0 };
        const cmf = attrVal(child, "customMarkFollows");
        if (cmf === "1" || cmf === "true") {
          er.customMarkFollows = true;
        }
        content.push(er);
        break;
      }
      case "drawing":
        parseDrawingContent(child, content);
        break;
      case "cr":
        content.push({ type: "carriageReturn" });
        break;
      case "noBreakHyphen":
        content.push({ type: "noBreakHyphen" });
        break;
      case "softHyphen":
        content.push({ type: "softHyphen" });
        break;
      case "lastRenderedPageBreak":
        content.push({ type: "lastRenderedPageBreak" });
        break;
      case "annotationRef":
        content.push({ type: "annotationReference", id: attrInt(child, "id") ?? 0 });
        break;
      case "commentReference":
        // This is annotationReference for comments inside runs
        content.push({ type: "annotationReference", id: attrInt(child, "id") ?? 0 });
        break;
    }
  }
  return content;
}

function parseDrawingContent(drawingEl: XmlElement, content: RunContent[]): void {
  // Look for wp:inline
  const inlineEl = findChild(drawingEl, "wp:inline");
  if (inlineEl) {
    const extentEl = findChild(inlineEl, "wp:extent");
    const docPrEl = findChild(inlineEl, "wp:docPr");
    const graphicEl = findChild(inlineEl, "a:graphic");
    const graphicDataEl = graphicEl ? findChild(graphicEl, "a:graphicData") : undefined;
    const picEl = graphicDataEl ? findChild(graphicDataEl, "pic:pic") : undefined;
    const blipFillEl = picEl ? findChild(picEl, "pic:blipFill") : undefined;
    const blipEl = blipFillEl ? findChild(blipFillEl, "a:blip") : undefined;

    const rId = blipEl?.attributes["r:embed"] ?? "";
    const cx = parseInt(extentEl?.attributes["cx"] ?? "0", 10);
    const cy = parseInt(extentEl?.attributes["cy"] ?? "0", 10);

    const img: any = {
      type: "image",
      rId,
      width: cx,
      height: cy,
      altText: docPrEl?.attributes["descr"],
      name: docPrEl?.attributes["name"],
      drawingId: docPrEl ? parseInt(docPrEl.attributes["id"] ?? "1", 10) : undefined
    };

    // Parse xfrm for rotation/flip
    const spPrEl = picEl ? findChild(picEl, "pic:spPr") : undefined;
    if (spPrEl) {
      const xfrmEl = findChild(spPrEl, "a:xfrm");
      if (xfrmEl) {
        const rot = xfrmEl.attributes["rot"];
        if (rot !== undefined && rot !== "") {
          img.rotation = parseInt(rot, 10);
        }
        if (xfrmEl.attributes["flipH"] === "1") {
          img.flipHorizontal = true;
        }
        if (xfrmEl.attributes["flipV"] === "1") {
          img.flipVertical = true;
        }
      }
      // Outline
      const lnEl = findChild(spPrEl, "a:ln");
      if (lnEl) {
        const outline: any = {};
        const w = lnEl.attributes["w"];
        if (w) {
          outline.width = parseInt(w, 10);
        }
        const sfEl = findChild(lnEl, "a:solidFill");
        const srgbEl = sfEl ? findChild(sfEl, "a:srgbClr") : undefined;
        if (srgbEl) {
          outline.color = srgbEl.attributes["val"];
        }
        img.outline = outline;
      }
    }

    // SVG blip in a:extLst
    if (blipEl) {
      const extLst = findChild(blipEl, "a:extLst");
      if (extLst) {
        for (const ext of findChildren(extLst, "a:ext")) {
          const svgBlip = findChild(ext, "asvg:svgBlip") ?? findChildNs(ext, "svgBlip");
          if (svgBlip) {
            const svgEmbed = svgBlip.attributes["r:embed"];
            if (svgEmbed) {
              img.svgRId = svgEmbed;
            }
          }
        }
      }
    }

    content.push(img);
  }
}

// =============================================================================
// Floating Image Parser
// =============================================================================

function parseFloatingImage(anchorEl: XmlElement): FloatingImage | undefined {
  const docPrEl = findChild(anchorEl, "wp:docPr");
  const extentEl = findChild(anchorEl, "wp:extent");
  const graphicEl = findChild(anchorEl, "a:graphic");
  const graphicDataEl = graphicEl ? findChild(graphicEl, "a:graphicData") : undefined;
  const picEl = graphicDataEl ? findChild(graphicDataEl, "pic:pic") : undefined;
  const blipFillEl = picEl ? findChild(picEl, "pic:blipFill") : undefined;
  const blipEl = blipFillEl ? findChild(blipFillEl, "a:blip") : undefined;

  const rId = blipEl?.attributes["r:embed"];
  if (!rId) {
    return undefined;
  }

  const cx = parseInt(extentEl?.attributes["cx"] ?? "0", 10);
  const cy = parseInt(extentEl?.attributes["cy"] ?? "0", 10);

  const img: any = {
    type: "floatingImage",
    rId,
    width: cx,
    height: cy,
    altText: docPrEl?.attributes["descr"],
    name: docPrEl?.attributes["name"],
    drawingId: docPrEl ? parseInt(docPrEl.attributes["id"] ?? "1", 10) : undefined
  };

  // Attributes
  if (anchorEl.attributes["behindDoc"] === "1") {
    img.behindDoc = true;
  }
  if (anchorEl.attributes["locked"] === "1") {
    img.lockAnchor = true;
  }
  if (anchorEl.attributes["layoutInCell"] === "0") {
    img.layoutInCell = false;
  }
  if (anchorEl.attributes["allowOverlap"] === "0") {
    img.allowOverlap = false;
  }
  const rh = anchorEl.attributes["relativeHeight"];
  if (rh) {
    img.relativeHeight = parseInt(rh, 10);
  }
  // Dist*
  const distT = anchorEl.attributes["distT"];
  if (distT) {
    img.distT = parseInt(distT, 10);
  }
  const distB = anchorEl.attributes["distB"];
  if (distB) {
    img.distB = parseInt(distB, 10);
  }
  const distL = anchorEl.attributes["distL"];
  if (distL) {
    img.distL = parseInt(distL, 10);
  }
  const distR = anchorEl.attributes["distR"];
  if (distR) {
    img.distR = parseInt(distR, 10);
  }

  // Simple positioning
  if (anchorEl.attributes["simplePos"] === "1") {
    const sposEl = findChild(anchorEl, "wp:simplePos");
    if (sposEl) {
      const x = parseInt(sposEl.attributes["x"] ?? "0", 10);
      const y = parseInt(sposEl.attributes["y"] ?? "0", 10);
      img.simplePos = { x, y };
    }
  }

  // Horizontal position
  const hPosEl = findChild(anchorEl, "wp:positionH");
  if (hPosEl) {
    const h: any = { relativeTo: hPosEl.attributes["relativeFrom"] };
    const offsetEl = findChild(hPosEl, "wp:posOffset");
    if (offsetEl) {
      h.offset = parseInt(textContent(offsetEl), 10);
    }
    const alignEl = findChild(hPosEl, "wp:align");
    if (alignEl) {
      h.align = textContent(alignEl);
    }
    img.horizontalPosition = h;
  }

  // Vertical position
  const vPosEl = findChild(anchorEl, "wp:positionV");
  if (vPosEl) {
    const v: any = { relativeTo: vPosEl.attributes["relativeFrom"] };
    const offsetEl = findChild(vPosEl, "wp:posOffset");
    if (offsetEl) {
      v.offset = parseInt(textContent(offsetEl), 10);
    }
    const alignEl = findChild(vPosEl, "wp:align");
    if (alignEl) {
      v.align = textContent(alignEl);
    }
    img.verticalPosition = v;
  }

  // Wrap
  for (const wrapChild of anchorEl.children) {
    if (wrapChild.type !== "element") {
      continue;
    }
    const wn = wrapChild.name;
    if (wn === "wp:wrapSquare") {
      img.wrap = { style: "square", side: wrapChild.attributes["wrapText"] };
    } else if (wn === "wp:wrapTight") {
      img.wrap = { style: "tight", side: wrapChild.attributes["wrapText"] };
    } else if (wn === "wp:wrapThrough") {
      img.wrap = { style: "through", side: wrapChild.attributes["wrapText"] };
    } else if (wn === "wp:wrapTopAndBottom") {
      img.wrap = { style: "topAndBottom" };
    } else if (wn === "wp:wrapNone") {
      img.wrap = { style: "none" };
    }
    if (img.wrap) {
      // Parse wrap margins
      const distT = anchorEl.attributes["distT"];
      const distB = anchorEl.attributes["distB"];
      const distL = anchorEl.attributes["distL"];
      const distR = anchorEl.attributes["distR"];
      if (distT || distB || distL || distR) {
        const margins: any = {};
        if (distT) {
          margins.top = parseInt(distT, 10);
        }
        if (distB) {
          margins.bottom = parseInt(distB, 10);
        }
        if (distL) {
          margins.left = parseInt(distL, 10);
        }
        if (distR) {
          margins.right = parseInt(distR, 10);
        }
        img.wrap.margins = margins;
      }
      break;
    }
  }

  // Rotation/flip from spPr
  const spPrEl = picEl ? findChild(picEl, "pic:spPr") : undefined;
  if (spPrEl) {
    const xfrmEl = findChild(spPrEl, "a:xfrm");
    if (xfrmEl) {
      const rot = xfrmEl.attributes["rot"];
      if (rot !== undefined && rot !== "") {
        img.rotation = parseInt(rot, 10);
      }
      if (xfrmEl.attributes["flipH"] === "1") {
        img.flipHorizontal = true;
      }
      if (xfrmEl.attributes["flipV"] === "1") {
        img.flipVertical = true;
      }
    }
    const lnEl = findChild(spPrEl, "a:ln");
    if (lnEl) {
      const outline: any = {};
      const w = lnEl.attributes["w"];
      if (w) {
        outline.width = parseInt(w, 10);
      }
      const sfEl = findChild(lnEl, "a:solidFill");
      const srgbEl = sfEl ? findChild(sfEl, "a:srgbClr") : undefined;
      if (srgbEl) {
        outline.color = srgbEl.attributes["val"];
      }
      img.outline = outline;
    }
  }

  // SVG blip in a:extLst
  if (blipEl) {
    const extLst = findChild(blipEl, "a:extLst");
    if (extLst) {
      for (const ext of findChildren(extLst, "a:ext")) {
        const svgBlip = findChild(ext, "asvg:svgBlip") ?? findChildNs(ext, "svgBlip");
        if (svgBlip) {
          const svgEmbed = svgBlip.attributes["r:embed"];
          if (svgEmbed) {
            img.svgRId = svgEmbed;
          }
        }
      }
    }
  }

  // Source rectangle (crop)
  if (blipFillEl) {
    const srcRectEl = findChild(blipFillEl, "a:srcRect");
    if (srcRectEl) {
      const sr: any = {};
      const lAttr = srcRectEl.attributes["l"];
      const tAttr = srcRectEl.attributes["t"];
      const rAttr = srcRectEl.attributes["r"];
      const bAttr = srcRectEl.attributes["b"];
      if (lAttr !== undefined) {
        sr.l = parseInt(lAttr, 10);
      }
      if (tAttr !== undefined) {
        sr.t = parseInt(tAttr, 10);
      }
      if (rAttr !== undefined) {
        sr.r = parseInt(rAttr, 10);
      }
      if (bAttr !== undefined) {
        sr.b = parseInt(bAttr, 10);
      }
      if (Object.keys(sr).length > 0) {
        img.srcRect = sr;
      }
    }
  }

  return img;
}

// =============================================================================
// DrawingML Shape Parser
// =============================================================================

function parseDrawingShape(anchorEl: XmlElement, wspEl: XmlElement): DrawingShape | undefined {
  const docPrEl = findChild(anchorEl, "wp:docPr");
  const extentEl = findChild(anchorEl, "wp:extent");

  const cx = parseInt(extentEl?.attributes["cx"] ?? "0", 10);
  const cy = parseInt(extentEl?.attributes["cy"] ?? "0", 10);

  // Parse preset shape type from wps:spPr > a:prstGeom
  const spPrEl = findChild(wspEl, "wps:spPr") ?? findChildNs(wspEl, "spPr");
  const prstGeomEl = spPrEl
    ? (findChild(spPrEl, "a:prstGeom") ?? findChildNs(spPrEl, "prstGeom"))
    : undefined;
  const shapeType = prstGeomEl?.attributes["prst"] ?? "rect";

  const shape: any = {
    type: "drawingShape",
    shapeType,
    width: cx,
    height: cy,
    altText: docPrEl?.attributes["descr"],
    name: docPrEl?.attributes["name"]
  };

  // Parse fill
  if (spPrEl) {
    const solidFill = findChild(spPrEl, "a:solidFill") ?? findChildNs(spPrEl, "solidFill");
    if (solidFill) {
      const srgb = findChild(solidFill, "a:srgbClr") ?? findChildNs(solidFill, "srgbClr");
      if (srgb) {
        shape.fillColor = srgb.attributes["val"];
      }
    }
    const noFill = findChild(spPrEl, "a:noFill") ?? findChildNs(spPrEl, "noFill");
    if (noFill) {
      shape.noFill = true;
    }

    // Parse outline
    const lnEl = findChild(spPrEl, "a:ln") ?? findChildNs(spPrEl, "ln");
    if (lnEl) {
      const w = lnEl.attributes["w"];
      if (w) {
        shape.outlineWidth = parseInt(w, 10);
      }
      const lnFill = findChild(lnEl, "a:solidFill") ?? findChildNs(lnEl, "solidFill");
      if (lnFill) {
        const srgb = findChild(lnFill, "a:srgbClr") ?? findChildNs(lnFill, "srgbClr");
        if (srgb) {
          shape.outlineColor = srgb.attributes["val"];
        }
      }
      const noLn = findChild(lnEl, "a:noFill") ?? findChildNs(lnEl, "noFill");
      if (noLn) {
        shape.noOutline = true;
      }
    }
  }

  // Parse text content (wps:txbx > w:txbxContent)
  const txbxEl = findChild(wspEl, "wps:txbx") ?? findChildNs(wspEl, "txbx");
  const txbxContentEl = txbxEl
    ? (findChild(txbxEl, "w:txbxContent") ?? findChildNs(txbxEl, "txbxContent"))
    : undefined;
  if (txbxContentEl) {
    const paras: Paragraph[] = [];
    for (const child of txbxContentEl.children) {
      if (child.type === "element" && child.name.replace(/^w:/, "") === "p") {
        paras.push(parseParagraph(child));
      }
    }
    if (paras.length > 0) {
      shape.textContent = paras;
    }
  }

  // Parse positioning
  const posH = findChild(anchorEl, "wp:positionH");
  if (posH) {
    const hp: any = { relativeTo: posH.attributes["relativeFrom"] };
    const offsetEl = findChild(posH, "wp:posOffset");
    if (offsetEl) {
      hp.offset = parseInt(textContent(offsetEl), 10);
    }
    const alignEl = findChild(posH, "wp:align");
    if (alignEl) {
      hp.align = textContent(alignEl);
    }
    shape.horizontalPosition = hp;
  }
  const posV = findChild(anchorEl, "wp:positionV");
  if (posV) {
    const vp: any = { relativeTo: posV.attributes["relativeFrom"] };
    const offsetEl = findChild(posV, "wp:posOffset");
    if (offsetEl) {
      vp.offset = parseInt(textContent(offsetEl), 10);
    }
    const alignEl = findChild(posV, "wp:align");
    if (alignEl) {
      vp.align = textContent(alignEl);
    }
    shape.verticalPosition = vp;
  }

  // Wrap
  for (const wrapChild of anchorEl.children) {
    if (wrapChild.type !== "element") {
      continue;
    }
    const wn = wrapChild.name;
    if (wn === "wp:wrapSquare") {
      shape.wrap = { style: "square", side: wrapChild.attributes["wrapText"] };
    } else if (wn === "wp:wrapTight") {
      shape.wrap = { style: "tight", side: wrapChild.attributes["wrapText"] };
    } else if (wn === "wp:wrapTopAndBottom") {
      shape.wrap = { style: "topAndBottom" };
    } else if (wn === "wp:wrapNone") {
      shape.wrap = { style: "none" };
    }
  }

  // Behind doc
  if (anchorEl.attributes["behindDoc"] === "1") {
    shape.behindDoc = true;
  }

  // Rotation
  if (spPrEl) {
    const xfrmEl = findChild(spPrEl, "a:xfrm") ?? findChildNs(spPrEl, "xfrm");
    if (xfrmEl?.attributes["rot"]) {
      shape.rotation = parseInt(xfrmEl.attributes["rot"], 10);
    }
  }

  return shape;
}

// =============================================================================
// Math Parser
// =============================================================================

function findMathChild(el: XmlElement, localName: string): XmlElement | undefined {
  return findChild(el, `m:${localName}`) ?? findChild(el, localName);
}

function mathAttrVal(el: XmlElement, name: string): string | undefined {
  return el.attributes[`m:${name}`] ?? el.attributes[name];
}

function findMathChildren(el: XmlElement, localName: string): XmlElement[] {
  const a = findChildren(el, `m:${localName}`);
  return a.length > 0 ? a : findChildren(el, localName);
}

function parseMathContent(el: XmlElement): MathContent[] {
  const result: MathContent[] = [];
  for (const child of el.children) {
    if (child.type !== "element") {
      continue;
    }
    const name = child.name.replace(/^m:/, "");
    switch (name) {
      case "r": {
        // Math run
        const tEl = findMathChild(child, "t");
        const mrPrEl = findMathChild(child, "rPr");
        const mr: any = { type: "mathRun", text: tEl ? textContent(tEl) : "" };
        if (mrPrEl) {
          const props: any = {};
          const sty = findMathChild(mrPrEl, "sty");
          if (sty) {
            const v = sty.attributes["m:val"] ?? sty.attributes["val"];
            if (v === "p" || v === "b") {
              props.italic = false;
            }
            if (v === "b" || v === "bi") {
              props.bold = true;
            }
          }
          if (Object.keys(props).length > 0) {
            mr.properties = props;
          }
        }
        result.push(mr);
        break;
      }
      case "f": {
        const fPrEl = findMathChild(child, "fPr");
        const num = findMathChild(child, "num");
        const den = findMathChild(child, "den");
        const frac: any = {
          type: "mathFraction",
          numerator: num ? parseMathContent(num) : [],
          denominator: den ? parseMathContent(den) : []
        };
        if (fPrEl) {
          const typeEl = findMathChild(fPrEl, "type");
          if (typeEl) {
            frac.fractionType = typeEl.attributes["m:val"] ?? typeEl.attributes["val"];
          }
        }
        result.push(frac);
        break;
      }
      case "sSup": {
        const base = findMathChild(child, "e");
        const sup = findMathChild(child, "sup");
        result.push({
          type: "mathSuperScript",
          base: base ? parseMathContent(base) : [],
          superScript: sup ? parseMathContent(sup) : []
        });
        break;
      }
      case "sSub": {
        const base = findMathChild(child, "e");
        const sub = findMathChild(child, "sub");
        result.push({
          type: "mathSubScript",
          base: base ? parseMathContent(base) : [],
          subScript: sub ? parseMathContent(sub) : []
        });
        break;
      }
      case "sSubSup": {
        const base = findMathChild(child, "e");
        const sub = findMathChild(child, "sub");
        const sup = findMathChild(child, "sup");
        result.push({
          type: "mathSubSuperScript",
          base: base ? parseMathContent(base) : [],
          subScript: sub ? parseMathContent(sub) : [],
          superScript: sup ? parseMathContent(sup) : []
        });
        break;
      }
      case "sPre": {
        const base = findMathChild(child, "e");
        const sub = findMathChild(child, "sub");
        const sup = findMathChild(child, "sup");
        result.push({
          type: "mathPreSubSuperScript",
          base: base ? parseMathContent(base) : [],
          preSubScript: sub ? parseMathContent(sub) : [],
          preSuperScript: sup ? parseMathContent(sup) : []
        });
        break;
      }
      case "phant": {
        const eEl = findMathChild(child, "e");
        const phantPrEl = findMathChild(child, "phantPr");
        const ph: any = {
          type: "mathPhantom",
          content: eEl ? parseMathContent(eEl) : []
        };
        if (phantPrEl) {
          const boolAttr = (name: string): boolean => {
            const el = findMathChild(phantPrEl, name);
            if (!el) {
              return false;
            }
            const v = mathAttrVal(el, "val");
            return v !== "0" && v !== "false";
          };
          if (boolAttr("show")) {
            ph.show = true;
          }
          if (boolAttr("zeroWid")) {
            ph.zeroWidth = true;
          }
          if (boolAttr("zeroAsc")) {
            ph.zeroAscent = true;
          }
          if (boolAttr("zeroDesc")) {
            ph.zeroDescent = true;
          }
          if (boolAttr("transp")) {
            ph.transparent = true;
          }
        }
        result.push(ph);
        break;
      }
      case "groupChr": {
        const eEl = findMathChild(child, "e");
        const prEl = findMathChild(child, "groupChrPr");
        const g: any = {
          type: "mathGroupChar",
          base: eEl ? parseMathContent(eEl) : []
        };
        if (prEl) {
          const chrEl = findMathChild(prEl, "chr");
          if (chrEl) {
            g.char = mathAttrVal(chrEl, "val");
          }
          const posEl = findMathChild(prEl, "pos");
          if (posEl) {
            const v = mathAttrVal(posEl, "val");
            if (v === "top" || v === "bottom") {
              g.position = v;
            }
          }
          const vjcEl = findMathChild(prEl, "vertJc");
          if (vjcEl) {
            const v = mathAttrVal(vjcEl, "val");
            if (v === "top" || v === "center" || v === "bottom") {
              g.verticalAlign = v;
            }
          }
        }
        result.push(g);
        break;
      }
      case "borderBox": {
        const eEl = findMathChild(child, "e");
        const prEl = findMathChild(child, "borderBoxPr");
        const b: any = {
          type: "mathBorderBox",
          content: eEl ? parseMathContent(eEl) : []
        };
        if (prEl) {
          const boolAttr = (name: string): boolean => {
            const el = findMathChild(prEl, name);
            if (!el) {
              return false;
            }
            const v = mathAttrVal(el, "val");
            return v !== "0" && v !== "false";
          };
          if (boolAttr("hideTop")) {
            b.hideTop = true;
          }
          if (boolAttr("hideBot")) {
            b.hideBottom = true;
          }
          if (boolAttr("hideLeft")) {
            b.hideLeft = true;
          }
          if (boolAttr("hideRight")) {
            b.hideRight = true;
          }
          if (boolAttr("strikeBLTR")) {
            b.strikeBlTr = true;
          }
          if (boolAttr("strikeTLBR")) {
            b.strikeTlBr = true;
          }
          if (boolAttr("strikeH")) {
            b.strikeH = true;
          }
          if (boolAttr("strikeV")) {
            b.strikeV = true;
          }
        }
        result.push(b);
        break;
      }
      case "rad": {
        const radPrEl = findMathChild(child, "radPr");
        const deg = findMathChild(child, "deg");
        const e = findMathChild(child, "e");
        const rad: any = {
          type: "mathRadical",
          content: e ? parseMathContent(e) : []
        };
        if (deg) {
          rad.degree = parseMathContent(deg);
        }
        if (radPrEl) {
          const hd = findMathChild(radPrEl, "degHide");
          if (hd) {
            const v = hd.attributes["m:val"] ?? hd.attributes["val"];
            if (v === "1" || v === "on" || v === "true") {
              rad.hideDegree = true;
            }
          }
        }
        result.push(rad);
        break;
      }
      case "d": {
        const dPrEl = findMathChild(child, "dPr");
        const delim: any = { type: "mathDelimiter", content: [] };
        if (dPrEl) {
          const bc = findMathChild(dPrEl, "begChr");
          if (bc) {
            delim.beginChar = bc.attributes["m:val"] ?? bc.attributes["val"];
          }
          const ec = findMathChild(dPrEl, "endChr");
          if (ec) {
            delim.endChar = ec.attributes["m:val"] ?? ec.attributes["val"];
          }
          const sc = findMathChild(dPrEl, "sepChr");
          if (sc) {
            delim.separatorChar = sc.attributes["m:val"] ?? sc.attributes["val"];
          }
        }
        for (const eEl of findMathChildren(child, "e")) {
          delim.content.push(parseMathContent(eEl));
        }
        result.push(delim);
        break;
      }
      case "nary": {
        const nPrEl = findMathChild(child, "naryPr");
        const sub = findMathChild(child, "sub");
        const sup = findMathChild(child, "sup");
        const e = findMathChild(child, "e");
        const nary: any = {
          type: "mathNary",
          content: e ? parseMathContent(e) : []
        };
        if (sub) {
          nary.sub = parseMathContent(sub);
        }
        if (sup) {
          nary.sup = parseMathContent(sup);
        }
        if (nPrEl) {
          const chrEl = findMathChild(nPrEl, "chr");
          if (chrEl) {
            nary.char = chrEl.attributes["m:val"] ?? chrEl.attributes["val"];
          }
          const limLoc = findMathChild(nPrEl, "limLoc");
          if (limLoc) {
            nary.limitsLocation = limLoc.attributes["m:val"] ?? limLoc.attributes["val"];
          }
          const sh = findMathChild(nPrEl, "supHide");
          if (sh && (sh.attributes["m:val"] ?? sh.attributes["val"]) === "1") {
            nary.supHide = true;
          }
          const sbh = findMathChild(nPrEl, "subHide");
          if (sbh && (sbh.attributes["m:val"] ?? sbh.attributes["val"]) === "1") {
            nary.subHide = true;
          }
        }
        result.push(nary);
        break;
      }
      case "func": {
        const fName = findMathChild(child, "fName");
        const e = findMathChild(child, "e");
        result.push({
          type: "mathFunction",
          name: fName ? parseMathContent(fName) : [],
          content: e ? parseMathContent(e) : []
        });
        break;
      }
      case "limLow": {
        const base = findMathChild(child, "e");
        const lim = findMathChild(child, "lim");
        result.push({
          type: "mathLimit",
          limitType: "lower",
          base: base ? parseMathContent(base) : [],
          limit: lim ? parseMathContent(lim) : []
        });
        break;
      }
      case "limUpp": {
        const base = findMathChild(child, "e");
        const lim = findMathChild(child, "lim");
        result.push({
          type: "mathLimit",
          limitType: "upper",
          base: base ? parseMathContent(base) : [],
          limit: lim ? parseMathContent(lim) : []
        });
        break;
      }
      case "m": {
        // Matrix
        const rows: MathContent[][][] = [];
        for (const mrEl of findMathChildren(child, "mr")) {
          const row: MathContent[][] = [];
          for (const eEl of findMathChildren(mrEl, "e")) {
            row.push(parseMathContent(eEl));
          }
          rows.push(row);
        }
        result.push({ type: "mathMatrix", rows });
        break;
      }
      case "acc": {
        const accPrEl = findMathChild(child, "accPr");
        const e = findMathChild(child, "e");
        const acc: any = {
          type: "mathAccent",
          content: e ? parseMathContent(e) : []
        };
        if (accPrEl) {
          const chr = findMathChild(accPrEl, "chr");
          if (chr) {
            acc.char = chr.attributes["m:val"] ?? chr.attributes["val"];
          }
        }
        result.push(acc);
        break;
      }
      case "bar": {
        const barPrEl = findMathChild(child, "barPr");
        const e = findMathChild(child, "e");
        let position: "top" | "bottom" = "top";
        if (barPrEl) {
          const pos = findMathChild(barPrEl, "pos");
          if (pos) {
            const v = pos.attributes["m:val"] ?? pos.attributes["val"];
            if (v === "bot") {
              position = "bottom";
            }
          }
        }
        result.push({
          type: "mathBar",
          position,
          content: e ? parseMathContent(e) : []
        });
        break;
      }
      case "box": {
        const e = findMathChild(child, "e");
        result.push({
          type: "mathBox",
          content: e ? parseMathContent(e) : []
        });
        break;
      }
      case "eqArr": {
        const rows: MathContent[][] = [];
        for (const eEl of findMathChildren(child, "e")) {
          rows.push(parseMathContent(eEl));
        }
        result.push({ type: "mathEquationArray", rows });
        break;
      }
      // Recurse into oMath elements
      case "oMath": {
        result.push(...parseMathContent(child));
        break;
      }
    }
  }
  return result;
}

function parseMathBlock(oMathParaEl: XmlElement): MathBlock {
  const content: MathContent[] = [];
  for (const child of oMathParaEl.children) {
    if (child.type === "element") {
      const n = child.name.replace(/^m:/, "");
      if (n === "oMath") {
        content.push(...parseMathContent(child));
      }
    }
  }
  return { type: "math", content };
}

// =============================================================================
// TextBox Parser
// =============================================================================

function parseTextBox(pictEl: XmlElement): TextBox | undefined {
  // Look for v:shape > v:textbox > w:txbxContent
  let txbxContentEl: XmlElement | undefined;
  let shapeEl: XmlElement | undefined;

  for (const child of pictEl.children) {
    if (child.type === "element" && (child.name === "v:shape" || child.name === "v:rect")) {
      shapeEl = child;
      for (const sc of child.children) {
        if (sc.type === "element" && sc.name === "v:textbox") {
          for (const tc of sc.children) {
            if (
              tc.type === "element" &&
              (tc.name === "w:txbxContent" || tc.name === "txbxContent")
            ) {
              txbxContentEl = tc;
            }
          }
        }
      }
    }
  }

  if (!txbxContentEl) {
    return undefined;
  }

  const paragraphs: Paragraph[] = [];
  for (const c of txbxContentEl.children) {
    if (c.type === "element" && c.name.replace(/^w:/, "") === "p") {
      paragraphs.push(parseParagraph(c));
    }
  }

  const tb: any = { type: "textBox", content: paragraphs };

  if (shapeEl) {
    const style = shapeEl.attributes["style"];
    if (style) {
      tb.style = style;
    }
    const sc = shapeEl.attributes["strokecolor"];
    if (sc) {
      tb.strokeColor = sc;
    }
    const fc = shapeEl.attributes["fillcolor"];
    if (fc) {
      tb.fillColor = fc;
    }
    if (shapeEl.attributes["stroked"] === "f") {
      tb.stroke = false;
    }
    if (shapeEl.attributes["filled"] === "f") {
      tb.fill = false;
    }
  }

  return tb;
}

// =============================================================================
// SDT / CheckBox / TOC Parser
// =============================================================================

function parseSdt(
  sdtEl: XmlElement
): StructuredDocumentTag | CheckBox | TableOfContents | undefined {
  const sdtPrEl = findChildNs(sdtEl, "sdtPr");
  const sdtContentEl = findChildNs(sdtEl, "sdtContent");

  // Check for checkbox (w14:checkbox)
  if (sdtPrEl) {
    const checkBoxEl = findChild(sdtPrEl, "w14:checkbox");
    if (checkBoxEl) {
      return parseCheckBox(checkBoxEl);
    }
  }

  // Check for TOC (contains docPartObj with docPartGallery "Table of Contents")
  if (sdtPrEl) {
    const docPartObjEl = findChildNs(sdtPrEl, "docPartObj");
    if (docPartObjEl) {
      const galleryEl = findChildNs(docPartObjEl, "docPartGallery");
      const galleryVal = galleryEl ? attrVal(galleryEl, "val") : undefined;
      if (galleryVal === "Table of Contents") {
        return parseTocFromSdt(sdtContentEl);
      }
    }
  }

  // Generic SDT
  const props: SdtProperties = {};
  if (sdtPrEl) {
    const tagEl = findChildNs(sdtPrEl, "tag");
    if (tagEl) {
      (props as any).tag = attrVal(tagEl, "val");
    }
    const aliasEl = findChildNs(sdtPrEl, "alias");
    if (aliasEl) {
      (props as any).alias = attrVal(aliasEl, "val");
    }
    const lockEl = findChildNs(sdtPrEl, "lock");
    if (lockEl) {
      const v = attrVal(lockEl, "val");
      if (v === "contentLocked" || v === "sdtContentLocked") {
        (props as any).lockContent = true;
      }
      if (v === "sdtLocked" || v === "sdtContentLocked") {
        (props as any).lockSdt = true;
      }
    }
    // Plain text
    if (findChildNs(sdtPrEl, "text")) {
      (props as any).plainText = true;
    }
    // showingPlcHdr is a toggle, not a property with a val
    if (findChildNs(sdtPrEl, "showingPlcHdr")) {
      const v = boolToggle(sdtPrEl, "showingPlcHdr");
      if (v !== false) {
        (props as any).showingPlaceholder = true;
      }
    }
    // w15:appearance (replaces the old misused showingPlcHdr)
    const appearanceEl = findChild(sdtPrEl, "w15:appearance");
    if (appearanceEl) {
      const v = appearanceEl.attributes["w15:val"] ?? appearanceEl.attributes["val"];
      if (v === "boundingBox" || v === "tags" || v === "hidden") {
        (props as any).appearance = v;
      }
    }
    // Dropdown list
    const ddlEl = findChildNs(sdtPrEl, "dropDownList");
    if (ddlEl) {
      const items: SdtListItem[] = [];
      for (const li of findChildrenNs(ddlEl, "listItem")) {
        const item: SdtListItem = { value: attrVal(li, "value") ?? "" };
        const dt = attrVal(li, "displayText");
        if (dt) {
          (item as any).displayText = dt;
        }
        items.push(item);
      }
      (props as any).dropdownList = items;
    }
    // ComboBox
    const cbEl = findChildNs(sdtPrEl, "comboBox");
    if (cbEl) {
      const items: SdtListItem[] = [];
      for (const li of findChildrenNs(cbEl, "listItem")) {
        const item: SdtListItem = { value: attrVal(li, "value") ?? "" };
        const dt = attrVal(li, "displayText");
        if (dt) {
          (item as any).displayText = dt;
        }
        items.push(item);
      }
      (props as any).comboBox = items;
    }
    // Date picker
    const dateEl = findChildNs(sdtPrEl, "date");
    if (dateEl) {
      const dateProp: SdtDateProperties = {} as any;
      const fullDate = attrVal(dateEl, "fullDate");
      if (fullDate) {
        (dateProp as any).fullDate = fullDate;
      }
      const dfEl = findChildNs(dateEl, "dateFormat");
      if (dfEl) {
        (dateProp as any).dateFormat = attrVal(dfEl, "val");
      }
      const lidEl = findChildNs(dateEl, "lid");
      if (lidEl) {
        (dateProp as any).lid = attrVal(lidEl, "val");
      }
      const storeEl = findChildNs(dateEl, "storeMappedDataAs");
      if (storeEl) {
        (dateProp as any).storeMappedDataAs = attrVal(storeEl, "val");
      }
      (props as any).date = dateProp;
    }
    // ID
    const idEl = findChildNs(sdtPrEl, "id");
    if (idEl) {
      const v = attrInt(idEl, "val");
      if (v !== undefined) {
        (props as any).id = v;
      }
    }
    // Data binding
    const dbEl = findChildNs(sdtPrEl, "dataBinding");
    if (dbEl) {
      const xpath = attrVal(dbEl, "xpath");
      const storeItemId = attrVal(dbEl, "storeItemID");
      if (xpath && storeItemId) {
        const binding: any = { xpath, storeItemId };
        const prefixMappings = attrVal(dbEl, "prefixMappings");
        if (prefixMappings) {
          binding.prefixMappings = prefixMappings;
        }
        (props as any).dataBinding = binding;
      }
    }
    // Placeholder
    const phEl = findChildNs(sdtPrEl, "placeholder");
    if (phEl) {
      const docPartEl = findChildNs(phEl, "docPart");
      if (docPartEl) {
        (props as any).placeholder = attrVal(docPartEl, "val");
      }
    }
    // Boolean marker elements
    if (findChildNs(sdtPrEl, "richText")) {
      (props as any).richText = true;
    }
    if (findChildNs(sdtPrEl, "picture")) {
      (props as any).picture = true;
    }
    if (findChildNs(sdtPrEl, "group")) {
      (props as any).group = true;
    }
    if (findChildNs(sdtPrEl, "equation")) {
      (props as any).equation = true;
    }
    if (findChildNs(sdtPrEl, "citation")) {
      (props as any).citation = true;
    }
    if (findChildNs(sdtPrEl, "bibliography")) {
      (props as any).bibliography = true;
    }
    if (findChildNs(sdtPrEl, "temporary")) {
      (props as any).temporary = true;
    }
    // w15: repeating section
    const rsEl = findChild(sdtPrEl, "w15:repeatingSection");
    if (rsEl) {
      const rs: any = {};
      // Read from child elements (correct per schema)
      const titleEl = findChild(rsEl, "w15:sectionTitle");
      if (titleEl) {
        const v = titleEl.attributes["w15:val"] ?? titleEl.attributes["val"];
        if (v !== undefined) {
          rs.sectionTitle = v;
        }
      }
      if (findChild(rsEl, "w15:doNotAllowInsertDeleteSection")) {
        rs.allowInsertDelete = false;
      }
      // Also accept attribute form for backwards compatibility
      const stAttr = rsEl.attributes["w15:sectionTitle"];
      if (stAttr !== undefined && rs.sectionTitle === undefined) {
        rs.sectionTitle = stAttr;
      }
      const noInsDelAttr = rsEl.attributes["w15:doNotAllowInsertDeleteSection"];
      if (noInsDelAttr !== undefined && rs.allowInsertDelete === undefined) {
        rs.allowInsertDelete = noInsDelAttr === "0";
      }
      (props as any).repeatingSection = rs;
    }
    if (findChild(sdtPrEl, "w15:repeatingSectionItem")) {
      (props as any).repeatingSectionItem = true;
    }
  }

  const content: (Paragraph | Run | Table)[] = [];
  if (sdtContentEl) {
    for (const child of sdtContentEl.children) {
      if (child.type !== "element") {
        continue;
      }
      const n = child.name.replace(/^w:/, "");
      if (n === "p") {
        content.push(parseParagraph(child));
      } else if (n === "tbl") {
        content.push(parseTable(child));
      } else if (n === "r") {
        content.push(parseRun(child));
      }
    }
  }

  return { type: "sdt", properties: props, content };
}

function parseCheckBox(checkBoxEl: XmlElement): CheckBox {
  const cb: any = { type: "checkBox" };
  const checkedEl = findChild(checkBoxEl, "w14:checked");
  if (checkedEl) {
    const v = checkedEl.attributes["w14:val"] ?? checkedEl.attributes["val"];
    cb.checked = v === "1" || v === "true";
  }
  const checkedStateEl = findChild(checkBoxEl, "w14:checkedState");
  if (checkedStateEl) {
    cb.checkedState = {
      value: checkedStateEl.attributes["w14:val"] ?? checkedStateEl.attributes["val"] ?? "",
      font: checkedStateEl.attributes["w14:font"] ?? checkedStateEl.attributes["font"]
    };
  }
  const uncheckedStateEl = findChild(checkBoxEl, "w14:uncheckedState");
  if (uncheckedStateEl) {
    cb.uncheckedState = {
      value: uncheckedStateEl.attributes["w14:val"] ?? uncheckedStateEl.attributes["val"] ?? "",
      font: uncheckedStateEl.attributes["w14:font"] ?? uncheckedStateEl.attributes["font"]
    };
  }
  return cb;
}

function parseTocFromSdt(sdtContentEl: XmlElement | undefined): TableOfContents {
  const toc: any = { type: "tableOfContents" };
  const cachedParagraphs: Paragraph[] = [];

  if (sdtContentEl) {
    // Collect all instrText to assemble the complete TOC field instruction
    let instrText = "";
    const collectInstr = (el: XmlElement): void => {
      for (const child of el.children) {
        if (child.type !== "element") {
          continue;
        }
        const name = child.name.replace(/^w:/, "");
        if (name === "instrText") {
          instrText += textContent(child);
        } else {
          collectInstr(child);
        }
      }
    };
    collectInstr(sdtContentEl);

    if (instrText.trim()) {
      parseTocInstruction(instrText, toc);
    }

    for (const child of sdtContentEl.children) {
      if (child.type !== "element") {
        continue;
      }
      const n = child.name.replace(/^w:/, "");
      if (n === "p") {
        cachedParagraphs.push(parseParagraph(child));
      }
    }
  }

  if (cachedParagraphs.length > 0) {
    toc.cachedParagraphs = cachedParagraphs;
  }

  return toc;
}

/** Parse a TOC field instruction string (e.g. `TOC \o "1-3" \h \t "Style,1" \c "Figure"`). */
function parseTocInstruction(instr: string, toc: any): void {
  const trimmed = instr.trim();
  if (!/^TOC\b/i.test(trimmed)) {
    return;
  }
  // Match switches: \<letter> followed by either "quoted" or non-quoted non-switch token.
  // The next-switch boundary must be respected: an unquoted value cannot start with \.
  const switchRe = /\\(\w)(?:\s+"([^"]*)"|\s+([^\\\s][^\s]*))?/g;
  let match: RegExpExecArray | null;
  while ((match = switchRe.exec(trimmed)) !== null) {
    const switchName = match[1].toLowerCase();
    const value = match[2] ?? match[3];
    switch (switchName) {
      case "o": // Heading level range e.g. "1-3"
        if (value) {
          toc.headingStyleRange = value;
        }
        break;
      case "h": // Hyperlinks
        toc.hyperlink = true;
        break;
      case "c": // Caption label (table of figures)
        if (value) {
          toc.captionLabel = value;
        }
        break;
      case "s": // Sequence field identifier
        if (value) {
          toc.sequenceFieldIdentifier = value;
        }
        break;
      case "p": // Page-number leader or style separator
        // In real TOC fields, \p is sometimes used for tab leader.
        // Common values: "." "-" "_"
        if (value === "." || value === "-" || value === "_") {
          toc.leader = "dot";
          if (value === "-") {
            toc.leader = "hyphen";
          } else if (value === "_") {
            toc.leader = "underscore";
          }
        }
        break;
      case "t": {
        // Styles with levels: "StyleName1,Level1;StyleName2,Level2;..."
        if (!value) {
          break;
        }
        const items: { styleName: string; level: number }[] = [];
        for (const part of value.split(";")) {
          const [styleName, levelStr] = part.split(",");
          if (styleName && levelStr) {
            items.push({ styleName: styleName.trim(), level: parseInt(levelStr, 10) });
          }
        }
        if (items.length > 0) {
          toc.stylesWithLevels = items;
        }
        break;
      }
    }
  }
}

// =============================================================================
// Paragraph Parser
// =============================================================================

function parseRun(el: XmlElement): Run {
  const rPrEl = findChildNs(el, "rPr");
  return {
    properties: rPrEl ? parseRunProperties(rPrEl) : undefined,
    content: parseRunContent(el)
  };
}

function parseParagraph(pEl: XmlElement): Paragraph {
  const pPrEl = findChildNs(pEl, "pPr");
  const children: ParagraphChild[] = [];

  // fldChar state machine: tracks field code assembly across runs
  let fieldState: "none" | "instrText" | "cached" = "none";
  let fieldInstr = "";
  let fieldCached = "";
  let fieldRunProps: RunProperties | undefined;
  let fieldFormField: FormField | undefined;

  for (const child of pEl.children) {
    if (child.type !== "element") {
      continue;
    }
    // Handle mc:AlternateContent — pick mc:Choice, fall back to mc:Fallback
    let resolved = child;
    if (child.name === "mc:AlternateContent") {
      const choice = findChild(child, "mc:Choice");
      const fallback = findChild(child, "mc:Fallback");
      const chosen = choice ?? fallback;
      if (chosen && chosen.children.length > 0) {
        // The first element child inside Choice/Fallback is the real element
        const inner = chosen.children.find(c => c.type === "element") as XmlElement | undefined;
        if (inner) {
          resolved = inner;
        } else {
          continue;
        }
      } else {
        continue;
      }
    }

    const name = resolved.name.replace(/^w:/, "");
    switch (name) {
      case "r": {
        // Check for fldChar and instrText inside the run
        let hasFldChar = false;
        for (const rc of resolved.children) {
          if (rc.type !== "element") {
            continue;
          }
          const rcName = rc.name.replace(/^w:/, "");
          if (rcName === "fldChar") {
            hasFldChar = true;
            const fldCharType = attrVal(rc, "fldCharType");
            if (fldCharType === "begin") {
              fieldState = "instrText";
              fieldInstr = "";
              fieldCached = "";
              // Capture run properties from this run for the field
              const rPrEl = findChildNs(resolved, "rPr");
              fieldRunProps = rPrEl ? parseRunProperties(rPrEl) : undefined;
              // Parse ffData for legacy form fields
              const ffDataEl = findChildNs(rc, "ffData");
              if (ffDataEl) {
                fieldFormField = parseFfData(ffDataEl);
              } else {
                fieldFormField = undefined;
              }
            } else if (fldCharType === "separate") {
              fieldState = "cached";
            } else if (fldCharType === "end") {
              // Emit the assembled field as a Run with FieldContent
              const fc: FieldContent = {
                type: "field",
                instruction: fieldInstr.trim(),
                cachedValue: fieldCached || undefined,
                formField: fieldFormField
              };
              children.push({
                properties: fieldRunProps,
                content: [fc]
              } satisfies Run);
              fieldState = "none";
              fieldInstr = "";
              fieldCached = "";
              fieldRunProps = undefined;
            }
          } else if (rcName === "instrText" && fieldState === "instrText") {
            hasFldChar = true;
            fieldInstr += textContent(rc);
          }
        }

        if (fieldState === "cached") {
          // Collect cached text from this run
          for (const rc of resolved.children) {
            if (rc.type !== "element") {
              continue;
            }
            const rcName = rc.name.replace(/^w:/, "");
            if (rcName === "t") {
              fieldCached += textContent(rc);
            } else if (rcName === "fldChar") {
              // Already handled above
            }
          }
          if (!hasFldChar) {
            continue; // Skip adding this run normally
          }
        }

        if (fieldState === "instrText" && hasFldChar) {
          continue; // Don't add begin/instrText runs as normal content
        }
        if (fieldState === "none" && !hasFldChar) {
          children.push(parseRun(resolved));
        }
        break;
      }
      case "fldSimple": {
        // Simple field: <w:fldSimple w:instr=" PAGE "><w:r>...</w:r></w:fldSimple>
        const instr = attrVal(resolved, "instr") ?? "";
        let cached = "";
        for (const fc of resolved.children) {
          if (fc.type === "element" && fc.name.replace(/^w:/, "") === "r") {
            for (const rc of fc.children) {
              if (rc.type === "element" && rc.name.replace(/^w:/, "") === "t") {
                cached += textContent(rc);
              }
            }
          }
        }
        const fc: FieldContent = {
          type: "field",
          instruction: instr.trim(),
          cachedValue: cached || undefined
        };
        children.push({
          properties: undefined,
          content: [fc]
        } satisfies Run);
        break;
      }
      case "hyperlink": {
        const rId = resolved.attributes["r:id"];
        const anchor = resolved.attributes["w:anchor"] ?? resolved.attributes["anchor"];
        const tooltip = resolved.attributes["w:tooltip"] ?? resolved.attributes["tooltip"];
        const historyAttr = resolved.attributes["w:history"] ?? resolved.attributes["history"];
        const tgtFrame = resolved.attributes["w:tgtFrame"] ?? resolved.attributes["tgtFrame"];
        const docLocation =
          resolved.attributes["w:docLocation"] ?? resolved.attributes["docLocation"];
        const hRuns: Run[] = [];
        for (const hChild of resolved.children) {
          if (hChild.type === "element" && hChild.name.replace(/^w:/, "") === "r") {
            hRuns.push(parseRun(hChild));
          }
        }
        // Resolve URL from relMap
        let url: string | undefined;
        if (rId) {
          const rel = _parseRelMap.get(rId);
          if (rel && rel.targetMode === "External") {
            url = rel.target;
          }
        }
        const hyperlink: any = {
          type: "hyperlink",
          rId,
          anchor,
          url,
          tooltip,
          children: hRuns
        };
        if (historyAttr === "1" || historyAttr === "true") {
          hyperlink.history = true;
        }
        if (tgtFrame) {
          hyperlink.tgtFrame = tgtFrame;
        }
        if (docLocation) {
          hyperlink.docLocation = docLocation;
        }
        children.push(hyperlink);
        break;
      }
      case "bookmarkStart": {
        const bm: any = {
          type: "bookmarkStart",
          id: parseInt(resolved.attributes["w:id"] ?? resolved.attributes["id"] ?? "0", 10),
          name: resolved.attributes["w:name"] ?? resolved.attributes["name"] ?? ""
        };
        const colFirst = resolved.attributes["w:colFirst"] ?? resolved.attributes["colFirst"];
        if (colFirst !== undefined) {
          bm.colFirst = parseInt(colFirst, 10);
        }
        const colLast = resolved.attributes["w:colLast"] ?? resolved.attributes["colLast"];
        if (colLast !== undefined) {
          bm.colLast = parseInt(colLast, 10);
        }
        const dcx =
          resolved.attributes["w:displacedByCustomXml"] ??
          resolved.attributes["displacedByCustomXml"];
        if (dcx === "next" || dcx === "prev") {
          bm.displacedByCustomXml = dcx;
        }
        children.push(bm);
        break;
      }
      case "bookmarkEnd":
        children.push({
          type: "bookmarkEnd",
          id: parseInt(resolved.attributes["w:id"] ?? resolved.attributes["id"] ?? "0", 10)
        });
        break;
      case "commentRangeStart":
        children.push({
          type: "commentRangeStart",
          id: parseInt(resolved.attributes["w:id"] ?? resolved.attributes["id"] ?? "0", 10)
        });
        break;
      case "commentRangeEnd":
        children.push({
          type: "commentRangeEnd",
          id: parseInt(resolved.attributes["w:id"] ?? resolved.attributes["id"] ?? "0", 10)
        });
        break;
      case "commentReference":
        children.push({
          type: "commentReference",
          id: parseInt(resolved.attributes["w:id"] ?? resolved.attributes["id"] ?? "0", 10)
        });
        break;
      case "ins": {
        // Inserted run (track changes)
        const rev = parseRevisionInfo(resolved);
        if (rev) {
          for (const insChild of resolved.children) {
            if (insChild.type === "element" && insChild.name.replace(/^w:/, "") === "r") {
              children.push({
                type: "insertedRun",
                revision: rev,
                run: parseRun(insChild)
              } satisfies InsertedRun);
            }
          }
        }
        break;
      }
      case "del": {
        // Deleted run (track changes)
        const rev = parseRevisionInfo(resolved);
        if (rev) {
          for (const delChild of resolved.children) {
            if (delChild.type === "element" && delChild.name.replace(/^w:/, "") === "r") {
              children.push({
                type: "deletedRun",
                revision: rev,
                run: parseDeletedRun(delChild)
              } satisfies DeletedRun);
            }
          }
        }
        break;
      }
      case "moveFrom": {
        const rev = parseRevisionInfo(resolved);
        if (rev) {
          for (const mfChild of resolved.children) {
            if (mfChild.type === "element" && mfChild.name.replace(/^w:/, "") === "r") {
              children.push({
                type: "movedFromRun",
                revision: rev,
                run: parseRun(mfChild)
              });
            }
          }
        }
        break;
      }
      case "moveTo": {
        const rev = parseRevisionInfo(resolved);
        if (rev) {
          for (const mtChild of resolved.children) {
            if (mtChild.type === "element" && mtChild.name.replace(/^w:/, "") === "r") {
              children.push({
                type: "movedToRun",
                revision: rev,
                run: parseRun(mtChild)
              });
            }
          }
        }
        break;
      }
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd": {
        const id = attrInt(resolved, "id");
        if (id !== undefined) {
          const marker: any = {
            type: name,
            id
          };
          const author = attrVal(resolved, "author");
          if (author) {
            marker.author = author;
          }
          const date = attrVal(resolved, "date");
          if (date) {
            marker.date = date;
          }
          const mName = attrVal(resolved, "name");
          if (mName) {
            marker.name = mName;
          }
          children.push(marker);
        }
        break;
      }
      case "customXmlInsRangeStart":
      case "customXmlInsRangeEnd":
      case "customXmlDelRangeStart":
      case "customXmlDelRangeEnd":
      case "customXmlMoveFromRangeStart":
      case "customXmlMoveFromRangeEnd":
      case "customXmlMoveToRangeStart":
      case "customXmlMoveToRangeEnd": {
        const id = attrInt(resolved, "id");
        if (id !== undefined) {
          const marker: any = { type: name, id };
          const author = attrVal(resolved, "author");
          if (author) {
            marker.author = author;
          }
          const date = attrVal(resolved, "date");
          if (date) {
            marker.date = date;
          }
          children.push(marker);
        }
        break;
      }
      case "smartTag":
      case "customXml":
      case "dir": {
        // Semantic wrappers: flatten their children into the current paragraph.
        // A smartTag/customXml/dir can contain runs, hyperlinks, nested wrappers, etc.
        // Re-use parseParagraph to recursively parse the contained children.
        const subPara = parseParagraph(resolved);
        for (const sub of subPara.children) {
          children.push(sub);
        }
        break;
      }
      case "proofErr":
      case "permStart":
      case "permEnd":
      case "lastRenderedPageBreak":
        // Non-semantic markers; safely ignored
        break;
    }
  }

  const paraId = pEl.attributes["w14:paraId"];
  const textId = pEl.attributes["w14:textId"];

  const result: any = {
    type: "paragraph",
    properties: pPrEl ? parseParagraphProperties(pPrEl) : undefined,
    children
  };
  if (paraId) {
    result.paraId = paraId;
  }
  if (textId) {
    result.textId = textId;
  }
  return result;
}

/** Parse a deleted run (w:delText instead of w:t). */
function parseDeletedRun(el: XmlElement): Run {
  const rPrEl = findChildNs(el, "rPr");
  const content: RunContent[] = [];
  for (const child of el.children) {
    if (child.type !== "element") {
      continue;
    }
    const name = child.name.replace(/^w:/, "");
    if (name === "delText") {
      content.push({ type: "text", text: textContent(child) });
    } else if (name === "t") {
      content.push({ type: "text", text: textContent(child) });
    } else if (name === "br") {
      content.push({ type: "break", breakType: attrVal(child, "type") as any });
    } else if (name === "tab") {
      content.push({ type: "tab" });
    }
  }
  return {
    properties: rPrEl ? parseRunProperties(rPrEl) : undefined,
    content
  };
}

// =============================================================================
// Table Parser
// =============================================================================

function parseTableBorders(el: XmlElement): TableBorders {
  const borders: any = {};
  for (const side of [
    "top",
    "left",
    "bottom",
    "right",
    "insideH",
    "insideV",
    "start",
    "end",
    "tl2br",
    "tr2bl"
  ] as const) {
    const sideEl = findChildNs(el, side);
    if (sideEl) {
      borders[side] = parseBorder(sideEl);
    }
  }
  return borders;
}

function parseTableCellMargins(el: XmlElement): TableCellMargins {
  const margins: any = {};
  for (const side of ["top", "left", "bottom", "right", "start", "end"] as const) {
    const sideEl = findChildNs(el, side);
    if (sideEl) {
      margins[side] = parseTableWidth(sideEl);
    }
  }
  return margins;
}

function parseTableProperties(el: XmlElement): TableProperties {
  const props: any = {};

  const styleEl = findChildNs(el, "tblStyle");
  if (styleEl) {
    props.style = attrVal(styleEl, "val");
  }

  const wEl = findChildNs(el, "tblW");
  if (wEl) {
    props.width = parseTableWidth(wEl);
  }

  const jcEl = findChildNs(el, "jc");
  if (jcEl) {
    props.alignment = attrVal(jcEl, "val");
  }

  const indEl = findChildNs(el, "tblInd");
  if (indEl) {
    props.indent = parseInt(indEl.attributes["w:w"] ?? indEl.attributes["w"] ?? "0", 10);
  }

  const bordersEl = findChildNs(el, "tblBorders");
  if (bordersEl) {
    props.borders = parseTableBorders(bordersEl);
  }

  const layoutEl = findChildNs(el, "tblLayout");
  if (layoutEl) {
    props.layout = attrVal(layoutEl, "type");
  }

  const cellMarEl = findChildNs(el, "tblCellMar");
  if (cellMarEl) {
    props.cellMargins = parseTableCellMargins(cellMarEl);
  }

  // TableLook
  const lookEl = findChildNs(el, "tblLook");
  if (lookEl) {
    const look: any = {};

    // Read individual attributes first (authoritative when explicit "0"/"1")
    const readFlag = (name: string): boolean | undefined => {
      const v = attrVal(lookEl, name);
      if (v === "1" || v === "true") {
        return true;
      }
      if (v === "0" || v === "false") {
        return false;
      }
      return undefined;
    };

    const firstRow = readFlag("firstRow");
    const lastRow = readFlag("lastRow");
    const firstColumn = readFlag("firstColumn");
    const lastColumn = readFlag("lastColumn");
    const noHBand = readFlag("noHBand");
    const noVBand = readFlag("noVBand");

    if (firstRow !== undefined) {
      look.firstRow = firstRow;
    }
    if (lastRow !== undefined) {
      look.lastRow = lastRow;
    }
    if (firstColumn !== undefined) {
      look.firstColumn = firstColumn;
    }
    if (lastColumn !== undefined) {
      look.lastColumn = lastColumn;
    }
    if (noHBand !== undefined) {
      look.noHBand = noHBand;
    }
    if (noVBand !== undefined) {
      look.noVBand = noVBand;
    }

    // Fall back to w:val bitmask ONLY if no individual attrs were specified
    if (Object.keys(look).length === 0) {
      const val = attrVal(lookEl, "val");
      if (val) {
        const v = parseInt(val, 16);
        if (v & 0x0020) {
          look.firstRow = true;
        }
        if (v & 0x0040) {
          look.lastRow = true;
        }
        if (v & 0x0080) {
          look.firstColumn = true;
        }
        if (v & 0x0100) {
          look.lastColumn = true;
        }
        if (v & 0x0200) {
          look.noHBand = true;
        }
        if (v & 0x0400) {
          look.noVBand = true;
        }
      }
    }

    if (Object.keys(look).length > 0) {
      props.look = look;
    }
  }

  // TableFloat
  const tblpPrEl = findChildNs(el, "tblpPr");
  if (tblpPrEl) {
    const tf: TableFloat = {} as any;
    const f: any = tf;
    const hAnchor = attrVal(tblpPrEl, "horzAnchor");
    if (hAnchor) {
      f.horizontalAnchor = hAnchor;
    }
    const vAnchor = attrVal(tblpPrEl, "vertAnchor");
    if (vAnchor) {
      f.verticalAnchor = vAnchor;
    }
    const tblpX = attrInt(tblpPrEl, "tblpX");
    if (tblpX !== undefined) {
      f.absoluteHorizontalPosition = tblpX;
    }
    const tblpY = attrInt(tblpPrEl, "tblpY");
    if (tblpY !== undefined) {
      f.absoluteVerticalPosition = tblpY;
    }
    const tblpXSpec = attrVal(tblpPrEl, "tblpXSpec");
    if (tblpXSpec) {
      f.relativeHorizontalPosition = tblpXSpec;
    }
    const tblpYSpec = attrVal(tblpPrEl, "tblpYSpec");
    if (tblpYSpec) {
      f.relativeVerticalPosition = tblpYSpec;
    }
    const topFromText = attrInt(tblpPrEl, "topFromText");
    if (topFromText !== undefined) {
      f.topFromText = topFromText;
    }
    const bottomFromText = attrInt(tblpPrEl, "bottomFromText");
    if (bottomFromText !== undefined) {
      f.bottomFromText = bottomFromText;
    }
    const leftFromText = attrInt(tblpPrEl, "leftFromText");
    if (leftFromText !== undefined) {
      f.leftFromText = leftFromText;
    }
    const rightFromText = attrInt(tblpPrEl, "rightFromText");
    if (rightFromText !== undefined) {
      f.rightFromText = rightFromText;
    }
    const overlap = attrVal(tblpPrEl, "overlap");
    if (overlap) {
      f.overlap = overlap;
    }
    props.float = tf;
  }

  // w:tblOverlap is a separate sibling element of w:tblpPr (value "never"|"overlap")
  const tblOverlapEl = findChildNs(el, "tblOverlap");
  if (tblOverlapEl && props.float) {
    const v = attrVal(tblOverlapEl, "val");
    if (v === "never" || v === "overlap") {
      (props.float as any).overlap = v;
    }
  }

  // Cell spacing
  const csEl = findChildNs(el, "tblCellSpacing");
  if (csEl) {
    props.cellSpacing = parseTableWidth(csEl);
  }

  // Bidi visual
  if (findChildNs(el, "bidiVisual")) {
    props.visuallyRightToLeft = true;
  }

  // Shading
  const shdEl = findChildNs(el, "shd");
  if (shdEl) {
    props.shading = parseShading(shdEl);
  }

  // Accessibility: caption and description
  const captionEl = findChildNs(el, "tblCaption");
  if (captionEl) {
    props.caption = attrVal(captionEl, "val");
  }
  const descEl = findChildNs(el, "tblDescription");
  if (descEl) {
    props.description = attrVal(descEl, "val");
  }

  // Table property change
  const tblPrChangeEl = findChildNs(el, "tblPrChange");
  if (tblPrChangeEl) {
    const rev = parseRevisionInfo(tblPrChangeEl);
    if (rev) {
      const prev = findChildNs(tblPrChangeEl, "tblPr");
      props.propertyChange = {
        revision: rev,
        previousProperties: prev ? parseTableProperties(prev) : undefined
      };
    }
  }

  return props;
}

function parseTableCell(el: XmlElement): TableCell {
  const tcPrEl = findChildNs(el, "tcPr");
  const content: (Paragraph | Table)[] = [];

  for (const child of el.children) {
    if (child.type !== "element") {
      continue;
    }
    const name = child.name.replace(/^w:/, "");
    if (name === "p") {
      content.push(parseParagraph(child));
    } else if (name === "tbl") {
      content.push(parseTable(child));
    }
  }

  let props: TableCellProperties | undefined;
  if (tcPrEl) {
    const p: any = {};
    const wEl = findChildNs(tcPrEl, "tcW");
    if (wEl) {
      p.width = parseTableWidth(wEl);
    }

    const gsEl = findChildNs(tcPrEl, "gridSpan");
    if (gsEl) {
      p.gridSpan = attrInt(gsEl, "val");
    }

    const vmEl = findChildNs(tcPrEl, "vMerge");
    if (vmEl) {
      p.verticalMerge = attrVal(vmEl, "val") ?? "continue";
    }

    const bordersEl = findChildNs(tcPrEl, "tcBorders");
    if (bordersEl) {
      p.borders = parseTableBorders(bordersEl);
    }

    const shdEl = findChildNs(tcPrEl, "shd");
    if (shdEl) {
      p.shading = parseShading(shdEl);
    }

    const vAlignEl = findChildNs(tcPrEl, "vAlign");
    if (vAlignEl) {
      p.verticalAlign = attrVal(vAlignEl, "val");
    }

    if (findChildNs(tcPrEl, "noWrap")) {
      p.noWrap = true;
    }

    const textDirEl = findChildNs(tcPrEl, "textDirection");
    if (textDirEl) {
      p.textDirection = attrVal(textDirEl, "val");
    }

    const marginsEl = findChildNs(tcPrEl, "tcMar");
    if (marginsEl) {
      p.margins = parseTableCellMargins(marginsEl);
    }

    // Conditional formatting
    const cnfEl = findChildNs(tcPrEl, "cnfStyle");
    if (cnfEl) {
      p.cnfStyle = attrVal(cnfEl, "val");
    }

    // Hide cell end-of-cell marker
    if (findChildNs(tcPrEl, "hideMark")) {
      p.hideMark = true;
    }

    // Fit text
    if (findChildNs(tcPrEl, "tcFitText")) {
      p.fitText = true;
    }

    // Cell-level revisions
    const cellInsEl = findChildNs(tcPrEl, "cellIns");
    if (cellInsEl) {
      const rev = parseRevisionInfo(cellInsEl);
      if (rev) {
        p.inserted = { revision: rev };
      }
    }
    const cellDelEl = findChildNs(tcPrEl, "cellDel");
    if (cellDelEl) {
      const rev = parseRevisionInfo(cellDelEl);
      if (rev) {
        p.deleted = { revision: rev };
      }
    }
    const cellMergeEl = findChildNs(tcPrEl, "cellMerge");
    if (cellMergeEl) {
      const vMerge = attrVal(cellMergeEl, "vMerge");
      const rev = parseRevisionInfo(cellMergeEl);
      if (rev && (vMerge === "cont" || vMerge === "rest")) {
        p.cellMerge = { vMerge, revision: rev };
      }
    }

    // tcPrChange
    const tcPrChangeEl = findChildNs(tcPrEl, "tcPrChange");
    if (tcPrChangeEl) {
      const rev = parseRevisionInfo(tcPrChangeEl);
      if (rev) {
        const prev = findChildNs(tcPrChangeEl, "tcPr");
        p.propertyChange = { revision: rev };
        if (prev) {
          // Minimal: previousProperties won't recurse (avoid infinite recursion).
          // Just capture the presence of the change marker here.
        }
      }
    }

    props = p;
  }

  return { properties: props, content };
}

function parseTableRow(el: XmlElement): TableRow {
  const trPrEl = findChildNs(el, "trPr");
  const tblPrExEl = findChildNs(el, "tblPrEx");
  const cells: TableCell[] = [];

  for (const child of el.children) {
    if (child.type === "element" && child.name.replace(/^w:/, "") === "tc") {
      cells.push(parseTableCell(child));
    }
  }

  let props: TableRowProperties | undefined;
  if (trPrEl || tblPrExEl) {
    const p: any = {};
    if (tblPrExEl) {
      p.tblPrEx = parseTableProperties(tblPrExEl);
    }
    if (trPrEl) {
      const heightEl = findChildNs(trPrEl, "trHeight");
      if (heightEl) {
        p.height = {
          value: attrInt(heightEl, "val") ?? 0,
          rule: attrVal(heightEl, "hRule")
        };
      }
      if (findChildNs(trPrEl, "tblHeader")) {
        p.tableHeader = true;
      }
      if (findChildNs(trPrEl, "cantSplit")) {
        p.cantSplit = true;
      }
      if (findChildNs(trPrEl, "hidden")) {
        p.hidden = true;
      }
      const csEl = findChildNs(trPrEl, "tblCellSpacing");
      if (csEl) {
        p.cellSpacing = parseTableWidth(csEl);
      }
      const insEl = findChildNs(trPrEl, "ins");
      if (insEl) {
        const rev = parseRevisionInfo(insEl);
        if (rev) {
          p.inserted = { revision: rev };
        }
      }
      const delEl = findChildNs(trPrEl, "del");
      if (delEl) {
        const rev = parseRevisionInfo(delEl);
        if (rev) {
          p.deleted = { revision: rev };
        }
      }
      const gbEl = findChildNs(trPrEl, "gridBefore");
      if (gbEl) {
        p.gridBefore = attrInt(gbEl, "val");
      }
      const gaEl = findChildNs(trPrEl, "gridAfter");
      if (gaEl) {
        p.gridAfter = attrInt(gaEl, "val");
      }
      const wbEl = findChildNs(trPrEl, "wBefore");
      if (wbEl) {
        p.widthBefore = parseTableWidth(wbEl);
      }
      const waEl = findChildNs(trPrEl, "wAfter");
      if (waEl) {
        p.widthAfter = parseTableWidth(waEl);
      }
      const cnfEl = findChildNs(trPrEl, "cnfStyle");
      if (cnfEl) {
        p.cnfStyle = attrVal(cnfEl, "val");
      }
      const trPrChangeEl = findChildNs(trPrEl, "trPrChange");
      if (trPrChangeEl) {
        const rev = parseRevisionInfo(trPrChangeEl);
        if (rev) {
          const prevTrPr = findChildNs(trPrChangeEl, "trPr");
          p.propertyChange = {
            revision: rev,
            previousProperties: prevTrPr ? parseRowPrInner(prevTrPr) : undefined
          };
        }
      }
    }
    props = p;
  }

  return { properties: props, cells };
}

/** Inner parse for row properties content (used by propertyChange recursion). */
function parseRowPrInner(trPrEl: XmlElement): TableRowProperties {
  const p: any = {};
  const heightEl = findChildNs(trPrEl, "trHeight");
  if (heightEl) {
    p.height = { value: attrInt(heightEl, "val") ?? 0, rule: attrVal(heightEl, "hRule") };
  }
  if (findChildNs(trPrEl, "tblHeader")) {
    p.tableHeader = true;
  }
  if (findChildNs(trPrEl, "cantSplit")) {
    p.cantSplit = true;
  }
  return p;
}

function parseTable(tblEl: XmlElement): Table {
  const tblPrEl = findChildNs(tblEl, "tblPr");
  const gridEl = findChildNs(tblEl, "tblGrid");
  const rows: TableRow[] = [];

  for (const child of tblEl.children) {
    if (child.type === "element" && child.name.replace(/^w:/, "") === "tr") {
      rows.push(parseTableRow(child));
    }
  }

  let columnWidths: number[] | undefined;
  if (gridEl) {
    columnWidths = [];
    for (const col of findChildrenNs(gridEl, "gridCol")) {
      columnWidths.push(parseInt(col.attributes["w:w"] ?? col.attributes["w"] ?? "0", 10));
    }
  }

  return {
    type: "table",
    properties: tblPrEl ? parseTableProperties(tblPrEl) : undefined,
    columnWidths,
    rows
  };
}

// =============================================================================
// Section Properties Parser
// =============================================================================

function parseSectionProperties(sectPrEl: XmlElement): SectionProperties {
  const sect: any = {};

  const pgSzEl = findChildNs(sectPrEl, "pgSz");
  if (pgSzEl) {
    // Per ECMA-376, w:orient defaults to "portrait" when absent
    const orient = attrVal(pgSzEl, "orient");
    sect.pageSize = {
      width: attrInt(pgSzEl, "w"),
      height: attrInt(pgSzEl, "h"),
      orientation: (orient === "landscape" ? "landscape" : "portrait") as any
    };
  }

  const pgMarEl = findChildNs(sectPrEl, "pgMar");
  if (pgMarEl) {
    sect.margins = {
      top: attrInt(pgMarEl, "top"),
      right: attrInt(pgMarEl, "right"),
      bottom: attrInt(pgMarEl, "bottom"),
      left: attrInt(pgMarEl, "left"),
      header: attrInt(pgMarEl, "header"),
      footer: attrInt(pgMarEl, "footer"),
      gutter: attrInt(pgMarEl, "gutter")
    };
  }

  const typeEl = findChildNs(sectPrEl, "type");
  if (typeEl) {
    sect.breakType = attrVal(typeEl, "val");
  }

  const colsEl = findChildNs(sectPrEl, "cols");
  if (colsEl) {
    const cols: any = {};
    cols.space = attrInt(colsEl, "space");
    cols.count = attrInt(colsEl, "num");
    const eqw = attrVal(colsEl, "equalWidth");
    if (eqw !== undefined) {
      cols.equalWidth = eqw === "1";
    }
    const sep = attrVal(colsEl, "sep");
    if (sep === "1" || sep === "true") {
      cols.separator = true;
    }
    const colDefs = findChildrenNs(colsEl, "col");
    if (colDefs.length > 0) {
      cols.columns = colDefs.map(c => ({
        width: attrInt(c, "w") ?? 0,
        space: attrInt(c, "space")
      }));
    }
    sect.columns = cols;
  }

  if (findChildNs(sectPrEl, "titlePg")) {
    sect.titlePage = true;
  }

  const pgNumEl = findChildNs(sectPrEl, "pgNumType");
  if (pgNumEl) {
    sect.pageNumbering = {
      start: attrInt(pgNumEl, "start"),
      format: attrVal(pgNumEl, "fmt")
    };
  }

  // Page borders
  const pgBordersEl = findChildNs(sectPrEl, "pgBorders");
  if (pgBordersEl) {
    const pb: PageBorders = {} as any;
    const p: any = pb;
    for (const side of ["top", "left", "bottom", "right"] as const) {
      const sideEl = findChildNs(pgBordersEl, side);
      if (sideEl) {
        p[side] = parseBorder(sideEl);
      }
    }
    const display = attrVal(pgBordersEl, "display");
    if (display) {
      p.display = display;
    }
    const offsetFrom = attrVal(pgBordersEl, "offsetFrom");
    if (offsetFrom) {
      p.offsetFrom = offsetFrom;
    }
    const zOrder = attrVal(pgBordersEl, "zOrder");
    if (zOrder) {
      p.zOrder = zOrder;
    }
    sect.pageBorders = pb;
  }

  // Vertical alignment
  const vAlignEl = findChildNs(sectPrEl, "vAlign");
  if (vAlignEl) {
    sect.verticalAlign = attrVal(vAlignEl, "val");
  }

  // Text direction
  const textDirEl = findChildNs(sectPrEl, "textDirection");
  if (textDirEl) {
    sect.textDirection = attrVal(textDirEl, "val");
  }

  // Bidi
  const bidiToggle = boolToggle(sectPrEl, "bidi");
  if (bidiToggle !== undefined) {
    sect.bidi = bidiToggle;
  }

  // RTL gutter
  const rtlGutterEl = findChildNs(sectPrEl, "rtlGutter");
  if (rtlGutterEl) {
    sect.rtlGutter = true;
  }

  // Form protection
  const formProtEl = findChildNs(sectPrEl, "formProt");
  if (formProtEl) {
    const v = attrVal(formProtEl, "val");
    sect.formProtection = v === "1" || v === "true";
  }

  // Document grid
  const docGridEl = findChildNs(sectPrEl, "docGrid");
  if (docGridEl) {
    const dg: any = {};
    const linePitch = attrInt(docGridEl, "linePitch");
    if (linePitch !== undefined) {
      dg.linePitch = linePitch;
    }
    const charSpace = attrInt(docGridEl, "charSpace");
    if (charSpace !== undefined) {
      dg.charSpace = charSpace;
    }
    const gridType = attrVal(docGridEl, "type");
    if (gridType) {
      dg.type = gridType;
    }
    sect.docGrid = dg;
  }

  // Line numbers
  const lnNumEl = findChildNs(sectPrEl, "lnNumType");
  if (lnNumEl) {
    const ln: any = {};
    const countBy = attrInt(lnNumEl, "countBy");
    if (countBy !== undefined) {
      ln.countBy = countBy;
    }
    const start = attrInt(lnNumEl, "start");
    if (start !== undefined) {
      ln.start = start;
    }
    const restart = attrVal(lnNumEl, "restart");
    if (restart) {
      ln.restart = restart;
    }
    const distance = attrInt(lnNumEl, "distance");
    if (distance !== undefined) {
      ln.distance = distance;
    }
    sect.lineNumbers = ln;
  }

  // Footnote properties
  const fnPrEl = findChildNs(sectPrEl, "footnotePr");
  if (fnPrEl) {
    sect.footnoteProperties = parseNoteProperties(fnPrEl);
  }

  // Endnote properties
  const enPrEl = findChildNs(sectPrEl, "endnotePr");
  if (enPrEl) {
    sect.endnoteProperties = parseNoteProperties(enPrEl);
  }

  // Headers/Footers refs
  const headerRefs: HeaderFooterRef[] = [];
  for (const hRef of findChildrenNs(sectPrEl, "headerReference")) {
    headerRefs.push({
      type: (attrVal(hRef, "type") ?? "default") as any,
      rId: hRef.attributes["r:id"] ?? ""
    });
  }
  if (headerRefs.length > 0) {
    sect.headers = headerRefs;
  }

  const footerRefs: HeaderFooterRef[] = [];
  for (const fRef of findChildrenNs(sectPrEl, "footerReference")) {
    footerRefs.push({
      type: (attrVal(fRef, "type") ?? "default") as any,
      rId: fRef.attributes["r:id"] ?? ""
    });
  }
  if (footerRefs.length > 0) {
    sect.footers = footerRefs;
  }

  // sectPrChange (track changes for section properties)
  const sectPrChangeEl = findChildNs(sectPrEl, "sectPrChange");
  if (sectPrChangeEl) {
    const rev = parseRevisionInfo(sectPrChangeEl);
    if (rev) {
      const prevSectPrEl = findChildNs(sectPrChangeEl, "sectPr");
      const change: SectionPropertyChange = {
        revision: rev,
        previousProperties: prevSectPrEl ? parseSectionProperties(prevSectPrEl) : undefined
      };
      sect.propertyChange = change;
    }
  }

  return sect;
}

// =============================================================================
// Styles Parser
// =============================================================================

function parseStyles(xmlStr: string): { docDefaults?: DocDefaults; styles: StyleDef[] } {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  let docDefaults: DocDefaults | undefined;
  const styles: StyleDef[] = [];

  const ddEl = findChildNs(root, "docDefaults");
  if (ddEl) {
    const dd: any = {};
    const rPrDefaultEl = findChildNs(ddEl, "rPrDefault");
    if (rPrDefaultEl) {
      const rPrEl = findChildNs(rPrDefaultEl, "rPr");
      if (rPrEl) {
        dd.runProperties = parseRunProperties(rPrEl);
      }
    }
    const pPrDefaultEl = findChildNs(ddEl, "pPrDefault");
    if (pPrDefaultEl) {
      const pPrEl = findChildNs(pPrDefaultEl, "pPr");
      if (pPrEl) {
        dd.paragraphProperties = parseParagraphProperties(pPrEl);
      }
    }
    docDefaults = dd;
  }

  for (const styleEl of findChildrenNs(root, "style")) {
    const s: any = {};
    s.type = attrVal(styleEl, "type");
    s.styleId = attrVal(styleEl, "styleId");
    s.isDefault = attrVal(styleEl, "default") === "1";
    if (attrVal(styleEl, "customStyle") === "1") {
      s.customStyle = true;
    }

    const nameEl = findChildNs(styleEl, "name");
    s.name = nameEl ? (attrVal(nameEl, "val") ?? "") : "";

    const basedOnEl = findChildNs(styleEl, "basedOn");
    if (basedOnEl) {
      s.basedOn = attrVal(basedOnEl, "val");
    }

    const nextEl = findChildNs(styleEl, "next");
    if (nextEl) {
      s.next = attrVal(nextEl, "val");
    }

    const linkEl = findChildNs(styleEl, "link");
    if (linkEl) {
      s.link = attrVal(linkEl, "val");
    }

    const uiPrEl = findChildNs(styleEl, "uiPriority");
    if (uiPrEl) {
      s.uiPriority = attrInt(uiPrEl, "val");
    }

    if (findChildNs(styleEl, "qFormat")) {
      s.qFormat = true;
    }
    if (findChildNs(styleEl, "semiHidden")) {
      s.semiHidden = true;
    }
    if (findChildNs(styleEl, "unhideWhenUsed")) {
      s.unhideWhenUsed = true;
    }
    if (findChildNs(styleEl, "hidden")) {
      s.hidden = true;
    }
    if (findChildNs(styleEl, "locked")) {
      s.locked = true;
    }
    if (findChildNs(styleEl, "autoRedefine")) {
      s.autoRedefine = true;
    }

    const pPrEl = findChildNs(styleEl, "pPr");
    if (pPrEl) {
      s.paragraphProperties = parseParagraphProperties(pPrEl);
    }

    const rPrEl = findChildNs(styleEl, "rPr");
    if (rPrEl) {
      s.runProperties = parseRunProperties(rPrEl);
    }

    // Table properties for table styles
    const tblPrEl = findChildNs(styleEl, "tblPr");
    if (tblPrEl) {
      s.tableProperties = parseTableProperties(tblPrEl);
    }

    // Table style conditional formats
    const tblStylePrs = findChildrenNs(styleEl, "tblStylePr");
    if (tblStylePrs.length > 0) {
      const conditions: TableStyleConditionalFormat[] = [];
      for (const tsp of tblStylePrs) {
        const cond: any = { type: attrVal(tsp, "type") };
        const cpPr = findChildNs(tsp, "pPr");
        if (cpPr) {
          cond.paragraphProperties = parseParagraphProperties(cpPr);
        }
        const crPr = findChildNs(tsp, "rPr");
        if (crPr) {
          cond.runProperties = parseRunProperties(crPr);
        }
        const ctblPr = findChildNs(tsp, "tblPr");
        if (ctblPr) {
          cond.tableProperties = parseTableProperties(ctblPr);
        }
        const ctrPr = findChildNs(tsp, "trPr");
        if (ctrPr) {
          const rp: any = {};
          const hEl = findChildNs(ctrPr, "trHeight");
          if (hEl) {
            rp.height = { value: attrInt(hEl, "val") ?? 0, rule: attrVal(hEl, "hRule") };
          }
          cond.rowProperties = rp;
        }
        const ctcPr = findChildNs(tsp, "tcPr");
        if (ctcPr) {
          const cp: any = {};
          const bEl = findChildNs(ctcPr, "tcBorders");
          if (bEl) {
            cp.borders = parseTableBorders(bEl);
          }
          const shd = findChildNs(ctcPr, "shd");
          if (shd) {
            cp.shading = parseShading(shd);
          }
          cond.cellProperties = cp;
        }
        conditions.push(cond);
      }
      s.tableStyleConditions = conditions;
    }

    styles.push(s);
  }

  return { docDefaults, styles };
}

// =============================================================================
// Numbering Parser
// =============================================================================

function parseNumberingXml(xmlStr: string): {
  abstractNums: AbstractNumbering[];
  instances: NumberingInstance[];
  numPicBullets: NumPicBullet[];
} {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const abstractNums: AbstractNumbering[] = [];
  const instances: NumberingInstance[] = [];
  const numPicBullets: NumPicBullet[] = [];

  // Parse picture bullets
  for (const pbEl of findChildrenNs(root, "numPicBullet")) {
    const id = attrInt(pbEl, "numPicBulletId");
    if (id === undefined) {
      continue;
    }
    const pb: any = { id };
    // Try to extract VML shape info
    const pictEl = findChildNs(pbEl, "pict");
    if (pictEl) {
      // Preserve raw VML for complete fidelity
      let rawVml = "";
      for (const child of pictEl.children) {
        if (child.type === "element") {
          rawVml += serializeElement(child);
        }
      }
      if (rawVml) {
        pb.rawVmlXml = rawVml;
      }
      // Extract rId from v:imagedata
      const shapeEl = findChild(pictEl, "v:shape");
      if (shapeEl) {
        const imgDataEl = findChild(shapeEl, "v:imagedata");
        if (imgDataEl) {
          const rId = imgDataEl.attributes["r:id"] ?? imgDataEl.attributes["r:pict"];
          if (rId) {
            pb.rId = rId;
          }
        }
        // Extract width/height from style
        const style = shapeEl.attributes["style"];
        if (style) {
          const wMatch = /width:([\d.]+)pt/i.exec(style);
          const hMatch = /height:([\d.]+)pt/i.exec(style);
          if (wMatch) {
            pb.width = Math.round(parseFloat(wMatch[1]) * 12700);
          }
          if (hMatch) {
            pb.height = Math.round(parseFloat(hMatch[1]) * 12700);
          }
        }
      }
    }
    numPicBullets.push(pb);
  }

  for (const absEl of findChildrenNs(root, "abstractNum")) {
    const levels: NumberingLevel[] = [];
    for (const lvlEl of findChildrenNs(absEl, "lvl")) {
      levels.push(parseLevel(lvlEl));
    }

    const abs: any = {
      abstractNumId: attrInt(absEl, "abstractNumId") ?? 0,
      levels
    };
    const mltEl = findChildNs(absEl, "multiLevelType");
    if (mltEl) {
      abs.multiLevelType = attrVal(mltEl, "val");
    }
    const numStyleLinkEl = findChildNs(absEl, "numStyleLink");
    if (numStyleLinkEl) {
      abs.numStyleLink = attrVal(numStyleLinkEl, "val");
    }
    const styleLinkEl = findChildNs(absEl, "styleLink");
    if (styleLinkEl) {
      abs.styleLink = attrVal(styleLinkEl, "val");
    }
    abstractNums.push(abs);
  }

  for (const numEl of findChildrenNs(root, "num")) {
    const absIdEl = findChildNs(numEl, "abstractNumId");
    const overrides: LevelOverride[] = [];
    for (const ovEl of findChildrenNs(numEl, "lvlOverride")) {
      const ov: any = { level: attrInt(ovEl, "ilvl") ?? 0 };
      const startOvEl = findChildNs(ovEl, "startOverride");
      if (startOvEl) {
        ov.startOverride = attrInt(startOvEl, "val");
      }
      // Level def override: parse full level definition
      const lvlEl = findChildNs(ovEl, "lvl");
      if (lvlEl) {
        ov.levelDef = parseLevel(lvlEl);
        // Inherit level index from parent if not specified
        if ((ov.levelDef as any).level === undefined) {
          (ov.levelDef as any).level = ov.level;
        }
      }
      overrides.push(ov);
    }
    instances.push({
      numId: attrInt(numEl, "numId") ?? 0,
      abstractNumId: absIdEl ? (attrInt(absIdEl, "val") ?? 0) : 0,
      overrides: overrides.length > 0 ? overrides : undefined
    });
  }

  return { abstractNums, instances, numPicBullets };
}

/** Parse a w:lvl element into a NumberingLevel (shared by abstractNum and lvlOverride). */
function parseLevel(lvlEl: XmlElement): NumberingLevel {
  const level: any = { level: attrInt(lvlEl, "ilvl") ?? 0 };

  const startEl = findChildNs(lvlEl, "start");
  if (startEl) {
    level.start = attrInt(startEl, "val");
  }
  const fmtEl = findChildNs(lvlEl, "numFmt");
  if (fmtEl) {
    level.format = attrVal(fmtEl, "val");
  }
  const textEl = findChildNs(lvlEl, "lvlText");
  if (textEl) {
    level.text = attrVal(textEl, "val") ?? "";
  }
  const pStyleEl = findChildNs(lvlEl, "pStyle");
  if (pStyleEl) {
    level.paragraphStyle = attrVal(pStyleEl, "val");
  }
  const jcEl = findChildNs(lvlEl, "lvlJc");
  if (jcEl) {
    level.justification = attrVal(jcEl, "val");
  }
  const pPrEl = findChildNs(lvlEl, "pPr");
  if (pPrEl) {
    level.paragraphProperties = parseParagraphProperties(pPrEl);
  }
  const rPrEl = findChildNs(lvlEl, "rPr");
  if (rPrEl) {
    level.runProperties = parseRunProperties(rPrEl);
  }
  const suffEl = findChildNs(lvlEl, "suff");
  if (suffEl) {
    level.suffix = attrVal(suffEl, "val");
  }
  if (findChildNs(lvlEl, "isLgl")) {
    level.isLegalNumberingStyle = true;
  }
  const lvlRestartEl = findChildNs(lvlEl, "lvlRestart");
  if (lvlRestartEl) {
    level.restartAfterLevel = attrInt(lvlRestartEl, "val");
  }
  const picBulletEl = findChildNs(lvlEl, "lvlPicBulletId");
  if (picBulletEl) {
    level.picBulletId = attrInt(picBulletEl, "val");
  }
  return level;
}

// =============================================================================
// Footnotes/Endnotes Parser
// =============================================================================

function parseNotesXml(
  xmlStr: string,
  elementName: string
): { id: number; type?: NoteType; content: Paragraph[] }[] {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const notes: { id: number; type?: NoteType; content: Paragraph[] }[] = [];

  for (const noteEl of findChildrenNs(root, elementName)) {
    const id = attrInt(noteEl, "id");
    const type = attrVal(noteEl, "type");
    // Skip auto-generated separator entries (default IDs -1 and 0)
    // Real separators/continuationSeparators are regenerated by the writer.
    if (type === "separator" || type === "continuationSeparator") {
      continue;
    }
    if (id === undefined) {
      continue;
    }

    const content: Paragraph[] = [];
    for (const child of noteEl.children) {
      if (child.type === "element" && child.name.replace(/^w:/, "") === "p") {
        content.push(parseParagraph(child));
      }
    }

    const note: { id: number; type?: NoteType; content: Paragraph[] } = { id, content };
    if (type === "continuationNotice" || type === "normal") {
      note.type = type;
    }
    notes.push(note);
  }

  return notes;
}

// =============================================================================
// Header/Footer Parser
// =============================================================================

function parseHeaderFooterXml(xmlStr: string): HeaderFooterContent {
  return parseHeaderFooterRoot(parseXml(xmlStr).root);
}

function parseHeaderFooterRoot(root: XmlElement): HeaderFooterContent {
  const children: (Paragraph | Table)[] = [];
  for (const child of root.children) {
    if (child.type !== "element") {
      continue;
    }
    const name = child.name.replace(/^w:/, "");
    if (name === "p") {
      children.push(parseParagraph(child));
    } else if (name === "tbl") {
      children.push(parseTable(child));
    }
  }
  return { children };
}

/** Detect watermark from a header's parsed XML root element. */
function detectWatermarkFromRoot(root: XmlElement): Watermark | undefined {
  // Look for VML shape with id containing "WaterMark"
  for (const pEl of root.children) {
    if (pEl.type !== "element") {
      continue;
    }
    for (const rEl of pEl.children) {
      if (rEl.type !== "element") {
        continue;
      }
      // Look for w:pict or w:r > w:pict
      const pictEls: XmlElement[] = [];
      const rName = rEl.name.replace(/^w:/, "");
      if (rName === "pict") {
        pictEls.push(rEl);
      } else if (rName === "r") {
        for (const rc of rEl.children) {
          if (rc.type === "element" && rc.name.replace(/^w:/, "") === "pict") {
            pictEls.push(rc);
          }
        }
      }
      for (const pictEl of pictEls) {
        for (const shapeEl of pictEl.children) {
          if (shapeEl.type !== "element") {
            continue;
          }
          const shapeId = shapeEl.attributes["id"] ?? "";
          if (!shapeId.toLowerCase().includes("watermark")) {
            continue;
          }
          // Found watermark shape
          const shapeType = shapeEl.attributes["type"] ?? "";
          if (shapeType.includes("136")) {
            // WordArt text watermark (shapetype 136)
            return parseTextWatermark(shapeEl);
          }
          // Check for image watermark (has v:imagedata)
          const imgData = findChild(shapeEl, "v:imagedata");
          if (imgData) {
            return parseImageWatermark(shapeEl, imgData);
          }
        }
      }
    }
  }
  return undefined;
}

function parseTextWatermark(shapeEl: XmlElement): TextWatermark {
  const fillColor = shapeEl.attributes["fillcolor"] ?? "#C0C0C0";
  const color = fillColor.replace(/^#/, "");

  // Parse rotation from style
  const style = shapeEl.attributes["style"] ?? "";
  let rotation = -45;
  const rotMatch = style.match(/rotation:\s*(-?\d+)/);
  if (rotMatch) {
    rotation = parseInt(rotMatch[1], 10);
  }

  // Get opacity from v:fill
  const fillEl = findChild(shapeEl, "v:fill");
  const opacity = fillEl?.attributes["opacity"] ?? ".5";
  const semiTransparent = opacity !== "1";

  // Get text and font from v:textpath
  const textpathEl = findChild(shapeEl, "v:textpath");
  const text = textpathEl?.attributes["string"] ?? "";
  const tpStyle = textpathEl?.attributes["style"] ?? "";
  let font: string | undefined;
  let fontSize: number | undefined;
  const fontMatch = tpStyle.match(/font-family:\s*"?([^";]+)"?/);
  if (fontMatch) {
    font = fontMatch[1].replace(/&quot;/g, "");
  }
  const sizeMatch = tpStyle.match(/font-size:\s*(\d+(?:\.\d+)?)\s*pt/);
  if (sizeMatch) {
    fontSize = Math.round(parseFloat(sizeMatch[1]) * 2); // convert pt to half-points
  }

  return {
    type: "text",
    text,
    font,
    fontSize,
    color,
    semiTransparent,
    rotation
  };
}

function parseImageWatermark(shapeEl: XmlElement, imgDataEl: XmlElement): ImageWatermark {
  const rId = imgDataEl.attributes["r:id"] ?? "";
  const gain = imgDataEl.attributes["gain"] ?? "";
  const washout = gain.startsWith("19661") || gain === "";

  return {
    type: "image",
    rId,
    washout
  };
}

// =============================================================================
// Comments Parser
// =============================================================================

function parseCommentsXml(xmlStr: string): CommentDef[] {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const comments: CommentDef[] = [];

  for (const commentEl of findChildrenNs(root, "comment")) {
    const id = attrInt(commentEl, "id");
    const author = attrVal(commentEl, "author");
    if (id === undefined || !author) {
      continue;
    }

    const content: Paragraph[] = [];
    for (const child of commentEl.children) {
      if (child.type === "element" && child.name.replace(/^w:/, "") === "p") {
        content.push(parseParagraph(child));
      }
    }

    const comment: any = { id, author, content };
    const date = attrVal(commentEl, "date");
    if (date) {
      comment.date = date;
    }
    const initials = attrVal(commentEl, "initials");
    if (initials) {
      comment.initials = initials;
    }
    comments.push(comment);
  }

  return comments;
}

/** Parse word/commentsExtended.xml — map paraId → { done, parentId }. */
function parseCommentsExtendedXml(
  xmlStr: string
): Map<string, { done?: boolean; parentId?: string }> {
  const map = new Map<string, { done?: boolean; parentId?: string }>();
  const doc = parseXml(xmlStr);
  const root = doc.root;
  for (const child of root.children) {
    if (child.type !== "element") {
      continue;
    }
    // w15:commentEx
    const name = child.name;
    if (!name.endsWith("commentEx")) {
      continue;
    }
    const paraId = child.attributes["w15:paraId"] ?? child.attributes["paraId"];
    if (!paraId) {
      continue;
    }
    const entry: { done?: boolean; parentId?: string } = {};
    const done = child.attributes["w15:done"] ?? child.attributes["done"];
    if (done === "1" || done === "true") {
      entry.done = true;
    } else if (done === "0" || done === "false") {
      entry.done = false;
    }
    const pid = child.attributes["w15:paraIdParent"] ?? child.attributes["paraIdParent"];
    if (pid) {
      entry.parentId = pid;
    }
    map.set(paraId, entry);
  }
  return map;
}

// =============================================================================
// Core Properties Parser
// =============================================================================

function parseCoreProps(xmlStr: string): CoreProperties {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const props: any = {};

  const fields: [string, string][] = [
    ["dc:title", "title"],
    ["dc:subject", "subject"],
    ["dc:creator", "creator"],
    ["dc:description", "description"],
    ["cp:keywords", "keywords"],
    ["cp:lastModifiedBy", "lastModifiedBy"],
    ["cp:revision", "revision"],
    ["cp:category", "category"]
  ];

  for (const [tag, prop] of fields) {
    const el = findChild(root, tag);
    if (el) {
      const val = textContent(el);
      if (val) {
        props[prop] = val;
      }
    }
  }

  const createdEl = findChild(root, "dcterms:created");
  if (createdEl) {
    const val = textContent(createdEl);
    if (val) {
      props.created = new Date(val);
    }
  }

  const modifiedEl = findChild(root, "dcterms:modified");
  if (modifiedEl) {
    const val = textContent(modifiedEl);
    if (val) {
      props.modified = new Date(val);
    }
  }

  return props;
}

// =============================================================================
// App Properties Parser
// =============================================================================

function parseAppProps(xmlStr: string): AppProperties {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const props: any = {};

  const strFields = ["Application", "AppVersion", "Company", "Manager"];
  const intFields = ["Pages", "Words", "Characters", "Lines", "Paragraphs"];

  for (const field of strFields) {
    const el = findChild(root, field);
    if (el) {
      const val = textContent(el);
      if (val) {
        props[field.charAt(0).toLowerCase() + field.slice(1)] = val;
      }
    }
  }

  for (const field of intFields) {
    const el = findChild(root, field);
    if (el) {
      const val = textContent(el);
      if (val) {
        props[field.charAt(0).toLowerCase() + field.slice(1)] = parseInt(val, 10);
      }
    }
  }

  return props;
}

// =============================================================================
// Theme Parser
// =============================================================================

const THEME_COLOR_NAMES: ThemeColorName[] = [
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

function parseThemeXml(xmlStr: string): DocumentTheme {
  const doc = parseXml(xmlStr);
  const root = doc.root;

  // Find a:themeElements
  const themeElements = findChild(root, "a:themeElements") ?? findChildNs(root, "themeElements");

  const defaultScheme = {
    name: "Office",
    colors: {
      dk1: "000000",
      lt1: "FFFFFF",
      dk2: "44546A",
      lt2: "E7E6E6",
      accent1: "4472C4",
      accent2: "ED7D31",
      accent3: "A5A5A5",
      accent4: "FFC000",
      accent5: "5B9BD5",
      accent6: "70AD47",
      hlink: "0563C1",
      folHlink: "954F72"
    } as Record<ThemeColorName, string>
  };
  const defaultFontScheme = { name: "Office", majorFont: "Calibri Light", minorFont: "Calibri" };

  if (!themeElements) {
    return {
      name: root.attributes["name"],
      colorScheme: defaultScheme,
      fontScheme: defaultFontScheme
    };
  }

  // Parse color scheme
  const clrSchemeEl =
    findChild(themeElements, "a:clrScheme") ?? findChildNs(themeElements, "clrScheme");
  const colorScheme = { ...defaultScheme };
  if (clrSchemeEl) {
    colorScheme.name = clrSchemeEl.attributes["name"] ?? "Office";
    for (const colorName of THEME_COLOR_NAMES) {
      const colorEl =
        findChild(clrSchemeEl, `a:${colorName}`) ?? findChildNs(clrSchemeEl, colorName);
      if (colorEl) {
        // Color can be sysClr (with lastClr) or srgbClr (with val)
        const srgb = findChild(colorEl, "a:srgbClr") ?? findChildNs(colorEl, "srgbClr");
        if (srgb) {
          const val = srgb.attributes["val"];
          if (val) {
            colorScheme.colors[colorName] = val;
          }
        } else {
          const sys = findChild(colorEl, "a:sysClr") ?? findChildNs(colorEl, "sysClr");
          if (sys) {
            const lastClr = sys.attributes["lastClr"];
            if (lastClr) {
              colorScheme.colors[colorName] = lastClr;
            }
          }
        }
      }
    }
  }

  // Parse font scheme
  const fontSchemeEl =
    findChild(themeElements, "a:fontScheme") ?? findChildNs(themeElements, "fontScheme");
  const fontScheme: any = { ...defaultFontScheme };
  if (fontSchemeEl) {
    fontScheme.name = fontSchemeEl.attributes["name"] ?? "Office";
    const majorEl =
      findChild(fontSchemeEl, "a:majorFont") ?? findChildNs(fontSchemeEl, "majorFont");
    if (majorEl) {
      const major = parseThemeFont(majorEl);
      fontScheme.major = major;
      if (major.latin) {
        fontScheme.majorFont = major.latin;
      }
    }
    const minorEl =
      findChild(fontSchemeEl, "a:minorFont") ?? findChildNs(fontSchemeEl, "minorFont");
    if (minorEl) {
      const minor = parseThemeFont(minorEl);
      fontScheme.minor = minor;
      if (minor.latin) {
        fontScheme.minorFont = minor.latin;
      }
    }
  }

  // Parse format scheme (preserve raw XML of its children for round-trip)
  const fmtSchemeEl =
    findChild(themeElements, "a:fmtScheme") ?? findChildNs(themeElements, "fmtScheme");
  let formatScheme: any;
  if (fmtSchemeEl) {
    let rawXml = "";
    for (const child of fmtSchemeEl.children) {
      if (child.type === "element") {
        rawXml += serializeElement(child);
      }
    }
    formatScheme = {
      name: fmtSchemeEl.attributes["name"] ?? "Office",
      rawXml: rawXml || undefined
    };
  }

  // Preserve extLst (theme extensions) as raw XML
  let extLstXml: string | undefined;
  const extLstEl = findChild(root, "a:extLst") ?? findChildNs(root, "extLst");
  if (extLstEl) {
    extLstXml = serializeElement(extLstEl);
  }

  return {
    name: root.attributes["name"],
    colorScheme,
    fontScheme,
    formatScheme,
    extLstXml
  };
}

/** Parse a theme font (a:majorFont or a:minorFont). */
function parseThemeFont(el: XmlElement): any {
  const font: any = {};
  const latin = findChild(el, "a:latin") ?? findChildNs(el, "latin");
  if (latin?.attributes["typeface"]) {
    font.latin = latin.attributes["typeface"];
  }
  const ea = findChild(el, "a:ea") ?? findChildNs(el, "ea");
  if (ea?.attributes["typeface"]) {
    font.eastAsia = ea.attributes["typeface"];
  }
  const cs = findChild(el, "a:cs") ?? findChildNs(el, "cs");
  if (cs?.attributes["typeface"]) {
    font.complexScript = cs.attributes["typeface"];
  }
  // Supplemental fonts (a:font script="..." typeface="...")
  const supplementalFonts: Record<string, string> = {};
  for (const child of el.children) {
    if (child.type === "element" && (child.name === "a:font" || child.name === "font")) {
      const script = child.attributes["script"];
      const typeface = child.attributes["typeface"];
      if (script && typeface) {
        supplementalFonts[script] = typeface;
      }
    }
  }
  if (Object.keys(supplementalFonts).length > 0) {
    font.supplementalFonts = supplementalFonts;
  }
  return font;
}

// =============================================================================
// Settings Parser
// =============================================================================

/** Parse word/webSettings.xml. */
function parseWebSettings(xmlStr: string): WebSettings {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const ws: any = {};

  const ofbEl = findChildNs(root, "optimizeForBrowser");
  if (ofbEl) {
    const ofb: any = {};
    const target = attrVal(ofbEl, "target");
    if (target) {
      ofb.target = target;
    }
    const mv = attrInt(ofbEl, "majorVersion");
    if (mv !== undefined) {
      ofb.majorVersion = mv;
    }
    ws.optimizeForBrowser = ofb;
  }
  if (findChildNs(root, "allowPNG")) {
    ws.allowPng = true;
  }
  if (findChildNs(root, "relyOnVML")) {
    ws.relyOnVml = true;
  }
  if (findChildNs(root, "doNotSaveAsSingleFile")) {
    ws.doNotSaveAsSingleFile = true;
  }
  if (findChildNs(root, "doNotOrganizeInFolder")) {
    ws.doNotOrganizeInFolder = true;
  }
  if (findChildNs(root, "useTargetMachineType")) {
    ws.useTargetMachineType = true;
  }
  return ws;
}

/** Parse word/people.xml. */
function parsePeople(xmlStr: string): PersonInfo[] {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const people: PersonInfo[] = [];
  for (const personEl of root.children) {
    if (personEl.type !== "element") {
      continue;
    }
    const author = personEl.attributes["w15:author"] ?? personEl.attributes["author"];
    if (!author) {
      continue;
    }
    const info: any = { author };
    // presenceInfo
    for (const child of personEl.children) {
      if (child.type === "element" && child.name.endsWith("presenceInfo")) {
        const pi: any = {};
        const providerId = child.attributes["w15:providerId"] ?? child.attributes["providerId"];
        if (providerId) {
          pi.providerId = providerId;
        }
        const userId = child.attributes["w15:userId"] ?? child.attributes["userId"];
        if (userId) {
          pi.userId = userId;
        }
        if (Object.keys(pi).length > 0) {
          info.presenceInfo = pi;
        }
        break;
      }
    }
    people.push(info);
  }
  return people;
}

function parseSettingsXml(xmlStr: string): DocumentSettings {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const settings: any = {};

  const zoomEl = findChildNs(root, "zoom");
  if (zoomEl) {
    settings.zoom = attrInt(zoomEl, "percent");
  }

  const tabEl = findChildNs(root, "defaultTabStop");
  if (tabEl) {
    settings.defaultTabStop = attrInt(tabEl, "val");
  }

  const csControlEl = findChildNs(root, "characterSpacingControl");
  if (csControlEl) {
    const v = attrVal(csControlEl, "val");
    if (
      v === "doNotCompress" ||
      v === "compressPunctuation" ||
      v === "compressPunctuationAndJapaneseKana"
    ) {
      settings.characterSpacingControl = v;
    }
  }

  // Extended settings
  if (findChildNs(root, "doNotTrackMoves")) {
    settings.doNotTrackMoves = true;
  }
  if (findChildNs(root, "doNotTrackFormatting")) {
    settings.doNotTrackFormatting = true;
  }
  if (findChildNs(root, "doNotDemoteNonCombiningChars")) {
    settings.doNotDemoteAsianTextFirstLine = true;
  }
  const ssFontsEl = findChildNs(root, "saveSubsetFonts");
  if (ssFontsEl) {
    const v = attrVal(ssFontsEl, "val");
    settings.saveSubsetFonts = v !== "0" && v !== "false";
  }
  if (findChildNs(root, "noPunctuationKerning")) {
    settings.noPunctuationKerning = true;
  }
  if (findChildNs(root, "bordersDoNotSurroundHeader")) {
    settings.bordersDoNotSurroundHeader = true;
  }
  if (findChildNs(root, "bordersDoNotSurroundFooter")) {
    settings.bordersDoNotSurroundFooter = true;
  }
  const clickStyleEl = findChildNs(root, "clickAndTypeStyle");
  if (clickStyleEl) {
    settings.clickAndTypeStyle = attrVal(clickStyleEl, "val");
  }
  const spfEl = findChildNs(root, "stylePaneFormatFilter");
  if (spfEl) {
    settings.stylePaneFormatFilter = attrVal(spfEl, "val");
  }
  const spsEl = findChildNs(root, "stylePaneSortMethod");
  if (spsEl) {
    settings.stylePaneSortMethod = attrVal(spsEl, "val");
  }
  const tflEl = findChildNs(root, "themeFontLang");
  if (tflEl) {
    const tfl: any = {};
    const v = attrVal(tflEl, "val");
    if (v) {
      tfl.val = v;
    }
    const ea = attrVal(tflEl, "eastAsia");
    if (ea) {
      tfl.eastAsia = ea;
    }
    const bd = attrVal(tflEl, "bidi");
    if (bd) {
      tfl.bidi = bd;
    }
    if (Object.keys(tfl).length > 0) {
      settings.themeFontLang = tfl;
    }
  }
  const dsEl = findChildNs(root, "decimalSymbol");
  if (dsEl) {
    settings.decimalSymbol = attrVal(dsEl, "val");
  }
  const lsEl = findChildNs(root, "listSeparator");
  if (lsEl) {
    settings.listSeparator = attrVal(lsEl, "val");
  }

  // RSID list
  const rsidsEl = findChildNs(root, "rsids");
  if (rsidsEl) {
    const rsids: any = {};
    const rootEl = findChildNs(rsidsEl, "rsidRoot");
    if (rootEl) {
      rsids.rsidRoot = attrVal(rootEl, "val");
    }
    const rsidList: string[] = [];
    for (const rsidEl of findChildrenNs(rsidsEl, "rsid")) {
      const v = attrVal(rsidEl, "val");
      if (v) {
        rsidList.push(v);
      }
    }
    if (rsidList.length > 0) {
      rsids.rsid = rsidList;
    }
    if (Object.keys(rsids).length > 0) {
      settings.rsids = rsids;
    }
  }

  if (findChildNs(root, "evenAndOddHeaders")) {
    settings.evenAndOddHeaders = true;
  }

  if (findChildNs(root, "trackRevisions")) {
    settings.trackRevisions = true;
  }

  if (findChildNs(root, "mirrorMargins")) {
    settings.mirrorMargins = true;
  }

  if (findChildNs(root, "gutterAtTop")) {
    settings.gutterAtTop = true;
  }

  if (findChildNs(root, "displayBackgroundShape")) {
    settings.displayBackgroundShape = true;
  }

  if (findChildNs(root, "updateFields")) {
    settings.updateFieldsOnOpen = true;
  }

  // Hyphenation
  const autoHyphEl = findChildNs(root, "autoHyphenation");
  if (autoHyphEl) {
    settings.autoHyphenation = true;
    const hyph: any = { autoHyphenation: true };
    const hzEl = findChildNs(root, "hyphenationZone");
    if (hzEl) {
      hyph.hyphenationZone = attrInt(hzEl, "val");
    }
    const chlEl = findChildNs(root, "consecutiveHyphenLimit");
    if (chlEl) {
      hyph.consecutiveHyphenLimit = attrInt(chlEl, "val");
    }
    if (findChildNs(root, "doNotHyphenateCaps")) {
      hyph.doNotHyphenateCaps = true;
    }
    settings.hyphenation = hyph;
  }

  // Document protection
  const protEl = findChildNs(root, "documentProtection");
  if (protEl) {
    settings.documentProtection = {
      type: attrVal(protEl, "edit") ?? "none",
      enforcement: attrVal(protEl, "enforcement") === "1"
    };
  }

  const compatEl = findChildNs(root, "compat");
  if (compatEl) {
    const compatSettings: any[] = [];
    const compatFlags: any[] = [];
    for (const csEl of compatEl.children) {
      if (csEl.type !== "element") {
        continue;
      }
      const localName = csEl.name.replace(/^w:/, "");
      if (localName === "compatSetting") {
        const name = attrVal(csEl, "name");
        const uri = attrVal(csEl, "uri");
        const val = attrVal(csEl, "val");
        if (name === "compatibilityMode" && val !== undefined) {
          settings.compatibilityMode = parseInt(val, 10);
        } else if (name !== undefined && uri !== undefined && val !== undefined) {
          compatSettings.push({ name, uri, val });
        }
      } else {
        // Legacy compat flags (w:useFELayout, w:balanceSingleByteDoubleByteWidth, etc.)
        compatFlags.push({ name: localName, val: attrVal(csEl, "val") });
      }
    }
    if (compatSettings.length > 0) {
      settings.compatSettings = compatSettings;
    }
    if (compatFlags.length > 0) {
      settings.compatFlags = compatFlags;
    }
  }

  // Mail merge settings (preserve as raw XML)
  const mailMergeEl = findChildNs(root, "mailMerge");
  if (mailMergeEl) {
    settings.mailMergeRawXml = serializeElement(mailMergeEl);
  }

  // Write protection
  const writeProtectionEl = findChildNs(root, "writeProtection");
  if (writeProtectionEl) {
    const wp: any = {};
    const recommended = attrVal(writeProtectionEl, "recommended");
    if (recommended === "1" || recommended === "true") {
      wp.recommended = true;
    }
    const algName = attrVal(writeProtectionEl, "algorithmName");
    if (algName) {
      wp.algorithmName = algName;
    }
    const hashValue = attrVal(writeProtectionEl, "hashValue");
    if (hashValue) {
      wp.hashValue = hashValue;
    }
    const saltValue = attrVal(writeProtectionEl, "saltValue");
    if (saltValue) {
      wp.saltValue = saltValue;
    }
    const spinCount = attrInt(writeProtectionEl, "spinCount");
    if (spinCount !== undefined) {
      wp.spinCount = spinCount;
    }
    settings.writeProtection = wp;
  }

  // Document variables
  const docVarsEl = findChildNs(root, "docVars");
  if (docVarsEl) {
    const vars = new Map<string, string>();
    for (const dvEl of findChildrenNs(docVarsEl, "docVar")) {
      const name = attrVal(dvEl, "name");
      const val = attrVal(dvEl, "val");
      if (name !== undefined && val !== undefined) {
        vars.set(name, val);
      }
    }
    if (vars.size > 0) {
      settings.docVars = vars;
    }
  }

  // Footnote/endnote properties at document level
  const fnPrEl = findChildNs(root, "footnotePr");
  if (fnPrEl) {
    const fnProps = parseNoteProperties(fnPrEl);
    if (fnProps) {
      settings.footnoteProperties = fnProps;
    }
  }
  const enPrEl = findChildNs(root, "endnotePr");
  if (enPrEl) {
    const enProps = parseNoteProperties(enPrEl);
    if (enProps) {
      settings.endnoteProperties = enProps;
    }
  }

  return settings;
}

// =============================================================================
// Custom Properties Parser
// =============================================================================

function parseCustomPropsXml(xmlStr: string): CustomProperty[] {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const props: CustomProperty[] = [];

  for (const propEl of root.children) {
    if (propEl.type !== "element" || propEl.name !== "property") {
      continue;
    }
    const name = propEl.attributes["name"];
    if (!name) {
      continue;
    }

    let value: CustomPropertyValue | undefined;
    for (const child of propEl.children) {
      if (child.type !== "element") {
        continue;
      }
      const tn = child.name;
      const tv = textContent(child);
      if (tn === "vt:lpwstr") {
        value = { type: "string", value: tv };
      } else if (tn === "vt:i4") {
        value = { type: "number", value: parseInt(tv, 10) };
      } else if (tn === "vt:r8") {
        value = { type: "number", value: parseFloat(tv) };
      } else if (tn === "vt:bool") {
        value = { type: "boolean", value: tv === "true" };
      } else if (tn === "vt:filetime") {
        value = { type: "date", value: new Date(tv) };
      }
    }

    if (value) {
      props.push({ name, value });
    }
  }

  return props;
}

// =============================================================================
// Font Table Parser
// =============================================================================

function parseFontTableXml(xmlStr: string): FontDef[] {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const fonts: FontDef[] = [];

  for (const fontEl of findChildrenNs(root, "font")) {
    const f: any = { name: attrVal(fontEl, "name") ?? "" };
    const p1 = findChildNs(fontEl, "panose1");
    if (p1) {
      f.panose1 = attrVal(p1, "val");
    }
    const cs = findChildNs(fontEl, "charset");
    if (cs) {
      f.charset = attrVal(cs, "val");
    }
    const fam = findChildNs(fontEl, "family");
    if (fam) {
      f.family = attrVal(fam, "val");
    }
    const pitch = findChildNs(fontEl, "pitch");
    if (pitch) {
      f.pitch = attrVal(pitch, "val");
    }
    // Signature
    const sigEl = findChildNs(fontEl, "sig");
    if (sigEl) {
      const sig: any = {};
      for (const key of ["usb0", "usb1", "usb2", "usb3", "csb0", "csb1"]) {
        const v = attrVal(sigEl, key);
        if (v !== undefined) {
          sig[key] = v;
        }
      }
      if (Object.keys(sig).length > 0) {
        f.sig = sig;
      }
    }
    // Embedded fonts
    for (const [tag, rIdKey, keyKey] of [
      ["embedRegular", "embedRegular", "embedRegularKey"],
      ["embedBold", "embedBold", "embedBoldKey"],
      ["embedItalic", "embedItalic", "embedItalicKey"],
      ["embedBoldItalic", "embedBoldItalic", "embedBoldItalicKey"]
    ] as const) {
      const el = findChildNs(fontEl, tag);
      if (el) {
        const rId = el.attributes["r:id"] ?? el.attributes["id"];
        if (rId) {
          f[rIdKey] = rId;
          const fontKey = attrVal(el, "fontKey");
          if (fontKey) {
            f[keyKey] = fontKey;
          }
        }
      }
    }
    fonts.push(f);
  }

  return fonts;
}

// =============================================================================
// Relationships Parser
// =============================================================================

interface ParsedRelationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

function parseRelationships(xmlStr: string): ParsedRelationship[] {
  const doc = parseXml(xmlStr);
  const rels: ParsedRelationship[] = [];

  for (const child of doc.root.children) {
    if (child.type === "element" && child.name === "Relationship") {
      rels.push({
        id: child.attributes["Id"] ?? "",
        type: child.attributes["Type"] ?? "",
        target: child.attributes["Target"] ?? "",
        targetMode: child.attributes["TargetMode"]
      });
    }
  }

  return rels;
}

// =============================================================================
// Main Document Parser
// =============================================================================

/** Recursively extract floating images, drawing shapes, and opaque drawings from an element tree. */
function extractFloatingContent(
  el: XmlElement,
  images: FloatingImage[],
  shapes: DrawingShape[],
  opaqueDrawings: OpaqueDrawing[]
): void {
  for (const child of el.children) {
    if (child.type !== "element") {
      continue;
    }
    if (child.name === "wp:anchor") {
      // Check if this is a pic (image) or wsp (shape)
      const graphicEl = findChild(child, "a:graphic");
      const graphicDataEl = graphicEl ? findChild(graphicEl, "a:graphicData") : undefined;
      const wspEl = graphicDataEl
        ? (findChild(graphicDataEl, "wps:wsp") ?? findChildNs(graphicDataEl, "wsp"))
        : undefined;
      if (wspEl) {
        const shape = parseDrawingShape(child, wspEl);
        if (shape) {
          shapes.push(shape);
        }
      } else {
        const fi = parseFloatingImage(child);
        if (fi) {
          images.push(fi);
        } else {
          // Unknown anchor content (chart, diagram, etc.) — preserve as opaque
          const drawingEl = findDrawingParent(child);
          if (drawingEl) {
            const rids = new Set<string>();
            collectRIds(drawingEl, rids);
            opaqueDrawings.push({
              type: "opaqueDrawing",
              rawXml: serializeElement(drawingEl),
              referencedRIds: [...rids]
            });
          }
        }
      }
    } else if (child.name === "wp:inline") {
      // Inline drawings that aren't images — check for chart etc.
      const graphicEl = findChild(child, "a:graphic");
      const graphicDataEl = graphicEl ? findChild(graphicEl, "a:graphicData") : undefined;
      if (graphicDataEl) {
        const picEl = findChild(graphicDataEl, "pic:pic") ?? findChildNs(graphicDataEl, "pic");
        if (!picEl) {
          // Not an image — opaque inline drawing
          // Find the w:drawing parent
          const rids = new Set<string>();
          collectRIds(child, rids);
          // Serialize the wp:inline element wrapped in w:drawing
          const rawXml = `<w:drawing>${serializeElement(child)}</w:drawing>`;
          opaqueDrawings.push({
            type: "opaqueDrawing",
            rawXml,
            referencedRIds: [...rids]
          });
        }
      }
    } else {
      extractFloatingContent(child, images, shapes, opaqueDrawings);
    }
  }
}

/** Find the w:drawing ancestor element for serialization. */
function findDrawingParent(anchorEl: XmlElement): XmlElement | undefined {
  // We don't have parent refs, so we construct a synthetic w:drawing wrapper
  return {
    type: "element",
    name: "w:drawing",
    attributes: {},
    children: [anchorEl]
  } as XmlElement;
}

function parseDocumentXml(xmlStr: string): {
  body: BodyContent[];
  sectionProperties?: SectionProperties;
  background?: DocumentBackground;
} {
  const doc = parseXml(xmlStr);
  const root = doc.root;

  // Parse background
  let background: DocumentBackground | undefined;
  const bgEl = findChildNs(root, "background");
  if (bgEl) {
    const bg: any = {};
    const color = attrVal(bgEl, "color");
    if (color) {
      bg.color = color;
    }
    const themeColor = attrVal(bgEl, "themeColor");
    if (themeColor) {
      bg.themeColor = themeColor;
    }
    const themeShade = attrVal(bgEl, "themeShade");
    if (themeShade) {
      bg.themeShade = themeShade;
    }
    const themeTint = attrVal(bgEl, "themeTint");
    if (themeTint) {
      bg.themeTint = themeTint;
    }
    background = bg;
  }

  const bodyEl = findChildNs(root, "body") ?? findChild(root, "w:body");
  if (!bodyEl) {
    throw new DocxParseError("Missing w:body element in document.xml");
  }

  const body: BodyContent[] = [];
  let sectionProperties: SectionProperties | undefined;

  // Collect floating images, drawing shapes, and opaque drawings from the whole body
  const floatingImages: FloatingImage[] = [];
  const drawingShapes: DrawingShape[] = [];
  const opaqueDrawings: OpaqueDrawing[] = [];
  extractFloatingContent(bodyEl, floatingImages, drawingShapes, opaqueDrawings);

  for (const child of bodyEl.children) {
    if (child.type !== "element") {
      continue;
    }
    const name = child.name.replace(/^w:/, "");

    switch (name) {
      case "p":
        body.push(parseParagraph(child));
        break;
      case "tbl":
        body.push(parseTable(child));
        break;
      case "sectPr":
        // Final section properties at the body level
        sectionProperties = parseSectionProperties(child);
        break;
      case "sdt": {
        const sdtResult = parseSdt(child);
        if (sdtResult) {
          body.push(sdtResult as BodyContent);
        }
        break;
      }
      case "altChunk": {
        const rId = child.attributes["r:id"] ?? child.attributes["id"];
        if (rId) {
          body.push({ type: "altChunk", rId });
        }
        break;
      }
      default: {
        // Check for math namespace
        if (child.name === "m:oMathPara") {
          body.push(parseMathBlock(child));
        } else if (child.name === "m:oMath") {
          body.push({ type: "math", content: parseMathContent(child) });
        }
        // Check for VML pict (textbox)
        if (name === "pict" || child.name === "w:pict") {
          const tb = parseTextBox(child);
          if (tb) {
            body.push(tb);
          }
        }
        break;
      }
    }
  }

  // Append floating images as top-level body content
  for (const fi of floatingImages) {
    body.push(fi);
  }

  // Append drawing shapes as top-level body content
  for (const ds of drawingShapes) {
    body.push(ds);
  }

  // Append opaque drawings as top-level body content
  for (const od of opaqueDrawings) {
    body.push(od);
  }

  return { body, sectionProperties, background };
}

// =============================================================================
// Public API - Read DOCX
// =============================================================================

/**
 * Read a DOCX file from a Uint8Array buffer and parse it into a DocxDocument model.
 */
export async function readDocx(buffer: Uint8Array): Promise<DocxDocument> {
  try {
    return await _readDocxInner(buffer);
  } catch (e) {
    if (e instanceof DocxError) {
      throw e;
    }
    const msg = e instanceof Error ? e.message : String(e);
    throw new DocxParseError(`Failed to read DOCX: ${msg}`, { cause: e });
  }
}

async function _readDocxInner(buffer: Uint8Array): Promise<DocxDocument> {
  const reader = unzip(buffer);
  const entries = new Map<string, Uint8Array>();

  for await (const entry of reader.entries()) {
    const data = await entry.bytes();
    // Normalize path: remove leading slash, normalize separators
    const path = entry.path.replace(/^\//, "").replace(/\\/g, "/");
    entries.set(path, data);
  }

  const decoder = new TextDecoder("utf-8");
  const consumedPaths = new Set<string>(["[Content_Types].xml"]);
  const getText = (path: string): string | undefined => {
    const data = entries.get(path);
    if (data) {
      consumedPaths.add(path);
    }
    return data ? decoder.decode(data) : undefined;
  };

  // Parse document relationships (must be before parseDocumentXml for hyperlink resolution)
  const docRelsXml = getText("word/_rels/document.xml.rels");
  const docRels = docRelsXml ? parseRelationships(docRelsXml) : [];
  const _relMap = new Map(docRels.map(r => [r.id, r]));

  // Set module-level context for parseParagraph hyperlink resolution
  _parseRelMap = _relMap;

  // Parse document.xml (required)
  const documentXml = getText("word/document.xml");
  if (!documentXml) {
    throw new DocxMissingPartError("word/document.xml");
  }
  const { body, sectionProperties, background } = parseDocumentXml(documentXml);

  // Parse styles
  const stylesXml = getText("word/styles.xml");
  const stylesResult = stylesXml ? parseStyles(stylesXml) : undefined;

  // Parse numbering
  const numberingXml = getText("word/numbering.xml");
  const numberingResult = numberingXml ? parseNumberingXml(numberingXml) : undefined;

  // Parse footnotes/endnotes
  const footnotesXml = getText("word/footnotes.xml");
  const footnotes = footnotesXml ? parseNotesXml(footnotesXml, "footnote") : undefined;

  const endnotesXml = getText("word/endnotes.xml");
  const endnotes = endnotesXml ? parseNotesXml(endnotesXml, "endnote") : undefined;

  // Parse headers/footers + detect watermarks
  const headers = new Map<string, HeaderDef>();
  const footers = new Map<string, FooterDef>();
  let watermark: Watermark | undefined;

  for (const rel of docRels) {
    if (rel.type === RelType.Header) {
      const xml = getText(resolvePartPath("word/document.xml", rel.target));
      if (xml) {
        // Parse XML once, re-use for both header content and watermark detection
        const headerRoot = parseXml(xml).root;
        headers.set(rel.id, { content: parseHeaderFooterRoot(headerRoot), rId: rel.id });
        if (!watermark) {
          watermark = detectWatermarkFromRoot(headerRoot);
        }
      }
    } else if (rel.type === RelType.Footer) {
      const xml = getText(resolvePartPath("word/document.xml", rel.target));
      if (xml) {
        footers.set(rel.id, { content: parseHeaderFooterXml(xml), rId: rel.id });
      }
    }
  }

  // Parse settings
  const settingsXml = getText("word/settings.xml");
  const settings = settingsXml ? parseSettingsXml(settingsXml) : undefined;

  // Parse web settings
  const webSettingsXml = getText("word/webSettings.xml");
  const webSettings = webSettingsXml ? parseWebSettings(webSettingsXml) : undefined;

  // Parse people
  const peopleXml = getText("word/people.xml");
  const people = peopleXml ? parsePeople(peopleXml) : undefined;

  // Parse thumbnail (from package rels)
  let thumbnail: DocxDocument["thumbnail"];
  const packageRelsXml = getText("_rels/.rels");
  if (packageRelsXml) {
    const pkgRels = parseRelationships(packageRelsXml);
    for (const rel of pkgRels) {
      if (rel.type.endsWith("/thumbnail")) {
        // Target in package rels is relative to package root; may include or exclude leading slash
        let target = rel.target;
        if (target.startsWith("/")) {
          target = target.substring(1);
        }
        // If the target doesn't include docProps/ prefix, add it (some writers emit bare filenames)
        const normalized = target.includes("/") ? target : `docProps/${target}`;
        consumedPaths.add(normalized);
        const thumbData = entries.get(normalized);
        if (thumbData) {
          const ext = normalized.split(".").pop()?.toLowerCase();
          const ct =
            ext === "jpeg" || ext === "jpg"
              ? "image/jpeg"
              : ext === "png"
                ? "image/png"
                : "image/x-wmf";
          thumbnail = { contentType: ct as any, data: thumbData };
        }
        break;
      }
    }
  }

  // Parse font table
  const fontTableXml = getText("word/fontTable.xml");
  const fonts = fontTableXml ? parseFontTableXml(fontTableXml) : undefined;

  // Parse embedded fonts
  let embeddedFonts: EmbeddedFont[] | undefined;
  const fontTableRelsXml = getText("word/_rels/fontTable.xml.rels");
  if (fontTableRelsXml && fonts) {
    const fontRels = parseRelationships(fontTableRelsXml);
    const efs: EmbeddedFont[] = [];
    // Build rId → { key } map from font table
    const rIdToKey = new Map<string, string>();
    for (const f of fonts) {
      if (f.embedRegular && f.embedRegularKey) {
        rIdToKey.set(f.embedRegular, f.embedRegularKey);
      }
      if (f.embedBold && f.embedBoldKey) {
        rIdToKey.set(f.embedBold, f.embedBoldKey);
      }
      if (f.embedItalic && f.embedItalicKey) {
        rIdToKey.set(f.embedItalic, f.embedItalicKey);
      }
      if (f.embedBoldItalic && f.embedBoldItalicKey) {
        rIdToKey.set(f.embedBoldItalic, f.embedBoldItalicKey);
      }
    }
    for (const rel of fontRels) {
      if (rel.type === RelType.Font) {
        const fontPath = resolvePartPath("word/fontTable.xml", rel.target);
        consumedPaths.add(fontPath);
        const data = entries.get(fontPath);
        if (data) {
          const fileName = rel.target.split("/").pop() ?? "";
          const fontKey = rIdToKey.get(rel.id);
          const ef: any = {
            rId: rel.id,
            data,
            fileName
          };
          if (fontKey) {
            ef.fontKey = fontKey;
          }
          efs.push(ef);
        }
      }
    }
    if (efs.length > 0) {
      embeddedFonts = efs;
    }
  }

  // Parse Custom XML parts (for SDT data binding)
  const customXmlParts: CustomXmlPart[] = [];
  for (const rel of docRels) {
    if (rel.type === RelType.CustomXml) {
      const targetPath = resolvePartPath("word/document.xml", rel.target);
      consumedPaths.add(targetPath);
      const xmlContent = getText(targetPath);
      if (!xmlContent) {
        continue;
      }

      // Parse itemProps*.xml to get storeItemID
      const fileName = targetPath.split("/").pop() ?? "";
      // itemProps file is typically at the same directory
      const dir = targetPath.substring(0, targetPath.lastIndexOf("/"));
      // Extract item number from fileName (e.g. "item1.xml" → "1")
      const match = fileName.match(/item(\d+)\.xml$/);
      let itemId = "";
      let schemaReferences: string[] | undefined;
      if (match) {
        const num = match[1];
        const propsPath = `${dir}/itemProps${num}.xml`;
        consumedPaths.add(propsPath);
        const propsXml = getText(propsPath);
        if (propsXml) {
          const propsDoc = parseXml(propsXml);
          const dsItemEl = propsDoc.root;
          const id = dsItemEl.attributes["ds:itemID"];
          if (id) {
            itemId = id.replace(/[{}]/g, "");
          }
          // Schema references
          const refs: string[] = [];
          const schemaRefsEl =
            findChild(dsItemEl, "ds:schemaRefs") ?? findChild(dsItemEl, "schemaRefs");
          if (schemaRefsEl) {
            for (const srChild of schemaRefsEl.children) {
              if (srChild.type === "element") {
                const uri = srChild.attributes["ds:uri"] ?? srChild.attributes["uri"];
                if (uri) {
                  refs.push(uri);
                }
              }
            }
          }
          if (refs.length > 0) {
            schemaReferences = refs;
          }
        }
      }

      customXmlParts.push({
        itemId,
        xmlContent,
        fileName,
        schemaReferences
      });
    }
  }

  // Parse core properties
  const corePropsXml = getText("docProps/core.xml");
  const coreProperties = corePropsXml ? parseCoreProps(corePropsXml) : undefined;

  // Parse app properties
  const appPropsXml = getText("docProps/app.xml");
  const appProperties = appPropsXml ? parseAppProps(appPropsXml) : undefined;

  // Parse comments
  const commentsXml = getText("word/comments.xml");
  let comments = commentsXml ? parseCommentsXml(commentsXml) : undefined;

  // Merge in commentsExtended.xml data if present
  const commentsExtXml = getText("word/commentsExtended.xml");
  if (commentsExtXml && comments) {
    const extMap = parseCommentsExtendedXml(commentsExtXml);
    comments = comments.map(c => {
      const firstPara = c.content[0];
      if (!firstPara?.paraId) {
        return c;
      }
      const ext = extMap.get(firstPara.paraId);
      if (!ext) {
        return c;
      }
      return {
        ...c,
        ...(ext.done !== undefined ? { done: ext.done } : {}),
        ...(ext.parentId !== undefined ? { parentId: ext.parentId } : {})
      };
    });
  }

  // Parse custom properties
  const customPropsXml = getText("docProps/custom.xml");
  const customProperties = customPropsXml ? parseCustomPropsXml(customPropsXml) : undefined;

  // Parse theme
  const themeXml = getText("word/theme/theme1.xml");
  const theme = themeXml ? parseThemeXml(themeXml) : undefined;

  // Collect images
  const images: ImageDef[] = [];
  for (const rel of docRels) {
    if (rel.type === RelType.Image) {
      const imgPath = resolvePartPath("word/document.xml", rel.target);
      consumedPaths.add(imgPath);
      const data = entries.get(imgPath);
      if (data) {
        const fileName = rel.target.split("/").pop() ?? "";
        const ext = fileName.split(".").pop()?.toLowerCase() ?? "png";
        images.push({
          data,
          mediaType: ext as ImageMediaType,
          fileName,
          rId: rel.id
        });
      }
    }
  }

  // Collect opaque (unrecognized) parts for round-trip preservation
  const opaqueParts: OpaquePart[] = [];
  for (const [path, data] of entries) {
    // Skip consumed paths and all .rels files (structural)
    if (consumedPaths.has(path) || path.includes("_rels/")) {
      continue;
    }
    // Parse rels for this part if they exist
    const partRelsPath = getPartRelsPath(path);
    const partRelsData = entries.get(partRelsPath);
    let relationships: OpaqueRelationship[] | undefined;
    if (partRelsData) {
      const rels = parseRelationships(decoder.decode(partRelsData));
      relationships = rels.map(r => ({
        id: r.id,
        type: r.type,
        target: r.target,
        targetMode: r.targetMode === "External" ? "External" : undefined
      }));
    }
    opaqueParts.push({ path, data, relationships });
  }

  // Resolve altChunk data: body elements of type "altChunk" reference a rId.
  // The target file is stored in docRels + entries. Move target data from entries
  // (and remove from opaqueParts, since it's now part of the altChunk body element).
  for (const item of body) {
    if (item.type === "altChunk" && item.rId) {
      const rel = _relMap.get(item.rId);
      if (rel) {
        const target = resolvePartPath("word/document.xml", rel.target);
        const targetData = entries.get(target);
        if (targetData) {
          const fileName = target.split("/").pop();
          (item as any).data = targetData;
          (item as any).fileName = fileName;
          // Infer content type from extension
          const ext = fileName?.split(".").pop()?.toLowerCase();
          if (ext === "html" || ext === "htm") {
            (item as any).contentType = "text/html";
          } else if (ext === "rtf") {
            (item as any).contentType = "text/rtf";
          } else if (ext === "txt") {
            (item as any).contentType = "text/plain";
          }
        }
      }
    }
  }

  return {
    body,
    sectionProperties,
    styles: stylesResult?.styles,
    docDefaults: stylesResult?.docDefaults,
    abstractNumberings: numberingResult?.abstractNums,
    numberingInstances: numberingResult?.instances,
    numPicBullets:
      numberingResult?.numPicBullets && numberingResult.numPicBullets.length > 0
        ? numberingResult.numPicBullets
        : undefined,
    headers: headers.size > 0 ? headers : undefined,
    footers: footers.size > 0 ? footers : undefined,
    footnotes: footnotes && footnotes.length > 0 ? footnotes : undefined,
    endnotes: endnotes && endnotes.length > 0 ? endnotes : undefined,
    images: images.length > 0 ? images : undefined,
    fonts: fonts && fonts.length > 0 ? fonts : undefined,
    embeddedFonts: embeddedFonts && embeddedFonts.length > 0 ? embeddedFonts : undefined,
    customXmlParts: customXmlParts.length > 0 ? customXmlParts : undefined,
    webSettings,
    thumbnail,
    people: people && people.length > 0 ? people : undefined,
    settings,
    coreProperties,
    appProperties,
    comments: comments && comments.length > 0 ? comments : undefined,
    background,
    customProperties:
      customProperties && customProperties.length > 0 ? customProperties : undefined,
    theme,
    watermark,
    opaqueParts: opaqueParts.length > 0 ? opaqueParts : undefined
  };
}
