/**
 * DOCX Module - OpenDocument Text (ODT) Format Support
 *
 * Implements reading and writing of ODT (OpenDocument Text) files.
 * ODT files are ZIP archives containing XML content in ODF namespaces.
 *
 * Main archive structure:
 * - content.xml — document body and automatic styles
 * - styles.xml — named styles, page layout, master pages
 * - meta.xml — document metadata
 * - META-INF/manifest.xml — manifest of all archive entries
 * - Pictures/ — embedded images
 *
 * @stability experimental
 */

import { zip } from "@archive/create-archive";
import { unzip } from "@archive/read-archive";
import { parseXml, findChild, findChildren, textContent } from "@xml/dom";
import type { XmlElement } from "@xml/types";
import { XmlWriter } from "@xml/writer";

import { sanitizeUrl, utf8Decoder, utf8Encoder } from "../../core/internal-utils";
import { isRun } from "../../core/text-utils";
import { DocxParseError } from "../../errors";
import type {
  DocxDocument,
  AbstractNumbering,
  BodyContent,
  LevelSuffix,
  NumberingInstance,
  NumberFormat,
  NumberingLevel,
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
  TableCellProperties,
  StyleDef,
  StyleType,
  Alignment,
  LineSpacing,
  Indentation,
  SectionProperties,
  PageSize,
  PageMargins,
  CoreProperties,
  ImageDef,
  ImageMediaType,
  InlineImageContent
} from "../../types";
import { EMU_PER_CM, EMU_PER_INCH, EMU_PER_POINT, EMU_PER_PX } from "../../units";

// =============================================================================
// ODF Namespace Constants
// =============================================================================

/** ODF namespace URIs. */
const NS = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  style: "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  draw: "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
  fo: "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
  xlink: "http://www.w3.org/1999/xlink",
  dc: "http://purl.org/dc/elements/1.1/",
  meta: "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
  svg: "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
  manifest: "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
} as const;

// =============================================================================
// Internal Types
// =============================================================================

/** Parsed ODF style (intermediate representation). */
interface OdfStyle {
  readonly name: string;
  readonly family: string;
  readonly parentStyle?: string;
  readonly paragraphProperties?: OdfParagraphProps;
  readonly textProperties?: OdfTextProps;
  readonly tableProperties?: OdfTableProps;
  readonly tableColumnProperties?: OdfTableColumnProps;
  readonly tableCellProperties?: OdfTableCellProps;
  readonly pageLayoutProperties?: OdfPageLayoutProps;
}

interface OdfParagraphProps {
  readonly textAlign?: string;
  readonly marginTop?: string;
  readonly marginBottom?: string;
  readonly marginLeft?: string;
  readonly marginRight?: string;
  readonly textIndent?: string;
  readonly lineHeight?: string;
  readonly breakBefore?: string;
  readonly keepWithNext?: string;
}

interface OdfTextProps {
  readonly fontName?: string;
  readonly fontSize?: string;
  readonly fontWeight?: string;
  readonly fontStyle?: string;
  readonly textDecoration?: string;
  readonly color?: string;
  readonly backgroundColor?: string;
  readonly textPosition?: string;
  readonly fontVariant?: string;
  readonly letterSpacing?: string;
}

interface OdfTableProps {
  readonly width?: string;
  readonly align?: string;
}

interface OdfTableColumnProps {
  readonly columnWidth?: string;
}

interface OdfTableCellProps {
  readonly padding?: string;
  readonly borderTop?: string;
  readonly borderBottom?: string;
  readonly borderLeft?: string;
  readonly borderRight?: string;
  readonly backgroundColor?: string;
  readonly verticalAlign?: string;
}

interface OdfPageLayoutProps {
  readonly pageWidth?: string;
  readonly pageHeight?: string;
  readonly marginTop?: string;
  readonly marginBottom?: string;
  readonly marginLeft?: string;
  readonly marginRight?: string;
}

/** Context for tracking list state during parsing. */
interface ListParseContext {
  readonly listStyleName?: string;
  readonly level: number;
  /** numId assigned to this list's style (resolved via the registry). */
  readonly numId: number;
}

// =============================================================================
// Unit Conversion Helpers
// =============================================================================

/** Parse an ODF length value (e.g. "1.27cm", "0.5in", "12pt") to twips. */
function odfLengthToTwips(value: string | undefined): number | undefined {
  if (!value) {
    return undefined;
  }
  const match = value.match(/^(-?\d+(?:\.\d+)?)\s*(cm|mm|in|pt|pc|px|em)$/);
  if (!match) {
    return undefined;
  }
  const num = parseFloat(match[1]);
  const unit = match[2];
  switch (unit) {
    case "cm":
      return Math.round(num * 567);
    case "mm":
      return Math.round(num * 56.7);
    case "in":
      return Math.round(num * 1440);
    case "pt":
      return Math.round(num * 20);
    case "pc":
      return Math.round(num * 240);
    case "px":
      // Approximate: 1px ≈ 0.75pt at 96dpi
      return Math.round(num * 15);
    case "em":
      // Approximate: 1em ≈ 12pt
      return Math.round(num * 240);
    default:
      return undefined;
  }
}

/** Parse an ODF length value to EMU (English Metric Units). */
function odfLengthToEmu(value: string | undefined): number | undefined {
  if (!value) {
    return undefined;
  }
  const match = value.match(/^(-?\d+(?:\.\d+)?)\s*(cm|mm|in|pt|pc|px)$/);
  if (!match) {
    return undefined;
  }
  const num = parseFloat(match[1]);
  const unit = match[2];
  switch (unit) {
    case "cm":
      return Math.round(num * EMU_PER_CM);
    case "mm":
      return Math.round((num * EMU_PER_CM) / 10);
    case "in":
      return Math.round(num * EMU_PER_INCH);
    case "pt":
      return Math.round(num * EMU_PER_POINT);
    case "pc":
      return Math.round(num * EMU_PER_POINT * 12); // 1 pica = 12 points
    case "px":
      return Math.round(num * EMU_PER_PX);
    default:
      return undefined;
  }
}

/** Convert twips to ODF length string (cm). */
function twipsToCm(twips: number): string {
  return (twips / 567).toFixed(3) + "cm";
}

/** Convert EMU to ODF length string (cm). */
function emuToCm(emu: number): string {
  return (emu / EMU_PER_CM).toFixed(3) + "cm";
}

/** Convert half-points to pt string. */
function halfPointsToPt(hp: number): string {
  return (hp / 2).toString() + "pt";
}

/** Parse a font size string (e.g. "12pt") to half-points. */
function parseFontSizeToHalfPoints(value: string | undefined): number | undefined {
  if (!value) {
    return undefined;
  }
  const match = value.match(/^(\d+(?:\.\d+)?)\s*pt$/);
  if (!match) {
    return undefined;
  }
  return Math.round(parseFloat(match[1]) * 2);
}

/** Parse a 6-digit hex color from ODF format (#RRGGBB). */
function parseOdfColor(value: string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }
  const match = value.match(/^#([0-9a-fA-F]{6})$/);
  if (!match) {
    return undefined;
  }
  return match[1].toUpperCase();
}

/** Convert a hex color to ODF format (#RRGGBB). */
function colorToOdf(hex: string): string {
  return `#${hex}`;
}

// =============================================================================
// XML Element Query Helpers (Namespace-Aware)
// =============================================================================

/**
 * Find a child element by namespace-prefixed name.
 * Tries both "prefix:local" and just "local" for flexibility.
 */
function findNsChild(el: XmlElement, prefix: string, local: string): XmlElement | undefined {
  return findChild(el, `${prefix}:${local}`) ?? findChild(el, local);
}

/** Find all children by namespace-prefixed name. */
function findNsChildren(el: XmlElement, prefix: string, local: string): XmlElement[] {
  const result = findChildren(el, `${prefix}:${local}`);
  if (result.length > 0) {
    return result;
  }
  return findChildren(el, local);
}

/** Get an attribute value with a namespace prefix. */
function nsAttr(el: XmlElement, prefix: string, local: string): string | undefined {
  return el.attributes[`${prefix}:${local}`] ?? el.attributes[local];
}

// =============================================================================
// ODF Style Parsing
// =============================================================================

/** Parse style properties from a style:style element. */
function parseOdfStyle(el: XmlElement): OdfStyle {
  const name = nsAttr(el, "style", "name") ?? "";
  const family = nsAttr(el, "style", "family") ?? "";
  const parentStyle = nsAttr(el, "style", "parent-style-name");

  let paragraphProperties: OdfParagraphProps | undefined;
  let textProperties: OdfTextProps | undefined;
  let tableProperties: OdfTableProps | undefined;
  let tableColumnProperties: OdfTableColumnProps | undefined;
  let tableCellProperties: OdfTableCellProps | undefined;

  const pPropsEl = findNsChild(el, "style", "paragraph-properties");
  if (pPropsEl) {
    paragraphProperties = {
      textAlign: nsAttr(pPropsEl, "fo", "text-align"),
      marginTop: nsAttr(pPropsEl, "fo", "margin-top"),
      marginBottom: nsAttr(pPropsEl, "fo", "margin-bottom"),
      marginLeft: nsAttr(pPropsEl, "fo", "margin-left"),
      marginRight: nsAttr(pPropsEl, "fo", "margin-right"),
      textIndent: nsAttr(pPropsEl, "fo", "text-indent"),
      lineHeight: nsAttr(pPropsEl, "fo", "line-height"),
      breakBefore: nsAttr(pPropsEl, "fo", "break-before"),
      keepWithNext: nsAttr(pPropsEl, "fo", "keep-with-next")
    };
  }

  const tPropsEl = findNsChild(el, "style", "text-properties");
  if (tPropsEl) {
    textProperties = {
      fontName: nsAttr(tPropsEl, "style", "font-name") ?? nsAttr(tPropsEl, "fo", "font-family"),
      fontSize: nsAttr(tPropsEl, "fo", "font-size"),
      fontWeight: nsAttr(tPropsEl, "fo", "font-weight"),
      fontStyle: nsAttr(tPropsEl, "fo", "font-style"),
      textDecoration:
        nsAttr(tPropsEl, "style", "text-underline-style") ??
        nsAttr(tPropsEl, "style", "text-line-through-style"),
      color: nsAttr(tPropsEl, "fo", "color"),
      backgroundColor: nsAttr(tPropsEl, "fo", "background-color"),
      textPosition: nsAttr(tPropsEl, "style", "text-position"),
      fontVariant: nsAttr(tPropsEl, "fo", "font-variant"),
      letterSpacing: nsAttr(tPropsEl, "fo", "letter-spacing")
    };
  }

  const tablePropsEl = findNsChild(el, "style", "table-properties");
  if (tablePropsEl) {
    tableProperties = {
      width: nsAttr(tablePropsEl, "style", "width"),
      align: nsAttr(tablePropsEl, "table", "align") ?? nsAttr(tablePropsEl, "fo", "margin-left")
    };
  }

  const tableColPropsEl = findNsChild(el, "style", "table-column-properties");
  if (tableColPropsEl) {
    tableColumnProperties = {
      columnWidth: nsAttr(tableColPropsEl, "style", "column-width")
    };
  }

  const tableCellPropsEl = findNsChild(el, "style", "table-cell-properties");
  if (tableCellPropsEl) {
    tableCellProperties = {
      padding: nsAttr(tableCellPropsEl, "fo", "padding"),
      borderTop: nsAttr(tableCellPropsEl, "fo", "border-top"),
      borderBottom: nsAttr(tableCellPropsEl, "fo", "border-bottom"),
      borderLeft: nsAttr(tableCellPropsEl, "fo", "border-left"),
      borderRight: nsAttr(tableCellPropsEl, "fo", "border-right"),
      backgroundColor: nsAttr(tableCellPropsEl, "fo", "background-color"),
      verticalAlign: nsAttr(tableCellPropsEl, "style", "vertical-align")
    };
  }

  return {
    name,
    family,
    parentStyle,
    paragraphProperties,
    textProperties,
    tableProperties,
    tableColumnProperties,
    tableCellProperties
  };
}

/** Parse page layout properties from a style:page-layout element. */
function parsePageLayout(el: XmlElement): OdfPageLayoutProps | undefined {
  const propsEl = findNsChild(el, "style", "page-layout-properties");
  if (!propsEl) {
    return undefined;
  }
  return {
    pageWidth: nsAttr(propsEl, "fo", "page-width"),
    pageHeight: nsAttr(propsEl, "fo", "page-height"),
    marginTop: nsAttr(propsEl, "fo", "margin-top"),
    marginBottom: nsAttr(propsEl, "fo", "margin-bottom"),
    marginLeft: nsAttr(propsEl, "fo", "margin-left"),
    marginRight: nsAttr(propsEl, "fo", "margin-right")
  };
}

// =============================================================================
// ODF Content Parsing → DocxDocument Model
// =============================================================================

/** Convert ODF text alignment to OOXML alignment. */
function odfAlignToAlignment(align: string | undefined): Alignment | undefined {
  if (!align) {
    return undefined;
  }
  switch (align) {
    case "start":
    case "left":
      return "left";
    case "center":
      return "center";
    case "end":
    case "right":
      return "right";
    case "justify":
      return "both";
    default:
      return undefined;
  }
}

/** Convert ODF paragraph properties to OOXML ParagraphProperties. */
function odfParagraphPropsToDocx(
  props: OdfParagraphProps | undefined,
  styleName?: string
): ParagraphProperties | undefined {
  if (!props && !styleName) {
    return undefined;
  }

  const alignment = odfAlignToAlignment(props?.textAlign);

  let indent: Indentation | undefined;
  const leftTwips = odfLengthToTwips(props?.marginLeft);
  const rightTwips = odfLengthToTwips(props?.marginRight);
  const firstLineTwips = odfLengthToTwips(props?.textIndent);
  if (leftTwips !== undefined || rightTwips !== undefined || firstLineTwips !== undefined) {
    indent = {
      ...(leftTwips !== undefined ? { left: leftTwips } : {}),
      ...(rightTwips !== undefined ? { right: rightTwips } : {}),
      ...(firstLineTwips !== undefined
        ? firstLineTwips >= 0
          ? { firstLine: firstLineTwips }
          : { hanging: -firstLineTwips }
        : {})
    };
  }

  let spacing: LineSpacing | undefined;
  const beforeTwips = odfLengthToTwips(props?.marginTop);
  const afterTwips = odfLengthToTwips(props?.marginBottom);
  if (beforeTwips !== undefined || afterTwips !== undefined) {
    spacing = {
      ...(beforeTwips !== undefined ? { before: beforeTwips } : {}),
      ...(afterTwips !== undefined ? { after: afterTwips } : {})
    };
  }

  const keepNext =
    props?.keepWithNext === "always" || props?.keepWithNext === "true" ? true : undefined;
  const pageBreakBefore = props?.breakBefore === "page" ? true : undefined;

  const result: ParagraphProperties = {
    ...(styleName ? { style: styleName } : {}),
    ...(alignment ? { alignment } : {}),
    ...(indent ? { indent } : {}),
    ...(spacing ? { spacing } : {}),
    ...(keepNext !== undefined ? { keepNext } : {}),
    ...(pageBreakBefore !== undefined ? { pageBreakBefore } : {})
  };

  return Object.keys(result).length > 0 ? result : undefined;
}

/** Convert ODF text properties to OOXML RunProperties. */
function odfTextPropsToDocx(props: OdfTextProps | undefined): RunProperties | undefined {
  if (!props) {
    return undefined;
  }

  const font = props.fontName || undefined;
  const size = parseFontSizeToHalfPoints(props.fontSize);
  const bold = props.fontWeight === "bold" ? true : undefined;
  const italic = props.fontStyle === "italic" ? true : undefined;

  let underline: "single" | undefined;
  if (props.textDecoration && props.textDecoration !== "none") {
    underline = "single";
  }

  let strike: boolean | undefined;
  if (props.textDecoration === "line-through") {
    strike = true;
    underline = undefined;
  }

  const color = parseOdfColor(props.color);
  const smallCaps = props.fontVariant === "small-caps" ? true : undefined;
  const spacingTwips = odfLengthToTwips(props.letterSpacing);

  let vertAlign: "superscript" | "subscript" | undefined;
  if (props.textPosition) {
    if (props.textPosition.startsWith("super") || props.textPosition.startsWith("33%")) {
      vertAlign = "superscript";
    } else if (props.textPosition.startsWith("sub") || props.textPosition.startsWith("-33%")) {
      vertAlign = "subscript";
    }
  }

  const result: RunProperties = {
    ...(font ? { font } : {}),
    ...(size !== undefined ? { size } : {}),
    ...(bold !== undefined ? { bold } : {}),
    ...(italic !== undefined ? { italic } : {}),
    ...(underline !== undefined ? { underline } : {}),
    ...(strike !== undefined ? { strike } : {}),
    ...(color !== undefined ? { color } : {}),
    ...(smallCaps !== undefined ? { smallCaps } : {}),
    ...(spacingTwips !== undefined ? { spacing: spacingTwips } : {}),
    ...(vertAlign !== undefined ? { vertAlign } : {})
  };

  return Object.keys(result).length > 0 ? result : undefined;
}

/** Parse inline text spans (text:span) within a paragraph. */
function parseTextSpan(el: XmlElement, styles: Map<string, OdfStyle>): Run {
  const styleName = nsAttr(el, "text", "style-name");
  const style = styleName ? styles.get(styleName) : undefined;
  const runProps = odfTextPropsToDocx(style?.textProperties);
  const content: RunContent[] = [];

  for (const child of el.children) {
    if (child.type === "text") {
      if (child.value) {
        content.push({ type: "text", text: child.value });
      }
    } else if (child.type === "element") {
      const local = getLocalName(child.name);
      if (local === "s") {
        // text:s — multiple spaces
        const count = parseInt(nsAttr(child, "text", "c") ?? "1", 10);
        content.push({ type: "text", text: " ".repeat(count) });
      } else if (local === "tab") {
        content.push({ type: "tab" });
      } else if (local === "line-break") {
        content.push({ type: "break" });
      } else if (local === "span") {
        // Nested span — flatten into current run content
        const nestedRun = parseTextSpan(child, styles);
        content.push(...nestedRun.content);
      }
    }
  }

  return {
    ...(runProps ? { properties: runProps } : {}),
    content
  };
}

/** Parse a text:p or text:h element into a Paragraph. */
function parseParagraph(
  el: XmlElement,
  styles: Map<string, OdfStyle>,
  listContext?: ListParseContext
): Paragraph {
  const styleName = nsAttr(el, "text", "style-name");
  const style = styleName ? styles.get(styleName) : undefined;
  const isHeading = getLocalName(el.name) === "h";
  const outlineLevel = isHeading
    ? parseInt(nsAttr(el, "text", "outline-level") ?? "1", 10)
    : undefined;

  let paraProps = odfParagraphPropsToDocx(style?.paragraphProperties, styleName);

  // If this is a heading, add outline level
  if (outlineLevel !== undefined && outlineLevel >= 1 && outlineLevel <= 9) {
    paraProps = {
      ...paraProps,
      outlineLevel: outlineLevel - 1,
      style: styleName ?? `Heading${outlineLevel}`
    };
  }

  // If inside a list, add numbering reference
  if (listContext) {
    paraProps = {
      ...paraProps,
      numbering: { numId: listContext.numId, level: listContext.level }
    };
  }

  const children: ParagraphChild[] = [];

  for (const child of el.children) {
    if (child.type === "text") {
      if (child.value) {
        children.push({
          content: [{ type: "text", text: child.value }]
        });
      }
    } else if (child.type === "element") {
      const local = getLocalName(child.name);
      if (local === "span") {
        children.push(parseTextSpan(child, styles));
      } else if (local === "a") {
        // Hyperlink. Pass the href through sanitizeUrl so dangerous
        // schemes (javascript:, vbscript:, data: with executable payloads,
        // etc.) coming from an untrusted ODT do not survive into the DOCX
        // model. Unsafe links degrade to plain text — the caller still sees
        // the link's children but no clickable hyperlink is created.
        const rawHref = nsAttr(child, "xlink", "href");
        const safeHref = sanitizeUrl(rawHref);
        const runs: Run[] = [];
        for (const linkChild of child.children) {
          if (linkChild.type === "text") {
            if (linkChild.value) {
              runs.push({ content: [{ type: "text", text: linkChild.value }] });
            }
          } else if (linkChild.type === "element" && getLocalName(linkChild.name) === "span") {
            runs.push(parseTextSpan(linkChild, styles));
          }
        }
        if (safeHref) {
          children.push({
            type: "hyperlink",
            url: safeHref,
            children: runs
          });
        } else if (rawHref && rawHref.startsWith("#")) {
          // Internal anchor — sanitizeUrl rejects fragment-only URLs but
          // they're safe and meaningful. Preserve as a Hyperlink with
          // `anchor` set, mirroring how DOCX represents in-document
          // links.
          children.push({
            type: "hyperlink",
            anchor: rawHref.slice(1),
            children: runs
          });
        } else {
          // Drop the wrapper — preserve text content as plain runs so no
          // user-visible content disappears.
          for (const r of runs) {
            children.push(r);
          }
        }
      } else if (local === "s") {
        const count = parseInt(nsAttr(child, "text", "c") ?? "1", 10);
        children.push({ content: [{ type: "text", text: " ".repeat(count) }] });
      } else if (local === "tab") {
        children.push({ content: [{ type: "tab" }] });
      } else if (local === "line-break") {
        children.push({ content: [{ type: "break" }] });
      } else if (local === "frame") {
        // Inline image frame
        const imageRun = parseDrawFrame(child);
        if (imageRun) {
          children.push(imageRun);
        }
      } else if (local === "bookmark-start") {
        const bkName = nsAttr(child, "text", "name") ?? "";
        children.push({ type: "bookmarkStart", id: 0, name: bkName });
      } else if (local === "bookmark-end") {
        children.push({ type: "bookmarkEnd", id: 0 });
      } else if (local === "note") {
        // Footnote/endnote — extract text content
        const noteBodyEl = findNsChild(child, "text", "note-body");
        if (noteBodyEl) {
          const noteText = textContent(noteBodyEl);
          if (noteText) {
            children.push({ content: [{ type: "text", text: `[${noteText}]` }] });
          }
        }
      }
    }
  }

  return {
    type: "paragraph",
    ...(paraProps ? { properties: paraProps } : {}),
    children
  };
}

/** Parse a draw:frame element for inline images. */
function parseDrawFrame(el: XmlElement): Run | undefined {
  const width = odfLengthToEmu(nsAttr(el, "svg", "width"));
  const height = odfLengthToEmu(nsAttr(el, "svg", "height"));
  const name = nsAttr(el, "draw", "name");

  const imageEl =
    findNsChild(el, "draw", "image") ?? findChild(el, "draw:image") ?? findChild(el, "image");
  if (!imageEl) {
    return undefined;
  }

  const href = nsAttr(imageEl, "xlink", "href");
  if (!href) {
    return undefined;
  }

  const imageContent: InlineImageContent = {
    type: "image",
    rId: href, // Use the path as rId placeholder; resolved during image collection
    width: width ?? EMU_PER_INCH, // Default 1 inch
    height: height ?? EMU_PER_INCH,
    ...(name ? { name } : {})
  };

  return {
    content: [imageContent]
  };
}

/** Parse a text:list element. */
function parseList(
  el: XmlElement,
  styles: Map<string, OdfStyle>,
  parentLevel: number,
  registry: OdtNumberingRegistry,
  parentNumId?: number
): Paragraph[] {
  const listStyleName = nsAttr(el, "text", "style-name");
  // The outermost list determines the numId; nested lists inherit it so a
  // single multi-level list resolves to one numbering definition. Only the
  // top-level list (no parent numId) consults its own style name.
  const numId = parentNumId ?? registry.numIdFor(listStyleName);
  const paragraphs: Paragraph[] = [];
  const level = parentLevel;

  const items = findNsChildren(el, "text", "list-item");
  for (const item of items) {
    for (const child of item.children) {
      if (child.type !== "element") {
        continue;
      }
      const local = getLocalName(child.name);
      if (local === "p" || local === "h") {
        registry.noteLevel(numId, level);
        paragraphs.push(parseParagraph(child, styles, { listStyleName, level, numId }));
      } else if (local === "list") {
        // Nested list — inherit the enclosing list's numId.
        paragraphs.push(...parseList(child, styles, level + 1, registry, numId));
      }
    }
  }

  return paragraphs;
}

/** Parse a table:table element into a Table. */
function parseTable(
  el: XmlElement,
  styles: Map<string, OdfStyle>,
  registry: OdtNumberingRegistry
): Table {
  const tableStyleName = nsAttr(el, "table", "style-name");
  const tableStyle = tableStyleName ? styles.get(tableStyleName) : undefined;

  // Parse column definitions for widths
  const columnWidths: number[] = [];
  const colElements = findNsChildren(el, "table", "table-column");
  for (const colEl of colElements) {
    const colStyleName = nsAttr(colEl, "table", "style-name");
    const colStyle = colStyleName ? styles.get(colStyleName) : undefined;
    const width = odfLengthToTwips(colStyle?.tableColumnProperties?.columnWidth);
    const repeatCount = parseInt(nsAttr(colEl, "table", "number-columns-repeated") ?? "1", 10);
    for (let i = 0; i < repeatCount; i++) {
      columnWidths.push(width ?? 2000);
    }
  }

  // Parse rows
  const rows: TableRow[] = [];
  const rowElements = findNsChildren(el, "table", "table-row");
  for (const rowEl of rowElements) {
    const cells: TableCell[] = [];

    // Iterate in document order (preserve natural order)
    const orderedCells: XmlElement[] = [];
    for (const child of rowEl.children) {
      if (child.type === "element") {
        const childLocal = getLocalName(child.name);
        if (childLocal === "table-cell" || childLocal === "covered-table-cell") {
          orderedCells.push(child);
        }
      }
    }

    for (const cellEl of orderedCells) {
      const cellLocal = getLocalName(cellEl.name);
      const isCovered = cellLocal === "covered-table-cell";

      const gridSpan = parseInt(nsAttr(cellEl, "table", "number-columns-spanned") ?? "1", 10);
      const rowSpan = parseInt(nsAttr(cellEl, "table", "number-rows-spanned") ?? "1", 10);

      const cellProps: TableCellProperties = {
        ...(gridSpan > 1 ? { gridSpan } : {}),
        ...(rowSpan > 1 ? { rowSpan } : {}),
        ...(isCovered ? { verticalMerge: "continue" as const } : {})
      };

      // Parse cell content
      const cellContent: (Paragraph | Table)[] = [];
      for (const cellChild of cellEl.children) {
        if (cellChild.type !== "element") {
          continue;
        }
        const cellChildLocal = getLocalName(cellChild.name);
        if (cellChildLocal === "p" || cellChildLocal === "h") {
          cellContent.push(parseParagraph(cellChild, styles));
        } else if (cellChildLocal === "table") {
          cellContent.push(parseTable(cellChild, styles, registry));
        } else if (cellChildLocal === "list") {
          cellContent.push(...parseList(cellChild, styles, 0, registry));
        }
      }

      // Ensure at least one paragraph per cell
      if (cellContent.length === 0) {
        cellContent.push({ type: "paragraph", children: [] });
      }

      cells.push({
        ...(Object.keys(cellProps).length > 0 ? { properties: cellProps } : {}),
        content: cellContent
      });
    }

    rows.push({ cells });
  }

  // Build table properties
  let tableProps: TableProperties | undefined;
  const widthTwips = odfLengthToTwips(tableStyle?.tableProperties?.width);
  if (widthTwips !== undefined || tableStyleName) {
    tableProps = {
      ...(tableStyleName ? { style: tableStyleName } : {}),
      ...(widthTwips !== undefined ? { width: { value: widthTwips, type: "dxa" as const } } : {})
    };
  }

  return {
    type: "table",
    ...(tableProps ? { properties: tableProps } : {}),
    ...(columnWidths.length > 0 ? { columnWidths } : {}),
    rows
  };
}

/** Parse document body content from the office:body/office:text element. */
function parseDocumentBody(
  bodyEl: XmlElement,
  styles: Map<string, OdfStyle>,
  registry: OdtNumberingRegistry
): BodyContent[] {
  const content: BodyContent[] = [];

  for (const child of bodyEl.children) {
    if (child.type !== "element") {
      continue;
    }
    const local = getLocalName(child.name);

    if (local === "p" || local === "h") {
      content.push(parseParagraph(child, styles));
    } else if (local === "table") {
      content.push(parseTable(child, styles, registry));
    } else if (local === "list") {
      content.push(...parseList(child, styles, 0, registry));
    } else if (local === "section") {
      // Sections in ODF are logical containers; flatten their content
      content.push(...parseDocumentBody(child, styles, registry));
    }
  }

  return content;
}

/** Get the local name from a possibly prefixed XML name. */
function getLocalName(name: string): string {
  const colonIdx = name.indexOf(":");
  if (colonIdx >= 0) {
    return name.substring(colonIdx + 1);
  }
  return name;
}

// =============================================================================
// Meta Parsing
// =============================================================================

/** Parse meta.xml into CoreProperties. */
function parseMetaXml(xml: string): CoreProperties | undefined {
  let doc;
  try {
    doc = parseXml(xml);
  } catch {
    return undefined;
  }

  const root = doc.root;
  const metaEl = findNsChild(root, "office", "meta") ?? findChild(root, "office:meta") ?? root;

  const title = getMetaText(metaEl, "dc", "title");
  const subject = getMetaText(metaEl, "dc", "subject");
  const creator =
    getMetaText(metaEl, "dc", "creator") ?? getMetaText(metaEl, "meta", "initial-creator");
  const description = getMetaText(metaEl, "dc", "description");
  const keywords = getMetaText(metaEl, "meta", "keyword");
  const createdStr = getMetaText(metaEl, "meta", "creation-date");
  const modifiedStr = getMetaText(metaEl, "dc", "date");

  const created = createdStr ? new Date(createdStr) : undefined;
  const modified = modifiedStr ? new Date(modifiedStr) : undefined;

  const result: CoreProperties = {
    ...(title ? { title } : {}),
    ...(subject ? { subject } : {}),
    ...(creator ? { creator } : {}),
    ...(description ? { description } : {}),
    ...(keywords ? { keywords } : {}),
    ...(created && !isNaN(created.getTime()) ? { created } : {}),
    ...(modified && !isNaN(modified.getTime()) ? { modified } : {})
  };

  return Object.keys(result).length > 0 ? result : undefined;
}

/** Get text content of a metadata element. */
function getMetaText(parent: XmlElement, prefix: string, local: string): string | undefined {
  const el = findNsChild(parent, prefix, local);
  if (!el) {
    return undefined;
  }
  const text = textContent(el).trim();
  return text || undefined;
}

// =============================================================================
// Page Layout Parsing
// =============================================================================

/** Parse page layout from styles.xml into SectionProperties. */
function parsePageLayoutToSection(
  pageLayout: OdfPageLayoutProps | undefined
): SectionProperties | undefined {
  if (!pageLayout) {
    return undefined;
  }

  let pageSize: PageSize | undefined;
  const widthTwips = odfLengthToTwips(pageLayout.pageWidth);
  const heightTwips = odfLengthToTwips(pageLayout.pageHeight);
  if (widthTwips !== undefined && heightTwips !== undefined) {
    pageSize = {
      width: widthTwips,
      height: heightTwips,
      ...(widthTwips > heightTwips ? { orientation: "landscape" as const } : {})
    };
  }

  let margins: PageMargins | undefined;
  const marginTop = odfLengthToTwips(pageLayout.marginTop);
  const marginBottom = odfLengthToTwips(pageLayout.marginBottom);
  const marginLeft = odfLengthToTwips(pageLayout.marginLeft);
  const marginRight = odfLengthToTwips(pageLayout.marginRight);
  if (
    marginTop !== undefined ||
    marginBottom !== undefined ||
    marginLeft !== undefined ||
    marginRight !== undefined
  ) {
    margins = {
      top: marginTop ?? 1440,
      right: marginRight ?? 1440,
      bottom: marginBottom ?? 1440,
      left: marginLeft ?? 1440
    };
  }

  if (!pageSize && !margins) {
    return undefined;
  }

  return {
    ...(pageSize ? { pageSize } : {}),
    ...(margins ? { margins } : {})
  };
}

// =============================================================================
// Styles Conversion → StyleDef[]
// =============================================================================

/** Convert parsed ODF styles to DocxDocument style definitions. */
function convertStylesToStyleDefs(styles: Map<string, OdfStyle>): StyleDef[] {
  const defs: StyleDef[] = [];

  for (const [, style] of styles) {
    // Only convert named/non-automatic styles
    if (!style.name) {
      continue;
    }

    let type: StyleType;
    switch (style.family) {
      case "paragraph":
        type = "paragraph";
        break;
      case "text":
        type = "character";
        break;
      case "table":
        type = "table";
        break;
      default:
        continue; // Skip unsupported style families
    }

    const def: StyleDef = {
      type,
      styleId: style.name,
      name: style.name,
      ...(style.parentStyle ? { basedOn: style.parentStyle } : {}),
      ...(style.paragraphProperties
        ? { paragraphProperties: odfParagraphPropsToDocx(style.paragraphProperties) }
        : {}),
      ...(style.textProperties ? { runProperties: odfTextPropsToDocx(style.textProperties) } : {})
    };

    defs.push(def);
  }

  return defs;
}

// =============================================================================
// ODT list numbering definitions (bullet & numbered, multi-level)
// =============================================================================

/** Bullet glyphs per nesting level, cycling every three levels. */
const ODT_BULLET_CHARS = ["•", "◦", "▪", "•", "◦", "▪", "•", "◦", "▪"] as const;

/** Default ODF `style:num-format` value emitted for each docx NumberFormat. */
function numberFormatToOdf(format: NumberFormat): string {
  switch (format) {
    case "decimal":
    case "decimalZero":
      return "1";
    case "lowerLetter":
      return "a";
    case "upperLetter":
      return "A";
    case "lowerRoman":
      return "i";
    case "upperRoman":
      return "I";
    default:
      // ODF has no equivalent for many Word-specific formats; fall back to
      // decimal so the list still renders as an ordered list (its ordered
      // nature is preserved even when the exact glyph set is not).
      return "1";
  }
}

/** Map an ODF `style:num-format` value to the closest docx NumberFormat. */
function odfNumFormatToDocx(numFormat: string | undefined): NumberFormat {
  switch (numFormat) {
    case "1":
      return "decimal";
    case "a":
      return "lowerLetter";
    case "A":
      return "upperLetter";
    case "i":
      return "lowerRoman";
    case "I":
      return "upperRoman";
    default:
      return "decimal";
  }
}

/** Per-level numbering description parsed from an ODF `text:list-style`. */
interface OdfListLevel {
  readonly format: NumberFormat;
  /** Bullet glyph (bullet format) or empty for numbered levels. */
  readonly bulletChar?: string;
}

/**
 * Parse a single `text:list-style` element into its per-level formats.
 *
 * Each `text:list-level-style-bullet` becomes a bullet level; each
 * `text:list-level-style-number` becomes a numbered level whose docx format
 * is derived from `style:num-format`. Levels are indexed by `text:level`
 * (1-based in ODF) and stored 0-based.
 */
function parseOdfListStyle(el: XmlElement): Map<number, OdfListLevel> {
  const levels = new Map<number, OdfListLevel>();
  for (const child of el.children) {
    if (child.type !== "element") {
      continue;
    }
    const local = getLocalName(child.name);
    const levelAttr = parseInt(nsAttr(child, "text", "level") ?? "1", 10);
    const level = Math.max(0, levelAttr - 1);
    if (local === "list-level-style-bullet") {
      levels.set(level, {
        format: "bullet",
        bulletChar: nsAttr(child, "text", "bullet-char") ?? ODT_BULLET_CHARS[level] ?? "•"
      });
    } else if (local === "list-level-style-number") {
      levels.set(level, {
        format: odfNumFormatToDocx(nsAttr(child, "style", "num-format"))
      });
    }
  }
  return levels;
}

/**
 * Index every `text:list-style` found in the ODF automatic/named styles by
 * name. Used by the reader to recover the original bullet-vs-number format of
 * each list so the produced DOCX renders the correct markers.
 */
function collectOdfListStyles(
  roots: readonly XmlElement[]
): Map<string, Map<number, OdfListLevel>> {
  const out = new Map<string, Map<number, OdfListLevel>>();
  const visit = (el: XmlElement): void => {
    for (const child of el.children) {
      if (child.type !== "element") {
        continue;
      }
      if (getLocalName(child.name) === "list-style") {
        const name = nsAttr(child, "style", "name");
        if (name) {
          out.set(name, parseOdfListStyle(child));
        }
      } else {
        visit(child);
      }
    }
  };
  for (const root of roots) {
    visit(root);
  }
  return out;
}

/**
 * Registry that assigns a stable `numId` to each distinct ODF list-style name
 * encountered while parsing the body, and produces the matching docx
 * `abstractNumberings` / `numberingInstances` on demand.
 *
 * A registry instance is threaded through `parseList` → `parseParagraph` so
 * list paragraphs reference the same `numId` that ultimately resolves to the
 * synthesized numbering definition. Lists without an explicit style name (or
 * referencing an unknown style) fall back to a default bulleted definition.
 */
class OdtNumberingRegistry {
  private readonly listStyles: Map<string, Map<number, OdfListLevel>>;
  private readonly numIdByStyle = new Map<string, number>();
  // The format chosen at each level for a given numId (level 0..8).
  private readonly levelsByNumId = new Map<number, Map<number, OdfListLevel>>();
  private nextNumId = 1;

  constructor(listStyles: Map<string, Map<number, OdfListLevel>>) {
    this.listStyles = listStyles;
  }

  /** Resolve (and lazily register) the numId for a given list-style name. */
  numIdFor(styleName: string | undefined): number {
    const key = styleName ?? "";
    const existing = this.numIdByStyle.get(key);
    if (existing !== undefined) {
      return existing;
    }
    const numId = this.nextNumId++;
    this.numIdByStyle.set(key, numId);
    this.levelsByNumId.set(numId, this.listStyles.get(key) ?? new Map());
    return numId;
  }

  /**
   * Record the level a paragraph used for this numId so the synthesized
   * abstract numbering covers at least the deepest level actually referenced
   * even when the ODF list-style omitted some levels.
   */
  noteLevel(numId: number, level: number): void {
    let levels = this.levelsByNumId.get(numId);
    if (!levels) {
      levels = new Map();
      this.levelsByNumId.set(numId, levels);
    }
    if (!levels.has(level)) {
      // Unknown level: default to bullet so the marker is still visible.
      levels.set(level, { format: "bullet", bulletChar: ODT_BULLET_CHARS[level] ?? "•" });
    }
  }

  /** Whether any list paragraph was registered. */
  get isEmpty(): boolean {
    return this.numIdByStyle.size === 0;
  }

  /** Build the docx numbering definitions for every registered list. */
  build(): {
    abstractNumberings?: AbstractNumbering[];
    numberingInstances?: NumberingInstance[];
  } {
    if (this.isEmpty) {
      return {};
    }
    const abstractNumberings: AbstractNumbering[] = [];
    const numberingInstances: NumberingInstance[] = [];
    for (const numId of this.levelsByNumId.keys()) {
      const odfLevels = this.levelsByNumId.get(numId) ?? new Map<number, OdfListLevel>();
      const levels: NumberingLevel[] = [];
      for (let level = 0; level < 9; level++) {
        const info =
          odfLevels.get(level) ??
          ({ format: "bullet", bulletChar: ODT_BULLET_CHARS[level] ?? "•" } as OdfListLevel);
        levels.push(makeNumberingLevel(level, info));
      }
      // abstractNumId mirrors numId for a clean 1:1 mapping.
      abstractNumberings.push({ abstractNumId: numId, levels });
      numberingInstances.push({ numId, abstractNumId: numId });
    }
    return { abstractNumberings, numberingInstances };
  }
}

/** Build a single docx NumberingLevel from a parsed ODF level. */
function makeNumberingLevel(level: number, info: OdfListLevel): NumberingLevel {
  const tabSuffix: LevelSuffix = "tab";
  const indent = { left: 720 * (level + 1), hanging: 360 };
  if (info.format === "bullet") {
    return {
      level,
      format: "bullet",
      text: info.bulletChar ?? ODT_BULLET_CHARS[level] ?? "•",
      start: 1,
      paragraphProperties: { indent },
      suffix: tabSuffix
    };
  }
  return {
    level,
    format: info.format,
    // Word level text uses %N placeholders; "%<level+1>." renders "1." etc.
    text: `%${level + 1}.`,
    start: 1,
    paragraphProperties: { indent },
    suffix: tabSuffix
  };
}

// =============================================================================
// Writer-side list numbering helpers
// =============================================================================

/** Deterministic ODF `text:list-style` name for a given docx numId. */
function listStyleNameForNumId(numId: number): string {
  return `L${numId}`;
}

/** Collect every distinct numId referenced by list paragraphs in the body. */
function collectListNumIds(blocks: readonly BodyContent[], out: Set<number>): void {
  for (const block of blocks) {
    if (block.type === "paragraph") {
      const numId = block.properties?.numbering?.numId;
      if (typeof numId === "number") {
        out.add(numId);
      }
    } else if (block.type === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          collectListNumIds(cell.content as readonly BodyContent[], out);
        }
      }
    }
  }
}

/**
 * Resolve the per-level NumberFormat for a numId from the document's numbering
 * definitions. Follows numId → numberingInstance → abstractNumbering → levels.
 * Levels missing a definition default to bullet so a marker is still emitted.
 */
function resolveNumIdLevelFormats(doc: DocxDocument, numId: number): NumberFormat[] {
  const formats: NumberFormat[] = new Array(9).fill("bullet");
  const inst = doc.numberingInstances?.find(n => n.numId === numId);
  const abstractId = inst?.abstractNumId;
  const abstract =
    abstractId !== undefined
      ? doc.abstractNumberings?.find(a => a.abstractNumId === abstractId)
      : undefined;
  if (abstract) {
    for (const lvl of abstract.levels) {
      if (lvl.level >= 0 && lvl.level < 9) {
        formats[lvl.level] = lvl.format;
      }
    }
  }
  return formats;
}

// =============================================================================
// readOdt — Main Entry Point
// =============================================================================

/**
 * Read an ODT file and convert to DocxDocument model.
 *
 * Extracts the ZIP archive, parses content.xml, styles.xml, and meta.xml,
 * and produces a unified DocxDocument representation.
 *
 * @param buffer - The ODT file as a Uint8Array.
 * @returns A DocxDocument representing the ODT content.
 * @throws {DocxParseError} If the ODT file is malformed or missing required parts.
 *
 * @stability experimental
 */
export async function readOdt(buffer: Uint8Array): Promise<DocxDocument> {
  // Extract ZIP contents
  const reader = unzip(buffer);
  const entries = new Map<string, Uint8Array>();

  for await (const entry of reader.entries()) {
    const data = await entry.bytes();
    const path = entry.path.replace(/^\//, "").replace(/\\/g, "/");
    entries.set(path, data);
  }

  const decoder = utf8Decoder;

  // Parse content.xml (required)
  const contentData = entries.get("content.xml");
  if (!contentData) {
    throw new DocxParseError("Required ODT part not found: content.xml");
  }
  const contentXml = decoder.decode(contentData);
  const contentDoc = parseXml(contentXml);

  // Parse styles.xml (optional)
  const stylesData = entries.get("styles.xml");
  let stylesXml: string | undefined;
  if (stylesData) {
    stylesXml = decoder.decode(stylesData);
  }

  // Parse meta.xml (optional)
  const metaData = entries.get("meta.xml");
  let coreProperties: CoreProperties | undefined;
  if (metaData) {
    coreProperties = parseMetaXml(decoder.decode(metaData));
  }

  // Collect all styles (from both content.xml and styles.xml)
  const allStyles = new Map<string, OdfStyle>();
  let pageLayoutProps: OdfPageLayoutProps | undefined;
  let stylesRootEl: XmlElement | undefined;

  // Parse styles from styles.xml
  if (stylesXml) {
    const stylesDoc = parseXml(stylesXml);
    const stylesRoot = stylesDoc.root;
    stylesRootEl = stylesRoot;

    // Named styles (office:styles)
    const officeStylesEl = findNsChild(stylesRoot, "office", "styles");
    if (officeStylesEl) {
      for (const child of officeStylesEl.children) {
        if (child.type === "element" && getLocalName(child.name) === "style") {
          const s = parseOdfStyle(child);
          if (s.name) {
            allStyles.set(s.name, s);
          }
        }
      }
    }

    // Automatic styles in styles.xml (office:automatic-styles)
    const autoStylesEl = findNsChild(stylesRoot, "office", "automatic-styles");
    if (autoStylesEl) {
      for (const child of autoStylesEl.children) {
        if (child.type === "element") {
          const childLocal = getLocalName(child.name);
          if (childLocal === "style") {
            const s = parseOdfStyle(child);
            if (s.name) {
              allStyles.set(s.name, s);
            }
          } else if (childLocal === "page-layout") {
            pageLayoutProps = parsePageLayout(child);
          }
        }
      }
    }

    // Master page styles for page layout reference
    const masterStylesEl = findNsChild(stylesRoot, "office", "master-styles");
    if (masterStylesEl && !pageLayoutProps) {
      // Look for the page layout reference in master pages
      for (const child of masterStylesEl.children) {
        if (child.type === "element" && getLocalName(child.name) === "master-page") {
          const pageLayoutName = nsAttr(child, "style", "page-layout-name");
          if (pageLayoutName && autoStylesEl) {
            for (const autoChild of autoStylesEl.children) {
              if (
                autoChild.type === "element" &&
                getLocalName(autoChild.name) === "page-layout" &&
                nsAttr(autoChild, "style", "name") === pageLayoutName
              ) {
                pageLayoutProps = parsePageLayout(autoChild);
                break;
              }
            }
          }
          break;
        }
      }
    }
  }

  // Parse automatic styles from content.xml
  const contentAutoStylesEl = findNsChild(contentDoc.root, "office", "automatic-styles");
  if (contentAutoStylesEl) {
    for (const child of contentAutoStylesEl.children) {
      if (child.type === "element") {
        const childLocal = getLocalName(child.name);
        if (childLocal === "style") {
          const s = parseOdfStyle(child);
          if (s.name) {
            allStyles.set(s.name, s);
          }
        } else if (childLocal === "page-layout" && !pageLayoutProps) {
          pageLayoutProps = parsePageLayout(child);
        }
      }
    }
  }

  // Parse document body
  const bodyEl = findNsChild(contentDoc.root, "office", "body");
  if (!bodyEl) {
    throw new DocxParseError("Invalid ODT: missing office:body element");
  }
  const textEl = findNsChild(bodyEl, "office", "text");
  if (!textEl) {
    throw new DocxParseError("Invalid ODT: missing office:text element");
  }

  // Index every `text:list-style` (from styles.xml and content.xml) so the
  // body parser can recover each list's bullet-vs-number format. The registry
  // assigns a stable numId per distinct list style and produces the matching
  // numbering definitions below.
  const listStyleRoots: XmlElement[] = [contentDoc.root];
  if (stylesRootEl) {
    listStyleRoots.push(stylesRootEl);
  }
  const registry = new OdtNumberingRegistry(collectOdfListStyles(listStyleRoots));

  const body = parseDocumentBody(textEl, allStyles, registry);

  // Convert styles to StyleDef[]
  const styleDefs = convertStylesToStyleDefs(allStyles);

  // Parse section properties from page layout
  const sectionProperties = parsePageLayoutToSection(pageLayoutProps);

  // Collect images
  const images: ImageDef[] = [];
  for (const [path, data] of entries) {
    if (path.startsWith("Pictures/") && data.length > 0) {
      const ext = path.substring(path.lastIndexOf(".") + 1).toLowerCase();
      let mediaType: string;
      switch (ext) {
        case "png":
          mediaType = "image/png";
          break;
        case "jpg":
        case "jpeg":
          mediaType = "image/jpeg";
          break;
        case "gif":
          mediaType = "image/gif";
          break;
        case "svg":
          mediaType = "image/svg+xml";
          break;
        case "emf":
          mediaType = "image/x-emf";
          break;
        case "wmf":
          mediaType = "image/x-wmf";
          break;
        case "tiff":
        case "tif":
          mediaType = "image/tiff";
          break;
        case "bmp":
          mediaType = "image/bmp";
          break;
        default:
          mediaType = "application/octet-stream";
          break;
      }
      images.push({
        rId: path,
        data,
        mediaType: mediaType as ImageMediaType,
        fileName: path.substring(path.lastIndexOf("/") + 1)
      });
    }
  }

  const doc: DocxDocument = {
    body,
    ...(sectionProperties ? { sectionProperties } : {}),
    ...(styleDefs.length > 0 ? { styles: styleDefs } : {}),
    ...(coreProperties ? { coreProperties } : {}),
    ...(images.length > 0 ? { images } : {}),
    ...registry.build()
  };

  return doc;
}

/**
 * Build a deterministic, ZIP-safe `Pictures/` path for each image in the
 * document. The same rId → path map is consumed by both the content
 * writer (`xlink:href`) and the archive writer so that what we point at
 * in content.xml actually exists in the package.
 *
 * Image identifiers (`rId`, `fileName`) come from arbitrary upstream
 * input — including untrusted DOCX files round-tripped through readDocx.
 * Without sanitisation a hostile rId like `../../etc/passwd` would
 * produce both an out-of-tree ZIP entry name and an out-of-package
 * `xlink:href`, which is a real vector for confusing downstream ODT
 * readers and tooling that resolves relative paths.
 */
function buildOdtImagePathMap(images: readonly { rId?: string; fileName?: string }[] | undefined): {
  byRId: Map<string, string>;
  byFileName: Map<string, string>;
} {
  const byRId = new Map<string, string>();
  const byFileName = new Map<string, string>();
  if (!images) {
    return { byRId, byFileName };
  }
  const used = new Set<string>();
  // Allocate a unique `Pictures/<safe>` path for each image. We generate
  // a single counter-based fallback when the preferred name is empty or
  // collides with an entry that's already taken; the previous version
  // double-incremented `counter` (once inside the collision loop, once
  // after) producing sparse names like `image1.bin`, `image3.bin`, ….
  let counter = 1;
  const allocate = (preferred: string | undefined): string => {
    const safeBase = sanitizeOdtPictureName(preferred);
    let candidate = safeBase || `image${counter++}.bin`;
    while (used.has(candidate)) {
      // Suffix a fresh counter to disambiguate; preserve the extension
      // when we have one so downstream readers still classify the file
      // correctly. `_bin` was the previous fallback when no extension
      // was available — keep that behaviour for the no-base case.
      const c = counter++;
      if (safeBase) {
        const dot = safeBase.lastIndexOf(".");
        const stem = dot >= 0 ? safeBase.slice(0, dot) : safeBase;
        const ext = dot >= 0 ? safeBase.slice(dot) : "";
        candidate = `${stem}_${c}${ext}`;
      } else {
        candidate = `image${c}.bin`;
      }
    }
    used.add(candidate);
    return `Pictures/${candidate}`;
  };
  for (const image of images) {
    const path = allocate(image.fileName ?? image.rId);
    if (image.rId) {
      byRId.set(image.rId, path);
    }
    if (image.fileName) {
      // Two distinct images might share a fileName (the inputs are
      // attacker-controlled in the round-trip path). The first one to
      // claim a fileName wins the byFileName lookup; later collisions
      // are still reachable via their own rId in byRId. Without this
      // guard the second insertion would silently overwrite the first
      // and inline-image lookups via fileName would resolve to the
      // wrong binary.
      if (!byFileName.has(image.fileName)) {
        byFileName.set(image.fileName, path);
      }
    }
  }
  return { byRId, byFileName };
}

/**
 * Sanitise a single ODT picture file-name component. Strips path
 * separators, parent-directory traversal, leading dots, and anything
 * outside a conservative whitelist. Returns an empty string if no usable
 * characters remain (caller falls back to a generated name).
 */
function sanitizeOdtPictureName(raw: string | undefined): string {
  if (!raw) {
    return "";
  }
  // Drop directory components — we only ever want a leaf file name.
  const lastSep = Math.max(raw.lastIndexOf("/"), raw.lastIndexOf("\\"));
  let leaf = lastSep >= 0 ? raw.substring(lastSep + 1) : raw;
  // Strip leading dots so names like "..png" or ".htaccess" don't sneak
  // through with attribute meaning to filesystems.
  while (leaf.startsWith(".")) {
    leaf = leaf.substring(1);
  }
  // Whitelist alnum, dash, underscore, dot. Replace everything else with
  // underscore. Empty result triggers fallback in the caller.
  leaf = leaf.replace(/[^A-Za-z0-9._-]/g, "_");
  // Avoid pathological double-dot anywhere mid-name (e.g. "foo..bin").
  leaf = leaf.replace(/\.{2,}/g, ".");
  return leaf;
}

// =============================================================================
// writeOdt — Main Entry Point
// =============================================================================

/**
 * Convert a DocxDocument to ODT (OpenDocument Text) format.
 *
 * Generates the ZIP archive structure with content.xml, styles.xml,
 * meta.xml, and META-INF/manifest.xml.
 *
 * @param doc - The DocxDocument to convert.
 * @returns A Uint8Array containing the ODT ZIP archive.
 *
 * @stability experimental
 */
export async function writeOdt(doc: DocxDocument): Promise<Uint8Array> {
  const encoder = utf8Encoder;
  // `noSort: true` preserves insertion order. The ODF spec (OpenDocument
  // v1.2 part 3, §3.3) requires the `mimetype` entry to be the FIRST entry in
  // the package and STORED (uncompressed) so the file type can be detected by
  // magic bytes. The default ZipArchive behaviour sorts entries alphabetically,
  // which would push `mimetype` after `content.xml` and break ODF detection.
  const archive = zip({ noSort: true });

  // Mimetype MUST be the first entry in the ZIP, uncompressed (ODF spec requirement)
  archive.add("mimetype", encoder.encode("application/vnd.oasis.opendocument.text"), { level: 0 });

  // Compute a sanitised rId → Pictures/<safe>.<ext> map up front so the
  // content writer and archive writer agree on every image path.
  const imageMap = buildOdtImagePathMap(doc.images);

  // Collect image paths for manifest
  const imagePaths: string[] = [];

  // Generate content.xml — the writer reads from imageMap so href values
  // are guaranteed in-package and free of traversal sequences.
  const contentXml = generateContentXml(doc, imagePaths, imageMap);
  archive.add("content.xml", encoder.encode(contentXml));

  // Generate styles.xml
  const stylesXml = generateStylesXml(doc);
  archive.add("styles.xml", encoder.encode(stylesXml));

  // Generate meta.xml
  const metaXml = generateMetaXml(doc);
  archive.add("meta.xml", encoder.encode(metaXml));

  // Add images at the sanitised paths.
  if (doc.images) {
    for (const image of doc.images) {
      if (!image.data) {
        continue;
      }
      const imgPath =
        (image.rId ? imageMap.byRId.get(image.rId) : undefined) ??
        (image.fileName ? imageMap.byFileName.get(image.fileName) : undefined);
      if (!imgPath) {
        continue;
      }
      archive.add(imgPath, image.data);
      imagePaths.push(imgPath);
    }
  }

  // Generate META-INF/manifest.xml
  const manifestXml = generateManifestXml(imagePaths);
  archive.add("META-INF/manifest.xml", encoder.encode(manifestXml));

  return archive.bytes();
}

// =============================================================================
// Content XML Generation
// =============================================================================

/**
 * Context for collecting automatic text styles during body writing.
 * Maps a deterministic key derived from RunProperties to an auto style name.
 */
interface OdtWriteContext {
  /** Map from run properties key → auto style name */
  readonly runStyleMap: Map<string, string>;
  /** Map from auto style name → RunProperties (for writing styles later) */
  readonly runStyleProps: Map<string, RunProperties>;
  /** Counter for generating unique style names */
  nextRunStyleId: number;
  /**
   * Map from bookmark `id` → bookmark `name`. Populated as we encounter
   * `bookmarkStart` markers so the matching `bookmarkEnd` can reference the
   * same name (ODF requires `text:bookmark-end` to carry the same name as
   * its `text:bookmark-start`; emitting an empty name turns every bookmark
   * range into an isolated point in compliant readers).
   */
  readonly bookmarkNames: Map<number, string>;
  /**
   * Map from image rId → sanitised `Pictures/<safe>.<ext>` path. Built up
   * front by `writeOdt` and consumed by inline-image writers so the
   * `xlink:href` we emit always agrees with the entry actually added to
   * the ZIP, and so we never let a hostile rId escape the package root.
   */
  readonly imagePathByRId: Map<string, string>;
  /** Same as `imagePathByRId` but keyed by `image.fileName` for legacy lookups. */
  readonly imagePathByFileName: Map<string, string>;
}

/** Create a new write context. */
function createWriteContext(imageMap?: {
  byRId: Map<string, string>;
  byFileName: Map<string, string>;
}): OdtWriteContext {
  return {
    runStyleMap: new Map(),
    runStyleProps: new Map(),
    nextRunStyleId: 1,
    bookmarkNames: new Map(),
    imagePathByRId: imageMap?.byRId ?? new Map(),
    imagePathByFileName: imageMap?.byFileName ?? new Map()
  };
}

/**
 * Generate a deterministic cache key for RunProperties.
 * Only includes formatting-relevant properties.
 */
function runPropsKey(props: RunProperties): string {
  const parts: string[] = [];
  if (props.font) {
    const fontName = typeof props.font === "string" ? props.font : props.font.ascii;
    if (fontName) {
      parts.push(`f:${fontName}`);
    }
  }
  if (props.size !== undefined) {
    parts.push(`s:${props.size}`);
  }
  if (props.bold) {
    parts.push("b");
  }
  if (props.italic) {
    parts.push("i");
  }
  if (props.underline) {
    parts.push("u");
  }
  if (props.strike) {
    parts.push("st");
  }
  if (props.color) {
    const colorVal = typeof props.color === "string" ? props.color : props.color.val;
    if (colorVal) {
      parts.push(`c:${colorVal}`);
    }
  }
  if (props.smallCaps) {
    parts.push("sc");
  }
  if (props.vertAlign) {
    parts.push(`va:${props.vertAlign}`);
  }
  if (props.spacing !== undefined) {
    parts.push(`sp:${props.spacing}`);
  }
  return parts.join("|");
}

/** Get or create an automatic style name for the given RunProperties. */
function getRunAutoStyleName(ctx: OdtWriteContext, props: RunProperties): string {
  const key = runPropsKey(props);
  let name = ctx.runStyleMap.get(key);
  if (!name) {
    name = `T${ctx.nextRunStyleId++}`;
    ctx.runStyleMap.set(key, name);
    ctx.runStyleProps.set(name, props);
  }
  return name;
}

/** Generate the content.xml for the ODT package. */
function generateContentXml(
  doc: DocxDocument,
  imagePaths: string[],
  imageMap?: { byRId: Map<string, string>; byFileName: Map<string, string> }
): string {
  // First pass: write body to collect automatic run styles
  const ctx = createWriteContext(imageMap);
  const bodyWriter = new XmlWriter();
  bodyWriter.openNode("office:body");
  bodyWriter.openNode("office:text");
  writeBlocks(bodyWriter, doc.body, doc, imagePaths, ctx);
  bodyWriter.closeNode(); // office:text
  bodyWriter.closeNode(); // office:body
  const bodyXml = bodyWriter.xml;

  // Second pass: assemble the full content.xml with collected automatic styles
  const w = new XmlWriter();
  w.openXml();
  w.openNode("office:document-content", {
    "xmlns:office": NS.office,
    "xmlns:style": NS.style,
    "xmlns:text": NS.text,
    "xmlns:table": NS.table,
    "xmlns:draw": NS.draw,
    "xmlns:fo": NS.fo,
    "xmlns:xlink": NS.xlink,
    "xmlns:svg": NS.svg,
    "office:version": "1.3"
  });

  // Automatic styles section
  w.openNode("office:automatic-styles");
  writeAutoStyles(w, doc);
  writeCollectedRunStyles(w, ctx);
  w.closeNode();

  // Body (inject pre-built body XML)
  w.writeRaw(bodyXml);

  w.closeNode(); // office:document-content

  return w.xml;
}

/** Write collected automatic run styles into the automatic-styles section. */
function writeCollectedRunStyles(w: XmlWriter, ctx: OdtWriteContext): void {
  for (const [styleName, props] of ctx.runStyleProps) {
    w.openNode("style:style", {
      "style:name": styleName,
      "style:family": "text"
    });
    writeTextPropertiesOdf(w, props);
    w.closeNode();
  }
}

/** Write automatic styles based on document content. */
function writeAutoStyles(w: XmlWriter, doc: DocxDocument): void {
  // Output styles from the StyleDef array that exist in the document.
  if (doc.styles) {
    for (const styleDef of doc.styles) {
      writeStyleDef(w, styleDef);
    }
  }

  // Emit one `text:list-style` per distinct numId referenced by the body so
  // bullet vs numbered (and multi-level mixes) round-trip with the correct
  // markers. Numbers are sorted for deterministic output.
  const numIds = new Set<number>();
  collectListNumIds(doc.body, numIds);
  for (const numId of [...numIds].sort((a, b) => a - b)) {
    writeListStyleDef(w, numId, resolveNumIdLevelFormats(doc, numId));
  }
}

/**
 * Write the automatic `text:list-style` for a numId.
 *
 * Each of the nine ODF list levels is emitted as either a
 * `text:list-level-style-bullet` or `text:list-level-style-number` based on
 * the document's per-level NumberFormat, giving compliant readers (Word,
 * LibreOffice) the right markers and letting the reader recover the format on
 * the way back in.
 */
function writeListStyleDef(w: XmlWriter, numId: number, formats: readonly NumberFormat[]): void {
  w.openNode("text:list-style", { "style:name": listStyleNameForNumId(numId) });
  for (let idx = 0; idx < 9; idx++) {
    const level = idx + 1; // ODF list levels are 1-based.
    const format = formats[idx] ?? "bullet";
    if (format === "bullet") {
      w.openNode("text:list-level-style-bullet", {
        "text:level": String(level),
        "text:bullet-char": ODT_BULLET_CHARS[idx] ?? "•"
      });
    } else {
      w.openNode("text:list-level-style-number", {
        "text:level": String(level),
        "style:num-format": numberFormatToOdf(format),
        "style:num-suffix": "."
      });
    }
    w.openNode("style:list-level-properties", {
      "text:list-level-position-and-space-mode": "label-alignment"
    });
    w.leafNode("style:list-level-label-alignment", {
      "text:label-followed-by": "listtab",
      "fo:text-indent": "-0.25in",
      "fo:margin-left": `${0.25 * level}in`
    });
    w.closeNode(); // style:list-level-properties
    w.closeNode(); // text:list-level-style-{bullet,number}
  }
  w.closeNode(); // text:list-style
}

/** Write a single StyleDef as an ODF style. */
function writeStyleDef(w: XmlWriter, def: StyleDef): void {
  const family = styleTypeToOdfFamily(def.type);
  if (!family) {
    return;
  }

  w.openNode("style:style", {
    "style:name": def.styleId,
    "style:family": family,
    ...(def.basedOn ? { "style:parent-style-name": def.basedOn } : {}),
    ...(def.name !== def.styleId ? { "style:display-name": def.name } : {})
  });

  // Paragraph properties
  if (def.paragraphProperties) {
    writeParagraphPropertiesOdf(w, def.paragraphProperties);
  }

  // Run/text properties
  if (def.runProperties) {
    writeTextPropertiesOdf(w, def.runProperties);
  }

  w.closeNode();
}

/** Convert DocxDocument style type to ODF style family. */
function styleTypeToOdfFamily(type: StyleType): string | undefined {
  switch (type) {
    case "paragraph":
      return "paragraph";
    case "character":
      return "text";
    case "table":
      return "table";
    default:
      return undefined;
  }
}

/** Write paragraph properties as ODF elements. */
function writeParagraphPropertiesOdf(w: XmlWriter, props: ParagraphProperties): void {
  const attrs: Record<string, string> = {};

  if (props.alignment) {
    attrs["fo:text-align"] = alignmentToOdf(props.alignment);
  }
  if (props.indent) {
    if (props.indent.left !== undefined) {
      attrs["fo:margin-left"] = twipsToCm(props.indent.left);
    }
    if (props.indent.right !== undefined) {
      attrs["fo:margin-right"] = twipsToCm(props.indent.right);
    }
    if (props.indent.firstLine !== undefined) {
      attrs["fo:text-indent"] = twipsToCm(props.indent.firstLine);
    } else if (props.indent.hanging !== undefined) {
      attrs["fo:text-indent"] = twipsToCm(-props.indent.hanging);
    }
  }
  if (props.spacing) {
    if (props.spacing.before !== undefined) {
      attrs["fo:margin-top"] = twipsToCm(props.spacing.before);
    }
    if (props.spacing.after !== undefined) {
      attrs["fo:margin-bottom"] = twipsToCm(props.spacing.after);
    }
  }
  if (props.keepNext) {
    attrs["fo:keep-with-next"] = "always";
  }
  if (props.pageBreakBefore) {
    attrs["fo:break-before"] = "page";
  }

  if (Object.keys(attrs).length > 0) {
    w.leafNode("style:paragraph-properties", attrs);
  }
}

/** Write run/text properties as ODF elements. */
function writeTextPropertiesOdf(w: XmlWriter, props: RunProperties): void {
  const attrs: Record<string, string> = {};

  if (props.font) {
    const fontName = typeof props.font === "string" ? props.font : props.font.ascii;
    if (fontName) {
      attrs["style:font-name"] = fontName;
    }
  }
  if (props.size !== undefined) {
    attrs["fo:font-size"] = halfPointsToPt(props.size);
  }
  if (props.bold) {
    attrs["fo:font-weight"] = "bold";
  }
  if (props.italic) {
    attrs["fo:font-style"] = "italic";
  }
  if (props.underline) {
    attrs["style:text-underline-style"] = "solid";
    attrs["style:text-underline-width"] = "auto";
    attrs["style:text-underline-color"] = "font-color";
  }
  if (props.strike) {
    attrs["style:text-line-through-style"] = "solid";
  }
  if (props.color) {
    const colorVal = typeof props.color === "string" ? props.color : props.color.val;
    if (colorVal && colorVal !== "auto") {
      attrs["fo:color"] = colorToOdf(colorVal);
    }
  }
  if (props.smallCaps) {
    attrs["fo:font-variant"] = "small-caps";
  }
  if (props.vertAlign === "superscript") {
    attrs["style:text-position"] = "super 58%";
  } else if (props.vertAlign === "subscript") {
    attrs["style:text-position"] = "sub 58%";
  }
  if (props.spacing !== undefined) {
    attrs["fo:letter-spacing"] = twipsToCm(props.spacing);
  }

  if (Object.keys(attrs).length > 0) {
    w.leafNode("style:text-properties", attrs);
  }
}

/** Convert OOXML alignment to ODF text-align value. */
function alignmentToOdf(alignment: Alignment): string {
  switch (alignment) {
    case "left":
    case "start":
      return "start";
    case "center":
      return "center";
    case "right":
    case "end":
      return "end";
    case "both":
      return "justify";
    default:
      return "start";
  }
}

/**
 * Determine whether a body block is a list paragraph (carries numbering).
 *
 * The ODT reader maps list-item paragraphs back to `numbering.numId`, so the
 * writer must round-trip them as `text:list` structures rather than bare
 * `text:p` (which would silently drop the list semantics).
 */
function isListParagraph(block: BodyContent): block is Paragraph {
  return block.type === "paragraph" && block.properties?.numbering !== undefined;
}

/** numId of a list paragraph (falls back to 1 if somehow absent). */
function paragraphNumId(block: BodyContent): number {
  return block.type === "paragraph" ? (block.properties?.numbering?.numId ?? 1) : 1;
}

/**
 * Write a sequence of block-level content elements, grouping runs of
 * consecutive list paragraphs that share a numId into nested `text:list`
 * structures.
 *
 * The ODF list model nests one `text:list` per indentation level, so a
 * paragraph at `numbering.level === N` is wrapped in `N + 1` nested lists.
 * Tracking an explicit level stack lets a single run contain mixed levels
 * (e.g. a sub-bullet between two top-level bullets) and still produce valid,
 * round-trippable markup that matches what `parseList` expects. Breaking on
 * numId changes keeps each `text:list` pointed at a single list-style.
 */
function writeBlocks(
  w: XmlWriter,
  blocks: readonly BodyContent[],
  doc: DocxDocument,
  imagePaths: string[],
  ctx: OdtWriteContext
): void {
  let i = 0;
  while (i < blocks.length) {
    const block = blocks[i];
    if (isListParagraph(block)) {
      // Consume the maximal run of consecutive list paragraphs sharing numId.
      const numId = paragraphNumId(block);
      let j = i;
      while (
        j < blocks.length &&
        isListParagraph(blocks[j]) &&
        paragraphNumId(blocks[j]) === numId
      ) {
        j++;
      }
      writeListGroup(w, blocks.slice(i, j) as Paragraph[], numId, doc, imagePaths, ctx);
      i = j;
    } else {
      writeBodyContent(w, block, doc, imagePaths, ctx);
      i++;
    }
  }
}

/**
 * Emit a run of list paragraphs (all sharing `numId`) as nested `text:list`
 * elements.
 *
 * ODF nests one `text:list` per indentation level, and every nested list lives
 * inside a `text:list-item` of its parent. We therefore track, for each open
 * list level, whether a `text:list-item` is currently open at that level
 * (`itemOpen[d]`). Raising the level opens wrapper items plus nested lists;
 * lowering it closes them in the correct order so the markup stays balanced
 * and matches what `parseList` expects on the way back in.
 */
function writeListGroup(
  w: XmlWriter,
  paras: readonly Paragraph[],
  numId: number,
  doc: DocxDocument,
  imagePaths: string[],
  ctx: OdtWriteContext
): void {
  const styleName = listStyleNameForNumId(numId);
  // `itemOpen[d]` is true when a `text:list-item` is open inside the list at
  // depth `d` (0-based). Its length equals the number of open `text:list`s.
  const itemOpen: boolean[] = [];

  const openList = (): void => {
    w.openNode("text:list", { "text:style-name": styleName });
    itemOpen.push(false);
  };
  const closeList = (): void => {
    const depth = itemOpen.length - 1;
    if (itemOpen[depth]) {
      w.closeNode(); // text:list-item
    }
    itemOpen.pop();
    w.closeNode(); // text:list
  };

  for (const para of paras) {
    const level = Math.max(0, para.properties?.numbering?.level ?? 0);
    const wantLists = level + 1;

    // Close lists until we are at or below the desired depth.
    while (itemOpen.length > wantLists) {
      closeList();
    }
    // Open nested lists until we reach the desired depth. Each nested list
    // must sit inside an open `text:list-item` of its parent list.
    while (itemOpen.length < wantLists) {
      const depth = itemOpen.length - 1;
      if (depth >= 0 && !itemOpen[depth]) {
        w.openNode("text:list-item");
        itemOpen[depth] = true;
      }
      openList();
    }

    // At the target depth start a fresh list-item for this paragraph.
    const depth = itemOpen.length - 1;
    if (itemOpen[depth]) {
      w.closeNode(); // previous text:list-item at this depth
    }
    w.openNode("text:list-item");
    itemOpen[depth] = true;
    writeParagraph(w, para, doc, imagePaths, ctx);
  }

  // Unwind any lists still open.
  while (itemOpen.length > 0) {
    closeList();
  }
}

/** Write a block-level content element to the XML writer. */
function writeBodyContent(
  w: XmlWriter,
  block: BodyContent,
  doc: DocxDocument,
  imagePaths: string[],
  ctx: OdtWriteContext
): void {
  switch (block.type) {
    case "paragraph":
      writeParagraph(w, block, doc, imagePaths, ctx);
      break;
    case "table":
      writeTable(w, block, doc, imagePaths, ctx);
      break;
    case "tableOfContents":
      // Write TOC cached paragraphs if available
      if (block.cachedParagraphs) {
        for (const p of block.cachedParagraphs) {
          writeParagraph(w, p, doc, imagePaths, ctx);
        }
      }
      break;
    case "math":
      // Math blocks are rendered as paragraphs with math text
      w.openNode("text:p");
      for (const mc of block.content) {
        if (mc.type === "mathRun") {
          w.writeText(mc.text);
        }
      }
      w.closeNode();
      break;
    default:
      // Other block types (floatingImage, textBox, etc.) — emit as empty paragraph placeholder
      w.leafNode("text:p");
      break;
  }
}

/** Write a paragraph element. */
function writeParagraph(
  w: XmlWriter,
  para: Paragraph,
  doc: DocxDocument,
  imagePaths: string[],
  ctx: OdtWriteContext
): void {
  const isHeading =
    para.properties?.outlineLevel !== undefined && para.properties.outlineLevel >= 0;
  const styleName = para.properties?.style;

  if (isHeading) {
    const level = (para.properties!.outlineLevel ?? 0) + 1;
    w.openNode("text:h", {
      ...(styleName ? { "text:style-name": styleName } : {}),
      "text:outline-level": String(level)
    });
  } else {
    w.openNode("text:p", {
      ...(styleName ? { "text:style-name": styleName } : {})
    });
  }

  for (const child of para.children) {
    writeParagraphChild(w, child, doc, imagePaths, ctx);
  }

  w.closeNode();
}

/** Write a paragraph child (Run, Hyperlink, etc.). */
function writeParagraphChild(
  w: XmlWriter,
  child: ParagraphChild,
  doc: DocxDocument,
  imagePaths: string[],
  ctx: OdtWriteContext
): void {
  if (isRun(child)) {
    writeRun(w, child, doc, imagePaths, ctx);
    return;
  }

  switch (child.type) {
    case "hyperlink": {
      // Defense in depth: even though hyperlinks coming through readDocx /
      // htmlImport / odtRead are already sanitized, models can be built
      // by hand. Run the URL through sanitizeUrl one more time before
      // writing it into the ODT, so a malicious model can't smuggle
      // javascript:/vbscript: URLs through this writer.
      const rawHref = child.url ?? (child.anchor ? `#${child.anchor}` : "");
      const href = rawHref.startsWith("#") ? rawHref : (sanitizeUrl(rawHref) ?? "");
      if (!href) {
        // Drop the link wrapper — write children as plain runs.
        for (const run of child.children) {
          writeRun(w, run, doc, imagePaths, ctx);
        }
        break;
      }
      w.openNode("text:a", {
        "xlink:type": "simple",
        "xlink:href": href
      });
      for (const run of child.children) {
        writeRun(w, run, doc, imagePaths, ctx);
      }
      w.closeNode();
      break;
    }
    case "bookmarkStart":
      ctx.bookmarkNames.set(child.id, child.name);
      w.leafNode("text:bookmark-start", { "text:name": child.name });
      break;
    case "bookmarkEnd": {
      const name = ctx.bookmarkNames.get(child.id);
      if (name !== undefined) {
        // Emit the matching name so range-aware ODT readers can pair the
        // start/end markers correctly.
        w.leafNode("text:bookmark-end", { "text:name": name });
      } else {
        // Stray end without a known start (shouldn't happen for documents
        // produced by this library) — emit with an empty name to keep the
        // XML well-formed.
        w.leafNode("text:bookmark-end", { "text:name": "" });
      }
      break;
    }
    case "insertedRun":
      writeRun(w, child.run, doc, imagePaths, ctx);
      break;
    case "deletedRun":
      // Deleted runs are typically not shown in the output
      break;
    default:
      // Other paragraph children (comments, etc.) — skip
      break;
  }
}

/** Write a Run (text:span or direct text). */
function writeRun(
  w: XmlWriter,
  run: Run,
  doc: DocxDocument,
  imagePaths: string[],
  ctx: OdtWriteContext
): void {
  const hasProps = run.properties && Object.keys(run.properties).length > 0;

  if (hasProps) {
    const styleName = getRunAutoStyleName(ctx, run.properties!);
    w.openNode("text:span", { "text:style-name": styleName });
  }

  for (const content of run.content) {
    writeRunContent(w, content, doc, imagePaths, ctx);
  }

  if (hasProps) {
    w.closeNode();
  }
}

/** Write run content elements. */
function writeRunContent(
  w: XmlWriter,
  content: RunContent,
  doc: DocxDocument,
  imagePaths: string[],
  ctx: OdtWriteContext
): void {
  switch (content.type) {
    case "text":
      writeOdfText(w, content.text);
      break;
    case "break":
      if (content.breakType === "page") {
        // ODF 1.2 §5.1.4 — `text:soft-page-break` is a hint that a layout
        // page break occurs at this position. Strict semantic page breaks
        // require `fo:break-before="page"` on the surrounding paragraph
        // automatic style; we emit the soft hint here so the break is at
        // least preserved (rather than silently dropped, which corrupted
        // documents on round-trip).
        w.leafNode("text:soft-page-break");
      } else if (content.breakType === "column") {
        // Column breaks degrade to a soft page break in ODF — better than
        // dropping silently.
        w.leafNode("text:soft-page-break");
      } else {
        w.leafNode("text:line-break");
      }
      break;
    case "tab":
      w.leafNode("text:tab");
      break;
    case "image": {
      // Write inline image as draw:frame
      const widthCm = emuToCm(content.width);
      const heightCm = emuToCm(content.height);
      // Resolve the safe Pictures/ path that writeOdt registered when it
      // built the image map. This keeps content.xml's xlink:href in
      // lockstep with the actual ZIP entry name, and prevents a hostile
      // rId from emitting a traversing `xlink:href`.
      const mapped =
        ctx.imagePathByRId.get(content.rId) ??
        (content.name ? ctx.imagePathByFileName.get(content.name) : undefined);
      const imgPath = mapped ?? `Pictures/${sanitizeOdtPictureName(content.rId) || "image.bin"}`;

      w.openNode("draw:frame", {
        "draw:name": content.name ?? "",
        "svg:width": widthCm,
        "svg:height": heightCm,
        "text:anchor-type": "as-char"
      });
      w.leafNode("draw:image", {
        "xlink:href": imgPath,
        "xlink:type": "simple",
        "xlink:show": "embed",
        "xlink:actuate": "onLoad"
      });
      w.closeNode();
      break;
    }
    case "carriageReturn":
      w.leafNode("text:line-break");
      break;
    case "noBreakHyphen":
      w.writeText("\u2011");
      break;
    case "softHyphen":
      w.writeText("\u00AD");
      break;
    default:
      // Other content types (field, footnoteRef, etc.) — skip for now
      break;
  }
}

/** Write text content, handling multiple spaces via text:s. */
function writeOdfText(w: XmlWriter, text: string): void {
  // ODF collapses whitespace like HTML. Use text:s for multiple spaces.
  let i = 0;
  while (i < text.length) {
    const spaceStart = text.indexOf("  ", i);
    if (spaceStart < 0) {
      w.writeText(text.substring(i));
      break;
    }

    // Write text before the spaces
    if (spaceStart > i) {
      w.writeText(text.substring(i, spaceStart));
    }

    // Count consecutive spaces
    let spaceCount = 0;
    let j = spaceStart;
    while (j < text.length && text[j] === " ") {
      spaceCount++;
      j++;
    }

    // First space is implicit, rest use text:s
    w.writeText(" ");
    if (spaceCount > 1) {
      w.leafNode("text:s", { "text:c": String(spaceCount - 1) });
    }
    i = j;
  }
}

/** Write a table element. */
function writeTable(
  w: XmlWriter,
  table: Table,
  doc: DocxDocument,
  imagePaths: string[],
  ctx: OdtWriteContext
): void {
  w.openNode("table:table", {
    ...(table.properties?.style ? { "table:style-name": table.properties.style } : {})
  });

  // Write column definitions
  if (table.columnWidths && table.columnWidths.length > 0) {
    for (const _width of table.columnWidths) {
      w.leafNode("table:table-column");
    }
  }

  // Write rows
  for (const row of table.rows) {
    writeTableRow(w, row, doc, imagePaths, ctx);
  }

  w.closeNode();
}

/** Write a table row. */
function writeTableRow(
  w: XmlWriter,
  row: TableRow,
  doc: DocxDocument,
  imagePaths: string[],
  ctx: OdtWriteContext
): void {
  w.openNode("table:table-row");

  for (const cell of row.cells) {
    writeTableCell(w, cell, doc, imagePaths, ctx);
  }

  w.closeNode();
}

/** Write a table cell. */
function writeTableCell(
  w: XmlWriter,
  cell: TableCell,
  doc: DocxDocument,
  imagePaths: string[],
  ctx: OdtWriteContext
): void {
  const attrs: Record<string, string> = {};

  if (cell.properties?.gridSpan && cell.properties.gridSpan > 1) {
    attrs["table:number-columns-spanned"] = String(cell.properties.gridSpan);
  }
  if (cell.properties?.rowSpan && cell.properties.rowSpan > 1) {
    attrs["table:number-rows-spanned"] = String(cell.properties.rowSpan);
  }

  if (cell.properties?.verticalMerge === "continue") {
    w.leafNode("table:covered-table-cell");
    return;
  }

  w.openNode("table:table-cell", attrs);

  writeBlocks(w, cell.content as readonly BodyContent[], doc, imagePaths, ctx);

  w.closeNode();
}

// =============================================================================
// Styles XML Generation
// =============================================================================

/** Generate styles.xml for the ODT package. */
function generateStylesXml(doc: DocxDocument): string {
  const w = new XmlWriter();
  w.openXml();
  w.openNode("office:document-styles", {
    "xmlns:office": NS.office,
    "xmlns:style": NS.style,
    "xmlns:text": NS.text,
    "xmlns:table": NS.table,
    "xmlns:fo": NS.fo,
    "xmlns:draw": NS.draw,
    "xmlns:svg": NS.svg,
    "office:version": "1.3"
  });

  // Office styles (named styles)
  w.openNode("office:styles");

  // Default paragraph style
  w.openNode("style:default-style", { "style:family": "paragraph" });
  if (doc.docDefaults?.paragraphProperties) {
    writeParagraphPropertiesOdf(w, doc.docDefaults.paragraphProperties);
  }
  if (doc.docDefaults?.runProperties) {
    writeTextPropertiesOdf(w, doc.docDefaults.runProperties);
  } else {
    w.leafNode("style:text-properties", {
      "fo:font-size": "12pt",
      "style:font-name": "Times New Roman"
    });
  }
  w.closeNode();

  // Default table style
  w.openNode("style:default-style", { "style:family": "table" });
  w.closeNode();

  w.closeNode(); // office:styles

  // Automatic styles (page layout)
  w.openNode("office:automatic-styles");

  // Page layout
  w.openNode("style:page-layout", { "style:name": "pm1" });
  const pageProps: Record<string, string> = {};
  if (doc.sectionProperties?.pageSize) {
    pageProps["fo:page-width"] = twipsToCm(doc.sectionProperties.pageSize.width);
    pageProps["fo:page-height"] = twipsToCm(doc.sectionProperties.pageSize.height);
    if (doc.sectionProperties.pageSize.orientation === "landscape") {
      pageProps["style:print-orientation"] = "landscape";
    }
  } else {
    // Default A4
    pageProps["fo:page-width"] = "21.001cm";
    pageProps["fo:page-height"] = "29.700cm";
  }
  if (doc.sectionProperties?.margins) {
    const m = doc.sectionProperties.margins;
    pageProps["fo:margin-top"] = twipsToCm(m.top);
    pageProps["fo:margin-bottom"] = twipsToCm(m.bottom);
    pageProps["fo:margin-left"] = twipsToCm(m.left);
    pageProps["fo:margin-right"] = twipsToCm(m.right);
  } else {
    // Default 1 inch margins
    pageProps["fo:margin-top"] = "2.540cm";
    pageProps["fo:margin-bottom"] = "2.540cm";
    pageProps["fo:margin-left"] = "2.540cm";
    pageProps["fo:margin-right"] = "2.540cm";
  }
  w.leafNode("style:page-layout-properties", pageProps);
  w.closeNode(); // style:page-layout

  w.closeNode(); // office:automatic-styles

  // Master styles
  w.openNode("office:master-styles");
  w.leafNode("style:master-page", {
    "style:name": "Default",
    "style:page-layout-name": "pm1"
  });
  w.closeNode();

  w.closeNode(); // office:document-styles

  return w.xml;
}

// =============================================================================
// Meta XML Generation
// =============================================================================

/** Generate meta.xml for the ODT package. */
function generateMetaXml(doc: DocxDocument): string {
  const w = new XmlWriter();
  w.openXml();
  w.openNode("office:document-meta", {
    "xmlns:office": NS.office,
    "xmlns:meta": NS.meta,
    "xmlns:dc": NS.dc,
    "office:version": "1.3"
  });

  w.openNode("office:meta");

  if (doc.coreProperties) {
    const cp = doc.coreProperties;
    if (cp.title) {
      w.leafNode("dc:title", undefined, cp.title);
    }
    if (cp.subject) {
      w.leafNode("dc:subject", undefined, cp.subject);
    }
    if (cp.creator) {
      w.leafNode("meta:initial-creator", undefined, cp.creator);
      w.leafNode("dc:creator", undefined, cp.creator);
    }
    if (cp.description) {
      w.leafNode("dc:description", undefined, cp.description);
    }
    if (cp.keywords) {
      w.leafNode("meta:keyword", undefined, cp.keywords);
    }
    if (cp.created) {
      w.leafNode("meta:creation-date", undefined, cp.created.toISOString());
    }
    if (cp.modified) {
      w.leafNode("dc:date", undefined, cp.modified.toISOString());
    }
  }

  // Generator
  w.leafNode("meta:generator", undefined, "excelts/odt");

  w.closeNode(); // office:meta
  w.closeNode(); // office:document-meta

  return w.xml;
}

// =============================================================================
// Manifest XML Generation
// =============================================================================

/** Generate META-INF/manifest.xml for the ODT package. */
function generateManifestXml(imagePaths: string[]): string {
  const w = new XmlWriter();
  w.openXml();
  w.openNode("manifest:manifest", {
    "xmlns:manifest": NS.manifest,
    "manifest:version": "1.3"
  });

  // Root entry
  w.leafNode("manifest:file-entry", {
    "manifest:full-path": "/",
    "manifest:version": "1.3",
    "manifest:media-type": "application/vnd.oasis.opendocument.text"
  });

  // Content parts
  w.leafNode("manifest:file-entry", {
    "manifest:full-path": "content.xml",
    "manifest:media-type": "text/xml"
  });
  w.leafNode("manifest:file-entry", {
    "manifest:full-path": "styles.xml",
    "manifest:media-type": "text/xml"
  });
  w.leafNode("manifest:file-entry", {
    "manifest:full-path": "meta.xml",
    "manifest:media-type": "text/xml"
  });

  // Images
  for (const imgPath of imagePaths) {
    const ext = imgPath.substring(imgPath.lastIndexOf(".") + 1).toLowerCase();
    let mediaType: string;
    switch (ext) {
      case "png":
        mediaType = "image/png";
        break;
      case "jpg":
      case "jpeg":
        mediaType = "image/jpeg";
        break;
      case "gif":
        mediaType = "image/gif";
        break;
      case "svg":
        mediaType = "image/svg+xml";
        break;
      default:
        mediaType = "application/octet-stream";
        break;
    }
    w.leafNode("manifest:file-entry", {
      "manifest:full-path": imgPath,
      "manifest:media-type": mediaType
    });
  }

  w.closeNode(); // manifest:manifest

  return w.xml;
}
