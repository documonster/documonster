/**
 * DOCX Module - Reader / Parser
 *
 * Reads a DOCX ZIP file and parses it into a DocxDocument model.
 * Uses the archive module for ZIP reading and XML module for parsing.
 */

import { unzipZip } from "@archive/read-archive";
import { appendAll } from "@utils/array";
import type { SchemaMatches } from "@utils/schema";
import { runSteps } from "@utils/yield";
import type { Steps } from "@utils/yield";
import { RelType, ContentType } from "@word/constants";
import { isFlatOpc } from "@word/convert/flat-opc";
import { WRITTEN, checkBytes, invalidArgument, readToken } from "@word/core/argument-checks";
import { OOXML_TO_GALLERY } from "@word/core/glossary-galleries";
import type { Mutable } from "@word/core/internal-utils";
import { base64ToBytes, utf8Decoder, utf8Encoder } from "@word/core/internal-utils";
import { checkSchema, w } from "@word/core/option-schema";
import {
  arrangeParagraphChildren,
  isRidingMarker,
  foldRangeMarkers,
  removeBlock,
  withRangeMarkersBefore
} from "@word/core/range-markers";
import { isRangeMarker, isRun } from "@word/core/text-utils";
import { findTocFieldPieces, parseTocInstruction } from "@word/core/toc-instruction";
import { walkBlocks } from "@word/core/walker";
import { headerSlots } from "@word/core/watermark-sections";
import type { HeaderSlot } from "@word/core/watermark-sections";
import {
  DocxError,
  DocxParseError,
  DocxMissingPartError,
  DocxLimitExceededError
} from "@word/errors";
import {
  replaceOpaqueCharts,
  replaceOpaqueChartExDrawings,
  parseChartXml,
  parseChartExXml
} from "@word/reader/chart-parser";
import {
  parseCommentsXml as parseCommentsXmlExternal,
  parseCommentsExtendedXml
} from "@word/reader/comments-parser";
import {
  parseCoreProps,
  parseAppProps,
  parseCustomPropsXml,
  parseFontTableXml
} from "@word/reader/doc-props-parsers";
import {
  frameFieldsOf,
  readDrawingFrame,
  readingPreservedLinks,
  resolvePreservedLinks
} from "@word/reader/drawing-frame";
import {
  expandBlockSimpleFields,
  fieldCodeTextOf,
  fieldPieceOf,
  isResultText,
  splitFieldPieces,
  splitFieldPiecesSteps
} from "@word/reader/field-pieces";
import type { FieldPiece } from "@word/reader/field-pieces";
import { parseFfData } from "@word/reader/form-field-parser";
import { parseDrawingContent, parseFloatingImage } from "@word/reader/image-parsers";
import { createImageTable } from "@word/reader/image-table";
import { parseMathContent, parseMathBlock } from "@word/reader/math-parser";
import {
  parseThemeXml,
  parseWebSettings,
  parsePeople,
  parseSettingsXml
} from "@word/reader/metadata-parsers";
import { parseNumberingXml } from "@word/reader/numbering-parser";
import {
  checkPackageSize,
  checkReadOptions,
  createPartBudget,
  readOptionsShape,
  flatOpcEntries
} from "@word/reader/package-limits";
import {
  parseParagraphProperties,
  parseSectionProperties
} from "@word/reader/paragraph-section-parsers";
import {
  attrVal,
  attrInt,
  findChildNs,
  isWordDocumentRoot,
  findChildrenNs,
  boolToggle,
  serializeElement,
  serializeElementSteps,
  collectRIds,
  getPartRelsPath,
  getFileName,
  getFileExt,
  resolvePartPath,
  resolveContentChild,
  resolveRelTarget,
  stripWPrefix
} from "@word/reader/parse-utils";
import { parsePartXml } from "@word/reader/part-xml";
import { parseBoundContentPart } from "@word/reader/preserved-xml";
import {
  parseRowHeight,
  parseRunProperties,
  parseShading,
  parseTableWidth,
  parseRevisionInfo
} from "@word/reader/properties-parsers";
import { parseRangeMarker, takeRangeMarker } from "@word/reader/range-markers";
import type {
  ParsedRelationship,
  PendingFieldMarker,
  ReaderContext
} from "@word/reader/reader-context";
import {
  createFieldState,
  createReaderContext,
  parseRelationships
} from "@word/reader/reader-context";
import { parseCheckBox } from "@word/reader/sdt-helpers";
import { preserveShapeMarkup } from "@word/reader/shape-markup";
import { parseStyles } from "@word/reader/styles-parser";
import {
  parseTableBorders,
  parseTableCellMargins,
  parseTableProperties
} from "@word/reader/table-properties-parsers";
import { findWatermark, withoutWatermark } from "@word/reader/watermark-parser";
import type { FoundWatermark } from "@word/reader/watermark-parser";
import { decryptIfEncrypted } from "@word/security/package-encryption";
import type { DocxSecurityPolicy } from "@word/security/policy";
import type {
  DocxDocument,
  BodyContent,
  NoteType,
  OpaqueParagraphChild,
  Paragraph,
  PreservedLink,
  RangeMarker,
  ParagraphChild,
  RevisionInfo,
  Run,
  RunContent,
  Table,
  TableRow,
  TableCell,
  TableRowProperties,
  TableCellProperties,
  SectionProperties,
  HeaderDef,
  FooterDef,
  HeaderFooterContent,
  EmbeddedFont,
  CustomXmlPart,
  ImageDef,
  CommentDef,
  AltChunk,
  InsertedRun,
  DeletedRun,
  FloatingImage,
  RubyContent,
  RubyProperties,
  PositionalTabContent,
  FootnoteRefContent,
  EndnoteRefContent,
  CommentReference,
  Hyperlink,
  HyperlinkNestedChild,
  SmartTagAttribute,
  SmartTagStart,
  TextBox,
  StructuredDocumentTag,
  SdtListItem,
  SdtDateProperties,
  BlockContentControl,
  SdtDocPart,
  SdtProperties,
  TableOfContents,
  CheckBox,
  DocumentBackground,
  FieldContent,
  Watermark,
  DrawingShape,
  OpaquePart,
  OpaqueRelationship,
  OpaqueDrawing,
  OleObjectPart,
  GlossaryDocument,
  BuildingBlock,
  Chart,
  ChartExContent,
  DocxDocumentType
} from "@word/types";
import { findChild, textContent } from "@xml/dom";
import type { XmlElement, XmlNode } from "@xml/types";

// =============================================================================
// Run Content Parser
// =============================================================================

function parseRunContent(
  el: XmlElement,
  content: RunContent[] = [],
  from = 0,
  to = el.children.length
): RunContent[] {
  const end = Math.min(to, el.children.length);
  for (let index = from; index < end; index++) {
    const child = el.children[index]!;
    if (child.type !== "element") {
      continue;
    }
    const name = stripWPrefix(child.name);
    switch (name) {
      case "t":
        content.push({ type: "text", text: textContent(child) });
        break;
      case "br": {
        const brType = attrVal(child, "type");
        content.push({
          type: "break",
          breakType: brType as "page" | "column" | "textWrapping" | undefined
        });
        break;
      }
      case "tab":
        content.push({ type: "tab" });
        break;
      case "ptab": {
        const alignment = attrVal(child, "alignment") ?? "left";
        const relativeTo = attrVal(child, "relativeTo") ?? "margin";
        const leader = attrVal(child, "leader");
        const ptab: Mutable<PositionalTabContent> = {
          type: "ptab",
          alignment: alignment as PositionalTabContent["alignment"],
          relativeTo: relativeTo as PositionalTabContent["relativeTo"]
        };
        if (leader) {
          ptab.leader = leader as PositionalTabContent["leader"];
        }
        content.push(ptab);
        break;
      }
      case "ruby":
        content.push(parseRuby(child));
        break;
      case "sym":
        content.push({
          type: "symbol",
          font: attrVal(child, "font") ?? "",
          char: attrVal(child, "char") ?? ""
        });
        break;
      case "footnoteReference": {
        const fr: Mutable<FootnoteRefContent> = {
          type: "footnoteRef",
          id: attrInt(child, "id") ?? 0
        };
        const cmf = attrVal(child, "customMarkFollows");
        if (cmf === "1" || cmf === "true") {
          fr.customMarkFollows = true;
        }
        content.push(fr);
        break;
      }
      case "endnoteReference": {
        const er: Mutable<EndnoteRefContent> = {
          type: "endnoteRef",
          id: attrInt(child, "id") ?? 0
        };
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
        content.push({
          type: "annotationReference",
          id: attrInt(child, "id") ?? 0
        });
        break;
      // `w:commentReference` is not run content in the model: a run in a
      // paragraph (or a tracked change) is split around it by
      // parseRunParts. Where a run cannot be split — a hyperlink's runs, a
      // run directly in an SDT, ruby text — it falls to the default below
      // and is kept verbatim, id and all.
      case "rPr":
      case "fldChar":
      case "instrText":
      case "delText":
        // Known structural elements handled elsewhere — skip silently
        break;
      default:
        // Unknown run child: preserve as opaque for round-trip fidelity
        content.push({
          type: "opaqueRun",
          rawXml: serializeElement(child)
        });
        break;
    }
  }
  return content;
}

// =============================================================================
// DrawingML Shape Parser
// =============================================================================

/**
 * A shape whose text box paragraphs are still to be read, and their elements: what {@link parseDrawingShape} leaves
 * for a read in steps, the text of one shape being a block of its own however many paragraphs it holds.
 */
interface PendingShapeText {
  /** The shape's `textContent`, to be filled in place: what the shape records of itself (`_preserved`) holds it too. */
  readonly textContent: Paragraph[];
  readonly paragraphs: readonly XmlElement[];
}

function parseDrawingShape(
  anchorEl: XmlElement,
  wspEl: XmlElement,
  ctx: ReaderContext,
  pendingText?: PendingShapeText[]
): DrawingShape | undefined {
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

  const shape: Mutable<DrawingShape> = {
    type: "drawingShape",
    shapeType: shapeType as DrawingShape["shapeType"],
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

    // Parse transform (rotation / flip) from a:xfrm
    const xfrmEl = findChild(spPrEl, "a:xfrm") ?? findChildNs(spPrEl, "xfrm");
    if (xfrmEl) {
      const rot = xfrmEl.attributes["rot"];
      if (rot) {
        const n = parseInt(rot, 10);
        if (!Number.isNaN(n)) {
          shape.rotation = n;
        }
      }
      if (xfrmEl.attributes["flipH"] === "1") {
        shape.flipHorizontal = true;
      }
      if (xfrmEl.attributes["flipV"] === "1") {
        shape.flipVertical = true;
      }
    }
  }

  // Parse text content (wps:txbx > w:txbxContent)
  const txbxEl = findChild(wspEl, "wps:txbx") ?? findChildNs(wspEl, "txbx");
  const txbxContentEl = txbxEl
    ? (findChild(txbxEl, "w:txbxContent") ?? findChildNs(txbxEl, "txbxContent"))
    : undefined;
  if (txbxContentEl) {
    const paragraphs = expandBlockSimpleFields(txbxContentEl.children).filter(
      (child): child is XmlElement => child.type === "element" && stripWPrefix(child.name) === "p"
    );
    if (paragraphs.length > 0) {
      if (pendingText) {
        // Its place kept, so the shape's fields are in the order a read of them all at once gives.
        const textContent: Paragraph[] = [];
        shape.textContent = textContent;
        pendingText.push({ textContent, paragraphs });
      } else {
        shape.textContent = paragraphs.map(child => parseParagraph(child, ctx));
      }
    }
  }

  // Parse text body vertical anchor (wps:bodyPr/@anchor)
  const bodyPrEl = findChild(wspEl, "wps:bodyPr") ?? findChildNs(wspEl, "bodyPr");
  const anchorAttr = bodyPrEl?.attributes["anchor"];
  if (
    anchorAttr === "t" ||
    anchorAttr === "ctr" ||
    anchorAttr === "b" ||
    anchorAttr === "just" ||
    anchorAttr === "dist"
  ) {
    shape.textBodyAnchor = anchorAttr;
  }

  // Parse positioning
  const posH = findChild(anchorEl, "wp:positionH");
  if (posH) {
    const hp: Mutable<NonNullable<DrawingShape["horizontalPosition"]>> = {};
    hp.relativeTo = readToken(
      posH.attributes["relativeFrom"],
      WRITTEN.horizontalRelative,
      hp,
      "relativeTo"
    ) as NonNullable<DrawingShape["horizontalPosition"]>["relativeTo"];
    const offsetEl = findChild(posH, "wp:posOffset");
    if (offsetEl) {
      hp.offset = parseInt(textContent(offsetEl), 10);
    }
    const alignEl = findChild(posH, "wp:align");
    if (alignEl) {
      hp.align = readToken(
        textContent(alignEl),
        WRITTEN.horizontalAlign,
        hp,
        "align"
      ) as NonNullable<DrawingShape["horizontalPosition"]>["align"];
    }
    shape.horizontalPosition = hp;
  }
  const posV = findChild(anchorEl, "wp:positionV");
  if (posV) {
    const vp: Mutable<NonNullable<DrawingShape["verticalPosition"]>> = {};
    vp.relativeTo = readToken(
      posV.attributes["relativeFrom"],
      WRITTEN.verticalRelative,
      vp,
      "relativeTo"
    ) as NonNullable<DrawingShape["verticalPosition"]>["relativeTo"];
    const offsetEl = findChild(posV, "wp:posOffset");
    if (offsetEl) {
      vp.offset = parseInt(textContent(offsetEl), 10);
    }
    const alignEl = findChild(posV, "wp:align");
    if (alignEl) {
      vp.align = readToken(textContent(alignEl), WRITTEN.verticalAlign, vp, "align") as NonNullable<
        DrawingShape["verticalPosition"]
      >["align"];
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
      const wrap: Mutable<NonNullable<DrawingShape["wrap"]>> = {
        style: "square"
      };
      wrap.side = readToken(wrapChild.attributes["wrapText"], WRITTEN.wrapSide, wrap, "side") as
        | NonNullable<DrawingShape["wrap"]>["side"]
        | undefined;
      shape.wrap = wrap;
    } else if (wn === "wp:wrapTight") {
      const wrap: Mutable<NonNullable<DrawingShape["wrap"]>> = {
        style: "tight"
      };
      wrap.side = readToken(wrapChild.attributes["wrapText"], WRITTEN.wrapSide, wrap, "side") as
        | NonNullable<DrawingShape["wrap"]>["side"]
        | undefined;
      shape.wrap = wrap;
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

  const preserved = preserveShapeMarkup(wspEl, shape);
  if (preserved) {
    shape._preserved = preserved;
  }
  shape._frame = readDrawingFrame(anchorEl, frameFieldsOf(shape));
  return shape;
}

// =============================================================================
// TextBox Parser
// =============================================================================

function parseTextBox(pictEl: XmlElement, ctx: ReaderContext): TextBox | undefined {
  const parts = textBoxParts(pictEl);
  return (
    parts &&
    finishTextBox(
      parts.shapeEl,
      parts.paragraphs.map(c => parseParagraph(c, ctx))
    )
  );
}

/** {@link parseTextBox} as steps, a paragraph of the box at a time. */
function* textBoxSteps(pictEl: XmlElement, ctx: ReaderContext): Steps<TextBox | undefined> {
  const parts = textBoxParts(pictEl);
  if (parts === undefined) {
    return undefined;
  }
  const paragraphs: Paragraph[] = [];
  for (const paragraph of parts.paragraphs) {
    paragraphs.push(yield* paragraphSteps(paragraph, ctx));
  }
  return finishTextBox(parts.shapeEl, paragraphs);
}

/** A VML text box's parts — `v:shape` > `v:textbox` > `w:txbxContent` — or `undefined` when `pictEl` holds none. */
function textBoxParts(
  pictEl: XmlElement
): { readonly shapeEl: XmlElement | undefined; readonly paragraphs: XmlElement[] } | undefined {
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
  const paragraphs = txbxContentEl.children.filter(
    (c): c is XmlElement => c.type === "element" && stripWPrefix(c.name) === "p"
  );
  return { shapeEl, paragraphs };
}

/** The text box the shape `shapeEl` makes with its `paragraphs` read. */
function finishTextBox(shapeEl: XmlElement | undefined, paragraphs: Paragraph[]): TextBox {
  const tb: Mutable<TextBox> = { type: "textBox", content: paragraphs };

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

/**
 * A block-level SDT as steps — one per block it holds, see {@link tableSteps} — so the streaming reader can give
 * the event loop a turn inside a large one.
 */
function* sdtSteps(
  sdtEl: XmlElement,
  ctx: ReaderContext
): Steps<StructuredDocumentTag | CheckBox | TableOfContents | undefined> {
  const sdtPrEl = findChildNs(sdtEl, "sdtPr");
  const sdtContentEl = findChildNs(sdtEl, "sdtContent");

  // Check for checkbox (w14:checkbox)
  if (sdtPrEl) {
    const checkBoxEl = findChild(sdtPrEl, "w14:checkbox");
    if (checkBoxEl) {
      return parseCheckBox(checkBoxEl);
    }
  }

  if (isTocSdt(sdtEl)) {
    return yield* tocSteps(sdtContentEl, ctx);
  }
  return yield* genericSdtSteps(sdtEl, ctx);
}

/** {@link sdtSteps} for a read that does not yield: the same control, without the steps. */
function parseSdt(
  sdtEl: XmlElement,
  ctx: ReaderContext
): StructuredDocumentTag | CheckBox | TableOfContents | undefined {
  const sdtPrEl = findChildNs(sdtEl, "sdtPr");
  const checkBoxEl = sdtPrEl ? findChild(sdtPrEl, "w14:checkbox") : undefined;
  if (checkBoxEl) {
    return parseCheckBox(checkBoxEl);
  }
  // A table of contents is one in a document: its steps are run.
  return isTocSdt(sdtEl)
    ? runSteps(tocSteps(findChildNs(sdtEl, "sdtContent"), ctx))
    : parseGenericSdt(sdtEl, ctx);
}

/**
 * The {@link CheckBox} a paragraph is when it holds nothing but a checkbox
 * content control with nothing beside `w14:checkbox` in its properties —
 * the form the writer gives a checkbox block — or `undefined`.
 */
function paragraphOnlyCheckBox(p: XmlElement): CheckBox | undefined {
  // The paragraph's one element, looked for without listing them all: a paragraph of runs stops at its second.
  let only: XmlElement | undefined;
  for (const c of p.children) {
    if (c.type === "element") {
      if (only !== undefined) {
        return undefined;
      }
      only = c;
    }
  }
  if (only === undefined || stripWPrefix(only.name) !== "sdt") {
    return undefined;
  }
  const sdtPr = findChildNs(only, "sdtPr");
  const properties = sdtPr?.children.filter((c): c is XmlElement => c.type === "element") ?? [];
  if (properties.length !== 1 || properties[0]!.name !== "w14:checkbox") {
    return undefined;
  }
  return parseCheckBox(properties[0]!);
}

/**
 * An SDT read as a generic content control: its properties as
 * {@link SdtProperties} and its content as it is. A checkbox or TOC nested in
 * another SDT, where the model has no place for one, is read this way and
 * keeps its content.
 */
function* genericSdtSteps(sdtEl: XmlElement, ctx: ReaderContext): Steps<StructuredDocumentTag> {
  const props = parseSdtProperties(findChildNs(sdtEl, "sdtPr"));
  const items: SdtItem[] = [];
  let runs = 0;
  for (const child of sdtContentChildren(sdtEl)) {
    if (child.type !== "element" || takeRangeMarker(child, items)) {
      continue;
    }
    const n = stripWPrefix(child.name);
    if (n === "p") {
      items.push(yield* paragraphSteps(child, ctx));
    } else if (n === "tbl") {
      items.push(yield* tableSteps(child, ctx));
    } else if (n === "r") {
      items.push(parseRun(child));
      // Runs standing in the content are blocks of no size: a step ends after a number of them.
      if (++runs % PARAGRAPH_STEP === 0) {
        yield;
      }
    } else if (n === "sdt") {
      items.push(yield* genericSdtSteps(child, ctx));
    }
  }
  return finishGenericSdt(props, items);
}

/** {@link genericSdtSteps} for a read that does not yield. */
function parseGenericSdt(sdtEl: XmlElement, ctx: ReaderContext): StructuredDocumentTag {
  const props = parseSdtProperties(findChildNs(sdtEl, "sdtPr"));
  const items: SdtItem[] = [];
  for (const child of sdtContentChildren(sdtEl)) {
    if (child.type !== "element" || takeRangeMarker(child, items)) {
      continue;
    }
    const n = stripWPrefix(child.name);
    if (n === "p") {
      items.push(parseParagraph(child, ctx));
    } else if (n === "tbl") {
      items.push(parseTable(child, ctx));
    } else if (n === "r") {
      items.push(parseRun(child));
    } else if (n === "sdt") {
      items.push(parseGenericSdt(child, ctx));
    }
  }
  return finishGenericSdt(props, items);
}

/** What a generic content control holds, as it is read. */
type SdtItem = Paragraph | Run | Table | StructuredDocumentTag | RangeMarker;

/**
 * The children of a content control's `w:sdtContent`, a block-level simple field read as the complex field it stands
 * for. A nested SDT (e.g. repeating section item SDTs) is read as a generic one, its own properties kept, so data
 * binding, alias, lock and similar metadata round-trip; a nested checkbox or TOC has no place in SDT content as such,
 * so it is read as a generic SDT and keeps its content — a TOC its field too, read in pieces where it spans paragraphs.
 */
function sdtContentChildren(sdtEl: XmlElement): XmlNode[] {
  const sdtContentEl = findChildNs(sdtEl, "sdtContent");
  return sdtContentEl ? expandBlockSimpleFields(sdtContentEl.children) : [];
}

/** The content control `props` and `items` make. */
function finishGenericSdt(props: SdtProperties, items: SdtItem[]): StructuredDocumentTag {
  const unhosted: RangeMarker[] = [];
  const sdt: Mutable<StructuredDocumentTag> = {
    type: "sdt",
    properties: props,
    content: foldRangeMarkers(items, unhosted)
  };
  if (unhosted.length > 0) {
    sdt.contentRangeMarkersAfter = unhosted;
  }
  return sdt;
}

/** A content control's `w:sdtPr` as {@link SdtProperties}. */
function parseSdtProperties(sdtPrEl: XmlElement | undefined): Mutable<SdtProperties> {
  const props: Mutable<SdtProperties> = {};
  if (sdtPrEl) {
    const tagEl = findChildNs(sdtPrEl, "tag");
    if (tagEl) {
      props.tag = attrVal(tagEl, "val");
    }
    const aliasEl = findChildNs(sdtPrEl, "alias");
    if (aliasEl) {
      props.alias = attrVal(aliasEl, "val");
    }
    const lockEl = findChildNs(sdtPrEl, "lock");
    if (lockEl) {
      const v = attrVal(lockEl, "val");
      if (v === "contentLocked" || v === "sdtContentLocked") {
        props.lockContent = true;
      }
      if (v === "sdtLocked" || v === "sdtContentLocked") {
        props.lockSdt = true;
      }
    }
    // Plain text
    if (findChildNs(sdtPrEl, "text")) {
      props.plainText = true;
    }
    // showingPlcHdr is a toggle, not a property with a val
    if (findChildNs(sdtPrEl, "showingPlcHdr")) {
      const v = boolToggle(sdtPrEl, "showingPlcHdr");
      if (v !== false) {
        props.showingPlaceholder = true;
      }
    }
    // w15:appearance (replaces the old misused showingPlcHdr)
    const appearanceEl = findChild(sdtPrEl, "w15:appearance");
    if (appearanceEl) {
      const v = appearanceEl.attributes["w15:val"] ?? appearanceEl.attributes["val"];
      if (v === "boundingBox" || v === "tags" || v === "hidden") {
        props.appearance = v;
      }
    }
    // Dropdown list
    const ddlEl = findChildNs(sdtPrEl, "dropDownList");
    if (ddlEl) {
      const items: SdtListItem[] = [];
      for (const li of findChildrenNs(ddlEl, "listItem")) {
        const item: Partial<Mutable<SdtListItem>> = {
          value: attrVal(li, "value") ?? ""
        };
        const dt = attrVal(li, "displayText");
        if (dt) {
          item.displayText = dt;
        }
        items.push(item as SdtListItem);
      }
      props.dropdownList = items;
    }
    // ComboBox
    const cbEl = findChildNs(sdtPrEl, "comboBox");
    if (cbEl) {
      const items: SdtListItem[] = [];
      for (const li of findChildrenNs(cbEl, "listItem")) {
        const item: Partial<Mutable<SdtListItem>> = {
          value: attrVal(li, "value") ?? ""
        };
        const dt = attrVal(li, "displayText");
        if (dt) {
          item.displayText = dt;
        }
        items.push(item as SdtListItem);
      }
      props.comboBox = items;
    }
    // Date picker
    const dateEl = findChildNs(sdtPrEl, "date");
    if (dateEl) {
      const dateProp: Partial<Mutable<SdtDateProperties>> = {};
      const fullDate = attrVal(dateEl, "fullDate");
      if (fullDate) {
        dateProp.fullDate = fullDate;
      }
      const dfEl = findChildNs(dateEl, "dateFormat");
      if (dfEl) {
        dateProp.dateFormat = attrVal(dfEl, "val");
      }
      const lidEl = findChildNs(dateEl, "lid");
      if (lidEl) {
        dateProp.lid = attrVal(lidEl, "val");
      }
      const storeEl = findChildNs(dateEl, "storeMappedDataAs");
      if (storeEl) {
        dateProp.storeMappedDataAs = attrVal(
          storeEl,
          "val"
        ) as SdtDateProperties["storeMappedDataAs"];
      }
      props.date = dateProp;
    }
    // ID
    const idEl = findChildNs(sdtPrEl, "id");
    if (idEl) {
      const v = attrInt(idEl, "val");
      if (v !== undefined) {
        props.id = v;
      }
    }
    // Data binding
    const dbEl = findChildNs(sdtPrEl, "dataBinding");
    if (dbEl) {
      const xpath = attrVal(dbEl, "xpath");
      const storeItemId = attrVal(dbEl, "storeItemID");
      if (xpath && storeItemId) {
        const binding: {
          xpath: string;
          storeItemId: string;
          prefixMappings?: string;
        } = {
          xpath,
          storeItemId
        };
        const prefixMappings = attrVal(dbEl, "prefixMappings");
        if (prefixMappings) {
          binding.prefixMappings = prefixMappings;
        }
        props.dataBinding = binding;
      }
    }
    // Placeholder
    const phEl = findChildNs(sdtPrEl, "placeholder");
    if (phEl) {
      const docPartEl = findChildNs(phEl, "docPart");
      if (docPartEl) {
        props.placeholder = attrVal(docPartEl, "val");
      }
    }
    // Boolean marker elements
    if (findChildNs(sdtPrEl, "richText")) {
      props.richText = true;
    }
    if (findChildNs(sdtPrEl, "picture")) {
      props.picture = true;
    }
    if (findChildNs(sdtPrEl, "group")) {
      props.group = true;
    }
    if (findChildNs(sdtPrEl, "equation")) {
      props.equation = true;
    }
    if (findChildNs(sdtPrEl, "citation")) {
      props.citation = true;
    }
    if (findChildNs(sdtPrEl, "bibliography")) {
      props.bibliography = true;
    }
    if (findChildNs(sdtPrEl, "temporary")) {
      props.temporary = true;
    }
    // w15: repeating section
    const rsEl = findChild(sdtPrEl, "w15:repeatingSection");
    if (rsEl) {
      const rs: { sectionTitle?: string; allowInsertDelete?: boolean } = {};
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
      props.repeatingSection = rs;
    }
    if (findChild(sdtPrEl, "w15:repeatingSectionItem")) {
      props.repeatingSectionItem = true;
    }
    // Building-block gallery entry (a cover page, a page number…)
    const docPartEl = findChildNs(sdtPrEl, "docPartObj") ?? findChildNs(sdtPrEl, "docPartList");
    if (docPartEl) {
      const docPart: Mutable<SdtDocPart> = {};
      const gallery = findChildNs(docPartEl, "docPartGallery");
      const galleryVal = gallery ? attrVal(gallery, "val") : undefined;
      if (galleryVal !== undefined) {
        docPart.gallery = galleryVal;
      }
      const category = findChildNs(docPartEl, "docPartCategory");
      const categoryVal = category ? attrVal(category, "val") : undefined;
      if (categoryVal !== undefined) {
        docPart.category = categoryVal;
      }
      if (
        findChildNs(docPartEl, "docPartUnique") &&
        boolToggle(docPartEl, "docPartUnique") !== false
      ) {
        docPart.unique = true;
      }
      if (stripWPrefix(docPartEl.name) === "docPartList") {
        docPart.list = true;
      }
      props.docPart = docPart;
    }
  }
  return props;
}

/**
 * Whether an SDT is a table of contents (`docPartGallery` "Table of Contents").
 */
function isTocSdt(sdtEl: XmlElement): boolean {
  const sdtPrEl = findChildNs(sdtEl, "sdtPr");
  const docPartObjEl = sdtPrEl ? findChildNs(sdtPrEl, "docPartObj") : undefined;
  const galleryEl = docPartObjEl ? findChildNs(docPartObjEl, "docPartGallery") : undefined;
  return galleryEl !== undefined && attrVal(galleryEl, "val") === "Table of Contents";
}

/**
 * The content of a block-level SDT, for a container whose model holds blocks
 * but no SDT (a table cell, a header or footer, a note). The SDT's own
 * properties are dropped; its content is kept whatever kind of control it is,
 * nested SDTs are flattened the same way, and runs standing directly in the
 * content are gathered into a paragraph. A note holds paragraphs only, so
 * `tables: false` keeps just a table's range markers, as before.
 *
 * A TOC keeps its field too: read field character by field character (see
 * `FieldContent.fieldChar`), it spans the paragraphs it did.
 */
function flattenSdt(
  sdtEl: XmlElement,
  ctx: ReaderContext,
  options: { tables: false }
): (Paragraph | RangeMarker)[];
function flattenSdt(
  sdtEl: XmlElement,
  ctx: ReaderContext,
  options: {
    tables: true;
    controls?: { list: BlockContentControl[]; base: number };
  }
): (Paragraph | Table | RangeMarker)[];
function flattenSdt(
  sdtEl: XmlElement,
  ctx: ReaderContext,
  options: {
    tables: boolean;
    controls?: { list: BlockContentControl[]; base: number };
  }
): (Paragraph | Table | RangeMarker)[] {
  const read = startFlatten(options);
  flattenInto(read, sdtEl, ctx);
  return read.out;
}

/** The content control `el` read into `read` by {@link flattenSdt}, the controls it holds with it. */
function flattenInto(read: FlattenRead, el: XmlElement, ctx: ReaderContext): void {
  const control = flattenOpen(read, el);
  for (const child of sdtContentChildren(el)) {
    if (child.type !== "element" || flattenTake(read, child)) {
      continue;
    }
    const name = stripWPrefix(child.name);
    if (name === "p") {
      flattenPush(read, parseParagraph(child, ctx));
    } else if (name === "tbl") {
      flattenTable(read, parseTable(child, ctx));
    } else if (name === "sdt") {
      flattenInto(read, child, ctx);
    }
  }
  flattenClose(read, control);
}

/** {@link flattenSdt} as steps — one per block its content holds, see {@link tableSteps}. */
function* flattenSdtSteps(
  sdtEl: XmlElement,
  ctx: ReaderContext,
  options: FlattenOptions
): Steps<FlattenedItem[]> {
  const read = startFlatten(options);
  function* flatten(el: XmlElement): Steps<void> {
    const control = flattenOpen(read, el);
    for (const child of sdtContentChildren(el)) {
      if (child.type !== "element") {
        continue;
      }
      if (flattenTake(read, child)) {
        // Runs standing in the content are blocks of no size: a step ends after a number of them.
        if (read.runCount % PARAGRAPH_STEP === 0 && stripWPrefix(child.name) === "r") {
          yield;
        }
        continue;
      }
      const name = stripWPrefix(child.name);
      if (name === "p") {
        flattenPush(read, yield* paragraphSteps(child, ctx));
      } else if (name === "tbl") {
        flattenTable(read, yield* tableSteps(child, ctx));
      } else if (name === "sdt") {
        yield* flatten(child);
      }
    }
    flattenClose(read, control);
  }
  yield* flatten(sdtEl);
  return read.out;
}

type FlattenedItem = Paragraph | Table | RangeMarker;

interface FlattenOptions {
  readonly tables: boolean;
  readonly controls?: { list: BlockContentControl[]; base: number };
}

/**
 * A {@link flattenSdt} under way: what it has read, shared by the read that yields and the one that does not. With
 * `controls`, each SDT is recorded around the blocks it held, by their index among the container's blocks (its range
 * markers are folded into blocks, so they do not count). A plain record and functions over it: one is made per
 * content control in a cell.
 */
interface FlattenRead {
  readonly out: FlattenedItem[];
  readonly options: FlattenOptions;
  /** The container's blocks read so far. */
  blocks: number;
  /** The paragraph the runs just read are gathered into, while runs follow one another. */
  runs: Run[] | undefined;
  runCount: number;
}

function startFlatten(options: FlattenOptions): FlattenRead {
  return { out: [], options, blocks: 0, runs: undefined, runCount: 0 };
}

function flattenAdd(read: FlattenRead, item: FlattenedItem): void {
  read.out.push(item);
  if (!isRangeMarker(item)) {
    read.blocks++;
  }
}

/** A block read in the content: the runs after it begin a paragraph of their own. */
function flattenPush(read: FlattenRead, item: FlattenedItem): void {
  read.runs = undefined;
  flattenAdd(read, item);
}

function flattenTable(read: FlattenRead, table: Table): void {
  read.runs = undefined;
  if (read.options.tables) {
    flattenAdd(read, table);
  } else {
    for (const item of table.rangeMarkersBefore ?? []) {
      flattenAdd(read, item);
    }
  }
}

/** A run, gathered into the paragraph the runs before it began, or a range marker: whether `child` was one. */
function flattenTake(read: FlattenRead, child: XmlElement): boolean {
  if (stripWPrefix(child.name) === "r") {
    if (!read.runs) {
      read.runs = [];
      flattenAdd(read, { type: "paragraph", children: read.runs });
    }
    read.runs.push(parseRun(child));
    read.runCount++;
    return true;
  }
  read.runs = undefined;
  return takeRangeMarker(child, read.out);
}

/** The content control `el` begun: its record, when they are kept. */
function flattenOpen(read: FlattenRead, el: XmlElement): Mutable<BlockContentControl> | undefined {
  read.runs = undefined;
  const controls = read.options.controls;
  if (!controls) {
    return undefined;
  }
  const properties = parseSdtProperties(findChildNs(el, "sdtPr"));
  const control: Mutable<BlockContentControl> = {
    start: controls.base + read.blocks,
    end: controls.base + read.blocks
  };
  if (Object.keys(properties).length > 0) {
    control.properties = properties;
  }
  controls.list.push(control);
  return control;
}

/** The content control `control` records ended. */
function flattenClose(read: FlattenRead, control: Mutable<BlockContentControl> | undefined): void {
  // The runs after a content control begin a paragraph of their own, as they began one before it.
  read.runs = undefined;
  const controls = read.options.controls;
  if (control && controls) {
    control.end = controls.base + read.blocks;
    if (control.end === control.start) {
      // Holding no block, it has nowhere to be written back in a cell.
      controls.list.splice(controls.list.indexOf(control), 1);
    }
  }
}

function* tocSteps(
  sdtContentEl: XmlElement | undefined,
  ctx: ReaderContext
): Steps<TableOfContents> {
  const toc: Mutable<TableOfContents> = { type: "tableOfContents" };
  if (!sdtContentEl) {
    return toc;
  }
  // A block-level simple field is read as the complex field it stands for.
  const paragraphs = expandBlockSimpleFields(sdtContentEl.children).filter(
    (child): child is XmlElement => child.type === "element" && stripWPrefix(child.name) === "p"
  );

  // The paragraphs are read as they are, the block's TOC field among them in
  // pieces where it is — Word opens it in the first entry's paragraph — so
  // the block is written back with the paragraphs it had.
  const cachedParagraphs: Paragraph[] = [];
  const savedField = ctx.field;
  ctx.field = createFieldState();
  const tocBegin = firstTocBegin(paragraphs);
  if (tocBegin) {
    ctx.field.forced = new Set([tocBegin]);
  }
  try {
    for (const p of paragraphs) {
      cachedParagraphs.push(ctx.paced ? yield* paragraphSteps(p, ctx) : parseParagraph(p, ctx));
    }
  } finally {
    ctx.field = savedField;
  }

  const field = findTocFieldPieces(cachedParagraphs);
  if (field) {
    toc.field = "entries";
    toc.instruction = field.instruction;
    parseTocInstruction(field.instruction, toc);
    if (!field.closed) {
      // The field cannot run on past its block: it ends with the block.
      const last = cachedParagraphs.length - 1;
      const end: Run = {
        content: [{ type: "field", fieldChar: "end", instruction: "" }]
      };
      cachedParagraphs[last] = {
        ...cachedParagraphs[last]!,
        children: [...cachedParagraphs[last]!.children, end]
      };
    }
  } else {
    // No field of its own — a hand-written table: the writer adds none.
    toc.field = "none";
  }
  if (cachedParagraphs.length > 0) {
    toc.cachedParagraphs = cachedParagraphs;
  }
  return toc;
}

/**
 * The `w:fldChar` that begins a TOC block's own field: the first field the
 * paragraphs' runs open, when its instruction is a `TOC` one.
 */
function firstTocBegin(paragraphs: readonly XmlElement[]): XmlElement | undefined {
  let begin: XmlElement | undefined;
  let instruction = "";
  for (const p of paragraphs) {
    for (const child of p.children) {
      if (child.type !== "element" || stripWPrefix(child.name) !== "r") {
        continue;
      }
      for (const rc of child.children) {
        if (rc.type !== "element") {
          continue;
        }
        const name = stripWPrefix(rc.name);
        if (name === "instrText" && begin) {
          instruction += textContent(rc);
        } else if (name === "fldChar") {
          if (begin) {
            return /^\s*TOC\b/i.test(instruction) ? begin : undefined;
          }
          if (attrVal(rc, "fldCharType") === "begin") {
            begin = rc;
          }
        }
      }
    }
  }
  return begin && /^\s*TOC\b/i.test(instruction) ? begin : undefined;
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

/** How many children of a run {@link parseRunSteps} reads in one step. */
const RUN_STEP = 4096;

/** A `w:ruby`: its properties, and the runs of its text (`w:rt`) and of its base (`w:rubyBase`). */
function parseRuby(child: XmlElement): RubyContent {
  const ruby = startRuby(child);
  for (const [runs, list] of rubyRunLists(child, ruby)) {
    for (const run of runs) {
      list.push(parseRun(run));
    }
  }
  return ruby;
}

/** {@link parseRuby} in steps of {@link PARAGRAPH_STEP} runs, for a run read in steps. */
function* rubySteps(child: XmlElement): Steps<RubyContent> {
  const ruby = startRuby(child);
  let read = 0;
  for (const [runs, list] of rubyRunLists(child, ruby)) {
    for (const run of runs) {
      list.push(parseRun(run));
      if (++read % PARAGRAPH_STEP === 0) {
        yield;
      }
    }
  }
  return ruby;
}

/** Whether `child` is a `w:ruby` holding more runs than a step ({@link PARAGRAPH_STEP}). */
function isLongRuby(child: XmlElement): boolean {
  if (stripWPrefix(child.name) !== "ruby") {
    return false;
  }
  let count = 0;
  for (const [runs] of rubyRunLists(child, startRuby(child, false))) {
    count += runs.length;
  }
  return count > PARAGRAPH_STEP;
}

/** A ruby's runs to read — the text's (`w:rt`), then the base's (`w:rubyBase`) — each with the list it is read into. */
function rubyRunLists(
  child: XmlElement,
  ruby: Mutable<RubyContent> & { rubyText: Run[]; baseText: Run[] }
): [XmlElement[], Run[]][] {
  const runsOf = (el: XmlElement | undefined): XmlElement[] =>
    el
      ? el.children.filter(
          (c): c is XmlElement => c.type === "element" && stripWPrefix(c.name) === "r"
        )
      : [];
  return [
    [runsOf(findChildNs(child, "rt")), ruby.rubyText],
    [runsOf(findChildNs(child, "rubyBase")), ruby.baseText]
  ];
}

/** A ruby with its properties read and no runs yet — or, without `withProperties`, a ruby to count runs into. */
function startRuby(
  child: XmlElement,
  withProperties = true
): Mutable<RubyContent> & { rubyText: Run[]; baseText: Run[] } {
  const ruby: Mutable<RubyContent> & {
    rubyText: Run[];
    baseText: Run[];
  } = {
    type: "ruby",
    rubyText: [],
    baseText: []
  };
  const rubyPrEl = withProperties ? findChildNs(child, "rubyPr") : undefined;
  if (rubyPrEl) {
    const props: Mutable<RubyProperties> = {};
    const alignEl = findChildNs(rubyPrEl, "rubyAlign");
    if (alignEl) {
      props.align = attrVal(alignEl, "val") as RubyProperties["align"];
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
  return ruby;
}

/**
 * {@link parseRun} in steps of {@link RUN_STEP} children, for a run read in steps: a run is one element however much
 * it holds, and text imported as one run of a thousand lines and breaks holds a child per line.
 */
function* parseRunSteps(el: XmlElement, read: RunContentReader = parseRunContent): Steps<Run> {
  const rPrEl = findChildNs(el, "rPr");
  const content: RunContent[] = [];
  for (let from = 0; from < el.children.length; from += RUN_STEP) {
    const to = Math.min(from + RUN_STEP, el.children.length);
    let start = from;
    for (let at = from; at < to; at++) {
      const child = el.children[at]!;
      if (child.type === "element" && read === parseRunContent && isLongRuby(child)) {
        read(el, content, start, at);
        content.push(yield* rubySteps(child));
        start = at + 1;
      }
    }
    read(el, content, start, to);
    yield;
  }
  return { properties: rPrEl ? parseRunProperties(rPrEl) : undefined, content };
}

/**
 * Whether the run `el` is longer than a step of its own: more than {@link RUN_STEP} children, or a ruby of more than a
 * step of runs ({@link isLongRuby}).
 */
function isLongRun(el: XmlElement): boolean {
  return (
    el.children.length > RUN_STEP || el.children.some(c => c.type === "element" && isLongRuby(c))
  );
}

/** How a run's content is read, from child `from` to `to`: `parseRunContent`, or `readDeletedRunContent`. */
type RunContentReader = (
  el: XmlElement,
  content: RunContent[],
  from: number,
  to: number
) => RunContent[];

/** {@link parseRunParts} as steps, each part's content read by {@link parseRunSteps}. */
function* runPartsSteps(el: XmlElement, read: RunContentReader): Steps<(Run | CommentReference)[]> {
  const parts = runParts(el);
  if (parts === undefined) {
    return [yield* parseRunSteps(el, read)];
  }
  const out: (Run | CommentReference)[] = [];
  for (const part of parts) {
    out.push(part.type === "element" ? yield* parseRunSteps(part, read) : part);
  }
  return out;
}

/** Whether `el` is one of a split field's marker runs, read as the piece or the instruction it stands for. */
function isFieldMarkerRun(el: XmlElement, marked: boolean): boolean {
  return marked && (fieldPieceOf(el) !== undefined || fieldCodeTextOf(el) !== undefined);
}

/**
 * A range marker met inside an open complex field. It is kept in the field
 * ({@link FieldContent.rangeMarkers}) when the field ends in the same
 * paragraph, and in place among the paragraph's children otherwise.
 */
interface FieldMarkerSlot {
  readonly fieldMarker: PendingFieldMarker;
}

/** What {@link readParagraphContent} collects: children, markers, and field marker slots. */
type ParagraphItem = ParagraphChild | RangeMarker | FieldMarkerSlot;

function isFieldMarkerSlot(item: ParagraphItem): item is FieldMarkerSlot {
  return "fieldMarker" in item;
}

/**
 * The parts of a run inside a revision: a field piece's marker run (see
 * `splitFieldPieces`) is that piece, any other run as `parseRunParts` reads it.
 */
function revisionRunParts(
  el: XmlElement,
  marked: boolean,
  parse: (el: XmlElement) => Run = parseRun
): (Run | CommentReference)[] {
  const piece = marked ? fieldPieceOf(el) : undefined;
  if (piece) {
    return [fieldPieceRun(piece)];
  }
  const code = marked ? fieldCodeTextOf(el) : undefined;
  return code !== undefined ? [fieldCodeRun(el, code)] : parseRunParts(el, parse);
}

/**
 * A `w:r` as the paragraph children it stands for. A comment's reference
 * mark (`w:commentReference`) is the paragraph child {@link CommentReference}
 * in the model, not run content, so a run holding one is split at each mark:
 * the content before and after it become runs of their own with the run's
 * formatting, and the mark a reference carrying that formatting too. The
 * anchor keeps its comment id and its place. A run without a mark is read as
 * one run, as `parse` reads it; a mark without an id anchors nothing and is
 * dropped.
 */
function parseRunParts(
  el: XmlElement,
  parse: (el: XmlElement) => Run = parseRun
): (Run | CommentReference)[] {
  const parts = runParts(el);
  return parts === undefined
    ? [parse(el)]
    : parts.map(part => (part.type === "element" ? parse(part) : part));
}

/**
 * The run `el` split at its comment marks: the runs around them, each with the run's formatting, and a reference per
 * mark carrying that formatting too — or `undefined` when it holds none, and is one run.
 */
function runParts(el: XmlElement): (XmlElement | CommentReference)[] | undefined {
  // Asked of every child of every run read.
  const isMark = (node: XmlNode): node is XmlElement =>
    node.type === "element" &&
    (node.name === "w:commentReference" || node.name === "commentReference");
  if (!el.children.some(isMark)) {
    return undefined;
  }
  const rPrEl = findChildNs(el, "rPr");
  const parts: (XmlElement | CommentReference)[] = [];
  let group: XmlNode[] = [];
  const flush = (): void => {
    if (group.some(node => node.type === "element")) {
      parts.push({ ...el, children: rPrEl ? [rPrEl, ...group] : group });
    }
    group = [];
  };
  for (const child of el.children) {
    if (child === rPrEl) {
      continue;
    }
    if (!isMark(child)) {
      group.push(child);
      continue;
    }
    const id = attrInt(child, "id");
    if (id === undefined) {
      continue;
    }
    flush();
    const reference: Mutable<CommentReference> = {
      type: "commentReference",
      id
    };
    if (rPrEl) {
      reference.properties = parseRunProperties(rPrEl);
    }
    parts.push(reference);
  }
  flush();
  return parts;
}

function parseParagraph(pEl: XmlElement, ctx: ReaderContext): Paragraph {
  const field = ctx.field;
  const read = beginParagraph(
    pEl,
    ctx,
    field.noFieldChars ? pEl : splitFieldPieces(pEl, field.open, field.forced)
  );
  try {
    readParagraphContent(read.split, ctx, read.items);
  } finally {
    ctx.field.marked = read.marked;
  }
  return finishParagraph(pEl, ctx, read.items);
}

/** How many elements of a paragraph's content {@link paragraphSteps} reads in one step. */
const PARAGRAPH_STEP = 256;

/**
 * {@link parseParagraph} as steps, for a read run paced (`ReaderContext.paced`; a read that is not calls
 * `parseParagraph`): a step ends after the paragraph, and one larger than {@link PARAGRAPH_STEP} elements of content —
 * counted through the wrappers it holds, a hyperlink, a tracked change, a smart tag — is read in steps of that many, so
 * a paragraph of tens of thousands of runs is too. One no larger — nearly every one — is read by `parseParagraph`: its
 * content read through generators measured 4–5% slower.
 *
 * Not split: the paragraph's finishing after the last step, under 3 ms per 100,000 runs.
 */
function* paragraphSteps(pEl: XmlElement, ctx: ReaderContext): Steps<Paragraph> {
  let paragraph: Paragraph;
  if (exceedsStep(pEl)) {
    const field = ctx.field;
    const read = beginParagraph(
      pEl,
      ctx,
      field.noFieldChars ? pEl : yield* splitFieldPiecesSteps(pEl, field.open, field.forced)
    );
    try {
      yield* paragraphContentSteps(read.split, ctx, read.items, { read: 0 });
    } finally {
      ctx.field.marked = read.marked;
    }
    paragraph = finishParagraph(pEl, ctx, read.items);
  } else {
    paragraph = parseParagraph(pEl, ctx);
  }
  yield;
  return paragraph;
}

/**
 * Whether the content of `pEl` holds more than {@link PARAGRAPH_STEP} elements, counting into everything but its runs —
 * anything else may be a wrapper read in steps or markup kept in steps — or a run longer than a step of its own
 * ({@link isLongRun}). Stops counting at the step, so it costs a short paragraph no more than a look at its children.
 */
function exceedsStep(pEl: XmlElement): boolean {
  let count = 0;
  const pending: XmlElement[] = [pEl];
  while (pending.length > 0) {
    for (const child of pending.pop()!.children) {
      if (child.type !== "element") {
        continue;
      }
      if (++count > PARAGRAPH_STEP) {
        return true;
      }
      const resolved = resolveContentChild(child);
      if (resolved === undefined) {
        continue;
      }
      if (stripWPrefix(resolved.name) !== "r") {
        pending.push(resolved);
      } else if (isLongRun(resolved)) {
        return true;
      }
    }
  }
  return false;
}

/**
 * {@link readParagraphContent} as steps of {@link PARAGRAPH_STEP} elements, counted in `budget` across the wrappers
 * it reads into. Each element is read by `readParagraphContent` itself — one child at a time — except a wrapper,
 * whose content this reads the same way, between the same markers and with the same field state as
 * `readParagraphContent` does: the pieces of each wrapper (`finishHyperlink`, `readRevisionContent`, `WRAPPERS`)
 * are shared, not copied. Markup `readParagraphContent` keeps as it is is serialised in steps too.
 */
function* paragraphContentSteps(
  el: XmlElement,
  ctx: ReaderContext,
  children: ParagraphItem[],
  budget: { read: number }
): Steps<void> {
  const opaque: PendingOpaqueChild[] = [];
  for (let index = 0; index < el.children.length; index++) {
    const child = el.children[index]!;
    if (child.type !== "element") {
      continue;
    }
    if (++budget.read % PARAGRAPH_STEP === 0) {
      yield;
    }
    const resolved = resolveContentChild(child);
    const name = resolved === undefined ? undefined : stripWPrefix(resolved.name);
    if (name === "r" && isLongRun(resolved!) && !isFieldMarkerRun(resolved!, ctx.field.marked)) {
      yield* fieldRunSteps(resolved!, ctx.field, children);
      continue;
    }
    if (name === "fldSimple") {
      yield* simpleFieldSteps(resolved!, children);
      continue;
    }
    if (resolved === undefined || name === undefined || !STEPPED_WRAPPERS.has(name)) {
      readParagraphContent(el, ctx, children, index, index + 1, opaque);
      for (const pending of opaque) {
        pending.child.rawXml = yield* serializeElementSteps(pending.element);
      }
      opaque.length = 0;
      continue;
    }
    if (name === "hyperlink") {
      const outerField = ctx.field;
      ctx.field = createFieldState();
      const linkItems: ParagraphItem[] = [];
      try {
        yield* paragraphContentSteps(resolved, ctx, linkItems, budget);
      } finally {
        ctx.field = outerField;
      }
      children.push(finishHyperlink(resolved, ctx, linkItems));
    } else if (isRevisionKind(name)) {
      const rev = parseRevisionInfo(resolved);
      if (rev) {
        const read = name === "del" ? readDeletedRunContent : parseRunContent;
        for (let from = 0; from < resolved.children.length; from += PARAGRAPH_STEP) {
          const to = Math.min(from + PARAGRAPH_STEP, resolved.children.length);
          for (let at = from; at < to; at++) {
            const run = resolved.children[at]!;
            if (run.type === "element" && isLongRevisionRun(run, ctx.field.marked)) {
              // A run longer than a step of its own, read in steps — as `readRevisionContent` reads it.
              for (const part of yield* runPartsSteps(run, read)) {
                children.push("type" in part ? part : revisionRun(name, rev, part));
              }
            } else {
              readRevisionContent(resolved, name, rev, children, ctx.field.marked, at, at + 1);
            }
          }
          budget.read += to - from;
          yield;
        }
      }
    } else {
      const wrapper = WRAPPERS[name as keyof typeof WRAPPERS];
      const start = wrapper.start(resolved, name);
      if (start) {
        children.push(start);
      }
      yield* paragraphContentSteps(
        wrappedContent(resolved, wrapper.properties),
        ctx,
        children,
        budget
      );
      if (wrapper.end) {
        children.push(wrapper.end);
      }
    }
  }
}

/**
 * The start of a paragraph's read: `split`, its children in source order with the fields not read whole split into
 * their field characters (`splitFieldPieces`, which the caller runs whole or in steps). `marked` is the field state's
 * flag as it was, which the caller restores once the content is read.
 */
function beginParagraph(
  pEl: XmlElement,
  ctx: ReaderContext,
  split: XmlElement
): { split: XmlElement; items: ParagraphItem[]; marked: boolean } {
  const field = ctx.field;
  const marked = field.marked;
  field.marked = split !== pEl;
  return { split, items: [], marked };
}

/** The paragraph `items` — its content, read — make, with range permissions moved onto the child that follows them. */
function finishParagraph(pEl: XmlElement, ctx: ReaderContext, items: ParagraphItem[]): Paragraph {
  const pPrEl = findChildNs(pEl, "pPr");
  // A marker inside a field the paragraph does not finish stays where it was.
  // (Nearly every paragraph has no marker inside a field: its items are its
  // children as they are.)
  // One pass says whether either of the steps below has anything to do.
  let hasSlot = false;
  let hasRiding = false;
  for (let i = 0; i < items.length; i++) {
    const item = items[i]!;
    if (isFieldMarkerSlot(item)) {
      hasSlot = true;
    } else if (isRidingMarker(item)) {
      hasRiding = true;
    }
  }
  let children = items as (ParagraphChild | RangeMarker)[];
  if (hasSlot) {
    children = [];
    const settled = new Set<PendingFieldMarker>();
    for (const item of items) {
      if (!isFieldMarkerSlot(item)) {
        children.push(item);
      } else if (!item.fieldMarker.embedded) {
        children.push(item.fieldMarker.marker);
        settled.add(item.fieldMarker);
      }
    }
    if (settled.size > 0) {
      ctx.field.markers = ctx.field.markers.filter(entry => !settled.has(entry));
    }
  }

  const paraId = pEl.attributes["w14:paraId"];
  const textId = pEl.attributes["w14:textId"];

  // Without a riding marker (a field's slots resolve to markers that are
  // paragraph children or ride, hence the second test) the items are the children.
  let paragraphChildren = children as ParagraphChild[];
  let rangeMarkersAfter: RangeMarker[] | undefined;
  if (hasRiding || hasSlot) {
    const arranged = arrangeParagraphChildren(children);
    paragraphChildren = arranged.children;
    rangeMarkersAfter = arranged.rangeMarkersAfter;
  }
  const result: Mutable<Paragraph> = {
    type: "paragraph",
    properties: pPrEl ? parseParagraphProperties(pPrEl) : undefined,
    children: paragraphChildren
  };
  if (rangeMarkersAfter) {
    result.rangeMarkersAfter = rangeMarkersAfter;
  }
  if (paraId) {
    result.paraId = paraId;
  }
  if (textId) {
    result.textId = textId;
  }
  return result;
}

/**
 * `name` with the `w:` prefix — itself, as Word writes nearly every name — so
 * the run loop compares names without taking each apart.
 */
function wName(name: string): string {
  return name.charCodeAt(0) === 0x77 && name.charCodeAt(1) === 0x3a ? name : `w:${name}`;
}

/** The run a field piece from `splitFieldPieces` is read as. */
function fieldPieceRun(piece: FieldPiece): Run {
  const fc: Mutable<FieldContent> = {
    type: "field",
    fieldChar: piece.fieldChar,
    instruction: piece.instruction
  };
  const ffDataEl = findChildNs(piece.fldChar, "ffData");
  if (ffDataEl) {
    fc.formField = parseFfData(ffDataEl);
  }
  return {
    properties: piece.rPr ? parseRunProperties(piece.rPr) : undefined,
    content: [fc]
  };
}

/** The run a field's instruction text kept in place by `splitFieldPieces` is read as. */
function fieldCodeRun(el: XmlElement, text: string): Run {
  const rPrEl = findChildNs(el, "rPr");
  return {
    properties: rPrEl ? parseRunProperties(rPrEl) : undefined,
    content: [{ type: "text", text, fieldCode: true }]
  };
}

/**
 * Add a range marker met in paragraph content to `children`: in place, or —
 * inside an open complex field — as a slot the field may take it from.
 */
function pushParagraphMarker(
  children: ParagraphItem[],
  field: ReaderContext["field"],
  marker: RangeMarker | undefined
): void {
  if (!marker) {
    return;
  }
  if (field.state === "none") {
    children.push(marker);
    return;
  }
  const entry: PendingFieldMarker = {
    marker,
    part: field.state === "instrText" ? "instruction" : "result",
    offset: field.state === "instrText" ? field.instr.length : field.cached.length,
    embedded: false
  };
  field.markers.push(entry);
  children.push({ fieldMarker: entry });
}

/** A wrapper element without its properties element `prName` — the content a wrapper is read for. */
function wrappedContent(wrapper: XmlElement, prName: string): XmlElement {
  return {
    ...wrapper,
    children: wrapper.children.filter(c => c.type !== "element" || stripWPrefix(c.name) !== prName)
  };
}

/** The tracked-change wrappers {@link readRevisionContent} reads. */
const REVISION_KINDS = ["ins", "del", "moveFrom", "moveTo"] as const;
type RevisionKind = (typeof REVISION_KINDS)[number];

function isRevisionKind(name: string): name is RevisionKind {
  return (REVISION_KINDS as readonly string[]).includes(name);
}

/**
 * A child `readParagraphContent` keeps as markup, its XML still to be written: a caller reading in steps serialises it
 * in steps (`serializeElementSteps`), however large the markup is.
 */
interface PendingOpaqueChild {
  readonly child: Mutable<OpaqueParagraphChild>;
  readonly element: XmlElement;
}

/**
 * The paragraph-content wrappers read in place: the properties element their content is read without, and the markers
 * that open and close it among the paragraph's children.
 */
const WRAPPERS: Readonly<
  Record<
    "smartTag" | "dir" | "bdo" | "customXml",
    {
      readonly properties: string;
      readonly start: (el: XmlElement, name: string) => ParagraphItem | undefined;
      readonly end?: ParagraphItem;
    }
  >
> = {
  smartTag: { properties: "smartTagPr", start: smartTagStart, end: { type: "smartTagEnd" } },
  dir: { properties: "", start: bidiStart, end: { type: "bidiEnd" } },
  bdo: { properties: "", start: bidiStart, end: { type: "bidiEnd" } },
  customXml: { properties: "customXmlPr", start: () => undefined }
};

/** The wrappers {@link paragraphContentSteps} reads into, a step at a time, rather than as one child. */
const STEPPED_WRAPPERS: ReadonlySet<string> = new Set([
  "hyperlink",
  ...REVISION_KINDS,
  ...Object.keys(WRAPPERS)
]);

/** The marker a `w:smartTag` opens with: its element, URI and attributes. */
function smartTagStart(el: XmlElement): SmartTagStart {
  const prEl = findChildNs(el, "smartTagPr");
  const tag: Mutable<SmartTagStart> = {
    type: "smartTagStart",
    element: attrVal(el, "element") ?? ""
  };
  const uri = attrVal(el, "uri");
  if (uri !== undefined) {
    tag.uri = uri;
  }
  const attributes: SmartTagAttribute[] = [];
  for (const attrEl of prEl ? findChildrenNs(prEl, "attr") : []) {
    const attr: Mutable<SmartTagAttribute> = {
      name: attrVal(attrEl, "name") ?? "",
      val: attrVal(attrEl, "val") ?? ""
    };
    const attrUri = attrVal(attrEl, "uri");
    if (attrUri !== undefined) {
      attr.uri = attrUri;
    }
    attributes.push(attr);
  }
  if (attributes.length > 0) {
    tag.attributes = attributes;
  }
  return tag;
}

/** The marker a `w:dir` (embedding) or `w:bdo` (override) opens with. */
function bidiStart(el: XmlElement, name: string): ParagraphItem {
  return {
    type: "bidiStart",
    kind: name === "bdo" ? "override" : "embedding",
    direction: attrVal(el, "val") === "rtl" ? "rtl" : "ltr"
  };
}

/**
 * The content of `el` — a paragraph, or a wrapper inside one — read into
 * `children` in source order: its children from index `from` to `to`, all of them by default. A child kept as markup
 * is serialised here, or — given `opaque` — left to the caller, listed there.
 */
function readParagraphContent(
  el: XmlElement,
  ctx: ReaderContext,
  children: ParagraphItem[],
  from = 0,
  to = el.children.length,
  opaque?: PendingOpaqueChild[]
): void {
  // Field state machine lives on ctx so that complex fields (TOC, INDEX,
  // long REF/SEQ chains) can span paragraph boundaries — the matching
  // `<w:fldChar fldCharType="end">` may occur in a later paragraph than the
  // `begin`. Storing state on ctx is also safe because part-scoped parsers
  // (header/footer/footnote/endnote/comment) save and reset it on entry.
  const field = ctx.field;

  const end = Math.min(to, el.children.length);
  for (let index = from; index < end; index++) {
    const child = el.children[index]!;
    if (child.type !== "element") {
      continue;
    }
    const resolved = resolveContentChild(child);
    if (resolved === undefined) {
      continue;
    }

    const name = stripWPrefix(resolved.name);
    switch (name) {
      case "r": {
        const piece = field.marked ? fieldPieceOf(resolved) : undefined;
        if (piece) {
          children.push(fieldPieceRun(piece));
          break;
        }
        const code = field.marked ? fieldCodeTextOf(resolved) : undefined;
        if (code !== undefined) {
          children.push(fieldCodeRun(resolved, code));
          break;
        }
        const run = startFieldRun();
        readFieldRunChars(resolved, field, children, run, 0, resolved.children.length);
        if (field.state === "cached") {
          readFieldRunCached(resolved, field, 0, resolved.children.length);
        }
        if (fieldRunStandsAlone(field, run)) {
          // A run holding comment reference marks becomes the runs around
          // them and a paragraph-level reference per mark.
          if (run.hasMark) {
            appendAll(children, parseRunParts(resolved));
          } else {
            children.push(parseRun(resolved));
          }
        }
        break;
      }
      case "fldSimple": {
        const read = startSimpleField(resolved);
        readSimpleField(read, resolved, 0, resolved.children.length);
        finishSimpleField(read, children);
        break;
      }
      case "hyperlink": {
        // A field inside the link is read with a field state of its own — see `finishHyperlink`.
        const outerField = ctx.field;
        ctx.field = createFieldState();
        const linkItems: ParagraphItem[] = [];
        try {
          readParagraphContent(resolved, ctx, linkItems);
        } finally {
          ctx.field = outerField;
        }
        children.push(finishHyperlink(resolved, ctx, linkItems));
        break;
      }
      case "bookmarkStart":
      case "bookmarkEnd":
      case "commentRangeStart":
      case "commentRangeEnd":
        pushParagraphMarker(children, field, parseRangeMarker(resolved, name));
        break;
      case "commentReference": {
        // A mark standing outside any run (malformed, but kept as the anchor).
        const id = attrInt(resolved, "id");
        if (id !== undefined) {
          children.push({ type: "commentReference", id });
        }
        break;
      }
      case "ins":
      case "del":
      case "moveFrom":
      case "moveTo": {
        // A tracked insertion, deletion or move: each run inside is a child of its own, carrying the revision.
        const rev = parseRevisionInfo(resolved);
        if (rev) {
          readRevisionContent(resolved, name, rev, children, field.marked);
        }
        break;
      }
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
      case "customXmlInsRangeStart":
      case "customXmlInsRangeEnd":
      case "customXmlDelRangeStart":
      case "customXmlDelRangeEnd":
      case "customXmlMoveFromRangeStart":
      case "customXmlMoveFromRangeEnd":
      case "customXmlMoveToRangeStart":
      case "customXmlMoveToRangeEnd":
        pushParagraphMarker(children, field, parseRangeMarker(resolved, name));
        break;
      case "smartTag":
      case "dir":
      case "bdo":
      case "customXml": {
        // A wrapper read in place, between the markers it makes (a custom XML wrapper makes none — it is normally
        // already removed by `unwrapCustomXml`).
        const wrapper = WRAPPERS[name];
        const start = wrapper.start(resolved, name);
        if (start) {
          children.push(start);
        }
        readParagraphContent(wrappedContent(resolved, wrapper.properties), ctx, children);
        if (wrapper.end) {
          children.push(wrapper.end);
        }
        break;
      }
      case "permStart":
      case "permEnd":
        pushParagraphMarker(children, field, parseRangeMarker(resolved, name));
        break;
      case "pPr":
        // Read into `properties` below; the writer renders it from there.
        break;
      case "proofErr":
      case "lastRenderedPageBreak":
        // Non-semantic markers; safely ignored
        break;
      default: {
        // Unknown paragraph child: preserve as opaque for round-trip fidelity
        const kept: Mutable<OpaqueParagraphChild> = { type: "opaqueParagraphChild", rawXml: "" };
        if (opaque) {
          opaque.push({ child: kept, element: resolved });
        } else {
          kept.rawXml = serializeElement(resolved);
        }
        children.push(kept);
        break;
      }
    }
  }
}

/** What the field state machine found in a run as {@link readParagraphContent} reads it. */
interface FieldRun {
  /** It holds a field character, or an instruction while one is read: it is the field's, not content of its own. */
  hasFldChar: boolean;
  hasMark: boolean;
}

function startFieldRun(): FieldRun {
  return { hasFldChar: false, hasMark: false };
}

/**
 * The field characters and instruction text of the run `resolved`, from child `from` to `to`, taken by the field
 * state machine `field` — a field it ends read into `children` — and what was met recorded in `run`.
 */
function readFieldRunChars(
  resolved: XmlElement,
  field: ReaderContext["field"],
  children: ParagraphItem[],
  run: FieldRun,
  from: number,
  to: number
): void {
  const end = Math.min(to, resolved.children.length);
  for (let index = from; index < end; index++) {
    const rc = resolved.children[index]!;
    if (rc.type !== "element") {
      continue;
    }
    const rcName = wName(rc.name);
    if (rcName === "w:commentReference") {
      run.hasMark = true;
    } else if (rcName === "w:fldChar") {
      run.hasFldChar = true;
      const fldCharType = attrVal(rc, "fldCharType");
      if (fldCharType === "begin") {
        field.state = "instrText";
        field.instr = "";
        field.cached = "";
        field.markers = [];
        // Capture run properties from this run for the field
        const rPrEl = findChildNs(resolved, "rPr");
        field.runProps = rPrEl ? parseRunProperties(rPrEl) : undefined;
        // Parse ffData for legacy form fields
        const ffDataEl = findChildNs(rc, "ffData");
        field.formField = ffDataEl ? parseFfData(ffDataEl) : undefined;
      } else if (fldCharType === "separate") {
        field.state = "cached";
      } else if (fldCharType === "end") {
        // Emit the assembled field as a Run with FieldContent
        const instruction = field.instr.trim();
        const lead = field.instr.length - field.instr.trimStart().length;
        const fc: Mutable<FieldContent> = {
          type: "field",
          instruction,
          cachedValue: field.cached || undefined,
          formField: field.formField
        };
        if (field.markers.length > 0) {
          fc.rangeMarkers = field.markers.map(entry => {
            entry.embedded = true;
            return {
              marker: entry.marker,
              part: entry.part,
              offset:
                entry.part === "instruction"
                  ? Math.min(Math.max(entry.offset - lead, 0), instruction.length)
                  : entry.offset
            };
          });
          field.markers = [];
        }
        children.push({
          properties: field.runProps,
          content: [fc]
        } satisfies Run);
        field.state = "none";
        field.instr = "";
        field.cached = "";
        field.runProps = undefined;
        field.formField = undefined;
      }
    } else if (rcName === "w:instrText" && field.state === "instrText") {
      run.hasFldChar = true;
      field.instr += textContent(rc);
    }
  }
}

/** The text of the run `resolved`, from child `from` to `to`, taken as the result of the field `field` is reading. */
function readFieldRunCached(
  resolved: XmlElement,
  field: ReaderContext["field"],
  from: number,
  to: number
): void {
  const end = Math.min(to, resolved.children.length);
  for (let index = from; index < end; index++) {
    const rc = resolved.children[index]!;
    if (rc.type !== "element") {
      continue;
    }
    const rcName = wName(rc.name);
    if (rcName === "w:t") {
      field.cached += textContent(rc);
    } else if (rcName === "w:br" && isResultText(rc)) {
      // A line break in the result, as the writer writes `\n`.
      field.cached += "\n";
    }
  }
}

/**
 * Whether a run the field state machine has read ({@link readFieldRunChars}, then {@link readFieldRunCached} while a
 * result is read) is content of its own — no field holds it.
 */
function fieldRunStandsAlone(field: ReaderContext["field"], run: FieldRun): boolean {
  return field.state === "none" && !run.hasFldChar;
}

/**
 * The run `resolved` as {@link readParagraphContent} reads it — not a split field's marker run — in steps of
 * {@link RUN_STEP} children: a run is one element however much it holds.
 */
function* fieldRunSteps(
  resolved: XmlElement,
  field: ReaderContext["field"],
  children: ParagraphItem[]
): Steps<void> {
  const run = startFieldRun();
  const length = resolved.children.length;
  for (let from = 0; from < length; from += RUN_STEP) {
    readFieldRunChars(resolved, field, children, run, from, from + RUN_STEP);
    yield;
  }
  if (field.state === "cached") {
    for (let from = 0; from < length; from += RUN_STEP) {
      readFieldRunCached(resolved, field, from, from + RUN_STEP);
      yield;
    }
  }
  if (fieldRunStandsAlone(field, run)) {
    appendAll(children, yield* runPartsSteps(resolved, parseRunContent));
  }
}

/**
 * The link `resolved` makes of `linkItems`, what was read inside it. Plain runs are the link's children; anything else
 * inside it is a nested child positioned among them, and range markers ride on the item after them (or
 * `rangeMarkersAfter` at the end), so all of it stays inside the link where it was.
 *
 * The content is read with a field state of its own, restored in a `finally` once it is read: a complex field inside
 * the link (a TOC entry's PAGEREF) neither disturbs nor is swallowed by a field around it, and a read abandoned inside
 * the link leaves nothing of the link's state behind.
 */
function finishHyperlink(
  resolved: XmlElement,
  ctx: ReaderContext,
  linkItems: readonly ParagraphItem[]
): Hyperlink {
  const rId = resolved.attributes["r:id"];
  const anchor = resolved.attributes["w:anchor"] ?? resolved.attributes["anchor"];
  const tooltip = resolved.attributes["w:tooltip"] ?? resolved.attributes["tooltip"];
  const historyAttr = resolved.attributes["w:history"] ?? resolved.attributes["history"];
  const tgtFrame = resolved.attributes["w:tgtFrame"] ?? resolved.attributes["tgtFrame"];
  const docLocation = resolved.attributes["w:docLocation"] ?? resolved.attributes["docLocation"];
  const hRuns: Run[] = [];
  const nested: HyperlinkNestedChild[] = [];
  let hMarkers: RangeMarker[] = [];
  for (const raw of linkItems) {
    if (isFieldMarkerSlot(raw) && raw.fieldMarker.embedded) {
      continue;
    }
    const item = isFieldMarkerSlot(raw) ? raw.fieldMarker.marker : raw;
    if (isRangeMarker(item)) {
      hMarkers.push(item);
    } else {
      const rides = hMarkers;
      if (rides.length > 0) {
        hMarkers = [];
      }
      if (isRun(item)) {
        hRuns.push(withRangeMarkersBefore(item, rides));
      } else {
        nested.push({
          before: hRuns.length,
          child: withRangeMarkersBefore(item as Exclude<ParagraphChild, Run>, rides)
        });
      }
    }
  }
  // Resolve URL from relMap. If the security policy disallows
  // external targets, skip URL resolution entirely so the resulting
  // Hyperlink only carries an anchor (or becomes a plain non-link
  // wrapper). Internal anchor-only hyperlinks are unaffected.
  let url: string | undefined;
  if (rId && ctx.securityPolicy.allowExternalTargets) {
    const rel = ctx.relMap.get(rId);
    if (rel && rel.targetMode === "External") {
      url = rel.target;
    }
  }
  const hyperlink: Mutable<Hyperlink> & { children: Run[] } = {
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
  if (hMarkers.length > 0) {
    hyperlink.rangeMarkersAfter = hMarkers;
  }
  if (nested.length > 0) {
    hyperlink.nestedChildren = nested;
  }
  return hyperlink;
}

/**
 * A `w:fldSimple` being read — `<w:fldSimple w:instr=" PAGE "><w:r>…</w:r></w:fldSimple>` — as the one run it
 * becomes: its instruction, the text of its result so far, and the range markers inside it. The field becomes one
 * run, so markers before its result go before it and markers after go after it: every range keeps both ends and still
 * covers the field.
 */
interface SimpleFieldRead {
  readonly instruction: string;
  cached: string;
  inResult: boolean;
  readonly markersBefore: RangeMarker[];
  readonly markersAfter: RangeMarker[];
}

function startSimpleField(el: XmlElement): SimpleFieldRead {
  return {
    instruction: attrVal(el, "instr") ?? "",
    cached: "",
    inResult: false,
    markersBefore: [],
    markersAfter: []
  };
}

/** The children of `el` from index `from` to `to`, read into `read`. */
function readSimpleField(read: SimpleFieldRead, el: XmlElement, from: number, to: number): void {
  const end = Math.min(to, el.children.length);
  for (let index = from; index < end; index++) {
    const fc = el.children[index]!;
    if (fc.type !== "element") {
      continue;
    }
    if (stripWPrefix(fc.name) === "r") {
      read.inResult = true;
      for (const rc of fc.children) {
        if (rc.type === "element" && stripWPrefix(rc.name) === "t") {
          read.cached += textContent(rc);
        }
      }
    } else {
      takeRangeMarker(fc, read.inResult ? read.markersAfter : read.markersBefore);
    }
  }
}

/** The field `read` holds, with the markers around it, appended to `children`. */
function finishSimpleField(read: SimpleFieldRead, children: ParagraphItem[]): void {
  const fc: FieldContent = {
    type: "field",
    instruction: read.instruction.trim(),
    cachedValue: read.cached || undefined
  };
  appendAll(children, read.markersBefore);
  children.push({ properties: undefined, content: [fc] } satisfies Run);
  appendAll(children, read.markersAfter);
}

/** A `w:fldSimple` read into `children` in steps of {@link PARAGRAPH_STEP} of its children. */
function* simpleFieldSteps(el: XmlElement, children: ParagraphItem[]): Steps<void> {
  const read = startSimpleField(el);
  for (let from = 0; from < el.children.length; from += PARAGRAPH_STEP) {
    readSimpleField(read, el, from, from + PARAGRAPH_STEP);
    yield;
  }
  finishSimpleField(read, children);
}

/**
 * The runs of a tracked change — `w:ins`, `w:del`, `w:moveFrom` or `w:moveTo` — read into `children`, its children from
 * index `from` to `to` (all by default): each run becomes a child carrying the revision, and a range marker inside the
 * revision keeps its place among them. A deleted run's text is in `w:delText`.
 */
function readRevisionContent(
  el: XmlElement,
  kind: RevisionKind,
  revision: RevisionInfo,
  children: ParagraphItem[],
  marked: boolean,
  from = 0,
  to = el.children.length
): void {
  const end = Math.min(to, el.children.length);
  for (let index = from; index < end; index++) {
    const child = el.children[index]!;
    if (child.type !== "element") {
      continue;
    }
    if (stripWPrefix(child.name) !== "r") {
      takeRangeMarker(child, children);
      continue;
    }
    const parts =
      kind === "del"
        ? revisionRunParts(child, marked, parseDeletedRun)
        : revisionRunParts(child, marked);
    for (const part of parts) {
      children.push("type" in part ? part : revisionRun(kind, revision, part));
    }
  }
}

/**
 * Whether `el`, a child of a tracked change, is a run longer than a step of its own ({@link isLongRun}) that
 * {@link readRevisionContent} reads by its parts (`revisionRunParts`) — not a split field's marker run.
 */
function isLongRevisionRun(el: XmlElement, marked: boolean): boolean {
  return stripWPrefix(el.name) === "r" && isLongRun(el) && !isFieldMarkerRun(el, marked);
}

/** `run` as the child a tracked change of `kind` makes of it. */
function revisionRun(kind: RevisionKind, revision: RevisionInfo, run: Run): ParagraphChild {
  switch (kind) {
    case "ins":
      return { type: "insertedRun", revision, run } satisfies InsertedRun;
    case "del":
      return { type: "deletedRun", revision, run } satisfies DeletedRun;
    case "moveFrom":
      return { type: "movedFromRun", revision, run };
    case "moveTo":
      return { type: "movedToRun", revision, run };
  }
}

/** Parse a deleted run (w:delText instead of w:t). */
function parseDeletedRun(el: XmlElement): Run {
  const rPrEl = findChildNs(el, "rPr");
  return {
    properties: rPrEl ? parseRunProperties(rPrEl) : undefined,
    content: readDeletedRunContent(el)
  };
}

/** A deleted run's content — its text in `w:delText` — from child `from` to `to`, read into `content`. */
function readDeletedRunContent(
  el: XmlElement,
  content: RunContent[] = [],
  from = 0,
  to = el.children.length
): RunContent[] {
  const end = Math.min(to, el.children.length);
  for (let index = from; index < end; index++) {
    const child = el.children[index]!;
    if (child.type !== "element") {
      continue;
    }
    const name = stripWPrefix(child.name);
    if (name === "delText") {
      content.push({ type: "text", text: textContent(child) });
    } else if (name === "t") {
      content.push({ type: "text", text: textContent(child) });
    } else if (name === "br") {
      content.push({
        type: "break",
        breakType: attrVal(child, "type") as "page" | "column" | "textWrapping" | undefined
      });
    } else if (name === "tab") {
      content.push({ type: "tab" });
    }
  }
  return content;
}

// =============================================================================
// Table Parser
// =============================================================================

/** A table cell as steps — see {@link tableSteps}. */
function* tableCellSteps(el: XmlElement, ctx: ReaderContext): Steps<TableCell> {
  const content: (Paragraph | Table | RangeMarker)[] = [];
  const controls: BlockContentControl[] = [];
  for (const child of expandBlockSimpleFields(el.children)) {
    if (child.type !== "element" || takeRangeMarker(child, content)) {
      continue;
    }
    const name = stripWPrefix(child.name);
    if (name === "p") {
      content.push(yield* paragraphSteps(child, ctx));
    } else if (name === "tbl") {
      content.push(yield* tableSteps(child, ctx));
    } else if (name === "sdt") {
      appendAll(content, yield* flattenSdtSteps(child, ctx, cellControlOptions(content, controls)));
    }
  }
  return finishTableCell(el, content, controls);
}

/** {@link tableCellSteps} for a read that does not yield: the same cell, without the steps. */
function parseTableCell(el: XmlElement, ctx: ReaderContext): TableCell {
  const content: (Paragraph | Table | RangeMarker)[] = [];
  const controls: BlockContentControl[] = [];
  for (const child of expandBlockSimpleFields(el.children)) {
    if (child.type !== "element" || takeRangeMarker(child, content)) {
      continue;
    }
    const name = stripWPrefix(child.name);
    if (name === "p") {
      content.push(parseParagraph(child, ctx));
    } else if (name === "tbl") {
      content.push(parseTable(child, ctx));
    } else if (name === "sdt") {
      appendAll(content, flattenSdt(child, ctx, cellControlOptions(content, controls)));
    }
  }
  return finishTableCell(el, content, controls);
}

/**
 * How a cell reads a block content control: the `TableCell.content` union does not include an SDT, so its content is
 * flattened into the cell, and the control is recorded around the blocks it held (`TableCell.contentControls`).
 */
function cellControlOptions(
  content: readonly (Paragraph | Table | RangeMarker)[],
  controls: BlockContentControl[]
): Parameters<typeof flattenSdt>[2] {
  const base = content.filter(item => !isRangeMarker(item)).length;
  return { tables: true, controls: { list: controls, base } };
}

/** The cell `el` makes of the content read from it. */
function finishTableCell(
  el: XmlElement,
  content: (Paragraph | Table | RangeMarker)[],
  controls: BlockContentControl[]
): TableCell {
  const tcPrEl = findChildNs(el, "tcPr");
  let props: TableCellProperties | undefined;
  if (tcPrEl) {
    const p: Mutable<TableCellProperties> = {};
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
      p.verticalMerge = (attrVal(vmEl, "val") ??
        "continue") as TableCellProperties["verticalMerge"];
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
      p.verticalAlign = readToken(
        attrVal(vAlignEl, "val"),
        WRITTEN.cellVerticalAlign,
        p,
        "verticalAlign"
      ) as TableCellProperties["verticalAlign"];
    }

    if (findChildNs(tcPrEl, "noWrap")) {
      p.noWrap = true;
    }

    const textDirEl = findChildNs(tcPrEl, "textDirection");
    if (textDirEl) {
      p.textDirection = readToken(
        attrVal(textDirEl, "val"),
        WRITTEN.textDirection,
        p,
        "textDirection"
      ) as TableCellProperties["textDirection"];
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

  const cell: Mutable<TableCell> = {
    properties: props,
    content: foldRangeMarkers(content)
  };
  if (controls.length > 0) {
    cell.contentControls = controls;
  }
  return cell;
}

/** A table row as steps — see {@link tableSteps}. */
function* tableRowSteps(el: XmlElement, ctx: ReaderContext): Steps<TableRow> {
  const cells: TableCell[] = [];
  const markers: RangeMarker[] = [];
  for (const child of el.children) {
    if (child.type !== "element" || takeRangeMarker(child, markers)) {
      continue;
    }
    if (stripWPrefix(child.name) === "tc") {
      cells.push(withMarkersTaken(yield* tableCellSteps(child, ctx), markers));
    }
  }
  return finishTableRow(el, cells, markers);
}

/** {@link tableRowSteps} for a read that does not yield. */
function parseTableRow(el: XmlElement, ctx: ReaderContext): TableRow {
  const cells: TableCell[] = [];
  const markers: RangeMarker[] = [];
  for (const child of el.children) {
    if (child.type !== "element" || takeRangeMarker(child, markers)) {
      continue;
    }
    if (stripWPrefix(child.name) === "tc") {
      cells.push(withMarkersTaken(parseTableCell(child, ctx), markers));
    }
  }
  return finishTableRow(el, cells, markers);
}

/** `item` given the range markers met before it as `rangeMarkersBefore`, which are taken from `markers`. */
function withMarkersTaken<T extends TableRow | TableCell>(item: T, markers: RangeMarker[]): T {
  if (markers.length === 0) {
    return item;
  }
  return { ...item, rangeMarkersBefore: markers.splice(0) };
}

/** The row `el` makes of its cells, and the range markers after the last of them. */
function finishTableRow(el: XmlElement, cells: TableCell[], markers: RangeMarker[]): TableRow {
  const trPrEl = findChildNs(el, "trPr");
  const tblPrExEl = findChildNs(el, "tblPrEx");
  let props: TableRowProperties | undefined;
  if (trPrEl || tblPrExEl) {
    const p: Mutable<TableRowProperties> = {};
    if (tblPrExEl) {
      p.tblPrEx = parseTableProperties(tblPrExEl);
    }
    if (trPrEl) {
      const heightEl = findChildNs(trPrEl, "trHeight");
      if (heightEl) {
        p.height = parseRowHeight(heightEl);
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

  const row: Mutable<TableRow> = { properties: props, cells };
  if (markers.length > 0) {
    row.rangeMarkersAfter = markers;
  }
  return row;
}

/** Inner parse for row properties content (used by propertyChange recursion). */
function parseRowPrInner(trPrEl: XmlElement): TableRowProperties {
  const p: Mutable<TableRowProperties> = {};
  const heightEl = findChildNs(trPrEl, "trHeight");
  if (heightEl) {
    p.height = parseRowHeight(heightEl);
  }
  if (findChildNs(trPrEl, "tblHeader")) {
    p.tableHeader = true;
  }
  if (findChildNs(trPrEl, "cantSplit")) {
    p.cantSplit = true;
  }
  return p;
}

/**
 * A table as steps, one per block it holds — each row, and each paragraph in a cell — at any depth of nesting.
 *
 * A table is one body element however many rows it has, and the streaming reader hands out a body element whole: a
 * 20,000-row table was read in one call, blocking the event loop for a tenth of a second. Read as steps, the streaming
 * reader can give the event loop a turn between its blocks, with the code every other reader runs through
 * `runSteps`.
 */
function* tableSteps(tblEl: XmlElement, ctx: ReaderContext): Steps<Table> {
  const rows: TableRow[] = [];
  const markers: RangeMarker[] = [];
  for (const child of tblEl.children) {
    if (child.type !== "element" || takeRangeMarker(child, markers)) {
      continue;
    }
    if (stripWPrefix(child.name) === "tr") {
      rows.push(withMarkersTaken(yield* tableRowSteps(child, ctx), markers));
      yield;
    }
  }
  return finishTable(tblEl, rows, markers);
}

/**
 * {@link tableSteps} for a read that does not yield: the same table, without the steps. Reading every table through
 * them cost a read that never yields a tenth more on a document of small cells.
 */
function parseTable(tblEl: XmlElement, ctx: ReaderContext): Table {
  const rows: TableRow[] = [];
  const markers: RangeMarker[] = [];
  for (const child of tblEl.children) {
    if (child.type !== "element" || takeRangeMarker(child, markers)) {
      continue;
    }
    if (stripWPrefix(child.name) === "tr") {
      rows.push(withMarkersTaken(parseTableRow(child, ctx), markers));
    }
  }
  return finishTable(tblEl, rows, markers);
}

/** The table `tblEl` makes of its rows, and the range markers after the last of them. */
function finishTable(tblEl: XmlElement, rows: TableRow[], markers: RangeMarker[]): Table {
  const tblPrEl = findChildNs(tblEl, "tblPr");
  const gridEl = findChildNs(tblEl, "tblGrid");
  let columnWidths: number[] | undefined;
  if (gridEl) {
    columnWidths = [];
    for (const col of findChildrenNs(gridEl, "gridCol")) {
      columnWidths.push(parseInt(col.attributes["w:w"] ?? col.attributes["w"] ?? "0", 10));
    }
  }

  const table: Mutable<Table> = {
    type: "table",
    properties: tblPrEl ? parseTableProperties(tblPrEl) : undefined,
    columnWidths,
    rows
  };
  if (markers.length > 0) {
    table.rangeMarkersAfter = markers;
  }
  return table;
}

// =============================================================================
// Footnotes/Endnotes Parser
// =============================================================================

/**
 * Each note's `w:type` by `w:id`, read off the start tags — `parseNotesXml`
 * drops the separators, which is exactly what settings references.
 */
function noteTypesById(xmlStr: string, elementName: string): Map<number, string> {
  const types = new Map<number, string>();
  for (const [, attrs] of xmlStr.matchAll(new RegExp(`<w:${elementName}\\b([^>]*)>`, "g"))) {
    const id = /\bw:id=["'](-?\d+)["']/.exec(attrs)?.[1];
    if (id !== undefined) {
      types.set(Number(id), /\bw:type=["']([^"']*)["']/.exec(attrs)?.[1] ?? "normal");
    }
  }
  return types;
}

function parseNotesXml(
  xmlStr: string,
  elementName: string,
  ctx: ReaderContext
): { id: number; type?: NoteType; content: Paragraph[] }[] {
  // Each note part is self-contained. Save and reset the field state so an
  // unterminated complex field in the document body cannot bleed into a
  // footnote/endnote and swallow its runs.
  const savedField = ctx.field;
  ctx.field = createFieldState();
  try {
    const doc = parseBoundContentPart(xmlStr, ctx.securityPolicy.maxXmlDepth);
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

      const items: (Paragraph | RangeMarker)[] = [];
      for (const child of expandBlockSimpleFields(noteEl.children)) {
        if (child.type !== "element" || takeRangeMarker(child, items)) {
          continue;
        }
        const ln = stripWPrefix(child.name);
        if (ln === "p") {
          items.push(parseParagraph(child, ctx));
        } else if (ln === "sdt") {
          // A note's content is `Paragraph[]`, so the SDT is flattened.
          appendAll(items, flattenSdt(child, ctx, { tables: false }));
        }
      }

      const note: { id: number; type?: NoteType; content: Paragraph[] } = {
        id,
        content: foldRangeMarkers(items)
      };
      if (type === "continuationNotice" || type === "normal") {
        note.type = type;
      }
      notes.push(note);
    }

    return notes;
  } finally {
    ctx.field = savedField;
  }
}

// =============================================================================
// Header/Footer Parser
// =============================================================================

function parseHeaderFooterXml(xmlStr: string, ctx: ReaderContext): HeaderFooterContent {
  return parseHeaderFooterRoot(
    parseBoundContentPart(xmlStr, ctx.securityPolicy.maxXmlDepth).root,
    ctx
  );
}

function parseHeaderFooterRoot(root: XmlElement, ctx: ReaderContext): HeaderFooterContent {
  // Header/footer parts are self-contained: reset field state on entry so an
  // unterminated complex field in the body does not consume header/footer runs.
  const savedField = ctx.field;
  ctx.field = createFieldState();
  try {
    const children: (Paragraph | Table | RangeMarker)[] = [];
    for (const child of expandBlockSimpleFields(root.children)) {
      if (child.type !== "element" || takeRangeMarker(child, children)) {
        continue;
      }
      const name = stripWPrefix(child.name);
      if (name === "p") {
        children.push(parseParagraph(child, ctx));
      } else if (name === "tbl") {
        children.push(parseTable(child, ctx));
      } else if (name === "sdt") {
        // HeaderFooterContent.children is `(Paragraph | Table)[]`, so the
        // SDT is flattened.
        appendAll(children, flattenSdt(child, ctx, { tables: true }));
      }
    }
    return { children: foldRangeMarkers(children) };
  } finally {
    ctx.field = savedField;
  }
}

/**
 * Whether a header, its watermark taken out, is the paragraph the writer
 * (and Word) puts a watermark in when the header has nothing else.
 */
function isWatermarkOnlyHeader(content: HeaderFooterContent): boolean {
  if (content.children.length !== 1) {
    return false;
  }
  const only = content.children[0];
  if (only.type !== "paragraph" || only.children.length > 0) {
    return false;
  }
  if (only.rangeMarkersBefore?.length || only.rangeMarkersAfter?.length) {
    return false;
  }
  const props = only.properties;
  return !props || Object.keys(props).every(k => k === "style" && props.style === "Header");
}

/** The section breaks inside the body, in document order. */
function innerSections(body: readonly BodyContent[]): SectionProperties[] {
  const out: SectionProperties[] = [];
  for (const item of body) {
    if (item.type === "paragraph" && item.properties?.sectionProperties) {
      out.push(item.properties.sectionProperties);
    }
  }
  return out;
}

/**
 * The document's watermark: the one every header shown anywhere draws, the
 * same in each — and where each draws it. `undefined` when a page shows no
 * header, or a header shown draws none or another.
 */
function findDocumentWatermark(
  slots: readonly HeaderSlot[],
  headerRoots: ReadonlyMap<string, { root: XmlElement }>
): { watermark: Watermark; found: Map<string, FoundWatermark> } | undefined {
  const found = new Map<string, FoundWatermark>();
  let key: string | undefined;
  for (const slot of slots) {
    if (slot.rId === undefined) {
      return undefined;
    }
    if (found.has(slot.rId)) {
      continue;
    }
    const entry = headerRoots.get(slot.rId);
    const hit = entry ? findWatermark(entry.root) : undefined;
    if (!hit) {
      return undefined;
    }
    const k = JSON.stringify(hit.watermark);
    if (key !== undefined && k !== key) {
      return undefined;
    }
    key = k;
    found.set(slot.rId, hit);
  }
  const first = found.values().next().value;
  return first ? { watermark: first.watermark, found } : undefined;
}

/**
 * The watermark-only headers to drop, references and all: those every
 * reference to which is a section showing it as its own while inheriting no
 * header of that type — where the writer names its watermark-only header
 * again.
 */
function droppableWatermarkHeaders(
  slots: readonly HeaderSlot[],
  sections: readonly SectionProperties[],
  watermarkOnly: ReadonlySet<string>
): Set<string> {
  const out = new Set<string>();
  if (watermarkOnly.size === 0) {
    return out;
  }
  const replaceable = new Set<string>();
  for (const slot of slots) {
    if (slot.own && slot.inherited === undefined && slot.rId !== undefined) {
      replaceable.add(`${slot.section}:${slot.type}:${slot.rId}`);
    }
  }
  for (const rId of watermarkOnly) {
    const refs = sections.flatMap((sect, index) =>
      (sect.headers ?? []).filter(h => h.rId === rId).map(h => `${index}:${h.type}:${rId}`)
    );
    if (refs.length > 0 && refs.every(ref => replaceable.has(ref))) {
      out.add(rId);
    }
  }
  return out;
}

/**
 * The watermark-only headers the reader drops, references and all (see
 * {@link droppableWatermarkHeaders}), decided from the sections and the
 * header parts alone — for the streaming reader, which does not keep the
 * headers but must give the sections the references `readDocx` gives them.
 */
export function watermarkHeadersToDrop(
  sections: readonly SectionProperties[],
  evenAndOddHeaders: boolean,
  headerRoots: ReadonlyMap<string, { readonly root: XmlElement }>,
  ctx: ReaderContext
): Set<string> {
  const slots = headerSlots(sections, evenAndOddHeaders);
  const documentWatermark = findDocumentWatermark(slots, headerRoots);
  const watermarkOnly = new Set<string>();
  for (const [rId, found] of documentWatermark?.found ?? []) {
    const entry = headerRoots.get(rId);
    try {
      if (
        entry &&
        isWatermarkOnlyHeader(parseHeaderFooterRoot(withoutWatermark(entry.root, found), ctx))
      ) {
        watermarkOnly.add(rId);
      }
    } catch {
      // A malformed header is skipped, as `readDocx` skips it.
    }
  }
  return droppableWatermarkHeaders(slots, sections, watermarkOnly);
}

/** `sect` without its references to the headers `rIds`. */
export function withoutHeaderReferences(
  sect: SectionProperties,
  rIds: ReadonlySet<string>
): SectionProperties {
  const next: Mutable<SectionProperties> = { ...sect };
  const headers = sect.headers?.filter(h => !rIds.has(h.rId));
  if (headers && headers.length > 0) {
    next.headers = headers;
  } else {
    delete next.headers;
  }
  return next;
}

// =============================================================================
// Comments Parser
// =============================================================================

function parseCommentsXmlFromCtx(xmlStr: string, ctx: ReaderContext): CommentDef[] {
  return parseCommentsXmlExternal(xmlStr, ctx, parseParagraph);
}

// =============================================================================
// Main Document Parser
// =============================================================================

/**
 * Floating images, drawing shapes and opaque drawings found in an element tree, and the serialized
 * `mc:AlternateContent` elements a drawing was taken from — created when the first is found, so the many paragraphs
 * without one allocate nothing.
 */
interface FloatingExtraction {
  readonly images: FloatingImage[];
  readonly shapes: DrawingShape[];
  readonly opaqueDrawings: OpaqueDrawing[];
  readonly ctx: ReaderContext;
  extracted: Set<string> | undefined;
  /** Where the text of the shapes found is left to be read in steps; read with each shape when absent. */
  readonly pendingText?: PendingShapeText[];
}

/** How many drawings `extraction` has found. */
function drawingsFound(extraction: FloatingExtraction): number {
  return extraction.images.length + extraction.shapes.length + extraction.opaqueDrawings.length;
}

/**
 * What `child` — a drawing or alternate content ({@link isFloatingCandidate}) — is to an extraction: the element to
 * walk next, which is empty for an anchored or inline drawing, taken and walked no further; an
 * `mc:AlternateContent` is walked in the branch a consumer would pick (as `parseParagraph` does), never both — the
 * branches are alternative spellings of one drawing, and when the branch yields a drawing, the run content's
 * verbatim copy of the whole element (an `opaqueRun`) is dropped by `stripExtractedDrawingRuns`.
 */
function visitFloating(extraction: FloatingExtraction, child: XmlElement): XmlElement {
  const { images, shapes, opaqueDrawings, ctx } = extraction;
  if (child.name === "wp:anchor") {
    // Check if this is a pic (image) or wsp (shape)
    const graphicEl = findChild(child, "a:graphic");
    const graphicDataEl = graphicEl ? findChild(graphicEl, "a:graphicData") : undefined;
    const wspEl = graphicDataEl
      ? (findChild(graphicDataEl, "wps:wsp") ?? findChildNs(graphicDataEl, "wsp"))
      : undefined;
    if (wspEl) {
      const shape = parseDrawingShape(child, wspEl, ctx, extraction.pendingText);
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
        // Not an image — opaque inline drawing. We deliberately keep this
        // path even though parseDrawingContent also emits an `opaqueRun`
        // for the same drawing: the body-level pass below removes the
        // duplicate opaqueRun once we know this OpaqueDrawing has been
        // captured. Inside table cells / headers / footers / SDTs (where
        // this extractor is not invoked) the opaqueRun is the only
        // representation, so the drawing still survives a round-trip.
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
  }
  if (child.name === "mc:AlternateContent") {
    return findChild(child, "mc:Choice") ?? findChild(child, "mc:Fallback") ?? NOTHING_TO_WALK;
  }
  return NOTHING_TO_WALK;
}

/** An element with nothing in it: what {@link visitFloating} answers for a drawing it took. */
const NOTHING_TO_WALK: XmlElement = { type: "element", name: "", attributes: {}, children: [] };

/** Whether {@link visitFloating} has anything to do with `child`: a drawing, or alternate content. */
function isFloatingCandidate(child: XmlElement): boolean {
  return (
    child.name === "wp:anchor" || child.name === "wp:inline" || child.name === "mc:AlternateContent"
  );
}

/** `child` when it is an `mc:AlternateContent`, whose walk {@link recordAlternate} records. */
function alternateOf(child: XmlElement): XmlElement | undefined {
  return child.name === "mc:AlternateContent" ? child : undefined;
}

/** `alternate`, recorded as extracted if the walk of its branch found a drawing since `before`. */
function recordAlternate(
  extraction: FloatingExtraction,
  alternate: XmlElement | undefined,
  before: number
): void {
  if (alternate !== undefined && drawingsFound(extraction) > before) {
    (extraction.extracted ??= new Set()).add(serializeElement(alternate));
  }
}

/** The whole of `el` walked into `extraction`, depth first, in one go. */
function extractFloatingAll(extraction: FloatingExtraction, el: XmlElement): void {
  for (const child of el.children) {
    if (child.type !== "element") {
      continue;
    }
    // Nearly every element is neither a drawing nor alternate content, and is walked: decided here, not in a call.
    const next = isFloatingCandidate(child) ? visitFloating(extraction, child) : child;
    if (next.children.length > 0) {
      const before = drawingsFound(extraction);
      extractFloatingAll(extraction, next);
      recordAlternate(extraction, alternateOf(child), before);
    }
  }
}

/**
 * {@link extractFloatingAll} as steps of `step` elements: the walk from explicit frames, which can stop and resume.
 */
function* extractFloatingSteps(
  extraction: FloatingExtraction,
  el: XmlElement,
  step: number
): Steps<void> {
  const frames: {
    readonly el: XmlElement;
    index: number;
    readonly alternate?: XmlElement;
    readonly before: number;
  }[] = [{ el, index: 0, before: 0 }];
  let visited = 0;
  while (frames.length > 0) {
    const frame = frames[frames.length - 1]!;
    if (frame.index >= frame.el.children.length) {
      frames.pop();
      recordAlternate(extraction, frame.alternate, frame.before);
      continue;
    }
    if (++visited % step === 0) {
      yield;
    }
    const child = frame.el.children[frame.index++]!;
    if (child.type !== "element") {
      continue;
    }
    const next = isFloatingCandidate(child) ? visitFloating(extraction, child) : child;
    if (next.children.length > 0) {
      frames.push({
        el: next,
        index: 0,
        alternate: alternateOf(child),
        before: drawingsFound(extraction)
      });
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

/**
 * A paragraph is considered "empty" for the purposes of synthetic-anchor
 * detection if it has no children, or if every child is a run whose content
 * is either absent or contains only zero-length text segments. Inline images,
 * fields, hyperlinks etc. all count as non-empty content. Properties (style
 * id, alignment, etc.) are intentionally ignored — a single floating drawing
 * that the writer wrapped in its own paragraph would never have meaningful
 * paragraph properties.
 */
function isEmptyParagraph(para: Paragraph): boolean {
  // Range markers (permissions riding on the paragraph) are content to keep.
  if ((para.rangeMarkersAfter?.length ?? 0) > 0 || (para.rangeMarkersBefore?.length ?? 0) > 0) {
    return false;
  }
  if (!para.children || para.children.length === 0) {
    return true;
  }
  for (const child of para.children) {
    if ((child.rangeMarkersBefore?.length ?? 0) > 0) {
      return false;
    }
    if (!isRun(child)) {
      // Anything with a `type` (hyperlink, bookmark, insertedRun, etc.) is
      // considered meaningful content.
      return false;
    }
    const run = child as Run;
    for (const c of run.content) {
      if (c.type === "text") {
        if (c.text.length > 0) {
          return false;
        }
      } else {
        // Any non-text run content (image, field, break, tab, ruby, etc.)
        // makes the paragraph non-empty.
        return false;
      }
    }
  }
  return true;
}

/**
 * Remove the `opaqueRun` entries of drawings the body-level extractor has
 * already turned into blocks of their own, which would otherwise be written
 * twice:
 *
 * - with `inlineDrawings`, every non-picture `<wp:inline>` drawing. These are
 *   emitted by parseDrawingContent so the drawing survives a round-trip when
 *   its containing paragraph lives inside a table cell, header, footer or SDT
 *   (places where the body-level extractor never runs). At the body level,
 *   however, the same drawings are also captured as `OpaqueDrawing` entries.
 * - every `mc:AlternateContent` in `alternateContent` (serialized), whose
 *   chosen branch the extractor read — a text box or shape the way Word writes
 *   one, DrawingML with a VML fallback.
 *
 * Mutates the run content in place.
 */
function stripExtractedDrawingRuns(
  para: Paragraph,
  inlineDrawings: boolean,
  alternateContent: ReadonlySet<string>
): void {
  const strip = (run: Run): void => {
    const content = run.content as RunContent[];
    let i = 0;
    while (i < content.length) {
      const c = content[i]!;
      if (
        c.type === "opaqueRun" &&
        ((inlineDrawings && c.rawXml.includes("<wp:inline") && !c.rawXml.includes("<pic:pic")) ||
          alternateContent.has(c.rawXml))
      ) {
        content.splice(i, 1);
      } else {
        i++;
      }
    }
  };
  for (const child of para.children) {
    if (isRun(child)) {
      strip(child);
    } else if (child.type === "hyperlink") {
      child.children.forEach(strip);
    } else if (
      child.type === "insertedRun" ||
      child.type === "movedFromRun" ||
      child.type === "movedToRun"
    ) {
      strip(child.run);
    }
  }
}

/**
 * Remove the `opaqueRun` copies parseDrawingContent keeps of anchored
 * drawings, once a top-level body paragraph has lifted every anchor out as a
 * floating image, shape or opaque drawing block. Mutates the runs in place.
 */
function stripAnchoredDrawingOpaqueRuns(para: Paragraph): void {
  // Every run, including those inside hyperlinks and tracked changes: the
  // extractor lifts anchors out of those too.
  walkBlocks([para], {
    enterRun(run) {
      const content = run.content as RunContent[];
      for (let i = content.length - 1; i >= 0; i--) {
        const c = content[i]!;
        if (c.type === "opaqueRun" && c.rawXml.includes("<wp:anchor")) {
          content.splice(i, 1);
        }
      }
    }
  });
}

function parseDocumentXml(
  xmlStr: string,
  ctx: ReaderContext
): {
  body: BodyContent[];
  sectionProperties?: SectionProperties;
  bodyRangeMarkersAfter?: RangeMarker[];
  background?: DocumentBackground;
} {
  const doc = parseBoundContentPart(xmlStr, ctx.securityPolicy.maxXmlDepth);
  const root = doc.root;

  // Parse background
  let background: DocumentBackground | undefined;
  const bgEl = findChildNs(root, "background");
  if (bgEl) {
    const bg: Mutable<DocumentBackground> = {};
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

  // The body is `w:body` in Word's namespace, under a root that is `w:document`: a part in another
  // namespace — or with `w:` bound to nothing — has no body Word would read.
  const bodyEl = isWordDocumentRoot(root) ? findChild(root, "w:body") : undefined;
  if (!bodyEl) {
    throw new DocxParseError("Missing w:body element in document.xml");
  }

  const body: BodyContent[] = [];
  let sectionProperties: SectionProperties | undefined;
  // Range markers between body blocks travel on the block that follows them.
  const pendingMarkers: RangeMarker[] = [];

  // Instead of extracting floating content from the entire body tree and
  // appending at the end (which loses positional information), we now extract
  // floating content per-paragraph and insert it immediately after the
  // paragraph it belongs to.

  // A body without field characters has no field to read in pieces (a
  // block-level simple field is read as the field characters it stands for).
  const field = ctx.field;
  // (One scan of a body without fields: "fld" is rare in one.)
  field.noFieldChars =
    !xmlStr.includes("fld") || (!xmlStr.includes("fldChar") && !xmlStr.includes("fldSimple"));
  try {
    readBodyElements(bodyEl, ctx, body, pendingMarkers, props => {
      sectionProperties = props;
    });
  } finally {
    field.noFieldChars = undefined;
  }

  return {
    body,
    sectionProperties,
    background,
    ...(pendingMarkers.length > 0 ? { bodyRangeMarkersAfter: pendingMarkers } : {})
  };
}

/** The children of `w:body` read into `body`, the markers after its last block left in `pendingMarkers`. */
function readBodyElements(
  bodyEl: XmlElement,
  ctx: ReaderContext,
  body: BodyContent[],
  pendingMarkers: RangeMarker[],
  setSectionProperties: (props: SectionProperties) => void
): void {
  for (const child of bodyEl.children) {
    if (child.type !== "element") {
      continue;
    }
    const name = stripWPrefix(child.name);

    // The body-level <w:sectPr> sets the final section properties rather than
    // contributing a BodyContent entry, so it is handled here directly; every
    // other element type is delegated to the shared `parseBodyChild` so the
    // streaming reader can reuse the exact same per-element logic.
    if (name === "sectPr") {
      setSectionProperties(parseSectionProperties(child));
      continue;
    }

    if (takeRangeMarker(child, pendingMarkers)) {
      continue;
    }
    for (const item of parseBodyChild(child, ctx)) {
      if (pendingMarkers.length > 0) {
        body.push(withRangeMarkersBefore(item, pendingMarkers.splice(0)));
      } else {
        body.push(item);
      }
    }
  }
}

/**
 * Parse a single body-level child element into zero or more {@link BodyContent}
 * entries.
 *
 * Extracted from {@link parseDocumentXml}'s body loop so that both the batch
 * reader and the streaming reader (`Stream.DocumentReader`) run the
 * exact same per-element logic — paragraph/table/sdt/altChunk/math/textbox
 * handling plus per-paragraph floating-content extraction. The caller is
 * responsible for handling the body-level `<w:sectPr>` (which sets section
 * properties rather than producing a BodyContent entry); this function never
 * emits anything for it.
 *
 * A single source element can expand to multiple entries: a paragraph carrying
 * anchored drawings yields the paragraph followed by each extracted floating
 * image / shape / opaque drawing, preserving document position.
 */
export function parseBodyChild(child: XmlElement, ctx: ReaderContext): BodyContent[] {
  // The blocks nearly every body is made of — paragraphs, tables, content controls — without the steps: a read that
  // does not yield need not pay for them. `bodyChildSteps` reads each the same way, through the same pieces.
  if (!ctx.paced) {
    switch (stripWPrefix(child.name)) {
      case "p": {
        const blocks = paragraphBlocks(child);
        if (blocks !== undefined) {
          return blocks;
        }
        const out: BodyContent[] = [];
        appendParagraph(child, parseParagraph(child, ctx), ctx, out);
        return out;
      }
      case "tbl":
        return [parseTable(child, ctx)];
      case "sdt": {
        const sdt = parseSdt(child, ctx);
        return sdt ? [sdt as BodyContent] : [];
      }
    }
  }
  return runSteps(bodyChildSteps(child, ctx));
}

/**
 * The blocks a body paragraph is read as when it stands for something else — a math block, a checkbox — or `undefined`
 * when it is a paragraph.
 */
function paragraphBlocks(child: XmlElement): BodyContent[] | undefined {
  // Per OOXML schema (CT_OMathPara is a member of EG_PContent), a
  // body-level math block is encoded as a paragraph containing a
  // single m:oMathPara child. Detect that shape and surface it as
  // a top-level MathBlock so the document model stays flat — the
  // writer reverses this by re-wrapping math blocks in <w:p>.
  // pPr is allowed beside the math; anything else (runs, hyperlinks, etc.) means this is not a synthetic math wrapper
  // and the paragraph is kept. Decided at the first such child: a paragraph of runs is not walked to its end.
  let mathParas: XmlElement[] | undefined;
  for (const c of child.children) {
    if (c.type !== "element" || c.name === "w:pPr") {
      continue;
    }
    if (c.name !== "m:oMathPara") {
      mathParas = undefined;
      break;
    }
    (mathParas ??= []).push(c);
  }
  if (mathParas !== undefined) {
    return mathParas.map(el => parseMathBlock(el));
  }

  // A checkbox block is written as a paragraph holding only its
  // run-level SDT (see the writer's "checkBox" case): read that shape back
  // as the block. A paragraph with properties or anything else beside
  // the control is a paragraph, and keeps the control as it was.
  const checkBox = paragraphOnlyCheckBox(child);
  return checkBox ? [checkBox] : undefined;
}

/** `para`, the body paragraph `child` was read as, appended to `out` with the floating content drawn in it. */
function appendParagraph(
  child: XmlElement,
  para: Paragraph,
  ctx: ReaderContext,
  out: BodyContent[]
): void {
  // Extract floating content from this paragraph element and insert
  // immediately after it to preserve document position.
  const extraction: FloatingExtraction = {
    images: [],
    shapes: [],
    opaqueDrawings: [],
    ctx,
    extracted: undefined
  };
  extractFloatingAll(extraction, child);
  appendExtracted(para, extraction, out);
}

/** {@link appendParagraph} as steps of {@link PARAGRAPH_STEP} elements of the paragraph's XML: it is walked whole. */
function* appendParagraphSteps(
  child: XmlElement,
  para: Paragraph,
  ctx: ReaderContext,
  out: BodyContent[]
): Steps<void> {
  const pendingText: PendingShapeText[] = [];
  const extraction: FloatingExtraction = {
    images: [],
    shapes: [],
    opaqueDrawings: [],
    ctx,
    extracted: undefined,
    pendingText
  };
  yield* extractFloatingSteps(extraction, child, PARAGRAPH_STEP * 8);
  for (const { textContent, paragraphs } of pendingText) {
    for (const paragraph of paragraphs) {
      textContent.push(yield* paragraphSteps(paragraph, ctx));
    }
  }
  appendExtracted(para, extraction, out);
}

/** `para`, appended to `out` with the floating content `extraction` found drawn in it. */
function appendExtracted(
  para: Paragraph,
  extraction: FloatingExtraction,
  out: BodyContent[]
): void {
  const {
    images: pFloatingImages,
    shapes: pDrawingShapes,
    opaqueDrawings: pOpaqueDrawings,
    extracted: extractedAlternateContent
  } = extraction;

  // parseDrawingContent (called from parseRunContent) already preserved
  // every non-picture inline drawing as an `opaqueRun` so the drawing
  // survives a round-trip even inside cells/headers/footers/SDTs where
  // this body-level extractor is not invoked. At the body level
  // the floating extraction has now also captured those drawings as
  // `OpaqueDrawing` entries — that is the form chart-parser is wired
  // to look for when promoting them to `ChartContent`. To avoid
  // duplicate output we strip any opaqueRun whose XML embeds a
  // <wp:inline> drawing from the paragraph here, and likewise the
  // verbatim copy of an mc:AlternateContent a drawing was read from.
  if (pOpaqueDrawings.length > 0 || extractedAlternateContent) {
    stripExtractedDrawingRuns(
      para,
      pOpaqueDrawings.length > 0,
      extractedAlternateContent ?? new Set()
    );
  }
  if (pFloatingImages.length > 0 || pDrawingShapes.length > 0 || pOpaqueDrawings.length > 0) {
    stripAnchoredDrawingOpaqueRuns(para);
  }

  // If the paragraph is otherwise empty AND we did extract anchored
  // content out of it, treat the paragraph as a synthetic carrier for
  // the floating drawing(s) and drop it. Otherwise keeping it would
  // cause a phantom empty paragraph to accumulate on every round-trip
  // (writer wraps floating images in their own <w:p>, reader pulls the
  // anchor out, leaving an empty <w:p> behind).
  const hasAnchoredContent =
    pFloatingImages.length > 0 || pDrawingShapes.length > 0 || pOpaqueDrawings.length > 0;
  // A section break ends its section whatever else the paragraph holds.
  const paragraphIsEmpty = isEmptyParagraph(para) && !para.properties?.sectionProperties;
  const keepParagraph = !(hasAnchoredContent && paragraphIsEmpty);
  if (keepParagraph) {
    out.push(para);
  }

  for (const fi of pFloatingImages) {
    out.push(keepParagraph ? { ...fi, inPrecedingParagraph: true } : fi);
  }
  for (const ds of pDrawingShapes) {
    out.push(ds);
  }
  for (const od of pOpaqueDrawings) {
    out.push(od);
  }
}

/**
 * {@link parseBodyChild} as steps — one per block the element holds, see {@link tableSteps} — for the streaming
 * reader, which runs it with `runStepsPaced`.
 */
export function* bodyChildSteps(child: XmlElement, ctx: ReaderContext): Steps<BodyContent[]> {
  const out: BodyContent[] = [];
  const name = stripWPrefix(child.name);

  switch (name) {
    case "p": {
      const blocks = paragraphBlocks(child);
      if (blocks !== undefined) {
        appendAll(out, blocks);
        break;
      }
      if (ctx.paced) {
        yield* appendParagraphSteps(child, yield* paragraphSteps(child, ctx), ctx, out);
      } else {
        appendParagraph(child, parseParagraph(child, ctx), ctx, out);
      }
      break;
    }
    case "tbl":
      out.push(ctx.paced ? yield* tableSteps(child, ctx) : parseTable(child, ctx));
      break;
    case "fldSimple":
      // A simple field around whole paragraphs, read as the complex field
      // it stands for (see `expandBlockSimpleFields`).
      for (const node of expandBlockSimpleFields([child])) {
        if (node.type === "element") {
          appendAll(out, yield* bodyChildSteps(node, ctx));
        }
      }
      break;
    case "sdt": {
      const sdtResult = ctx.paced ? yield* sdtSteps(child, ctx) : parseSdt(child, ctx);
      if (sdtResult) {
        out.push(sdtResult as BodyContent);
      }
      break;
    }
    case "altChunk": {
      const rId = child.attributes["r:id"] ?? child.attributes["id"];
      if (rId) {
        out.push({ type: "altChunk", rId });
      }
      break;
    }
    default: {
      // Check for math namespace
      if (child.name === "m:oMathPara") {
        out.push(parseMathBlock(child));
      } else if (child.name === "m:oMath") {
        out.push({ type: "math", content: parseMathContent(child) });
      }
      // Check for VML pict (textbox)
      if (name === "pict" || child.name === "w:pict") {
        const tb = ctx.paced ? yield* textBoxSteps(child, ctx) : parseTextBox(child, ctx);
        if (tb) {
          out.push(tb);
        }
      }
      break;
    }
  }

  return out;
}

// =============================================================================
// Public API - Read DOCX
// =============================================================================

export { checkReadOptions } from "@word/reader/package-limits";

const readDocxOptionsSchema = w.record(readOptionsShape);
const base64OptionSchema = w.record({ base64: w.boolean().optional() }).optional();
const _readDocxOptionsMatch: SchemaMatches<typeof readDocxOptionsSchema, ReadDocxOptions> = true;
void _readDocxOptionsMatch;

/** Options for reading a DOCX file. */
export interface ReadDocxOptions {
  /** Password for decrypting an encrypted DOCX file. */
  readonly password?: string;
  /**
   * Optional security policy that controls hard limits enforced while reading
   * the package (max total size, max single-part size, max part count). Falls
   * back to {@link DEFAULT_SECURITY_POLICY} when omitted. Helpful as a
   * defense-in-depth control against ZIP bombs / pathological inputs.
   */
  readonly securityPolicy?: DocxSecurityPolicy;
  /**
   * Whether each part is checked against the CRC-32 the package records for it. A
   * mismatch fails the read with a `DocxParseError` whose `cause` is the archive's
   * `ZipCrc32MismatchError`. Set `false` to read a damaged package as far as it goes.
   * @default true
   */
  readonly checkCrc32?: boolean;
}

/**
 * Read a DOCX file and parse it into a DocxDocument model. The file is given as its
 * bytes — a `Uint8Array` (a Node `Buffer` is one), any other `ArrayBufferView` (a
 * `DataView` or typed array, read as the bytes it views) or an `ArrayBuffer` — or as
 * text: the Flat OPC XML `Document.toFlatOpc` returns, or, with `{ base64: true }`, the
 * base64 `Document.toBase64` returns (as `Workbook.read` takes base64). `base64` applies to
 * a string only.
 *
 * If the file is encrypted (CFB format), provide a password via the options parameter
 * to decrypt it automatically. A Flat OPC document — the whole package as one XML file,
 * as `Document.toFlatOpc` writes it and Word saves an "XML Document" — is read as the package
 * it holds.
 */
export async function readDocx(
  input: string | Uint8Array | ArrayBuffer | ArrayBufferView,
  options?: ReadDocxOptions & {
    /** A string `input` is base64 (as `Document.toBase64` writes), not Flat OPC XML. */
    readonly base64?: boolean;
  }
): Promise<DocxDocument> {
  const where = "Document.read";
  // A number or `null` failed on `.length` with an untyped `TypeError`. An `ArrayBuffer` is recognised by its tag,
  // which one from another realm also carries.
  const bytes = typeof input === "string" ? undefined : checkBytes(input, "input", where);
  const policy = checkReadOptions(options, where);
  checkSchema(base64OptionSchema, options, where);
  return readDocxBytes(
    bytes ?? textPackageBytes(input as string, options?.base64 === true),
    options,
    policy
  );
}

const BASE64_TEXT = /^[A-Za-z0-9+/_-]*={0,2}$/;

/** The package bytes of a string given to `Document.read`: base64, or Flat OPC XML. */
function textPackageBytes(text: string, base64: boolean): Uint8Array {
  if (base64) {
    const compact = text.replace(/\s+/g, "");
    if (!BASE64_TEXT.test(compact)) {
      invalidArgument("Document.read", "input", "base64 text", text);
    }
    return base64ToBytes(compact);
  }
  if (isFlatOpc(text)) {
    return utf8Encoder.encode(text);
  }
  return invalidArgument(
    "Document.read",
    "input",
    "bytes, Flat OPC XML text, or base64 text with { base64: true }",
    text
  );
}

async function readDocxBytes(
  input: Uint8Array,
  options: ReadDocxOptions | undefined,
  policy: Required<DocxSecurityPolicy>
): Promise<DocxDocument> {
  // Defense-in-depth: reject obviously oversized packages up-front. The same
  // limit is also enforced incrementally during entry decompression so a
  // pathological deflate stream can't slip past this check.
  checkPackageSize(input, policy);
  // An encrypted DOCX (CFB format) is decrypted before the ZIP is parsed.
  const buffer = await decryptIfEncrypted(input, options?.password, policy.maxPackageSize);

  try {
    // A Flat OPC document (`Document.toFlatOpc`, Word's "XML Document") is the same
    // package as one XML file.
    if (isFlatOpc(buffer)) {
      return readDocxEntries(flatOpcEntries(buffer, policy), policy);
    }
    return await _readDocxInner(buffer, policy, options?.checkCrc32);
  } catch (e) {
    if (e instanceof DocxError) {
      throw e;
    }
    const msg = e instanceof Error ? e.message : String(e);
    throw new DocxParseError(`Failed to read DOCX: ${msg}`, { cause: e });
  }
}

async function _readDocxInner(
  buffer: Uint8Array,
  policy: Required<DocxSecurityPolicy>,
  checkCrc32?: boolean
): Promise<DocxDocument> {
  const reader = unzipZip(buffer, {
    checkCrc32,
    maxEntrySize: policy.maxPartSize
  });
  const entries = new Map<string, Uint8Array>();

  const budget = createPartBudget(policy);
  for await (const entry of reader.entries()) {
    // A folder entry is not a part: a part name never ends with "/" (OPC §9.1.1.1),
    // and carrying one through would declare it in [Content_Types].xml.
    if (entry.type === "directory" || /[/\\]$/.test(entry.path)) {
      continue;
    }
    budget.count();
    const data = await budget.read(entry.path, entry);
    // Normalize path: remove leading slash, normalize separators
    const path = entry.path.replace(/^\//, "").replace(/\\/g, "/");
    entries.set(path, data);
  }

  return readDocxEntries(entries, policy);
}

/** Build the document model from unzipped entries. */
function readDocxEntries(
  entries: Map<string, Uint8Array>,
  policy: Required<DocxSecurityPolicy>
): DocxDocument {
  const decoder = utf8Decoder;
  const depth = policy.maxXmlDepth;
  const consumedPaths = new Set<string>(["[Content_Types].xml"]);

  // Best-effort parse for non-critical parts (settings, numbering, styles,
  // theme, fontTable, comments, charts, headers, footers, notes, …). A
  // malformed auxiliary part should not prevent us from returning the main
  // document body. Only parse failures on document.xml itself are fatal.
  const tryParse = <T>(fn: () => T): T | undefined => {
    try {
      return fn();
    } catch (error) {
      // A resource limit is a verdict on the package, not a malformed part.
      if (error instanceof DocxLimitExceededError) {
        throw error;
      }
      return undefined;
    }
  };

  // Parse [Content_Types].xml for accurate opaque part content types
  const contentTypesXml = entries.get("[Content_Types].xml");
  const contentTypeOverrides = new Map<string, string>();
  const contentTypeDefaults = new Map<string, string>();
  if (contentTypesXml) {
    const ctDoc = parsePartXml(decoder.decode(contentTypesXml), depth);
    for (const child of ctDoc.root.children) {
      if (child.type !== "element") {
        continue;
      }
      if (child.name === "Override") {
        const partName = child.attributes["PartName"] ?? "";
        const ct = child.attributes["ContentType"] ?? "";
        if (partName && ct) {
          // Normalize: remove leading slash
          contentTypeOverrides.set(partName.replace(/^\//, ""), ct);
        }
      } else if (child.name === "Default") {
        const ext = child.attributes["Extension"] ?? "";
        const ct = child.attributes["ContentType"] ?? "";
        if (ext && ct) {
          contentTypeDefaults.set(ext.toLowerCase(), ct);
        }
      }
    }
  }

  const getText = (path: string): string | undefined => {
    const data = entries.get(path);
    if (data) {
      consumedPaths.add(path);
    }
    return data ? decoder.decode(data) : undefined;
  };

  // Parse document relationships (must be before parseDocumentXml for hyperlink resolution)
  // First, try to discover document path via package rels (supports Strict conformance)
  let documentPartPath = "word/document.xml";
  const packageRelsXmlEarly = getText("_rels/.rels");
  if (packageRelsXmlEarly) {
    const pkgRelsEarly = parseRelationships(packageRelsXmlEarly, depth);
    for (const rel of pkgRelsEarly) {
      if (rel.type === RelType.OfficeDocument) {
        let target = rel.target;
        if (target.startsWith("/")) {
          target = target.substring(1);
        }
        documentPartPath = target;
        break;
      }
    }
  }

  /** A part's own relationships, empty when it has none. */
  const readPartRels = (partPath: string): ParsedRelationship[] => {
    const relsPath = getPartRelsPath(partPath);
    const relsXml = getText(relsPath);
    if (!relsXml) {
      return [];
    }
    consumedPaths.add(relsPath);
    return tryParse(() => parseRelationships(relsXml, depth)) ?? [];
  };

  const docRelsPath = getPartRelsPath(documentPartPath);
  const docRelsXml = getText(docRelsPath);
  const docRels = docRelsXml ? parseRelationships(docRelsXml, depth) : [];
  const _relMap = new Map(docRels.map(r => [r.id, r]));

  // Every part's image relationships are bound into one document-wide table
  // before the part is parsed, so each part's content names images by keys
  // that cannot collide with another part's local ids.
  const allRelationshipIds = new Set<string>();
  for (const [path, data] of entries) {
    if (path.endsWith(".rels")) {
      for (const rel of tryParse(() => parseRelationships(decoder.decode(data), depth)) ?? []) {
        allRelationshipIds.add(rel.id);
      }
    }
  }
  const imageTable = createImageTable(
    entries,
    allRelationshipIds,
    consumedPaths,
    path => contentTypeOverrides.get(path) ?? contentTypeDefaults.get(getFileExt(path))
  );
  imageTable.bindDocument(documentPartPath, docRels);

  // Create reader context for this parse session (replaces module-level _session)
  const ctx = createReaderContext(policy);
  ctx.relMap = _relMap;

  // Parse document.xml (required)
  const documentXml = getText(documentPartPath);
  if (!documentXml) {
    throw new DocxMissingPartError(documentPartPath);
  }
  // The hyperlinks markup kept verbatim in the body refers to, resolved once
  // the body's drawings are all read (charts are read after the body).
  const bodyLinks: PreservedLink[] = [];
  const parsedDocument = readingPreservedLinks(bodyLinks, () => parseDocumentXml(documentXml, ctx));
  const { body, background, bodyRangeMarkersAfter } = parsedDocument;
  let { sectionProperties } = parsedDocument;
  let trailingMarkers = bodyRangeMarkersAfter;

  // Parse styles (resolve path via relationship, fallback to hardcoded)
  const stylesPath =
    resolveRelTarget(docRels, RelType.Styles, documentPartPath) ?? "word/styles.xml";
  const stylesXml = getText(stylesPath);
  const stylesResult = stylesXml ? tryParse(() => parseStyles(stylesXml, depth)) : undefined;

  // Parse numbering
  const numberingPath =
    resolveRelTarget(docRels, RelType.Numbering, documentPartPath) ?? "word/numbering.xml";
  const numberingSource = getText(numberingPath);
  const numberingXml = numberingSource
    ? imageTable.bindPart(numberingPath, readPartRels(numberingPath), numberingSource)
    : undefined;
  const numberingResult = numberingXml
    ? tryParse(() => parseNumberingXml(numberingXml, depth))
    : undefined;

  // Parse footnotes/endnotes — swap ctx.relMap to the notes part's own
  // .rels (footnotes.xml.rels / endnotes.xml.rels) so hyperlinks and images
  // inside notes resolve against the correct relationship map. Without this,
  // any rId used in a footnote silently resolves to undefined.
  const footnotesPath =
    resolveRelTarget(docRels, RelType.Footnotes, documentPartPath) ?? "word/footnotes.xml";
  const footnotesSource = getText(footnotesPath);
  let footnotes: ReturnType<typeof parseNotesXml> | undefined;
  if (footnotesSource) {
    const footnotesRels = readPartRels(footnotesPath);
    const footnotesXml = imageTable.bindPart(footnotesPath, footnotesRels, footnotesSource);
    const savedRelMap = ctx.relMap;
    ctx.relMap = new Map(footnotesRels.map(r => [r.id, r]));
    const links: PreservedLink[] = [];
    footnotes = readingPreservedLinks(links, () =>
      tryParse(() => parseNotesXml(footnotesXml, "footnote", ctx))
    );
    resolvePreservedLinks(links, ctx.relMap);
    ctx.relMap = savedRelMap;
  }

  const endnotesPath =
    resolveRelTarget(docRels, RelType.Endnotes, documentPartPath) ?? "word/endnotes.xml";
  const endnotesSource = getText(endnotesPath);
  let endnotes: ReturnType<typeof parseNotesXml> | undefined;
  if (endnotesSource) {
    const endnotesRels = readPartRels(endnotesPath);
    const endnotesXml = imageTable.bindPart(endnotesPath, endnotesRels, endnotesSource);
    const savedRelMap = ctx.relMap;
    ctx.relMap = new Map(endnotesRels.map(r => [r.id, r]));
    const links: PreservedLink[] = [];
    endnotes = readingPreservedLinks(links, () =>
      tryParse(() => parseNotesXml(endnotesXml, "endnote", ctx))
    );
    resolvePreservedLinks(links, ctx.relMap);
    ctx.relMap = savedRelMap;
  }

  // Parse settings — before the headers: whether the document has distinct
  // even-page headers decides which headers its pages show.
  const settingsPath =
    resolveRelTarget(docRels, RelType.Settings, documentPartPath) ?? "word/settings.xml";
  const settingsXml = getText(settingsPath);
  const settings = settingsXml
    ? tryParse(() =>
        parseSettingsXml(settingsXml, depth, {
          footnote: footnotesSource ? noteTypesById(footnotesSource, "footnote") : undefined,
          endnote: endnotesSource ? noteTypesById(endnotesSource, "endnote") : undefined
        })
      )
    : undefined;

  // Parse headers/footers + detect watermarks
  const headers = new Map<string, HeaderDef>();
  const footers = new Map<string, FooterDef>();
  // `DocxDocument.watermark` is the document's watermark: one every header a
  // section shows draws, the same in each (see `headerSlots`). It is taken
  // out of those headers, since the writer draws it there; elsewhere — a
  // header some pages show without it — it is header content.
  const headerRoots = new Map<string, { rels: ParsedRelationship[]; root: XmlElement }>();
  for (const rel of docRels) {
    if (rel.type !== RelType.Header) {
      continue;
    }
    const headerPartPath = resolvePartPath(documentPartPath, rel.target);
    const source = getText(headerPartPath);
    if (source) {
      const headerRels = readPartRels(headerPartPath);
      const xml = imageTable.bindPart(headerPartPath, headerRels, source);
      try {
        headerRoots.set(rel.id, {
          rels: headerRels,
          root: parseBoundContentPart(xml, depth).root
        });
      } catch {
        // Skip a malformed header; preserve other headers and the document.
      }
    }
  }
  const sections = [...innerSections(body), sectionProperties ?? {}];
  const slots = headerSlots(sections, settings?.evenAndOddHeaders === true);
  const documentWatermark = findDocumentWatermark(slots, headerRoots);
  const watermark = documentWatermark?.watermark;
  const watermarkOnly = new Set<string>();

  for (const rel of docRels) {
    if (rel.type === RelType.Header) {
      const entry = headerRoots.get(rel.id);
      if (entry) {
        // Switch ctx.relMap to the header's own rels so hyperlinks referenced
        // inside the header resolve against its own relationship map.
        const savedRelMap = ctx.relMap;
        ctx.relMap = new Map(entry.rels.map(r => [r.id, r]));
        try {
          const found = documentWatermark?.found.get(rel.id);
          const headerRoot = found ? withoutWatermark(entry.root, found) : entry.root;
          const links: PreservedLink[] = [];
          const content = readingPreservedLinks(links, () =>
            parseHeaderFooterRoot(headerRoot, ctx)
          );
          resolvePreservedLinks(links, ctx.relMap);
          if (found && isWatermarkOnlyHeader(content)) {
            watermarkOnly.add(rel.id);
          }
          headers.set(rel.id, { content, rId: rel.id });
        } catch {
          // Skip a malformed header; preserve other headers and the document.
        }
        ctx.relMap = savedRelMap;
      }
    } else if (rel.type === RelType.Footer) {
      const footerPartPath = resolvePartPath(documentPartPath, rel.target);
      const source = getText(footerPartPath);
      if (source) {
        // Switch ctx.relMap to the footer's own rels so hyperlinks referenced
        // inside the footer resolve against its own relationship map.
        const footerRels = readPartRels(footerPartPath);
        const xml = imageTable.bindPart(footerPartPath, footerRels, source);
        const savedRelMap = ctx.relMap;
        ctx.relMap = new Map(footerRels.map(r => [r.id, r]));
        try {
          const links: PreservedLink[] = [];
          const content = readingPreservedLinks(links, () => parseHeaderFooterXml(xml, ctx));
          resolvePreservedLinks(links, ctx.relMap);
          footers.set(rel.id, { content, rId: rel.id });
        } catch {
          // Skip a malformed footer; preserve other footers and the document.
        }
        ctx.relMap = savedRelMap;
      }
    }
  }

  // A header left holding nothing but the watermark is not kept where the
  // writer would give the section one again: a section that shows it as its
  // own and inherits no header of that type. Kept anywhere else, it stays an
  // (empty) header the writer draws the watermark in.
  const dropped = droppableWatermarkHeaders(slots, sections, watermarkOnly);
  if (dropped.size > 0) {
    for (const rId of dropped) {
      headers.delete(rId);
    }
    if (sectionProperties) {
      sectionProperties = withoutHeaderReferences(sectionProperties, dropped);
    }
    for (let i = 0; i < body.length; i++) {
      const item = body[i]!;
      const sect = item.type === "paragraph" ? item.properties?.sectionProperties : undefined;
      if (item.type === "paragraph" && sect?.headers?.some(h => dropped.has(h.rId))) {
        body[i] = {
          ...item,
          properties: {
            ...item.properties,
            sectionProperties: withoutHeaderReferences(sect, dropped)
          }
        };
      }
    }
  }

  // Parse web settings
  const webSettingsPath =
    resolveRelTarget(docRels, RelType.WebSettings, documentPartPath) ?? "word/webSettings.xml";
  const webSettingsXml = getText(webSettingsPath);
  const webSettings = webSettingsXml
    ? tryParse(() => parseWebSettings(webSettingsXml, depth))
    : undefined;

  // Parse people
  const peoplePath =
    resolveRelTarget(docRels, RelType.People, documentPartPath) ?? "word/people.xml";
  const peopleXml = getText(peoplePath);
  const people = peopleXml ? tryParse(() => parsePeople(peopleXml, depth)) : undefined;

  // Parse thumbnail (from package rels — reuse already-parsed rels)
  let thumbnail: DocxDocument["thumbnail"];
  if (packageRelsXmlEarly) {
    const pkgRels = parseRelationships(packageRelsXmlEarly, depth);
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
          const ext = getFileExt(normalized);
          const ct =
            ext === "jpeg" || ext === "jpg"
              ? "image/jpeg"
              : ext === "png"
                ? "image/png"
                : "image/x-wmf";
          thumbnail = {
            contentType: ct as "image/jpeg" | "image/x-wmf" | "image/png",
            data: thumbData
          };
        }
        break;
      }
    }
  }

  // Parse font table
  const fontTablePath =
    resolveRelTarget(docRels, RelType.FontTable, documentPartPath) ?? "word/fontTable.xml";
  const fontTableXml = getText(fontTablePath);
  const fonts = fontTableXml ? tryParse(() => parseFontTableXml(fontTableXml, depth)) : undefined;

  // Parse embedded fonts
  let embeddedFonts: EmbeddedFont[] | undefined;
  const fontTableRelsXml = getText("word/_rels/fontTable.xml.rels");
  if (fontTableRelsXml && fonts) {
    const fontRels = parseRelationships(fontTableRelsXml, depth);
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
          const fileName = getFileName(rel.target);
          const fontKey = rIdToKey.get(rel.id);
          const ef: Mutable<EmbeddedFont> = {
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
      const targetPath = resolvePartPath(documentPartPath, rel.target);
      consumedPaths.add(targetPath);
      const xmlContent = getText(targetPath);
      if (!xmlContent) {
        continue;
      }

      // Parse itemProps*.xml to get storeItemID
      const fileName = getFileName(targetPath);
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
          const propsDoc = parsePartXml(propsXml, depth);
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
  const coreProperties = corePropsXml
    ? tryParse(() => parseCoreProps(corePropsXml, depth))
    : undefined;

  // Parse app properties
  const appPropsXml = getText("docProps/app.xml");
  const appProperties = appPropsXml ? tryParse(() => parseAppProps(appPropsXml, depth)) : undefined;

  // Parse comments — switch ctx.relMap to comments.xml.rels so any
  // hyperlinks/images referenced from inside comment paragraphs resolve
  // against the comment part's own relationships rather than document.xml.rels.
  const commentsSource = getText("word/comments.xml");
  let comments: CommentDef[] | undefined;
  if (commentsSource) {
    const commentsRels = readPartRels("word/comments.xml");
    const commentsXml = imageTable.bindPart("word/comments.xml", commentsRels, commentsSource);
    const savedRelMap = ctx.relMap;
    ctx.relMap = new Map(commentsRels.map(r => [r.id, r]));
    const links: PreservedLink[] = [];
    comments = readingPreservedLinks(links, () =>
      tryParse(() => parseCommentsXmlFromCtx(commentsXml, ctx))
    );
    resolvePreservedLinks(links, ctx.relMap);
    ctx.relMap = savedRelMap;
  }

  // Merge in commentsExtended.xml data if present
  const commentsExtXml = getText("word/commentsExtended.xml");
  if (commentsExtXml && comments) {
    const extMap = tryParse(() => parseCommentsExtendedXml(commentsExtXml, depth));
    if (extMap) {
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
  }

  // Parse custom properties
  const customPropsXml = getText("docProps/custom.xml");
  const customProperties = customPropsXml
    ? tryParse(() => parseCustomPropsXml(customPropsXml, depth))
    : undefined;

  // Parse theme
  const themePath =
    resolveRelTarget(docRels, RelType.Theme, documentPartPath) ?? "word/theme/theme1.xml";
  const themeXml = getText(themePath);
  const theme = themeXml ? tryParse(() => parseThemeXml(themeXml, depth)) : undefined;

  const images = imageTable.images as ImageDef[];

  // Parse chart parts and replace opaque drawings with typed ChartContent
  const chartRIdToChart = new Map<string, Chart>();
  for (const rel of docRels) {
    if (rel.type === RelType.Chart) {
      const chartPath = resolvePartPath(documentPartPath, rel.target);
      // Read without consuming: a part the model cannot read stays opaque.
      const chartData = entries.get(chartPath);
      const chartXml = chartData ? decoder.decode(chartData) : undefined;
      if (chartXml) {
        // The chart's own relationships (its style and colour parts, its
        // embedded workbook, its user shapes); their targets stay opaque parts.
        const relationships = readPartRels(chartPath).map(r => ({
          id: r.id,
          type: r.type,
          target: r.target,
          ...(r.targetMode !== undefined && { targetMode: r.targetMode })
        }));
        const chart = tryParse(() => parseChartXml(chartXml, depth, relationships));
        // A part the model cannot read stays an opaque part, with its rels.
        if (chart) {
          consumedPaths.add(chartPath);
          chartRIdToChart.set(rel.id, chart);
        }
      }
    }
  }

  // Replace OpaqueDrawing items that reference chart rIds with proper ChartContent
  if (chartRIdToChart.size > 0) {
    readingPreservedLinks(bodyLinks, () => replaceOpaqueCharts(body, chartRIdToChart));
  }

  // Parse ChartEx parts and replace opaque drawings with typed ChartExContent
  const chartExRIdToContent = new Map<string, ChartExContent>();
  for (const rel of docRels) {
    if (rel.type === RelType.ChartEx) {
      const chartExPath = resolvePartPath(documentPartPath, rel.target);
      consumedPaths.add(chartExPath);
      const chartExXml = getText(chartExPath);
      if (chartExXml) {
        const data = tryParse(() => parseChartExXml(chartExXml, depth));
        const content: ChartExContent = {
          type: "chartEx",
          chartExXml,
          ...(data !== undefined && { data })
        };
        chartExRIdToContent.set(rel.id, content);
      }
    }
  }

  // Replace OpaqueDrawing items that reference ChartEx rIds with proper ChartExContent
  if (chartExRIdToContent.size > 0) {
    readingPreservedLinks(bodyLinks, () => replaceOpaqueChartExDrawings(body, chartExRIdToContent));
  }
  resolvePreservedLinks(bodyLinks, _relMap);

  // Detect document type from main document part content type
  let docType: DocxDocumentType | undefined;
  const mainDocCT =
    contentTypeOverrides.get(documentPartPath) ?? contentTypeOverrides.get(`/${documentPartPath}`);
  if (mainDocCT) {
    if (mainDocCT.includes("template.main") && mainDocCT.includes("macroEnabled")) {
      docType = "macroEnabledTemplate";
    } else if (mainDocCT.includes("template.main")) {
      docType = "template";
    } else if (mainDocCT.includes("macroEnabled")) {
      docType = "macroEnabledDocument";
    }
    // "document" is the default — only set if non-standard
  }

  // Extract VBA project binary for .docm/.dotm round-trip.
  // Honour `preserveVbaProject`: if disabled, mark the relationship's
  // target consumed (so opaqueParts won't retain it either) but leave
  // `vbaProject` undefined so the produced model does not surface macro
  // payloads to downstream consumers.
  let vbaProject: Uint8Array | undefined;
  for (const rel of docRels) {
    if (rel.type === RelType.VbaProject) {
      const vbaPath = resolvePartPath(documentPartPath, rel.target);
      consumedPaths.add(vbaPath);
      if (policy.preserveVbaProject) {
        vbaProject = entries.get(vbaPath);
      }
      break;
    }
  }

  // Extract OLE embedded objects wired on document.xml.rels. We surface
  // them as structured `oleObjects` so callers can query/round-trip them
  // (getOleObjectData/extractOleObjects) without depending on the part's
  // own .rels (OLE binaries carry no .rels — their relationship lives on
  // document.xml.rels). The body still references each object through an
  // opaqueDrawing carrying the same r:id. Honours `preserveOleObjects`.
  let oleObjects: OleObjectPart[] | undefined;
  if (policy.preserveOleObjects) {
    const collected: OleObjectPart[] = [];
    // progId may be recoverable from the body's <o:OLEObject ProgID="…">.
    // The <w:object> markup is preserved either as a body-level opaqueDrawing
    // or (more commonly) as a run-level opaqueRun inside a paragraph. Collect
    // raw XML from both so the ProgID round-trips.
    const progIdByRId = new Map<string, string>();
    const scanOleMarkup = (rawXml: string): void => {
      // Match ProgID + r:id from within the same <o:OLEObject> element so a
      // preview <v:imagedata r:id="…"> earlier in the markup is not mistaken
      // for the OLE binary's relationship id.
      const oleTagRe = /<o:OLEObject\b[^>]*>/gi;
      let m: RegExpExecArray | null;
      while ((m = oleTagRe.exec(rawXml)) !== null) {
        const tag = m[0];
        const progMatch = tag.match(/ProgID="([^"]+)"/i);
        const ridMatch = tag.match(/r:id="([^"]+)"/i);
        if (progMatch && ridMatch) {
          progIdByRId.set(ridMatch[1]!, progMatch[1]!);
        }
      }
    };
    for (const item of body) {
      if (item.type === "opaqueDrawing") {
        scanOleMarkup(item.rawXml);
      } else if (item.type === "paragraph") {
        for (const child of item.children) {
          if (isRun(child)) {
            for (const rc of child.content) {
              if (rc.type === "opaqueRun") {
                scanOleMarkup(rc.rawXml);
              }
            }
          }
        }
      }
    }
    for (const rel of docRels) {
      if (rel.type !== RelType.Package) {
        continue;
      }
      const olePath = resolvePartPath(documentPartPath, rel.target);
      if (!olePath.startsWith("word/embeddings/")) {
        continue;
      }
      const oleData = entries.get(olePath);
      if (!oleData) {
        continue;
      }
      consumedPaths.add(olePath);
      collected.push({
        path: olePath,
        data: oleData,
        rId: rel.id,
        progId: progIdByRId.get(rel.id),
        contentType: ContentType.OleObject
      });
    }
    if (collected.length > 0) {
      oleObjects = collected;
    }
  }

  // Glossary document (Building Blocks). The glossary is a self-contained
  // sub-document: word/glossary/document.xml plus its own companion parts
  // (styles/settings/webSettings/fontTable) referenced from
  // word/glossary/_rels/document.xml.rels. We carry document.xml verbatim as
  // `rawXml` and ALL sibling glossary parts (including the .rels) verbatim in
  // `rawParts`, so a read→write round-trip re-emits a complete, Word-valid
  // glossary rather than dropping it (opaqueParts would lose both the
  // glossaryDocument relationship and the companion .rels). The structured
  // `blocks` are not reverse-parsed.
  let glossary: GlossaryDocument | undefined;
  for (const rel of docRels) {
    if (rel.type !== RelType.Glossary) {
      continue;
    }
    const glossaryPath = resolvePartPath(documentPartPath, rel.target);
    const glossaryData = entries.get(glossaryPath);
    if (glossaryData) {
      // Collect every word/glossary/** part (document.xml, its .rels and all
      // companions) verbatim, and mark them consumed so they are not also
      // emitted via opaqueParts.
      const rawParts = new Map<string, Uint8Array>();
      for (const [path, data] of entries) {
        if (path.startsWith("word/glossary/")) {
          rawParts.set(path, data);
          consumedPaths.add(path);
        }
      }
      const rawXml = decoder.decode(glossaryData);
      const relsData = rawParts.get(getPartRelsPath(glossaryPath));
      const glossaryRels = relsData
        ? (tryParse(() => parseRelationships(decoder.decode(relsData), depth)) ?? [])
        : [];
      glossary = parseGlossaryDocument(rawXml, rawParts, glossaryRels, ctx.securityPolicy);
    }
    break;
  }

  // Resolve altChunk data: body elements of type "altChunk" reference a rId.
  // The target file is stored in docRels + entries. We populate the altChunk
  // body item with its data here AND mark the target path as consumed so the
  // opaqueParts collector below does not retain a duplicate copy that would
  // later be written back to the ZIP twice.
  //
  // Honour `preserveAltChunks`: when disabled, we still consume the target
  // path (so it doesn't leak into opaqueParts) but skip data attachment
  // and remove altChunk entries from the body before the document is
  // returned. Embedded HTML/RTF in altChunks is a common attack vector
  // for downstream renderers, so strict mode strips them entirely.
  for (const item of body) {
    if (item.type === "altChunk" && item.rId) {
      const rel = _relMap.get(item.rId);
      if (rel) {
        const target = resolvePartPath(documentPartPath, rel.target);
        const targetData = entries.get(target);
        if (targetData) {
          consumedPaths.add(target);
          if (policy.preserveAltChunks) {
            const fileName = getFileName(target);
            const mItem = item as Mutable<AltChunk>;
            mItem.data = targetData;
            mItem.fileName = fileName;
            // Infer content type from extension
            const ext = fileName ? getFileExt(fileName) : "";
            if (ext === "html" || ext === "htm") {
              mItem.contentType = "text/html";
            } else if (ext === "rtf") {
              mItem.contentType = "text/rtf";
            } else if (ext === "txt") {
              mItem.contentType = "text/plain";
            }
          }
        }
      }
    }
  }
  // Remove altChunk body entries entirely when not preserving them.
  if (!policy.preserveAltChunks) {
    for (let i = body.length - 1; i >= 0; i--) {
      if (body[i]!.type === "altChunk") {
        // Its range markers stay where it was.
        const { markersAfter } = removeBlock(body, i);
        if (markersAfter.length > 0) {
          trailingMarkers = [...markersAfter, ...(trailingMarkers ?? [])];
        }
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
    // Honour `preserveOleObjects`: when disabled, drop OLE embedding
    // binaries (word/embeddings/*.bin and similar) before they reach the
    // returned model. The relationship targets remain in their parent
    // part's .rels, so the caller is responsible for stripping or
    // ignoring those if they need a fully-clean document.
    if (
      !policy.preserveOleObjects &&
      (path.startsWith("word/embeddings/") || (path.endsWith(".bin") && path.includes("embed")))
    ) {
      continue;
    }
    // Parse rels for this part if they exist
    const partRelsPath = getPartRelsPath(path);
    const partRelsData = entries.get(partRelsPath);
    let relationships: OpaqueRelationship[] | undefined;
    if (partRelsData) {
      const rels = parseRelationships(decoder.decode(partRelsData), depth);
      relationships = rels.map(r => ({
        id: r.id,
        type: r.type,
        target: r.target,
        // Preserve the source string verbatim ("External", "Internal", or
        // any non-standard value) so opaque round-trip is byte-faithful.
        targetMode: r.targetMode
      }));
    }
    // Resolve content type from [Content_Types].xml (override > default by extension)
    let contentType: string | undefined = contentTypeOverrides.get(path);
    if (!contentType) {
      const ext = getFileExt(path);
      contentType = contentTypeDefaults.get(ext);
    }
    opaqueParts.push({ path, data, contentType, relationships });
  }

  return {
    ...(docType ? { docType } : {}),
    body,
    sectionProperties,
    ...(trailingMarkers ? { bodyRangeMarkersAfter: trailingMarkers } : {}),
    styles: stylesResult?.styles,
    docDefaults: stylesResult?.docDefaults,
    ...(stylesResult?.preservedXml ? { stylesPreservedXml: stylesResult.preservedXml } : {}),
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
    opaqueParts: opaqueParts.length > 0 ? opaqueParts : undefined,
    vbaProject,
    oleObjects,
    glossary
  };
}

/**
 * A glossary read from `word/glossary/document.xml`: its entries as
 * {@link BuildingBlock}s — those in a gallery the model names; the others
 * (content-control placeholder text and Word's internal entries) are kept as
 * markup — plus the part itself and its companions, verbatim, so an unedited
 * glossary is written back byte for byte.
 */
function parseGlossaryDocument(
  rawXml: string,
  rawParts: ReadonlyMap<string, Uint8Array>,
  rels: readonly ParsedRelationship[],
  securityPolicy: Required<DocxSecurityPolicy>
): GlossaryDocument {
  const blocks: BuildingBlock[] = [];
  const unlisted: string[] = [];
  let parsed: { root: XmlElement } | undefined;
  try {
    parsed = parseBoundContentPart(rawXml, securityPolicy.maxXmlDepth);
  } catch (e) {
    // A part too deep for the policy is still refused; markup the parser
    // cannot read is kept, as before, and lists no blocks.
    if (e instanceof DocxError) {
      throw e;
    }
  }
  // The glossary is a document of its own: its relationships and field state
  // are not the main document's.
  const ctx = createReaderContext(securityPolicy);
  ctx.relMap = new Map(rels.map(r => [r.id, r]));
  const docParts = parsed ? findChildNs(parsed.root, "docParts") : undefined;
  for (const part of docParts ? findChildrenNs(docParts, "docPart") : []) {
    const block = parseDocPart(part, ctx);
    if (block) {
      blocks.push(block);
    } else {
      unlisted.push(serializeElement(part));
    }
  }
  return {
    blocks,
    rawXml,
    rawParts,
    _blocksAsRead: JSON.stringify(blocks),
    ...(unlisted.length > 0 ? { _unlistedDocParts: unlisted } : {})
  };
}

/** One `w:docPart` as a building block, or `undefined` when its gallery is none the model names. */
function parseDocPart(part: XmlElement, ctx: ReaderContext): BuildingBlock | undefined {
  const pr = findChildNs(part, "docPartPr");
  const category = pr ? findChildNs(pr, "category") : undefined;
  const galleryEl = category ? findChildNs(category, "gallery") : undefined;
  const gallery = galleryEl ? OOXML_TO_GALLERY.get(attrVal(galleryEl, "val") ?? "") : undefined;
  const nameEl = pr ? findChildNs(pr, "name") : undefined;
  const name = nameEl ? attrVal(nameEl, "val") : undefined;
  if (!gallery || name === undefined) {
    return undefined;
  }
  const categoryNameEl = category ? findChildNs(category, "name") : undefined;
  const descriptionEl = pr ? findChildNs(pr, "description") : undefined;
  const guidEl = pr ? findChildNs(pr, "guid") : undefined;
  const body = findChildNs(part, "docPartBody");
  const content: BodyContent[] = [];
  let sectionProperties: SectionProperties | undefined;
  for (const child of body?.children ?? []) {
    if (child.type !== "element") {
      continue;
    }
    if (stripWPrefix(child.name) === "sectPr") {
      sectionProperties = parseSectionProperties(child);
      continue;
    }
    appendAll(content, parseBodyChild(child, ctx));
  }
  const description = descriptionEl ? attrVal(descriptionEl, "val") : undefined;
  const guid = guidEl ? attrVal(guidEl, "val") : undefined;
  return {
    name,
    gallery,
    category: (categoryNameEl ? attrVal(categoryNameEl, "val") : undefined) ?? "General",
    ...(description !== undefined ? { description } : {}),
    content,
    ...(sectionProperties ? { sectionProperties } : {}),
    ...(guid !== undefined ? { guid } : {})
  };
}
