/**
 * DOCX Module - Reader / Parser
 *
 * Reads a DOCX ZIP file and parses it into a DocxDocument model.
 * Uses the archive module for ZIP reading and XML module for parsing.
 */

import { unzip } from "@archive/read-archive";
import { parseXml, findChild, textContent } from "@xml/dom";
import type { XmlElement } from "@xml/types";

import { RelType, ContentType } from "../constants";
import { type Mutable, utf8Decoder } from "../core/internal-utils";
import { isRun } from "../core/text-utils";
import {
  DocxError,
  DocxParseError,
  DocxMissingPartError,
  DocxEncryptedError,
  DocxLimitExceededError
} from "../errors";
import { decryptDocx } from "../security/encryption";
import { resolveSecurityPolicy, type WordSecurityPolicy } from "../security/policy";
import type {
  DocxDocument,
  BodyContent,
  NoteType,
  Paragraph,
  ParagraphChild,
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
  ImageMediaType,
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
  Hyperlink,
  BookmarkStart,
  TextBox,
  StructuredDocumentTag,
  SdtListItem,
  SdtDateProperties,
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
  Chart,
  ChartExContent,
  DocxDocumentType
} from "../types";
import {
  replaceOpaqueCharts,
  replaceOpaqueChartExDrawings,
  parseChartXml,
  parseChartExXml
} from "./chart-parser";
import {
  parseCommentsXml as parseCommentsXmlExternal,
  parseCommentsExtendedXml
} from "./comments-parser";
import {
  parseCoreProps,
  parseAppProps,
  parseCustomPropsXml,
  parseFontTableXml
} from "./doc-props-parsers";
import { parseFfData } from "./form-field-parser";
import { parseDrawingContent, parseFloatingImage } from "./image-parsers";
import { parseMathContent, parseMathBlock } from "./math-parser";
import { parseThemeXml, parseWebSettings, parsePeople, parseSettingsXml } from "./metadata-parsers";
import { parseNumberingXml } from "./numbering-parser";
import { parseParagraphProperties, parseSectionProperties } from "./paragraph-section-parsers";
import {
  attrVal,
  attrInt,
  findChildNs,
  findChildrenNs,
  boolToggle,
  serializeElement,
  collectRIds,
  getPartRelsPath,
  getFileName,
  getFileExt,
  resolvePartPath,
  resolveRelTarget
} from "./parse-utils";
import {
  parseRunProperties,
  parseShading,
  parseTableWidth,
  parseRevisionInfo
} from "./properties-parsers";
import type { ReaderContext } from "./reader-context";
import { createFieldState, createReaderContext, parseRelationships } from "./reader-context";
import { parseCheckBox, parseTocInstruction } from "./sdt-helpers";
import { parseStyles } from "./styles-parser";
import {
  parseTableBorders,
  parseTableCellMargins,
  parseTableProperties
} from "./table-properties-parsers";
import { detectWatermarkFromRoot } from "./watermark-parser";

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
      case "ruby": {
        const ruby: Mutable<RubyContent> & { rubyText: Run[]; baseText: Run[] } = {
          type: "ruby",
          rubyText: [],
          baseText: []
        };
        const rubyPrEl = findChildNs(child, "rubyPr");
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
        // Parse w:rt (ruby text)
        const rtEl = findChildNs(child, "rt");
        if (rtEl) {
          for (const rtChild of rtEl.children) {
            if (rtChild.type === "element" && rtChild.name.replace(/^w:/, "") === "r") {
              ruby.rubyText.push(parseRun(rtChild));
            }
          }
        }
        // Parse w:rubyBase
        const baseEl = findChildNs(child, "rubyBase");
        if (baseEl) {
          for (const bChild of baseEl.children) {
            if (bChild.type === "element" && bChild.name.replace(/^w:/, "") === "r") {
              ruby.baseText.push(parseRun(bChild));
            }
          }
        }
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
        content.push({ type: "annotationReference", id: attrInt(child, "id") ?? 0 });
        break;
      case "commentReference":
        // This is annotationReference for comments inside runs
        content.push({ type: "annotationReference", id: attrInt(child, "id") ?? 0 });
        break;
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

function parseDrawingShape(
  anchorEl: XmlElement,
  wspEl: XmlElement,
  ctx: ReaderContext
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
    const paras: Paragraph[] = [];
    for (const child of txbxContentEl.children) {
      if (child.type === "element" && child.name.replace(/^w:/, "") === "p") {
        paras.push(parseParagraph(child, ctx));
      }
    }
    if (paras.length > 0) {
      shape.textContent = paras;
    }
  }

  // Parse text body vertical anchor (wps:bodyPr/@anchor)
  const bodyPrEl = findChild(wspEl, "wps:bodyPr") ?? findChildNs(wspEl, "bodyPr");
  const anchorAttr = bodyPrEl?.attributes["anchor"];
  if (anchorAttr === "t" || anchorAttr === "ctr" || anchorAttr === "b") {
    shape.textBodyAnchor = anchorAttr;
  }

  // Parse positioning
  const posH = findChild(anchorEl, "wp:positionH");
  if (posH) {
    const hp: Mutable<NonNullable<DrawingShape["horizontalPosition"]>> = {
      relativeTo: posH.attributes["relativeFrom"] as NonNullable<
        DrawingShape["horizontalPosition"]
      >["relativeTo"]
    };
    const offsetEl = findChild(posH, "wp:posOffset");
    if (offsetEl) {
      hp.offset = parseInt(textContent(offsetEl), 10);
    }
    const alignEl = findChild(posH, "wp:align");
    if (alignEl) {
      hp.align = textContent(alignEl) as NonNullable<DrawingShape["horizontalPosition"]>["align"];
    }
    shape.horizontalPosition = hp;
  }
  const posV = findChild(anchorEl, "wp:positionV");
  if (posV) {
    const vp: Mutable<NonNullable<DrawingShape["verticalPosition"]>> = {
      relativeTo: posV.attributes["relativeFrom"] as NonNullable<
        DrawingShape["verticalPosition"]
      >["relativeTo"]
    };
    const offsetEl = findChild(posV, "wp:posOffset");
    if (offsetEl) {
      vp.offset = parseInt(textContent(offsetEl), 10);
    }
    const alignEl = findChild(posV, "wp:align");
    if (alignEl) {
      vp.align = textContent(alignEl) as NonNullable<DrawingShape["verticalPosition"]>["align"];
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
      shape.wrap = {
        style: "square",
        side: wrapChild.attributes["wrapText"] as NonNullable<DrawingShape["wrap"]>["side"]
      };
    } else if (wn === "wp:wrapTight") {
      shape.wrap = {
        style: "tight",
        side: wrapChild.attributes["wrapText"] as NonNullable<DrawingShape["wrap"]>["side"]
      };
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
// TextBox Parser
// =============================================================================

function parseTextBox(pictEl: XmlElement, ctx: ReaderContext): TextBox | undefined {
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
      paragraphs.push(parseParagraph(c, ctx));
    }
  }

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

function parseSdt(
  sdtEl: XmlElement,
  ctx: ReaderContext
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
        return parseTocFromSdt(sdtContentEl, ctx);
      }
    }
  }

  // Generic SDT
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
        const item: Partial<Mutable<SdtListItem>> = { value: attrVal(li, "value") ?? "" };
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
        const item: Partial<Mutable<SdtListItem>> = { value: attrVal(li, "value") ?? "" };
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
        const binding: { xpath: string; storeItemId: string; prefixMappings?: string } = {
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
  }

  const content: (Paragraph | Run | Table | StructuredDocumentTag)[] = [];
  if (sdtContentEl) {
    for (const child of sdtContentEl.children) {
      if (child.type !== "element") {
        continue;
      }
      const n = child.name.replace(/^w:/, "");
      if (n === "p") {
        content.push(parseParagraph(child, ctx));
      } else if (n === "tbl") {
        content.push(parseTable(child, ctx));
      } else if (n === "r") {
        content.push(parseRun(child));
      } else if (n === "sdt") {
        // Nested SDT (e.g. repeating section item SDTs). Preserve the
        // inner SDT verbatim — including its own properties — so data
        // binding, alias, lock and similar metadata round-trip correctly.
        const inner = parseSdt(child, ctx);
        if (inner && inner.type === "sdt") {
          content.push(inner);
        }
      }
    }
  }

  return { type: "sdt", properties: props, content };
}

function parseTocFromSdt(
  sdtContentEl: XmlElement | undefined,
  ctx: ReaderContext
): TableOfContents {
  const toc: Mutable<TableOfContents> = { type: "tableOfContents" };
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
        cachedParagraphs.push(parseParagraph(child, ctx));
      }
    }
  }

  if (cachedParagraphs.length > 0) {
    toc.cachedParagraphs = cachedParagraphs;
  }

  return toc;
}

/** Parse a TOC field instruction string (e.g. `TOC \o "1-3" \h \t "Style,1" \c "Figure"`). */

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

function parseParagraph(pEl: XmlElement, ctx: ReaderContext): Paragraph {
  const pPrEl = findChildNs(pEl, "pPr");
  const children: ParagraphChild[] = [];

  // Field state machine lives on ctx so that complex fields (TOC, INDEX,
  // long REF/SEQ chains) can span paragraph boundaries — the matching
  // `<w:fldChar fldCharType="end">` may occur in a later paragraph than the
  // `begin`. Storing state on ctx is also safe because part-scoped parsers
  // (header/footer/footnote/endnote/comment) save and reset it on entry.
  const field = ctx.field;

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
              field.state = "instrText";
              field.instr = "";
              field.cached = "";
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
              const fc: FieldContent = {
                type: "field",
                instruction: field.instr.trim(),
                cachedValue: field.cached || undefined,
                formField: field.formField
              };
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
          } else if (rcName === "instrText" && field.state === "instrText") {
            hasFldChar = true;
            field.instr += textContent(rc);
          }
        }

        if (field.state === "cached") {
          // Collect cached text from this run
          for (const rc of resolved.children) {
            if (rc.type !== "element") {
              continue;
            }
            const rcName = rc.name.replace(/^w:/, "");
            if (rcName === "t") {
              field.cached += textContent(rc);
            } else if (rcName === "fldChar") {
              // Already handled above
            }
          }
          if (!hasFldChar) {
            continue; // Skip adding this run normally
          }
        }

        if (field.state === "instrText" && hasFldChar) {
          continue; // Don't add begin/instrText runs as normal content
        }
        if (field.state === "none" && !hasFldChar) {
          // Detect a degenerate `<w:r>` whose only meaningful child is
          // `<w:commentReference>`. The OOXML schema requires the leaf
          // to live inside a w:r, but at the model level we represent
          // it as a paragraph-child `commentReference`. Hoisting here
          // means a round-trip preserves the model shape instead of
          // collapsing to `annotationReference`.
          let onlyCommentRefId: number | undefined;
          let onlyCommentRefSeen = false;
          let hasOtherMeaningfulChild = false;
          for (const rcc of resolved.children) {
            if (rcc.type !== "element") {
              continue;
            }
            const rccName = rcc.name.replace(/^w:/, "");
            if (rccName === "rPr") {
              continue;
            }
            if (rccName === "commentReference") {
              if (onlyCommentRefSeen) {
                // Multiple commentReferences in one run is malformed;
                // fall through to the generic run parser.
                hasOtherMeaningfulChild = true;
                break;
              }
              onlyCommentRefSeen = true;
              const idAttr = rcc.attributes["w:id"] ?? rcc.attributes["id"];
              const id = idAttr !== undefined ? parseInt(idAttr, 10) : NaN;
              if (!Number.isNaN(id)) {
                onlyCommentRefId = id;
              }
            } else {
              hasOtherMeaningfulChild = true;
              break;
            }
          }
          if (onlyCommentRefSeen && !hasOtherMeaningfulChild && onlyCommentRefId !== undefined) {
            children.push({ type: "commentReference", id: onlyCommentRefId });
          } else {
            children.push(parseRun(resolved));
          }
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
        children.push(hyperlink);
        break;
      }
      case "bookmarkStart": {
        const idAttr = resolved.attributes["w:id"] ?? resolved.attributes["id"];
        const id = idAttr !== undefined ? parseInt(idAttr, 10) : NaN;
        if (Number.isNaN(id)) {
          // Without a valid id we can't pair this with a bookmarkEnd; drop it
          // rather than fabricate id=0 (which would collide with every other
          // bookmark missing an id and corrupt cross-references on round-trip).
          break;
        }
        const bm: Mutable<BookmarkStart> = {
          type: "bookmarkStart",
          id,
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
      case "bookmarkEnd": {
        const idAttr = resolved.attributes["w:id"] ?? resolved.attributes["id"];
        const id = idAttr !== undefined ? parseInt(idAttr, 10) : NaN;
        if (Number.isNaN(id)) {
          break;
        }
        children.push({ type: "bookmarkEnd", id });
        break;
      }
      case "commentRangeStart": {
        const idAttr = resolved.attributes["w:id"] ?? resolved.attributes["id"];
        const id = idAttr !== undefined ? parseInt(idAttr, 10) : NaN;
        if (Number.isNaN(id)) {
          break;
        }
        children.push({ type: "commentRangeStart", id });
        break;
      }
      case "commentRangeEnd": {
        const idAttr = resolved.attributes["w:id"] ?? resolved.attributes["id"];
        const id = idAttr !== undefined ? parseInt(idAttr, 10) : NaN;
        if (Number.isNaN(id)) {
          break;
        }
        children.push({ type: "commentRangeEnd", id });
        break;
      }
      case "commentReference": {
        const idAttr = resolved.attributes["w:id"] ?? resolved.attributes["id"];
        const id = idAttr !== undefined ? parseInt(idAttr, 10) : NaN;
        if (Number.isNaN(id)) {
          break;
        }
        children.push({ type: "commentReference", id });
        break;
      }
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
          const marker: {
            type: string;
            id: number;
            author?: string;
            date?: string;
            name?: string;
          } = {
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
          children.push(marker as ParagraphChild);
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
          const marker: { type: string; id: number; author?: string; date?: string } = {
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
          children.push(marker as ParagraphChild);
        }
        break;
      }
      case "smartTag":
      case "customXml":
      case "dir": {
        // Semantic wrappers: flatten their children into the current
        // paragraph. The wrapper's own properties element (smartTagPr,
        // customXmlPr, …) is not a paragraph child and would otherwise
        // fall through to the `default` branch below and be emitted as a
        // bogus `opaqueParagraphChild` containing the properties XML —
        // poisoning the paragraph on round-trip. Build a synthetic element
        // that excludes those `*Pr` siblings before recursing.
        const filteredChildren = resolved.children.filter(c => {
          if (c.type !== "element") {
            return true;
          }
          const ln = c.name.replace(/^w:/, "");
          return ln !== "smartTagPr" && ln !== "customXmlPr";
        });
        const surrogate: XmlElement = {
          ...resolved,
          children: filteredChildren
        };
        const subPara = parseParagraph(surrogate, ctx);
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
      default:
        // Unknown paragraph child: preserve as opaque for round-trip fidelity
        children.push({
          type: "opaqueParagraphChild",
          rawXml: serializeElement(resolved)
        });
        break;
    }
  }

  const paraId = pEl.attributes["w14:paraId"];
  const textId = pEl.attributes["w14:textId"];

  const result: Mutable<Paragraph> = {
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
      content.push({
        type: "break",
        breakType: attrVal(child, "type") as "page" | "column" | "textWrapping" | undefined
      });
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

function parseTableCell(el: XmlElement, ctx: ReaderContext): TableCell {
  const tcPrEl = findChildNs(el, "tcPr");
  const content: (Paragraph | Table)[] = [];

  for (const child of el.children) {
    if (child.type !== "element") {
      continue;
    }
    const name = child.name.replace(/^w:/, "");
    if (name === "p") {
      content.push(parseParagraph(child, ctx));
    } else if (name === "tbl") {
      content.push(parseTable(child, ctx));
    } else if (name === "sdt") {
      // SDT inside a table cell. The TableCell.content union does not
      // include StructuredDocumentTag, so we flatten the SDT's inner
      // paragraphs/tables into the cell. SDT-level metadata (data binding,
      // alias, repeating section, …) is lost on round-trip but visible
      // content is preserved — better than dropping the runs entirely.
      const sdt = parseSdt(child, ctx);
      if (sdt && sdt.type === "sdt") {
        for (const c of sdt.content) {
          if ((c as { type?: string }).type === "paragraph") {
            content.push(c as Paragraph);
          } else if ((c as { type?: string }).type === "table") {
            content.push(c as Table);
          }
          // Run-only and nested-SDT children cannot live as direct
          // siblings of <w:p>/<w:tbl> in a <w:tc>, so they are dropped.
        }
      }
    }
  }

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
      p.verticalAlign = attrVal(vAlignEl, "val") as TableCellProperties["verticalAlign"];
    }

    if (findChildNs(tcPrEl, "noWrap")) {
      p.noWrap = true;
    }

    const textDirEl = findChildNs(tcPrEl, "textDirection");
    if (textDirEl) {
      p.textDirection = attrVal(textDirEl, "val") as TableCellProperties["textDirection"];
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

function parseTableRow(el: XmlElement, ctx: ReaderContext): TableRow {
  const trPrEl = findChildNs(el, "trPr");
  const tblPrExEl = findChildNs(el, "tblPrEx");
  const cells: TableCell[] = [];

  for (const child of el.children) {
    if (child.type === "element" && child.name.replace(/^w:/, "") === "tc") {
      cells.push(parseTableCell(child, ctx));
    }
  }

  let props: TableRowProperties | undefined;
  if (trPrEl || tblPrExEl) {
    const p: Mutable<TableRowProperties> = {};
    if (tblPrExEl) {
      p.tblPrEx = parseTableProperties(tblPrExEl);
    }
    if (trPrEl) {
      const heightEl = findChildNs(trPrEl, "trHeight");
      if (heightEl) {
        p.height = {
          value: attrInt(heightEl, "val") ?? 0,
          rule: attrVal(heightEl, "hRule") as NonNullable<TableRowProperties["height"]>["rule"]
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
  const p: Mutable<TableRowProperties> = {};
  const heightEl = findChildNs(trPrEl, "trHeight");
  if (heightEl) {
    p.height = {
      value: attrInt(heightEl, "val") ?? 0,
      rule: attrVal(heightEl, "hRule") as NonNullable<TableRowProperties["height"]>["rule"]
    };
  }
  if (findChildNs(trPrEl, "tblHeader")) {
    p.tableHeader = true;
  }
  if (findChildNs(trPrEl, "cantSplit")) {
    p.cantSplit = true;
  }
  return p;
}

function parseTable(tblEl: XmlElement, ctx: ReaderContext): Table {
  const tblPrEl = findChildNs(tblEl, "tblPr");
  const gridEl = findChildNs(tblEl, "tblGrid");
  const rows: TableRow[] = [];

  for (const child of tblEl.children) {
    if (child.type === "element" && child.name.replace(/^w:/, "") === "tr") {
      rows.push(parseTableRow(child, ctx));
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
// Footnotes/Endnotes Parser
// =============================================================================

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
        if (child.type !== "element") {
          continue;
        }
        const ln = child.name.replace(/^w:/, "");
        if (ln === "p") {
          content.push(parseParagraph(child, ctx));
        } else if (ln === "sdt") {
          // SDT inside a footnote/endnote: the model's content type is
          // `Paragraph[]`, so flatten the SDT's inner paragraphs (and their
          // descendants reachable as paragraphs). SDT-level metadata is
          // dropped here on round-trip — better than losing the visible
          // text completely.
          const sdt = parseSdt(child, ctx);
          if (sdt && sdt.type === "sdt") {
            for (const c of sdt.content) {
              if ((c as { type?: string }).type === "paragraph") {
                content.push(c as Paragraph);
              }
            }
          }
        }
      }

      const note: { id: number; type?: NoteType; content: Paragraph[] } = { id, content };
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
  return parseHeaderFooterRoot(parseXml(xmlStr).root, ctx);
}

function parseHeaderFooterRoot(root: XmlElement, ctx: ReaderContext): HeaderFooterContent {
  // Header/footer parts are self-contained: reset field state on entry so an
  // unterminated complex field in the body does not consume header/footer runs.
  const savedField = ctx.field;
  ctx.field = createFieldState();
  try {
    const children: (Paragraph | Table)[] = [];
    for (const child of root.children) {
      if (child.type !== "element") {
        continue;
      }
      const name = child.name.replace(/^w:/, "");
      if (name === "p") {
        children.push(parseParagraph(child, ctx));
      } else if (name === "tbl") {
        children.push(parseTable(child, ctx));
      } else if (name === "sdt") {
        // Flatten SDT children. HeaderFooterContent.children is
        // `(Paragraph | Table)[]` so we hoist the inner paragraphs/tables;
        // SDT-level metadata is dropped on round-trip but visible content
        // is preserved (better than losing the runs entirely).
        const sdt = parseSdt(child, ctx);
        if (sdt && sdt.type === "sdt") {
          for (const c of sdt.content) {
            if ((c as { type?: string }).type === "paragraph") {
              children.push(c as Paragraph);
            } else if ((c as { type?: string }).type === "table") {
              children.push(c as Table);
            }
          }
        }
      }
    }
    return { children };
  } finally {
    ctx.field = savedField;
  }
}

/** Detect watermark from a header's parsed XML root element. */
// =============================================================================
// Comments Parser
// =============================================================================

function parseCommentsXmlFromCtx(xmlStr: string, ctx: ReaderContext): CommentDef[] {
  return parseCommentsXmlExternal(xmlStr, ctx, parseParagraph);
}

// =============================================================================
// Main Document Parser
// =============================================================================

/** Recursively extract floating images, drawing shapes, and opaque drawings from an element tree. */
function extractFloatingContent(
  el: XmlElement,
  images: FloatingImage[],
  shapes: DrawingShape[],
  opaqueDrawings: OpaqueDrawing[],
  ctx: ReaderContext
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
        const shape = parseDrawingShape(child, wspEl, ctx);
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
    } else {
      extractFloatingContent(child, images, shapes, opaqueDrawings, ctx);
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
  if (!para.children || para.children.length === 0) {
    return true;
  }
  for (const child of para.children) {
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
 * Remove `opaqueRun` entries that wrap a non-picture `<wp:inline>` drawing.
 *
 * These are emitted by parseDrawingContent so the drawing survives a
 * round-trip when its containing paragraph lives inside a table cell, header,
 * footer or SDT (places where the body-level extractor never runs). At the
 * body level, however, the same drawings are also captured as `OpaqueDrawing`
 * entries by extractFloatingContent — keeping both would duplicate the
 * drawing in the produced document. Mutates `para.children`/run content in
 * place.
 */
function stripInlineDrawingOpaqueRuns(para: Paragraph): void {
  for (const child of para.children) {
    if (!isRun(child)) {
      continue;
    }
    const run = child as Mutable<Run> & { content: RunContent[] };
    let i = 0;
    while (i < run.content.length) {
      const c = run.content[i];
      if (
        c.type === "opaqueRun" &&
        c.rawXml.includes("<wp:inline") &&
        !c.rawXml.includes("<pic:pic")
      ) {
        run.content.splice(i, 1);
      } else {
        i++;
      }
    }
  }
}

function parseDocumentXml(
  xmlStr: string,
  ctx: ReaderContext
): {
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

  const bodyEl = findChildNs(root, "body") ?? findChild(root, "w:body");
  if (!bodyEl) {
    throw new DocxParseError("Missing w:body element in document.xml");
  }

  const body: BodyContent[] = [];
  let sectionProperties: SectionProperties | undefined;

  // Instead of extracting floating content from the entire body tree and
  // appending at the end (which loses positional information), we now extract
  // floating content per-paragraph and insert it immediately after the
  // paragraph it belongs to.

  for (const child of bodyEl.children) {
    if (child.type !== "element") {
      continue;
    }
    const name = child.name.replace(/^w:/, "");

    switch (name) {
      case "p": {
        // Per OOXML schema (CT_OMathPara is a member of EG_PContent), a
        // body-level math block is encoded as a paragraph containing a
        // single m:oMathPara child. Detect that shape and surface it as
        // a top-level MathBlock so the document model stays flat — the
        // writer reverses this by re-wrapping math blocks in <w:p>.
        const mathParaChildren = child.children.filter(
          c => c.type === "element" && c.name === "m:oMathPara"
        );
        const otherChildren = child.children.filter(c => {
          if (c.type !== "element") {
            return false;
          }
          // pPr is allowed; everything else (runs, hyperlinks, etc.) means
          // we're NOT a synthetic math wrapper and must keep the paragraph.
          return c.name !== "w:pPr" && c.name !== "m:oMathPara";
        });
        if (mathParaChildren.length > 0 && otherChildren.length === 0) {
          for (const oMathPara of mathParaChildren) {
            if (oMathPara.type === "element") {
              body.push(parseMathBlock(oMathPara));
            }
          }
          break;
        }

        const para = parseParagraph(child, ctx);
        // Extract floating content from this paragraph element and insert
        // immediately after it to preserve document position.
        const pFloatingImages: FloatingImage[] = [];
        const pDrawingShapes: DrawingShape[] = [];
        const pOpaqueDrawings: OpaqueDrawing[] = [];
        extractFloatingContent(child, pFloatingImages, pDrawingShapes, pOpaqueDrawings, ctx);

        // parseDrawingContent (called from parseRunContent) already preserved
        // every non-picture inline drawing as an `opaqueRun` so the drawing
        // survives a round-trip even inside cells/headers/footers/SDTs where
        // this body-level extractor is not invoked. At the body level
        // extractFloatingContent has now also captured those drawings as
        // `OpaqueDrawing` entries — that is the form chart-parser is wired
        // to look for when promoting them to `ChartContent`. To avoid
        // duplicate output we strip any opaqueRun whose XML embeds a
        // <wp:inline> drawing from the paragraph here.
        if (pOpaqueDrawings.length > 0) {
          stripInlineDrawingOpaqueRuns(para);
        }

        // If the paragraph is otherwise empty AND we did extract anchored
        // content out of it, treat the paragraph as a synthetic carrier for
        // the floating drawing(s) and drop it. Otherwise keeping it would
        // cause a phantom empty paragraph to accumulate on every round-trip
        // (writer wraps floating images in their own <w:p>, reader pulls the
        // anchor out, leaving an empty <w:p> behind).
        const hasAnchoredContent =
          pFloatingImages.length > 0 || pDrawingShapes.length > 0 || pOpaqueDrawings.length > 0;
        const paragraphIsEmpty = isEmptyParagraph(para);
        if (!(hasAnchoredContent && paragraphIsEmpty)) {
          body.push(para);
        }

        for (const fi of pFloatingImages) {
          body.push(fi);
        }
        for (const ds of pDrawingShapes) {
          body.push(ds);
        }
        for (const od of pOpaqueDrawings) {
          body.push(od);
        }
        break;
      }
      case "tbl":
        body.push(parseTable(child, ctx));
        break;
      case "sectPr":
        // Final section properties at the body level
        sectionProperties = parseSectionProperties(child);
        break;
      case "sdt": {
        const sdtResult = parseSdt(child, ctx);
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
          const tb = parseTextBox(child, ctx);
          if (tb) {
            body.push(tb);
          }
        }
        break;
      }
    }
  }

  return { body, sectionProperties, background };
}

// =============================================================================
// Public API - Read DOCX
// =============================================================================

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
  readonly securityPolicy?: WordSecurityPolicy;
}

/**
 * Read a DOCX file from a Uint8Array buffer and parse it into a DocxDocument model.
 *
 * If the file is encrypted (CFB format), provide a password via the options parameter
 * to decrypt it automatically.
 */
export async function readDocx(
  buffer: Uint8Array,
  options?: ReadDocxOptions
): Promise<DocxDocument> {
  const policy = resolveSecurityPolicy(options?.securityPolicy);

  // Defense-in-depth: reject obviously oversized packages up-front. The same
  // limit is also enforced incrementally during entry decompression so a
  // pathological deflate stream can't slip past this check.
  if (buffer.length > policy.maxPackageSize) {
    throw new DocxLimitExceededError(
      "packageSize",
      policy.maxPackageSize,
      buffer.length,
      "compressed input larger than maxPackageSize"
    );
  }

  // Detect encrypted DOCX (CFB format) before attempting ZIP parse.
  // CFB signature: D0 CF 11 E0 A1 B1 1A E1
  if (
    buffer.length >= 8 &&
    buffer[0] === 0xd0 &&
    buffer[1] === 0xcf &&
    buffer[2] === 0x11 &&
    buffer[3] === 0xe0 &&
    buffer[4] === 0xa1 &&
    buffer[5] === 0xb1 &&
    buffer[6] === 0x1a &&
    buffer[7] === 0xe1
  ) {
    if (options?.password != null) {
      // Pass the security policy's package-size cap so a hostile CFB cannot
      // claim a multi-GiB decrypted size and force a huge buffer allocation
      // before the unzip stage even runs.
      const decryptedZip = await decryptDocx(buffer, options.password, policy.maxPackageSize);
      return readDocx(decryptedZip, options);
    }
    throw new DocxEncryptedError();
  }

  try {
    return await _readDocxInner(buffer, policy);
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
  policy: Required<WordSecurityPolicy>
): Promise<DocxDocument> {
  const reader = unzip(buffer);
  const entries = new Map<string, Uint8Array>();

  let totalUncompressed = 0;
  let entryCount = 0;
  for await (const entry of reader.entries()) {
    entryCount++;
    if (entryCount > policy.maxPartCount) {
      throw new DocxLimitExceededError(
        "partCount",
        policy.maxPartCount,
        entryCount,
        "ZIP contains more entries than maxPartCount"
      );
    }
    const data = await entry.bytes();
    if (data.length > policy.maxPartSize) {
      throw new DocxLimitExceededError(
        "partSize",
        policy.maxPartSize,
        data.length,
        `entry "${entry.path}" exceeds maxPartSize`
      );
    }
    totalUncompressed += data.length;
    if (totalUncompressed > policy.maxPackageSize) {
      throw new DocxLimitExceededError(
        "packageSize",
        policy.maxPackageSize,
        totalUncompressed,
        "cumulative uncompressed entry size exceeds maxPackageSize"
      );
    }
    // Normalize path: remove leading slash, normalize separators
    const path = entry.path.replace(/^\//, "").replace(/\\/g, "/");
    entries.set(path, data);
  }

  const decoder = utf8Decoder;
  const consumedPaths = new Set<string>(["[Content_Types].xml"]);

  // Best-effort parse for non-critical parts (settings, numbering, styles,
  // theme, fontTable, comments, charts, headers, footers, notes, …). A
  // malformed auxiliary part should not prevent us from returning the main
  // document body. Only parse failures on document.xml itself are fatal.
  const tryParse = <T>(fn: () => T): T | undefined => {
    try {
      return fn();
    } catch {
      return undefined;
    }
  };

  // Parse [Content_Types].xml for accurate opaque part content types
  const contentTypesXml = entries.get("[Content_Types].xml");
  const contentTypeOverrides = new Map<string, string>();
  const contentTypeDefaults = new Map<string, string>();
  if (contentTypesXml) {
    const ctDoc = parseXml(decoder.decode(contentTypesXml));
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
    const pkgRelsEarly = parseRelationships(packageRelsXmlEarly);
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

  const docRelsPath = getPartRelsPath(documentPartPath);
  const docRelsXml = getText(docRelsPath);
  const docRels = docRelsXml ? parseRelationships(docRelsXml) : [];
  const _relMap = new Map(docRels.map(r => [r.id, r]));

  // Create reader context for this parse session (replaces module-level _session)
  const ctx = createReaderContext(policy);
  ctx.relMap = _relMap;

  // Parse document.xml (required)
  const documentXml = getText(documentPartPath);
  if (!documentXml) {
    throw new DocxMissingPartError(documentPartPath);
  }
  const { body, sectionProperties, background } = parseDocumentXml(documentXml, ctx);

  // Parse styles (resolve path via relationship, fallback to hardcoded)
  const stylesPath =
    resolveRelTarget(docRels, RelType.Styles, documentPartPath) ?? "word/styles.xml";
  const stylesXml = getText(stylesPath);
  const stylesResult = stylesXml ? tryParse(() => parseStyles(stylesXml)) : undefined;

  // Parse numbering
  const numberingPath =
    resolveRelTarget(docRels, RelType.Numbering, documentPartPath) ?? "word/numbering.xml";
  const numberingXml = getText(numberingPath);
  const numberingResult = numberingXml
    ? tryParse(() => parseNumberingXml(numberingXml))
    : undefined;

  // Parse footnotes/endnotes — swap ctx.relMap to the notes part's own
  // .rels (footnotes.xml.rels / endnotes.xml.rels) so hyperlinks and images
  // inside notes resolve against the correct relationship map. Without this,
  // any rId used in a footnote silently resolves to undefined.
  const footnotesPath =
    resolveRelTarget(docRels, RelType.Footnotes, documentPartPath) ?? "word/footnotes.xml";
  const footnotesXml = getText(footnotesPath);
  let footnotes: ReturnType<typeof parseNotesXml> | undefined;
  if (footnotesXml) {
    const footnotesRelsPath = getPartRelsPath(footnotesPath);
    const footnotesRelsXml = getText(footnotesRelsPath);
    const savedRelMap = ctx.relMap;
    if (footnotesRelsXml) {
      const footnotesRels = parseRelationships(footnotesRelsXml);
      ctx.relMap = new Map(footnotesRels.map(r => [r.id, r]));
      consumedPaths.add(footnotesRelsPath);
    } else {
      ctx.relMap = new Map();
    }
    footnotes = tryParse(() => parseNotesXml(footnotesXml, "footnote", ctx));
    ctx.relMap = savedRelMap;
  }

  const endnotesPath =
    resolveRelTarget(docRels, RelType.Endnotes, documentPartPath) ?? "word/endnotes.xml";
  const endnotesXml = getText(endnotesPath);
  let endnotes: ReturnType<typeof parseNotesXml> | undefined;
  if (endnotesXml) {
    const endnotesRelsPath = getPartRelsPath(endnotesPath);
    const endnotesRelsXml = getText(endnotesRelsPath);
    const savedRelMap = ctx.relMap;
    if (endnotesRelsXml) {
      const endnotesRels = parseRelationships(endnotesRelsXml);
      ctx.relMap = new Map(endnotesRels.map(r => [r.id, r]));
      consumedPaths.add(endnotesRelsPath);
    } else {
      ctx.relMap = new Map();
    }
    endnotes = tryParse(() => parseNotesXml(endnotesXml, "endnote", ctx));
    ctx.relMap = savedRelMap;
  }

  // Parse headers/footers + detect watermarks
  const headers = new Map<string, HeaderDef>();
  const footers = new Map<string, FooterDef>();
  let watermark: Watermark | undefined;

  for (const rel of docRels) {
    if (rel.type === RelType.Header) {
      const headerPartPath = resolvePartPath(documentPartPath, rel.target);
      const xml = getText(headerPartPath);
      if (xml) {
        // Parse header's own rels and switch ctx.relMap so hyperlinks/images
        // referenced inside the header resolve against its own relationship map.
        const headerRelsPath = getPartRelsPath(headerPartPath);
        const headerRelsXml = getText(headerRelsPath);
        const savedRelMap = ctx.relMap;
        if (headerRelsXml) {
          const headerRels = parseRelationships(headerRelsXml);
          const headerRelMap = new Map(headerRels.map(r => [r.id, r]));
          ctx.relMap = headerRelMap;
          consumedPaths.add(headerRelsPath);
        } else {
          ctx.relMap = new Map();
        }
        try {
          // Parse XML once, re-use for both header content and watermark detection
          const headerRoot = parseXml(xml).root;
          headers.set(rel.id, { content: parseHeaderFooterRoot(headerRoot, ctx), rId: rel.id });
          if (!watermark) {
            watermark = detectWatermarkFromRoot(headerRoot);
          }
        } catch {
          // Skip a malformed header; preserve other headers and the document.
        }
        ctx.relMap = savedRelMap;
      }
    } else if (rel.type === RelType.Footer) {
      const footerPartPath = resolvePartPath(documentPartPath, rel.target);
      const xml = getText(footerPartPath);
      if (xml) {
        // Parse footer's own rels and switch ctx.relMap so hyperlinks/images
        // referenced inside the footer resolve against its own relationship map.
        const footerRelsPath = getPartRelsPath(footerPartPath);
        const footerRelsXml = getText(footerRelsPath);
        const savedRelMap = ctx.relMap;
        if (footerRelsXml) {
          const footerRels = parseRelationships(footerRelsXml);
          const footerRelMap = new Map(footerRels.map(r => [r.id, r]));
          ctx.relMap = footerRelMap;
          consumedPaths.add(footerRelsPath);
        } else {
          ctx.relMap = new Map();
        }
        try {
          footers.set(rel.id, { content: parseHeaderFooterXml(xml, ctx), rId: rel.id });
        } catch {
          // Skip a malformed footer; preserve other footers and the document.
        }
        ctx.relMap = savedRelMap;
      }
    }
  }

  // Parse settings
  const settingsPath =
    resolveRelTarget(docRels, RelType.Settings, documentPartPath) ?? "word/settings.xml";
  const settingsXml = getText(settingsPath);
  const settings = settingsXml ? tryParse(() => parseSettingsXml(settingsXml)) : undefined;

  // Parse web settings
  const webSettingsPath =
    resolveRelTarget(docRels, RelType.WebSettings, documentPartPath) ?? "word/webSettings.xml";
  const webSettingsXml = getText(webSettingsPath);
  const webSettings = webSettingsXml ? tryParse(() => parseWebSettings(webSettingsXml)) : undefined;

  // Parse people
  const peoplePath =
    resolveRelTarget(docRels, RelType.People, documentPartPath) ?? "word/people.xml";
  const peopleXml = getText(peoplePath);
  const people = peopleXml ? tryParse(() => parsePeople(peopleXml)) : undefined;

  // Parse thumbnail (from package rels — reuse already-parsed rels)
  let thumbnail: DocxDocument["thumbnail"];
  if (packageRelsXmlEarly) {
    const pkgRels = parseRelationships(packageRelsXmlEarly);
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
  const fonts = fontTableXml ? tryParse(() => parseFontTableXml(fontTableXml)) : undefined;

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
  const coreProperties = corePropsXml ? tryParse(() => parseCoreProps(corePropsXml)) : undefined;

  // Parse app properties
  const appPropsXml = getText("docProps/app.xml");
  const appProperties = appPropsXml ? tryParse(() => parseAppProps(appPropsXml)) : undefined;

  // Parse comments — switch ctx.relMap to comments.xml.rels so any
  // hyperlinks/images referenced from inside comment paragraphs resolve
  // against the comment part's own relationships rather than document.xml.rels.
  const commentsXml = getText("word/comments.xml");
  let comments: CommentDef[] | undefined;
  if (commentsXml) {
    const commentsRelsPath = "word/_rels/comments.xml.rels";
    const commentsRelsXml = getText(commentsRelsPath);
    const savedRelMap = ctx.relMap;
    if (commentsRelsXml) {
      const commentsRels = parseRelationships(commentsRelsXml);
      ctx.relMap = new Map(commentsRels.map(r => [r.id, r]));
      consumedPaths.add(commentsRelsPath);
    } else {
      ctx.relMap = new Map();
    }
    comments = tryParse(() => parseCommentsXmlFromCtx(commentsXml, ctx));
    ctx.relMap = savedRelMap;
  }

  // Merge in commentsExtended.xml data if present
  const commentsExtXml = getText("word/commentsExtended.xml");
  if (commentsExtXml && comments) {
    const extMap = tryParse(() => parseCommentsExtendedXml(commentsExtXml));
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
    ? tryParse(() => parseCustomPropsXml(customPropsXml))
    : undefined;

  // Parse theme
  const themePath =
    resolveRelTarget(docRels, RelType.Theme, documentPartPath) ?? "word/theme/theme1.xml";
  const themeXml = getText(themePath);
  const theme = themeXml ? tryParse(() => parseThemeXml(themeXml)) : undefined;

  // Collect images from main document relationships
  const images: ImageDef[] = [];
  for (const rel of docRels) {
    if (rel.type === RelType.Image) {
      const imgPath = resolvePartPath(documentPartPath, rel.target);
      consumedPaths.add(imgPath);
      const data = entries.get(imgPath);
      if (data) {
        const fileName = getFileName(rel.target);
        const ext = getFileExt(fileName) || "png";
        images.push({
          data,
          mediaType: ext as ImageMediaType,
          fileName,
          rId: rel.id
        });
      }
    }
  }

  // Also collect images from header/footer relationships to ensure full round-trip.
  // Headers and footers have their own .rels files which may reference unique
  // images, OR they may share a media file with the main document. In the
  // latter case we keep the existing ImageDef but record the local rId as an
  // alias so the packager can rebuild header1.xml.rels with the original
  // (header-local) id intact.
  const collectedImagePaths = new Map<string, ImageDef>();
  for (const img of images) {
    collectedImagePaths.set(img.fileName, img);
  }
  for (const rel of docRels) {
    if (rel.type !== RelType.Header && rel.type !== RelType.Footer) {
      continue;
    }
    const partPath = resolvePartPath(documentPartPath, rel.target);
    const partRelsPath = getPartRelsPath(partPath);
    const partRelsXml = entries.get(partRelsPath);
    if (!partRelsXml) {
      continue;
    }
    const partRels = parseRelationships(decoder.decode(partRelsXml));
    for (const pRel of partRels) {
      if (pRel.type !== RelType.Image) {
        continue;
      }
      const imgPath = resolvePartPath(partPath, pRel.target);
      consumedPaths.add(imgPath);
      const data = entries.get(imgPath);
      if (!data) {
        continue;
      }
      const fileName = getFileName(pRel.target);
      const existing = collectedImagePaths.get(fileName);
      if (existing) {
        // Same physical file as one we already know — keep one ImageDef and
        // append this part-local rId to its aliases (if it differs from the
        // primary rId and isn't already recorded).
        if (pRel.id && pRel.id !== existing.rId) {
          const aliases = existing.aliasRIds ? [...existing.aliasRIds] : [];
          if (!aliases.includes(pRel.id)) {
            aliases.push(pRel.id);
            (existing as { aliasRIds?: readonly string[] }).aliasRIds = aliases;
          }
        }
        continue;
      }
      const ext = getFileExt(fileName) || "png";
      const newImg: ImageDef = {
        data,
        mediaType: ext as ImageMediaType,
        fileName,
        rId: pRel.id
      };
      images.push(newImg);
      collectedImagePaths.set(fileName, newImg);
    }
  }

  // Parse chart parts and replace opaque drawings with typed ChartContent
  const chartRIdToChart = new Map<string, Chart>();
  for (const rel of docRels) {
    if (rel.type === RelType.Chart) {
      const chartPath = resolvePartPath(documentPartPath, rel.target);
      consumedPaths.add(chartPath);
      const chartXml = getText(chartPath);
      if (chartXml) {
        const chart = tryParse(() => parseChartXml(chartXml));
        if (chart) {
          chartRIdToChart.set(rel.id, chart);
        }
      }
    }
  }

  // Replace OpaqueDrawing items that reference chart rIds with proper ChartContent
  if (chartRIdToChart.size > 0) {
    replaceOpaqueCharts(body, chartRIdToChart);
  }

  // Parse ChartEx parts and replace opaque drawings with typed ChartExContent
  const chartExRIdToContent = new Map<string, ChartExContent>();
  for (const rel of docRels) {
    if (rel.type === RelType.ChartEx) {
      const chartExPath = resolvePartPath(documentPartPath, rel.target);
      consumedPaths.add(chartExPath);
      const chartExXml = getText(chartExPath);
      if (chartExXml) {
        const data = tryParse(() => parseChartExXml(chartExXml));
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
    replaceOpaqueChartExDrawings(body, chartExRIdToContent);
  }

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

  // Glossary document (Building Blocks). Carried verbatim so a read→write
  // round-trip re-emits it (and re-registers the glossaryDocument relationship
  // + content type) rather than dropping it into opaqueParts where the
  // relationship would be lost. The structured `blocks` are not reverse-parsed.
  let glossary: GlossaryDocument | undefined;
  for (const rel of docRels) {
    if (rel.type !== RelType.Glossary) {
      continue;
    }
    const glossaryPath = resolvePartPath(documentPartPath, rel.target);
    const glossaryData = entries.get(glossaryPath);
    if (glossaryData) {
      consumedPaths.add(glossaryPath);
      glossary = { blocks: [], rawXml: decoder.decode(glossaryData) };
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
        body.splice(i, 1);
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
      const rels = parseRelationships(decoder.decode(partRelsData));
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
    opaqueParts: opaqueParts.length > 0 ? opaqueParts : undefined,
    vbaProject,
    oleObjects,
    glossary
  };
}
