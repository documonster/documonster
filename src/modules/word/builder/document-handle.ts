/**
 * Document Handle & Namespace for building DOCX documents.
 *
 * Provides the Document namespace with free functions for constructing
 * documents via an opaque DocumentHandle.
 */

import type {
  DocxDocument,
  BodyContent,
  Paragraph,
  ParagraphProperties,
  ParagraphChild,
  RunProperties,
  Table,
  SectionProperties,
  StyleDef,
  DocDefaults,
  AbstractNumbering,
  NumberingInstance,
  NumPicBullet,
  HeaderDef,
  FooterDef,
  HeaderFooterContent,
  FootnoteDef,
  EndnoteDef,
  ImageDef,
  ImageMediaType,
  CoreProperties,
  AppProperties,
  DocumentSettings,
  DocumentTheme,
  FontDef,
  EmbeddedFont,
  TableWidth,
  Emu,
  Twips,
  CommentDef,
  CustomXmlPart,
  DocumentBackground,
  CustomPropertyValue,
  OpaquePart,
  PersonInfo,
  TableOfContents,
  MathContent,
  FloatingImage,
  Watermark,
  WebSettings
} from "../types";
import { paragraph, textParagraph, heading } from "./paragraph-builders";
import { pageBreak, floatingImage, mathBlock } from "./run-builders";
import { gridBorders, simpleTable } from "./table-builders";

// =============================================================================
// Document Handle & Namespace
// =============================================================================

/**
 * Internal state for a document being built.
 * Consumers receive an opaque `DocumentHandle` — they cannot construct it directly.
 */
interface _DocumentState {
  body: BodyContent[];
  sectionProperties?: SectionProperties;
  styles: StyleDef[];
  docDefaults?: DocDefaults;
  abstractNumberings: AbstractNumbering[];
  numberingInstances: NumberingInstance[];
  numPicBullets?: NumPicBullet[];
  headers: Map<string, HeaderDef>;
  footers: Map<string, FooterDef>;
  footnotes: FootnoteDef[];
  endnotes: EndnoteDef[];
  images: ImageDef[];
  fonts: FontDef[];
  embeddedFonts?: EmbeddedFont[];
  settings?: DocumentSettings;
  coreProperties?: CoreProperties;
  appProperties?: AppProperties;
  comments: CommentDef[];
  background?: DocumentBackground;
  customProperties: Array<{ name: string; value: CustomPropertyValue }>;
  watermark?: Watermark;
  // --- Round-trip preservation surface ---
  // These fields aren't populated by any builder helper, but they MUST
  // survive `Document.create()` → mutation by the caller →
  // `Document.build()` so a `readDocx → mutate → packageDocx` workflow
  // does not silently drop large chunks of the original package.
  theme?: DocumentTheme;
  webSettings?: WebSettings;
  people?: PersonInfo[];
  customXmlParts?: CustomXmlPart[];
  thumbnail?: DocxDocument["thumbnail"];
  opaqueParts?: OpaquePart[];
  vbaProject?: Uint8Array;
  docType?: DocxDocument["docType"];
  nextImageId: number;
  nextFootnoteId: number;
  nextEndnoteId: number;
  nextBookmarkId: number;
  nextAbstractNumId: number;
  nextNumId: number;
  nextDrawingId: number;
  nextCommentId: number;
}

declare const _documentBrand: unique symbol;

/**
 * Opaque handle representing a document being built.
 * Created via `Document.create()`, passed to `Document.*` functions.
 */
export type DocumentHandle = { readonly [_documentBrand]: true };

/** Cast internal state to opaque handle. */
function _toHandle(state: _DocumentState): DocumentHandle {
  return state as unknown as DocumentHandle;
}

/** Cast opaque handle back to internal state. */
function _toState(handle: DocumentHandle): _DocumentState {
  return handle as unknown as _DocumentState;
}

/**
 * Ensure a numbering instance exists for the given abstractNumId and return
 * its numId. If multiple instances reference the same abstract, the first
 * one is returned. Used by addBulletList / addNumberedList to look up (or
 * create) the runtime numId after they ensure the abstract definition
 * exists. Robust against state where an abstract has no instance yet —
 * the previous code crashed with `find(...)!.numId` in that case.
 */
function _ensureNumberingInstance(s: _DocumentState, abstractNumId: number): number {
  const existing = s.numberingInstances.find(n => n.abstractNumId === abstractNumId);
  if (existing) {
    return existing.numId;
  }
  const numId = s.nextNumId++;
  s.numberingInstances.push({ numId, abstractNumId });
  return numId;
}

/**
 * Namespace of free functions for building DOCX documents.
 *
 * Replaces the former `DocumentBuilder` class with tree-shakeable free functions.
 * Each function operates on an opaque `DocumentHandle`.
 *
 * @example
 * ```ts
 * import { Document, toBuffer } from "excelts/word";
 *
 * const doc = Document.create();
 * Document.addHeading(doc, "Hello World", 1);
 * Document.addParagraph(doc, "This is a paragraph.");
 * Document.addTable(doc, [["Name", "Age"], ["Alice", "30"]]);
 * const bytes = await toBuffer(Document.build(doc));
 * ```
 */
export const Document = {
  /** Create a new document handle. */
  create(): DocumentHandle {
    return _toHandle({
      body: [],
      styles: [],
      abstractNumberings: [],
      numberingInstances: [],
      headers: new Map(),
      footers: new Map(),
      footnotes: [],
      endnotes: [],
      images: [],
      fonts: [],
      comments: [],
      customProperties: [],
      nextImageId: 1,
      nextFootnoteId: 1,
      nextEndnoteId: 1,
      nextBookmarkId: 0,
      nextAbstractNumId: 0,
      nextNumId: 1,
      nextDrawingId: 1,
      nextCommentId: 0
    });
  },

  /** Add raw body content. */
  addContent(doc: DocumentHandle, content: BodyContent): void {
    _toState(doc).body.push(content);
  },

  /** Insert body content at a specific index. */
  insertContentAt(doc: DocumentHandle, index: number, content: BodyContent): void {
    _toState(doc).body.splice(index, 0, content);
  },

  /** Remove body content at a specific index. Returns the removed item. */
  removeContent(doc: DocumentHandle, index: number): BodyContent | undefined {
    const body = _toState(doc).body;
    if (index < 0 || index >= body.length) {
      return undefined;
    }
    return body.splice(index, 1)[0];
  },

  /** Get the number of body content items. */
  getContentCount(doc: DocumentHandle): number {
    return _toState(doc).body.length;
  },

  /** Get body content at a specific index. */
  getContent(doc: DocumentHandle, index: number): BodyContent | undefined {
    return _toState(doc).body[index];
  },

  /** Add a paragraph with runs. */
  addParagraphElement(doc: DocumentHandle, para: Paragraph): void {
    _toState(doc).body.push(para);
  },

  /** Add a simple text paragraph. */
  addParagraph(
    doc: DocumentHandle,
    content: string,
    properties?: ParagraphProperties & { run?: RunProperties }
  ): void {
    _toState(doc).body.push(textParagraph(content, properties));
  },

  /** Add a heading. Accepts plain text or an array of ParagraphChild for mixed formatting. */
  addHeading(
    doc: DocumentHandle,
    content: string | ParagraphChild[],
    level: 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 = 1
  ): void {
    _toState(doc).body.push(heading(content, level));
  },

  /** Add a page break. */
  addPageBreak(doc: DocumentHandle): void {
    _toState(doc).body.push(paragraph([pageBreak()]));
  },

  /** Add a table from a 2D array. */
  addTable(
    doc: DocumentHandle,
    data: string[][],
    options?: { headerRow?: boolean; borders?: boolean; width?: TableWidth; columnWidths?: Twips[] }
  ): void {
    _toState(doc).body.push(simpleTable(data, options));
  },

  /** Add a table element. */
  addTableElement(doc: DocumentHandle, tbl: Table): void {
    _toState(doc).body.push(tbl);
  },

  /** Add an inline image. Returns the image relationship ID and drawing ID. */
  addImage(
    doc: DocumentHandle,
    data: Uint8Array,
    mediaType: ImageMediaType,
    width: Emu,
    height: Emu,
    options?: { altText?: string; name?: string }
  ): { rId: string; drawingId: number } {
    const s = _toState(doc);
    const fileName = `image${s.nextImageId}.${mediaType}`;
    const rId = `__img_${s.nextImageId}`;
    const drawingId = s.nextDrawingId++;

    s.images.push({ data, mediaType, fileName, rId });

    s.body.push(
      paragraph([
        {
          content: [
            {
              type: "image",
              rId,
              width,
              height,
              altText: options?.altText,
              name: options?.name ?? `Picture ${s.nextImageId}`,
              drawingId
            }
          ]
        }
      ])
    );

    s.nextImageId++;
    return { rId, drawingId };
  },

  /** Add a floating image. Returns the image relationship ID. */
  addFloatingImage(
    doc: DocumentHandle,
    data: Uint8Array,
    mediaType: ImageMediaType,
    width: Emu,
    height: Emu,
    options?: {
      altText?: string;
      name?: string;
      horizontalPosition?: FloatingImage["horizontalPosition"];
      verticalPosition?: FloatingImage["verticalPosition"];
      wrap?: FloatingImage["wrap"];
      behindDoc?: boolean;
      lockAnchor?: boolean;
      layoutInCell?: boolean;
      allowOverlap?: boolean;
      distT?: Emu;
      distB?: Emu;
      distL?: Emu;
      distR?: Emu;
      rotation?: number;
      flipHorizontal?: boolean;
      flipVertical?: boolean;
    }
  ): string {
    const s = _toState(doc);
    const fileName = `image${s.nextImageId}.${mediaType}`;
    const rId = `__img_${s.nextImageId}`;

    s.images.push({ data, mediaType, fileName, rId });

    s.body.push(
      floatingImage({
        rId,
        width,
        height,
        altText: options?.altText,
        name: options?.name ?? `Picture ${s.nextImageId}`,
        horizontalPosition: options?.horizontalPosition,
        verticalPosition: options?.verticalPosition,
        wrap: options?.wrap,
        behindDoc: options?.behindDoc,
        lockAnchor: options?.lockAnchor,
        layoutInCell: options?.layoutInCell,
        allowOverlap: options?.allowOverlap,
        distT: options?.distT,
        distB: options?.distB,
        distL: options?.distL,
        distR: options?.distR,
        rotation: options?.rotation,
        flipHorizontal: options?.flipHorizontal,
        flipVertical: options?.flipVertical
      })
    );

    s.nextImageId++;
    return rId;
  },

  /** Add a custom font definition. */
  addFont(doc: DocumentHandle, font: FontDef): void {
    _toState(doc).fonts.push(font);
  },

  /** Set a text watermark on the document. */
  setWatermark(doc: DocumentHandle, watermark: Watermark): void {
    _toState(doc).watermark = watermark;
  },

  /** Add a footnote. Returns the footnote ID. */
  addFootnote(doc: DocumentHandle, content: string | Paragraph[]): number {
    const s = _toState(doc);
    const id = s.nextFootnoteId++;
    const paras = typeof content === "string" ? [textParagraph(content)] : content;
    s.footnotes.push({ id, content: paras });
    return id;
  },

  /** Add an endnote. Returns the endnote ID. */
  addEndnote(doc: DocumentHandle, content: string | Paragraph[]): number {
    const s = _toState(doc);
    const id = s.nextEndnoteId++;
    const paras = typeof content === "string" ? [textParagraph(content)] : content;
    s.endnotes.push({ id, content: paras });
    return id;
  },

  /** Add a comment. Returns the comment ID. */
  addComment(
    doc: DocumentHandle,
    author: string,
    content: string | Paragraph[],
    options?: { date?: string; initials?: string }
  ): number {
    const s = _toState(doc);
    const id = s.nextCommentId++;
    const paras = typeof content === "string" ? [textParagraph(content)] : content;
    s.comments.push({
      id,
      author,
      date: options?.date,
      initials: options?.initials,
      content: paras
    });
    return id;
  },

  /** Add a Table of Contents. */
  addTableOfContents(doc: DocumentHandle, options?: Partial<Omit<TableOfContents, "type">>): void {
    _toState(doc).body.push({
      type: "tableOfContents",
      headingStyleRange: options?.headingStyleRange ?? "1-3",
      hyperlink: options?.hyperlink ?? true,
      ...options
    });
  },

  /** Add a math equation block. */
  addMath(doc: DocumentHandle, content: MathContent[]): void {
    _toState(doc).body.push(mathBlock(content));
  },

  /** Add a text box. */
  addTextBox(
    doc: DocumentHandle,
    content: string | Paragraph[],
    options?: { width?: Twips; height?: Twips; stroke?: boolean; fill?: boolean }
  ): void {
    const paras = typeof content === "string" ? [textParagraph(content)] : content;
    _toState(doc).body.push({
      type: "textBox",
      content: paras,
      width: options?.width,
      height: options?.height,
      stroke: options?.stroke,
      fill: options?.fill
    });
  },

  /**
   * Add a bullet list. Items can be plain strings or arrays of ParagraphChild for rich formatting.
   */
  addBulletList(doc: DocumentHandle, items: (string | ParagraphChild[])[], level = 0): void {
    const s = _toState(doc);
    // Create abstract numbering for bullets if not exists
    let bulletAbsId = s.abstractNumberings.find(
      a => a.levels[0]?.format === "bullet"
    )?.abstractNumId;

    if (bulletAbsId === undefined) {
      bulletAbsId = s.nextAbstractNumId++;
      s.abstractNumberings.push({
        abstractNumId: bulletAbsId,
        multiLevelType: "hybridMultilevel",
        levels: [
          {
            level: 0,
            start: 1,
            format: "bullet",
            text: "\uF0B7",
            justification: "left",
            paragraphProperties: { indent: { left: 720, hanging: 360 } },
            runProperties: { font: { ascii: "Symbol", hAnsi: "Symbol" } }
          },
          {
            level: 1,
            start: 1,
            format: "bullet",
            text: "o",
            justification: "left",
            paragraphProperties: { indent: { left: 1440, hanging: 360 } },
            runProperties: { font: { ascii: "Courier New", hAnsi: "Courier New" } }
          },
          {
            level: 2,
            start: 1,
            format: "bullet",
            text: "\uF0A7",
            justification: "left",
            paragraphProperties: { indent: { left: 2160, hanging: 360 } },
            runProperties: { font: { ascii: "Wingdings", hAnsi: "Wingdings" } }
          }
        ]
      });
    }

    const numId = _ensureNumberingInstance(s, bulletAbsId);

    for (const item of items) {
      if (typeof item === "string") {
        s.body.push(textParagraph(item, { numbering: { numId, level } }));
      } else {
        s.body.push(paragraph(item, { numbering: { numId, level } }));
      }
    }
  },

  /** Add a numbered list. Items can be plain strings or arrays of ParagraphChild for rich formatting. */
  addNumberedList(doc: DocumentHandle, items: (string | ParagraphChild[])[], level = 0): void {
    const s = _toState(doc);
    let numAbsId = s.abstractNumberings.find(a => a.levels[0]?.format === "decimal")?.abstractNumId;

    if (numAbsId === undefined) {
      numAbsId = s.nextAbstractNumId++;
      s.abstractNumberings.push({
        abstractNumId: numAbsId,
        multiLevelType: "hybridMultilevel",
        levels: [
          {
            level: 0,
            start: 1,
            format: "decimal",
            text: "%1.",
            justification: "left",
            paragraphProperties: { indent: { left: 720, hanging: 360 } }
          },
          {
            level: 1,
            start: 1,
            format: "lowerLetter",
            text: "%2.",
            justification: "left",
            paragraphProperties: { indent: { left: 1440, hanging: 360 } }
          },
          {
            level: 2,
            start: 1,
            format: "lowerRoman",
            text: "%3.",
            justification: "right",
            paragraphProperties: { indent: { left: 2160, hanging: 180 } }
          }
        ]
      });
    }

    const numId = _ensureNumberingInstance(s, numAbsId);

    for (const item of items) {
      if (typeof item === "string") {
        s.body.push(textParagraph(item, { numbering: { numId, level } }));
      } else {
        s.body.push(paragraph(item, { numbering: { numId, level } }));
      }
    }
  },

  /** Set section properties (page size, margins, etc.). */
  setSectionProperties(doc: DocumentHandle, props: SectionProperties): void {
    _toState(doc).sectionProperties = props;
  },

  /** Set document defaults. */
  setDocDefaults(doc: DocumentHandle, defaults: DocDefaults): void {
    _toState(doc).docDefaults = defaults;
  },

  /** Add a style definition. */
  addStyle(doc: DocumentHandle, style: StyleDef): void {
    _toState(doc).styles.push(style);
  },

  /** Set default styles (Normal, Heading1-6, Hyperlink, etc.). */
  useDefaultStyles(doc: DocumentHandle): void {
    const s = _toState(doc);
    s.docDefaults = {
      runProperties: {
        font: { ascii: "Calibri", hAnsi: "Calibri", eastAsia: "SimSun", cs: "Times New Roman" },
        size: 22,
        sizeCs: 22,
        language: { val: "en-US" }
      },
      paragraphProperties: {
        spacing: { after: 160, line: 259, lineRule: "auto" }
      }
    };

    s.styles.push(
      { type: "paragraph", styleId: "Normal", name: "Normal", isDefault: true, qFormat: true },
      {
        type: "paragraph",
        styleId: "Heading1",
        name: "heading 1",
        basedOn: "Normal",
        next: "Normal",
        qFormat: true,
        uiPriority: 9,
        paragraphProperties: {
          keepNext: true,
          keepLines: true,
          spacing: { before: 240, after: 0 }
        },
        runProperties: { font: "Calibri Light", color: "2F5496", size: 32 }
      },
      {
        type: "paragraph",
        styleId: "Heading2",
        name: "heading 2",
        basedOn: "Normal",
        next: "Normal",
        qFormat: true,
        uiPriority: 9,
        paragraphProperties: { keepNext: true, keepLines: true, spacing: { before: 40, after: 0 } },
        runProperties: { font: "Calibri Light", color: "2F5496", size: 26 }
      },
      {
        type: "paragraph",
        styleId: "Heading3",
        name: "heading 3",
        basedOn: "Normal",
        next: "Normal",
        qFormat: true,
        uiPriority: 9,
        paragraphProperties: { keepNext: true, keepLines: true, spacing: { before: 40, after: 0 } },
        runProperties: { font: "Calibri Light", color: "1F3763", size: 24 }
      },
      {
        type: "character",
        styleId: "Hyperlink",
        name: "Hyperlink",
        uiPriority: 99,
        runProperties: { color: "0563C1", underline: "single" }
      },
      {
        type: "table",
        styleId: "TableGrid",
        name: "Table Grid",
        basedOn: "TableNormal",
        uiPriority: 39,
        tableProperties: { borders: gridBorders(4, "auto") }
      }
    );
  },

  /** Set a header for the given type. */
  setHeader(doc: DocumentHandle, type: string, content: HeaderFooterContent): void {
    _toState(doc).headers.set(type, { content });
  },

  /** Set a footer for the given type. */
  setFooter(doc: DocumentHandle, type: string, content: HeaderFooterContent): void {
    _toState(doc).footers.set(type, { content });
  },

  /** Set document settings. */
  setSettings(doc: DocumentHandle, settings: DocumentSettings): void {
    _toState(doc).settings = settings;
  },

  /** Set core properties (metadata). */
  setCoreProperties(doc: DocumentHandle, props: CoreProperties): void {
    _toState(doc).coreProperties = props;
  },

  /** Set application properties. */
  setAppProperties(doc: DocumentHandle, props: AppProperties): void {
    _toState(doc).appProperties = props;
  },

  /** Set document background. */
  setBackground(doc: DocumentHandle, background: DocumentBackground): void {
    _toState(doc).background = background;
  },

  /** Add a custom document property. */
  addCustomProperty(doc: DocumentHandle, name: string, value: CustomPropertyValue): void {
    _toState(doc).customProperties.push({ name, value });
  },

  /** Add a section break with properties. */
  addSectionBreak(doc: DocumentHandle, props: SectionProperties): void {
    const s = _toState(doc);
    // Insert as the last paragraph's section properties
    if (s.body.length > 0) {
      const last = s.body[s.body.length - 1];
      if (last.type === "paragraph") {
        const existingProps = last.properties ?? {};
        s.body[s.body.length - 1] = {
          ...last,
          properties: { ...existingProps, sectionProperties: props }
        };
        return;
      }
    }
    // If no previous paragraph, add an empty one with section properties
    s.body.push(paragraph([], { sectionProperties: props }));
  },

  /** Get next available bookmark ID. */
  nextBookmarkId(doc: DocumentHandle): number {
    return _toState(doc).nextBookmarkId++;
  },

  /** Build the DocxDocument model from the handle. */
  build(doc: DocumentHandle): DocxDocument {
    const s = _toState(doc);
    return {
      docType: s.docType,
      body: s.body,
      sectionProperties: s.sectionProperties ?? {
        pageSize: { width: 12240, height: 15840 },
        margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      },
      styles: s.styles.length > 0 ? s.styles : undefined,
      docDefaults: s.docDefaults,
      abstractNumberings: s.abstractNumberings.length > 0 ? s.abstractNumberings : undefined,
      numberingInstances: s.numberingInstances.length > 0 ? s.numberingInstances : undefined,
      numPicBullets: s.numPicBullets && s.numPicBullets.length > 0 ? s.numPicBullets : undefined,
      headers: s.headers.size > 0 ? s.headers : undefined,
      footers: s.footers.size > 0 ? s.footers : undefined,
      footnotes: s.footnotes.length > 0 ? s.footnotes : undefined,
      endnotes: s.endnotes.length > 0 ? s.endnotes : undefined,
      images: s.images.length > 0 ? s.images : undefined,
      fonts: s.fonts.length > 0 ? s.fonts : undefined,
      embeddedFonts: s.embeddedFonts && s.embeddedFonts.length > 0 ? s.embeddedFonts : undefined,
      customXmlParts:
        s.customXmlParts && s.customXmlParts.length > 0 ? s.customXmlParts : undefined,
      settings: s.settings,
      coreProperties: s.coreProperties,
      appProperties: s.appProperties,
      comments: s.comments.length > 0 ? s.comments : undefined,
      background: s.background,
      customProperties: s.customProperties.length > 0 ? s.customProperties : undefined,
      watermark: s.watermark,
      // Round-trip preservation surface — passes through any field the
      // caller stored on the handle, so `readDocx → mutate → packageDocx`
      // does not silently drop these. Builder helpers don't populate
      // them; users who need to manipulate them assign directly.
      theme: s.theme,
      webSettings: s.webSettings,
      people: s.people && s.people.length > 0 ? s.people : undefined,
      thumbnail: s.thumbnail,
      opaqueParts: s.opaqueParts && s.opaqueParts.length > 0 ? s.opaqueParts : undefined,
      vbaProject: s.vbaProject
    };
  }
};
