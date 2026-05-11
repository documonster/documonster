/**
 * Layout Document Model
 *
 * Structured output of the layout engine. Contains positioned elements
 * that can be consumed by any renderer (SVG, PDF, Canvas, etc.).
 *
 * The layout engine (`layout.ts`) produces this model from a DocxDocument.
 * Renderers consume it to produce visual output.
 */

/** A point in page coordinates (origin: top-left of page content area). */
export interface LayoutPoint {
  readonly x: number; // points
  readonly y: number; // points
}

/** A rectangle in page coordinates. */
export interface LayoutRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/** A positioned text run within a line box. */
export interface PositionedRun {
  readonly text: string;
  readonly x: number; // x offset from line start
  readonly width: number;
  readonly font: string;
  readonly fontSize: number; // points
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: string; // hex RGB
  readonly underline?: boolean;
  readonly strikethrough?: boolean;
}

/** A line box — a single line of text within a paragraph. */
export interface LineBox {
  readonly y: number; // y offset from paragraph start
  readonly height: number;
  readonly baseline: number; // distance from top of line to baseline
  readonly runs: readonly PositionedRun[];
  readonly alignment: "left" | "center" | "right" | "justify";
}

/** A positioned paragraph on a page. */
export interface LayoutParagraph {
  readonly type: "paragraph";
  readonly rect: LayoutRect;
  readonly lines: readonly LineBox[];
  readonly sourceIndex: number; // index in doc.body
}

/** A positioned table on a page. */
export interface LayoutTable {
  readonly type: "table";
  readonly rect: LayoutRect;
  readonly cells: readonly LayoutTableCell[];
  readonly sourceIndex: number;
}

/** A positioned table cell. */
export interface LayoutTableCell {
  readonly rect: LayoutRect;
  readonly row: number;
  readonly col: number;
  readonly content: readonly (LayoutParagraph | LayoutTable)[];
  readonly backgroundColor?: string;
  readonly borders?: {
    readonly top?: { width: number; color: string };
    readonly bottom?: { width: number; color: string };
    readonly left?: { width: number; color: string };
    readonly right?: { width: number; color: string };
  };
}

/** A positioned image on a page. */
export interface LayoutImage {
  readonly type: "image";
  readonly rect: LayoutRect;
  readonly data: Uint8Array;
  readonly mimeType: string;
  readonly altText?: string;
}

/** A positioned floating object on a page. */
export interface LayoutFloat {
  readonly type: "float";
  readonly rect: LayoutRect;
  readonly content: LayoutImage | LayoutParagraph;
  readonly behindText?: boolean;
}

/** Content that appears on a single page. */
export type PageContent = LayoutParagraph | LayoutTable | LayoutImage | LayoutFloat;

/** Page geometry (margins, dimensions). */
export interface PageGeometry {
  readonly width: number; // total page width in points
  readonly height: number; // total page height in points
  readonly marginTop: number;
  readonly marginBottom: number;
  readonly marginLeft: number;
  readonly marginRight: number;
  /** Usable content area. */
  readonly contentWidth: number;
  readonly contentHeight: number;
}

/** A fully laid-out page. */
export interface LayoutPage {
  readonly pageNumber: number;
  readonly geometry: PageGeometry;
  readonly content: readonly PageContent[];
  readonly header?: readonly LayoutParagraph[];
  readonly footer?: readonly LayoutParagraph[];
  readonly footnoteArea?: readonly LayoutParagraph[];
}

/** The complete layout result — a document broken into positioned pages. */
export interface LayoutDocument {
  readonly pages: readonly LayoutPage[];
  readonly totalPages: number;
  /** Bookmark → page number mapping. */
  readonly bookmarkPages: ReadonlyMap<string, number>;
  /** Section boundaries (page indices where sections start). */
  readonly sectionBreaks: readonly number[];
}
