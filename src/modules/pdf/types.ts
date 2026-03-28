/**
 * Type definitions for the PDF module.
 * Covers input data models, export options, page layout, and internal rendering models.
 *
 * The input data models (PdfWorkbook, PdfSheetData, etc.) are fully independent of
 * the Excel module, allowing the PDF engine to be used standalone.
 */

// =============================================================================
// PDF Input Data Model (Excel-independent)
// =============================================================================

/**
 * Cell value type discriminator for the PDF engine.
 */
export const PdfCellType = {
  Empty: 0,
  String: 1,
  Number: 2,
  Boolean: 3,
  Date: 4,
  RichText: 5,
  Error: 6,
  Formula: 7,
  Hyperlink: 8,
  Merge: 9
} as const;

export type PdfCellTypeValue = (typeof PdfCellType)[keyof typeof PdfCellType];

/** Color data used across the PDF input model. */
export interface PdfColorData {
  argb?: string;
  theme?: number;
  tint?: number;
}

/** Font style in the PDF input model. */
export interface PdfFontStyle {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  underline?: boolean | string;
  color?: PdfColorData;
}

/** Fill data in the PDF input model. */
export interface PdfFillData {
  type: "pattern" | "gradient";
  pattern?: string;
  fgColor?: PdfColorData;
  stops?: Array<{ position?: number; color: PdfColorData }>;
}

/** A single border edge in the PDF input model. */
export interface PdfBorderSideData {
  style?: string;
  color?: PdfColorData;
}

/** Border data in the PDF input model. */
export interface PdfBordersData {
  top?: Partial<PdfBorderSideData>;
  right?: Partial<PdfBorderSideData>;
  bottom?: Partial<PdfBorderSideData>;
  left?: Partial<PdfBorderSideData>;
}

/** Alignment data in the PDF input model. */
export interface PdfAlignmentData {
  horizontal?: string;
  vertical?: string;
  wrapText?: boolean;
  indent?: number;
  textRotation?: number;
}

/** Cell style in the PDF input model. */
export interface PdfCellStyle {
  font?: Partial<PdfFontStyle>;
  numFmt?: string | { formatCode: string };
  fill?: PdfFillData;
  border?: Partial<PdfBordersData>;
  alignment?: Partial<PdfAlignmentData>;
}

/** A single run of rich text. */
export interface PdfRichTextRunData {
  text: string;
  font?: Partial<PdfFontStyle>;
}

/** A cell in the PDF input model. */
export interface PdfCellData {
  type: PdfCellTypeValue;
  value: unknown;
  /** Pre-computed display text */
  text: string;
  style?: Partial<PdfCellStyle>;
  hyperlink?: string;
  /** Formula result (for formula cells) */
  result?: unknown;
  /** Column number (1-based) */
  col: number;
}

/** A row in the PDF input model. */
export interface PdfRowData {
  hidden?: boolean;
  height?: number;
  /** Cells keyed by 1-based column number */
  cells: Map<number, PdfCellData>;
}

/** A column in the PDF input model. */
export interface PdfColumnData {
  hidden?: boolean;
  width?: number;
}

/** Page setup configuration. */
export interface PdfPageSetupData {
  orientation?: string;
  paperSize?: number;
  margins?: { left: number; right: number; top: number; bottom: number };
  scale?: number;
  printTitlesRow?: string;
  showGridLines?: boolean;
  printArea?: string;
}

/** An image embedded in a sheet. */
export interface PdfSheetImage {
  data: Uint8Array;
  format: "jpeg" | "png";
  range: {
    tl: {
      col: number;
      row: number;
      nativeCol?: number;
      nativeRow?: number;
      nativeColOff?: number;
      nativeRowOff?: number;
    };
    br?: {
      col: number;
      row: number;
      nativeCol?: number;
      nativeRow?: number;
      nativeColOff?: number;
      nativeRowOff?: number;
    };
    ext?: { width: number; height: number };
  };
}

/** A single sheet in the PDF input model. */
export interface PdfSheetData {
  name: string;
  state?: "visible" | "hidden" | "veryHidden";
  /** Data bounds (1-based) */
  bounds: { top: number; left: number; bottom: number; right: number };
  /** Columns keyed by 1-based column number */
  columns: Map<number, PdfColumnData>;
  /** Rows keyed by 1-based row number */
  rows: Map<number, PdfRowData>;
  /** Merge ranges in "A1:B2" format */
  merges?: string[];
  pageSetup?: PdfPageSetupData;
  /** Row numbers where manual page breaks occur */
  rowBreaks?: number[];
  /** Column numbers where manual page breaks occur */
  colBreaks?: number[];
  /** Embedded images */
  images?: PdfSheetImage[];
}

/**
 * A workbook data structure for PDF generation.
 * This is a plain data object — not tied to the Excel module.
 */
export interface PdfWorkbook {
  title?: string;
  creator?: string;
  subject?: string;
  sheets: PdfSheetData[];
}

// =============================================================================
// Page Size Definitions
// =============================================================================

/**
 * Standard page sizes in PDF points (1 point = 1/72 inch).
 */
export interface PdfPageSize {
  /** Width in points */
  width: number;
  /** Height in points */
  height: number;
}

/**
 * Predefined page size names.
 */
export type PageSizeName = "A3" | "A4" | "A5" | "LETTER" | "LEGAL" | "TABLOID";

/**
 * Predefined page sizes.
 */
export const PageSizes: Record<PageSizeName, PdfPageSize> = {
  A3: { width: 841.89, height: 1190.55 },
  A4: { width: 595.28, height: 841.89 },
  A5: { width: 419.53, height: 595.28 },
  LETTER: { width: 612, height: 792 },
  LEGAL: { width: 612, height: 1008 },
  TABLOID: { width: 792, height: 1224 }
};

// =============================================================================
// PDF Export Options
// =============================================================================

/**
 * Page orientation for PDF export.
 */
export type PdfOrientation = "portrait" | "landscape";

/**
 * Options for controlling PDF export behavior.
 */
export interface PdfExportOptions {
  /**
   * Page size. Can be a predefined name or custom dimensions.
   * @default "A4"
   */
  pageSize?: PageSizeName | PdfPageSize;

  /**
   * Page orientation. If not set, uses the sheet's pageSetup.orientation.
   * @default "portrait"
   */
  orientation?: PdfOrientation;

  /**
   * Page margins in points (1/72 inch).
   * @default { top: 72, right: 72, bottom: 72, left: 72 }
   */
  margins?: Partial<PdfMargins>;

  /**
   * Which sheets to include. Accepts sheet names or 1-based positions.
   * If omitted, all visible sheets are included.
   */
  sheets?: (string | number)[];

  /**
   * Whether to auto-fit column widths to page width.
   * When true, columns are scaled proportionally to fit the page.
   * @default true
   */
  fitToPage?: boolean;

  /**
   * Scale factor (0.1 to 3.0). Applied after fitToPage.
   * @default 1.0
   */
  scale?: number;

  /**
   * Whether to show grid lines on the page.
   * @default false
   */
  showGridLines?: boolean;

  /**
   * Grid line color as an ARGB hex string (e.g. "FFD0D0D0").
   * @default "FFD0D0D0"
   */
  gridLineColor?: string;

  /**
   * Whether to repeat row headers on each page.
   * Can be a number (row count from top) or false to disable.
   * @default false
   */
  repeatRows?: number | false;

  /**
   * Default font family for cells without an explicit font.
   * @default "Helvetica"
   */
  defaultFontFamily?: string;

  /**
   * Default font size in points for cells without an explicit font size.
   * @default 11
   */
  defaultFontSize?: number;

  /**
   * Whether to include sheet names as page headers.
   * @default false
   */
  showSheetNames?: boolean;

  /**
   * Whether to include page numbers in the footer.
   * @default false
   */
  showPageNumbers?: boolean;

  /**
   * PDF document title metadata.
   */
  title?: string;

  /**
   * PDF document author metadata.
   */
  author?: string;

  /**
   * PDF document subject metadata.
   */
  subject?: string;

  /**
   * PDF document creator metadata.
   * @default "excelts"
   */
  creator?: string;

  /**
   * TrueType font file (.ttf) data for Unicode text support.
   * When provided, all text rendering uses this font, enabling
   * CJK (Chinese, Japanese, Korean), Arabic, Hindi, Cyrillic, and other scripts.
   *
   * Pass the raw bytes of a .ttf font file:
   * ```typescript
   * import { readFileSync } from "fs";
   * const font = readFileSync("NotoSansSC-Regular.ttf");
   * excelToPdf(workbook, { font });
   * ```
   */
  font?: Uint8Array;

  /**
   * Encryption options for password-protecting the PDF.
   *
   * @example
   * ```typescript
   * excelToPdf(workbook, {
   *   encryption: {
   *     ownerPassword: "secret",
   *     userPassword: "open",
   *     permissions: { print: true, copy: false }
   *   }
   * });
   * ```
   */
  encryption?: {
    /** User password (required to open). Empty string = no open password. */
    userPassword?: string;
    /** Owner password (grants full access). Required. */
    ownerPassword: string;
    /** Permissions when opened with user password. */
    permissions?: Partial<{
      print: boolean;
      modify: boolean;
      copy: boolean;
      annotate: boolean;
      fillForms: boolean;
      accessibility: boolean;
      assemble: boolean;
      printHighQuality: boolean;
    }>;
  };
}

// =============================================================================
// Internal Layout Models
// =============================================================================

/**
 * Page margins in PDF points.
 */
export interface PdfMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

/**
 * Resolved (normalized) export options with all defaults applied.
 */
export interface ResolvedPdfOptions {
  pageSize: PdfPageSize;
  orientation: PdfOrientation;
  margins: PdfMargins;
  fitToPage: boolean;
  scale: number;
  showGridLines: boolean;
  gridLineColor: PdfColor;
  repeatRows: number | false;
  defaultFontFamily: string;
  defaultFontSize: number;
  showSheetNames: boolean;
  showPageNumbers: boolean;
  title: string;
  author: string;
  subject: string;
  creator: string;
}

// =============================================================================
// Internal Rendering Types
// =============================================================================

/**
 * RGBA color used internally for PDF rendering.
 * Each component is 0-1. Alpha defaults to 1 (fully opaque) if omitted.
 */
export interface PdfColor {
  r: number;
  g: number;
  b: number;
  /** Opacity: 0 = fully transparent, 1 = fully opaque. Default 1. */
  a?: number;
}

/**
 * A rectangular region in PDF coordinate space (origin = bottom-left).
 */
export interface PdfRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * A laid-out cell ready for rendering.
 */
export interface LayoutCell {
  /** Cell text content */
  text: string;
  /** Rectangle in page coordinates (PDF points, origin bottom-left) */
  rect: PdfRect;
  /** Font family resolved for this cell */
  fontFamily: string;
  /** Font size in points */
  fontSize: number;
  /** Whether the font is bold */
  bold: boolean;
  /** Whether the font is italic */
  italic: boolean;
  /** Whether the font has strikethrough */
  strike: boolean;
  /** Underline style */
  underline: boolean;
  /** Text color */
  textColor: PdfColor;
  /** Background fill color (null = transparent) */
  fillColor: PdfColor | null;
  /** Horizontal alignment */
  horizontalAlign: "left" | "center" | "right";
  /** Vertical alignment */
  verticalAlign: "top" | "middle" | "bottom";
  /** Whether text wrapping is enabled */
  wrapText: boolean;
  /** Border definitions for this cell */
  borders: LayoutBorders;
  /** Number of columns this cell spans (for merged cells) */
  colSpan: number;
  /** Number of rows this cell spans (for merged cells) */
  rowSpan: number;
  /** Hyperlink URL (if this cell is a hyperlink) */
  hyperlink: string | null;
  /** Rich text runs (null if plain text) */
  richText: LayoutRichTextRun[] | null;
  /** Indent level (0 = none) */
  indent: number;
  /** Text rotation in degrees (0-90 ccw, 91-180 cw) or "vertical" for stacked */
  textRotation: number | "vertical";
}

/**
 * A single run within a rich text cell.
 */
export interface LayoutRichTextRun {
  text: string;
  fontFamily: string;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  strike: boolean;
  underline: boolean;
  textColor: PdfColor;
}

/**
 * Border definitions for a single cell.
 */
export interface LayoutBorders {
  top: LayoutBorder | null;
  right: LayoutBorder | null;
  bottom: LayoutBorder | null;
  left: LayoutBorder | null;
}

/**
 * A single border edge.
 */
export interface LayoutBorder {
  /** Line width in points */
  width: number;
  /** Border color */
  color: PdfColor;
  /** Dash pattern (empty array = solid) */
  dashPattern: number[];
}

/**
 * A single page of laid-out content.
 */
export interface LayoutPage {
  /** Page number (1-based) */
  pageNumber: number;
  /** Resolved rendering options for the sheet that produced this page */
  options: ResolvedPdfOptions;
  /** Cells to render on this page */
  cells: LayoutCell[];
  /** Page width in points */
  width: number;
  /** Page height in points */
  height: number;
  /** Sheet name for this page */
  sheetName: string;
  /** Sheet column numbers included on this page */
  sheetCols: number[];
  /** Column x-offsets (left edges) relative to page content area */
  columnOffsets: number[];
  /** Column widths in points */
  columnWidths: number[];
  /** Sheet row numbers included on this page */
  sheetRows: number[];
  /** Row y-offsets (top edges) in page coordinates (PDF bottom-left origin) */
  rowYPositions: number[];
  /** Row heights in points */
  rowHeights: number[];
  /** Images to render on this page */
  images: LayoutImage[];
}

/**
 * A positioned image on a PDF page.
 */
export interface LayoutImage {
  /** Image data bytes (JPEG or PNG) */
  data: Uint8Array;
  /** Image format */
  format: "jpeg" | "png";
  /** Rectangle in page coordinates (PDF points, origin bottom-left) */
  rect: PdfRect;
}
