/**
 * DOCX Module - Constants
 *
 * XML namespaces, relationship types, content types, and other constants
 * for the OOXML WordprocessingML format.
 */

// =============================================================================
// XML Namespaces
// =============================================================================

/** WordprocessingML main namespace (w:). */
export const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** Relationships namespace (r:). */
export const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** DrawingML Word Processing Drawing namespace (wp:). */
export const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

/** DrawingML main namespace (a:). */
export const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";

/** DrawingML picture namespace (pic:). */
export const NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";

/** Markup compatibility namespace (mc:). */
export const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/** VML namespace (v:). */
export const NS_V = "urn:schemas-microsoft-com:vml";

/** Office VML extensions (o:). */
export const NS_O = "urn:schemas-microsoft-com:office:office";

/** Math namespace (m:). */
export const NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

/** Word 2010 extensions (w14:). */
export const NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

/** Word 2013 extensions (w15:). */
export const NS_W15 = "http://schemas.microsoft.com/office/word/2012/wordml";

/** Word Processing Group (wpg:). */
export const NS_WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";

/** Word Processing Ink (wpi:). */
export const NS_WPI = "http://schemas.microsoft.com/office/word/2010/wordprocessingInk";

/** Word Processing Shape (wps:). */
export const NS_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

/** Word Processing Canvas (wpc:). */
export const NS_WPC = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";

/** Word Processing Drawing 2010 (wp14:). */
export const NS_WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";

/** Office Word (w10:). */
export const NS_W10 = "urn:schemas-microsoft-com:office:word";

/** SVG extension (asvg:). */
export const NS_ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

/** SVG extension GUID for a:ext. */
export const GUID_SVG = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

/** Dublin Core elements (dc:). */
export const NS_DC = "http://purl.org/dc/elements/1.1/";

/** Dublin Core terms (dcterms:). */
export const NS_DCTERMS = "http://purl.org/dc/terms/";

/** Dublin Core DCMI type (dcmitype:). */
export const NS_DCMITYPE = "http://purl.org/dc/dcmitype/";

/** Core properties namespace (cp:). */
export const NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

/** Extended properties namespace. */
export const NS_EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

/** Doc props VTypes namespace (vt:). */
export const NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

/** Custom properties namespace. */
export const NS_CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

/** XML Schema instance namespace (xsi:). */
export const NS_XSI = "http://www.w3.org/2001/XMLSchema-instance";

/** Content Types namespace. */
export const NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";

/** Package relationships namespace. */
export const NS_PKG_RELS = "http://schemas.openxmlformats.org/package/2006/relationships";

/** DrawingML picture URI for graphicData. */
export const URI_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";

// =============================================================================
// Namespace Attributes (for document.xml root)
// =============================================================================

/** Standard namespace attributes for w:document root element. */
export const DOCUMENT_NAMESPACES: Record<string, string> = {
  "xmlns:wpc": NS_WPC,
  "xmlns:mc": NS_MC,
  "xmlns:o": NS_O,
  "xmlns:r": NS_R,
  "xmlns:m": NS_M,
  "xmlns:v": NS_V,
  "xmlns:wp": NS_WP,
  "xmlns:w10": NS_W10,
  "xmlns:w": NS_W,
  "xmlns:w14": NS_W14,
  "xmlns:w15": NS_W15,
  "xmlns:wpg": NS_WPG,
  "xmlns:wpi": NS_WPI,
  "xmlns:wps": NS_WPS,
  "xmlns:wp14": NS_WP14,
  "mc:Ignorable": "w14 w15 wp14"
};

// =============================================================================
// Relationship Types
// =============================================================================

const REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_BASE = "http://schemas.openxmlformats.org/package/2006/relationships";

/** OOXML relationship type URIs. */
export const RelType = {
  OfficeDocument: `${REL_BASE}/officeDocument`,
  CoreProperties: `${PKG_BASE}/metadata/core-properties`,
  ExtendedProperties: `${REL_BASE}/extended-properties`,
  CustomProperties: `${REL_BASE}/custom-properties`,
  Styles: `${REL_BASE}/styles`,
  Settings: `${REL_BASE}/settings`,
  FontTable: `${REL_BASE}/fontTable`,
  Numbering: `${REL_BASE}/numbering`,
  Footnotes: `${REL_BASE}/footnotes`,
  Endnotes: `${REL_BASE}/endnotes`,
  Header: `${REL_BASE}/header`,
  Footer: `${REL_BASE}/footer`,
  Image: `${REL_BASE}/image`,
  Hyperlink: `${REL_BASE}/hyperlink`,
  Theme: `${REL_BASE}/theme`,
  Comments: `${REL_BASE}/comments`,
  CommentsExtended: `${REL_BASE}/commentsExtended`,
  WebSettings: `${REL_BASE}/webSettings`,
  Glossary: `${REL_BASE}/glossaryDocument`,
  Font: `${REL_BASE}/font`,
  Chart: `${REL_BASE}/chart`,
  Diagram: `${REL_BASE}/diagramData`,
  CustomXml: `${REL_BASE}/customXml`,
  CustomXmlProps: `${REL_BASE}/customXmlProps`,
  VbaProject: `${REL_BASE}/vbaProject`,
  People: `${REL_BASE}/people`,
  DigitalSignature: `${REL_BASE}/digital-signature/signature`
} as const;

// =============================================================================
// Content Types
// =============================================================================

/** OOXML content type strings. */
export const ContentType = {
  Relationships: "application/vnd.openxmlformats-package.relationships+xml",
  Xml: "application/xml",
  Document: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  Styles: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
  Settings: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
  FontTable: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
  Numbering: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
  Footnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
  Endnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
  Header: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
  Footer: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
  Theme: "application/vnd.openxmlformats-officedocument.theme+xml",
  CoreProperties: "application/vnd.openxmlformats-package.core-properties+xml",
  ExtendedProperties: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
  CustomProperties: "application/vnd.openxmlformats-officedocument.custom-properties+xml",
  Comments: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
  CommentsExtended: "application/vnd.ms-word.commentsExtended+xml",
  People: "application/vnd.ms-word.people+xml",
  WebSettings: "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
  Png: "image/png",
  Jpeg: "image/jpeg",
  Gif: "image/gif",
  Bmp: "image/bmp",
  Tiff: "image/tiff",
  Svg: "image/svg+xml",
  Webp: "image/webp",
  Emf: "image/x-emf",
  Wmf: "image/x-wmf",
  ObfuscatedFont: "application/vnd.openxmlformats-officedocument.obfuscatedFont",
  Chart: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
  CustomXml: "application/xml",
  VbaProject: "application/vnd.ms-office.vbaProject"
} as const;

/** Map from image file extension to content type. */
export const IMAGE_CONTENT_TYPES: Record<string, string> = {
  png: ContentType.Png,
  jpeg: ContentType.Jpeg,
  jpg: ContentType.Jpeg,
  gif: ContentType.Gif,
  bmp: ContentType.Bmp,
  tiff: ContentType.Tiff,
  tif: ContentType.Tiff,
  svg: ContentType.Svg,
  webp: ContentType.Webp,
  emf: ContentType.Emf,
  wmf: ContentType.Wmf
};

// =============================================================================
// Standard XML Declaration
// =============================================================================

/** Standard XML declaration attributes. */
export const STD_DOC_ATTRIBUTES = {
  version: "1.0",
  encoding: "UTF-8",
  standalone: "yes"
};

// =============================================================================
// Default Section Properties (Letter size, 1 inch margins)
// =============================================================================

/** Default page size: US Letter (8.5" x 11") in twips. */
export const DEFAULT_PAGE_WIDTH = 12240; // 8.5 inches
export const DEFAULT_PAGE_HEIGHT = 15840; // 11 inches

/** A4 page size in twips. */
export const A4_PAGE_WIDTH = 11906; // 210mm
export const A4_PAGE_HEIGHT = 16838; // 297mm

/** Default margins: 1 inch all around. */
export const DEFAULT_MARGIN = 1440; // 1 inch in twips
export const DEFAULT_HEADER_FOOTER_MARGIN = 720; // 0.5 inch

/** Default column space. */
export const DEFAULT_COLUMN_SPACE = 720;

// =============================================================================
// Part Paths
// =============================================================================

/** Standard DOCX part file paths within the ZIP package. */
export const PartPath = {
  ContentTypes: "[Content_Types].xml",
  PackageRels: "_rels/.rels",
  Document: "word/document.xml",
  DocumentRels: "word/_rels/document.xml.rels",
  Styles: "word/styles.xml",
  Settings: "word/settings.xml",
  FontTable: "word/fontTable.xml",
  Numbering: "word/numbering.xml",
  Footnotes: "word/footnotes.xml",
  Endnotes: "word/endnotes.xml",
  Comments: "word/comments.xml",
  CommentsExtended: "word/commentsExtended.xml",
  People: "word/people.xml",
  WebSettings: "word/webSettings.xml",
  Theme: "word/theme/theme1.xml",
  CoreProps: "docProps/core.xml",
  AppProps: "docProps/app.xml",
  CustomProps: "docProps/custom.xml",
  Thumbnail: "docProps/thumbnail.jpeg",
  header: (n: number) => `word/header${n}.xml`,
  headerRels: (n: number) => `word/_rels/header${n}.xml.rels`,
  footer: (n: number) => `word/footer${n}.xml`,
  footerRels: (n: number) => `word/_rels/footer${n}.xml.rels`,
  media: (name: string) => `word/media/${name}`
} as const;
