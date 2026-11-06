/**
 * DOCX Module - Document Validation (Schema Validation)
 *
 * Provides structural validation of DocxDocument models against the
 * OOXML WordprocessingML constraints. This validates the document model
 * at a semantic level (not XML schema validation) to ensure the generated
 * DOCX will be well-formed and compliant.
 *
 * Validation rules are based on:
 * - ECMA-376 Part 1, Chapter 17 (WordprocessingML)
 * - Practical interoperability constraints from Word/LibreOffice
 */

import type { SchemaMatches } from "@utils/schema";
import { findUnrecognizedValues } from "@utils/unrecognized-values";
import { checkHandle } from "@word/core/argument-checks";
import { indexImagesByKey } from "@word/core/image-keys";
import { EXTENSION_MEDIA_TYPE_MAP, KNOWN_IMAGE_MEDIA_TYPES } from "@word/core/image-media";
import { checkSchema, w } from "@word/core/option-schema";
import { isRangeMarker } from "@word/core/text-utils";
import { walkBlocks } from "@word/core/walker";
import type {
  DocxDocument,
  BodyContent,
  Paragraph,
  Table,
  TableRow,
  Run,
  RunContent,
  SectionProperties,
  StyleDef,
  AbstractNumbering,
  NumberingInstance,
  HeaderDef,
  FooterDef,
  ImageDef,
  FloatingImage,
  CommentDef,
  FootnoteDef,
  EndnoteDef,
  DocDefaults,
  Hyperlink,
  ParagraphChild,
  FontDef,
  CustomXmlPart
} from "@word/types";

// =============================================================================
// Types
// =============================================================================

/** Validation severity level. */
export type ValidationSeverity = "error" | "warning" | "info";

/** A single validation issue. */
export interface ValidationIssue {
  /** Severity of the issue. */
  readonly severity: ValidationSeverity;
  /** Human-readable description of the issue. */
  readonly message: string;
  /** Path to the problematic element (e.g. "body[3].rows[0].cells[1]"). */
  readonly path: string;
  /** Rule identifier for programmatic handling. */
  readonly rule: string;
}

/** Validation result. */
export interface ValidationResult {
  /** Whether the document is valid (no errors). */
  readonly valid: boolean;
  /** All issues found. */
  readonly issues: readonly ValidationIssue[];
  /** Count of errors. */
  readonly errorCount: number;
  /** Count of warnings. */
  readonly warningCount: number;
}

/** Validation options. */
export interface ValidationOptions {
  /** Maximum severity to report. Default: all levels. */
  readonly maxSeverity?: ValidationSeverity;
  /** Stop after this many errors. Default: unlimited. */
  readonly maxErrors?: number;
  /** Check for compatibility with specific Word version. */
  readonly compatibilityMode?: "word2007" | "word2010" | "word2013" | "word2016" | "word2019";
  /** Enable strict mode (treats warnings as errors). */
  readonly strict?: boolean;
}

const validationOptionsSchema = w.record({
  maxSeverity: w.oneOf(["error", "warning", "info"]).optional(),
  maxErrors: w.limit().optional(),
  compatibilityMode: w
    .oneOf(["word2007", "word2010", "word2013", "word2016", "word2019"])
    .optional(),
  strict: w.boolean().optional()
});
const _validationOptionsMatch: SchemaMatches<typeof validationOptionsSchema, ValidationOptions> =
  true;
void _validationOptionsMatch;

// =============================================================================
// Validation Engine
// =============================================================================

/**
 * Validate a DocxDocument model for structural correctness and OOXML compliance.
 *
 * @param doc - The document model to validate.
 * @param options - Validation options.
 * @returns Validation result with all issues found.
 */
export function validateDocument(
  doc: DocxDocument,
  options: ValidationOptions = {}
): ValidationResult {
  checkHandle(doc, "document", "doc", "Validation.validate");
  checkSchema(validationOptionsSchema, options, "Validation.validate");
  const issues: ValidationIssue[] = [];
  const maxErrors = options.maxErrors ?? Infinity;

  const addIssue = (
    severity: ValidationSeverity,
    message: string,
    path: string,
    rule: string
  ): void => {
    if (options.maxSeverity === "error" && severity !== "error") {
      return;
    }
    if (options.maxSeverity === "warning" && severity === "info") {
      return;
    }
    issues.push({ severity, message, path, rule });
  };

  const errorCount = (): number => issues.filter(i => i.severity === "error").length;

  // --- Document-level checks ---
  validateBody(doc.body, addIssue, errorCount, maxErrors);
  validateStyles(doc.styles, doc.docDefaults, addIssue);
  validateNumberings(doc.abstractNumberings, doc.numberingInstances, addIssue);
  validateHeaders(doc.headers, addIssue);
  validateFooters(doc.footers, addIssue);
  validateImages(doc.images, doc.body, addIssue);
  validateComments(doc.comments, addIssue);
  validateFootnotes(doc.footnotes, addIssue);
  validateEndnotes(doc.endnotes, addIssue);
  validateSectionProperties(doc.sectionProperties, addIssue);
  validateSettings(doc, addIssue);
  validateCrossReferences(doc, addIssue);

  // --- OpenXML compliance checks ---
  validateRelationshipConsistency(doc, addIssue);
  validateContentTypeConsistency(doc, addIssue);
  validateStructuralIntegrity(doc, addIssue);
  validateNamespaceCompliance(doc, addIssue);

  // --- Enhanced validation rules ---
  validateNumberingAbstractConsistency(doc.abstractNumberings, doc.numberingInstances, addIssue);
  validateStyleBasedOnReferences(doc.styles, addIssue);
  validateHeaderFooterReferences(doc.sectionProperties, doc.headers, doc.footers, addIssue);
  validateCommentIdUniqueness(doc.comments, addIssue);
  validateNoteAndCommentReferences(doc, addIssue);
  validateBookmarkNameUniqueness(doc.body, addIssue);
  validateTableCellContent(doc.body, addIssue);
  validateCustomXmlPartItemIds(doc.customXmlParts, addIssue);
  validateImageExtensionMediaType(doc.images, addIssue);
  validateUnrecognizedValues(doc, addIssue);

  const errors = issues.filter(i => i.severity === "error").length;
  const warnings = issues.filter(i => i.severity === "warning").length;

  return {
    valid: options.strict ? errors === 0 && warnings === 0 : errors === 0,
    issues,
    errorCount: errors,
    warningCount: warnings
  };
}

// =============================================================================
// Body Validation
// =============================================================================

// Exhaustive over the model unions: adding a member to `BodyContent` or
// `ParagraphChild` fails to compile until it is listed here.
const BODY_CONTENT_TYPES: Record<BodyContent["type"], true> = {
  paragraph: true,
  table: true,
  floatingImage: true,
  tableOfContents: true,
  math: true,
  textBox: true,
  checkBox: true,
  drawingShape: true,
  opaqueDrawing: true,
  chart: true,
  chartEx: true,
  altChunk: true,
  sdt: true
};

const PARAGRAPH_CHILD_TYPES: Record<Extract<ParagraphChild, { type: string }>["type"], true> = {
  hyperlink: true,
  bookmarkStart: true,
  bookmarkEnd: true,
  commentRangeStart: true,
  commentRangeEnd: true,
  commentReference: true,
  insertedRun: true,
  deletedRun: true,
  movedFromRun: true,
  movedToRun: true,
  moveFromRangeStart: true,
  moveFromRangeEnd: true,
  moveToRangeStart: true,
  moveToRangeEnd: true,
  customXmlInsRangeStart: true,
  customXmlInsRangeEnd: true,
  customXmlDelRangeStart: true,
  customXmlDelRangeEnd: true,
  customXmlMoveFromRangeStart: true,
  customXmlMoveFromRangeEnd: true,
  customXmlMoveToRangeStart: true,
  customXmlMoveToRangeEnd: true,
  opaqueParagraphChild: true
};

function validateBody(
  body: readonly BodyContent[],
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void,
  errorCount: () => number,
  maxErrors: number
): void {
  if (!body || body.length === 0) {
    addIssue("warning", "Document body is empty", "body", "empty-body");
    return;
  }

  for (let i = 0; i < body.length; i++) {
    if (errorCount() >= maxErrors) {
      return;
    }
    const element = body[i]!;
    const path = `body[${i}]`;

    if (!("type" in element)) {
      addIssue("error", "Body element missing 'type' field", path, "missing-type");
      continue;
    }

    switch (element.type) {
      case "paragraph":
        validateParagraph(element as Paragraph, path, addIssue);
        break;
      case "table":
        validateTable(element as Table, path, addIssue);
        break;
      case "floatingImage":
        validateFloatingImage(element as FloatingImage, path, addIssue);
        break;
      default:
        if (Object.hasOwn(BODY_CONTENT_TYPES, element.type)) {
          break;
        }
        addIssue(
          "error",
          `Unknown body content type: "${(element as { type: string }).type}"`,
          path,
          "unknown-body-type"
        );
    }
  }
}

function validateParagraph(
  p: Paragraph,
  path: string,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!p.children || !Array.isArray(p.children)) {
    addIssue("error", "Paragraph missing 'children' array", path, "paragraph-no-children");
    return;
  }

  for (let i = 0; i < p.children.length; i++) {
    const child = p.children[i]!;
    const childPath = `${path}.children[${i}]`;

    if ("type" in child) {
      if (
        !Object.hasOwn(PARAGRAPH_CHILD_TYPES, (child as { type: string }).type) &&
        !isRangeMarker(child)
      ) {
        addIssue(
          "info",
          `Paragraph child has unusual type: "${(child as { type: string }).type}"`,
          childPath,
          "unusual-para-child"
        );
      }
    } else {
      // It should be a Run
      validateRun(child as Run, childPath, addIssue);
    }
  }

  // Validate paragraph properties
  if (p.properties) {
    if (p.properties.numbering) {
      if (typeof p.properties.numbering.numId !== "number") {
        addIssue(
          "warning",
          "Paragraph numbering reference missing numId",
          `${path}.properties.numbering`,
          "numbering-no-numid"
        );
      }
    }
  }
}

function validateRun(
  run: Run,
  path: string,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!run.content || !Array.isArray(run.content)) {
    addIssue("error", "Run missing 'content' array", path, "run-no-content");
    return;
  }

  for (let i = 0; i < run.content.length; i++) {
    const rc = run.content[i]!;
    const rcPath = `${path}.content[${i}]`;

    if (!rc.type) {
      addIssue("error", "RunContent missing 'type' field", rcPath, "rc-no-type");
    }
  }
}

function validateTable(
  t: Table,
  path: string,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!t.rows || !Array.isArray(t.rows)) {
    addIssue("error", "Table missing 'rows' array", path, "table-no-rows");
    return;
  }

  if (t.rows.length === 0) {
    addIssue("warning", "Table has no rows", path, "table-empty");
    return;
  }

  const gridCols = gridColumnCount(t as Table);

  for (let ri = 0; ri < t.rows.length; ri++) {
    const row = t.rows[ri]!;
    const rowPath = `${path}.rows[${ri}]`;

    if (!row.cells || !Array.isArray(row.cells)) {
      addIssue("error", "TableRow missing 'cells' array", rowPath, "row-no-cells");
      continue;
    }

    const colCount = rowGridColumns(row);

    if (gridCols !== undefined && colCount > gridCols) {
      addIssue(
        "warning",
        `Row ${ri} occupies ${colCount} grid columns, expected ${gridCols}`,
        rowPath,
        "table-col-mismatch"
      );
    }

    // Validate cells
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci]!;
      const cellPath = `${rowPath}.cells[${ci}]`;

      if (!cell.content || !Array.isArray(cell.content)) {
        addIssue("error", "TableCell missing 'content' array", cellPath, "cell-no-content");
      } else if (cell.content.length === 0) {
        addIssue(
          "warning",
          "TableCell has no content (Word requires at least one paragraph)",
          cellPath,
          "cell-empty"
        );
      }
    }
  }
}

/**
 * Grid columns a row occupies (ECMA-376 §17.4.14/§17.4.15/§17.4.17):
 * gridBefore + Σ gridSpan + gridAfter. A vMerge continuation cell is still a
 * `w:tc`, so it is counted like any other cell.
 */
function rowGridColumns(row: TableRow): number {
  let count = (row.properties?.gridBefore ?? 0) + (row.properties?.gridAfter ?? 0);
  for (const cell of row.cells) {
    count += cell.properties?.gridSpan ?? 1;
  }
  return count;
}

/**
 * Column count declared by `w:tblGrid`, or `undefined` when the table has none.
 * Only a row that needs more columns than the grid is a mismatch: a row may
 * end short of the grid with its trailing cells omitted (ECMA-376 §17.4.14,
 * §17.4.15 — `w:gridAfter` is optional), and Word opens such a row unchanged.
 */
function gridColumnCount(table: Table): number | undefined {
  const n = table.columnWidths?.length;
  return n ? n : undefined;
}

function validateFloatingImage(
  fi: FloatingImage,
  path: string,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!fi.rId) {
    addIssue(
      "warning",
      "FloatingImage missing rId (will be assigned during packaging)",
      path,
      "fi-no-rid"
    );
  }
  if (!fi.width || fi.width <= 0) {
    addIssue("error", "FloatingImage must have positive width (EMU)", path, "fi-no-width");
  }
  if (!fi.height || fi.height <= 0) {
    addIssue("error", "FloatingImage must have positive height (EMU)", path, "fi-no-height");
  }
}

// =============================================================================
// Style Validation
// =============================================================================

function validateStyles(
  styles: readonly StyleDef[] | undefined,
  docDefaults: DocDefaults | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!styles) {
    return;
  }

  const styleIds = new Set<string>();
  const defaultTypes = new Set<string>();
  for (let i = 0; i < styles.length; i++) {
    const style = styles[i]!;
    const path = `styles[${i}]`;

    if (style.isDefault) {
      // Word applies the last one; the others are dead and the file is ambiguous.
      if (defaultTypes.has(style.type)) {
        addIssue(
          "warning",
          `More than one default ${style.type} style; "${style.styleId}" is the one Word uses`,
          path,
          "style-dup-default"
        );
      }
      defaultTypes.add(style.type);
    }

    if (!style.styleId) {
      addIssue("error", "Style missing 'styleId'", path, "style-no-id");
      continue;
    }

    if (styleIds.has(style.styleId)) {
      addIssue("error", `Duplicate style ID: "${style.styleId}"`, path, "style-dup-id");
    }
    styleIds.add(style.styleId);

    if (!style.name) {
      addIssue("warning", `Style "${style.styleId}" missing display name`, path, "style-no-name");
    }

    if (!style.type) {
      addIssue("error", `Style "${style.styleId}" missing type`, path, "style-no-type");
    }
    // basedOn cross-reference is validated in validateCrossReferences (second pass)
  }
}

// =============================================================================
// Numbering Validation
// =============================================================================

function validateNumberings(
  abstractNums: readonly AbstractNumbering[] | undefined,
  instances: readonly NumberingInstance[] | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!abstractNums && !instances) {
    return;
  }

  const abstractIds = new Set<number>();
  if (abstractNums) {
    for (let i = 0; i < abstractNums.length; i++) {
      const an = abstractNums[i]!;
      if (abstractIds.has(an.abstractNumId)) {
        addIssue(
          "error",
          `Duplicate abstractNumId: ${an.abstractNumId}`,
          `abstractNumberings[${i}]`,
          "num-dup-abstract"
        );
      }
      abstractIds.add(an.abstractNumId);

      if (!an.levels || an.levels.length === 0) {
        addIssue(
          "warning",
          `AbstractNumbering ${an.abstractNumId} has no levels`,
          `abstractNumberings[${i}]`,
          "num-no-levels"
        );
      }
    }
  }

  if (instances) {
    const numIds = new Set<number>();
    for (let i = 0; i < instances.length; i++) {
      const inst = instances[i]!;
      if (numIds.has(inst.numId)) {
        addIssue(
          "error",
          `Duplicate numId: ${inst.numId}`,
          `numberingInstances[${i}]`,
          "num-dup-instance"
        );
      }
      numIds.add(inst.numId);

      if (!abstractIds.has(inst.abstractNumId)) {
        addIssue(
          "error",
          `NumberingInstance ${inst.numId} references non-existent abstractNumId ${inst.abstractNumId}`,
          `numberingInstances[${i}]`,
          "num-missing-abstract"
        );
      }
    }
  }
}

// =============================================================================
// Header / Footer Validation
// =============================================================================

function validateHeaders(
  headers: ReadonlyMap<string, HeaderDef> | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!headers) {
    return;
  }
  for (const [key, header] of headers) {
    if (!header.content || !header.content.children) {
      addIssue(
        "error",
        `Header "${key}" missing content.children`,
        `headers["${key}"]`,
        "header-no-content"
      );
    }
  }
}

function validateFooters(
  footers: ReadonlyMap<string, FooterDef> | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!footers) {
    return;
  }
  for (const [key, footer] of footers) {
    if (!footer.content || !footer.content.children) {
      addIssue(
        "error",
        `Footer "${key}" missing content.children`,
        `footers["${key}"]`,
        "footer-no-content"
      );
    }
  }
}

// =============================================================================
// Image Validation
// =============================================================================

function validateImages(
  images: readonly ImageDef[] | undefined,
  body: readonly BodyContent[],
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!images) {
    return;
  }

  const fileNames = new Set<string>();
  for (let i = 0; i < images.length; i++) {
    const img = images[i]!;
    const path = `images[${i}]`;

    if (!img.data || img.data.length === 0) {
      addIssue("error", "Image has empty data", path, "image-empty-data");
    }

    if (!img.fileName) {
      addIssue("error", "Image missing fileName", path, "image-no-filename");
    } else if (fileNames.has(img.fileName)) {
      addIssue("error", `Duplicate image fileName: "${img.fileName}"`, path, "image-dup-filename");
    } else {
      fileNames.add(img.fileName);
    }

    if (!img.mediaType) {
      addIssue("warning", "Image missing mediaType", path, "image-no-mediatype");
    }
  }
}

// =============================================================================
// Comment / Note Validation
// =============================================================================

function validateComments(
  comments: readonly CommentDef[] | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!comments) {
    return;
  }
  const ids = new Set<number>();
  for (let i = 0; i < comments.length; i++) {
    const c = comments[i]!;
    if (ids.has(c.id)) {
      addIssue("error", `Duplicate comment ID: ${c.id}`, `comments[${i}]`, "comment-dup-id");
    }
    ids.add(c.id);

    if (!c.author) {
      addIssue("warning", `Comment ${c.id} missing author`, `comments[${i}]`, "comment-no-author");
    }
  }
}

function validateFootnotes(
  footnotes: readonly FootnoteDef[] | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!footnotes) {
    return;
  }
  const ids = new Set<number>();
  for (let i = 0; i < footnotes.length; i++) {
    const fn = footnotes[i]!;
    if (ids.has(fn.id)) {
      addIssue("error", `Duplicate footnote ID: ${fn.id}`, `footnotes[${i}]`, "footnote-dup-id");
    }
    ids.add(fn.id);
  }
}

function validateEndnotes(
  endnotes: readonly EndnoteDef[] | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!endnotes) {
    return;
  }
  const ids = new Set<number>();
  for (let i = 0; i < endnotes.length; i++) {
    const en = endnotes[i]!;
    if (ids.has(en.id)) {
      addIssue("error", `Duplicate endnote ID: ${en.id}`, `endnotes[${i}]`, "endnote-dup-id");
    }
    ids.add(en.id);
  }
}

// =============================================================================
// Section Properties Validation
// =============================================================================

function validateSectionProperties(
  sp: SectionProperties | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!sp) {
    return;
  }

  if (sp.pageSize) {
    if (sp.pageSize.width && sp.pageSize.width <= 0) {
      addIssue(
        "error",
        "Page width must be positive",
        "sectionProperties.pageSize",
        "section-bad-width"
      );
    }
    if (sp.pageSize.height && sp.pageSize.height <= 0) {
      addIssue(
        "error",
        "Page height must be positive",
        "sectionProperties.pageSize",
        "section-bad-height"
      );
    }
  }

  if (sp.columns && sp.columns.count !== undefined && sp.columns.count < 1) {
    addIssue(
      "error",
      "Column count must be at least 1",
      "sectionProperties.columns",
      "section-bad-cols"
    );
  }
}

// =============================================================================
// Settings Validation
// =============================================================================

function validateSettings(
  doc: DocxDocument,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (doc.settings?.documentProtection) {
    const dp = doc.settings.documentProtection;
    if (dp.enforcement && !dp.edit) {
      addIssue(
        "warning",
        "Document protection enforced but no edit restriction type specified",
        "settings.documentProtection",
        "protection-no-edit"
      );
    }
  }
}

// =============================================================================
// Cross-Reference Validation
// =============================================================================

function validateCrossReferences(
  doc: DocxDocument,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  // Collect valid numbering IDs
  const validNumIds = new Set<number>();
  if (doc.numberingInstances) {
    for (const inst of doc.numberingInstances) {
      validNumIds.add(inst.numId);
    }
  }

  // Collect valid style IDs
  const validStyleIds = new Set<string>();
  if (doc.styles) {
    for (const style of doc.styles) {
      if (style.styleId) {
        validStyleIds.add(style.styleId);
      }
    }
  }

  // Validate paragraph numbering references
  let paraIdx = 0;
  for (const element of doc.body) {
    if (element.type === "paragraph") {
      const numRef = element.properties?.numbering;
      if (numRef && typeof numRef.numId === "number" && validNumIds.size > 0) {
        if (!validNumIds.has(numRef.numId)) {
          addIssue(
            "warning",
            `Paragraph references numId ${numRef.numId} which does not exist in numberingInstances`,
            `body[${paraIdx}].properties.numbering`,
            "xref-numbering-missing"
          );
        }
      }
      // Validate style reference
      if (element.properties?.style && validStyleIds.size > 0) {
        if (!validStyleIds.has(element.properties.style)) {
          addIssue(
            "info",
            `Paragraph references style "${element.properties.style}" which is not defined`,
            `body[${paraIdx}].properties.style`,
            "xref-style-missing"
          );
        }
      }
    }
    paraIdx++;
  }

  // Validate basedOn style references (second pass — all styles now collected)
  if (doc.styles) {
    for (let i = 0; i < doc.styles.length; i++) {
      const style = doc.styles[i]!;
      if (style.basedOn && !validStyleIds.has(style.basedOn)) {
        addIssue(
          "warning",
          `Style "${style.styleId}" references basedOn style "${style.basedOn}" which does not exist`,
          `styles[${i}]`,
          "xref-basedOn-missing"
        );
      }
    }
  }
}

// =============================================================================
// Relationship Consistency Validation
// =============================================================================

function validateRelationshipConsistency(
  doc: DocxDocument,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  // Every key an image answers to — its rId and aliasRIds — resolved as the writer resolves them
  const validImageRIds = new Set<string>(indexImagesByKey(doc.images).keys());

  // Validate body image references (InlineImageContent and FloatingImage)
  for (let i = 0; i < doc.body.length; i++) {
    const element = doc.body[i]!;
    const path = `body[${i}]`;

    if (element.type === "floatingImage") {
      const fi = element as FloatingImage;
      if (fi.rId && doc.images && doc.images.length > 0 && !validImageRIds.has(fi.rId)) {
        addIssue(
          "error",
          `FloatingImage references rId "${fi.rId}" which does not exist in images`,
          path,
          "rel-image-missing"
        );
      }
    } else if (element.type === "paragraph") {
      const para = element as Paragraph;
      validateParagraphImageRefs(para, path, validImageRIds, doc.images, addIssue);
      validateParagraphHyperlinkRefs(para, path, addIssue);
    }
  }

  // Validate header/footer content image references
  if (doc.headers) {
    for (const [key, header] of doc.headers) {
      if (header.content && header.content.children) {
        validateHeaderFooterContentImageRefs(
          header.content.children,
          `headers["${key}"]`,
          validImageRIds,
          doc.images,
          addIssue
        );
      }
    }
  }

  if (doc.footers) {
    for (const [key, footer] of doc.footers) {
      if (footer.content && footer.content.children) {
        validateHeaderFooterContentImageRefs(
          footer.content.children,
          `footers["${key}"]`,
          validImageRIds,
          doc.images,
          addIssue
        );
      }
    }
  }
}

function validateParagraphImageRefs(
  para: Paragraph,
  path: string,
  validImageRIds: Set<string>,
  images: readonly ImageDef[] | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!para.children) {
    return;
  }
  for (let ci = 0; ci < para.children.length; ci++) {
    const child = para.children[ci]!;
    const childPath = `${path}.children[${ci}]`;

    if ("type" in child && (child as ParagraphChild & { type: string }).type === "hyperlink") {
      const hl = child as Hyperlink;
      validateRunsImageRefs(hl.children, childPath, validImageRIds, images, addIssue);
    } else if (!("type" in child)) {
      // It's a Run
      validateRunImageRefs(child as Run, childPath, validImageRIds, images, addIssue);
    }
  }
}

function validateRunsImageRefs(
  runs: readonly Run[],
  path: string,
  validImageRIds: Set<string>,
  images: readonly ImageDef[] | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  for (let i = 0; i < runs.length; i++) {
    validateRunImageRefs(runs[i]!, `${path}.children[${i}]`, validImageRIds, images, addIssue);
  }
}

function validateRunImageRefs(
  run: Run,
  path: string,
  validImageRIds: Set<string>,
  images: readonly ImageDef[] | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!run.content) {
    return;
  }
  for (let i = 0; i < run.content.length; i++) {
    const rc = run.content[i]!;
    if (rc.type === "image") {
      const img = rc as RunContent & { type: "image"; rId: string };
      if (img.rId && images && images.length > 0 && !validImageRIds.has(img.rId)) {
        addIssue(
          "error",
          `InlineImage references rId "${img.rId}" which does not exist in images`,
          `${path}.content[${i}]`,
          "rel-image-missing"
        );
      }
    }
  }
}

function validateParagraphHyperlinkRefs(
  para: Paragraph,
  path: string,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!para.children) {
    return;
  }
  for (let ci = 0; ci < para.children.length; ci++) {
    const child = para.children[ci]!;
    if ("type" in child && (child as ParagraphChild & { type: string }).type === "hyperlink") {
      const hl = child as Hyperlink;
      if (hl.rId && !hl.url && !hl.anchor) {
        addIssue(
          "warning",
          `Hyperlink has rId "${hl.rId}" but no url or anchor`,
          `${path}.children[${ci}]`,
          "rel-hyperlink-no-url"
        );
      }
    }
  }
}

function validateHeaderFooterContentImageRefs(
  children: readonly (Paragraph | Table)[],
  basePath: string,
  validImageRIds: Set<string>,
  images: readonly ImageDef[] | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  walkBlocks(children as readonly BodyContent[], {
    visitRunContent(content, _run, path) {
      if (content.type === "image") {
        const img = content as RunContent & { type: "image"; rId: string };
        if (img.rId && images && images.length > 0 && !validImageRIds.has(img.rId)) {
          addIssue(
            "error",
            `InlineImage references rId "${img.rId}" which does not exist in images`,
            `${basePath}:depth${path.depth}`,
            "rel-image-missing"
          );
        }
      }
    }
  });
}

// =============================================================================
// Content Type Consistency Validation
// =============================================================================

function validateContentTypeConsistency(
  doc: DocxDocument,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  // Validate image media types
  if (doc.images) {
    for (let i = 0; i < doc.images.length; i++) {
      const img = doc.images[i]!;
      if (img.mediaType && !img.contentType && !KNOWN_IMAGE_MEDIA_TYPES.includes(img.mediaType)) {
        addIssue(
          "error",
          `Image has invalid mediaType "${img.mediaType}". Valid types: ${KNOWN_IMAGE_MEDIA_TYPES.join(", ")}`,
          `images[${i}]`,
          "content-type-invalid-media"
        );
      }
    }
  }

  // Validate embedded fonts: if a font references an embed rId, there should be
  // a corresponding entry in embeddedFonts
  if (doc.fonts) {
    const embeddedRIds = new Set<string>();
    if (doc.embeddedFonts) {
      for (const ef of doc.embeddedFonts) {
        embeddedRIds.add(ef.rId);
      }
    }

    for (let i = 0; i < doc.fonts.length; i++) {
      const font = doc.fonts[i]!;
      const fontPath = `fonts[${i}]`;
      const embedRIds = collectFontEmbedRIds(font);

      for (const rId of embedRIds) {
        if (!embeddedRIds.has(rId)) {
          addIssue(
            "warning",
            `Font "${font.name}" references embed rId "${rId}" but no corresponding embeddedFont exists`,
            fontPath,
            "content-type-font-embed-missing"
          );
        }
      }
    }
  }
}

function collectFontEmbedRIds(font: FontDef): string[] {
  const rIds: string[] = [];
  if (font.embedRegular) {
    rIds.push(font.embedRegular);
  }
  if (font.embedBold) {
    rIds.push(font.embedBold);
  }
  if (font.embedItalic) {
    rIds.push(font.embedItalic);
  }
  if (font.embedBoldItalic) {
    rIds.push(font.embedBoldItalic);
  }
  return rIds;
}

// =============================================================================
// Structural Integrity Validation
// =============================================================================

function validateStructuralIntegrity(
  doc: DocxDocument,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  // Collect valid numbering IDs
  const validNumIds = new Set<number>();
  if (doc.numberingInstances) {
    for (const inst of doc.numberingInstances) {
      validNumIds.add(inst.numId);
    }
  }

  // Collect valid style IDs
  const validStyleIds = new Set<string>();
  if (doc.styles) {
    for (const style of doc.styles) {
      if (style.styleId) {
        validStyleIds.add(style.styleId);
      }
    }
  }

  // Validate body content structural integrity
  for (let i = 0; i < doc.body.length; i++) {
    const element = doc.body[i]!;
    const path = `body[${i}]`;

    if (element.type === "table") {
      validateTableColumnConsistency(element as Table, path, addIssue);
    } else if (element.type === "paragraph") {
      const para = element as Paragraph;

      // Validate numbering references
      if (para.properties?.numbering && typeof para.properties.numbering.numId === "number") {
        if (doc.numberingInstances && !validNumIds.has(para.properties.numbering.numId)) {
          addIssue(
            "error",
            `Paragraph references numId ${para.properties.numbering.numId} which has no corresponding numberingInstance`,
            `${path}.properties.numbering`,
            "struct-numbering-missing"
          );
        }
      }

      // Validate paragraph style references
      if (para.properties?.style && doc.styles) {
        if (!validStyleIds.has(para.properties.style)) {
          addIssue(
            "warning",
            `Paragraph references style "${para.properties.style}" which is not defined in styles`,
            `${path}.properties.style`,
            "struct-style-missing"
          );
        }
      }

      // Validate run style references
      if (para.children) {
        validateRunStyleRefs(para.children, path, validStyleIds, doc.styles, addIssue);
      }
    }
  }
}

function validateTableColumnConsistency(
  table: Table,
  path: string,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!table.rows || table.rows.length === 0) {
    return;
  }

  const gridCols = gridColumnCount(table);
  if (gridCols === undefined) {
    return;
  }

  for (let ri = 0; ri < table.rows.length; ri++) {
    const row = table.rows[ri]!;
    if (!row.cells || !Array.isArray(row.cells)) {
      continue;
    }

    const colCount = rowGridColumns(row);

    if (colCount > gridCols) {
      addIssue(
        "error",
        `Table row ${ri} occupies ${colCount} grid columns, expected ${gridCols}`,
        `${path}.rows[${ri}]`,
        "struct-table-col-mismatch"
      );
    }
  }
}

function validateRunStyleRefs(
  children: readonly ParagraphChild[],
  path: string,
  validStyleIds: Set<string>,
  styles: readonly StyleDef[] | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!styles) {
    return;
  }
  for (let ci = 0; ci < children.length; ci++) {
    const child = children[ci]!;
    const childPath = `${path}.children[${ci}]`;

    if ("type" in child && (child as ParagraphChild & { type: string }).type === "hyperlink") {
      const hl = child as Hyperlink;
      for (let ri = 0; ri < hl.children.length; ri++) {
        const run = hl.children[ri]!;
        if (run.properties?.style && !validStyleIds.has(run.properties.style)) {
          addIssue(
            "warning",
            `Run references style "${run.properties.style}" which is not defined in styles`,
            `${childPath}.children[${ri}].properties.style`,
            "struct-run-style-missing"
          );
        }
      }
    } else if (!("type" in child)) {
      const run = child as Run;
      if (run.properties?.style && !validStyleIds.has(run.properties.style)) {
        addIssue(
          "warning",
          `Run references style "${run.properties.style}" which is not defined in styles`,
          `${childPath}.properties.style`,
          "struct-run-style-missing"
        );
      }
    }
  }
}

// =============================================================================
// Namespace / Standard Compliance Validation
// =============================================================================

/** Maximum reasonable page size in twips (50000 twips ≈ 34.7 inches). */
const MAX_PAGE_SIZE_TWIPS = 50000;

/** Valid header/footer type values per OOXML spec. */
const VALID_HEADER_FOOTER_TYPES: readonly string[] = ["default", "first", "even"];

function validateNamespaceCompliance(
  doc: DocxDocument,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  // Validate page size range
  if (doc.sectionProperties?.pageSize) {
    const ps = doc.sectionProperties.pageSize;
    if (ps.width <= 0) {
      addIssue(
        "error",
        "Page width must be greater than 0",
        "sectionProperties.pageSize.width",
        "ns-page-width-invalid"
      );
    } else if (ps.width >= MAX_PAGE_SIZE_TWIPS) {
      addIssue(
        "warning",
        `Page width ${ps.width} twips exceeds reasonable maximum (${MAX_PAGE_SIZE_TWIPS})`,
        "sectionProperties.pageSize.width",
        "ns-page-width-excessive"
      );
    }

    if (ps.height <= 0) {
      addIssue(
        "error",
        "Page height must be greater than 0",
        "sectionProperties.pageSize.height",
        "ns-page-height-invalid"
      );
    } else if (ps.height >= MAX_PAGE_SIZE_TWIPS) {
      addIssue(
        "warning",
        `Page height ${ps.height} twips exceeds reasonable maximum (${MAX_PAGE_SIZE_TWIPS})`,
        "sectionProperties.pageSize.height",
        "ns-page-height-excessive"
      );
    }
  }

  // Validate header/footer type values in section references
  if (doc.sectionProperties?.headers) {
    for (let i = 0; i < doc.sectionProperties.headers.length; i++) {
      const ref = doc.sectionProperties.headers[i]!;
      if (!VALID_HEADER_FOOTER_TYPES.includes(ref.type)) {
        addIssue(
          "error",
          `Header reference has invalid type "${ref.type}". Valid types: ${VALID_HEADER_FOOTER_TYPES.join(", ")}`,
          `sectionProperties.headers[${i}]`,
          "ns-header-type-invalid"
        );
      }
    }
  }

  if (doc.sectionProperties?.footers) {
    for (let i = 0; i < doc.sectionProperties.footers.length; i++) {
      const ref = doc.sectionProperties.footers[i]!;
      if (!VALID_HEADER_FOOTER_TYPES.includes(ref.type)) {
        addIssue(
          "error",
          `Footer reference has invalid type "${ref.type}". Valid types: ${VALID_HEADER_FOOTER_TYPES.join(", ")}`,
          `sectionProperties.footers[${i}]`,
          "ns-footer-type-invalid"
        );
      }
    }
  }
}

// =============================================================================
// Numbering Instance / Abstract Consistency (Enhanced)
// =============================================================================

function validateNumberingAbstractConsistency(
  abstractNums: readonly AbstractNumbering[] | undefined,
  instances: readonly NumberingInstance[] | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!instances || !abstractNums) {
    return;
  }

  const abstractIds = new Set<number>();
  for (const an of abstractNums) {
    abstractIds.add(an.abstractNumId);
  }

  for (let i = 0; i < instances.length; i++) {
    const inst = instances[i]!;
    if (!abstractIds.has(inst.abstractNumId)) {
      addIssue(
        "error",
        `NumberingInstance numId=${inst.numId} references abstractNumId ${inst.abstractNumId} which does not exist in abstractNumberings`,
        `numberingInstances[${i}]`,
        "num-abstract-ref-missing"
      );
    }
  }
}

// =============================================================================
// Style basedOn Reference Validation
// =============================================================================

function validateStyleBasedOnReferences(
  styles: readonly StyleDef[] | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!styles) {
    return;
  }

  const styleIds = new Set<string>();
  for (const style of styles) {
    if (style.styleId) {
      styleIds.add(style.styleId);
    }
  }

  for (let i = 0; i < styles.length; i++) {
    const style = styles[i]!;
    if (style.basedOn && !styleIds.has(style.basedOn)) {
      addIssue(
        "error",
        `Style "${style.styleId}" has basedOn="${style.basedOn}" which does not exist in defined styles`,
        `styles[${i}].basedOn`,
        "style-basedOn-ref-missing"
      );
    }
  }
}

// =============================================================================
// Header/Footer Reference Validation
// =============================================================================

function validateHeaderFooterReferences(
  sp: SectionProperties | undefined,
  headers: ReadonlyMap<string, HeaderDef> | undefined,
  footers: ReadonlyMap<string, FooterDef> | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!sp) {
    return;
  }

  if (sp.headers && headers) {
    for (let i = 0; i < sp.headers.length; i++) {
      const ref = sp.headers[i]!;
      if (ref.rId && !headers.has(ref.rId)) {
        addIssue(
          "error",
          `Section header reference rId="${ref.rId}" (type="${ref.type}") does not exist in headers map`,
          `sectionProperties.headers[${i}]`,
          "hf-ref-header-missing"
        );
      }
    }
  }

  if (sp.footers && footers) {
    for (let i = 0; i < sp.footers.length; i++) {
      const ref = sp.footers[i]!;
      if (ref.rId && !footers.has(ref.rId)) {
        addIssue(
          "error",
          `Section footer reference rId="${ref.rId}" (type="${ref.type}") does not exist in footers map`,
          `sectionProperties.footers[${i}]`,
          "hf-ref-footer-missing"
        );
      }
    }
  }
}

// =============================================================================
// Comment ID Uniqueness
// =============================================================================

function validateCommentIdUniqueness(
  comments: readonly CommentDef[] | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!comments) {
    return;
  }

  const ids = new Set<number>();
  for (let i = 0; i < comments.length; i++) {
    const c = comments[i]!;
    if (ids.has(c.id)) {
      addIssue(
        "error",
        `Comment ID ${c.id} is not unique — duplicates found`,
        `comments[${i}]`,
        "comment-id-not-unique"
      );
    }
    ids.add(c.id);
  }
}

// =============================================================================
// Reference Closure (footnote/endnote/comment IDs referenced in body)
// =============================================================================

/**
 * Walk every Run.content in the document and ensure that any
 * `footnoteRef` / `endnoteRef` / `commentReference` ids are actually
 * defined in `doc.footnotes` / `doc.endnotes` / `doc.comments`.
 *
 * A dangling reference produces an XML document that opens in Word but
 * shows a broken reference number, which is one of the more confusing
 * failure modes for downstream tooling. We surface it as an error.
 */
function validateNoteAndCommentReferences(
  doc: DocxDocument,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  // Build the id sets once.
  const footnoteIds = new Set<number>();
  if (doc.footnotes) {
    for (const fn of doc.footnotes) {
      footnoteIds.add(fn.id);
    }
  }
  const endnoteIds = new Set<number>();
  if (doc.endnotes) {
    for (const en of doc.endnotes) {
      endnoteIds.add(en.id);
    }
  }
  const commentIds = new Set<number>();
  if (doc.comments) {
    for (const c of doc.comments) {
      commentIds.add(c.id);
    }
  }

  // Visit every run in body / headers / footers / footnotes / endnotes
  // / comments. Footnotes themselves can reference other footnotes in
  // theory but Word does not allow it, so we only scan the obvious
  // surfaces.
  const checkRun = (run: Run, path: string): void => {
    for (const c of run.content) {
      if (c.type === "footnoteRef") {
        if (footnoteIds.size > 0 && !footnoteIds.has(c.id)) {
          addIssue(
            "error",
            `footnoteRef points at id=${c.id} which is not defined in doc.footnotes`,
            path,
            "ref-footnote-missing"
          );
        }
      } else if (c.type === "endnoteRef") {
        if (endnoteIds.size > 0 && !endnoteIds.has(c.id)) {
          addIssue(
            "error",
            `endnoteRef points at id=${c.id} which is not defined in doc.endnotes`,
            path,
            "ref-endnote-missing"
          );
        }
      }
    }
  };

  // CommentReference / commentRangeStart / commentRangeEnd live as
  // ParagraphChild siblings (alongside runs) — scan paragraph children
  // directly. We use the existing walker to keep dispatch consistent.
  // Comment ranges between blocks (not inside a paragraph) are checked the same way.
  let inParagraph = 0;
  const checkBlockCommentRange = (
    marker: { type: "commentRangeStart" | "commentRangeEnd"; id: number },
    path: { index: number }
  ): void => {
    if (inParagraph === 0 && commentIds.size > 0 && !commentIds.has(marker.id)) {
      addIssue(
        "error",
        `${marker.type} points at id=${marker.id} which is not defined in doc.comments`,
        `body[${path.index}]`,
        "ref-comment-range-missing"
      );
    }
  };
  walkBlocks(doc.body, {
    visitCommentRangeStart: checkBlockCommentRange,
    visitCommentRangeEnd: checkBlockCommentRange,
    leaveParagraph() {
      inParagraph--;
    },
    enterParagraph(para, path) {
      inParagraph++;
      const basePath = `body[${path.index}]`;
      for (let i = 0; i < para.children.length; i++) {
        const child = para.children[i]!;
        const childPath = `${basePath}.children[${i}]`;
        if ("type" in child) {
          if (child.type === "commentReference") {
            if (commentIds.size > 0 && !commentIds.has(child.id)) {
              addIssue(
                "error",
                `commentReference points at id=${child.id} which is not defined in doc.comments`,
                childPath,
                "ref-comment-missing"
              );
            }
          } else if (child.type === "commentRangeStart" || child.type === "commentRangeEnd") {
            if (commentIds.size > 0 && !commentIds.has(child.id)) {
              addIssue(
                "error",
                `${child.type} points at id=${child.id} which is not defined in doc.comments`,
                childPath,
                "ref-comment-range-missing"
              );
            }
          } else if (child.type === "hyperlink") {
            // Hyperlink children are runs.
            for (const r of child.children) {
              checkRun(r, `${childPath}.children`);
            }
          } else if (
            child.type === "insertedRun" ||
            child.type === "deletedRun" ||
            child.type === "movedFromRun" ||
            child.type === "movedToRun"
          ) {
            checkRun(child.run, `${childPath}.run`);
          }
        } else {
          // Bare run.
          checkRun(child as Run, childPath);
        }
      }
    }
  });
  for (const marker of doc.bodyRangeMarkersAfter ?? []) {
    if (marker.type === "commentRangeStart" || marker.type === "commentRangeEnd") {
      checkBlockCommentRange(marker, { index: doc.body.length });
    }
  }
}

// =============================================================================
// Bookmark Name Uniqueness
// =============================================================================

function validateBookmarkNameUniqueness(
  body: readonly BodyContent[],
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  const bookmarkNames = new Map<string, string>();

  walkBlocks(body, {
    visitBookmarkStart(bm, path) {
      const pathStr = `body[${path.index}]:depth${path.depth}`;
      if (bookmarkNames.has(bm.name)) {
        addIssue(
          "error",
          `Bookmark name "${bm.name}" is not unique — first defined at ${bookmarkNames.get(bm.name)!}`,
          pathStr,
          "bookmark-name-not-unique"
        );
      } else {
        bookmarkNames.set(bm.name, pathStr);
      }
    }
  });
}

// =============================================================================
// Table Cell Content Validation (OOXML requires at least one paragraph)
// =============================================================================

function validateTableCellContent(
  body: readonly BodyContent[],
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  for (let i = 0; i < body.length; i++) {
    const element = body[i]!;
    if (element.type === "table") {
      validateTableCellParagraphs(element as Table, `body[${i}]`, addIssue);
    }
  }
}

function validateTableCellParagraphs(
  table: Table,
  basePath: string,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!table.rows) {
    return;
  }
  for (let ri = 0; ri < table.rows.length; ri++) {
    const row = table.rows[ri]!;
    if (!row.cells) {
      continue;
    }
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci]!;
      const cellPath = `${basePath}.rows[${ri}].cells[${ci}]`;

      if (!cell.content || cell.content.length === 0) {
        addIssue(
          "error",
          "Table cell must contain at least one paragraph (OOXML requirement)",
          cellPath,
          "cell-no-paragraph"
        );
        continue;
      }

      // Check that at least one paragraph exists in content
      const hasParagraph = cell.content.some(c => c.type === "paragraph");
      if (!hasParagraph) {
        addIssue(
          "error",
          "Table cell must contain at least one paragraph (OOXML requirement)",
          cellPath,
          "cell-no-paragraph"
        );
      }

      // Recursively validate nested tables
      for (let pi = 0; pi < cell.content.length; pi++) {
        const element = cell.content[pi]!;
        if (element.type === "table") {
          validateTableCellParagraphs(element as Table, `${cellPath}.content[${pi}]`, addIssue);
        }
      }
    }
  }
}

// =============================================================================
// Custom XML Part itemId Uniqueness
// =============================================================================

function validateCustomXmlPartItemIds(
  customXmlParts: readonly CustomXmlPart[] | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!customXmlParts) {
    return;
  }

  const itemIds = new Set<string>();
  for (let i = 0; i < customXmlParts.length; i++) {
    const part = customXmlParts[i]!;
    // Normalize itemId by stripping braces and lowercasing for comparison
    const normalizedId = part.itemId.replace(/[{}]/g, "").toLowerCase();
    if (itemIds.has(normalizedId)) {
      addIssue(
        "error",
        `CustomXmlPart itemId "${part.itemId}" is not unique`,
        `customXmlParts[${i}]`,
        "custom-xml-itemid-not-unique"
      );
    }
    itemIds.add(normalizedId);
  }
}

// =============================================================================
// Image File Extension / MediaType Consistency
// =============================================================================

/**
 * Values the reader kept as read because they are outside the vocabulary of
 * their field — see `readToken`. They are written back unchanged, so the
 * document is no worse than its source; but Word may repair them, and one that
 * the schema does allow means a vocabulary here is too narrow.
 */
function validateUnrecognizedValues(
  doc: DocxDocument,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  for (const { path, value } of findUnrecognizedValues(doc)) {
    addIssue(
      "warning",
      `${JSON.stringify(value)} is not a value this field takes; it was kept as read and is written back unchanged`,
      path,
      "unrecognized-value"
    );
  }
}

function validateImageExtensionMediaType(
  images: readonly ImageDef[] | undefined,
  addIssue: (s: ValidationSeverity, m: string, p: string, r: string) => void
): void {
  if (!images) {
    return;
  }

  for (let i = 0; i < images.length; i++) {
    const img = images[i]!;
    // A picture typed by its source package is written with that type.
    if (!img.fileName || !img.mediaType || img.contentType) {
      continue;
    }

    const dotIdx = img.fileName.lastIndexOf(".");
    if (dotIdx === -1) {
      addIssue(
        "warning",
        `Image fileName "${img.fileName}" has no file extension`,
        `images[${i}]`,
        "image-no-extension"
      );
      continue;
    }

    const ext = img.fileName.slice(dotIdx + 1).toLowerCase();
    const expectedTypes = EXTENSION_MEDIA_TYPE_MAP.get(ext);

    if (!expectedTypes) {
      addIssue(
        "warning",
        `Image fileName "${img.fileName}" has unrecognized extension ".${ext}"`,
        `images[${i}]`,
        "image-unknown-extension"
      );
      continue;
    }

    if (!expectedTypes.includes(img.mediaType)) {
      addIssue(
        "error",
        `Image fileName "${img.fileName}" has extension ".${ext}" but mediaType is "${img.mediaType}" (expected: ${expectedTypes.join(" or ")})`,
        `images[${i}]`,
        "image-ext-mediatype-mismatch"
      );
    }
  }
}
