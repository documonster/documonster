/**
 * The fields of the Word model that the schema's narrower property types have no place for.
 *
 * A container's properties come in more than one schema type: a paragraph's `w:pPr` is `CT_PPr`, but the previous
 * properties a `w:pPrChange` holds are `CT_PPrBase` and a style's, a default's or a numbering level's are
 * `CT_PPrGeneral` — neither has the paragraph's mark or section. The model has one type for all of them, so these
 * lists say which of its fields cannot be in the narrower ones. The reader sets such a field's element aside as not
 * written back, the argument checks and the writer refuse one a caller sets, and rejecting a change keeps the current
 * value of each, because the previous properties do not hold it. One list for all of them, so they cannot disagree.
 */

import type {
  ParagraphMarkRevisions,
  ParagraphProperties,
  SectionProperties,
  TableCellProperties,
  TableProperties,
  TableRowProperties
} from "@word/types";

/**
 * The fields of a paragraph's mark — its run properties and its insertion, deletion and move, all one `w:rPr`
 * (`CT_ParaRPr`).
 */
export const PARAGRAPH_MARK_PROPERTIES = [
  "markRunProperties",
  "paragraphInsertion",
  "paragraphDeletion",
  "paragraphMoveFrom",
  "paragraphMoveTo"
] as const satisfies readonly (keyof ParagraphProperties)[];

/**
 * A paragraph mark's revisions, in `CT_ParaRPr`'s order: the element, the paragraph's field for the mark's current
 * one, and the `ParagraphMarkRevisions` field for one a change record's previous properties hold.
 */
export const PARAGRAPH_MARK_REVISIONS = [
  { element: "ins", field: "paragraphInsertion", previous: "insertion" },
  { element: "del", field: "paragraphDeletion", previous: "deletion" },
  { element: "moveFrom", field: "paragraphMoveFrom", previous: "moveFrom" },
  { element: "moveTo", field: "paragraphMoveTo", previous: "moveTo" }
] as const satisfies readonly {
  readonly element: string;
  readonly field: keyof ParagraphProperties;
  readonly previous: keyof ParagraphMarkRevisions;
}[];

/**
 * The `CT_PPr` fields `CT_PPrBase` and `CT_PPrGeneral` do not have: the mark ({@link PARAGRAPH_MARK_PROPERTIES}) and
 * the section (`w:sectPr`).
 */
export const PARAGRAPH_ONLY_PROPERTIES = [
  ...PARAGRAPH_MARK_PROPERTIES,
  "sectionProperties"
] as const satisfies readonly (keyof ParagraphProperties)[];

/**
 * The row fields a row's previous properties (`CT_TrPrBase`, in `w:trPrChange`) do not have: its insertion and
 * deletion, which are `CT_TrPr`'s, and its table property exceptions, which are not in `w:trPr` at all but beside it
 * in the row.
 */
export const ROW_ONLY_PROPERTIES = [
  "inserted",
  "deleted",
  "tblPrEx"
] as const satisfies readonly (keyof TableRowProperties)[];

/**
 * The section fields a section's previous properties (`CT_SectPrBase`, in `w:sectPrChange`) do not have: its header
 * and footer references, which are `CT_SectPr`'s.
 */
export const SECTION_ONLY_PROPERTIES = [
  "headers",
  "footers"
] as const satisfies readonly (keyof SectionProperties)[];

/**
 * Which schema type a paragraph's properties are read and written as: a paragraph's own (`CT_PPr`), a style's, a
 * default's or a numbering level's (`general`, `CT_PPrGeneral`), or a change record's previous ones (`previous`,
 * `CT_PPrBase`, which has no change record either).
 */
export type ParagraphPropertiesScope = "paragraph" | "general" | "previous";

/** How a refusal names the properties a field has no place in, so the argument checks and the writer agree. */
export const NARROWER_PROPERTIES = {
  generalParagraph: "a style's, a default's or a numbering level's paragraph properties",
  previousParagraph: "a change record's previous paragraph properties",
  previousRun: "a change record's previous run properties",
  baseTable: "a style's table properties and a change record's previous ones",
  previousRow: "a change record's previous row properties",
  previousCell: "a change record's previous cell properties",
  previousExceptions: "a change record's previous row exceptions",
  previousSection: "a change record's previous section properties",
  styleTable: "a table style's table properties, as Word reads them,",
  conditionTable: "a table style's conditional table properties, as Word reads them,",
  conditionRow: "a table style's conditional row properties, as Word reads them,",
  conditionCell: "a table style's conditional cell properties, as Word reads them,"
} as const;

/**
 * What Word does not allow in a table style, by element, with the field the model reads it into. ECMA-376 gives a
 * style the table, row and cell properties a table has; Word takes only part of them ([MS-OI29500] §17.7.6.3, .4, .8,
 * .9, .10, .11), drops the rest, and the Open XML SDK's validator — Office's schema — refuses the part. The lists are
 * the validator's, probed one element at a time. The reader sets such an element aside as not written back, and the
 * argument checks and the writer refuse the field.
 */
export const TABLE_STYLE_EXCLUDED = {
  styleTable: {
    tblStyle: "style",
    tblpPr: "float",
    tblOverlap: "float",
    bidiVisual: "visuallyRightToLeft",
    tblW: "width",
    tblLayout: "layout",
    tblLook: "look",
    tblCaption: "caption",
    tblDescription: "description"
  },
  conditionTable: {
    tblStyle: "style",
    tblpPr: "float",
    tblOverlap: "float",
    bidiVisual: "visuallyRightToLeft",
    tblStyleRowBandSize: "rowBandSize",
    tblStyleColBandSize: "colBandSize",
    tblW: "width",
    tblLayout: "layout",
    tblLook: "look",
    tblCaption: "caption",
    tblDescription: "description"
  },
  conditionRow: {
    cnfStyle: "cnfStyle",
    gridBefore: "gridBefore",
    gridAfter: "gridAfter",
    wBefore: "widthBefore",
    wAfter: "widthAfter",
    trHeight: "height",
    ins: "inserted",
    del: "deleted",
    trPrChange: "propertyChange"
  },
  conditionCell: {
    cnfStyle: "cnfStyle",
    tcW: "width",
    gridSpan: "gridSpan",
    vMerge: "verticalMerge",
    textDirection: "textDirection",
    tcFitText: "fitText",
    hideMark: "hideMark",
    cellIns: "inserted",
    cellDel: "deleted",
    cellMerge: "cellMerge",
    tcPrChange: "propertyChange"
  }
} as const satisfies {
  readonly styleTable: Readonly<Record<string, keyof TableProperties>>;
  readonly conditionTable: Readonly<Record<string, keyof TableProperties>>;
  readonly conditionRow: Readonly<Record<string, keyof TableRowProperties>>;
  readonly conditionCell: Readonly<Record<string, keyof TableCellProperties>>;
};

/** Which of a table style's properties: the style's own table properties, or a conditional format's. */
export type TableStyleProperties = keyof typeof TABLE_STYLE_EXCLUDED;

/**
 * The fields a table style's properties of `kind` cannot hold: those {@link TABLE_STYLE_EXCLUDED} reads elements into,
 * and the model's own that would write one — a cell's `rowSpan` is its `vMerge`, and a row's exceptions (`tblPrEx`)
 * sit beside a table's row, where a style has none.
 */
export function tableStyleExcludedFields(kind: TableStyleProperties): string[] {
  const fields = new Set<string>(Object.values(TABLE_STYLE_EXCLUDED[kind]));
  if (kind === "conditionCell") {
    fields.add("rowSpan");
  }
  if (kind === "conditionRow") {
    fields.add("tblPrEx");
  }
  return [...fields];
}

/** What a refusal says the field must be: `absent: a change record's previous run properties have no place for it`. */
export function absentFrom(properties: string): string {
  return `absent: ${properties} have no place for it`;
}

/**
 * The edges of a table's and a cell's borders and margins that have two names: `start`/`left` and `end`/`right`.
 * ECMA-376 lists both names in each content model, but they are one edge — the second edition renamed `left` and
 * `right` to `start` and `end` — and Office's schema takes one of each pair: the Open XML SDK's validator refuses a
 * `w:tblCellMar`, `w:tcMar`, `w:tblBorders` or `w:tcBorders` holding both. The model has a field for each name,
 * so the argument checks and the writer refuse both set, and the reader keeps the first a file has.
 */
export const EDGE_NAMES = [
  ["start", "left"],
  ["end", "right"]
] as const;

/** What a refusal says the second name of an edge must be. */
export function oneEdgeName(first: string): string {
  return `absent when ${first} is given: the two name one edge, and Office takes one`;
}
