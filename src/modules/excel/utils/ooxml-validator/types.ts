/**
 * Public types for the OOXML validator.
 *
 * The validator's job is to detect xlsx packages that Excel will refuse to
 * open (or silently repair). Each check emits zero or more `Problem` entries
 * tagged by `kind`; callers can branch on `kind` for structured handling.
 *
 * Problem kinds are grouped by subsystem to make ad-hoc extension easy.
 * New kinds MUST be added to the union below so consumers can enumerate them
 * at the type level. Removing or renaming kinds is a breaking change.
 */

// -----------------------------------------------------------------------------
// Package structure
// -----------------------------------------------------------------------------

export type PackageStructureKind = "missing-part" | "xml-malformed" | "part-name-invalid";

// -----------------------------------------------------------------------------
// Content types
// -----------------------------------------------------------------------------

export type ContentTypesKind =
  | "content-types-missing"
  | "content-types-malformed"
  | "content-types-missing-default"
  | "content-types-missing-for-part"
  | "content-types-duplicate-override"
  | "content-types-wrong-for-part";

// -----------------------------------------------------------------------------
// Relationships
// -----------------------------------------------------------------------------

export type RelationshipsKind =
  | "root-rels-missing-officeDocument"
  | "rels-malformed"
  | "rels-missing-target"
  | "rels-duplicate-id"
  | "rels-empty-target"
  | "rels-missing-id-attr"
  | "rels-missing-type-attr"
  | "rels-invalid-target-path"
  | "rels-source-missing"
  | "rels-type-target-mismatch";

// -----------------------------------------------------------------------------
// Workbook
// -----------------------------------------------------------------------------

export type WorkbookKind =
  | "workbook-sheet-missing-rel"
  | "workbook-sheet-wrong-rel-type"
  | "workbook-duplicate-sheetId"
  | "workbook-duplicate-sheet-rid"
  | "workbook-sheet-missing-name"
  | "workbook-sheet-name-too-long"
  | "workbook-sheet-name-invalid-chars"
  | "workbook-sheet-name-duplicate";

// -----------------------------------------------------------------------------
// Worksheet
// -----------------------------------------------------------------------------

export type WorksheetKind =
  | "sheet-missing-rels"
  | "sheet-child-out-of-order"
  | "sheet-legacyDrawing-after-controls"
  | "sheet-controls-missing-drawing"
  | "sheet-control-missing-rel"
  | "sheet-control-wrong-rel-type"
  | "sheet-legacyDrawing-missing-rel"
  | "sheet-legacyDrawing-wrong-rel-type"
  | "sheet-drawing-missing-rel"
  | "sheet-drawing-wrong-rel-type"
  | "sheet-comments-missing-rel"
  | "sheet-comments-wrong-rel-type"
  | "sheet-hyperlink-missing-rel"
  | "sheet-hyperlink-wrong-rel-type"
  | "sheet-tablePart-missing-rel"
  | "sheet-tablePart-wrong-rel-type"
  | "sheet-cell-ref-missing"
  | "sheet-cell-ref-invalid"
  | "sheet-cell-ref-row-mismatch"
  | "sheet-cell-ref-out-of-bounds"
  | "sheet-cell-style-index-oob"
  | "sheet-cell-sst-index-oob"
  | "sheet-merge-invalid-range"
  | "sheet-merge-overlap"
  | "sheet-row-index-out-of-bounds"
  | "sheet-sharedFormula-master-missing"
  | "sheet-sharedFormula-duplicate-master";

// -----------------------------------------------------------------------------
// Styles
// -----------------------------------------------------------------------------

export type StylesKind =
  | "styles-malformed"
  | "styles-numFmt-missing-for-xf"
  | "styles-font-index-oob"
  | "styles-fill-index-oob"
  | "styles-border-index-oob";

// -----------------------------------------------------------------------------
// Table
// -----------------------------------------------------------------------------
//
// These three kinds cover the two "Removed Records: Table from
// /xl/tables/tableN.xml" patterns that appear most often in real Excel
// logs: redundant per-column `<filterColumn hiddenButton="1"/>` (Excel
// treats a fully-hidden autoFilter on a live table as inconsistent),
// `<totalsRowFormula>` paired with a built-in `totalsRowFunction`
// (schema allows the child only when the function is "custom"), and
// autoFilter ref that covers the totals row (the filter range must end
// above the totals row). All three cause Excel to drop the whole table
// on open.
//
export type TableKind =
  | "table-filterColumn-redundant-per-column"
  | "table-totalsRowFormula-with-builtin-function"
  | "table-autoFilter-covers-totalsRow";

// -----------------------------------------------------------------------------
// Chart / drawing / chartsheet / pivot
// -----------------------------------------------------------------------------

export type ChartKind =
  | "chart-missing-chart"
  | "chart-missing-plotArea"
  | "chartEx-missing-chart"
  | "chartEx-missing-plotArea"
  | "chartEx-missing-series"
  | "chartEx-series-missing-layoutId"
  | "chartEx-series-missing-data-id"
  | "chartEx-series-missing-axis-id"
  | "chartEx-externalData-missing-rel"
  // Schema-conformance checks for chartEx. Each maps to an "Excel rejects the
  // chartEx part and drops the parent drawing" pattern observed in
  // real Excel repair logs. See `check-chart.ts` for details and
  // `__tests__/ooxml-validator/chart.test.ts` for negative samples.
  | "chartEx-series-too-many-dataId"
  | "chartEx-typed-element-text-form"
  | "chartEx-invalid-auto-element"
  | "chartEx-paretoLine-in-layoutPr"
  | "chartEx-title-direct-layout"
  // Tier-2 schema conformance — structurally detectable patterns that
  // need a little more semantic context than a pure XPath lookup.
  | "chartEx-axis-missing-pos-and-type"
  | "chartEx-f-uses-direct-range-not-defined-name"
  | "chartEx-waterfall-missing-subtotals"
  | "chartEx-chartStyle-stub-form"
  | "chartEx-chartColorStyle-stub-form";

export type DrawingKind =
  | "drawing-graphicFrame-missing-chart"
  | "drawing-anchor-missing"
  | "drawing-anchor-invalid-coords"
  // ChartEx-specific drawing-shape wrappers. Every cx:chart reference
  // must live inside an `<mc:AlternateContent>` with a non-empty
  // `<mc:Fallback>`, and the AlternateContent must sit INSIDE the
  // `<xdr:twoCellAnchor>` shape slot — not around it.
  | "drawing-chartEx-missing-alternateContent-wrap"
  | "drawing-chartEx-alternateContent-empty-fallback"
  | "drawing-chartEx-alternateContent-outer-wrap"
  // Office creation-id extension on chartEx drawings. Strict Excel
  // builds have been observed to reject chartEx drawings lacking this
  // extension. Surfaced as a warning because legacy drawings authored
  // before the 2014 extension was registered do load successfully.
  | "drawing-chartEx-missing-creationId";

export type ChartsheetKind = "chartsheet-missing-drawing";

export type PivotKind =
  | "pivot-missing-cacheId"
  | "pivot-cacheId-not-in-workbook"
  | "pivot-cacheRecords-missing";

// -----------------------------------------------------------------------------
// Union
// -----------------------------------------------------------------------------

/**
 * All problem kinds emitted by the validator. The historical legacy kind
 * `sheet-legacyDrawing-after-controls` is preserved for backwards
 * compatibility — new ordering violations use the generic
 * `sheet-child-out-of-order` kind.
 */
export type OoxmlProblemKind =
  | PackageStructureKind
  | ContentTypesKind
  | RelationshipsKind
  | WorkbookKind
  | WorksheetKind
  | StylesKind
  | TableKind
  | ChartKind
  | DrawingKind
  | ChartsheetKind
  | PivotKind;

/**
 * Historical alias: the old public API exposed a separate union for the
 * single ordering check. We preserve the name so downstream type-level
 * references keep compiling.
 */
export type OoxmlOrderingProblemKind = "sheet-legacyDrawing-after-controls";

// -----------------------------------------------------------------------------
// Severity
// -----------------------------------------------------------------------------

/**
 * `error`   — Excel will refuse the file or show a repair prompt.
 * `warning` — Likely to work but deviates from OPC/OOXML invariants; Excel
 *             may silently recover.
 */
export type OoxmlProblemSeverity = "error" | "warning";

// -----------------------------------------------------------------------------
// Problem shape
// -----------------------------------------------------------------------------

export interface OoxmlValidationProblem {
  kind: OoxmlProblemKind;
  severity: OoxmlProblemSeverity;
  /** The file this problem is reported for, if any. */
  file?: string;
  /** Human-readable description. Stable enough for snapshot tests. */
  message: string;
}

/**
 * Historical alias retained for callers that used the narrower type before
 * severity was added. The shape is identical to `OoxmlValidationProblem`.
 */
export type OoxmlOrderingValidationProblem = OoxmlValidationProblem;

// -----------------------------------------------------------------------------
// Report
// -----------------------------------------------------------------------------

export interface OoxmlValidationStats {
  entryCount: number;
  xmlLikeCount: number;
  relsCount: number;
}

export interface OoxmlValidationReport {
  /** `true` when there are no `error`-severity problems. Warnings do not fail. */
  ok: boolean;
  problems: OoxmlValidationProblem[];
  stats: OoxmlValidationStats;
}

// -----------------------------------------------------------------------------
// Options
// -----------------------------------------------------------------------------

export interface OoxmlValidateOptions {
  /**
   * Check every XML-like entry (.xml / .rels / .vml) for well-formedness.
   * Default: true.
   */
  checkXmlWellFormed?: boolean;

  /**
   * Validate relationship targets exist (TargetMode=External is skipped).
   * Default: true.
   */
  checkRelationshipTargets?: boolean;

  /**
   * Validate the content types table against package parts.
   * Default: true.
   */
  checkContentTypesOverrides?: boolean;

  /**
   * Validate worksheet r:id wiring (controls / drawing / comments / etc.)
   * and structural invariants (child element order, cell r= consistency,
   * merge overlaps, style index bounds).
   * Default: true.
   */
  checkWorksheetControlWiring?: boolean;

  /**
   * Validate chart / chartEx / drawing / chartsheet / pivot internal
   * structure. Default: true.
   */
  checkChartStructure?: boolean;

  /**
   * Validate styles.xml internal integrity (numFmt/font/fill/border
   * back-references from xf records). Default: true.
   */
  checkStylesIntegrity?: boolean;

  /**
   * Stop after this many problems across all checks. Useful for large
   * packages where the first few problems imply the rest.
   */
  maxProblems?: number;

  /**
   * Include warnings in the output. When `false` (default), warnings are
   * silently dropped. Errors are always reported regardless of this flag.
   */
  includeWarnings?: boolean;
}
