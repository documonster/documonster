/**
 * Style Resolution (Effective/Computed Style)
 *
 * Resolve the effective style for a paragraph by walking the style inheritance chain.
 */

import { checkHandle, checkNode } from "@word/core/argument-checks";
import type {
  DocxDocument,
  Paragraph,
  ParagraphProperties,
  Run,
  RunProperties,
  Shading,
  StyleDef,
  Table,
  TabStop,
  TableCell,
  TableLook,
  LevelJustification,
  NumberFormat,
  NumberingLevel,
  NumberingRef,
  TableProperties,
  TableStyleConditionalFormat,
  TableStyleConditionType
} from "@word/types";

// =============================================================================
// Types
// =============================================================================

/** Context for style resolution when a paragraph is inside a table. */
export interface StyleResolveContext {
  /** If the paragraph is inside a table, provide table context. */
  readonly tableContext?: {
    /** The table's `w:tblStyle`; omitted, the document's default table style. */
    readonly tableStyleId?: string;
    readonly tblLook?: TableLook;
    readonly rowIndex: number;
    readonly colIndex: number;
    readonly totalRows: number;
    readonly totalCols: number;
    /** `w:tblStyleRowBandSize` — rows per stripe. Defaults to 1. */
    readonly rowBandSize?: number;
    /** `w:tblStyleColBandSize` — columns per stripe. Defaults to 1. */
    readonly colBandSize?: number;
  };
}

/** Resolved paragraph properties with all inherited values merged. */
export interface ResolvedParagraphStyle {
  /** The style chain (from most specific to base). */
  readonly chain: readonly string[];
  /** Merged paragraph properties (inherited + own). */
  readonly paragraphProperties: ParagraphProperties;
  /** Merged run properties (inherited + own). */
  readonly runProperties: RunProperties;
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Styles by `styleId`. Build once with {@link indexStyles} and pass it to the
 * resolvers below to skip re-indexing `doc.styles` on every call. The index is
 * a snapshot: rebuild it after editing the document's styles.
 */
export type StyleIndex = ReadonlyMap<string, StyleDef>;

/** Index `doc.styles` by `styleId` (later duplicates win, as before). */
export function indexStyles(doc: DocxDocument): StyleIndex {
  checkHandle(doc, "document", "doc", "Query.createStyleIndex");
  const map = new Map<string, StyleDef>();
  for (const s of doc.styles ?? []) {
    map.set(s.styleId, s);
  }
  return map;
}

/**
 * Resolve the effective (computed) style for a paragraph by walking the style inheritance chain.
 *
 * Merges properties from the document defaults → base style chain → table conditional formats → paragraph's own properties.
 *
 * @param doc - The document containing styles and defaults.
 * @param para - The paragraph to resolve styles for.
 * @param context - Optional context providing table position for conditional format overlay.
 * @param styles - Optional prebuilt {@link indexStyles} index, reused across calls.
 * @returns The fully resolved paragraph style with all inherited properties merged.
 */
export function resolveStyle(
  doc: DocxDocument,
  para: Paragraph,
  context?: StyleResolveContext,
  styles?: StyleIndex
): ResolvedParagraphStyle {
  const styleMap = styles ?? indexStyles(doc);

  // Walk the chain from paragraph's style to root
  const chain: string[] = [];
  // A paragraph naming no style — or a style the document does not define —
  // takes the default paragraph style, as Word does (ECMA-376 §17.7.4.17).
  const named = para.properties?.style;
  const styleId =
    named && styleMap.has(named) ? named : findDefaultStyle(doc.styles, "paragraph")?.styleId;
  if (styleId) {
    let current: string | undefined = styleId;
    const visited = new Set<string>();
    while (current && !visited.has(current)) {
      visited.add(current);
      const def = styleMap.get(current);
      // A `basedOn` naming a style the document does not define ends the chain there.
      if (!def) {
        break;
      }
      chain.push(current);
      current = def.basedOn;
    }
  }

  // Build merged properties: start from doc defaults, apply chain bottom-up, then paragraph
  let mergedPProps: Record<string, unknown> = {};
  let mergedRProps: Record<string, unknown> = {};

  // Document defaults
  if (doc.docDefaults) {
    if (doc.docDefaults.paragraphProperties) {
      mergedPProps = mergeProperties({}, doc.docDefaults.paragraphProperties);
    }
    if (doc.docDefaults.runProperties) {
      mergedRProps = mergeProperties({}, doc.docDefaults.runProperties);
    }
  }

  // Toggle properties combine per *level* of the hierarchy, not per merge
  // step (see TOGGLE_RUN_PROPERTY_KEYS), so the paragraph-style level and the
  // table-style level are each collected on their own as well as merged.
  const docDefaultRProps: Record<string, unknown> = mergedRProps;
  let paragraphStyleLevel: Record<string, unknown> = {};
  let tableStyleLevel: Record<string, unknown> = {};

  // A paragraph in a table takes its table style's paragraph and run
  // properties above the document defaults and below its own paragraph style
  // (ECMA-376 §17.7.2). A table naming no style uses the default table style.
  const tblCtx = context?.tableContext;
  const tableChain = tblCtx ? tableStyleChain(doc, tblCtx.tableStyleId, styleMap) : [];
  for (let i = tableChain.length - 1; i >= 0; i--) {
    const def = tableChain[i];
    if (def.paragraphProperties) {
      mergedPProps = mergeProperties(mergedPProps, def.paragraphProperties);
    }
    if (def.runProperties) {
      mergedRProps = mergeProperties(mergedRProps, def.runProperties);
      tableStyleLevel = mergeProperties(tableStyleLevel, def.runProperties);
    }
  }

  // Apply style chain (from base to most specific)
  for (let i = chain.length - 1; i >= 0; i--) {
    const def = styleMap.get(chain[i]);
    if (!def) {
      continue;
    }
    if (def.paragraphProperties) {
      mergedPProps = mergeProperties(mergedPProps, def.paragraphProperties);
    }
    if (def.runProperties) {
      mergedRProps = mergeProperties(mergedRProps, def.runProperties);
      paragraphStyleLevel = mergeProperties(paragraphStyleLevel, def.runProperties);
    }
    // Linked styles: if a paragraph style links to a character style, merge its runProperties.
    // The linked character style's runProperties layer on top of the paragraph style's own
    // runProperties at the same level, giving the linked style slightly higher specificity.
    if (def.type === "paragraph" && def.link) {
      const linkedDef = styleMap.get(def.link);
      if (linkedDef?.type === "character" && linkedDef.runProperties) {
        mergedRProps = mergeProperties(mergedRProps, linkedDef.runProperties);
        paragraphStyleLevel = mergeProperties(paragraphStyleLevel, linkedDef.runProperties);
      }
    }
  }

  // Apply table conditional format overlay (higher priority than base table/paragraph style,
  // but lower priority than paragraph's own direct properties), base style first.
  if (tblCtx && tableChain.length > 0) {
    let rowBand = tblCtx.rowBandSize;
    let colBand = tblCtx.colBandSize;
    for (let i = tableChain.length - 1; i >= 0; i--) {
      rowBand = rowBand ?? tableChain[i].tableProperties?.rowBandSize;
      colBand = colBand ?? tableChain[i].tableProperties?.colBandSize;
    }
    for (let i = tableChain.length - 1; i >= 0; i--) {
      const conditions = tableChain[i].tableStyleConditions;
      if (!conditions) {
        continue;
      }
      const matchingConditions = getMatchingTableConditions(
        conditions,
        tblCtx.tblLook,
        tblCtx.rowIndex,
        tblCtx.colIndex,
        tblCtx.totalRows,
        tblCtx.totalCols,
        { row: rowBand, column: colBand }
      );
      for (const cond of matchingConditions) {
        if (cond.paragraphProperties) {
          mergedPProps = mergeProperties(mergedPProps, cond.paragraphProperties);
        }
        if (cond.runProperties) {
          mergedRProps = mergeProperties(mergedRProps, cond.runProperties);
          tableStyleLevel = mergeProperties(tableStyleLevel, cond.runProperties);
        }
      }
    }
  }

  applyToggleLevels(mergedRProps, docDefaultRProps, [paragraphStyleLevel, tableStyleLevel]);

  // Apply paragraph's own properties (most specific)
  if (para.properties) {
    const { style: _s, sectionProperties: _sp, ...ownPProps } = para.properties;
    mergedPProps = mergeProperties(mergedPProps, ownPProps);
  }

  return {
    chain,
    paragraphProperties: mergedPProps as ParagraphProperties,
    runProperties: mergedRProps as RunProperties
  };
}

// =============================================================================
// Internal helpers
// =============================================================================

/**
 * Property keys whose value OOXML expresses as a bag of independently
 * inherited attributes — or, for `borders` and `cellMargins`, as a container of
 * independently inherited child elements.
 *
 * Word merges these one attribute at a time down the style hierarchy: a style
 * declaring `<w:spacing w:after="154"/>` does **not** reset the `w:line` it
 * inherited, and `<w:ind w:left="720"/>` does not reset `w:firstLine`.
 * Replacing the whole value instead silently dropped whatever a lower level had
 * set, which is how `ListParagraph` — declaring only `after` — lost the
 * document default's 1.31 line spacing and rendered every list item at single
 * spacing while the paragraphs around it stayed 24% looser.
 *
 * A key is listed only when the merge can actually change something. A `Border`
 * or a `TableWidth` has a required member, so no level can declare half of one
 * and replacing it whole is already the correct answer.
 */
const MERGEABLE_PROPERTY_KEYS: ReadonlySet<string> = new Set([
  "spacing", // w:spacing — before / after / line / lineRule
  "indent", // w:ind — left / right / firstLine / hanging / start / end
  "borders", // w:pBdr / w:tblBorders / w:tcBorders — per side
  "cellMargins", // w:tblCellMar — per side
  "shading", // w:shd — val / color / fill
  "font", // w:rFonts — ascii / hAnsi / eastAsia / cs and their themes
  "language", // w:lang — val / eastAsia / bidi
  "color", // w:color — val / themeColor / themeTint / themeShade
  "underline", // w:u — val / color
  "look", // w:tblLook
  "float", // w:tblpPr
  "frame", // w:framePr
  "markRunProperties" // the paragraph mark's w:rPr — merges like any other rPr
]);

/** Whether a value is a mergeable object rather than a scalar or an array. */
function isPropertyBag(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Layer one style-hierarchy level onto the levels below it.
 *
 * An absent member means "inherit", not "reset", so `undefined` never erases
 * what a lower level set. The attribute-bag keys above are merged rather than
 * replaced, and recursively: `markRunProperties` therefore gets the same
 * treatment as a top-level run-properties bag, and a `borders` container merges
 * per side while each side is still replaced whole.
 *
 * The `isPropertyBag` guard on both sides is what lets a single key set serve
 * paragraph, run and table properties even where a name is reused for different
 * shapes — `RunProperties.spacing` is a `Twips` and `TableProperties.indent` a
 * number, and a scalar never takes the merge branch.
 */
function mergeProperties(base: Record<string, unknown>, override: object): Record<string, unknown> {
  const result: Record<string, unknown> = { ...base };
  for (const [key, value] of Object.entries(override)) {
    if (value === undefined) {
      continue;
    }
    const existing = result[key];
    if (key === "tabs" && Array.isArray(existing) && Array.isArray(value)) {
      result[key] = mergeTabStops(existing as readonly TabStop[], value as readonly TabStop[]);
      continue;
    }
    if (MERGEABLE_PROPERTY_KEYS.has(key) && isPropertyBag(existing) && isPropertyBag(value)) {
      result[key] = mergeProperties(existing, value);
      continue;
    }
    result[key] = value;
  }
  return result;
}

/**
 * Tab stops accumulate down the style hierarchy, one position at a time: a
 * level's stop replaces the one below at the same position, and its `clear`
 * removes it (ECMA-376 §17.3.1.37). Replacing the whole list dropped the stops
 * a style set — the Header style's centre and right stops under a paragraph
 * that added one of its own.
 */
function mergeTabStops(base: readonly TabStop[], override: readonly TabStop[]): TabStop[] {
  const at = new Set(override.map(t => t.position));
  return [
    ...base.filter(t => !at.has(t.position)),
    ...override.filter(t => t.type !== "clear")
  ].sort((a, b) => a.position - b.position);
}

/**
 * Determine which table style conditions match the given cell position.
 *
 * Returns matching conditions in application order (banding → whole row/col → corner cells).
 * Within Word, the specificity order from lowest to highest is:
 *   banding → first/last row/col → corner cells.
 */
function getMatchingTableConditions(
  conditions: readonly TableStyleConditionalFormat[],
  look: TableLook | undefined,
  rowIndex: number,
  colIndex: number,
  totalRows: number,
  totalCols: number,
  bandSize?: { readonly row?: number; readonly column?: number }
): TableStyleConditionalFormat[] {
  // Resolve effective banding flags: noHBand means band rows are disabled,
  // noVBand means band columns are disabled.
  const bandRow = look?.noHBand !== true;
  const bandCol = look?.noVBand !== true;

  // `w:tblStyleRowBandSize` / `w:tblStyleColBandSize`: how many rows or columns
  // make up one stripe. Alternating per single row regardless of it turned every
  // "banded by twos" table style into a one-row zebra.
  const rowBand = Math.max(1, Math.trunc(bandSize?.row ?? 1));
  const colBand = Math.max(1, Math.trunc(bandSize?.column ?? 1));

  // Build a set of applicable condition types based on position and tblLook.
  const applicable = new Set<TableStyleConditionType>();

  // Banding (lowest priority among conditions)
  if (bandRow) {
    if (Math.floor(rowIndex / rowBand) % 2 === 0) {
      applicable.add("oddRowBanding");
    } else {
      applicable.add("evenRowBanding");
    }
  }
  if (bandCol) {
    if (Math.floor(colIndex / colBand) % 2 === 0) {
      applicable.add("oddColumnBanding");
    } else {
      applicable.add("evenColumnBanding");
    }
  }

  // Whole row/column conditions
  if (rowIndex === 0 && look?.firstRow !== false) {
    applicable.add("firstRow");
  }
  if (rowIndex === totalRows - 1 && look?.lastRow !== false) {
    applicable.add("lastRow");
  }
  if (colIndex === 0 && look?.firstColumn !== false) {
    applicable.add("firstColumn");
  }
  if (colIndex === totalCols - 1 && look?.lastColumn !== false) {
    applicable.add("lastColumn");
  }

  // Corner cells (highest priority among conditions)
  if (rowIndex === 0 && colIndex === 0 && look?.firstRow !== false && look?.firstColumn !== false) {
    applicable.add("topLeftCell");
  }
  if (
    rowIndex === 0 &&
    colIndex === totalCols - 1 &&
    look?.firstRow !== false &&
    look?.lastColumn !== false
  ) {
    applicable.add("topRightCell");
  }
  if (
    rowIndex === totalRows - 1 &&
    colIndex === 0 &&
    look?.lastRow !== false &&
    look?.firstColumn !== false
  ) {
    applicable.add("bottomLeftCell");
  }
  if (
    rowIndex === totalRows - 1 &&
    colIndex === totalCols - 1 &&
    look?.lastRow !== false &&
    look?.lastColumn !== false
  ) {
    applicable.add("bottomRightCell");
  }

  // Filter and sort conditions by specificity order:
  // banding < first/last row/col < corner cells
  const priorityOrder: readonly TableStyleConditionType[] = [
    "oddRowBanding",
    "evenRowBanding",
    "oddColumnBanding",
    "evenColumnBanding",
    "firstRow",
    "lastRow",
    "firstColumn",
    "lastColumn",
    "topLeftCell",
    "topRightCell",
    "bottomLeftCell",
    "bottomRightCell"
  ];
  const priorityMap = new Map<TableStyleConditionType, number>();
  for (let i = 0; i < priorityOrder.length; i++) {
    priorityMap.set(priorityOrder[i], i);
  }

  return conditions
    .filter(c => applicable.has(c.type))
    .sort((a, b) => (priorityMap.get(a.type) ?? 0) - (priorityMap.get(b.type) ?? 0));
}

// =============================================================================
// Table cell formatting
// =============================================================================

/** Where a cell sits in its table, which is what selects the conditional formats. */
export interface TableCellPosition {
  readonly rowIndex: number;
  readonly colIndex: number;
  readonly totalRows: number;
  readonly totalCols: number;
}

/** Cell-level formatting a table style contributes to one cell position. */
export interface ResolvedTableCellStyle {
  /** Effective background shading, or undefined for no fill. */
  readonly shading?: Shading;
}

/** The default table style of each style index (`w:default="1"`), found once. */
const defaultTableStyles = new WeakMap<StyleIndex, StyleDef | null>();

/**
 * The table style a table uses: the one it names (`w:tblStyle`), or — for a
 * table that names none — the document's default table style
 * (`w:default="1"`), which Word applies to such a table exactly as if it named
 * it. Undefined when there is neither. This is the one place that rule lives;
 * everything that applies a table style goes through it.
 */
export function resolveEffectiveTableStyleId(
  doc: DocxDocument,
  tableStyleId: string | undefined,
  styles?: StyleIndex
): string | undefined {
  if (tableStyleId || !doc.styles) {
    return tableStyleId;
  }
  const styleMap = styles ?? indexStyles(doc);
  let fallback = defaultTableStyles.get(styleMap);
  if (fallback === undefined) {
    fallback = findDefaultStyle(doc.styles, "table") ?? null;
    defaultTableStyles.set(styleMap, fallback);
  }
  return fallback?.styleId;
}

/**
 * The default style of `type` — the one marked `w:default="1"`. A document may
 * mark more than one; Word takes the last (Word for Mac 16.114: two default
 * table styles, the tables take the second's formatting), so this does too.
 */
export function findDefaultStyle(
  styles: readonly StyleDef[] | undefined,
  type: StyleDef["type"]
): StyleDef | undefined {
  return styles?.findLast(s => s.type === type && s.isDefault === true);
}

/**
 * The style chain of a table, most specific first, starting from
 * {@link resolveEffectiveTableStyleId}. Empty when the table has no style.
 */
function tableStyleChain(
  doc: DocxDocument,
  tableStyleId: string | undefined,
  styles?: StyleIndex
): StyleDef[] {
  if (!doc.styles) {
    return [];
  }
  const styleMap = styles ?? indexStyles(doc);
  let current = resolveEffectiveTableStyleId(doc, tableStyleId, styleMap);
  const chain: StyleDef[] = [];
  const visited = new Set<string>();
  while (current && !visited.has(current)) {
    visited.add(current);
    const def = styleMap.get(current);
    if (!def) {
      break;
    }
    chain.push(def);
    current = def.basedOn;
  }
  return chain;
}

/**
 * A table's effective properties: its style chain's `w:tblPr` (from
 * {@link resolveEffectiveTableStyleId}, base first), then its own on top,
 * merged attribute by attribute — with `style` set to the effective style id.
 * Returns `table.properties` itself when the table has no style.
 */
export function resolveTableProperties(
  doc: DocxDocument,
  table: Table,
  styles?: StyleIndex
): TableProperties | undefined {
  const chain = tableStyleChain(doc, table.properties?.style, styles);
  if (chain.length === 0) {
    return table.properties;
  }
  let merged: Record<string, unknown> = {};
  for (let i = chain.length - 1; i >= 0; i--) {
    const own = chain[i].tableProperties;
    if (own) {
      merged = mergeProperties(merged, own);
    }
  }
  if (table.properties) {
    merged = mergeProperties(merged, table.properties);
  }
  return { ...(merged as TableProperties), style: chain[0].styleId };
}

/**
 * Resolve the cell-level formatting a table style contributes at one position.
 *
 * This is where Word's built-in table styles keep the look that makes a table
 * read as a table: the header band, the alternating row stripes, the emphasised
 * corner cells. They live in `tableStyleConditions[].cellProperties`, which
 * `resolveStyle` deliberately ignores (it answers a question about a paragraph),
 * so cells need their own resolution pass.
 *
 * Precedence, lowest to highest — matching Word:
 *   1. the table style chain's base `tableProperties.shading`
 *   2. its conditional formats, in `getMatchingTableConditions` order
 *      (banding → first/last row/column → corner cells)
 *
 * A cell's own `w:shd` and the table's direct `w:tblPr/w:shd` are *direct*
 * formatting and outrank everything here; callers apply them on top.
 */
export function resolveTableCellStyle(
  doc: DocxDocument,
  tableStyleId: string | undefined,
  look: TableLook | undefined,
  position: TableCellPosition,
  bandSize?: { readonly row?: number; readonly column?: number },
  styles?: StyleIndex
): ResolvedTableCellStyle {
  const chain = tableStyleChain(doc, tableStyleId, styles);
  if (chain.length === 0) {
    return {};
  }

  let shading: Shading | undefined;
  // Band sizes may be declared by the table or by any style in its chain; the
  // table's own value wins, then the most derived style that sets one.
  let rowBand = bandSize?.row;
  let colBand = bandSize?.column;
  for (let i = chain.length - 1; i >= 0; i--) {
    rowBand = rowBand ?? chain[i].tableProperties?.rowBandSize;
    colBand = colBand ?? chain[i].tableProperties?.colBandSize;
  }
  // Walk base → specific so a derived style overrides the one it is based on.
  for (let i = chain.length - 1; i >= 0; i--) {
    const def = chain[i];
    if (def.tableProperties?.shading) {
      shading = def.tableProperties.shading;
    }
    if (!def.tableStyleConditions) {
      continue;
    }
    const matching = getMatchingTableConditions(
      def.tableStyleConditions,
      look,
      position.rowIndex,
      position.colIndex,
      position.totalRows,
      position.totalCols,
      { row: rowBand, column: colBand }
    );
    for (const cond of matching) {
      if (cond.cellProperties?.shading) {
        shading = cond.cellProperties.shading;
      }
    }
  }

  return shading ? { shading } : {};
}

/**
 * The background colour of a table cell as a bare 6-digit hex string, or
 * undefined when nothing should be painted.
 *
 * Applies the full OOXML precedence in one place so every renderer agrees:
 *
 *   cell `w:tcPr/w:shd`  →  table `w:tblPr/w:shd`  →  table style
 *
 * `fill: "auto"` means "let the consumer decide", which for a cell background is
 * no fill; `pattern: "nil"` is an explicit absence of shading. The striped and
 * percentage patterns are approximated by their fill colour, which is what they
 * amount to at document scale.
 */
export function resolveTableCellFill(
  doc: DocxDocument,
  table: Table,
  cell: TableCell,
  position: TableCellPosition,
  styles?: StyleIndex
): string | undefined {
  // Direct formatting wins even when it paints nothing: `w:shd w:val="clear"
  // w:fill="auto"` is exactly how Word records "this cell has no shading",
  // which has to override whatever the table style would have contributed.
  // Treating that as "unspecified" and falling through would make it impossible
  // to clear a styled header band.
  if (cell.properties?.shading) {
    return resolveShadingFill(cell.properties.shading);
  }
  if (table.properties?.shading) {
    return resolveShadingFill(table.properties.shading);
  }
  const fromStyle = resolveTableCellStyle(
    doc,
    table.properties?.style,
    table.properties?.look,
    position,
    { row: table.properties?.rowBandSize, column: table.properties?.colBandSize },
    styles
  ).shading;
  return resolveShadingFill(fromStyle);
}

/** A `RRGGBB` hex string normalised from `w:fill` / `w:color`, or undefined. */
function normalizeHex(raw: string | undefined): string | undefined {
  if (!raw || raw === "auto") {
    return undefined;
  }
  const stripped = raw.replace(/^#/, "");
  if (/^[0-9a-fA-F]{6}$/.test(stripped)) {
    return stripped.toUpperCase();
  }
  if (/^[0-9a-fA-F]{3}$/.test(stripped)) {
    return stripped
      .split("")
      .map(ch => ch + ch)
      .join("")
      .toUpperCase();
  }
  return undefined;
}

/**
 * How much of a shading pattern's area its pattern colour covers, 0–1.
 *
 * A `w:shd` is a two-colour pattern: `w:fill` behind, `w:color` in the pattern
 * itself. `pct25` is literally 25 % coverage, and the named hatches are close to
 * fixed ratios. Flattening each to its coverage-weighted blend reproduces the
 * tone the pattern reads as at document scale, which is what a reader sees —
 * far closer than dropping the pattern colour entirely and painting bare `fill`.
 *
 * Returns 0 for a pattern that shows only the fill, and 1 for one that hides it.
 */
function patternCoverage(pattern: string | undefined): number {
  if (!pattern || pattern === "clear") {
    return 0;
  }
  if (pattern === "solid") {
    return 1;
  }
  const pct = /^pct(\d+)$/.exec(pattern);
  if (pct) {
    return Math.min(100, Math.max(0, Number.parseInt(pct[1], 10))) / 100;
  }
  // Hatches: the `thin*` variants use a 1px stroke, the rest a heavier one.
  if (pattern.startsWith("thin")) {
    return pattern.includes("Cross") ? 0.25 : 0.15;
  }
  if (pattern.endsWith("Cross")) {
    return 0.5;
  }
  if (pattern.endsWith("Stripe")) {
    return 0.3;
  }
  return 0;
}

/** Blend `over` onto `under` at `ratio`, both `RRGGBB`. */
function blendHex(under: string, over: string, ratio: number): string {
  let out = "";
  for (let i = 0; i < 6; i += 2) {
    const a = Number.parseInt(under.slice(i, i + 2), 16);
    const b = Number.parseInt(over.slice(i, i + 2), 16);
    out += Math.round(a + (b - a) * ratio)
      .toString(16)
      .toUpperCase()
      .padStart(2, "0");
  }
  return out;
}

/**
 * A shading's paintable fill as a bare `RRGGBB` hex string, or undefined when it
 * paints nothing.
 *
 * `fill: "auto"` means "let the consumer decide", which for a background is no
 * fill; `pattern: "nil"` is an explicit absence of shading. A pattern blends its
 * `w:color` into the fill by its coverage (see `patternCoverage`). Exported so
 * table cells and paragraphs cannot disagree about what a given `w:shd` paints.
 */
export function resolveShadingFill(shading: Shading | undefined): string | undefined {
  if (!shading || shading.pattern === "nil") {
    return undefined;
  }
  const fill = normalizeHex(shading.fill);
  const patternColor = normalizeHex(shading.color);
  const coverage = patternCoverage(shading.pattern);

  if (patternColor === undefined || coverage === 0) {
    return fill;
  }
  if (fill === undefined) {
    // No fill behind the pattern: the pattern colour is all there is to paint,
    // and only in proportion to how much of the area it covers. Blend onto
    // white, the page it sits on.
    return coverage >= 1 ? patternColor : blendHex("FFFFFF", patternColor, coverage);
  }
  return coverage >= 1 ? patternColor : blendHex(fill, patternColor, coverage);
}

// =============================================================================
// Extended Style Resolution APIs
// =============================================================================

/**
 * Toggle properties (ECMA-376 Part 1 §17.7.3): `b`, `bCs`, `caps`, `emboss`,
 * `i`, `iCs`, `imprint`, `outline`, `shadow`, `smallCaps`, `strike`, `vanish`.
 * `dstrike` is not one — its element definition (§17.3.2.9) does not call it a
 * toggle property, unlike `strike` (§17.3.2.37) and `vanish` (§17.3.2.41).
 *
 * Rules, from §17.7.3 as clarified by Microsoft's implementer notes
 * ([MS-OI29500] "Part 1 Section 17.7.3, Toggle Properties",
 * https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/f7130225-2368-48f3-acae-a9d278d0fb25):
 *
 * - Direct run formatting is absolute and never toggles.
 * - The *levels* of the style hierarchy are the table style, the paragraph
 *   style and the character style. Within one level the `basedOn` chain is
 *   ordinary inheritance: a derived style that does not set the property
 *   inherits its base's value, and a nearer setting overrides (note c: a
 *   paragraph style that does not set bold, based on one that sets it true, is
 *   bold). A paragraph style's linked character style belongs to its level.
 * - A level that sets nothing takes the document default (note b).
 * - Across levels the values toggle. With a `false` document default the result
 *   is true iff an odd number of levels are true; with a `true` default, Word
 *   makes it true iff an even number of levels are false (note a). Both are
 *   "the default, flipped once per level whose value differs from it" — which is
 *   {@link applyToggleLevels}.
 */
const TOGGLE_RUN_PROPERTY_KEYS = [
  "bold",
  "boldCs",
  "italic",
  "italicCs",
  "caps",
  "smallCaps",
  "strike",
  "outline",
  "shadow",
  "emboss",
  "imprint",
  "vanish"
] as const satisfies readonly (keyof RunProperties)[];

/**
 * Overwrite each toggle property in `target` with the §17.7.3 combination of
 * `defaults` and the style `levels` (see {@link TOGGLE_RUN_PROPERTY_KEYS}).
 * Keys no input sets are left as `target` has them.
 */
function applyToggleLevels(
  target: Record<string, unknown>,
  defaults: Record<string, unknown>,
  levels: readonly Record<string, unknown>[]
): void {
  for (const key of TOGGLE_RUN_PROPERTY_KEYS) {
    const base = defaults[key] === true;
    let value = base;
    let specified = typeof defaults[key] === "boolean";
    for (const level of levels) {
      const own = level[key];
      if (typeof own === "boolean") {
        specified = true;
        if (own !== base) {
          value = !value;
        }
      }
    }
    if (specified) {
      target[key] = value;
    }
  }
}

/** Resolved run style with full inheritance chain. */
export interface ResolvedRunStyle {
  /** Style chain (most specific → base). */
  readonly chain: readonly string[];
  /** Merged run properties. */
  readonly runProperties: RunProperties;
}

/**
 * Resolve the effective style for a single Run by walking the character
 * style inheritance chain.
 *
 * Resolution order (low → high specificity):
 * 1. Document defaults
 * 2. Paragraph's resolved style (if `paragraphRunProperties` provided)
 * 3. Run's character style chain (if `run.properties.style` is set);
 *    toggle properties combine with layers 1–2 per §17.7.3 — see
 *    `TOGGLE_RUN_PROPERTY_KEYS`. `paragraphRunProperties` must therefore come
 *    from {@link resolveStyle}, which has already combined the paragraph- and
 *    table-style levels.
 * 4. Run's own direct properties
 *
 * @param doc - The document containing styles.
 * @param run - The run to resolve.
 * @param paragraphRunProperties - Optional inherited run properties from the
 *   parent paragraph's resolved style. Pass `resolveStyle(doc, para).runProperties`
 *   to layer the paragraph style on top of doc defaults.
 * @param styles - Optional prebuilt {@link indexStyles} index, reused across calls.
 * @returns The fully resolved run style.
 */
export function resolveRunStyle(
  doc: DocxDocument,
  run: Run,
  paragraphRunProperties?: RunProperties,
  styles?: StyleIndex
): ResolvedRunStyle {
  const styleMap = styles ?? indexStyles(doc);

  // Build chain from run's character style
  const chain: string[] = [];
  // A run naming no character style takes the default one, as a paragraph does.
  const named = run.properties?.style;
  const runStyleId =
    named && styleMap.has(named) ? named : findDefaultStyle(doc.styles, "character")?.styleId;
  if (runStyleId) {
    let current: string | undefined = runStyleId;
    const visited = new Set<string>();
    while (current && !visited.has(current)) {
      visited.add(current);
      const def = styleMap.get(current);
      if (!def) {
        break;
      }
      chain.push(current);
      current = def.basedOn;
    }
  }

  let merged: Record<string, unknown> = {};

  // 1. Document defaults
  if (doc.docDefaults?.runProperties) {
    merged = mergeProperties({}, doc.docDefaults.runProperties);
  }

  // 2. Inherited from paragraph's resolved style
  if (paragraphRunProperties) {
    merged = mergeProperties(merged, paragraphRunProperties);
  }

  // 3. Run's character style chain (base → specific): one level of the
  // hierarchy. Within it a nearer style overrides; against the levels below,
  // toggle properties flip (see TOGGLE_RUN_PROPERTY_KEYS).
  let characterStyle: Record<string, unknown> = {};
  for (let i = chain.length - 1; i >= 0; i--) {
    const def = styleMap.get(chain[i]);
    if (def?.runProperties) {
      characterStyle = mergeProperties(characterStyle, def.runProperties);
    }
  }
  const inherited = merged;
  merged = mergeProperties(merged, characterStyle);
  const defaults = doc.docDefaults?.runProperties ?? {};
  for (const key of TOGGLE_RUN_PROPERTY_KEYS) {
    const own = characterStyle[key];
    if (typeof own !== "boolean") {
      continue;
    }
    const base = defaults[key] === true;
    const inheritedValue = inherited[key];
    const below = typeof inheritedValue === "boolean" ? inheritedValue : base;
    merged[key] = own === base ? below : !below;
  }

  // 4. Run's own direct properties (highest priority)
  if (run.properties) {
    const { style: _s, ...own } = run.properties;
    merged = mergeProperties(merged, own);
  }

  return {
    chain,
    runProperties: merged as RunProperties
  };
}

/** Resolved numbering level information. */
export interface ResolvedNumberingLevel {
  /** The level index (0-8). */
  readonly level: number;
  /** Start value (with any `w:startOverride` applied). */
  readonly start?: number;
  /** Number format. */
  readonly format?: NumberFormat;
  /** Level text template (e.g. `"%1."`). */
  readonly text?: string;
  /** Justification. */
  readonly justification?: LevelJustification;
  /** Run properties for the numbering text itself (bullet/number marker). */
  readonly runProperties?: RunProperties;
  /** Paragraph properties from the level (indent, alignment, etc.). */
  readonly paragraphProperties?: ParagraphProperties;
}

/**
 * Resolve the numbering level definition that is in effect for `numId` +
 * `level` (the `w:num` instance and `w:ilvl`).
 *
 * Walks `w:num` → `w:abstractNum` → `w:lvl`, then applies the instance's
 * `w:lvlOverride` for that level: a replacement `w:lvl` supersedes the
 * abstract definition, and `w:startOverride` replaces its start value. This is
 * the single place numbering is looked up — every converter goes through it,
 * so a list restarted or re-typed by an override renders the same everywhere.
 *
 * @param doc - The document.
 * @param numId - Numbering instance ID. `0` means "no numbering".
 * @param level - Level index (0-8).
 * @returns The effective level, or undefined if it cannot be resolved.
 */
export function resolveNumberingLevelDef(
  doc: DocxDocument,
  numId: number,
  level: number
): NumberingLevel | undefined {
  if (numId === 0) {
    return undefined;
  }
  const instance = doc.numberingInstances?.find(n => n.numId === numId);
  if (!instance) {
    return undefined;
  }
  const override = instance.overrides?.find(o => o.level === level);
  let levelDef: NumberingLevel | undefined = override?.levelDef;
  if (!levelDef) {
    const absNum = doc.abstractNumberings?.find(a => a.abstractNumId === instance.abstractNumId);
    levelDef = absNum?.levels.find(l => l.level === level);
  }
  if (!levelDef) {
    return undefined;
  }
  if (override?.startOverride !== undefined) {
    return { ...levelDef, start: override.startOverride };
  }
  return levelDef;
}

/**
 * Resolve the effective numbering reference (`w:numPr`) of a paragraph.
 *
 * Direct paragraph numbering wins; otherwise the paragraph style chain
 * (`w:pStyle` → `basedOn` …) is searched for the nearest `w:numPr`. A
 * `numId` of `0` at whichever level supplies it means "no numbering" and
 * cancels anything inherited, so it resolves to undefined.
 *
 * @param doc - The document.
 * @param para - The paragraph.
 * @param styles - Optional prebuilt {@link indexStyles} index, reused across calls.
 * @returns The effective numbering reference, or undefined when not a list item.
 */
export function resolveParagraphNumbering(
  doc: DocxDocument,
  para: Paragraph,
  styles?: StyleIndex
): NumberingRef | undefined {
  const direct = para.properties?.numbering;
  if (direct) {
    return direct.numId === 0 ? undefined : direct;
  }
  let current = para.properties?.style;
  if (!current || !doc.styles) {
    return undefined;
  }
  const styleMap = styles ?? indexStyles(doc);
  const visited = new Set<string>();
  while (current && !visited.has(current)) {
    visited.add(current);
    const def = styleMap.get(current);
    const numbering = def?.paragraphProperties?.numbering;
    if (numbering) {
      return numbering.numId === 0 ? undefined : numbering;
    }
    current = def?.basedOn;
  }
  return undefined;
}

/**
 * Resolve the numbering level for a paragraph.
 *
 * Uses the paragraph's effective numbering ({@link resolveParagraphNumbering})
 * and the effective level definition ({@link resolveNumberingLevelDef}), so
 * style-inherited numbering and `w:lvlOverride` are both honoured.
 *
 * @param doc - The document.
 * @param para - The paragraph.
 * @returns The resolved level, or undefined if no numbering or level not found.
 */
export function resolveNumberingLevel(
  doc: DocxDocument,
  para: Paragraph
): ResolvedNumberingLevel | undefined {
  checkHandle(doc, "document", "doc", "Query.getNumberingLevel");
  checkNode(para, "paragraph", "para", "Query.getNumberingLevel");
  const numRef = resolveParagraphNumbering(doc, para);
  if (!numRef) {
    return undefined;
  }
  const levelDef = resolveNumberingLevelDef(doc, numRef.numId, numRef.level);
  if (!levelDef) {
    return undefined;
  }

  return {
    level: levelDef.level,
    start: levelDef.start,
    format: levelDef.format,
    text: levelDef.text,
    justification: levelDef.justification,
    runProperties: levelDef.runProperties,
    paragraphProperties: levelDef.paragraphProperties
  };
}

/** A table style resolved along its `basedOn` chain (`Query.getTableStyle`). */
export interface ResolvedTableStyle {
  /** The style chain (from most specific to base). */
  readonly chain: readonly string[];
  /** Merged paragraph properties of the chain. */
  readonly paragraphProperties: ParagraphProperties;
  /** Merged run properties of the chain. */
  readonly runProperties: RunProperties;
  /** Merged table properties of the chain, if any style sets them. */
  readonly tableProperties?: TableProperties;
}

/**
 * Resolve table-level styles for a given table.
 *
 * Walks the table style inheritance chain (basedOn) to merge table properties.
 *
 * @param doc - The document.
 * @param tableStyleId - The starting table style ID.
 * @returns The merged table-level style chain.
 */
export function resolveTableStyle(
  doc: DocxDocument,
  tableStyleId: string,
  styles?: StyleIndex
): {
  chain: string[];
  paragraphProperties: ParagraphProperties;
  runProperties: RunProperties;
  tableProperties?: TableProperties;
} {
  const styleMap = styles ?? indexStyles(doc);

  const chain: string[] = [];
  let current: string | undefined = tableStyleId;
  const visited = new Set<string>();
  while (current && !visited.has(current)) {
    visited.add(current);
    const def = styleMap.get(current);
    // An id the document does not define is not part of the chain: a
    // misspelt id resolves to the document defaults, with an empty chain.
    if (!def) {
      break;
    }
    chain.push(current);
    current = def.basedOn;
  }

  let pProps: Record<string, unknown> = {};
  let rProps: Record<string, unknown> = {};
  let tProps: Record<string, unknown> = {};

  // Apply doc defaults first
  if (doc.docDefaults?.paragraphProperties) {
    pProps = mergeProperties({}, doc.docDefaults.paragraphProperties);
  }
  if (doc.docDefaults?.runProperties) {
    rProps = mergeProperties({}, doc.docDefaults.runProperties);
  }

  // Apply chain (base → specific)
  for (let i = chain.length - 1; i >= 0; i--) {
    const def = styleMap.get(chain[i]);
    if (!def) {
      continue;
    }
    if (def.paragraphProperties) {
      pProps = mergeProperties(pProps, def.paragraphProperties);
    }
    if (def.runProperties) {
      rProps = mergeProperties(rProps, def.runProperties);
    }
    if (def.tableProperties) {
      tProps = mergeProperties(tProps, def.tableProperties);
    }
  }

  return {
    chain,
    paragraphProperties: pProps as ParagraphProperties,
    runProperties: rProps as RunProperties,
    tableProperties: Object.keys(tProps).length > 0 ? (tProps as TableProperties) : undefined
  };
}
