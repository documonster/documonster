/**
 * Style Resolution (Effective/Computed Style)
 *
 * Resolve the effective style for a paragraph by walking the style inheritance chain.
 */

import type {
  DocxDocument,
  Paragraph,
  ParagraphProperties,
  Run,
  RunProperties,
  StyleDef,
  TableLook,
  NumberingLevel,
  TableProperties,
  TableStyleConditionalFormat,
  TableStyleConditionType
} from "../types";

// =============================================================================
// Types
// =============================================================================

/** Context for style resolution when a paragraph is inside a table. */
export interface StyleResolveContext {
  /** If the paragraph is inside a table, provide table context. */
  readonly tableContext?: {
    readonly tableStyleId?: string;
    readonly tblLook?: TableLook;
    readonly rowIndex: number;
    readonly colIndex: number;
    readonly totalRows: number;
    readonly totalCols: number;
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
 * Resolve the effective (computed) style for a paragraph by walking the style inheritance chain.
 *
 * Merges properties from the document defaults → base style chain → table conditional formats → paragraph's own properties.
 *
 * @param doc - The document containing styles and defaults.
 * @param para - The paragraph to resolve styles for.
 * @param context - Optional context providing table position for conditional format overlay.
 * @returns The fully resolved paragraph style with all inherited properties merged.
 */
export function resolveStyle(
  doc: DocxDocument,
  para: Paragraph,
  context?: StyleResolveContext
): ResolvedParagraphStyle {
  const styleMap = new Map<string, StyleDef>();
  if (doc.styles) {
    for (const s of doc.styles) {
      styleMap.set(s.styleId, s);
    }
  }

  // Walk the chain from paragraph's style to root
  const chain: string[] = [];
  const styleId = para.properties?.style;
  if (styleId) {
    let current: string | undefined = styleId;
    const visited = new Set<string>();
    while (current && !visited.has(current)) {
      visited.add(current);
      chain.push(current);
      const def = styleMap.get(current);
      current = def?.basedOn;
    }
  }

  // Build merged properties: start from doc defaults, apply chain bottom-up, then paragraph
  let mergedPProps: Record<string, unknown> = {};
  let mergedRProps: Record<string, unknown> = {};

  // Document defaults
  if (doc.docDefaults) {
    if (doc.docDefaults.paragraphProperties) {
      mergedPProps = { ...doc.docDefaults.paragraphProperties };
    }
    if (doc.docDefaults.runProperties) {
      mergedRProps = { ...doc.docDefaults.runProperties };
    }
  }

  // Apply style chain (from base to most specific)
  for (let i = chain.length - 1; i >= 0; i--) {
    const def = styleMap.get(chain[i]);
    if (!def) {
      continue;
    }
    if (def.paragraphProperties) {
      mergedPProps = { ...mergedPProps, ...stripUndefined(def.paragraphProperties) };
    }
    if (def.runProperties) {
      mergedRProps = { ...mergedRProps, ...stripUndefined(def.runProperties) };
    }
    // Linked styles: if a paragraph style links to a character style, merge its runProperties.
    // The linked character style's runProperties layer on top of the paragraph style's own
    // runProperties at the same level, giving the linked style slightly higher specificity.
    if (def.type === "paragraph" && def.link) {
      const linkedDef = styleMap.get(def.link);
      if (linkedDef?.type === "character" && linkedDef.runProperties) {
        mergedRProps = { ...mergedRProps, ...stripUndefined(linkedDef.runProperties) };
      }
    }
  }

  // Apply table conditional format overlay (higher priority than base table/paragraph style,
  // but lower priority than paragraph's own direct properties).
  if (context?.tableContext?.tableStyleId) {
    const tblCtx = context.tableContext;
    const tableStyleId: string = tblCtx.tableStyleId!;
    const tblStyleDef = styleMap.get(tableStyleId);
    if (tblStyleDef?.tableStyleConditions) {
      const matchingConditions = getMatchingTableConditions(
        tblStyleDef.tableStyleConditions,
        tblCtx.tblLook,
        tblCtx.rowIndex,
        tblCtx.colIndex,
        tblCtx.totalRows,
        tblCtx.totalCols
      );
      for (const cond of matchingConditions) {
        if (cond.paragraphProperties) {
          mergedPProps = { ...mergedPProps, ...stripUndefined(cond.paragraphProperties) };
        }
        if (cond.runProperties) {
          mergedRProps = { ...mergedRProps, ...stripUndefined(cond.runProperties) };
        }
      }
    }
  }

  // Apply paragraph's own properties (most specific)
  if (para.properties) {
    const { style: _s, sectionProperties: _sp, ...ownPProps } = para.properties;
    mergedPProps = { ...mergedPProps, ...stripUndefined(ownPProps) };
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

/** Remove undefined values from an object (so spreading doesn't override with undefined). */
function stripUndefined<T extends object>(obj: T): Record<string, unknown> {
  const result: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(obj)) {
    if (v !== undefined) {
      result[k] = v;
    }
  }
  return result;
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
  totalCols: number
): TableStyleConditionalFormat[] {
  // Resolve effective banding flags: noHBand means band rows are disabled,
  // noVBand means band columns are disabled.
  const bandRow = look?.noHBand !== true;
  const bandCol = look?.noVBand !== true;

  // Build a set of applicable condition types based on position and tblLook.
  const applicable = new Set<TableStyleConditionType>();

  // Banding (lowest priority among conditions)
  if (bandRow) {
    if (rowIndex % 2 === 0) {
      applicable.add("oddRowBanding");
    } else {
      applicable.add("evenRowBanding");
    }
  }
  if (bandCol) {
    if (colIndex % 2 === 0) {
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
// Extended Style Resolution APIs
// =============================================================================

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
 * 3. Run's character style chain (if `run.properties.style` is set)
 * 4. Run's own direct properties
 *
 * @param doc - The document containing styles.
 * @param run - The run to resolve.
 * @param paragraphRunProperties - Optional inherited run properties from the
 *   parent paragraph's resolved style. Pass `resolveStyle(doc, para).runProperties`
 *   to layer the paragraph style on top of doc defaults.
 * @returns The fully resolved run style.
 */
export function resolveRunStyle(
  doc: DocxDocument,
  run: Run,
  paragraphRunProperties?: RunProperties
): ResolvedRunStyle {
  const styleMap = new Map<string, StyleDef>();
  if (doc.styles) {
    for (const s of doc.styles) {
      styleMap.set(s.styleId, s);
    }
  }

  // Build chain from run's character style
  const chain: string[] = [];
  const runStyleId = run.properties?.style;
  if (runStyleId) {
    let current: string | undefined = runStyleId;
    const visited = new Set<string>();
    while (current && !visited.has(current)) {
      visited.add(current);
      chain.push(current);
      const def = styleMap.get(current);
      current = def?.basedOn;
    }
  }

  let merged: Record<string, unknown> = {};

  // 1. Document defaults
  if (doc.docDefaults?.runProperties) {
    merged = { ...doc.docDefaults.runProperties };
  }

  // 2. Inherited from paragraph's resolved style
  if (paragraphRunProperties) {
    merged = { ...merged, ...stripUndefined(paragraphRunProperties) };
  }

  // 3. Run's character style chain (base → specific)
  for (let i = chain.length - 1; i >= 0; i--) {
    const def = styleMap.get(chain[i]);
    if (def?.runProperties) {
      merged = { ...merged, ...stripUndefined(def.runProperties) };
    }
  }

  // 4. Run's own direct properties (highest priority)
  if (run.properties) {
    const { style: _s, ...own } = run.properties;
    merged = { ...merged, ...stripUndefined(own) };
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
  /** Number format. */
  readonly format?: string;
  /** Level text template (e.g. `"%1."`). */
  readonly text?: string;
  /** Justification. */
  readonly justification?: string;
  /** Run properties for the numbering text itself (bullet/number marker). */
  readonly runProperties?: RunProperties;
  /** Paragraph properties from the level (indent, alignment, etc.). */
  readonly paragraphProperties?: ParagraphProperties;
}

/**
 * Resolve the numbering level for a paragraph that has a numbering reference.
 *
 * Walks: paragraph.numbering → numberingInstances → abstractNumberings → level definition.
 * Also applies `LevelOverride` if present.
 *
 * @param doc - The document.
 * @param para - The paragraph (must have `numbering` set).
 * @returns The resolved level, or undefined if no numbering or level not found.
 */
export function resolveNumberingLevel(
  doc: DocxDocument,
  para: Paragraph
): ResolvedNumberingLevel | undefined {
  const numRef = para.properties?.numbering;
  if (!numRef) {
    return undefined;
  }

  // Find numbering instance
  const instance = doc.numberingInstances?.find(n => n.numId === numRef.numId);
  if (!instance) {
    return undefined;
  }

  // Check for level override first
  const override = instance.overrides?.find(o => o.level === numRef.level);
  let levelDef: NumberingLevel | undefined;
  if (override?.levelDef) {
    levelDef = override.levelDef;
  } else {
    // Walk to abstract numbering
    const absNum = doc.abstractNumberings?.find(a => a.abstractNumId === instance.abstractNumId);
    levelDef = absNum?.levels.find(l => l.level === numRef.level);
  }

  if (!levelDef) {
    return undefined;
  }

  return {
    level: levelDef.level,
    format: levelDef.format,
    text: levelDef.text,
    justification: levelDef.justification,
    runProperties: levelDef.runProperties,
    paragraphProperties: levelDef.paragraphProperties
  };
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
  tableStyleId: string
): {
  chain: string[];
  paragraphProperties: ParagraphProperties;
  runProperties: RunProperties;
  tableProperties?: TableProperties;
} {
  const styleMap = new Map<string, StyleDef>();
  if (doc.styles) {
    for (const s of doc.styles) {
      styleMap.set(s.styleId, s);
    }
  }

  const chain: string[] = [];
  let current: string | undefined = tableStyleId;
  const visited = new Set<string>();
  while (current && !visited.has(current)) {
    visited.add(current);
    chain.push(current);
    const def = styleMap.get(current);
    current = def?.basedOn;
  }

  let pProps: Record<string, unknown> = {};
  let rProps: Record<string, unknown> = {};
  let tProps: Record<string, unknown> = {};

  // Apply doc defaults first
  if (doc.docDefaults?.paragraphProperties) {
    pProps = { ...doc.docDefaults.paragraphProperties };
  }
  if (doc.docDefaults?.runProperties) {
    rProps = { ...doc.docDefaults.runProperties };
  }

  // Apply chain (base → specific)
  for (let i = chain.length - 1; i >= 0; i--) {
    const def = styleMap.get(chain[i]);
    if (!def) {
      continue;
    }
    if (def.paragraphProperties) {
      pProps = { ...pProps, ...stripUndefined(def.paragraphProperties) };
    }
    if (def.runProperties) {
      rProps = { ...rProps, ...stripUndefined(def.runProperties) };
    }
    if (def.tableProperties) {
      tProps = { ...tProps, ...stripUndefined(def.tableProperties) };
    }
  }

  return {
    chain,
    paragraphProperties: pProps as ParagraphProperties,
    runProperties: rProps as RunProperties,
    tableProperties: Object.keys(tProps).length > 0 ? (tProps as TableProperties) : undefined
  };
}
