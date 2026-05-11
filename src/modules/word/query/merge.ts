/**
 * Document Merge API
 *
 * Merge multiple DocxDocument bodies into a single document.
 *
 * The first document is taken as the base (its styles, settings, page layout,
 * etc. become the foundation). Each subsequent document is appended after a
 * section break. Numbering definitions, instances and the body references
 * that point to them are rewritten when the appended document collides with
 * IDs already present in the base. Body content from appended documents is
 * deep-cloned so that callers' models are never mutated.
 */

import type {
  DocxDocument,
  AbstractNumbering,
  BodyContent,
  NumberingInstance,
  NumberingRef,
  Paragraph,
  ParagraphProperties,
  StructuredDocumentTag,
  Table,
  TableCell,
  TableRow
} from "../types";

// =============================================================================
// Types
// =============================================================================

/** Options for merging documents. */
export interface MergeOptions {
  /** Break type between merged documents. Default: "nextPage". */
  readonly sectionBreak?: "continuous" | "nextPage" | "evenPage" | "oddPage";
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Merge multiple DocxDocument bodies into a single document.
 *
 * The first document is used as the base (preserving its styles, numbering, settings, etc.).
 * Subsequent documents' body content is appended with section breaks between them.
 * Numbering instance/abstract IDs that collide with the base are remapped, and
 * each appended document's body is deep-cloned to keep callers' models intact.
 *
 * @param documents - Array of documents to merge (at least 1).
 * @param options - Optional merge settings.
 * @returns A new merged DocxDocument.
 */
export function mergeDocuments(
  documents: readonly DocxDocument[],
  options?: MergeOptions
): DocxDocument {
  if (documents.length === 0) {
    return { body: [] };
  }
  if (documents.length === 1) {
    return documents[0];
  }

  const base = documents[0];
  const breakType = options?.sectionBreak ?? "nextPage";
  const mergedBody: BodyContent[] = [...base.body];

  const mergedImages = base.images ? [...base.images] : [];
  const mergedStyles = base.styles ? [...base.styles] : [];
  const existingStyleIds = new Set(mergedStyles.map(s => s.styleId));
  const mergedAbstractNums: AbstractNumbering[] = base.abstractNumberings
    ? [...base.abstractNumberings]
    : [];
  const mergedNumInstances: NumberingInstance[] = base.numberingInstances
    ? [...base.numberingInstances]
    : [];

  for (let i = 1; i < documents.length; i++) {
    const doc = documents[i];

    // Insert section break before appending next document
    const sectionBreakPara: Paragraph = {
      type: "paragraph",
      properties: {
        sectionProperties: {
          breakType
        }
      },
      children: []
    };
    mergedBody.push(sectionBreakPara);

    // Compute id remappings BEFORE cloning the body so we can rewrite refs
    // during the clone in a single pass.
    const { absIdMap, numIdMap } = mergeNumberingDefinitions(
      doc,
      mergedAbstractNums,
      mergedNumInstances
    );

    // Deep-clone body and rewrite numbering refs as needed.
    const cloned = doc.body.map(b => cloneBlockWithNumIdRemap(b, numIdMap));
    mergedBody.push(...cloned);

    // Merge images (avoid duplicates by fileName)
    if (doc.images) {
      for (const img of doc.images) {
        if (!mergedImages.some(m => m.fileName === img.fileName)) {
          mergedImages.push(img);
        }
      }
    }

    // Merge styles (avoid duplicates by styleId)
    if (doc.styles) {
      for (const style of doc.styles) {
        if (!existingStyleIds.has(style.styleId)) {
          mergedStyles.push(style);
          existingStyleIds.add(style.styleId);
        }
      }
    }

    // absIdMap is consumed inside mergeNumberingDefinitions; nothing else to do.
    void absIdMap;
  }

  return {
    ...base,
    body: mergedBody,
    images: mergedImages.length > 0 ? mergedImages : undefined,
    styles: mergedStyles.length > 0 ? mergedStyles : undefined,
    abstractNumberings: mergedAbstractNums.length > 0 ? mergedAbstractNums : undefined,
    numberingInstances: mergedNumInstances.length > 0 ? mergedNumInstances : undefined,
    // Use the final document's section properties (page layout) if available
    sectionProperties: documents[documents.length - 1].sectionProperties ?? base.sectionProperties
  };
}

// =============================================================================
// Numbering merge
// =============================================================================

/**
 * Append `doc`'s numbering definitions/instances into the running merge state,
 * remapping abstractNumIds and numIds that collide with what is already there.
 * The returned maps describe the remapping that callers must apply to body refs.
 */
function mergeNumberingDefinitions(
  doc: DocxDocument,
  mergedAbstractNums: AbstractNumbering[],
  mergedNumInstances: NumberingInstance[]
): { absIdMap: Map<number, number>; numIdMap: Map<number, number> } {
  const absIdMap = new Map<number, number>();
  const numIdMap = new Map<number, number>();

  if (doc.abstractNumberings && doc.abstractNumberings.length > 0) {
    const existingAbsIds = new Set(mergedAbstractNums.map(a => a.abstractNumId));
    let nextAbsId = existingAbsIds.size > 0 ? Math.max(...existingAbsIds) + 1 : 0;
    for (const abs of doc.abstractNumberings) {
      if (existingAbsIds.has(abs.abstractNumId)) {
        const newId = nextAbsId++;
        absIdMap.set(abs.abstractNumId, newId);
        existingAbsIds.add(newId);
        mergedAbstractNums.push({ ...abs, abstractNumId: newId });
      } else {
        existingAbsIds.add(abs.abstractNumId);
        mergedAbstractNums.push(abs);
      }
    }
  }

  if (doc.numberingInstances && doc.numberingInstances.length > 0) {
    const existingNumIds = new Set(mergedNumInstances.map(n => n.numId));
    let nextNumId = existingNumIds.size > 0 ? Math.max(...existingNumIds) + 1 : 1;
    for (const inst of doc.numberingInstances) {
      const remappedAbsId = absIdMap.get(inst.abstractNumId) ?? inst.abstractNumId;
      let numId = inst.numId;
      if (existingNumIds.has(numId)) {
        numId = nextNumId++;
        numIdMap.set(inst.numId, numId);
      }
      existingNumIds.add(numId);
      mergedNumInstances.push({ ...inst, numId, abstractNumId: remappedAbsId });
    }
  }

  return { absIdMap, numIdMap };
}

// =============================================================================
// Body cloning with numId remap
// =============================================================================

/**
 * Deep-clone a body content block while rewriting any paragraph numbering
 * references through `numIdMap`. The original block is never mutated.
 */
function cloneBlockWithNumIdRemap(block: BodyContent, numIdMap: Map<number, number>): BodyContent {
  switch (block.type) {
    case "paragraph":
      return cloneParagraph(block, numIdMap);
    case "table":
      return cloneTable(block, numIdMap);
    case "sdt":
      return cloneSdt(block as StructuredDocumentTag, numIdMap);
    default:
      // For other block types (textBox, tableOfContents, floatingImage, etc.)
      // the numId never appears at this depth; structuredClone gives us a safe
      // independent copy without us having to enumerate every kind.
      return structuredClone(block);
  }
}

function cloneParagraph(para: Paragraph, numIdMap: Map<number, number>): Paragraph {
  const clonedChildren = structuredClone(para.children);
  const clonedProps = para.properties ? cloneParagraphProps(para.properties, numIdMap) : undefined;
  const out: Paragraph = {
    ...para,
    properties: clonedProps,
    children: clonedChildren
  };
  return out;
}

function cloneParagraphProps(
  props: ParagraphProperties,
  numIdMap: Map<number, number>
): ParagraphProperties {
  const cloned = structuredClone(props) as ParagraphProperties & { numbering?: NumberingRef };
  if (cloned.numbering && numIdMap.has(cloned.numbering.numId)) {
    (cloned.numbering as { numId: number; level: number }).numId = numIdMap.get(
      cloned.numbering.numId
    )!;
  }
  return cloned;
}

function cloneTable(table: Table, numIdMap: Map<number, number>): Table {
  return {
    ...table,
    rows: table.rows.map(r => cloneRow(r, numIdMap))
  } as Table;
}

function cloneRow(row: TableRow, numIdMap: Map<number, number>): TableRow {
  return {
    ...row,
    cells: row.cells.map(c => cloneCell(c, numIdMap))
  } as TableRow;
}

function cloneCell(cell: TableCell, numIdMap: Map<number, number>): TableCell {
  return {
    ...cell,
    content: cell.content.map(item => cloneBlockWithNumIdRemap(item as BodyContent, numIdMap))
  } as TableCell;
}

function cloneSdt(
  sdt: StructuredDocumentTag,
  numIdMap: Map<number, number>
): StructuredDocumentTag {
  return {
    ...sdt,
    content: sdt.content.map(c => {
      if ("type" in c && (c.type === "paragraph" || c.type === "table")) {
        return cloneBlockWithNumIdRemap(c as BodyContent, numIdMap) as Paragraph | Table;
      }
      // Run / other in-line children: structural clone is enough.
      return structuredClone(c);
    })
  } as StructuredDocumentTag;
}
