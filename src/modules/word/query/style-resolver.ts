/**
 * A document's style resolution, bound once per conversion or layout.
 *
 * The free functions in `style-resolve` and `heading` index `doc.styles` on
 * every call, which is what a caller editing the document between calls needs.
 * A converter or layout pass walks a document that does not change underneath
 * it, so it creates one resolver up front and every paragraph and run reuses
 * the same index. Create a new resolver after editing the document's styles.
 */

import { resolveHeadingLevel } from "@word/query/heading";
import type { HeadingInfo } from "@word/query/heading";
import {
  indexStyles,
  resolveParagraphNumbering,
  resolveRunStyle,
  resolveStyle,
  resolveTableCellFill
} from "@word/query/style-resolve";
import type {
  ResolvedParagraphStyle,
  ResolvedRunStyle,
  StyleIndex,
  StyleResolveContext,
  TableCellPosition
} from "@word/query/style-resolve";
import type {
  DocxDocument,
  NumberingRef,
  Paragraph,
  Run,
  RunProperties,
  Table,
  TableCell
} from "@word/types";

/** Style resolution bound to one document snapshot. */
export interface StyleResolver {
  readonly doc: DocxDocument;
  readonly styles: StyleIndex;
  /** {@link resolveStyle} */
  paragraph(para: Paragraph, context?: StyleResolveContext): ResolvedParagraphStyle;
  /** {@link resolveRunStyle} */
  run(run: Run, paragraphRunProperties?: RunProperties): ResolvedRunStyle;
  /** {@link resolveParagraphNumbering} */
  numbering(para: Paragraph): NumberingRef | undefined;
  /** {@link resolveHeadingLevel} */
  heading(para: Paragraph): HeadingInfo | undefined;
  /** {@link resolveTableCellFill} */
  tableCellFill(table: Table, cell: TableCell, position: TableCellPosition): string | undefined;
}

/** Bind style resolution to `doc`'s current styles. */
export function createStyleResolver(doc: DocxDocument): StyleResolver {
  const styles = indexStyles(doc);
  return {
    doc,
    styles,
    paragraph: (para, context) => resolveStyle(doc, para, context, styles),
    run: (run, paragraphRunProperties) => resolveRunStyle(doc, run, paragraphRunProperties, styles),
    numbering: para => resolveParagraphNumbering(doc, para, styles),
    heading: para => resolveHeadingLevel(doc, para, styles),
    tableCellFill: (table, cell, position) =>
      resolveTableCellFill(doc, table, cell, position, styles)
  };
}
