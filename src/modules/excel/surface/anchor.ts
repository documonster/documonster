/**
 * `Anchor` namespace surface — drawing anchor handle (cell coordinate + offset).
 *
 * `import { Anchor } from "documonster/excel"` →
 *   `Anchor.create(ws, "B2")`, `Anchor.col(a)`, `Anchor.setRow(a, 4)`,
 *   `Anchor.model(a)`, `Anchor.clone(a)`.
 */
export {
  anchorCreate as create,
  anchorAsInstance as asInstance,
  isAnchorData as isAnchor,
  anchorCol as col,
  anchorSetCol as setCol,
  anchorRow as row,
  anchorSetRow as setRow,
  anchorColWidth as colWidth,
  anchorRowHeight as rowHeight,
  anchorModel as model,
  anchorSetModel as setModel,
  anchorClone as clone
} from "@excel/core/anchor";

/** An anchor handle. */
export type { AnchorData as Handle } from "@excel/core/anchor";
