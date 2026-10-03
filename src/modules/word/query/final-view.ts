/**
 * Final view — what a reader of the document sees: every tracked change
 * accepted (inserted and moved-to text shown, deleted and moved-from text
 * not), and hidden text (`w:vanish`) omitted.
 */

import { isHyperlink, isRun } from "@word/core/text-utils";
import type { ParagraphChild, Run, RunProperties } from "@word/types";

/**
 * The run a paragraph child contributes to the final view.
 *
 * Unwraps `insertedRun` / `movedToRun` to their inner run, returns a plain run
 * as is, and returns `undefined` for `deletedRun` / `movedFromRun` and for
 * every non-run child (hyperlinks, bookmarks, comment marks, …) — a caller
 * handles those itself.
 */
export function finalViewRun(child: ParagraphChild): Run | undefined {
  if (isRun(child)) {
    return child;
  }
  switch (child.type) {
    case "insertedRun":
    case "movedToRun":
      return child.run;
    default:
      return undefined;
  }
}

/**
 * Whether a run is hidden text. Pass the run's *resolved* properties
 * (`resolveRunStyle`), since `w:vanish` is commonly set by a style.
 */
export function isHiddenRun(runProperties: RunProperties | undefined): boolean {
  return runProperties?.vanish === true;
}

/**
 * The runs of a paragraph as the final view shows them, in document order:
 * revision wrappers resolved by {@link finalViewRun}, hyperlinks descended
 * into. Hidden text is left to the caller, which owns the resolved properties.
 */
export function* finalViewRuns(children: readonly ParagraphChild[]): Generator<Run> {
  for (const child of children) {
    const run = finalViewRun(child);
    if (run) {
      yield run;
    } else if (isHyperlink(child)) {
      yield* finalViewRuns(child.children as readonly ParagraphChild[]);
    }
  }
}
