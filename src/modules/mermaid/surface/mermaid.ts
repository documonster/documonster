/**
 * `Mermaid` namespace surface — Mermaid text to a display list, SVG, a parse tree or
 * a laid-out graph.
 *
 * `import { Mermaid } from "documonster/mermaid"` →
 *   `Mermaid.toSvg(source)`, `Mermaid.toPng(source)`, `Mermaid.toDrawList(source)`, `Mermaid.parse(source)`,
 *   `Mermaid.layoutGraph(diagram)`, `Mermaid.resolveTheme(options)`,
 *   `Mermaid.createCenteredText(box, lines, style)`.
 *
 * `layoutGraph` lays out any graph-shaped diagram — flowchart, state, class, ER,
 * requirement, C4 or architecture. Each is converted to the flowchart form and placed by
 * the one layered layout, with its boxes sized as the renderer draws them, so the name
 * says what it lays out rather than which diagram it was first written for.
 */

import type { DrawNode, DrawTextStyle } from "@draw/types";
import { DEFAULT_TEXT_FAMILY } from "@draw/types";
import { checkTextBox, checkTextStyleSize } from "@mermaid/arguments";
import { centeredText } from "@mermaid/render/shared";

export {
  layoutGraph,
  mermaidToDrawList as toDrawList,
  mermaidToPng as toPng,
  mermaidToSvg as toSvg
} from "@mermaid/diagram";
export { parseMermaid as parse } from "@mermaid/parse/index";
export { resolveTheme } from "@mermaid/theme";

/**
 * A text node with `lines` stacked in the middle of `box`, for a caller drawing its own boxes
 * from a layout (`MermaidNodeBox`, `MermaidGroupBox` and `MermaidEdgeRoute.label` all have the box shape).
 * `MermaidNodeBox.lines` and `MermaidEdgeRoute.label.lines` are already wrapped; a `MermaidGroupBox` has a single
 * `text` instead, so pass `[group.text]`. The style is any text style; its anchor is always
 * `middle`, and an omitted `fill` is black (where a raw text node without one is not drawn)
 * and an omitted `family` is the engine's default face.
 *
 * @throws {MermaidArgumentError} a box coordinate is not finite, `lines` is not a list of strings, or
 *   `style.size` is not a number > 0.
 */
export function createCenteredText(
  box: { x: number; y: number; width: number; height: number },
  lines: readonly string[],
  style: DrawTextStyle
): DrawNode {
  const where = "Mermaid.createCenteredText";
  checkTextBox(where, box, lines);
  checkTextStyleSize(where, style);
  const { size, family, fill, anchor: _anchor, ...extra } = style;
  return centeredText(box, lines, size, family ?? DEFAULT_TEXT_FAMILY, fill ?? BLACK, extra);
}

const BLACK = { r: 0, g: 0, b: 0, a: 1 } as const;
