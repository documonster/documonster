/**
 * Mermaid source (or a parsed tree) to a display list, and the SVG convenience over it.
 *
 * Lives outside `index.ts` so the `Mermaid` namespace surface can re-export it.
 */

import { toSvg } from "@draw/svg";
import type { DrawList, DrawNode } from "@draw/types";
import { MermaidArgumentError } from "@mermaid/errors";
import { layoutFlowchart } from "@mermaid/layout/flowchart";
import type { FlowLayout, LayoutOptions } from "@mermaid/layout/flowchart";
import { parseMermaid } from "@mermaid/parse/index";
import {
  architectureNodes,
  architectureToFlowchart,
  kanbanDrawList,
  packetDrawList,
  radarDrawList
} from "@mermaid/render/beta";
import { quadrantDrawList, sankeyDrawList, xyDrawList } from "@mermaid/render/chart";
import { classDiagramNodes, classNodeSizer, classToFlowchart } from "@mermaid/render/class";
import { erDiagramNodes, erNodeSizer, erToFlowchart } from "@mermaid/render/er";
import { flowchartNodes } from "@mermaid/render/flowchart";
import { ganttDrawList } from "@mermaid/render/gantt";
import {
  blockDrawList,
  c4DiagramNodes,
  c4NodeSizer,
  c4ToFlowchart,
  requirementDiagramNodes,
  requirementNodeSizer,
  requirementToFlowchart
} from "@mermaid/render/model";
import { backdrop } from "@mermaid/render/shared";
import { pieDrawList, sequenceDrawList } from "@mermaid/render/simple";
import { stateToFlowchart } from "@mermaid/render/state";
import { journeyDrawList, timelineDrawList } from "@mermaid/render/track";
import { gitGraphDrawList, mindmapDrawList } from "@mermaid/render/tree";
import { checkColour, checkThemeOptions, resolveTheme } from "@mermaid/theme";
import type { Theme, ThemeOptions } from "@mermaid/theme";
import type {
  ArchitectureDiagram,
  C4Diagram,
  ClassDiagram,
  ErDiagram,
  FlowchartDiagram,
  MermaidDiagram,
  RequirementDiagram,
  StateDiagram
} from "@mermaid/types";
import { describeValue as describeOption } from "@utils/describe-value";

/** How to draw a diagram. */
export interface MermaidRenderOptions extends LayoutOptions {
  /** Colour overrides; anything omitted keeps the default theme. */
  readonly theme?: ThemeOptions;
  /**
   * Background fill, a CSS colour. `"transparent"` — the default — leaves the list without a
   * backdrop, so a caller compositing the diagram onto something else gets what they asked for.
   *
   * The same setting as `theme.background`, which is what `Mermaid.resolveTheme` takes; when
   * both are given this one wins (`"transparent"` included). The backdrop is the list's first
   * node, drawn once, so every backend — SVG, pixels, a PDF page — shows the same colour.
   */
  readonly background?: string;
}

/** The diagram kinds {@link layoutGraph} lays out. */
const GRAPH_KINDS: ReadonlySet<string> = new Set([
  "flowchart",
  "state",
  "class",
  "er",
  "requirement",
  "c4",
  "architecture"
]);

/** Every diagram kind the renderer draws. */
const DIAGRAM_KINDS: ReadonlySet<string> = new Set([
  ...GRAPH_KINDS,
  "pie",
  "gantt",
  "timeline",
  "journey",
  "mindmap",
  "git",
  "quadrant",
  "xy",
  "sankey",
  "packet",
  "kanban",
  "radar",
  "block",
  "sequence"
]);

function describeValue(value: unknown): string {
  if (value === null) {
    return "null";
  }
  if (typeof value === "object") {
    const kind = (value as { kind?: unknown }).kind;
    return typeof kind === "string" ? `a "${kind}" diagram` : "an object that is not a diagram";
  }
  return typeof value;
}

function optionError(where: string, field: string, expected: string, value: unknown): never {
  throw new MermaidArgumentError(
    field,
    `${where}: ${field} must be ${expected}; got ${describeOption(value)}.`
  );
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Check `LayoutOptions`. A `NaN` font size used to lay the diagram out as a 0×0 picture, and
 * a negative one as a plausible-looking but wrong one.
 */
function checkLayoutOptions(where: string, options: unknown): void {
  if (!isPlainObject(options)) {
    optionError(where, "options", "an options object", options);
  }
  const positive = (field: string): void => {
    const value = options[field];
    if (value !== undefined && (typeof value !== "number" || !(value > 0) || value === Infinity)) {
      optionError(where, `options.${field}`, "a finite number > 0", value);
    }
  };
  const nonNegative = (field: string): void => {
    const value = options[field];
    if (value !== undefined && (typeof value !== "number" || !(value >= 0) || value === Infinity)) {
      optionError(where, `options.${field}`, "a finite number >= 0", value);
    }
  };
  positive("fontSize");
  positive("maxLabelWidth");
  positive("maxEdgeLabelWidth");
  nonNegative("rankGap");
  nonNegative("nodeGap");
  nonNegative("padding");
  if (options.fontFamily !== undefined && typeof options.fontFamily !== "string") {
    optionError(where, "options.fontFamily", "a string", options.fontFamily);
  }
  if (options.measureNode !== undefined && typeof options.measureNode !== "function") {
    optionError(where, "options.measureNode", "a function", options.measureNode);
  }
}

/** Check `MermaidRenderOptions`: the layout options plus the theme. */
function checkRenderOptions(where: string, options: unknown): void {
  checkLayoutOptions(where, options);
  checkThemeOptions(where, (options as { theme?: unknown }).theme, "options.theme");
  const background = (options as { background?: unknown }).background;
  if (background !== undefined && typeof background !== "string") {
    optionError(where, "options.background", "a CSS colour string", background);
  }
  if (background !== undefined) {
    checkColour(where, "options.background", background);
  }
}

/**
 * Refuse, by name, an argument that is not a diagram of the kinds a function handles. A caller passing the wrong
 * thing used to get a `TypeError` from deep inside the layout ("diagram.nodes is not iterable"), which names neither
 * the function nor the argument.
 */
function checkDiagram(
  value: unknown,
  where: string,
  field: string,
  kinds: ReadonlySet<string>,
  expected: string
): void {
  const kind =
    typeof value === "object" && value !== null ? (value as { kind?: unknown }).kind : undefined;
  if (typeof kind !== "string" || !kinds.has(kind)) {
    throw new MermaidArgumentError(
      field,
      `${where}: ${field} must be ${expected}; got ${describeValue(value)}.`
    );
  }
}

const DEFAULT_FONT_SIZE = 14;
const DEFAULT_FONT_FAMILY = "Arial";
const DEFAULT_PADDING = 16;

/**
 * The diagrams laid out as a graph of boxes: each is converted to the flowchart form and
 * placed by the same layered layout.
 */
export type GraphDiagram =
  | FlowchartDiagram
  | StateDiagram
  | ClassDiagram
  | ErDiagram
  | RequirementDiagram
  | C4Diagram
  | ArchitectureDiagram;

/** A layout and the flowchart form it was computed from. */
interface GraphLayout {
  readonly flow: FlowchartDiagram;
  readonly layout: FlowLayout;
}

/**
 * Convert a graph-shaped diagram to the flowchart form and lay it out, sizing each box the
 * way that diagram type draws it — a class compartment stack, an entity table, a C4
 * element with its technology line. `measureNode`, when given, replaces that sizing.
 */
function layoutGraphDiagram(
  diagram: GraphDiagram,
  options: LayoutOptions,
  measureNode: LayoutOptions["measureNode"]
): GraphLayout {
  const fontSize = options.fontSize ?? DEFAULT_FONT_SIZE;
  const fontFamily = options.fontFamily ?? DEFAULT_FONT_FAMILY;
  const sized = (
    flow: FlowchartDiagram,
    sizer: NonNullable<LayoutOptions["measureNode"]>
  ): GraphLayout => ({
    flow,
    layout: layoutFlowchart(flow, { ...options, measureNode: measureNode ?? sizer })
  });
  const plain = (flow: FlowchartDiagram): GraphLayout => ({
    flow,
    layout: layoutFlowchart(flow, measureNode === undefined ? options : { ...options, measureNode })
  });
  switch (diagram.kind) {
    case "class":
      return sized(classToFlowchart(diagram), classNodeSizer(diagram, fontSize, fontFamily));
    case "er":
      return sized(erToFlowchart(diagram), erNodeSizer(diagram, fontSize, fontFamily));
    case "requirement":
      return sized(
        requirementToFlowchart(diagram),
        requirementNodeSizer(diagram, fontSize, fontFamily)
      );
    case "c4":
      return sized(c4ToFlowchart(diagram), c4NodeSizer(diagram, fontSize, fontFamily));
    case "architecture":
      return plain(architectureToFlowchart(diagram));
    case "state":
      // The same picture in a different language; converting it means one layout and one
      // renderer serve both.
      return plain(stateToFlowchart(diagram));
    default:
      return plain(diagram);
  }
}

/**
 * Lay out a graph-shaped diagram — flowchart, state, class, ER, requirement, C4 or
 * architecture — and return the coordinates of its boxes, edges and groups.
 *
 * Every one of them is converted to the flowchart form and placed by the same layered
 * layout, with each box sized as the built-in renderer draws it, so the geometry matches
 * `Mermaid.toSvg` for the same diagram. Supply `measureNode` to size boxes yourself.
 * Diagrams that are not graphs (Gantt, pie, sequence, …) have no such layout.
 */
export function layoutGraph(diagram: GraphDiagram, options: LayoutOptions = {}): FlowLayout {
  checkLayoutOptions("Mermaid.layoutGraph", options);
  checkDiagram(
    diagram,
    "Mermaid.layoutGraph",
    "diagram",
    GRAPH_KINDS,
    "a graph diagram (flowchart, state, class, ER, requirement, C4 or architecture) from Mermaid.parse"
  );
  return layoutGraphDiagram(diagram, options, options.measureNode).layout;
}

/**
 * Wrap a laid-out diagram's nodes into a display list.
 *
 * Five diagram types reached this same four-step shape — take the layout's size, put the
 * background down first if one was asked for, then the nodes — and a shape repeated five
 * times is four chances for one of them to forget the background or report the wrong size.
 */
function fromLayout(layout: FlowLayout, theme: Theme, nodes: readonly DrawNode[]): DrawList {
  return {
    width: layout.width,
    height: layout.height,
    // Background first: several callers depend on reading it as the list's opening node.
    children: [...backdrop(theme, layout.width, layout.height), ...nodes]
  };
}

/**
 * Draw a diagram, from source or from an already-parsed tree.
 *
 * @throws {MermaidParseError} when a string names no supported diagram.
 */
export function mermaidToDrawList(
  source: string | MermaidDiagram,
  options: MermaidRenderOptions = {}
): DrawList {
  return drawDiagram("Mermaid.toDrawList", source, options);
}

/** {@link mermaidToDrawList}, its argument errors named for the public call `where`. */
function drawDiagram(
  where: string,
  source: string | MermaidDiagram,
  options: MermaidRenderOptions
): DrawList {
  if (typeof source !== "string") {
    checkDiagram(
      source,
      where,
      "source",
      DIAGRAM_KINDS,
      "Mermaid text or a diagram from Mermaid.parse"
    );
  }
  checkRenderOptions(where, options);
  const diagram = typeof source === "string" ? parseMermaid(source) : source;
  const fontSize = options.fontSize ?? DEFAULT_FONT_SIZE;
  const fontFamily = options.fontFamily ?? DEFAULT_FONT_FAMILY;
  const padding = options.padding ?? DEFAULT_PADDING;
  const theme = resolveTheme(
    options.background === undefined
      ? options.theme
      : { ...options.theme, background: options.background }
  );

  if (diagram.kind === "pie") {
    return pieDrawList(diagram, theme, fontSize, fontFamily, padding);
  }
  if (diagram.kind === "gantt") {
    return ganttDrawList(diagram, theme, fontSize, fontFamily, padding);
  }
  if (diagram.kind === "timeline") {
    return timelineDrawList(diagram, theme, fontSize, fontFamily, padding);
  }
  if (diagram.kind === "journey") {
    return journeyDrawList(diagram, theme, fontSize, fontFamily, padding);
  }
  if (diagram.kind === "mindmap") {
    return mindmapDrawList(diagram, theme, fontSize, fontFamily, padding);
  }
  if (diagram.kind === "git") {
    return gitGraphDrawList(diagram, theme, fontSize, fontFamily, padding);
  }
  if (diagram.kind === "quadrant") {
    return quadrantDrawList(diagram, theme, fontSize, fontFamily, padding);
  }
  if (diagram.kind === "xy") {
    return xyDrawList(diagram, theme, fontSize, fontFamily, padding);
  }
  if (diagram.kind === "sankey") {
    return sankeyDrawList(diagram, theme, fontSize, fontFamily, padding);
  }
  if (diagram.kind === "packet") {
    return packetDrawList(diagram, theme, fontSize, fontFamily, padding);
  }
  if (diagram.kind === "kanban") {
    return kanbanDrawList(diagram, theme, fontSize, fontFamily, padding);
  }
  if (diagram.kind === "radar") {
    return radarDrawList(diagram, theme, fontSize, fontFamily, padding);
  }
  if (diagram.kind === "architecture") {
    const archLayout = layoutGraphDiagram(diagram, options, undefined).layout;
    return fromLayout(
      archLayout,
      theme,
      architectureNodes(archLayout, theme, fontSize, fontFamily)
    );
  }
  if (diagram.kind === "block") {
    return blockDrawList(diagram, theme, fontSize, fontFamily, padding);
  }
  if (diagram.kind === "requirement" || diagram.kind === "c4") {
    // Both are graphs of boxes that size themselves, so the shared layout does the hard
    // part and only the box design differs.
    const modelLayout = layoutGraphDiagram(diagram, options, undefined).layout;
    return fromLayout(
      modelLayout,
      theme,
      diagram.kind === "requirement"
        ? requirementDiagramNodes(diagram, modelLayout, theme, fontSize, fontFamily)
        : c4DiagramNodes(diagram, modelLayout, theme, fontSize, fontFamily)
    );
  }
  if (diagram.kind === "sequence") {
    return sequenceDrawList(diagram, theme, fontSize, fontFamily, padding);
  }

  // A class diagram sizes its own boxes — a compartment stack is not a centred label — but
  // everything difficult about a layout is shared.
  if (diagram.kind === "class") {
    const { layout } = layoutGraphDiagram(diagram, options, undefined);
    return fromLayout(
      layout,
      theme,
      classDiagramNodes(diagram, layout, theme, fontSize, fontFamily)
    );
  }

  if (diagram.kind === "er") {
    const { layout } = layoutGraphDiagram(diagram, options, undefined);
    return fromLayout(layout, theme, erDiagramNodes(diagram, layout, theme, fontSize, fontFamily));
  }

  const { flow, layout } = layoutGraphDiagram(diagram, options, undefined);
  return fromLayout(
    layout,
    theme,
    flowchartNodes(layout, theme, flow.classDefs, fontSize, fontFamily)
  );
}

/**
 * Draw a diagram as SVG.
 *
 * A convenience over {@link mermaidToDrawList} and `toSvg`; a caller that wants pixels or
 * a PDF page passes the list to the matching backend instead.
 *
 * @throws {MermaidParseError} when the source names no supported diagram.
 */
export function mermaidToSvg(
  source: string | MermaidDiagram,
  options: MermaidRenderOptions = {}
): string {
  // The backdrop, when one was asked for, is already the list's first node. Asking
  // `toSvg` for one as well painted it twice — the second copy from the display list,
  // which is the one every other backend sees.
  return toSvg(drawDiagram("Mermaid.toSvg", source, options));
}
