/**
 * `mermaid` — Mermaid diagram text to a drawable display list.
 *
 * The module is a *producer*: it turns a diagram into a {@link DrawList} and stops there.
 * Nothing here knows what SVG or a PDF page is, and that is what lets one implementation
 * serve every output — `documonster/draw` renders the list to markup or pixels, and
 * `documonster/pdf` renders it onto a page.
 *
 * Twenty-one diagram types are drawn: `flowchart`, `stateDiagram`, `classDiagram`,
 * `erDiagram`, `sequenceDiagram`, `requirementDiagram`, `C4`, `architecture`, `gantt`,
 * `timeline`, `journey`, `kanban`, `mindmap`, `gitGraph`, `quadrantChart`, `xychart`,
 * `radar`, `sankey`, `packet`, `block` and `pie`.
 *
 * @example
 * ```ts
 * import { Mermaid } from "documonster/mermaid";
 *
 * const svg = Mermaid.toSvg(`flowchart TD
 *   A[Start] --> B{Ready?}
 *   B -->|yes| C[Ship it]
 *   B -->|no| A`);
 * ```
 *
 * @example Any backend, from the same list.
 * ```ts
 * import { Mermaid } from "documonster/mermaid";
 * import { Draw } from "documonster/draw";
 *
 * const list = Mermaid.toDrawList("pie title Split\n \"A\" : 60\n \"B\" : 40");
 * const svg = Draw.toSvg(list);
 * const pixels = Draw.toRgba(list, { scale: 2 });
 * ```
 */

import { checkFlowchart, checkTextBox } from "@mermaid/arguments";
import { layoutFlowchart as layoutFlowchartImpl } from "@mermaid/layout/flowchart";
import type { LayoutOptions as FlowLayoutOptions } from "@mermaid/layout/flowchart";
import { centeredText as centeredTextImpl } from "@mermaid/render/shared";
import * as MermaidNs from "@mermaid/surface/mermaid";
import type { FlowchartDiagram as FlowchartDiagramType } from "@mermaid/types";

// =============================================================================
// Public API — the `Mermaid` namespace (tree-shaken via `export * as`)
// =============================================================================

export * as Mermaid from "@mermaid/surface/mermaid";

export {
  MermaidError,
  MermaidArgumentError,
  MermaidParseError,
  MermaidSyntaxError,
  isMermaidError
} from "@mermaid/errors";
export type { MermaidParseErrorOptions } from "@mermaid/errors";
export type { GraphDiagram, MermaidRenderOptions } from "@mermaid/diagram";
export type { LayoutOptions } from "@mermaid/layout/flowchart";
export type { ThemeOptions } from "@mermaid/theme";
export type { EdgeRoute, FlowLayout, GroupBox, NodeBox } from "@mermaid/layout/flowchart";
/**
 * Draw a laid-out diagram yourself, in the colours and at the text positions the built-in
 * renderer would have used.
 *
 * `Mermaid.layoutGraph` is public, so "use the layout and render it your own way" is a
 * supported route — a caller wanting a different node shape, an annotation of their own, or
 * an HTML/canvas target rather than a `DrawList` takes it. It was half-open: the geometry
 * came back, and everything needed to *paint* it was internal.
 *
 * {@link Theme} was published without either the function that produces one or any signature
 * accepting one, which made it a name for a thing a consumer could not obtain, could not
 * construct — ten required `Rgba01` fields plus a `paletteText` callback — and could not pass
 * anywhere. `Mermaid.resolveTheme` is what it was missing: `MermaidRenderOptions.theme` takes
 * CSS strings, and this is the same resolution the renderer applies to them, so a
 * self-rendered diagram matches one from `Mermaid.toSvg` instead of approximating it.
 * Without it the only route to the palette was to copy the hex literals out of this module's
 * source and hope they did not move.
 *
 * `Mermaid.createCenteredText` is exported rather than the two constants behind it (`LINE_HEIGHT`,
 * `BASELINE_SHIFT`) on purpose. A display list positions text by its baseline, because that
 * is the one thing every backend can honour, so stacking `NodeBox.lines` in the middle of a
 * box is arithmetic the caller has to do — and handing out the constants invites a ninth copy
 * of that arithmetic, free to disagree with the eight inside. The box shape it takes is
 * exactly what the layout already returns: `NodeBox`, `GroupBox` and `EdgeRoute.label` all
 * carry `x`/`y`/`width`/`height`. `NodeBox` and `EdgeRoute.label` carry the wrapped `lines`
 * with them; a `GroupBox` carries its title as one unwrapped `text`, so pass `[group.text]`.
 */
export type { Theme } from "@mermaid/theme";
export type {
  ClassBox,
  ClassDef,
  ClassDiagram,
  ClassLink,
  ClassMember,
  ClassRelation,
  Cardinality,
  EdgeEnd,
  Entity,
  EntityAttribute,
  EntityRelation,
  ErDiagram,
  GanttDiagram,
  GanttTask,
  JourneyDiagram,
  JourneyTask,
  GitCommit,
  GitGraphDiagram,
  MindNode,
  MindShape,
  MindmapDiagram,
  ArchitectureDiagram,
  ArchitectureEdge,
  ArchitectureNode,
  BlockCell,
  BlockDiagram,
  C4Boundary,
  C4Diagram,
  C4Element,
  C4Relation,
  KanbanCard,
  KanbanColumn,
  KanbanDiagram,
  PacketDiagram,
  PacketField,
  QuadrantDiagram,
  QuadrantPoint,
  SankeyDiagram,
  RadarDiagram,
  RadarSeries,
  Requirement,
  RequirementDiagram,
  RequirementElement,
  RequirementLink,
  SankeyLink,
  XyDiagram,
  XySeries,
  TimelineDiagram,
  TimelinePeriod,
  EdgeStroke,
  FlowDirection,
  FlowEdge,
  FlowNode,
  FlowShape,
  FlowSubgraph,
  FlowchartDiagram,
  MermaidDiagram,
  MermaidParseReport,
  MermaidSkippedStatement,
  PieDiagram,
  StateDiagram,
  StateNode,
  StateTransition,
  TaskState,
  Visibility,
  PieSlice,
  SequenceDiagram,
  SequenceLine,
  SequenceMessage,
  SequenceParticipant,
  // `SequenceDiagram.body` is part of the public surface, so the types naming its elements
  // have to be too: without these a caller can read `diagram.body` but cannot declare a
  // variable to hold one of its entries.
  SequenceBlock,
  SequenceEntry,
  SequenceMessageEntry,
  SequenceNote,
  SequenceSection,
  NotePlacement
} from "@mermaid/types";

// =============================================================================
// Deprecated flat names — forwards to the `Mermaid` namespace
// =============================================================================

/** @deprecated Use {@link MermaidNs.toDrawList | Mermaid.toDrawList}. Removed in 1.0. */
export const mermaidToDrawList = MermaidNs.toDrawList;
/** @deprecated Use {@link MermaidNs.toSvg | Mermaid.toSvg}. Removed in 1.0. */
export const mermaidToSvg = MermaidNs.toSvg;
/** @deprecated Use {@link MermaidNs.parse | Mermaid.parse}. Removed in 1.0. */
export const parseMermaid = MermaidNs.parse;
/**
 * @deprecated Use {@link MermaidNs.layoutGraph | Mermaid.layoutGraph}, which also lays out
 * state, class, ER, requirement, C4 and architecture diagrams. Removed in 1.0.
 */
export function layoutFlowchart(
  diagram: FlowchartDiagramType,
  options?: FlowLayoutOptions
): ReturnType<typeof layoutFlowchartImpl> {
  checkFlowchart("layoutFlowchart", diagram);
  return layoutFlowchartImpl(diagram, options);
}
/**
 * @deprecated Use {@link MermaidNs.createCenteredText | Mermaid.createCenteredText}, which
 * takes the font as one style object. Removed in 1.0.
 */
export function centredText(
  ...args: Parameters<typeof centeredTextImpl>
): ReturnType<typeof centeredTextImpl> {
  checkTextBox("centredText", args[0], args[1]);
  return centeredTextImpl(...args);
}
/** @deprecated Use {@link MermaidNs.resolveTheme | Mermaid.resolveTheme}. Removed in 1.0. */
export const resolveTheme = MermaidNs.resolveTheme;

// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";
