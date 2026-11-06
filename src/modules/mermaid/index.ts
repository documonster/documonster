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
 *
 * @example A PNG file directly.
 * ```ts
 * import { Mermaid } from "documonster/mermaid";
 *
 * const { png, uncoveredCodePoints } = Mermaid.toPng("flowchart LR\n  A --> B", { scale: 2 });
 * ```
 */

import { checkFlowchart, checkTextBox } from "@mermaid/arguments";
import type { GraphDiagram as MermaidGraphDiagram } from "@mermaid/diagram";
import { layoutFlowchart as layoutFlowchartImpl } from "@mermaid/layout/flowchart";
import type {
  EdgeRoute as MermaidEdgeRoute,
  FlowLayout as MermaidFlowLayout,
  GroupBox as MermaidGroupBox,
  LayoutOptions as MermaidLayoutOptions,
  NodeBox as MermaidNodeBox
} from "@mermaid/layout/flowchart";
import { centeredText as centeredTextImpl } from "@mermaid/render/shared";
import * as MermaidNs from "@mermaid/surface/mermaid";
import type { Theme as MermaidTheme, ThemeOptions as MermaidThemeOptions } from "@mermaid/theme";
import type {
  ClassBox as MermaidClassBox,
  ClassDef as MermaidClassDef,
  ClassDiagram as MermaidClassDiagram,
  ClassLink as MermaidClassLink,
  ClassMember as MermaidClassMember,
  ClassRelation as MermaidClassRelation,
  Cardinality as MermaidCardinality,
  EdgeEnd as MermaidEdgeEnd,
  Entity as MermaidEntity,
  EntityAttribute as MermaidEntityAttribute,
  EntityRelation as MermaidEntityRelation,
  ErDiagram as MermaidErDiagram,
  GanttDiagram as MermaidGanttDiagram,
  GanttTask as MermaidGanttTask,
  JourneyDiagram as MermaidJourneyDiagram,
  JourneyTask as MermaidJourneyTask,
  GitCommit as MermaidGitCommit,
  GitGraphDiagram as MermaidGitGraphDiagram,
  MindNode as MermaidMindNode,
  MindShape as MermaidMindShape,
  MindmapDiagram as MermaidMindmapDiagram,
  ArchitectureDiagram as MermaidArchitectureDiagram,
  ArchitectureEdge as MermaidArchitectureEdge,
  ArchitectureNode as MermaidArchitectureNode,
  BlockCell as MermaidBlockCell,
  BlockDiagram as MermaidBlockDiagram,
  C4Boundary as MermaidC4Boundary,
  C4Diagram as MermaidC4Diagram,
  C4Element as MermaidC4Element,
  C4Relation as MermaidC4Relation,
  KanbanCard as MermaidKanbanCard,
  KanbanColumn as MermaidKanbanColumn,
  KanbanDiagram as MermaidKanbanDiagram,
  PacketDiagram as MermaidPacketDiagram,
  PacketField as MermaidPacketField,
  QuadrantDiagram as MermaidQuadrantDiagram,
  QuadrantPoint as MermaidQuadrantPoint,
  SankeyDiagram as MermaidSankeyDiagram,
  RadarDiagram as MermaidRadarDiagram,
  RadarSeries as MermaidRadarSeries,
  Requirement as MermaidRequirement,
  RequirementDiagram as MermaidRequirementDiagram,
  RequirementElement as MermaidRequirementElement,
  RequirementLink as MermaidRequirementLink,
  SankeyLink as MermaidSankeyLink,
  XyDiagram as MermaidXyDiagram,
  XySeries as MermaidXySeries,
  TimelineDiagram as MermaidTimelineDiagram,
  TimelinePeriod as MermaidTimelinePeriod,
  EdgeStroke as MermaidEdgeStroke,
  FlowDirection as MermaidFlowDirection,
  FlowEdge as MermaidFlowEdge,
  FlowNode as MermaidFlowNode,
  FlowShape as MermaidFlowShape,
  FlowSubgraph as MermaidFlowSubgraph,
  FlowchartDiagram as MermaidFlowchartDiagram,
  PieDiagram as MermaidPieDiagram,
  StateDiagram as MermaidStateDiagram,
  StateNode as MermaidStateNode,
  StateTransition as MermaidStateTransition,
  TaskState as MermaidTaskState,
  Visibility as MermaidVisibility,
  PieSlice as MermaidPieSlice,
  SequenceDiagram as MermaidSequenceDiagram,
  SequenceLine as MermaidSequenceLine,
  SequenceMessage as MermaidSequenceMessage,
  SequenceParticipant as MermaidSequenceParticipant,
  SequenceBlock as MermaidSequenceBlock,
  SequenceEntry as MermaidSequenceEntry,
  SequenceMessageEntry as MermaidSequenceMessageEntry,
  SequenceNote as MermaidSequenceNote,
  SequenceSection as MermaidSequenceSection,
  NotePlacement as MermaidNotePlacement
} from "@mermaid/types";

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
export type { MermaidRenderOptions, MermaidToPngOptions } from "@mermaid/diagram";
export type {
  MermaidGraphDiagram,
  MermaidLayoutOptions,
  MermaidThemeOptions,
  MermaidEdgeRoute,
  MermaidFlowLayout,
  MermaidGroupBox,
  MermaidNodeBox
};
/**
 * Draw a laid-out diagram yourself, in the colours and at the text positions the built-in
 * renderer would have used.
 *
 * `Mermaid.layoutGraph` is public, so "use the layout and render it your own way" is a
 * supported route — a caller wanting a different node shape, an annotation of their own, or
 * an HTML/canvas target rather than a `DrawList` takes it. It was half-open: the geometry
 * came back, and everything needed to *paint* it was internal.
 *
 * {@link MermaidTheme} was published without either the function that produces one or any signature
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
 * is the one thing every backend can honour, so stacking `MermaidNodeBox.lines` in the middle of a
 * box is arithmetic the caller has to do — and handing out the constants invites a ninth copy
 * of that arithmetic, free to disagree with the eight inside. The box shape it takes is
 * exactly what the layout already returns: `MermaidNodeBox`, `MermaidGroupBox` and `MermaidEdgeRoute.label` all
 * carry `x`/`y`/`width`/`height`. `MermaidNodeBox` and `MermaidEdgeRoute.label` carry the wrapped `lines`
 * with them; a `MermaidGroupBox` carries its title as one unwrapped `text`, so pass `[group.text]`.
 */
export type { MermaidTheme };
export type {
  MermaidClassBox,
  MermaidClassDef,
  MermaidClassDiagram,
  MermaidClassLink,
  MermaidClassMember,
  MermaidClassRelation,
  MermaidCardinality,
  MermaidEdgeEnd,
  MermaidEntity,
  MermaidEntityAttribute,
  MermaidEntityRelation,
  MermaidErDiagram,
  MermaidGanttDiagram,
  MermaidGanttTask,
  MermaidJourneyDiagram,
  MermaidJourneyTask,
  MermaidGitCommit,
  MermaidGitGraphDiagram,
  MermaidMindNode,
  MermaidMindShape,
  MermaidMindmapDiagram,
  MermaidArchitectureDiagram,
  MermaidArchitectureEdge,
  MermaidArchitectureNode,
  MermaidBlockCell,
  MermaidBlockDiagram,
  MermaidC4Boundary,
  MermaidC4Diagram,
  MermaidC4Element,
  MermaidC4Relation,
  MermaidKanbanCard,
  MermaidKanbanColumn,
  MermaidKanbanDiagram,
  MermaidPacketDiagram,
  MermaidPacketField,
  MermaidQuadrantDiagram,
  MermaidQuadrantPoint,
  MermaidSankeyDiagram,
  MermaidRadarDiagram,
  MermaidRadarSeries,
  MermaidRequirement,
  MermaidRequirementDiagram,
  MermaidRequirementElement,
  MermaidRequirementLink,
  MermaidSankeyLink,
  MermaidXyDiagram,
  MermaidXySeries,
  MermaidTimelineDiagram,
  MermaidTimelinePeriod,
  MermaidEdgeStroke,
  MermaidFlowDirection,
  MermaidFlowEdge,
  MermaidFlowNode,
  MermaidFlowShape,
  MermaidFlowSubgraph,
  MermaidFlowchartDiagram,
  MermaidPieDiagram,
  MermaidStateDiagram,
  MermaidStateNode,
  MermaidStateTransition,
  MermaidTaskState,
  MermaidVisibility,
  MermaidPieSlice,
  MermaidSequenceDiagram,
  MermaidSequenceLine,
  MermaidSequenceMessage,
  MermaidSequenceParticipant,
  // `SequenceDiagram.body` is part of the public surface, so the types naming its elements
  // have to be too: without these a caller can read `diagram.body` but cannot declare a
  // variable to hold one of its entries.
  MermaidSequenceBlock,
  MermaidSequenceEntry,
  MermaidSequenceMessageEntry,
  MermaidSequenceNote,
  MermaidSequenceSection,
  MermaidNotePlacement
};
export type { MermaidDiagram, MermaidParseReport, MermaidSkippedStatement } from "@mermaid/types";

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
  diagram: MermaidFlowchartDiagram,
  options?: MermaidLayoutOptions
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

// =============================================================================
// Deprecated unprefixed type names — every flat type of this entry carries `Mermaid`
// =============================================================================

/** @deprecated Use {@link MermaidGraphDiagram}. Removed in 1.0. */
export type GraphDiagram = MermaidGraphDiagram;
/** @deprecated Use {@link MermaidEdgeRoute}. Removed in 1.0. */
export type EdgeRoute = MermaidEdgeRoute;
/** @deprecated Use {@link MermaidFlowLayout}. Removed in 1.0. */
export type FlowLayout = MermaidFlowLayout;
/** @deprecated Use {@link MermaidGroupBox}. Removed in 1.0. */
export type GroupBox = MermaidGroupBox;
/** @deprecated Use {@link MermaidLayoutOptions}. Removed in 1.0. */
export type LayoutOptions = MermaidLayoutOptions;
/** @deprecated Use {@link MermaidNodeBox}. Removed in 1.0. */
export type NodeBox = MermaidNodeBox;
/** @deprecated Use {@link MermaidTheme}. Removed in 1.0. */
export type Theme = MermaidTheme;
/** @deprecated Use {@link MermaidThemeOptions}. Removed in 1.0. */
export type ThemeOptions = MermaidThemeOptions;
/** @deprecated Use {@link MermaidClassBox}. Removed in 1.0. */
export type ClassBox = MermaidClassBox;
/** @deprecated Use {@link MermaidClassDef}. Removed in 1.0. */
export type ClassDef = MermaidClassDef;
/** @deprecated Use {@link MermaidClassDiagram}. Removed in 1.0. */
export type ClassDiagram = MermaidClassDiagram;
/** @deprecated Use {@link MermaidClassLink}. Removed in 1.0. */
export type ClassLink = MermaidClassLink;
/** @deprecated Use {@link MermaidClassMember}. Removed in 1.0. */
export type ClassMember = MermaidClassMember;
/** @deprecated Use {@link MermaidClassRelation}. Removed in 1.0. */
export type ClassRelation = MermaidClassRelation;
/** @deprecated Use {@link MermaidCardinality}. Removed in 1.0. */
export type Cardinality = MermaidCardinality;
/** @deprecated Use {@link MermaidEdgeEnd}. Removed in 1.0. */
export type EdgeEnd = MermaidEdgeEnd;
/** @deprecated Use {@link MermaidEntity}. Removed in 1.0. */
export type Entity = MermaidEntity;
/** @deprecated Use {@link MermaidEntityAttribute}. Removed in 1.0. */
export type EntityAttribute = MermaidEntityAttribute;
/** @deprecated Use {@link MermaidEntityRelation}. Removed in 1.0. */
export type EntityRelation = MermaidEntityRelation;
/** @deprecated Use {@link MermaidErDiagram}. Removed in 1.0. */
export type ErDiagram = MermaidErDiagram;
/** @deprecated Use {@link MermaidGanttDiagram}. Removed in 1.0. */
export type GanttDiagram = MermaidGanttDiagram;
/** @deprecated Use {@link MermaidGanttTask}. Removed in 1.0. */
export type GanttTask = MermaidGanttTask;
/** @deprecated Use {@link MermaidJourneyDiagram}. Removed in 1.0. */
export type JourneyDiagram = MermaidJourneyDiagram;
/** @deprecated Use {@link MermaidJourneyTask}. Removed in 1.0. */
export type JourneyTask = MermaidJourneyTask;
/** @deprecated Use {@link MermaidGitCommit}. Removed in 1.0. */
export type GitCommit = MermaidGitCommit;
/** @deprecated Use {@link MermaidGitGraphDiagram}. Removed in 1.0. */
export type GitGraphDiagram = MermaidGitGraphDiagram;
/** @deprecated Use {@link MermaidMindNode}. Removed in 1.0. */
export type MindNode = MermaidMindNode;
/** @deprecated Use {@link MermaidMindShape}. Removed in 1.0. */
export type MindShape = MermaidMindShape;
/** @deprecated Use {@link MermaidMindmapDiagram}. Removed in 1.0. */
export type MindmapDiagram = MermaidMindmapDiagram;
/** @deprecated Use {@link MermaidArchitectureDiagram}. Removed in 1.0. */
export type ArchitectureDiagram = MermaidArchitectureDiagram;
/** @deprecated Use {@link MermaidArchitectureEdge}. Removed in 1.0. */
export type ArchitectureEdge = MermaidArchitectureEdge;
/** @deprecated Use {@link MermaidArchitectureNode}. Removed in 1.0. */
export type ArchitectureNode = MermaidArchitectureNode;
/** @deprecated Use {@link MermaidBlockCell}. Removed in 1.0. */
export type BlockCell = MermaidBlockCell;
/** @deprecated Use {@link MermaidBlockDiagram}. Removed in 1.0. */
export type BlockDiagram = MermaidBlockDiagram;
/** @deprecated Use {@link MermaidC4Boundary}. Removed in 1.0. */
export type C4Boundary = MermaidC4Boundary;
/** @deprecated Use {@link MermaidC4Diagram}. Removed in 1.0. */
export type C4Diagram = MermaidC4Diagram;
/** @deprecated Use {@link MermaidC4Element}. Removed in 1.0. */
export type C4Element = MermaidC4Element;
/** @deprecated Use {@link MermaidC4Relation}. Removed in 1.0. */
export type C4Relation = MermaidC4Relation;
/** @deprecated Use {@link MermaidKanbanCard}. Removed in 1.0. */
export type KanbanCard = MermaidKanbanCard;
/** @deprecated Use {@link MermaidKanbanColumn}. Removed in 1.0. */
export type KanbanColumn = MermaidKanbanColumn;
/** @deprecated Use {@link MermaidKanbanDiagram}. Removed in 1.0. */
export type KanbanDiagram = MermaidKanbanDiagram;
/** @deprecated Use {@link MermaidPacketDiagram}. Removed in 1.0. */
export type PacketDiagram = MermaidPacketDiagram;
/** @deprecated Use {@link MermaidPacketField}. Removed in 1.0. */
export type PacketField = MermaidPacketField;
/** @deprecated Use {@link MermaidQuadrantDiagram}. Removed in 1.0. */
export type QuadrantDiagram = MermaidQuadrantDiagram;
/** @deprecated Use {@link MermaidQuadrantPoint}. Removed in 1.0. */
export type QuadrantPoint = MermaidQuadrantPoint;
/** @deprecated Use {@link MermaidSankeyDiagram}. Removed in 1.0. */
export type SankeyDiagram = MermaidSankeyDiagram;
/** @deprecated Use {@link MermaidRadarDiagram}. Removed in 1.0. */
export type RadarDiagram = MermaidRadarDiagram;
/** @deprecated Use {@link MermaidRadarSeries}. Removed in 1.0. */
export type RadarSeries = MermaidRadarSeries;
/** @deprecated Use {@link MermaidRequirement}. Removed in 1.0. */
export type Requirement = MermaidRequirement;
/** @deprecated Use {@link MermaidRequirementDiagram}. Removed in 1.0. */
export type RequirementDiagram = MermaidRequirementDiagram;
/** @deprecated Use {@link MermaidRequirementElement}. Removed in 1.0. */
export type RequirementElement = MermaidRequirementElement;
/** @deprecated Use {@link MermaidRequirementLink}. Removed in 1.0. */
export type RequirementLink = MermaidRequirementLink;
/** @deprecated Use {@link MermaidSankeyLink}. Removed in 1.0. */
export type SankeyLink = MermaidSankeyLink;
/** @deprecated Use {@link MermaidXyDiagram}. Removed in 1.0. */
export type XyDiagram = MermaidXyDiagram;
/** @deprecated Use {@link MermaidXySeries}. Removed in 1.0. */
export type XySeries = MermaidXySeries;
/** @deprecated Use {@link MermaidTimelineDiagram}. Removed in 1.0. */
export type TimelineDiagram = MermaidTimelineDiagram;
/** @deprecated Use {@link MermaidTimelinePeriod}. Removed in 1.0. */
export type TimelinePeriod = MermaidTimelinePeriod;
/** @deprecated Use {@link MermaidEdgeStroke}. Removed in 1.0. */
export type EdgeStroke = MermaidEdgeStroke;
/** @deprecated Use {@link MermaidFlowDirection}. Removed in 1.0. */
export type FlowDirection = MermaidFlowDirection;
/** @deprecated Use {@link MermaidFlowEdge}. Removed in 1.0. */
export type FlowEdge = MermaidFlowEdge;
/** @deprecated Use {@link MermaidFlowNode}. Removed in 1.0. */
export type FlowNode = MermaidFlowNode;
/** @deprecated Use {@link MermaidFlowShape}. Removed in 1.0. */
export type FlowShape = MermaidFlowShape;
/** @deprecated Use {@link MermaidFlowSubgraph}. Removed in 1.0. */
export type FlowSubgraph = MermaidFlowSubgraph;
/** @deprecated Use {@link MermaidFlowchartDiagram}. Removed in 1.0. */
export type FlowchartDiagram = MermaidFlowchartDiagram;
/** @deprecated Use {@link MermaidPieDiagram}. Removed in 1.0. */
export type PieDiagram = MermaidPieDiagram;
/** @deprecated Use {@link MermaidStateDiagram}. Removed in 1.0. */
export type StateDiagram = MermaidStateDiagram;
/** @deprecated Use {@link MermaidStateNode}. Removed in 1.0. */
export type StateNode = MermaidStateNode;
/** @deprecated Use {@link MermaidStateTransition}. Removed in 1.0. */
export type StateTransition = MermaidStateTransition;
/** @deprecated Use {@link MermaidTaskState}. Removed in 1.0. */
export type TaskState = MermaidTaskState;
/** @deprecated Use {@link MermaidVisibility}. Removed in 1.0. */
export type Visibility = MermaidVisibility;
/** @deprecated Use {@link MermaidPieSlice}. Removed in 1.0. */
export type PieSlice = MermaidPieSlice;
/** @deprecated Use {@link MermaidSequenceDiagram}. Removed in 1.0. */
export type SequenceDiagram = MermaidSequenceDiagram;
/** @deprecated Use {@link MermaidSequenceLine}. Removed in 1.0. */
export type SequenceLine = MermaidSequenceLine;
/** @deprecated Use {@link MermaidSequenceMessage}. Removed in 1.0. */
export type SequenceMessage = MermaidSequenceMessage;
/** @deprecated Use {@link MermaidSequenceParticipant}. Removed in 1.0. */
export type SequenceParticipant = MermaidSequenceParticipant;
/** @deprecated Use {@link MermaidSequenceBlock}. Removed in 1.0. */
export type SequenceBlock = MermaidSequenceBlock;
/** @deprecated Use {@link MermaidSequenceEntry}. Removed in 1.0. */
export type SequenceEntry = MermaidSequenceEntry;
/** @deprecated Use {@link MermaidSequenceMessageEntry}. Removed in 1.0. */
export type SequenceMessageEntry = MermaidSequenceMessageEntry;
/** @deprecated Use {@link MermaidSequenceNote}. Removed in 1.0. */
export type SequenceNote = MermaidSequenceNote;
/** @deprecated Use {@link MermaidSequenceSection}. Removed in 1.0. */
export type SequenceSection = MermaidSequenceSection;
/** @deprecated Use {@link MermaidNotePlacement}. Removed in 1.0. */
export type NotePlacement = MermaidNotePlacement;

// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";
