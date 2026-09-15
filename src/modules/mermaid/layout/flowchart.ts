/**
 * Flowchart layout.
 *
 * A layered ("Sugiyama") layout, which is what every flowchart tool converges on because
 * it matches how the diagram is read: edges point one way, so nodes should advance one
 * way. Four passes, each fixing one thing the next depends on:
 *
 * 1. **Rank** — how far along the flow each node sits. Longest-path over the acyclic part
 *    of the graph, so an edge always spans at least one rank forwards.
 * 2. **Order** — the sequence within a rank, chosen to reduce edge crossings. A crossing
 *    is the single thing that makes a diagram unreadable, and the median heuristic
 *    removes most of them for a fraction of the cost of doing it exactly (which is
 *    NP-hard).
 * 3. **Position** — real coordinates, with a priority pass that pulls each node towards
 *    the average of its neighbours so chains come out straight instead of stepped.
 * 4. **Route** — the points an edge follows, including the ranks it passes through.
 *
 * Cycles are handled by ignoring the back edge when ranking — it still gets drawn, it
 * just does not get a say in the ordering. A graph with a cycle has no layering that
 * satisfies every edge, so something has to give, and the alternative (reversing the edge
 * and drawing it upside down) hides from the reader that the cycle exists.
 */

import { measureText, wrapText } from "@draw/text";
import type { DrawTextStyle } from "@draw/types";
import { LINE_HEIGHT } from "@mermaid/theme";
import type { FlowEdge, FlowNode, FlowShape, FlowchartDiagram } from "@mermaid/types";

/** A laid-out node, in the diagram's own Y-down space. */
export interface NodeBox {
  readonly id: string;
  readonly shape: FlowShape;
  readonly classes: readonly string[];
  /** Wrapped label lines. */
  readonly lines: readonly string[];
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/** A laid-out edge: the polyline it follows, plus where its label sits. */
export interface EdgeRoute {
  readonly edge: FlowEdge;
  readonly points: readonly { x: number; y: number }[];
  readonly label?: {
    text: string;
    lines: readonly string[];
    x: number;
    y: number;
    width: number;
    height: number;
  };
}

/** A laid-out subgraph frame. */
export interface GroupBox {
  readonly id: string;
  readonly text: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/** The whole diagram, positioned. */
export interface FlowLayout {
  readonly width: number;
  readonly height: number;
  readonly nodes: readonly NodeBox[];
  readonly edges: readonly EdgeRoute[];
  readonly groups: readonly GroupBox[];
  readonly title?: { text: string; x: number; y: number };
}

/** Tunable geometry. Every number is a diagram-space unit, which the SVG maps 1:1. */
export interface LayoutOptions {
  readonly fontSize?: number;
  readonly fontFamily?: string;
  /** Gap between adjacent ranks, along the flow. */
  readonly rankGap?: number;
  /** Gap between siblings within a rank, across the flow. */
  readonly nodeGap?: number;
  /** Longest line of a node's own text before it wraps. */
  readonly maxLabelWidth?: number;
  /**
   * Longest line of an *edge* label before it wraps. Defaults to {@link maxLabelWidth}.
   *
   * Separate because the two tolerate wrapping quite differently, and one number for both
   * produced a bad picture whenever a diagram had to be narrowed to fit a page: an edge label
   * is loose text in the space between ranks and reads fine over three or four short lines,
   * whereas a node's text is the thing the box is drawn around. Wrapping a ten-word node label
   * at seventy points turns it into a column one or two words wide and a dozen lines deep — the
   * diagram is then mostly empty box, and twice as tall as it needs to be.
   */
  readonly maxEdgeLabelWidth?: number;
  readonly padding?: number;
  /**
   * Size a node yourself.
   *
   * A class box is three compartments, not a centred label, so it knows its own size in a
   * way this module's measuring pass cannot. Supplying it lets a diagram type with its own
   * box design still use the ranking, ordering, straightening and routing here, which is
   * where the difficult parts of a layout live.
   */
  readonly measureNode?: (
    node: FlowNode
  ) => { width: number; height: number; lines: readonly string[] } | undefined;
}

/**
 * The options after defaults have been applied.
 *
 * Spelled out rather than `ResolvedOptions`, because `measureNode` is genuinely
 * optional — there is no default measurement, only the built-in one — and `Required` would
 * demand a callback nobody can supply.
 */
type ResolvedOptions = Required<Omit<LayoutOptions, "measureNode">> &
  Pick<LayoutOptions, "measureNode">;

const DEFAULTS = {
  fontSize: 14,
  fontFamily: "Arial",
  rankGap: 56,
  nodeGap: 34,
  maxLabelWidth: 220,
  maxEdgeLabelWidth: 220,
  padding: 16
} as const;

/** Space inside a node, around its label. */
const NODE_PAD_X = 16;
const NODE_PAD_Y = 11;
/** Extra width a slanted or pointed outline needs so its label still fits inside. */
const SHAPE_SLACK: Partial<Record<FlowShape, { x: number; y: number }>> = {
  rhombus: { x: 34, y: 20 },
  hexagon: { x: 20, y: 0 },
  parallelogram: { x: 18, y: 0 },
  parallelogramAlt: { x: 18, y: 0 },
  trapezoid: { x: 26, y: 0 },
  trapezoidAlt: { x: 26, y: 0 },
  asymmetric: { x: 12, y: 0 },
  cylinder: { x: 0, y: 12 },
  subroutine: { x: 14, y: 0 },
  stadium: { x: 12, y: 0 }
};

/** Lay out a flowchart. */
export function layoutFlowchart(
  diagram: FlowchartDiagram,
  options: LayoutOptions = {}
): FlowLayout {
  const base = { ...DEFAULTS, ...stripUndefined(options) };
  // An author who narrows labels without saying which meant both.
  const config = {
    ...base,
    maxEdgeLabelWidth: options.maxEdgeLabelWidth ?? base.maxLabelWidth
  };
  const style: DrawTextStyle = { size: config.fontSize, family: config.fontFamily };
  const labelStyle: DrawTextStyle = { size: config.fontSize - 2, family: config.fontFamily };

  const boxes = measureNodes(diagram, config, style);
  if (boxes.size === 0) {
    return {
      width: config.padding * 2,
      height: config.padding * 2,
      nodes: [],
      edges: [],
      groups: []
    };
  }

  // An edge may name a subgraph rather than a node. Ranking and ordering see it as an edge
  // to every member, which is what puts a group below the node that points at it; routing
  // draws one line, from the frame.
  const clusters = clusterMembers(diagram, boxes);
  const expanded = expandClusterEdges(diagram.edges, clusters);
  const graphEdges = expanded.edges;

  const ranks = assignRanks(boxes, graphEdges);
  // Long edges become chains of dummies, one per rank they pass through. Without them an
  // edge spanning three ranks has no say in how those ranks are ordered, so it is left to
  // find its own way afterwards — around the outside, which is why a diagram with several
  // skip edges grew a bundle of lanes down its margin. With them the edge is ordered like
  // any other and drawn straight through the gap that was left for it.
  const horizontal = diagram.direction === "LR" || diagram.direction === "RL";
  const chains = insertDummies(graphEdges, boxes, ranks, config, labelStyle, horizontal);
  aliasExpandedRoutes(chains, expanded.representative);
  const orders = orderRanks(diagram, chains.edges, ranks, boxes);
  const neighbours = adjacency(chains.edges);
  const spans = labelSpans(graphEdges, boxes, labelStyle, config.maxEdgeLabelWidth);
  place(
    boxes,
    orders,
    neighbours,
    config,
    horizontal,
    chains.dummies,
    spans,
    rankGaps(diagram, boxes, orders.length, config)
  );

  evictOutsiders(diagram, boxes, orders, config, horizontal);

  // Across positions are settled, so the lanes the labels need can be measured rather than
  // guessed — and the bands recomputed with room for them. Only the flow axis moves.
  // Lanes belong to visible edges, not to the member-edge copies used only for
  // ranking a cluster. Assigning them on `graphEdges` turned one labelled edge
  // from a three-node subgraph into three labels and reserved lanes for text
  // that would never be rendered, stretching five `alignment` links over two
  // horizontal trunks. Frames already have their stable across-axis bounds;
  // only the along-axis bands move below.
  const preliminaryGroups = frameGroups(diagram, boxes, config);
  const routeBoxes = withClusterBoxes(boxes, preliminaryGroups, clusters);
  const lanes = assignLanes(diagram, routeBoxes, labelStyle, config, horizontal, orders.length);
  assignBands(
    boxes,
    orders,
    config,
    horizontal,
    rankGaps(diagram, boxes, orders.length, config, lanes)
  );

  // Dummies are scaffolding: they steer the layout and then leave, handing their positions
  // to the edge that put them there.
  const nodes = [...boxes.values()].filter(box => !chains.dummies.has(box.id));
  const groups = frameGroups(diagram, boxes, config);
  const routes = routeEdges(
    diagram,
    withClusterBoxes(boxes, groups, clusters),
    new Set(clusters.keys()),
    labelStyle,
    config,
    chains.waypoints,
    lanes,
    chains.labelled
  );

  return finish(diagram, nodes, routes, groups, config, style);
}

/**
 * Subgraphs an edge may legitimately name, and the members that stand in for them.
 *
 * An id that is *also* a real node is not one of these: the author wrote a node and a group
 * with the same name, and the node is what an edge to that name means.
 */
function clusterMembers(
  diagram: FlowchartDiagram,
  boxes: ReadonlyMap<string, Mutable>
): Map<string, string[]> {
  const out = new Map<string, string[]>();
  for (const group of diagram.subgraphs) {
    if (boxes.has(group.id)) {
      continue;
    }
    const members = group.nodeIds.filter(id => boxes.has(id));
    if (members.length > 0) {
      out.set(group.id, members);
    }
  }
  return out;
}

/**
 * Replace each cluster endpoint with the members it stands for.
 *
 * Ranking works on nodes, and a group has no rank of its own — it is wherever its members
 * ended up. So `DM --> ALIGNMENT` has to reach the ranker as `DM --> B`, `DM --> T`,
 * `DM --> V`, or the group's members keep whatever rank their own edges gave them: in a
 * diagram whose grouped nodes had no other edges, that was rank 0, beside the node pointing
 * at them rather than below it.
 *
 * A non-cluster edge is passed through *by identity*, because `waypoints` and the sibling
 * fan are keyed on the edge object.
 */
interface ExpandedEdges {
  readonly edges: readonly FlowEdge[];
  /** One expanded edge whose route stands in for each original visible edge. */
  readonly representative: ReadonlyMap<FlowEdge, FlowEdge>;
}

function expandClusterEdges(
  edges: readonly FlowEdge[],
  clusters: ReadonlyMap<string, readonly string[]>
): ExpandedEdges {
  if (clusters.size === 0) {
    return { edges: [...edges], representative: new Map(edges.map(edge => [edge, edge])) };
  }
  const out: FlowEdge[] = [];
  const representative = new Map<FlowEdge, FlowEdge>();
  for (const edge of edges) {
    const froms = clusters.get(edge.from) ?? [edge.from];
    const tos = clusters.get(edge.to) ?? [edge.to];
    const expanded: FlowEdge[] = [];
    for (const from of froms) {
      for (const to of tos) {
        if (from === to) {
          continue;
        }
        const copy = from === edge.from && to === edge.to ? edge : { ...edge, from, to };
        out.push(copy);
        expanded.push(copy);
      }
    }
    const middle = expanded[Math.floor((expanded.length - 1) / 2)];
    if (middle !== undefined) {
      representative.set(edge, middle);
    }
  }
  return { edges: out, representative };
}

/** Carry routing scaffolding from the representative expanded edge back to the visible edge. */
function aliasExpandedRoutes(
  chains: Chains,
  representative: ReadonlyMap<FlowEdge, FlowEdge>
): void {
  for (const [original, expanded] of representative) {
    if (original === expanded) {
      continue;
    }
    const waypoints = chains.waypoints.get(expanded);
    if (waypoints !== undefined) {
      chains.waypoints.set(original, waypoints);
    }
    const label = chains.labelled.get(expanded);
    if (label !== undefined) {
      chains.labelled.set(original, label);
    }
  }
}

/** Carry the representative expanded edge's lane back to the visible cluster edge. */
/**
 * The node map the router sees, with a stand-in box for every cluster an edge names.
 *
 * The frame is only known once its members are placed, which is why this is built here and
 * not with the rest of the boxes. A cluster's rank is its members' first, so an edge into it
 * still reads as pointing forwards.
 */
function withClusterBoxes(
  boxes: Map<string, Mutable>,
  groups: readonly GroupBox[],
  clusters: ReadonlyMap<string, readonly string[]>
): Map<string, Mutable> {
  if (clusters.size === 0) {
    return boxes;
  }
  const merged = new Map(boxes);
  for (const group of groups) {
    const members = clusters.get(group.id);
    if (members === undefined || merged.has(group.id)) {
      continue;
    }
    const ranks = members
      .map(id => boxes.get(id)?.rank)
      .filter((r): r is number => r !== undefined);
    merged.set(group.id, {
      id: group.id,
      shape: "rect",
      classes: [],
      lines: [],
      x: group.x,
      y: group.y,
      width: group.width,
      height: group.height,
      rank: ranks.length === 0 ? 0 : Math.min(...ranks),
      order: 0
    });
  }
  return merged;
}

/** The across-axis extent of an edge's label, measured the way the router will measure it. */
function labelWidthOf(edge: FlowEdge, labelStyle: DrawTextStyle, maxLabelWidth: number): number {
  if (edge.label === undefined || edge.label === "") {
    return 0;
  }
  return wrapText(edge.label, labelStyle, maxLabelWidth).reduce(
    (max, line) => Math.max(max, measureText(line, labelStyle)),
    0
  );
}

/**
 * Panel padding plus daylight between two neighbouring edge labels.
 *
 * The renderer draws a label on a panel four units wider than its text on each side, so two
 * panels touch as soon as their texts are sixteen units apart; the rest is the gap that
 * makes them read as two labels rather than one band.
 */
const LABEL_CLEARANCE = 24;

/** Clear space between one lane's labels and the next's. */
const LANE_DAYLIGHT = 8;

/**
 * How much room each node has to reserve across the axis for the labels on its own links.
 *
 * A label sits at the midpoint of its edge, and that is the whole problem: for a fan out of
 * one parent the label of `parent → child` lands at `(parent + child) / 2`, so the labels
 * spread only *half* as far as the children do. Five children 220 units apart therefore
 * carry five labels 110 units apart, and a label 230 units wide overlaps both its
 * neighbours — which is what turned a row of five identical link labels into one unreadable
 * grey band whatever the node gap was set to.
 *
 * dagre answers this by giving every label a rank and a width of its own, so the ordering
 * pass separates labels exactly as it separates nodes. The same guarantee is expressed here
 * as a separation constraint between the nodes at the ends, which leaves the routing —
 * ports, sibling fans, detours — untouched.
 *
 * Only links between *adjacent* ranks count. One that spans further is threaded through
 * dummies and carries its label wherever the lane went, and one that runs backwards is
 * routed around the outside where nothing is in the way.
 */
function labelSpans(
  graphEdges: readonly FlowEdge[],
  boxes: ReadonlyMap<string, Mutable>,
  labelStyle: DrawTextStyle,
  maxLabelWidth: number
): Map<string, number> {
  const spans = new Map<string, number>();
  for (const edge of graphEdges) {
    const from = boxes.get(edge.from);
    const to = boxes.get(edge.to);
    if (!from || !to || from === to || to.rank !== from.rank + 1) {
      continue;
    }
    const width = labelWidthOf(edge, labelStyle, maxLabelWidth);
    if (width === 0) {
      continue;
    }
    // Both ends, because two labelled links arriving at one node from two parents are
    // separated by how far apart the *parents* sit, by the same halving.
    for (const id of [edge.from, edge.to]) {
      spans.set(id, Math.max(spans.get(id) ?? 0, width));
    }
  }
  return spans;
}

/** A node while it is being laid out. */
interface Mutable extends Omit<NodeBox, "x" | "y"> {
  x: number;
  y: number;
  rank: number;
  order: number;
}

function stripUndefined<T extends object>(value: T): Partial<T> {
  return Object.fromEntries(Object.entries(value).filter(([, v]) => v !== undefined)) as Partial<T>;
}

/** Wrap each label and size the box that has to contain it. */
function measureNodes(
  diagram: FlowchartDiagram,
  config: ResolvedOptions,
  style: DrawTextStyle
): Map<string, Mutable> {
  const out = new Map<string, Mutable>();
  for (const node of diagram.nodes) {
    const supplied = config.measureNode?.(node);
    if (supplied) {
      out.set(node.id, {
        id: node.id,
        shape: node.shape,
        classes: node.classes,
        lines: supplied.lines,
        x: 0,
        y: 0,
        width: supplied.width,
        height: supplied.height,
        rank: 0,
        order: 0
      });
      continue;
    }
    const lines = wrapText(node.text, style, config.maxLabelWidth);
    const textWidth = lines.reduce((max, line) => Math.max(max, measureText(line, style)), 0);
    const slack = SHAPE_SLACK[node.shape] ?? { x: 0, y: 0 };
    let width = textWidth + NODE_PAD_X * 2 + slack.x;
    let height = lines.length * config.fontSize * LINE_HEIGHT + NODE_PAD_Y * 2 + slack.y;
    // A start/end marker is a fixed-size dot: the minimums below would inflate it to a box.
    const isMarker = node.shape === "stateStart" || node.shape === "stateEnd";
    if (isMarker) {
      // A marker is a fixed disc: it carries no text, so nothing about the label should
      // decide how big it is.
      width = node.shape === "stateStart" ? 20 : 26;
      height = width;
    } else if (node.shape === "circle" || node.shape === "doubleCircle") {
      // A circle has to contain the label's diagonal, not its width.
      const diameter = Math.hypot(textWidth, height) + NODE_PAD_X;
      width = diameter;
      height = diameter;
    }
    out.set(node.id, {
      id: node.id,
      shape: node.shape,
      classes: node.classes,
      lines,
      x: 0,
      y: 0,
      width: isMarker ? width : Math.max(width, 36),
      height: isMarker ? height : Math.max(height, 30),
      rank: 0,
      order: 0
    });
  }
  return out;
}

/** What the dummy pass produced. */
interface Chains {
  /** The edge list the ordering should see: long edges replaced by their segments. */
  readonly edges: readonly FlowEdge[];
  /** Ids that exist only to steer the layout. */
  readonly dummies: ReadonlySet<string>;
  /** Per original edge, the ranks it passes through, in order. */
  readonly waypoints: Map<FlowEdge, readonly string[]>;
  /** For a threaded edge that carries a label, the dummy standing in for that label. */
  readonly labelled: Map<FlowEdge, string>;
}

/**
 * Replace every forward edge that skips a rank with a chain through it.
 *
 * This is the piece a layered layout needs to place long edges well, and the reason
 * `dagre` and every other implementation does it: an edge that is not represented in the
 * intermediate ranks cannot be ordered against the nodes there, so it has to be routed
 * afterwards around whatever ended up in the way.
 *
 * A dummy is narrow — it is a place for a line to pass through, not a box — so it costs
 * almost nothing in width while buying the edge a lane of its own.
 */
function insertDummies(
  graphEdges: readonly FlowEdge[],
  boxes: Map<string, Mutable>,
  layers: string[][],
  config: ResolvedOptions,
  labelStyle: DrawTextStyle,
  horizontal: boolean
): Chains {
  const edges: FlowEdge[] = [];
  const dummies = new Set<string>();
  const waypoints = new Map<FlowEdge, readonly string[]>();
  const labelled = new Map<FlowEdge, string>();
  let serial = 0;

  for (const edge of graphEdges) {
    const from = boxes.get(edge.from);
    const to = boxes.get(edge.to);
    if (!from || !to || from === to || to.rank <= from.rank + 1) {
      // Same rank, adjacent ranks, a self-edge or a back edge: nothing to thread.
      edges.push(edge);
      continue;
    }

    const through: string[] = [];
    let previous = edge.from;
    // A link that skips a rank cannot be given a lane: it bends at each dummy it threads
    // through, not once at a boundary. Its label instead *becomes* one of those dummies —
    // sized to the text, so the ordering pass separates it from its neighbours exactly as it
    // separates nodes, and the line passes through it because that is where the lane went.
    const labelRank = from.rank + 1 + Math.floor((to.rank - from.rank - 2) / 2);
    const label = labelBox(edge, labelStyle, config.maxEdgeLabelWidth, horizontal);
    for (let rank = from.rank + 1; rank < to.rank; rank++) {
      const id = `\u0000dummy${serial++}`;
      dummies.add(id);
      const carries = label !== undefined && rank === labelRank;
      if (carries) {
        labelled.set(edge, id);
      }
      boxes.set(id, {
        id,
        shape: "rect",
        classes: [],
        lines: [],
        x: 0,
        y: 0,
        // Wide enough to keep its neighbours off the line, narrow enough to cost nothing —
        // unless it is holding a label, in which case it is exactly as big as the label.
        width: carries ? label.width : Math.min(14, config.nodeGap),
        height: carries ? label.height : Math.min(14, config.nodeGap),
        rank,
        order: 0
      });
      layers[rank].push(id);
      edges.push({
        from: previous,
        to: id,
        stroke: edge.stroke,
        startEnd: "none",
        endEnd: "none",
        minRankSpan: 1
      });
      through.push(id);
      previous = id;
    }
    edges.push({
      from: previous,
      to: edge.to,
      stroke: edge.stroke,
      startEnd: "none",
      endEnd: edge.endEnd,
      minRankSpan: 1
    });
    waypoints.set(edge, through);
  }

  return { edges, dummies, waypoints, labelled };
}

/**
 * The box a label needs, in layout axes.
 *
 * A label is laid out across the flow, so for a left-to-right chart its text width is the box's
 * *height*. Getting that the wrong way round reserves the space in the axis that had room.
 */
function labelBox(
  edge: FlowEdge,
  labelStyle: DrawTextStyle,
  maxLabelWidth: number,
  horizontal: boolean
): { width: number; height: number } | undefined {
  if (edge.label === undefined || edge.label === "") {
    return undefined;
  }
  const lines = wrapText(edge.label, labelStyle, maxLabelWidth);
  const text = lines.reduce((max, line) => Math.max(max, measureText(line, labelStyle)), 0);
  const stack = lines.length * labelStyle.size * LINE_HEIGHT;
  return horizontal
    ? { width: stack + LABEL_CLEARANCE, height: text + LABEL_CLEARANCE }
    : { width: text + LABEL_CLEARANCE, height: stack + LABEL_CLEARANCE };
}

/**
 * Assign each node a rank.
 *
 * Longest path over the edges that point forwards. Back edges — the ones that would close
 * a cycle — are found with a depth-first walk and left out of the calculation; including
 * them has no solution, and dropping them keeps every other edge's constraint intact.
 */
function assignRanks(boxes: Map<string, Mutable>, graphEdges: readonly FlowEdge[]): string[][] {
  const outgoing = new Map<string, string[]>();
  for (const id of boxes.keys()) {
    outgoing.set(id, []);
  }
  const back = new Set<FlowEdge>();
  for (const edge of graphEdges) {
    if (boxes.has(edge.from) && boxes.has(edge.to) && edge.from !== edge.to) {
      outgoing.get(edge.from)!.push(edge.to);
    }
  }

  // Depth-first: an edge back into the current path closes a cycle.
  const state = new Map<string, 0 | 1 | 2>();
  const walk = (id: string): void => {
    state.set(id, 1);
    for (const next of outgoing.get(id) ?? []) {
      const seen = state.get(next) ?? 0;
      if (seen === 1) {
        for (const edge of graphEdges) {
          if (edge.from === id && edge.to === next) {
            back.add(edge);
          }
        }
        continue;
      }
      if (seen === 0) {
        walk(next);
      }
    }
    state.set(id, 2);
  };
  for (const id of boxes.keys()) {
    if ((state.get(id) ?? 0) === 0) {
      walk(id);
    }
  }

  const forward = graphEdges.filter(
    edge => !back.has(edge) && boxes.has(edge.from) && boxes.has(edge.to) && edge.from !== edge.to
  );

  // Longest path by relaxation: repeat until nothing moves. Bounded by the node count,
  // because each pass advances at least one node by at least one rank.
  const rank = new Map<string, number>();
  for (const id of boxes.keys()) {
    rank.set(id, 0);
  }
  for (let pass = 0; pass < boxes.size; pass++) {
    let moved = false;
    for (const edge of forward) {
      const want = rank.get(edge.from)! + Math.max(1, edge.minRankSpan);
      if (want > rank.get(edge.to)!) {
        rank.set(edge.to, want);
        moved = true;
      }
    }
    if (!moved) {
      break;
    }
  }

  const layers: string[][] = [];
  for (const [id, value] of rank) {
    boxes.get(id)!.rank = value;
    (layers[value] ??= []).push(id);
  }
  for (let i = 0; i < layers.length; i++) {
    layers[i] ??= [];
  }
  return layers;
}

/**
 * Order the nodes within each rank to reduce crossings.
 *
 * The median heuristic, swept forwards and backwards a few times: a node wants to sit at
 * the median position of its neighbours in the adjacent rank. Exact crossing minimisation
 * is NP-hard, and the median gets most of the benefit in a handful of passes.
 */
function orderRanks(
  diagram: FlowchartDiagram,
  edges: readonly FlowEdge[],
  layers: string[][],
  boxes: ReadonlyMap<string, Mutable>
): string[][] {
  return groupTogether(diagram, medianOrder(edges, layers, boxes));
}

/**
 * Pull each subgraph's members together within their rank.
 *
 * A frame is drawn around the box its members occupy, so a non-member ordered between two
 * of them ends up inside it — and the picture then claims a membership the source never
 * stated, which is worse than an ugly diagram. Members keep their relative order, and the
 * block lands at the position of its leftmost member, so this disturbs the crossing count
 * as little as it can while making the frame honest.
 */
function groupTogether(diagram: FlowchartDiagram, orders: string[][]): string[][] {
  if (diagram.subgraphs.length === 0) {
    return orders;
  }
  const owner = new Map<string, string>();
  for (const group of diagram.subgraphs) {
    for (const id of group.nodeIds) {
      // A node in two groups belongs to the innermost, which is the last one closed.
      owner.set(id, group.id);
    }
  }

  return orders.map(layer => {
    const seen = new Set<string>();
    const out: string[] = [];
    for (const id of layer) {
      if (seen.has(id)) {
        continue;
      }
      const group = owner.get(id);
      if (group === undefined) {
        out.push(id);
        seen.add(id);
        continue;
      }
      for (const member of layer) {
        if (!seen.has(member) && owner.get(member) === group) {
          out.push(member);
          seen.add(member);
        }
      }
    }
    return out;
  });
}

/**
 * Order the nodes within each rank to reduce crossings.
 *
 * Three ideas, and the third is what makes the other two worth having:
 *
 * 1. **Median.** A node wants to sit at the median position of its neighbours in the
 *    adjacent rank, swept forwards and backwards. Cheap, and it removes most crossings.
 * 2. **Transpose.** After a median pass, try swapping each adjacent pair and keep the swap
 *    when it removes crossings. The median is a heuristic about *position* and cannot see
 *    that two particular neighbours are the wrong way round; this can.
 * 3. **Keep the best.** Count the crossings each iteration actually produces and return
 *    the best arrangement seen. Without this the answer is whatever the last sweep left,
 *    which is regularly worse than something already passed through — a heuristic that
 *    does not measure its own output has no reason to improve it.
 *
 * Exact minimisation is NP-hard; this is the standard combination and gets most of the
 * benefit for a few passes over the graph.
 */
/** Append to a list held in a map, creating it on first use. */
function appendTo(index: Map<string, string[]>, key: string, value: string): void {
  const list = index.get(key);
  if (list === undefined) {
    index.set(key, [value]);
  } else {
    list.push(value);
  }
}

function medianOrder(
  edges: readonly FlowEdge[],
  layers: string[][],
  boxes: ReadonlyMap<string, Mutable>
): string[][] {
  const previous = new Map<string, string[]>();
  const next = new Map<string, string[]>();
  for (const edge of edges) {
    if (!boxes.has(edge.from) || !boxes.has(edge.to) || edge.from === edge.to) {
      continue;
    }
    const a = boxes.get(edge.from)!;
    const b = boxes.get(edge.to)!;
    if (a.rank === b.rank) {
      continue;
    }
    const [lower, upper] = a.rank < b.rank ? [edge.from, edge.to] : [edge.to, edge.from];
    // Appended in place. Rebuilding the array per edge copies everything already in it, which
    // is quadratic in the degree of a node — and a hub with many edges is exactly the shape
    // that makes a diagram slow to lay out.
    appendTo(next, lower, upper);
    appendTo(previous, upper, lower);
  }

  let current = layers.map(layer => [...layer]);
  let best = current.map(layer => [...layer]);
  let bestCount = crossings(best, previous);

  for (let pass = 0; pass < 8 && bestCount > 0; pass++) {
    const forwards = pass % 2 === 0;
    const range = forwards
      ? [...current.keys()].slice(1)
      : [...current.keys()].slice(0, -1).reverse();

    for (const index of range) {
      const reference = current[forwards ? index - 1 : index + 1];
      const position = new Map(reference.map((id, i) => [id, i]));
      const links = forwards ? previous : next;
      const scored = current[index].map((id, i) => ({
        id,
        key: median((links.get(id) ?? []).map(other => position.get(other)).filter(isNumber)) ?? i,
        tie: i
      }));
      scored.sort((a, b) => a.key - b.key || a.tie - b.tie);
      current[index] = scored.map(entry => entry.id);
    }

    transpose(current, previous, next);

    const count = crossings(current, previous);
    if (count < bestCount) {
      bestCount = count;
      best = current.map(layer => [...layer]);
    }
    current = current.map(layer => [...layer]);
  }

  return best;
}

/**
 * Swap adjacent pairs while it helps.
 *
 * Only the two ranks a pair sits between can change, so each trial is a local count rather
 * than a whole-diagram one.
 */
function transpose(
  layers: string[][],
  previous: ReadonlyMap<string, string[]>,
  next: ReadonlyMap<string, string[]>
): void {
  let improved = true;
  let rounds = 0;
  while (improved && rounds < 4) {
    improved = false;
    rounds++;
    for (let rank = 0; rank < layers.length; rank++) {
      const layer = layers[rank];
      // The ranks on either side are indexed once per rank, not once per trial: the inner
      // loop only ever swaps within `layer`, so neither neighbour moves while it runs. This
      // used to be rebuilt on all four calls for every adjacent pair, which made the pass
      // quadratic in rank width for no gain.
      const abovePos = positionsOf(layers[rank - 1]);
      const belowPos = positionsOf(layers[rank + 1]);
      for (let i = 0; i + 1 < layer.length; i++) {
        const before =
          pairCrossings(layer, i, previous, abovePos) + pairCrossings(layer, i, next, belowPos);
        [layer[i], layer[i + 1]] = [layer[i + 1], layer[i]];
        const after =
          pairCrossings(layer, i, previous, abovePos) + pairCrossings(layer, i, next, belowPos);
        if (after < before) {
          improved = true;
        } else {
          [layer[i], layer[i + 1]] = [layer[i + 1], layer[i]];
        }
      }
    }
  }
}

/** Where each id sits in a rank, or `undefined` past either end of the diagram. */
function positionsOf(
  layer: readonly string[] | undefined
): ReadonlyMap<string, number> | undefined {
  return layer === undefined ? undefined : new Map(layer.map((id, i) => [id, i]));
}

/**
 * Crossings contributed by one adjacent pair against the rank on one side.
 *
 * `position` is supplied by the caller because it is the same for every pair in the rank;
 * this is the innermost function of the ordering pass and ran a quarter of the layout.
 */
function pairCrossings(
  layer: readonly string[],
  index: number,
  links: ReadonlyMap<string, string[]>,
  position: ReadonlyMap<string, number> | undefined
): number {
  if (position === undefined) {
    return 0;
  }
  const left = links.get(layer[index]);
  const right = links.get(layer[index + 1]);
  if (left === undefined || right === undefined) {
    return 0;
  }
  let count = 0;
  for (const from of left) {
    const a = position.get(from);
    if (a === undefined) {
      continue;
    }
    for (const to of right) {
      const b = position.get(to);
      if (b !== undefined && a > b) {
        count++;
      }
    }
  }
  return count;
}

/**
 * How many pairs of edges cross, over the whole arrangement.
 *
 * Counted between each adjacent pair of ranks by the standard inversion argument: two
 * edges cross exactly when their endpoints are in the opposite order on the two sides.
 */
function crossings(layers: readonly string[][], previous: ReadonlyMap<string, string[]>): number {
  let total = 0;
  for (let rank = 1; rank < layers.length; rank++) {
    const above = new Map(layers[rank - 1].map((id, i) => [id, i]));
    const pairs: number[] = [];
    for (const id of layers[rank]) {
      for (const source of previous.get(id) ?? []) {
        const at = above.get(source);
        if (at !== undefined) {
          pairs.push(at);
        }
      }
    }
    for (let i = 0; i < pairs.length; i++) {
      for (let j = i + 1; j < pairs.length; j++) {
        if (pairs[i] > pairs[j]) {
          total++;
        }
      }
    }
  }
  return total;
}

function isNumber(value: number | undefined): value is number {
  return value !== undefined;
}

function median(values: number[]): number | undefined {
  if (values.length === 0) {
    return undefined;
  }
  const sorted = [...values].sort((a, b) => a - b);
  const middle = sorted.length >> 1;
  return sorted.length % 2 === 1 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
}

/**
 * Give each labelled link crossing a rank boundary a lane of its own to bend in.
 *
 * A bent path turns at the midpoint of the flow axis, so *every* link crossing a boundary
 * turns at the same coordinate: their horizontal runs land on one line and merge into a
 * single stroke, and their labels — which sit on that run — compete for one strip. Widening
 * the ranks does not help, because a label sits at the midpoint of its own link and therefore
 * moves only half as far as the nodes do.
 *
 * The separation rule keeps two labels apart when their links share an endpoint. It cannot
 * express the rest: with three sources and four targets fully joined, the midpoints form an
 * interleaved lattice, and `(source1 + target2) / 2` falls to the left of
 * `(source3 + target1) / 2` however far apart the ranks are spread. That is a constraint
 * between two ranks at once, which per-rank packing has no way to state — it is why dagre
 * gives a label a rank of its own instead.
 *
 * Lanes say it directly: sort the boundary's labels across the axis and put each one in the
 * first lane where it clears everything already there. Two labels in the same lane cannot
 * overlap because that is the test, and two in different lanes cannot overlap because the gap
 * reserved for the boundary stacks them. The link still passes through its own label, because
 * the lane *is* where the link bends.
 */
interface Lanes {
  /** Which lane each labelled edge bends in, counted from the middle of the boundary. */
  readonly of: Map<FlowEdge, number>;
  /** How many lanes each boundary needs. */
  readonly count: readonly number[];
  /**
   * Distance between lanes, per boundary.
   *
   * Per boundary and measured from the labels there, not one line's height: a label wraps to as
   * many lines as it needs, and a pitch that assumed one put a two-line label six units into its
   * neighbour's lane. That only showed once the labels were narrow enough to wrap — so a layout
   * verified at its natural width was still overlapping at the width a page had room for.
   */
  readonly pitch: readonly number[];
}

function assignLanes(
  diagram: FlowchartDiagram,
  boxes: ReadonlyMap<string, Mutable>,
  labelStyle: DrawTextStyle,
  config: ResolvedOptions,
  horizontal: boolean,
  rankCount: number
): Lanes {
  const of = new Map<FlowEdge, number>();
  const count = new Array<number>(Math.max(0, rankCount - 1)).fill(1);
  const pitch = new Array<number>(Math.max(0, rankCount - 1)).fill(0);

  // Grouped by the boundary the link crosses. Only a link between adjacent ranks bends at a
  // boundary: one that spans further is threaded through dummies and bends at each of them,
  // and one that runs backwards goes round the outside.
  const byBoundary = new Map<number, { edge: FlowEdge; at: number; half: number }[]>();
  for (const edge of diagram.edges) {
    const from = boxes.get(edge.from);
    const to = boxes.get(edge.to);
    if (!from || !to || from === to || to.rank !== from.rank + 1) {
      continue;
    }
    const box = labelBox(edge, labelStyle, config.maxEdgeLabelWidth, horizontal);
    if (box === undefined) {
      continue;
    }
    const centre = (node: Mutable): number =>
      horizontal ? node.y + node.height / 2 : node.x + node.width / 2;
    appendToList(byBoundary, from.rank, {
      edge,
      at: (centre(from) + centre(to)) / 2,
      // The visible edge-label panel is four units wider on each side.
      // Another eight units between panels is enough daylight. `labelBox`
      // includes the larger general-purpose dummy clearance, which made two
      // labels that visibly fit share different lanes and stretched the links.
      half:
        (horizontal
          ? wrapText(edge.label!, labelStyle, config.maxEdgeLabelWidth).length *
            labelStyle.size *
            LINE_HEIGHT
          : labelWidthOf(edge, labelStyle, config.maxEdgeLabelWidth)) /
          2 +
        8
    });
    // The pitch has to clear the tallest label the boundary carries, in the flow axis.
    pitch[from.rank] = Math.max(
      pitch[from.rank],
      (horizontal ? box.width : box.height) + LANE_DAYLIGHT
    );
  }

  for (const [boundary, items] of byBoundary) {
    // Widest first: a long label placed after two short ones has to go in a third lane even
    // when the three would have fitted in two.
    const sorted = [...items].sort((a, b) => b.half - a.half || a.at - b.at);
    const lanes: { from: number; to: number }[][] = [];
    for (const item of sorted) {
      const span = { from: item.at - item.half, to: item.at + item.half };
      let index = lanes.findIndex(
        lane => !lane.some(other => span.from < other.to && other.from < span.to)
      );
      if (index === -1) {
        index = lanes.length;
        lanes.push([]);
      }
      lanes[index].push(span);
      of.set(item.edge, index);
    }
    count[boundary] = Math.max(1, lanes.length);
  }

  return { of, count, pitch };
}

function appendToList<K, V>(index: Map<K, V[]>, key: K, value: V): void {
  const existing = index.get(key);
  if (existing) {
    existing.push(value);
  } else {
    index.set(key, [value]);
  }
}

/**
 * How much room to leave between each pair of consecutive ranks.
 *
 * A constant is wrong as soon as a subgraph is involved. A frame is drawn `framePadding`
 * outside its members with a title bar above them, and all of that comes out of the gap the
 * edges were going to use: in a diagram where one group's last rank was followed directly by
 * another's first, the two frames ended nine units apart and five links had to share a lane
 * that ran along the second group's title. The frame is not decoration drawn afterwards — it
 * occupies space, so the space is reserved here.
 *
 * Only a boundary a frame actually starts or ends at pays. A group spanning the boundary
 * simply continues across it and needs nothing extra.
 */
function rankGaps(
  diagram: FlowchartDiagram,
  boxes: ReadonlyMap<string, Mutable>,
  rankCount: number,
  config: ResolvedOptions,
  lanes?: Lanes
): number[] {
  const gaps = new Array<number>(Math.max(0, rankCount - 1)).fill(config.rankGap);
  // Every lane has to fit between the ranks — including the first. Reserving only the *stack*
  // was the same mistake in a subtler form: a boundary with one lane got the plain gap, so a
  // four-line label 140 units tall was posted into a 24-unit hole and covered the nodes at both
  // ends. `pitch` is a label's height plus daylight, so `count × pitch` is what the labels
  // actually occupy.
  if (lanes) {
    for (let boundary = 0; boundary < gaps.length; boundary++) {
      const need = (lanes.count[boundary] ?? 1) * (lanes.pitch[boundary] ?? 0);
      if (need > 0) {
        gaps[boundary] = Math.max(gaps[boundary], need + LANE_DAYLIGHT * 2);
      }
    }
  }
  for (const group of diagram.subgraphs) {
    const ranks = group.nodeIds
      .map(id => boxes.get(id)?.rank)
      .filter((rank): rank is number => rank !== undefined);
    if (ranks.length === 0) {
      continue;
    }
    const pad = framePadding(diagram, group, config);
    const first = Math.min(...ranks);
    const last = Math.max(...ranks);
    // Below the last rank the frame needs its padding; above the first it needs padding and
    // the title bar as well.
    if (last < gaps.length) {
      gaps[last] += pad;
    }
    if (first > 0) {
      gaps[first - 1] += pad + GROUP_TITLE_HEIGHT;
    }
  }
  return gaps;
}

/**
 * Turn ranks and orders into coordinates.
 *
 * Each rank gets a band along the flow axis, sized by its tallest (or widest) member. Across
 * the axis the nodes are packed by the separation rule, then centred — so a rank with one
 * node sits under the middle of a rank with three, which is what makes a fork look like a
 * fork.
 */
function place(
  boxes: Map<string, Mutable>,
  orders: string[][],
  neighbours: Map<string, string[]>,
  config: ResolvedOptions,
  horizontal: boolean,
  dummies: ReadonlySet<string>,
  spans: ReadonlyMap<string, number>,
  gaps: readonly number[]
): void {
  const across = (box: Mutable): number => (horizontal ? box.height : box.width);

  assignBands(boxes, orders, config, horizontal, gaps);

  // Centre-to-centre positions within each rank, from the separation rule rather than a
  // constant gap: what has to clear is not only the boxes but the labels their links carry.
  const centres = orders.map(layer => {
    const out: number[] = [];
    layer.forEach((id, index) => {
      const box = boxes.get(id)!;
      const previous = index === 0 ? undefined : boxes.get(layer[index - 1])!;
      out.push(
        previous === undefined
          ? across(box) / 2
          : out[index - 1] + separation(previous, box, spans, config, horizontal)
      );
    });
    return out;
  });

  // Widest rank decides the centre line everything else is centred on.
  const widths = orders.map((layer, rank) =>
    layer.length === 0 ? 0 : centres[rank][layer.length - 1] + across(boxes.get(layer.at(-1)!)!) / 2
  );
  const widest = Math.max(0, ...widths);

  orders.forEach((layer, rank) => {
    const offset = (widest - widths[rank]) / 2;
    layer.forEach((id, index) => {
      const box = boxes.get(id)!;
      box.order = index;
      const centreAcross = offset + centres[rank][index];
      if (horizontal) {
        box.y = centreAcross - box.height / 2;
      } else {
        box.x = centreAcross - box.width / 2;
      }
    });
  });

  straighten(boxes, orders, neighbours, config, horizontal, dummies, spans);
}

/**
 * Set each box's coordinate *along* the flow, and nothing else.
 *
 * Separated from the rest of placement because the two axes are independent: a rank's band
 * offset cannot affect where anything sits across the axis, and neither can the gap before
 * it. That is what lets the edge lanes be measured once the across positions are settled and
 * the bands then be recomputed with room for them, instead of guessing the gap up front or
 * iterating the whole layout to a fixed point.
 */
function assignBands(
  boxes: Map<string, Mutable>,
  orders: readonly string[][],
  config: ResolvedOptions,
  horizontal: boolean,
  gaps: readonly number[]
): void {
  const along = (box: Mutable): number => (horizontal ? box.width : box.height);
  let cursor = 0;
  orders.forEach((layer, rank) => {
    const extent = layer.reduce((max, id) => Math.max(max, along(boxes.get(id)!)), 0);
    const band = cursor + extent / 2;
    for (const id of layer) {
      const box = boxes.get(id)!;
      if (horizontal) {
        box.x = band - box.width / 2;
      } else {
        box.y = band - box.height / 2;
      }
    }
    cursor += extent + (gaps[rank] ?? config.rankGap);
  });
}

/**
 * The least distance that may separate the centres of two neighbours in a rank.
 *
 * One definition, because the initial packing and the straightening pass have to agree: a
 * pass that packs to one rule and is then relaxed against a looser one gives the relaxation
 * room to undo it, and the labels collide again after the third sweep.
 */
function separation(
  a: Mutable,
  b: Mutable,
  spans: ReadonlyMap<string, number>,
  config: ResolvedOptions,
  horizontal: boolean
): number {
  const size = (box: Mutable): number => (horizontal ? box.height : box.width);
  const boxes = size(a) / 2 + config.nodeGap + size(b) / 2;
  const labelA = spans.get(a.id) ?? 0;
  const labelB = spans.get(b.id) ?? 0;
  // Doubled, because a label sits at the midpoint of its link: the labels of two links into
  // this pair are only ever half as far apart as the pair itself.
  const labels = labelA === 0 && labelB === 0 ? 0 : labelA + labelB + LABEL_CLEARANCE;
  return Math.max(boxes, labels);
}

/**
 * Pull each node towards the average of its neighbours' cross positions.
 *
 * Three steps per rank, and the third is the one that matters. Wanting a position is not
 * enough — two nodes can want the same one — so the wants are made feasible by pushing
 * left to right until every pair clears the gap. That push is one-directional, and on its
 * own it drags a whole rank rightwards: the trunk of a diagram visibly bent away from a
 * fork because the fork's children had been shoved right and their parent followed them.
 * Translating the finished rank back by its mean displacement fixes that without breaking
 * feasibility, because moving every node by the same amount cannot re-introduce an
 * overlap.
 */
function straighten(
  boxes: Map<string, Mutable>,
  orders: string[][],
  neighbours: Map<string, string[]>,
  config: ResolvedOptions,
  horizontal: boolean,
  dummies: ReadonlySet<string>,
  spans: ReadonlyMap<string, number>
): void {
  const centre = (box: Mutable): number =>
    horizontal ? box.y + box.height / 2 : box.x + box.width / 2;
  const setCentre = (box: Mutable, value: number): void => {
    if (horizontal) {
      box.y = value - box.height / 2;
    } else {
      box.x = value - box.width / 2;
    }
  };

  // Alternate the sweep direction: influence has to travel both up and down the ranks, or
  // a node in the first rank only ever sees where its child was on the previous pass and
  // the chain settles one step behind.
  for (let pass = 0; pass < 8; pass++) {
    const sweep = pass % 2 === 0 ? [...orders.keys()] : [...orders.keys()].reverse();
    for (const rank of sweep) {
      const layer = orders[rank].map(id => boxes.get(id)!).filter(box => box !== undefined);
      if (layer.length === 0) {
        continue;
      }

      const desired = layer.map(box => {
        // Both means are accumulated in one walk of the neighbour list. The readable form —
        // map, filter, filter again, reduce — allocated three arrays per node per sweep, and
        // this runs eight times over every rank; a node with many neighbours made that the
        // most expensive thing in the pass. Summed in list order, so the arithmetic is
        // unchanged down to the last bit.
        let allSum = 0;
        let allCount = 0;
        let solidSum = 0;
        let solidCount = 0;
        for (const id of neighbours.get(box.id) ?? []) {
          const other = boxes.get(id);
          if (other === undefined || other.rank === box.rank) {
            continue;
          }
          const at = centre(other);
          allSum += at;
          allCount++;
          if (!dummies.has(id)) {
            solidSum += at;
            solidCount++;
          }
        }
        if (allCount === 0) {
          return centre(box);
        }
        // A real node follows the nodes it is joined to, not the scaffolding threading past
        // it: letting a dummy pull on its neighbours bent the trunk of a diagram towards
        // whichever long edge happened to run alongside. A dummy, having nothing else to
        // go on, follows everything.
        return !dummies.has(box.id) && solidCount > 0 ? solidSum / solidCount : allSum / allCount;
      });

      const placed = [...desired];
      for (let i = 1; i < placed.length; i++) {
        const minimum =
          placed[i - 1] + separation(layer[i - 1], layer[i], spans, config, horizontal);
        placed[i] = Math.max(placed[i], minimum);
      }
      const drift = placed.reduce((sum, value, i) => sum + (value - desired[i]), 0) / placed.length;

      layer.forEach((box, i) => {
        setCentre(box, placed[i] - drift);
      });
    }
  }
}

/** Adjacency in both directions, which is what straightening reads. */
function adjacency(edges: readonly FlowEdge[]): Map<string, string[]> {
  const out = new Map<string, string[]>();
  const add = (a: string, b: string): void => {
    const list = out.get(a);
    if (list) {
      list.push(b);
    } else {
      out.set(a, [b]);
    }
  };
  for (const edge of edges) {
    if (edge.from === edge.to) {
      continue;
    }
    add(edge.from, edge.to);
    add(edge.to, edge.from);
  }
  return out;
}

/**
 * Where an edge leaves and enters, and the points it follows between.
 *
 * The line is aimed at the two centres and then clipped to each outline, so it stops on
 * the border rather than under it — an arrowhead buried inside a box reads as a missing
 * arrow. A pair of nodes that are not aligned gets one bend at the midpoint of the flow
 * axis, which is the shape a reader expects from a flowchart and keeps the arrowhead
 * arriving square to the border.
 */
function routeEdges(
  diagram: FlowchartDiagram,
  boxes: Map<string, Mutable>,
  clusterIds: ReadonlySet<string>,
  labelStyle: DrawTextStyle,
  config: ResolvedOptions,
  waypoints: ReadonlyMap<FlowEdge, readonly string[]>,
  lanes: Lanes,
  labelled: ReadonlyMap<FlowEdge, string>
): EdgeRoute[] {
  const horizontal = diagram.direction === "LR" || diagram.direction === "RL";
  const routes: EdgeRoute[] = [];
  const rankEdges = rankFlowEdges(diagram, boxes, horizontal);

  // How far apart siblings need to sit is decided by what they carry, not by a constant:
  // two labels thirty units wide overlap at any step narrower than that, and the fan exists
  // to stop exactly that from happening.
  const fanStep = new Map<string, number>();
  for (const edge of diagram.edges) {
    const key = pairKeyOf(edge);
    const width =
      edge.label === undefined || edge.label === ""
        ? 0
        : wrapText(edge.label, labelStyle, config.maxEdgeLabelWidth).reduce(
            (max, line) => Math.max(max, measureText(line, labelStyle)),
            0
          );
    // Panels are four units wider than their text on each side, so the step has to clear
    // both halves plus both panels before there is any daylight between them.
    fanStep.set(key, Math.max(fanStep.get(key) ?? 0, width + 20));
  }

  const all = [...boxes.values()];
  const bounds = {
    minX: Math.min(...all.map(box => box.x)),
    maxX: Math.max(...all.map(box => box.x + box.width)),
    minY: Math.min(...all.map(box => box.y)),
    maxY: Math.max(...all.map(box => box.y + box.height))
  };

  // Two edges between the same pair get the same geometry unless something tells them
  // apart, so they land exactly on top of each other and one label hides the other. Each
  // is given its rank among its siblings and fanned out by it. Detours are counted too:
  // several back edges sharing one lane merge into a single line on the way round.
  const siblingIndex = new Map<FlowEdge, number>();
  const seenPair = new Map<string, number>();
  let detourCount = 0;
  const detourIndex = new Map<FlowEdge, number>();
  for (const edge of diagram.edges) {
    const from = boxes.get(edge.from);
    const to = boxes.get(edge.to);
    if (!from || !to) {
      continue;
    }
    const key = pairKeyOf(edge);
    const seen = seenPair.get(key) ?? 0;
    siblingIndex.set(edge, seen);
    seenPair.set(key, seen + 1);
    // Only a *backward* edge still detours. A forward edge that skips a rank now has a
    // chain of dummies holding a lane open for it, so it goes through rather than round.
    if (from !== to && to.rank <= from.rank && !waypoints.has(edge)) {
      detourIndex.set(edge, detourCount++);
    }
  }

  // Ports: where along a border each edge attaches. Bending at the midpoint of the flow
  // axis makes every incoming edge arrive at the *centre* of its target's leading edge,
  // whatever it came from — so a node with three parents collected three arrowheads on one
  // pixel, and in a class diagram the marks that tell inheritance from composition landed
  // on top of each other. Computed after the detours are known, because an edge going
  // round the outside does not use a port.
  const ports = assignPorts(diagram, boxes, detourIndex, horizontal);

  for (const edge of diagram.edges) {
    const from = boxes.get(edge.from);
    const to = boxes.get(edge.to);
    if (!from || !to) {
      continue;
    }

    // A straight line only stays clear of the other nodes when it joins adjacent ranks.
    // An edge that runs backwards — the one closing a cycle — or skips a rank would cross
    // whatever sits between, and the part of it that shows is the segment between two
    // *other* boxes: a "Draft → Review" pair grew a second arrowhead that neither edge
    // had asked for. Those go round the outside instead, where nothing is in the way.
    const detour = detourIndex.get(edge);
    const through = waypoints.get(edge);
    const bent =
      from === to || detour !== undefined || (through !== undefined && through.length > 0)
        ? undefined
        : bendedPath(
            from,
            to,
            horizontal,
            siblingIndex.get(edge) ?? 0,
            fanStep.get(pairKeyOf(edge)) ?? 0,
            ports.get(edge),
            laneOffset(edge, lanes, from.rank),
            bendAxisFor(from, to, rankEdges),
            clusterIds.has(edge.from) || clusterIds.has(edge.to) ? 6 : 1
          );
    const points =
      from === to
        ? selfLoop(from, config, siblingIndex.get(edge) ?? 0)
        : detour !== undefined
          ? aroundPath(from, to, horizontal, bounds, config, detour)
          : through !== undefined && through.length > 0
            ? threadedPath(from, to, through, boxes, horizontal)
            : bent!.points;

    const route: {
      -readonly [K in keyof EdgeRoute]: EdgeRoute[K];
    } = { edge, points };

    if (edge.label !== undefined && edge.label !== "") {
      const lines = wrapText(edge.label, labelStyle, config.maxEdgeLabelWidth);
      const width = lines.reduce((max, line) => Math.max(max, measureText(line, labelStyle)), 0);
      const height = lines.length * labelStyle.size * LINE_HEIGHT;
      // Where the router put it, not where the finished polyline happens to be halfway along:
      // a bent path's arc midpoint is not on its crossing run once the bend moves into a lane,
      // and a self-loop or a detour has no crossing run to be on.
      const carrier = labelled.get(edge);
      const held = carrier === undefined ? undefined : boxes.get(carrier);
      const anchor =
        bent?.labelAt ??
        (held === undefined
          ? midpoint(points)
          : { x: held.x + held.width / 2, y: held.y + held.height / 2 });
      route.label = {
        text: edge.label,
        lines,
        x: anchor.x - width / 2,
        y: anchor.y - height / 2,
        width,
        height
      };
    }
    routes.push(route);
  }
  return routes;
}

/**
 * One shared bend axis per adjacent pair of ranks.
 *
 * A fan-out is one visual trunk. Computing the gap midpoint independently for
 * each child gives several nearly parallel horizontal runs whenever the child
 * boxes have different heights. The rank boundary is the gap between the
 * trailing edge of the deepest source box and the leading edge of the earliest
 * target box, so every edge crossing it bends on the same line.
 */
function rankFlowEdges(
  diagram: FlowchartDiagram,
  boxes: ReadonlyMap<string, Mutable>,
  horizontal: boolean
): ReadonlyMap<number, number> {
  const bounds = new Map<number, { lead: number; trail: number }>();
  for (const box of boxes.values()) {
    const lead = horizontal ? box.x : box.y;
    const trail = horizontal ? box.x + box.width : box.y + box.height;
    const current = bounds.get(box.rank);
    bounds.set(box.rank, {
      lead: Math.min(current?.lead ?? Infinity, lead),
      trail: Math.max(current?.trail ?? -Infinity, trail)
    });
  }
  const out = new Map<number, number>();
  for (const edge of diagram.edges) {
    const from = boxes.get(edge.from);
    const to = boxes.get(edge.to);
    if (!from || !to || to.rank !== from.rank + 1) {
      continue;
    }
    const before = bounds.get(from.rank);
    const after = bounds.get(to.rank);
    if (before && after) {
      out.set(from.rank, (before.trail + after.lead) / 2);
    }
  }
  return out;
}

function bendAxisFor(
  from: Mutable,
  to: Mutable,
  axes: ReadonlyMap<number, number>
): number | undefined {
  return to.rank === from.rank + 1 ? axes.get(from.rank) : undefined;
}

/**
 * Where along the flow this link bends, relative to the middle of the boundary.
 *
 * Centred, so a boundary carrying one label bends exactly halfway and a diagram with no
 * labels is laid out as it always was.
 */
function laneOffset(edge: FlowEdge, lanes: Lanes, fromRank: number): number {
  const index = lanes.of.get(edge);
  if (index === undefined) {
    return 0;
  }
  const total = lanes.count[fromRank] ?? 1;
  return (index - (total - 1) / 2) * (lanes.pitch[fromRank] ?? 0);
}

/**
 * Centre-to-centre with one bend, clipped to both outlines.
 *
 * Returns the label's anchor as well as the line, because the two cannot be allowed to
 * disagree. Read back off the finished polyline the anchor was wrong twice over: the
 * arc-length midpoint of a bent path slides off the crossing run onto a vertical once the
 * bend moves away from centre, and a sibling fan bows the line sideways while a
 * centre-of-the-two-nodes formula does not — so two links between one pair drew their lines
 * apart and stacked their labels on top of each other in the middle.
 */
function bendedPath(
  from: Mutable,
  to: Mutable,
  horizontal: boolean,
  sibling: number,
  labelStep: number,
  port: { exit: number; enter: number } | undefined,
  bend = 0,
  bendAxis?: number,
  alignTolerance = 1
): { points: { x: number; y: number }[]; labelAt: { x: number; y: number } } {
  // Two spreads, because the two ends of the problem have different limits. The
  // *attachment* can only move as far as the border it has to land on, or the arrowhead
  // detaches from the node; the *middle* has to move as far as the labels need, or they
  // overlap. Spreading only the attachment leaves narrow nodes with colliding labels, and
  // spreading only the middle piles every arrowhead onto one pixel.
  // A port already separates this edge from its neighbours at the border; the sibling fan
  // only has to add what two labels between the *same* pair need on top of that.
  const attach = port === undefined ? fanOffset(sibling, from, to, horizontal, labelStep) : 0;
  const middle = fanMiddle(sibling, labelStep);
  const exitOffset = (port?.exit ?? 0) + attach;
  const enterOffset = (port?.enter ?? 0) + attach;
  const shift = (point: { x: number; y: number }, by: number): { x: number; y: number } => ({
    x: point.x + (horizontal ? 0 : by),
    y: point.y + (horizontal ? by : 0)
  });
  const a = shift({ x: from.x + from.width / 2, y: from.y + from.height / 2 }, exitOffset);
  const b = shift({ x: to.x + to.width / 2, y: to.y + to.height / 2 }, enterOffset);
  // Tiny centre differences are layout noise, not an authored offset. Drawing
  // a 3–4 unit horizontal dogleg between two boxes that visibly share a centre
  // makes an otherwise vertical arrow look crooked. Snap within six units;
  // genuinely displaced nodes still take the orthogonal bend below.
  const aligned = horizontal
    ? Math.abs(a.y - b.y) <= alignTolerance
    : Math.abs(a.x - b.x) <= alignTolerance;

  if (aligned) {
    const axis = horizontal ? (a.y + b.y) / 2 : (a.x + b.x) / 2;
    const snappedA = horizontal ? { x: a.x, y: axis } : { x: axis, y: a.y };
    const snappedB = horizontal ? { x: b.x, y: axis } : { x: axis, y: b.y };
    const start = clip(from, snappedA, snappedB);
    const finish = clip(to, snappedB, snappedA);
    const centre = shift({ x: (start.x + finish.x) / 2, y: (start.y + finish.y) / 2 }, middle);
    if (middle === 0) {
      return { points: [start, finish], labelAt: centre };
    }
    // Bow out to where the label sits, then back: the line stays attached at both ends.
    return { points: [start, centre, finish], labelAt: centre };
  }
  // Bend halfway along the flow axis — or in this link's own lane, so its crossing run and
  // the label riding on it do not land on top of every other link's.
  // Bend halfway through the *clear gap between the two boxes*, not halfway
  // between their centres. Those are only the same when the boxes have equal
  // size. A tall parent over short children put the centre midpoint just a few
  // units below the parent's border, giving the fan-out an uncomfortably short
  // stem even though the rank gap itself was ample.
  const gapMid =
    bendAxis ??
    (horizontal
      ? b.x >= a.x
        ? (from.x + from.width + to.x) / 2
        : (to.x + to.width + from.x) / 2
      : b.y >= a.y
        ? (from.y + from.height + to.y) / 2
        : (to.y + to.height + from.y) / 2);
  const mid = gapMid + bend;
  const via1 = horizontal ? { x: mid, y: a.y } : { x: a.x, y: mid };
  const via2 = horizontal ? { x: mid, y: b.y } : { x: b.x, y: mid };
  return {
    points: [clip(from, a, via1), via1, via2, clip(to, b, via2)],
    // The middle of the crossing run: on the line, in this link's lane, and offset by the
    // same fan and ports the two ends were.
    labelAt: horizontal ? { x: mid, y: (a.y + b.y) / 2 } : { x: (a.x + b.x) / 2, y: mid }
  };
}

/**
 * Follow the lane the dummies held open.
 *
 * The chain gives one point per rank the edge passes through, so the line can bend at each
 * of them instead of taking a detour around everything. The ends are clipped to their real
 * nodes; the middle is exactly where the ordering decided this edge should be.
 */
function threadedPath(
  from: Mutable,
  to: Mutable,
  through: readonly string[],
  boxes: ReadonlyMap<string, Mutable>,
  horizontal: boolean
): { x: number; y: number }[] {
  const centres = through
    .map(id => boxes.get(id))
    .filter((box): box is Mutable => box !== undefined)
    .map(box => ({ x: box.x + box.width / 2, y: box.y + box.height / 2 }));
  if (centres.length === 0) {
    return bendedPath(from, to, horizontal, 0, 0, undefined).points;
  }
  const a = { x: from.x + from.width / 2, y: from.y + from.height / 2 };
  const b = { x: to.x + to.width / 2, y: to.y + to.height / 2 };
  return [clip(from, a, centres[0]), ...centres, clip(to, b, centres[centres.length - 1])];
}

/**
 * Route an edge around the outside of the diagram.
 *
 * The lane sits just past whichever side is nearer to both endpoints, so a back edge in a
 * tall chart hugs the chart rather than swinging across it. Leaving and arriving on the
 * same side keeps the arrowhead square to the border it meets.
 */
function aroundPath(
  from: Mutable,
  to: Mutable,
  horizontal: boolean,
  bounds: { minX: number; maxX: number; minY: number; maxY: number },
  config: ResolvedOptions,
  laneIndex: number
): { x: number; y: number }[] {
  // Each detour gets its own lane, or two of them merge into one line on the way round
  // and the reader cannot tell which end belongs to which.
  const clearance = config.nodeGap + laneIndex * 14;
  const a = { x: from.x + from.width / 2, y: from.y + from.height / 2 };
  const b = { x: to.x + to.width / 2, y: to.y + to.height / 2 };

  if (horizontal) {
    // Flow runs left to right, so detour above or below.
    const above = a.y + b.y < bounds.minY + bounds.maxY;
    const lane = above ? bounds.minY - clearance : bounds.maxY + clearance;
    const exit = { x: a.x, y: above ? from.y : from.y + from.height };
    const enter = { x: b.x, y: above ? to.y : to.y + to.height };
    return [exit, { x: exit.x, y: lane }, { x: enter.x, y: lane }, enter];
  }
  const left = a.x + b.x < bounds.minX + bounds.maxX;
  const lane = left ? bounds.minX - clearance : bounds.maxX + clearance;
  const exit = { x: left ? from.x : from.x + from.width, y: a.y };
  const enter = { x: left ? to.x : to.x + to.width, y: b.y };
  return [exit, { x: lane, y: exit.y }, { x: lane, y: enter.y }, enter];
}

/**
 * A self-edge, drawn as a loop out of one side and back.
 *
 * It cannot be a straight line — the two endpoints coincide — so it gets its own shape
 * rather than degenerating to a zero-length segment the arrowhead cannot be aimed along.
 */
function selfLoop(
  box: Mutable,
  config: ResolvedOptions,
  sibling: number
): { x: number; y: number }[] {
  const reach = Math.min(config.rankGap, 40) + sibling * 12;
  const right = box.x + box.width;
  const top = box.y + box.height * 0.3;
  const bottom = box.y + box.height * 0.7;
  return [
    { x: right, y: top },
    { x: right + reach, y: top },
    { x: right + reach, y: bottom },
    { x: right, y: bottom }
  ];
}

/**
 * Give every edge a place on the borders it joins.
 *
 * Grouped by node and by which side the edge uses, then ordered by where the other end
 * sits: an edge coming from the left attaches to the left of one coming from the right, so
 * the lines arrive in the same order they left and do not cross each other at the border.
 * The spacing is capped by the border it has to fit on.
 */
function assignPorts(
  diagram: FlowchartDiagram,
  boxes: Map<string, Mutable>,
  detourIndex: ReadonlyMap<FlowEdge, number>,
  horizontal: boolean
): Map<FlowEdge, { exit: number; enter: number }> {
  const across = (box: Mutable): number => (horizontal ? box.height : box.width);
  const centreAcross = (box: Mutable): number =>
    horizontal ? box.y + box.height / 2 : box.x + box.width / 2;

  interface Slot {
    readonly edge: FlowEdge;
    readonly end: "exit" | "enter";
    readonly reference: number;
    /** Declaration order, so siblings pointing at the same place still get distinct slots. */
    readonly ordinal: number;
  }
  const sides = new Map<string, Slot[]>();
  let ordinal = 0;

  for (const edge of diagram.edges) {
    const from = boxes.get(edge.from);
    const to = boxes.get(edge.to);
    if (!from || !to || from === to || detourIndex.has(edge)) {
      continue;
    }
    const push = (box: Mutable, other: Mutable, end: "exit" | "enter"): void => {
      // Leading or trailing side, decided by which way the edge runs along the flow.
      const side = `${box.id}\u0000${end === "exit" ? "out" : "in"}`;
      const list = sides.get(side);
      const slot = { edge, end, reference: centreAcross(other), ordinal: ordinal++ };
      if (list) {
        list.push(slot);
      } else {
        sides.set(side, [slot]);
      }
    };
    push(from, to, "exit");
    push(to, from, "enter");
  }

  const out = new Map<FlowEdge, { exit: number; enter: number }>();
  const offsetOf = (edge: FlowEdge): { exit: number; enter: number } => {
    let entry = out.get(edge);
    if (!entry) {
      entry = { exit: 0, enter: 0 };
      out.set(edge, entry);
    }
    return entry;
  };

  for (const [key, slots] of sides) {
    const box = boxes.get(key.slice(0, key.indexOf("\u0000")));
    if (!box || slots.length < 2) {
      continue;
    }
    slots.sort((a, b) => a.reference - b.reference || a.ordinal - b.ordinal);
    // Leave a margin at each end so a port never sits on a corner — and never spread wider
    // than the border itself, or the outermost strand attaches beside the node instead of
    // to it.
    const usable = across(box) * 0.72;
    const spacing = Math.min(26, usable / Math.max(1, slots.length - 1));
    // The edge running straight through — the one whose other end is directly opposite —
    // keeps the centre, and the rest fan around it. Spreading every port evenly moved the
    // trunk of a diagram by a few units at each rank, and a chain of small steps reads as
    // a bend even though no single one does.
    // Evenly spread and centred on the border. Anchoring on "the straightest edge" instead
    // was an attempt to keep a trunk perfectly vertical, and it cost more than it bought:
    // the fan then ran off one side of the node and had to be clamped, which collapsed two
    // strands onto the same point. A centred fan keeps every strand on its border, and the
    // trunk's own step is half a spacing — small enough not to read as a bend.
    slots.forEach((slot, index) => {
      const offset = (index - (slots.length - 1) / 2) * spacing;
      offsetOf(slot.edge)[slot.end] = offset;
    });
  }
  return out;
}

/**
 * How far to one side an edge sits, given its rank among the edges sharing its endpoints.
 *
 * Alternating sides keeps a pair straddling the straight line rather than both leaning one
 * way, and the step is capped by the narrower of the two nodes so a fan cannot walk off
 * the border it is supposed to attach to.
 */
function fanOffset(
  sibling: number,
  from: Mutable,
  to: Mutable,
  horizontal: boolean,
  labelStep: number
): number {
  if (sibling === 0) {
    return 0;
  }
  const across = (box: Mutable): number => (horizontal ? box.height : box.width);
  const room = Math.min(across(from), across(to)) / 2;
  // Wide enough for what the siblings carry, but never so wide that the line leaves the
  // border it is supposed to attach to.
  const step = Math.max(14, Math.min(labelStep, room * 0.6));
  const rank = Math.floor((sibling + 1) / 2);
  const side = sibling % 2 === 1 ? 1 : -1;
  return Math.max(-room * 0.75, Math.min(room * 0.75, rank * step * side));
}

/**
 * How far the middle of an edge sits from the straight line, which is where its label goes.
 *
 * Unclamped by the node: a label needs the room it needs, and the line reaches it by
 * bowing rather than by detaching from the border.
 */
function fanMiddle(sibling: number, labelStep: number): number {
  if (sibling === 0) {
    return 0;
  }
  const rank = Math.floor((sibling + 1) / 2);
  const side = sibling % 2 === 1 ? 1 : -1;
  return rank * Math.max(14, labelStep) * side;
}

/** Both directions between one pair share a key: they are siblings however they point. */
function pairKeyOf(edge: FlowEdge): string {
  return edge.from < edge.to ? `${edge.from}\u0000${edge.to}` : `${edge.to}\u0000${edge.from}`;
}

/**
 * Clip a ray at a node's outline.
 *
 * The outline is approximated by its bounding box for everything but a circle and a
 * rhombus, whose real borders sit far enough inside the box that an arrow would visibly
 * stop short. Those two are worth the exact solution; the rest are not, because their
 * borders touch the box along the direction an edge actually arrives from.
 */
function clip(
  box: Mutable,
  from: { x: number; y: number },
  towards: { x: number; y: number }
): { x: number; y: number } {
  const dx = towards.x - from.x;
  const dy = towards.y - from.y;
  if (dx === 0 && dy === 0) {
    return { ...from };
  }
  const halfW = box.width / 2;
  const halfH = box.height / 2;

  if (
    box.shape === "circle" ||
    box.shape === "doubleCircle" ||
    box.shape === "stateStart" ||
    box.shape === "stateEnd"
  ) {
    const length = Math.hypot(dx, dy);
    const radius = Math.min(halfW, halfH);
    return { x: from.x + (dx / length) * radius, y: from.y + (dy / length) * radius };
  }

  if (box.shape === "rhombus") {
    // |x|/halfW + |y|/halfH = 1 — the diamond's border.
    const t = 1 / (Math.abs(dx) / halfW + Math.abs(dy) / halfH);
    return { x: from.x + dx * t, y: from.y + dy * t };
  }

  const scaleX = dx === 0 ? Infinity : halfW / Math.abs(dx);
  const scaleY = dy === 0 ? Infinity : halfH / Math.abs(dy);
  const t = Math.min(scaleX, scaleY);
  return { x: from.x + dx * t, y: from.y + dy * t };
}

/** The point halfway along a polyline, by arc length. */
function midpoint(points: readonly { x: number; y: number }[]): { x: number; y: number } {
  if (points.length === 1) {
    return { ...points[0] };
  }
  let total = 0;
  for (let i = 1; i < points.length; i++) {
    total += Math.hypot(points[i].x - points[i - 1].x, points[i].y - points[i - 1].y);
  }
  let walked = 0;
  for (let i = 1; i < points.length; i++) {
    const segment = Math.hypot(points[i].x - points[i - 1].x, points[i].y - points[i - 1].y);
    if (walked + segment >= total / 2) {
      const t = segment === 0 ? 0 : (total / 2 - walked) / segment;
      return {
        x: points[i - 1].x + (points[i].x - points[i - 1].x) * t,
        y: points[i - 1].y + (points[i].y - points[i - 1].y) * t
      };
    }
    walked += segment;
  }
  return { ...points[points.length - 1] };
}

/**
 * Push non-members out of the box a subgraph's frame will occupy.
 *
 * A frame is drawn around the extent of its members, and that extent is two-dimensional:
 * keeping members contiguous within a rank is not enough, because the frame spans several
 * ranks and a stranger in any of them can fall inside it. The picture then states a
 * membership the source never did, which is a worse failure than an ugly diagram — a
 * reader has no way to tell it is being lied to.
 *
 * The eviction moves the offender to the nearer edge of the frame and carries every node
 * beyond it along by the same amount, so the rank's order and its gaps both survive.
 * Repeated a few times because one eviction can push a node into a second group.
 */
function evictOutsiders(
  diagram: FlowchartDiagram,
  boxes: Map<string, Mutable>,
  orders: string[][],
  config: ResolvedOptions,
  horizontal: boolean
): void {
  if (diagram.subgraphs.length === 0) {
    return;
  }
  // Each group's own padding: the frame that will be drawn is what has to be cleared, and a
  // group with others nested inside it is drawn wider than one without.
  const members = diagram.subgraphs.map(group => ({
    ids: new Set(group.nodeIds),
    boxes: group.nodeIds
      .map(id => boxes.get(id))
      .filter((box): box is Mutable => box !== undefined),
    pad: framePadding(diagram, group, config)
  }));

  const lowCross = (box: Mutable): number => (horizontal ? box.y : box.x);
  const highCross = (box: Mutable): number => (horizontal ? box.y + box.height : box.x + box.width);
  const lowAlong = (box: Mutable): number => (horizontal ? box.x : box.y);
  const highAlong = (box: Mutable): number => (horizontal ? box.x + box.width : box.y + box.height);
  const shift = (box: Mutable, by: number): void => {
    if (horizontal) {
      box.y += by;
    } else {
      box.x += by;
    }
  };

  for (let pass = 0; pass < 3; pass++) {
    let moved = false;
    for (const group of members) {
      if (group.boxes.length === 0) {
        continue;
      }
      const crossMin = Math.min(...group.boxes.map(lowCross)) - group.pad;
      const crossMax = Math.max(...group.boxes.map(highCross)) + group.pad;
      const alongMin = Math.min(...group.boxes.map(lowAlong)) - group.pad - GROUP_TITLE_HEIGHT;
      const alongMax = Math.max(...group.boxes.map(highAlong)) + group.pad;

      for (const layer of orders) {
        for (let index = 0; index < layer.length; index++) {
          const box = boxes.get(layer[index]);
          if (!box || group.ids.has(box.id)) {
            continue;
          }
          // Only a node that overlaps the frame in *both* axes is inside it.
          if (highAlong(box) <= alongMin || lowAlong(box) >= alongMax) {
            continue;
          }
          if (highCross(box) <= crossMin || lowCross(box) >= crossMax) {
            continue;
          }
          const toLeft = crossMin - highCross(box);
          const toRight = crossMax - lowCross(box);
          const by = Math.abs(toLeft) <= Math.abs(toRight) ? toLeft : toRight;
          // Carry the rest of the rank along, so the order and the gaps both survive.
          for (let rest = index; rest < layer.length; rest++) {
            const follower = boxes.get(layer[rest]);
            if (follower) {
              shift(follower, by);
            }
          }
          if (by !== 0) {
            moved = true;
          }
        }
      }
    }
    if (!moved) {
      return;
    }
  }
}

/** Frame each subgraph around the nodes it holds. */
function frameGroups(
  diagram: FlowchartDiagram,
  boxes: Map<string, Mutable>,
  config: ResolvedOptions
): GroupBox[] {
  const out: GroupBox[] = [];
  for (const group of diagram.subgraphs) {
    const members = group.nodeIds.map(id => boxes.get(id)).filter((b): b is Mutable => !!b);
    if (members.length === 0) {
      continue;
    }
    const pad = framePadding(diagram, group, config);
    const x = Math.min(...members.map(m => m.x)) - pad;
    const y = Math.min(...members.map(m => m.y)) - pad - GROUP_TITLE_HEIGHT;
    const right = Math.max(...members.map(m => m.x + m.width)) + pad;
    const bottom = Math.max(...members.map(m => m.y + m.height)) + pad;
    out.push({ id: group.id, text: group.text, x, y, width: right - x, height: bottom - y });
  }
  // Outermost first, so an inner frame is drawn over the one that contains it.
  return out.sort((a, b) => b.width * b.height - a.width * a.height);
}

/** Room reserved above a group's members for its own name. */
const GROUP_TITLE_HEIGHT = 22;

/**
 * How much clear space a group's frame takes outside its members.
 *
 * One definition, because two consumers need exactly the same answer: the pass that pushes
 * strangers out of a frame, and the pass that draws it. They had a padding each — the second
 * one larger, to leave room for the frames nested inside — so an outer group was drawn wider
 * than the area that had been cleared for it and could still enclose a node that was never
 * a member. A frame that encloses a non-member is the diagram claiming a membership its
 * source never stated, which is the failure the eviction exists to prevent.
 */
function framePadding(
  diagram: FlowchartDiagram,
  group: FlowchartDiagram["subgraphs"][number],
  config: ResolvedOptions
): number {
  const depth = (id: string): number => {
    const children = diagram.subgraphs.filter(other => other.parent === id);
    return children.length === 0 ? 0 : 1 + Math.max(...children.map(child => depth(child.id)));
  };
  return config.nodeGap / 2 + depth(group.id) * (config.nodeGap / 2 + GROUP_TITLE_HEIGHT);
}

/**
 * Translate everything into a padded, positive-coordinate box.
 *
 * Layout works from an arbitrary origin — straightening can push a node to a negative
 * coordinate — so the last pass measures what was produced and moves it into view.
 */
function finish(
  diagram: FlowchartDiagram,
  nodes: readonly Mutable[],
  routes: readonly EdgeRoute[],
  groups: readonly GroupBox[],
  config: ResolvedOptions,
  style: DrawTextStyle
): FlowLayout {
  const xs: number[] = [];
  const ys: number[] = [];
  const note = (x: number, y: number): void => {
    xs.push(x);
    ys.push(y);
  };
  for (const node of nodes) {
    note(node.x, node.y);
    note(node.x + node.width, node.y + node.height);
  }
  for (const group of groups) {
    note(group.x, group.y);
    note(group.x + group.width, group.y + group.height);
  }
  for (const route of routes) {
    for (const point of route.points) {
      note(point.x, point.y);
    }
    if (route.label) {
      note(route.label.x, route.label.y);
      note(route.label.x + route.label.width, route.label.y + route.label.height);
    }
  }

  const titleHeight = diagram.title ? style.size * 1.9 : 0;
  const minX = Math.min(...xs);
  const minY = Math.min(...ys);
  // A title wider than the picture has to widen the picture, or it runs off both edges.
  // It is centred, so the extra room is split between the two sides — which means the
  // shift has to be known *before* anything is placed rather than patched in afterwards.
  const contentWidth = Math.max(...xs) - minX + config.padding * 2;
  const titleWidth =
    diagram.title === undefined
      ? 0
      : measureText(diagram.title, { ...style, size: style.size * 1.35, bold: true }) +
        config.padding * 2;
  const width = Math.max(contentWidth, titleWidth);
  const dx = config.padding - minX + (width - contentWidth) / 2;
  const dy = config.padding - minY + titleHeight;
  // The box the mirror reflects within: everything below the title, plus the padding that
  // frames it. Reflecting inside the whole canvas would swap the title's margin for the
  // bottom one and push the diagram under its own heading.
  const contentHeight = Math.max(...ys) + dy + config.padding;
  // `BT` and `RL` are the same layout read from the other end, so ranking and ordering run
  // in the canonical direction and the finished picture is mirrored. Reversing the ranks
  // instead would have to be undone in every pass that reads them — straightening, ports,
  // detours — and each would be a separate chance to get the sign wrong.
  const flipY = diagram.direction === "BT";
  const flipX = diagram.direction === "RL";
  const mirrorX = (x: number, w: number): number => (flipX ? width - (x + w) : x);
  const mirrorY = (y: number, h: number): number => (flipY ? contentHeight - (y + h) : y);

  const shifted: NodeBox[] = nodes.map(node => ({
    id: node.id,
    shape: node.shape,
    classes: node.classes,
    lines: node.lines,
    x: mirrorX(node.x + dx, node.width),
    y: mirrorY(node.y + dy, node.height),
    width: node.width,
    height: node.height
  }));
  const shiftedGroups: GroupBox[] = groups.map(group => ({
    ...group,
    x: mirrorX(group.x + dx, group.width),
    y: mirrorY(group.y + dy, group.height)
  }));
  const shiftedRoutes: EdgeRoute[] = routes.map(route => ({
    edge: route.edge,
    points: route.points.map(point => ({
      x: mirrorX(point.x + dx, 0),
      y: mirrorY(point.y + dy, 0)
    })),
    ...(route.label === undefined
      ? {}
      : {
          label: {
            ...route.label,
            x: mirrorX(route.label.x + dx, route.label.width),
            y: mirrorY(route.label.y + dy, route.label.height)
          }
        })
  }));

  const height = contentHeight;

  return {
    width,
    height,
    nodes: shifted,
    edges: shiftedRoutes,
    groups: shiftedGroups,
    ...(diagram.title === undefined
      ? {}
      : { title: { text: diagram.title, x: width / 2, y: config.padding + style.size } })
  };
}
