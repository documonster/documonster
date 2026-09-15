/**
 * Edge labels, groups, and fitting a chart to a narrow measure.
 *
 * Run with `pnpm example --filter labels-and-groups`; output lands in `tmp/mermaid-labels`.
 *
 * Four things a flowchart does that are easy to get wrong, each written as a *pair* of
 * pictures so the difference is visible rather than asserted:
 *
 * 1. `-. text .->` — the dotted link with a label. Its head is `-.` and its tail `.->`, which
 *    is a different shape from every other link form, and a parser that misses it does not
 *    complain: it drops the edge. Five of eleven edges vanishing from one chart left its two
 *    clusters with nothing to rank against, and they were laid out side by side, fourteen times
 *    wider than tall.
 *
 * 2. Many labelled links crossing one rank boundary. A label sits at the midpoint of its own
 *    link, so for a fan out of one node the labels spread only *half* as far as the targets do
 *    — widening the chart cannot separate them. Each link gets a lane of its own to bend in
 *    instead, and the labels stack rather than collide.
 *
 * 3. An edge naming a `subgraph`. It attaches to the group, not to a node of that name, and the
 *    group's members are ranked with respect to it — which is what puts the group *below* the
 *    node pointing at it rather than beside it.
 *
 * 4. `maxLabelWidth` against `maxEdgeLabelWidth`. Narrowing a chart to fit a page means
 *    wrapping text harder, and the two kinds of text tolerate that very differently: an edge
 *    label is loose text between ranks and reads fine over three short lines, while a node's
 *    text is what its box is drawn around. One number for both turns a ten-word node label into
 *    a column two words wide and ten lines deep.
 */

import { mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";

import { encodePng } from "@archive/png";
import { rasterizeToRgba, toSvg } from "@draw/index";
import { layoutFlowchart, mermaidToDrawList, parseMermaid } from "@mermaid/index";
import type { FlowchartDiagram } from "@mermaid/types";

const OUT = process.argv[2] ?? path.join(process.cwd(), "tmp", "mermaid-labels");
mkdirSync(OUT, { recursive: true });

/** Write one diagram as both SVG and PNG, and report its size and label count. */
function draw(name: string, source: string, options: Parameters<typeof mermaidToDrawList>[1] = {}) {
  const list = mermaidToDrawList(source, options);
  writeFileSync(path.join(OUT, `${name}.svg`), toSvg(list, { background: "#ffffff" }));
  const pixels = rasterizeToRgba(list, { scale: 2 });
  writeFileSync(
    path.join(OUT, `${name}.png`),
    encodePng(pixels.data, pixels.width, pixels.height, { dpi: 144 })
  );
  return list;
}

/** How many pairs of edge labels share any area. Zero is the only acceptable answer. */
function overlappingLabels(source: string, options: Parameters<typeof layoutFlowchart>[1] = {}) {
  const boxes = layoutFlowchart(parseMermaid(source) as FlowchartDiagram, options)
    .edges.map(route => route.label)
    .filter((label): label is NonNullable<typeof label> => label !== undefined);
  let hits = 0;
  for (let i = 0; i < boxes.length; i++) {
    for (let j = i + 1; j < boxes.length; j++) {
      const a = boxes[i];
      const b = boxes[j];
      if (
        Math.min(a.x + a.width, b.x + b.width) > Math.max(a.x, b.x) &&
        Math.min(a.y + a.height, b.y + b.height) > Math.max(a.y, b.y)
      ) {
        hits++;
      }
    }
  }
  return { labels: boxes.length, overlaps: hits };
}

// ---------------------------------------------------------------------------
// 1. The dotted link with a label
// ---------------------------------------------------------------------------

const DOTTED = `flowchart TB
    DM["Digital Modernisation"]

    subgraph ALIGNMENT["Central Alignment"]
        B["Business priority"]
        T["Technical standards"]
        V["Programme visibility"]
    end

    D[("Distribution team")]
    RM[("RM team")]
    IM[("Ingestion team")]

    DM --> ALIGNMENT
    ALIGNMENT -. alignment .-> D
    ALIGNMENT -. alignment .-> RM
    ALIGNMENT -. alignment .-> IM

    D --> OUT["One integrated outcome"]
    RM --> OUT
    IM --> OUT`;

const parsed = parseMermaid(DOTTED) as FlowchartDiagram;
const dotted = parsed.edges.filter(edge => edge.stroke === "dotted");
console.log("1. `-. text .->` and an edge naming a subgraph");
console.log(
  `   edges parsed: ${parsed.edges.length} (all of them; dotted with a label: ${dotted.length})`
);
console.log(`   labels on those: ${dotted.map(edge => edge.label ?? "-").join(", ")}`);
// No node called ALIGNMENT: the edge attached to the group.
console.log(`   nodes: ${parsed.nodes.map(node => node.id).join(", ")}`);

const laid = layoutFlowchart(parsed);
const frame = laid.groups.find(group => group.id === "ALIGNMENT")!;
const top = laid.nodes.find(node => node.id === "DM")!;
console.log(
  `   DM ends at y=${(top.y + top.height).toFixed(0)}, the ALIGNMENT frame starts at y=${frame.y.toFixed(0)} — below it, not beside it`
);
draw("1-dotted-labels-and-group", DOTTED);

// ---------------------------------------------------------------------------
// 2. A fan of labelled links across one boundary
// ---------------------------------------------------------------------------

const FAN = `flowchart TB
    P["Programme governance"]
${["Distribution", "RM", "Ingestion", "Data migration", "Future teams"]
  .map(
    name =>
      `    P -->|"Programme direction and coordination"| ${name.replace(/ /g, "_")}["${name}"]`
  )
  .join("\n")}`;

console.log("\n2. Five identical labels across one rank boundary");
const fan = overlappingLabels(FAN);
console.log(`   labels: ${fan.labels}, overlapping pairs: ${fan.overlaps}`);
const fanList = draw("2-label-lanes", FAN);
console.log(`   drawing is ${fanList.width.toFixed(0)}x${fanList.height.toFixed(0)}`);

// ---------------------------------------------------------------------------
// 3. Sources and targets fully joined — the case a per-rank rule cannot state
// ---------------------------------------------------------------------------

const LATTICE = `flowchart TB
${["S1", "S2", "S3"]
  .flatMap(source =>
    ["T1", "T2", "T3", "T4"].map(
      target =>
        `    ${source}["${source}"] -. "dependencies and timeline alignment" .-> ${target}["${target}"]`
    )
  )
  .join("\n")}`;

console.log("\n3. Three sources, four targets, every pair labelled");
const lattice = overlappingLabels(LATTICE);
console.log(`   labels: ${lattice.labels}, overlapping pairs: ${lattice.overlaps}`);
console.log("   the midpoints interleave — (S1+T2)/2 falls left of (S3+T1)/2 — so no rule about");
console.log("   neighbours within one rank can separate them; the lanes do");
draw("3-interleaved-lattice", LATTICE);

// ---------------------------------------------------------------------------
// 4. Narrowing a chart: which text to wrap
// ---------------------------------------------------------------------------

const WORDY = `flowchart TB
    LEAD(["Technical lead"])
    ROLE["Technical alignment role<br/>Architecture principles<br/>Integration patterns<br/>Data and domain principles"]
    LEAD --> ROLE
    ROLE -. "technical alignment" .-> A[("Distribution")]
    ROLE -. "technical alignment" .-> B[("RM")]
    ROLE -. "technical alignment" .-> C[("Ingestion and matching")]
    ROLE -. "technical alignment" .-> D[("Data migration")]`;

console.log("\n4. Narrowing to a 468pt text column (6.5in less one-inch margins)");
for (const [label, options] of [
  ["natural", {}],
  ["both narrowed to 70", { maxLabelWidth: 70, maxEdgeLabelWidth: 70 }],
  ["only edge labels narrowed", { maxEdgeLabelWidth: 70 }]
] as const) {
  const list = mermaidToDrawList(WORDY, options);
  const fit = Math.min(1, 468 / list.width);
  console.log(
    `   ${label.padEnd(26)} ${list.width.toFixed(0).padStart(4)}x${list.height.toFixed(0).padStart(4)} ` +
      `-> on the page ${(list.width * fit).toFixed(0)}x${(list.height * fit).toFixed(0)}pt, ` +
      `14pt text arrives at ${(14 * fit).toFixed(1)}pt`
  );
}
draw("4a-natural", WORDY);
draw("4b-both-narrowed", WORDY, { maxLabelWidth: 70, maxEdgeLabelWidth: 70 });
draw("4c-edge-labels-only", WORDY, { maxEdgeLabelWidth: 70 });
console.log("   compare 4b with 4c: the same width budget, spent on different text");

console.log(`\nWrote SVG and PNG for each to ${OUT}`);
