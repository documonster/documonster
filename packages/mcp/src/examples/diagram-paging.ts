/**
 * Example: fitting and paging diagrams in a Word document
 *
 * Run: node --import @oxc-node/core/register src/examples/diagram-paging.ts
 * Output: tmp/mcp-examples/paging/
 *
 * Two problems that only appear once a diagram has to live on a page, and the
 * answer to each. Both are what `doc_write` and `doc_convert` do for every
 * ```mermaid fence they meet; this drives the same code directly so the numbers
 * are visible.
 *
 * ## Fitting is not the same as being readable
 *
 * A page has a fixed measure — US Letter less one-inch margins is 468 points — and
 * Word does not shrink an oversized inline image, it runs it off the edge of the
 * paper. So an image must be scaled to the column. Scaling *alone* is where this
 * goes wrong: a 1941-point flowchart in a 468-point column is drawn at 24%, so its
 * 14-point labels arrive at 3.4pt. In bounds, and legible to nobody.
 *
 * The fix is to make the diagram intrinsically narrower and let it grow downwards,
 * because a page has one width and as many continuations as it needs. Which text to
 * wrap is the whole question, and the answer is ordered by what each concession
 * costs a reader: the gaps first (invisible), then edge labels (loose text between
 * ranks, fine over three short lines), and a node's own text last of all.
 *
 * ## A diagram taller than a page has to be cut, not shrunk
 *
 * Shrinking it back into one page gives away the legibility the narrowing just
 * bought. Cutting it needs a row that crosses nothing important — and "a row of
 * pixels that is entirely background" does not exist in a top-to-bottom flowchart,
 * because the links between ranks cross every gap. Cutting a link is fine: it
 * continues at the top of the next page, which is how any split figure reads. What
 * must not be cut is a box or a word.
 */

import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

import { Io } from "documonster/word";
import { markdownToDocx } from "documonster/word/markdown";

import {
  fitToColumn,
  MAX_EMBED_HEIGHT_POINTS,
  prepareMarkdownDiagrams
} from "../tools/diagram-markdown.js";
import { planSlices } from "../tools/diagram-slice.js";
import { buildDrawList, toRenderOptions } from "../tools/diagram.js";

const here = path.dirname(fileURLToPath(import.meta.url));
const out = path.resolve(here, "../../../../tmp/mcp-examples/paging");
fs.mkdirSync(out, { recursive: true });

/** The measure a Word page leaves for text, and the tallest piece that may go on one. */
const COLUMN_POINTS = 468;

/** A chain `n` ranks deep: as tall as we like, and narrow enough to need no fitting. */
const chain = (n: number, label: (i: number) => string): string =>
  `flowchart TB\n${Array.from(
    { length: n },
    (_, i) => `    S${i}["${label(i)}"] --> S${i + 1}["${label(i + 1)}"]`
  ).join("\n")}`;

const STEPS = [
  "Register the requirement",
  "Confirm the sponsor",
  "Capture acceptance criteria",
  "Identify affected components",
  "Assess cross team impact",
  "Draft the solution outline",
  "Review architecture fit",
  "Confirm data implications",
  "Agree integration approach",
  "Estimate the delivery effort",
  "Allocate to a delivery team",
  "Plan into the sprint",
  "Implement the change",
  "Write the automated tests",
  "Peer review the change",
  "Run the regression suite",
  "Close the requirement"
];

const CASES: Array<[string, string]> = [
  [
    "fits on one page",
    `flowchart TB
    A["Receive request"] --> B{"Valid?"}
    B -- yes --> C["Queue for processing"]
    B -- no --> D["Reject with reason"]
    C --> E["Done"]
    D --> E`
  ],
  ["needs two pages", chain(10, i => STEPS[i % STEPS.length]!)],
  ["needs three pages", chain(15, i => STEPS[i % STEPS.length]!)],
  [
    "wide: fitting alone would make it illegible",
    `flowchart TB
    P["Programme governance"]
${["Distribution", "RM", "Ingestion", "Data migration", "Future teams"]
  .map(
    name =>
      `    P -->|"Programme direction and coordination"| ${name.replace(/ /g, "_")}["${name} team"]`
  )
  .join("\n")}`
  ],
  [
    "a single box taller than a page",
    `flowchart TB
    BIG["Policy text<br/>${Array.from({ length: 42 }, (_, i) => `Clause ${i + 1}: the box keeps going.`).join("<br/>")}"]`
  ]
];

// ---------------------------------------------------------------------------
// What the fitting and the cut plan decide, per diagram
// ---------------------------------------------------------------------------

const options = toRenderOptions({});
console.log(`column ${COLUMN_POINTS}pt, one piece at most ${MAX_EMBED_HEIGHT_POINTS}pt\n`);

for (const [name, source] of CASES) {
  const natural = buildDrawList(source, options);
  const naturalFit = Math.min(1, COLUMN_POINTS / natural.width);
  console.log(`${name}`);
  console.log(
    `  as authored        ${natural.width.toFixed(0).padStart(5)}x${natural.height.toFixed(0).padStart(5)}  ` +
      `scaled to the column, 14pt text would arrive at ${(14 * naturalFit).toFixed(1)}pt`
  );

  // What the pipeline actually uses: the same drawing when it already fits, a narrower
  // layout of it when scaling alone would leave the text under 8pt.
  const { list, fit, short } = fitToColumn(source, options, true);
  if (list.width !== natural.width || list.height !== natural.height) {
    console.log(
      `  after fitting      ${list.width.toFixed(0).padStart(5)}x${list.height.toFixed(0).padStart(5)}  ` +
        `body text now arrives at ${(14 * fit).toFixed(1)}pt` +
        (short ? " — still under 8pt" : "")
    );
  }

  const cuts = planSlices(list, fit, MAX_EMBED_HEIGHT_POINTS);
  console.log(
    `  cut into ${cuts.length} piece(s): ${cuts
      .map(cut => `${(cut.bottom - cut.top).toFixed(0)}pt${cut.forced ? " (forced)" : ""}`)
      .join(", ")}`
  );
  if (cuts.some(cut => cut.forced)) {
    console.log("    forced — a box taller than a page leaves no legal row to cut at");
  }
  console.log();
}

// ---------------------------------------------------------------------------
// The same thing end to end: a real .docx, with the pieces embedded in order
// ---------------------------------------------------------------------------

const markdown = [
  "# Diagram paging",
  "",
  "Each section below holds one case. The prose is here so you can see that a diagram",
  "no longer strands the heading above it on a page of its own.",
  "",
  ...CASES.flatMap(([name, source], index) => [
    `## ${index + 1}. ${name.charAt(0).toUpperCase()}${name.slice(1)}`,
    "",
    "```mermaid",
    source,
    "```",
    "",
    "Text after the diagram, to show what shares the page with it.",
    ""
  ])
].join("\n");

const prepared = await prepareMarkdownDiagrams(markdown);
console.log("notes the tool would report:");
for (const note of prepared.notes) {
  console.log(`  ${note}`);
}

const doc = await markdownToDocx(prepared.markdown, {
  ...(prepared.resolveImage === undefined ? {} : { resolveImage: prepared.resolveImage })
});
const file = path.join(out, "diagram-paging.docx");
// Re-runnable: the example is run again against its own previous output, and
// `Io.writeFile` replaces rather than refusing.
await Io.writeFile(doc, file);

const written = await Io.readFile(file);
console.log(
  `\nwrote ${path.relative(process.cwd(), file)} — ${written.images?.length ?? 0} image(s) for ${CASES.length} diagram(s)`
);
for (const image of written.images ?? []) {
  const view = new DataView(image.data.buffer, image.data.byteOffset);
  // PNG IHDR: width and height are big-endian at byte 16 and 20. Rendered at 2x the
  // display size, so 144 pixels is one inch on the page.
  console.log(
    `  ${(view.getUint32(16) / 144).toFixed(2)}in x ${(view.getUint32(20) / 144).toFixed(2)}in`
  );
}
