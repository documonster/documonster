/**
 * Mermaid fences inside Markdown, for `doc_write` and `doc_convert`.
 *
 * A ` ```mermaid ` fence carried into a Word document as monospace text is
 * useless — the whole point of the fence is that it is a picture. So each one is
 * rendered to a PNG and spliced in as an inline image before the Markdown reaches
 * the converter.
 *
 * The seam is `markdownToDocx`'s `resolveImage` callback: the fence is rewritten
 * to `![alt](documonster-diagram:N)` and the callback answers that one URL scheme
 * and nothing else. Rewriting to a *file* path instead would need a writable
 * scratch directory and would leak the diagrams as loose files beside the
 * document; going through the callback keeps them in memory and inside the
 * package.
 *
 * The width cap is not cosmetic. A flowchart is routinely wider than a page's text
 * column, and Word does not shrink an oversized inline image — it runs off the
 * edge of the paper. Fitting to the text width is the difference between a
 * document and a broken one.
 */

import { encodePng } from "documonster/archive";
import { rasterizeToRgba, translate, type DrawList } from "documonster/draw";
import type { MermaidRenderOptions } from "documonster/mermaid";
import type { MarkdownImageData, MarkdownImportOptions } from "documonster/word/markdown";

import { McpToolError, toolError } from "../errors.js";
import { planSlices } from "./diagram-slice.js";
import {
  EMU_PER_POINT,
  buildDrawList,
  findMermaidFences,
  parseDiagram,
  toRenderOptions,
  type DiagramStyleArgs
} from "./diagram.js";
import { newImageBudget, type ImageBudget } from "./image.js";

/**
 * Widest an embedded diagram may be, in points.
 *
 * US Letter (the Word writer's default page) less one-inch margins is 6.5 inches
 * of text column. A wider image is not clipped by Word, it overflows the page.
 */
const MAX_EMBED_WIDTH_POINTS = 468;

/**
 * Tallest a single embedded piece may be, in points: the page less its margins, less a little.
 *
 * The slack is not decoration. An image exactly as tall as the text area cannot share a page
 * with *anything* — not even the heading it belongs to — so Word pushes it alone onto the next
 * page and leaves the heading stranded at the bottom of the previous one, which is how a
 * document grew a near-empty page before every large diagram.
 */
// Leaves room on the first slice's page for an H2 heading and the image
// paragraph's own spacing. At 600pt both individually fit a 648pt text area,
// but not together, which stranded the heading on the preceding page.
export const MAX_EMBED_HEIGHT_POINTS = 550;

/**
 * Smallest body text an embedded diagram may end up with on the page, in points.
 *
 * Scaling a diagram to the column is necessary but not sufficient, and treating it as
 * sufficient is what made these documents unreadable: a 1941-point flowchart fitted to a
 * 468-point column is drawn at 24%, so its 14-point labels arrive at 3.4pt — present in the
 * file, legible to nobody. Eleven of sixteen diagrams in one real document came out under
 * 8pt, and the two worst were the ones a reader complained about.
 *
 * 8pt is the floor because it is roughly the smallest footnote a printed page uses.
 */
const MIN_EMBED_FONT_POINTS = 8;

/**
 * Successively tighter layouts to try when a diagram will not fit the column legibly.
 *
 * Scaling is the wrong lever and there is only one right one: make the diagram *intrinsically*
 * narrower and let it grow downwards, where a page is free to continue onto the next one.
 *
 * **Ordered by what it costs the reader, not by how much width it saves.** That ordering is
 * the whole design, and having it wrong produced a diagram nobody wanted to look at: the gaps
 * cost nothing at all to tighten, an edge label is loose text in the space between ranks and
 * reads perfectly well over three short lines, but a node's text is the thing its box is drawn
 * around. One number governed all three, so narrowing a chart to fit wrapped a ten-word node
 * label at seventy points — a column two words wide and ten lines deep, in a diagram that was
 * then nine inches tall and mostly empty box. The same chart, with only its edge labels
 * narrowed, is 6.5 × 5.4 inches with every node on four readable lines.
 *
 * Note what does *not* work, since it looks like it should: rendering with a smaller
 * `fontSize`. The whole diagram shrinks with it, so the fit rises by exactly the factor the
 * text falls and the size on the page is unchanged.
 */
const FIT_LADDER: readonly MermaidRenderOptions[] = [
  // Gaps first: invisible to a reader, and worth about a tenth of the width.
  { nodeGap: 26, rankGap: 40 },
  { nodeGap: 22, rankGap: 32 },
  // Then edge labels.
  { nodeGap: 22, rankGap: 32, maxEdgeLabelWidth: 140 },
  { nodeGap: 22, rankGap: 32, maxEdgeLabelWidth: 100 },
  { nodeGap: 20, rankGap: 30, maxEdgeLabelWidth: 70 },
  // Only then the node text, and gently.
  { nodeGap: 20, rankGap: 30, maxEdgeLabelWidth: 70, maxLabelWidth: 190 },
  { nodeGap: 20, rankGap: 30, maxEdgeLabelWidth: 70, maxLabelWidth: 160 },
  { nodeGap: 20, rankGap: 30, maxEdgeLabelWidth: 70, maxLabelWidth: 140 },
  { nodeGap: 18, rankGap: 44, maxEdgeLabelWidth: 60, maxLabelWidth: 120 },
  { nodeGap: 18, rankGap: 44, maxEdgeLabelWidth: 60, maxLabelWidth: 100 }
];

/**
 * Pixels per point for an embedded diagram — 216 DPI.
 *
 * PDF and DOCX embed byte-for-byte equivalent pixels, but Preview applies a
 * better downsampling filter than Word. At 144 DPI Word made one- and two-pixel
 * connector strokes look kinked or misaligned even though the geometry was
 * identical. A third sample per point gives its renderer enough coverage while
 * costing materially less than 288 DPI.
 */
const EMBED_SCALE = 3;

/** URL scheme the rewritten fences use. Deliberately not a real one. */
const DIAGRAM_URL_PREFIX = "documonster-diagram:";

/** Fences one document may carry, so a pathological input cannot exhaust memory. */
const MAX_EMBEDDED_DIAGRAMS = 20;

/** Aggregate rendering budget for one document's fences. */
const MAX_EMBED_TOTAL_BYTES = 64 * 1024 * 1024;
const MAX_EMBED_TOTAL_PIXELS = 80_000_000;
/** Aggregate image relationships after tall diagrams are split. */
const MAX_EMBEDDED_PIECES = 60;

export interface PreparedMarkdown {
  /** The Markdown with every mermaid fence replaced by an image reference. */
  readonly markdown: string;
  /**
   * Pass as `markdownToDocx`'s `resolveImage`. Present only when at least one
   * fence was rendered, so the no-diagram path behaves exactly as before.
   */
  readonly resolveImage?: MarkdownImportOptions["resolveImage"];
  /** Lines worth reporting to the caller. */
  readonly notes: readonly string[];
}

/**
 * The narrowest layout of a diagram that the text column can show legibly.
 *
 * Walks {@link FIT_LADDER} and stops at the first rung whose text clears
 * {@link MIN_EMBED_FONT_POINTS} — first, not best, because the ladder is ordered by what each
 * concession costs a reader, so the earliest rung that works is the least damage. Height is
 * not a constraint here: a diagram taller than a page is cut across pages instead of being
 * shrunk back into one, which would give away the legibility the narrowing just bought.
 *
 * If no rung clears the bar — a rank with nine nodes in it is wide however hard the labels
 * wrap — the narrowest attempt is used and the caller is told, because a diagram nobody can
 * read is worth a line of output rather than silence.
 *
 * A caller who set any of the layout numbers themselves is left alone: they have said what the
 * layout should be, and second-guessing an explicit argument is worse than a small image.
 */
export function fitToColumn(
  source: string,
  options: MermaidRenderOptions & { readonly background: string },
  tune: boolean
): { list: DrawList; fit: number; short: boolean } {
  const measure = (list: DrawList): { list: DrawList; fit: number; short: boolean } => {
    const fit = Math.min(1, MAX_EMBED_WIDTH_POINTS / list.width);
    // The floor applies to the diagram's body text. Edge labels are authored a
    // step smaller on purpose; forcing the smallest label to 8pt over-wraps
    // every node, turning a balanced one-page chart into a tall two-page strip.
    const size = (options.fontSize ?? DEFAULT_DIAGRAM_FONT_POINTS) * fit;
    return { list, fit, short: size < MIN_EMBED_FONT_POINTS };
  };

  let best = measure(buildDrawList(source, options));
  if (!best.short || !tune) {
    return best;
  }
  for (const step of FIT_LADDER) {
    const attempt = measure(buildDrawList(source, { ...options, ...step }));
    if (attempt.list.width < best.list.width) {
      best = attempt;
    }
    if (!attempt.short) {
      return attempt;
    }
  }
  return best;
}

/** The layout's own default, which `toRenderOptions` leaves unset when the caller does. */
const DEFAULT_DIAGRAM_FONT_POINTS = 14;

/**
 * Render every mermaid fence in `markdown` and rewrite it as an inline image.
 *
 * A fence that does not parse is left exactly as it was — a code block — and
 * reported. Failing the whole document because one diagram is malformed would
 * throw away the nine paragraphs that were fine, and the note names the line to
 * fix.
 */
export async function prepareMarkdownDiagrams(
  markdown: string,
  style: DiagramStyleArgs = {}
): Promise<PreparedMarkdown> {
  const fences = findMermaidFences(markdown);
  if (fences.length === 0) {
    return { markdown, notes: [] };
  }

  if (fences.length > MAX_EMBEDDED_DIAGRAMS) {
    throw toolError.tooLarge(
      `this document has ${fences.length} mermaid fences, over the ${MAX_EMBEDDED_DIAGRAMS} limit for one call`,
      "Every diagram is rendered and held in memory until the document is written. Split the document, or pass diagrams: false and render the ones you need with diagram_render."
    );
  }

  const options = toRenderOptions(style);
  // Only tune a layout the caller has not spoken about.
  const tune =
    style.maxLabelWidth === undefined && style.nodeGap === undefined && style.rankGap === undefined;
  const images = new Map<string, MarkdownImageData>();
  const failures: string[] = [];
  const cramped: number[] = [];
  const split: { line: number; pieces: number }[] = [];
  const forced: number[] = [];
  // One budget across every fence: ten diagrams each just under the rasteriser's
  // own per-image cap is gigabytes, and each is held until the document is written.
  const budget = newImageBudget();
  let drawn = 0;
  let rewritten = markdown;

  // Back to front, so each splice leaves the earlier fences' offsets valid.
  for (const fence of [...fences].reverse()) {
    let references: string;
    try {
      const diagram = parseDiagram(fence.source);
      const { list, fit, short } = fitToColumn(fence.source, options, tune);
      if (short) {
        cramped.push(fence.line);
      }
      const alt = diagram.title ?? `${diagram.kind} diagram`;
      const cuts = planSlices(list, fit, MAX_EMBED_HEIGHT_POINTS);
      if (cuts.some(cut => cut.forced)) {
        forced.push(fence.line);
      }
      if (cuts.length > 1) {
        split.push({ line: fence.line, pieces: cuts.length });
      }
      const parts: string[] = [];
      const pending: Array<{ url: string; image: MarkdownImageData }> = [];
      const fenceBudget = { ...budget };
      cuts.forEach((cut, index) => {
        // Paint this viewport, not the whole diagram then crop it. A diagram
        // taller than the rasteriser's pixel limit is exactly the case slicing
        // is meant to make possible; allocating it first defeats that contract.
        const sliceHeight = (cut.bottom - cut.top) / fit;
        const slice: DrawList = {
          ...list,
          height: sliceHeight,
          children: [
            {
              kind: "group",
              transform: translate(0, -cut.top / fit),
              children: list.children
            }
          ]
        };
        const piece = rasterizeToRgba(slice, { scale: EMBED_SCALE * fit });
        const encoded = {
          bytes: encodePng(piece.data, piece.width, piece.height, {
            // Pixels and the OOXML extent have both already been reduced by
            // `fit`, so their ratio remains 2 px/pt = 144 DPI. Multiplying by
            // fit again made a 50%-fitted PNG falsely claim 72 DPI.
            dpi: 72 * EMBED_SCALE
          }),
          width: piece.width,
          height: piece.height
        };
        spendEmbedBudget(fenceBudget, encoded, fence.line);
        const url =
          cuts.length === 1
            ? `${DIAGRAM_URL_PREFIX}${fence.ordinal}`
            : `${DIAGRAM_URL_PREFIX}${fence.ordinal}.${index + 1}`;
        pending.push({
          url,
          image: {
            data: encoded.bytes,
            mediaType: "png",
            width: Math.round(list.width * fit * EMU_PER_POINT),
            height: Math.round((cut.bottom - cut.top) * EMU_PER_POINT)
          }
        });
        // Each piece names its place in the sequence, so a reader who meets the second half
        // three pages later knows it is the second half.
        const label = cuts.length === 1 ? alt : `${alt} (${index + 1} of ${cuts.length})`;
        parts.push(`![${escapeAlt(label)}](${url})`);
      });
      // Commit one fence atomically. A later slice can exhaust the aggregate
      // budget; registering earlier slices first leaves unreachable image data
      // retained after the fence falls back to a code block.
      for (const pendingImage of pending) {
        images.set(pendingImage.url, pendingImage.image);
      }
      Object.assign(budget, fenceBudget);
      drawn++;
      references = parts.join("\n\n");
    } catch (cause) {
      // A budget is a hard execution boundary, not a per-diagram rendering
      // failure. Continuing would keep spending the resource that ran out.
      if (cause instanceof McpToolError && cause.code === "too_large") {
        throw cause;
      }
      failures.push(
        `- **diagram at line ${fence.line} left as a code block**: ${cause instanceof Error ? cause.message : String(cause)}`
      );
      continue;
    }

    // A blank line either side: an image reference that lands against a
    // neighbouring line is parsed as part of that paragraph.
    rewritten = `${rewritten.slice(0, fence.start)}\n${references}\n${rewritten.slice(fence.end)}`;
  }

  const notes =
    images.size === 0
      ? failures
      : [
          `- ${drawn} mermaid diagram(s) rendered and embedded as PNG, fitted to the text column`,
          ...[...split]
            .reverse()
            .map(
              entry =>
                `- diagram at line ${entry.line} is taller than one page: cut into ${entry.pieces} pieces at rows that cross no box or label`
            ),
          ...(forced.length === 0
            ? []
            : [
                `- diagram(s) at line ${[...forced].reverse().join(", ")} contain a box taller than a page, so a cut had to pass through it`
              ]),
          ...(cramped.length === 0
            ? []
            : [
                `- diagram(s) at line ${[...cramped].reverse().join(", ")} are wider than the text column allows: their text lands under ${MIN_EMBED_FONT_POINTS}pt even at the tightest layout. Split the widest rank, or render them separately with diagram_render.`
              ]),
          ...failures
        ];

  return {
    markdown: rewritten,
    ...(images.size === 0 ? {} : { resolveImage: (url: string) => images.get(url) }),
    notes
  };
}

/**
 * Charge one rendered diagram against the call's budget.
 *
 * A `too_large` here aborts the whole conversion rather than being collected as a
 * per-fence failure: the limit is about the call, and continuing would keep spending
 * exactly the resource that ran out.
 */
function spendEmbedBudget(
  budget: ImageBudget,
  rendered: { readonly bytes: Uint8Array; readonly width: number; readonly height: number },
  line: number
): void {
  budget.count += 1;
  budget.bytes += rendered.bytes.length;
  budget.pixels += rendered.width * rendered.height;
  if (
    budget.count > MAX_EMBEDDED_PIECES ||
    budget.bytes > MAX_EMBED_TOTAL_BYTES ||
    budget.pixels > MAX_EMBED_TOTAL_PIXELS
  ) {
    throw toolError.tooLarge(
      `the diagrams in this document exceed the per-call rendering budget (reached at the fence on line ${line})`,
      "Split the document, or pass diagrams: false and render the large ones separately with diagram_render."
    );
  }
}

/** Escape image-alt delimiters without deleting part of the accessible name. */
function escapeAlt(text: string): string {
  return text.replace(/\\/g, "\\\\").replace(/\r?\n/g, " ").replace(/[[\]]/g, "\\$&");
}
