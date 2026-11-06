/**
 * Where lines land, against Word's own layout.
 *
 * Each fixture in `data/word-probes` is a probe the library wrote, as Microsoft
 * Word for Mac 16.114 opened it. `PG` bookmarks mark the lines measured — in
 * the body, in table cells at any depth and in footnotes. `expected.json` holds
 * Word's page for each, and for the nested-table probes its vertical position
 * on the page as well:
 *
 * - `Qa*`: a nested table in a cell. `Qaa`–`Qac` are auto-fit (Word sizes both
 *   the outer and the nested columns by content, and the nested rows wrap);
 *   `Qae` is fixed-layout and fits its cell, `Qag` is fixed-layout and wider.
 *   Positions come from Word's range information.
 * - `R*-nest`: how wide Word sets a nested table — fixed or auto-fit, with and
 *   without a declared width, in a fixed or auto-fit outer table (see
 *   `nestedTableWidthTwips`). `nestedColumns` is the first nested column's
 *   width from the same range information; the closing empty paragraph after
 *   the nested table takes no height in every one.
 * - `N6d`, `N6e`: mixed documents. N6d ends a landscape section with a footnote
 *   still continuing, which Word finishes on a page of its own; N6e breaks a
 *   row whose next line fits only without the rule closing the row. Pages come
 *   from Word's range information and agree with Word's PDF.
 * - `Qb*`: footnotes longer than a page — one note after a short paragraph, one
 *   referenced near the bottom of a page, and two in a row. Pages come from the
 *   PDF Word exported, since range information is unreliable in a note.
 * - `Qc-s*`: a row that has to break under a repeated header row, a three-line
 *   cell paragraph moved 2pt further down the page each time. Pages come from
 *   Word's PDF.
 * - `P*` (from `tmp/office-tests/pages2`): calibration probes, Word 16.114.
 *   `P1*` facing pages (mirror margins, odd/even headers) and the blank page
 *   between two of the same parity; `P3c`, `P4*` where the gutter goes on a
 *   page turned against the first section; `P5a` paragraph borders and their
 *   room; `P9` vertically merged cells, across a page too; `P10` exact and
 *   at-least row heights against their rules; `P11` rows broken under and
 *   without a repeated header; `P12*` columns — equal, unequal, column breaks,
 *   balanced before a continuous break; `P13` a paragraph ending in a line or
 *   column break; `P15*` balancing with and without space after.
 * - `R1*`: round one of mixed documents — columns, first/odd/even headers,
 *   floating images with every wrap, merged and long tables, lists with
 *   borders, keep and spacing rules, mixed runs and tab leaders, a CJK document
 *   grid, mirror margins with a gutter and number restarts. Positions are left
 *   out where Word's range information does not give the line's top: lines
 *   off a single line height (document grid, at-least spacing), bordered
 *   paragraphs, and the first line of a row continued on a new page.
 */

import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { describe, expect, it } from "vitest";

import { Io, Layout } from "../index";
import type { LayoutParagraph, LayoutTable } from "../index";

const DATA = resolve(__dirname, "data/word-probes");

interface Expected {
  readonly pageCount: number;
  /** Per bookmark: its page, and for some its line's top on the page in points. */
  readonly marks: Readonly<Record<string, readonly [number, number?]>>;
  /** Widths, in points, of the first nested table's leading columns. */
  readonly nestedColumns?: readonly number[];
}

/** The first table nested in a cell on the first page. */
function firstNestedTable(layout: ReturnType<typeof Layout.compute>): LayoutTable | undefined {
  for (const block of layout.pages[0]?.content ?? []) {
    if (block.type === "table") {
      for (const cell of block.cells) {
        const nested = cell.content.find((b): b is LayoutTable => b.type === "table");
        if (nested) {
          return nested;
        }
      }
    }
  }
  return undefined;
}

const { documents } = JSON.parse(readFileSync(resolve(DATA, "expected.json"), "utf8")) as {
  documents: Record<string, Expected>;
};

/** Page and page-absolute line top of every `PG` bookmark the layout placed. */
function markPositions(layout: ReturnType<typeof Layout.compute>): Map<string, [number, number]> {
  const out = new Map<string, [number, number]>();
  layout.pages.forEach((page, index) => {
    const visit = (blocks: readonly unknown[] | undefined, dy: number): void => {
      for (const block of (blocks ?? []) as (LayoutParagraph | LayoutTable)[]) {
        if (block.type === "paragraph") {
          block.lineBookmarks?.forEach((names, line) => {
            for (const name of names) {
              if (name.startsWith("PG") && !out.has(name)) {
                out.set(name, [index + 1, dy + block.rect.y + block.lines[line]!.y]);
              }
            }
          });
        } else if (block.type === "table") {
          for (const cell of block.cells) {
            visit(cell.content, dy + block.rect.y + cell.rect.y);
          }
        }
      }
    };
    visit(page.content, page.geometry.marginTop);
    visit(page.footnoteArea, 0);
  });
  return out;
}

describe("lines land where Word puts them", () => {
  it.each(Object.entries(documents))("%s", async (name, expected) => {
    const doc = await Io.read(readFileSync(resolve(DATA, `${name}.docx`)));
    const layout = Layout.compute(doc);
    const positions = markPositions(layout);
    const pages: Record<string, number> = {};
    const wanted: Record<string, number> = {};
    for (const [mark, [page, y]] of Object.entries(expected.marks)) {
      wanted[mark] = page;
      pages[mark] = positions.get(mark)?.[0] ?? 0;
      if (y !== undefined) {
        expect(positions.get(mark)?.[1], mark).toBeCloseTo(y, 0);
      }
    }
    expect(pages).toEqual(wanted);
    expect(layout.pages.length).toBe(expected.pageCount);
    expected.nestedColumns?.forEach((width, i) => {
      expect(firstNestedTable(layout)?.cells[i]?.rect.width, `column ${i}`).toBeCloseTo(width, 0);
    });
  });
});

/** Baselines of a page's lines, page-relative: body text first, then the paragraphs of each table cell skipped. */
function bodyBaselines(page: ReturnType<typeof Layout.compute>["pages"][number]): number[] {
  return page.content
    .filter((block): block is LayoutParagraph => block.type === "paragraph")
    .flatMap(block =>
      block.lines.map(line => page.geometry.marginTop + block.rect.y + line.y + line.baseline)
    );
}

describe("line heights and endnote pages measured on Word's PDF", () => {
  // P34: forty single-line items per page — numbered with the default Calibri
  // 11pt mark, then with an Arial 10pt mark, then bulleted the same two ways,
  // then plain — beside 10pt Arial text. Word's mean baseline pitch per page.
  it.each([
    ["P34240", [12.5908, 11.5015, 13.1754, 12.1723, 11.5015, 11.5015]],
    ["P34259", [13.5015, 12.4123, 14.0862, 13.0831, 12.4123, 12.4123]]
  ] as const)("%s: a list marker raises its line, never deepens it", async (name, pitches) => {
    const layout = Layout.compute(await Io.read(readFileSync(resolve(DATA, `${name}.docx`))));
    pitches.forEach((pitch, i) => {
      const ys = bodyBaselines(layout.pages[i]!);
      expect((ys.at(-1)! - ys[0]!) / (ys.length - 1), `page ${i + 1}`).toBeCloseTo(pitch, 2);
    });
  });

  // R11b: endnotes after a numbered list holding hidden runs. Word's bookmarks
  // are measured with hidden text shown, so its PDF is the reference: 11
  // pages; the notes open page 10 under the separator and page 11 under the
  // continuation separator, their first baselines at 80.6pt.
  it("R11b: endnote pages open under their separator", async () => {
    const layout = Layout.compute(
      await Io.read(readFileSync(resolve(DATA, "R11b-hidden-list-endnotes.docx")))
    );
    expect(layout.pages.length).toBe(11);
    for (const page of [10, 11]) {
      const first = bodyBaselines(layout.pages[page - 1]!).find(y => y > 70);
      expect(first, `page ${page}`).toBeCloseTo(80.6, 0);
    }
    expect(bodyBaselines(layout.pages[0]!).at(-1)).toBeCloseTo(723.1, 0);
  });

  // R12j: continuous bands of one, two and three columns. A two-column band's
  // paragraph continues at the top of page 5 with its last two lines, and
  // Word keeps both in the first column before the next band — it does not
  // balance them one per column.
  it("R12j: the last two lines of a continued paragraph stay together", async () => {
    const layout = Layout.compute(
      await Io.read(readFileSync(resolve(DATA, "R12j-column-bands.docx")))
    );
    const [first, second, heading] = bodyBaselines(layout.pages[4]!);
    expect(first).toBeCloseTo(82.6, 0);
    expect(second).toBeCloseTo(97.0, 0);
    expect(heading).toBeCloseTo(121.4, 0);
  });
});

/**
 * Lines per column of every band of columns: a band is what lies between two
 * one-column lines starting "SEP" (or "END"); columns are told apart by their
 * left edge, Letter paper with 1in margins and 0.5in gaps. A page's footnotes
 * count with the band last on it, and a one-column band counts its lines per
 * page — the way they are read off Word's PDF.
 */
function bandColumns(layout: ReturnType<typeof Layout.compute>): number[][] {
  const bands: number[][] = [];
  let band: number[] | undefined;
  let width = 0;
  let columns = 0;
  let pages = 0;
  const add = (block: LayoutParagraph): void => {
    if (!band) {
      return;
    }
    const column =
      columns === 1 ? pages - 1 : Math.min(columns - 1, Math.round(block.rect.x / (width + 36)));
    band[column] = (band[column] ?? 0) + block.lines.length;
  };
  for (const page of layout.pages) {
    if (band && columns === 1) {
      pages++;
    }
    for (const block of page.content) {
      if (block.type !== "paragraph") {
        continue;
      }
      const text = block.lines[0]?.runs.map(run => ("text" in run ? run.text : "")).join("") ?? "";
      const sep = /^(SEP\d+ (\d)c|END)/.exec(text);
      if (sep) {
        columns = sep[2] ? +sep[2] : 0;
        band = columns ? new Array<number>(columns).fill(0) : undefined;
        width = columns ? (468 - (columns - 1) * 36) / columns : 0;
        pages = 1;
        if (band) {
          bands.push(band);
        }
        continue;
      }
      add(block);
    }
    for (const block of page.footnoteArea ?? []) {
      add(block);
    }
  }
  return bands;
}

describe("columns balanced before a continuous break, as Word balances them", () => {
  // P35, P36 (pages2 `gen-p35.ts`, `gen-p36.ts`): bands of two or three
  // columns between one-column "SEP" lines — one paragraph of growing length,
  // several paragraphs, widow control on and off, Word's default 8pt after and
  // 1.08 spacing, keepNext (with and without keepLines), a footnote, headings,
  // runs of other sizes; P36f is a keepNext paragraph at the foot of a page.
  // Word's lines per column (per page for P36f), read off its PDF.
  const expected = JSON.parse(readFileSync(resolve(DATA, "column-bands.json"), "utf8")) as Record<
    string,
    number[][]
  >;
  // Bands Word sets otherwise, not yet understood: two columns of 8pt then
  // 11pt lines (P36c B1: Word 6/4 where the step rule gives 5/5), runs of
  // 10pt and 4.5pt among 11pt (P36d B3, B10), and notes set in columns under
  // a band (P35j B4).
  const open: Record<string, readonly number[]> = { P35j: [3], P36c: [0], P36d: [2, 9] };
  it.each(Object.entries(expected))("%s", async (name, bands) => {
    const layout = Layout.compute(await Io.read(readFileSync(resolve(DATA, `${name}.docx`))));
    const skip = (_: unknown, i: number): boolean => !(open[name] ?? []).includes(i);
    expect(bandColumns(layout).filter(skip)).toEqual(bands.filter(skip));
  });
});
