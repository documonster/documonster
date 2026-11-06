/**
 * The pruning rules of `page-index.ts` against the plain-walk oracle (`field-pruning.oracle.ts`)
 * on field graphs built to be awkward — nodes shared between parents, /Kids cycles, fields
 * listed twice, widgets with kids of their own, annotations on several pages — for every
 * selection of pages, and the established answers for the cases the rules single out.
 */
import { buildPageIndex, pruneRemovedPages, pruneToPage } from "@pdf/builder/page-index";
import { PdfDocument } from "@pdf/reader/pdf-document";
import { createRng } from "@test/rng";
import { describe, expect, it } from "vitest";

import { expectPruningMatchesOracle } from "./field-pruning.oracle";
import { buildRawPdf } from "./raw-pdf.helpers";
import { expectSplitMatchesReference } from "./split-pages.helpers";

const CONTENT = "<< /Length 0 >>\nstream\n\nendstream";

/** Pages are objects 3…, the content stream 9, field nodes 10…; `annots[p]` lists page p's. */
function formPdf(fields: string, annots: string[][], nodes: Record<number, string>): Uint8Array {
  const bodies: string[] = [
    `<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [${fields}] >> >>`,
    `<< /Type /Pages /Kids [${annots.map((_, p) => `${3 + p} 0 R`).join(" ")}] /Count ${annots.length} >>`
  ];
  for (let p = 0; p < 6; p++) {
    bodies.push(
      p < annots.length
        ? `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 9 0 R /Annots [${annots[p].join(" ")}] >>`
        : "null"
    );
  }
  bodies.push(CONTENT);
  const max = Math.max(9, ...Object.keys(nodes).map(Number));
  for (let n = 10; n <= max; n++) {
    bodies.push(nodes[n] ?? "null");
  }
  return buildRawPdf(bodies);
}

/** Every subset of `count` pages, as the pages removed. */
function everyRemoval(count: number): Array<Set<number>> {
  const out: Array<Set<number>> = [];
  for (let mask = 1; mask < 1 << count; mask++) {
    out.push(new Set(Array.from({ length: count }, (_, i) => i).filter(i => mask & (1 << i))));
  }
  return out;
}

function randomForm(seed: number): Uint8Array {
  const rng = createRng(seed);
  const pageCount = rng.int(2, 4);
  const nodeCount = rng.int(3, 9);
  const ref = (): string => `${10 + rng.int(0, nodeCount - 1)} 0 R`;
  const annots: string[][] = Array.from({ length: pageCount }, () => []);
  const nodes: Record<number, string> = {};
  for (let i = 0; i < nodeCount; i++) {
    const num = 10 + i;
    const widget = rng.next() < 0.5;
    if (widget) {
      // On one page usually, sometimes two, sometimes listed twice on the same one.
      for (let k = rng.next() < 0.25 ? 2 : 1; k > 0; k--) {
        annots[rng.int(0, pageCount - 1)].push(`${num} 0 R`);
      }
    }
    const kidCount = rng.next() < (widget ? 0.25 : 0.7) ? rng.int(0, 3) : -1;
    const kids =
      kidCount < 0
        ? ""
        : `/Kids [${Array.from({ length: kidCount }, () =>
            rng.next() < 0.1 ? "<< /T (inline) >>" : ref()
          ).join(" ")}]`;
    nodes[num] = widget
      ? `<< /Type /Annot /Subtype /Widget /Rect [0 0 10 10] /T (n${i}) ${kids} >>`
      : `<< /T (n${i}) ${kids} >>`;
  }
  const fields = Array.from({ length: rng.int(1, 4) }, () =>
    rng.next() < 0.1 ? "<< /T (top) /Kids [10 0 R] >>" : ref()
  ).join(" ");
  return formPdf(fields, annots, nodes);
}

describe("page-index pruning agrees with the plain walk", () => {
  for (let seed = 1; seed <= 300; seed++) {
    it(`random field graph ${seed}, every selection of pages`, () => {
      const bytes = randomForm(seed);
      const count = new PdfDocument(bytes).getPagesWithObjInfo().length;
      expectPruningMatchesOracle(bytes, "", everyRemoval(count));
    });
  }

  for (const seed of [3, 17, 42, 99, 256]) {
    it(`random field graph ${seed}: splitPages() writes the reference`, async () => {
      await expectSplitMatchesReference(randomForm(seed));
    });
  }
});

describe("page-index pruning: the established answers", () => {
  const widget = (extra = ""): string =>
    `<< /Type /Annot /Subtype /Widget /Rect [0 0 10 10] ${extra} >>`;
  /** What removing `removed` leaves out, among `keys`, and the /Fields it writes. */
  function removing(bytes: Uint8Array, removed: number[], keys: number[]) {
    const doc = new PdfDocument(bytes);
    const index = buildPageIndex(doc, doc.getPagesWithObjInfo());
    const pruning = pruneRemovedPages(index, new Set(removed));
    const show = (list: unknown[] | undefined) =>
      list?.map(item => (item as { objNum: number }).objNum);
    return {
      fields: show(pruning.fields()),
      dropped: keys.filter(k => pruning.isDropped(`${k}:0`)),
      kids: Object.fromEntries(
        keys.flatMap(k => {
          const kids = pruning.kidsOverride(`${k}:0`);
          return kids ? [[k, show(kids)]] : [];
        })
      )
    };
  }

  it("a /Kids cycle of kept nodes keeps itself, though no widget under it survives", () => {
    // 10 ⇄ 11, and 11 holds the widget on page 1.
    const bytes = formPdf("10 0 R", [[], ["12 0 R"]], {
      10: "<< /T (a) /Kids [11 0 R] >>",
      11: "<< /T (b) /Kids [10 0 R 12 0 R] >>",
      12: widget()
    });
    expect(removing(bytes, [1], [10, 11, 12])).toEqual({
      fields: [10],
      dropped: [12],
      kids: { 11: [10] }
    });
    expectPruningMatchesOracle(bytes, "", everyRemoval(2));
  });

  it("a cycle through a dropped widget does not keep itself", () => {
    // 10 (widget, page 1) → 11 → 10: with page 1 removed, 11's only kid is gone.
    const bytes = formPdf("10 0 R 11 0 R", [[], ["10 0 R"]], {
      10: widget("/Kids [11 0 R]"),
      11: "<< /T (b) /Kids [10 0 R] >>"
    });
    expect(removing(bytes, [1], [10, 11])).toEqual({ fields: [], dropped: [10, 11], kids: {} });
    expectPruningMatchesOracle(bytes, "", everyRemoval(2));
  });

  it("a kept widget with kids is dropped when none of its kids survives", () => {
    const bytes = formPdf("10 0 R", [["10 0 R"], ["11 0 R"]], {
      10: widget("/Kids [11 0 R]"),
      11: widget()
    });
    expect(removing(bytes, [1], [10, 11])).toEqual({ fields: [], dropped: [10, 11], kids: {} });
    expectPruningMatchesOracle(bytes, "", everyRemoval(2));
  });

  it("a node reached only through a dropped widget is left as it is", () => {
    // 11 sits under the widget 10 on page 1, and has only page 1's widget 12 beneath it.
    const bytes = formPdf("10 0 R", [[], ["10 0 R 12 0 R"]], {
      10: widget("/Kids [11 0 R]"),
      11: "<< /T (inner) /Kids [12 0 R] >>",
      12: widget()
    });
    expect(removing(bytes, [1], [10, 11, 12])).toEqual({ fields: [], dropped: [10, 12], kids: {} });
    expectPruningMatchesOracle(bytes, "", everyRemoval(2));
  });

  it("a shared node and a field listed twice keep every listing that survives", () => {
    const bytes = formPdf("10 0 R 12 0 R 10 0 R", [["13 0 R"], ["14 0 R"]], {
      10: "<< /T (a) /Kids [11 0 R 14 0 R] >>",
      11: "<< /T (shared) /Kids [13 0 R 14 0 R] >>",
      12: "<< /T (b) /Kids [11 0 R] >>",
      13: widget(),
      14: widget()
    });
    expect(removing(bytes, [1], [10, 11, 12, 13, 14])).toEqual({
      fields: [10, 12, 10],
      dropped: [14],
      kids: { 10: [11], 11: [13] }
    });
    expectPruningMatchesOracle(bytes, "", everyRemoval(2));
  });

  it("a split part answers every page from one evaluation", () => {
    const bytes = formPdf("10 0 R", [["11 0 R"], ["12 0 R"], ["13 0 R"]], {
      10: "<< /T (group) /Kids [11 0 R 12 0 R 13 0 R] >>",
      11: widget(),
      12: widget(),
      13: widget()
    });
    const doc = new PdfDocument(bytes);
    const index = buildPageIndex(doc, doc.getPagesWithObjInfo());
    const kidsOf = (page: number) =>
      pruneToPage(index, page)
        .kidsOverride("10:0")
        ?.map(kid => (kid as { objNum: number }).objNum);
    expect([0, 1, 2].map(kidsOf)).toEqual([[11], [12], [13]]);
    expect(pruneToPage(index, -1).isDropped("10:0")).toBe(true);
  });
});
