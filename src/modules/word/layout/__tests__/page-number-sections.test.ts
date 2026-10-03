/**
 * Page-number fields in headers and footers across sections, by Word's rule:
 * `w:pgNumType w:start` restarts `PAGE`, a section without it continues from
 * the previous displayed number, `SECTIONPAGES` counts the page's section and
 * `NUMPAGES` the whole document. The field engine's body `PAGE` results follow
 * the same rule.
 */

import { updateFields } from "@word/advanced/field-engine";
import { layoutDocumentFull } from "@word/layout/layout-full";
import type { LayoutPage } from "@word/layout/layout-model";
import { formatListCounter } from "@word/query/page-field";
import type {
  DocxDocument,
  FooterDef,
  FootnoteDef,
  Paragraph,
  RunContent,
  SectionProperties
} from "@word/types";
import { describe, expect, it } from "vitest";

function para(
  content: RunContent[],
  sectionProperties?: SectionProperties,
  pageBreakBefore?: boolean
): Paragraph {
  const properties = {
    ...(sectionProperties ? { sectionProperties } : {}),
    ...(pageBreakBefore ? { pageBreakBefore } : {})
  };
  return {
    type: "paragraph",
    children: [{ content }],
    ...(Object.keys(properties).length > 0 ? { properties } : {})
  } as Paragraph;
}

const text = (t: string): RunContent => ({ type: "text", text: t });
const field = (name: string): RunContent => ({
  type: "field",
  instruction: ` ${name} `,
  cachedValue: "?"
});

/** Footer showing `<label> PAGE/SECTIONPAGES/NUMPAGES`. */
function footer(label: string): FooterDef {
  return {
    content: {
      children: [
        para([
          text(`${label} `),
          field("PAGE"),
          text("/"),
          field("SECTIONPAGES"),
          text("/"),
          field("NUMPAGES")
        ])
      ]
    }
  };
}

const footers = new Map<string, FooterDef>([
  ["fDefault", footer("D")],
  ["fFirst", footer("F")]
]);

function section(extra: Partial<SectionProperties> = {}): SectionProperties {
  return { footers: [{ type: "default", rId: "fDefault" }], ...extra } as SectionProperties;
}

function footerText(page: LayoutPage): string {
  return (page.footer ?? [])
    .flatMap(block => (block.type === "paragraph" ? block.lines : []))
    .flatMap(line => line.runs.map(r => (r.type === "image" ? "" : r.text)))
    .join("")
    .trim();
}

/**
 * Section 1: pages 1–2. Section 2 (no restart): continues 3–4. Section 3
 * (`start` 1, `titlePage`): restarts at 1, with the "first" footer on its own
 * first page. Every body page carries a `PAGE` field for the field engine.
 */
function threeSections(): DocxDocument {
  const body = (label: string) => [text(label), field("PAGE")];
  return {
    body: [
      para(body("s1a ")),
      para(body("s1b "), section(), true),
      para(body("s2a ")),
      para(body("s2b "), section(), true),
      para(body("s3a ")),
      para(body("s3b "), undefined, true)
    ],
    sectionProperties: section({
      pageNumbering: { start: 1 },
      titlePage: true,
      footers: [
        { type: "default", rId: "fDefault" },
        { type: "first", rId: "fFirst" }
      ]
    }),
    footers
  } as unknown as DocxDocument;
}

describe("page-number fields across sections", () => {
  it("restarts, continues and counts per section", () => {
    const pages = layoutDocumentFull(threeSections()).pages;
    expect(pages.map(footerText)).toEqual([
      "D 1/2/6",
      "D 2/2/6",
      "D 3/2/6",
      "D 4/2/6",
      "F 1/2/6",
      "D 2/2/6"
    ]);
  });

  it("field engine PAGE results agree with the layout's footers", () => {
    const d = threeSections();
    const fromFooters = layoutDocumentFull(d).pages.map(p => footerText(p).split(/[ /]/)[1]);
    const updated = updateFields(d);
    const bodyPages = updated.body.map(item => {
      const content = (item as Paragraph).children.flatMap(c =>
        "content" in c ? (c.content as RunContent[]) : []
      );
      const f = content.find(c => c.type === "field");
      return f?.type === "field" ? f.cachedValue : undefined;
    });
    expect(bodyPages).toEqual(fromFooters);
  });

  it("a footnote-overflow page belongs to the last body page's section", () => {
    // Notes each nearly a page tall: the queue outlives the body, so the
    // trailing pages hold notes only.
    const tall = Array.from({ length: 40 }, (_, i) => para([text(`note line ${i}`)]));
    const footnotes: FootnoteDef[] = [1, 2, 3].map(id => ({ id, content: tall }));
    const d = {
      body: [
        para([text("first section")], section()),
        para([
          text("second"),
          { type: "footnoteRef", id: 1 },
          { type: "footnoteRef", id: 2 },
          { type: "footnoteRef", id: 3 }
        ])
      ],
      sectionProperties: section({ pageNumbering: { start: 10 } }),
      footers,
      footnotes
    } as unknown as DocxDocument;
    const pages = layoutDocumentFull(d).pages;
    expect(pages.length).toBeGreaterThan(2);
    const sectionTwo = pages.length - 1;
    expect(pages.map(footerText)).toEqual([
      `D 1/1/${pages.length}`,
      ...pages.slice(1).map((_, i) => `D ${10 + i}/${sectionTwo}/${pages.length}`)
    ]);
    // The overflow pages use the section's geometry, not a fallback.
    expect(pages.at(-1)!.geometry).toEqual(pages[1].geometry);
  });

  // The cases below reproduce Microsoft Word's own PDF export of the same
  // documents (footer text per page, read back from Word's PDF). A section's
  // `w:type` is how *that* section starts, and parity is judged on the
  // displayed page number.
  const twoSections = (opening: Partial<SectionProperties>) =>
    ({
      body: [para([text("one")], section()), para([text("two")])],
      sectionProperties: section(opening),
      footers
    }) as unknown as DocxDocument;

  it("odd-page break, continuing numbering: a bare blank page takes number 2", () => {
    const pages = layoutDocumentFull(twoSections({ breakType: "oddPage" })).pages;
    // Word: "PG=1/3/1", "", "PG=3/3/1" (PAGE/NUMPAGES/SECTIONPAGES).
    expect(pages.map(footerText)).toEqual(["D 1/1/3", "", "D 3/1/3"]);
    expect(pages[1]!.content).toHaveLength(0);
  });

  it("odd-page break, restarting at an even number: Word skips a number, not a page", () => {
    const pages = layoutDocumentFull(
      twoSections({ breakType: "oddPage", pageNumbering: { start: 2 } })
    ).pages;
    // Word: "PG=1/2/1", "PG=3/2/1".
    expect(pages.map(footerText)).toEqual(["D 1/1/2", "D 3/1/2"]);
  });

  it("even-page break, restarting at 1: shows 2 on the section's first page", () => {
    const pages = layoutDocumentFull(
      twoSections({ breakType: "evenPage", pageNumbering: { start: 1 } })
    ).pages;
    expect(pages.map(footerText)).toEqual(["D 1/1/2", "D 2/1/2"]);
  });

  it("odd-page break that already lands on an odd number needs nothing", () => {
    const pages = layoutDocumentFull(
      twoSections({ breakType: "oddPage", pageNumbering: { start: 1 } })
    ).pages;
    expect(pages.map(footerText)).toEqual(["D 1/1/2", "D 1/1/2"]);
  });

  it("a continuous section stays on the page", () => {
    const pages = layoutDocumentFull(twoSections({ breakType: "continuous" })).pages;
    expect(pages).toHaveLength(1);
  });

  it("the break type on the closing section does not apply to the next one", () => {
    const d = {
      body: [para([text("one")], section({ breakType: "oddPage" })), para([text("two")])],
      sectionProperties: section(),
      footers
    } as unknown as DocxDocument;
    expect(layoutDocumentFull(d).pages.map(footerText)).toEqual(["D 1/1/2", "D 2/1/2"]);
  });

  it("even/odd footers follow the displayed page number, as Word does", () => {
    const d = {
      body: [para([text("one")]), para([text("two")], undefined, true)],
      sectionProperties: section({
        pageNumbering: { start: 2 },
        footers: [
          { type: "default", rId: "fDefault" },
          { type: "even", rId: "fFirst" }
        ]
      }),
      settings: { evenAndOddHeaders: true },
      footers
    } as unknown as DocxDocument;
    // Word: physical page 1 shows number 2 and takes the even footer.
    expect(layoutDocumentFull(d).pages.map(footerText)).toEqual(["F 2/2/2", "D 3/2/2"]);
  });
});

describe("alphabetic counters", () => {
  it("repeat the letter per pass through the alphabet, as Word numbers", () => {
    expect([1, 26, 27, 28, 53].map(n => formatListCounter(n, "lowerLetter"))).toEqual([
      "a",
      "z",
      "aa",
      "bb",
      "aaa"
    ]);
    expect(formatListCounter(52, "upperLetter")).toBe("ZZ");
  });

  it("fall back to digits instead of building an unbounded string", () => {
    expect(formatListCounter(2_000_000_000, "lowerLetter")).toBe("2000000000");
  });
});
