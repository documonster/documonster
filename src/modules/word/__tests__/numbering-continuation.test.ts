/**
 * List numbering across interruptions and instances.
 *
 * Word counts per numbering instance (`numId`) and level: an item after an
 * interrupting paragraph continues its instance's count, an adjacent item of
 * another instance starts a list of its own, and a deeper level restarts when
 * a shallower one advances. CommonMark and HTML number a list from its first
 * marker, so the converters must emit the real numbers and split lists where
 * Word does. The Markdown is parsed back with the repository's importer to
 * prove it means what was intended.
 */

import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { docxToSemantic } from "@word/convert/docx-to-semantic";
import { renderToHtml } from "@word/convert/html/html-renderer";
import { markdownToDocx } from "@word/convert/markdown/markdown-import";
import { renderToMarkdown } from "@word/convert/markdown/markdown-renderer";
import { readOdt, writeOdt } from "@word/convert/odt/odt";
import { layoutDocumentFull } from "@word/layout/layout-full";
import { createNumberingCounter } from "@word/query/numbering-counter";
import { resolveNumberingLevelDef, resolveParagraphNumbering } from "@word/query/style-resolve";
import type { DocxDocument, NumberingRef, Paragraph, Run } from "@word/types";
import { describe, it, expect } from "vitest";

function para(text: string, numbering?: NumberingRef): Paragraph {
  return {
    type: "paragraph",
    properties: numbering ? { numbering } : {},
    children: [{ content: [{ type: "text", text }] } as Run]
  };
}

const decimal = (abstractNumId: number) => ({
  abstractNumId,
  levels: [
    { level: 0, format: "decimal", text: "%1.", start: 1 },
    { level: 1, format: "lowerLetter", text: "%2.", start: 1 }
  ]
});

/** numId 1 and 2 are independent decimal lists; numId 3 restarts at 5. */
function makeDoc(body: DocxDocument["body"]): DocxDocument {
  return {
    body,
    styles: [],
    abstractNumberings: [decimal(1), decimal(2)],
    numberingInstances: [
      { numId: 1, abstractNumId: 1 },
      { numId: 2, abstractNumId: 2 },
      { numId: 3, abstractNumId: 2, overrides: [{ level: 0, startOverride: 5 }] }
    ],
    headers: new Map(),
    footers: new Map(),
    footnotes: [],
    endnotes: [],
    comments: [],
    images: [],
    fonts: [],
    embeddedFonts: [],
    customXmlParts: [],
    customProperties: [],
    opaqueParts: []
  } as unknown as DocxDocument;
}

/** Text and displayed number of every ordered list paragraph, in order. */
function displayed(doc: DocxDocument): string[] {
  const counter = createNumberingCounter(doc);
  const out: string[] = [];
  for (const block of doc.body) {
    if (block.type !== "paragraph") {
      continue;
    }
    const ref = resolveParagraphNumbering(doc, block);
    if (!ref) {
      continue;
    }
    // Bullets count too: they restart the levels beneath them.
    const value = counter.next(ref.numId, ref.level);
    if (resolveNumberingLevelDef(doc, ref.numId, ref.level)?.format === "bullet") {
      continue;
    }
    const text = block.children
      .flatMap(c => ("content" in c ? (c as Run).content : []))
      .map(c => (c.type === "text" ? c.text : ""))
      .join("");
    out.push(`${"  ".repeat(ref.level)}${value} ${text}`);
  }
  return out;
}

const interrupted = (): DocxDocument =>
  makeDoc([
    para("one", { numId: 1, level: 0 }),
    para("two", { numId: 1, level: 0 }),
    para("Interruption"),
    para("three", { numId: 1, level: 0 })
  ]);

const adjacent = (): DocxDocument =>
  makeDoc([
    para("a1", { numId: 1, level: 0 }),
    para("a2", { numId: 1, level: 0 }),
    para("b1", { numId: 2, level: 0 }),
    para("c5", { numId: 3, level: 0 })
  ]);

const nested = (): DocxDocument =>
  makeDoc([
    para("one", { numId: 1, level: 0 }),
    para("a", { numId: 1, level: 1 }),
    para("b", { numId: 1, level: 1 }),
    para("two", { numId: 1, level: 0 }),
    para("a again", { numId: 1, level: 1 })
  ]);

describe("interrupted list of one instance continues", () => {
  it("Markdown numbers the resumed item 3 and round-trips", async () => {
    const md = renderToMarkdown(interrupted());
    expect(md).toContain("3. three");
    expect(displayed(await markdownToDocx(md))).toEqual(["1 one", "2 two", "3 three"]);
  });

  it("HTML reopens the list with start=3", () => {
    const { html } = renderToHtml(interrupted());
    expect(html).toMatch(/<\/p>\s*<ol start="3">\s*<li>\s*three/);
  });

  it("semantic IR starts the resumed list at 3", () => {
    const lists = docxToSemantic(interrupted()).document.blocks.filter(b => b.type === "list");
    expect(lists).toHaveLength(2);
    expect(lists[1]).toMatchObject({ ordered: true, start: 3 });
  });

  it("ODT marks the resumed list continue-numbering and reads it back", async () => {
    const bytes = await writeOdt(interrupted());
    const xml = new TextDecoder().decode((await extractAll(bytes)).get("content.xml")!.data);
    expect(xml).toMatch(/<text:list[^>]*text:continue-numbering="true"/);
    expect(displayed(await readOdt(bytes))).toEqual(["1 one", "2 two", "3 three"]);
  });
});

describe("adjacent lists of different instances stay separate", () => {
  it("Markdown separates them and round-trips the starts", async () => {
    const md = renderToMarkdown(adjacent());
    expect(md).toMatch(/2\. a2\n\n<!-- -->\n\n1\. b1\n\n<!-- -->\n\n5\. c5/);
    expect(displayed(await markdownToDocx(md))).toEqual(["1 a1", "2 a2", "1 b1", "5 c5"]);
  });

  it("HTML emits separate <ol>s", () => {
    const { html } = renderToHtml(adjacent());
    expect(html.match(/<ol/g)).toHaveLength(3);
    expect(html).toMatch(/<\/ol>\s*<ol start="5">\s*<li>\s*c5/);
  });

  it("semantic IR emits one list per instance", () => {
    const lists = docxToSemantic(adjacent()).document.blocks.filter(b => b.type === "list");
    expect(lists).toHaveLength(3);
    expect(lists[0]).not.toHaveProperty("start");
    expect(lists[1]).not.toHaveProperty("start");
    expect(lists[2]).toMatchObject({ start: 5 });
  });
});

describe("nested levels restart when the parent advances", () => {
  it("Markdown, HTML and the importer agree", async () => {
    const md = renderToMarkdown(nested());
    expect(md).toContain("\n   1. a again");
    expect(renderToHtml(nested()).html).not.toContain("start=");
    expect(displayed(await markdownToDocx(md))).toEqual([
      "1 one",
      "  1 a",
      "  2 b",
      "2 two",
      "  1 a again"
    ]);
  });
});

describe("ODT start values", () => {
  it("round-trips startOverride 5 through the list style", async () => {
    const doc = makeDoc([para("x", { numId: 3, level: 0 }), para("y", { numId: 3, level: 0 })]);
    const restored = await readOdt(await writeOdt(doc));
    const p = restored.body[0] as Paragraph;
    const ref = p.properties!.numbering!;
    expect(resolveNumberingLevelDef(restored, ref.numId, ref.level)?.start).toBe(5);
    expect(displayed(restored)).toEqual(["5 x", "6 y"]);
  });

  /** Write `doc`, rewrite its content.xml, and read the result back. */
  async function rewritten(
    doc: DocxDocument,
    edit: (xml: string) => string
  ): Promise<DocxDocument> {
    const entries = await extractAll(await writeOdt(doc));
    const dec = new TextDecoder();
    const enc = new TextEncoder();
    return readOdt(
      await createZip(
        [...entries].map(([name, e]) => ({
          name,
          data: name === "content.xml" ? enc.encode(edit(dec.decode(e.data))) : e.data
        }))
      )
    );
  }

  it("restarts a second list of the same style unless it continues", async () => {
    const restored = await rewritten(interrupted(), xml =>
      xml.replace(/ text:continue-numbering="true"/, "")
    );
    expect(displayed(restored)).toEqual(["1 one", "2 two", "1 three"]);
  });

  it("honours text:start-value on a list item", async () => {
    const doc = makeDoc([
      para("one", { numId: 1, level: 0 }),
      para("ten", { numId: 1, level: 0 }),
      para("eleven", { numId: 1, level: 0 })
    ]);
    let seen = 0;
    const restored = await rewritten(doc, xml =>
      xml.replace(/<text:list-item>/g, m =>
        ++seen === 2 ? '<text:list-item text:start-value="10">' : m
      )
    );
    expect(displayed(restored)).toEqual(["1 one", "10 ten", "11 eleven"]);
  });
});

/** List marker texts the page layout (the Word → PDF path) emits, in order. */
function layoutMarkers(doc: DocxDocument): string[] {
  const markers: string[] = [];
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      value.forEach(visit);
    } else if (value && typeof value === "object") {
      const text = (value as { text?: unknown }).text;
      if (typeof text === "string" && /^\S+ {2}$/.test(text)) {
        markers.push(text.trim());
      }
      Object.values(value).forEach(visit);
    }
  };
  visit(layoutDocumentFull(doc).pages);
  return markers;
}

describe("page layout numbers lists by the same rules as the converters", () => {
  it("continues an interrupted instance (1, 2, 3) like Markdown, HTML and the IR", () => {
    expect(layoutMarkers(interrupted())).toEqual(["1.", "2.", "3."]);
  });

  it("separates adjacent instances and honours startOverride", () => {
    expect(layoutMarkers(adjacent())).toEqual(["1.", "2.", "1.", "5."]);
  });

  it("restarts a deeper level when the shallower one advances", () => {
    expect(layoutMarkers(nested())).toEqual(["1.", "a.", "b.", "2.", "a."]);
  });

  it("numbers a paragraph whose numbering comes from its style", () => {
    const doc = {
      ...makeDoc([
        para("one", { numId: 1, level: 0 }),
        { ...para("styled"), properties: { style: "ListStyled" } }
      ]),
      styles: [
        {
          type: "paragraph",
          styleId: "ListStyled",
          paragraphProperties: { numbering: { numId: 1, level: 0 } }
        }
      ]
    } as unknown as DocxDocument;
    expect(layoutMarkers(doc)).toEqual(["1.", "2."]);
    expect(renderToMarkdown(doc)).toContain("2. styled");
  });

  it("fills each %N of a multi-level template from its own level", () => {
    const base = makeDoc([
      para("one", { numId: 1, level: 0 }),
      para("two", { numId: 1, level: 0 }),
      para("two-a", { numId: 1, level: 1 }),
      para("two-b", { numId: 1, level: 1 })
    ]);
    const doc = {
      ...base,
      abstractNumberings: [
        {
          abstractNumId: 1,
          levels: [
            { level: 0, format: "decimal", text: "%1.", start: 1 },
            { level: 1, format: "decimal", text: "%1.%2.", start: 1 }
          ]
        }
      ]
    } as unknown as DocxDocument;
    expect(layoutMarkers(doc)).toEqual(["1.", "2.", "2.1.", "2.2."]);
  });

  it("restarts the ordered level beneath a bullet that advances", async () => {
    const base = makeDoc([
      para("bullet one", { numId: 4, level: 0 }),
      para("x", { numId: 4, level: 1 }),
      para("y", { numId: 4, level: 1 }),
      para("bullet two", { numId: 4, level: 0 }),
      para("x again", { numId: 4, level: 1 })
    ]);
    const doc = {
      ...base,
      abstractNumberings: [
        ...base.abstractNumberings!,
        {
          abstractNumId: 4,
          levels: [
            { level: 0, format: "bullet", text: "•", start: 1 },
            { level: 1, format: "decimal", text: "%2.", start: 1 }
          ]
        }
      ],
      numberingInstances: [...base.numberingInstances!, { numId: 4, abstractNumId: 4 }]
    } as unknown as DocxDocument;
    expect(layoutMarkers(doc)).toEqual(["•", "1.", "2.", "•", "1."]);
    const expected = ["  1 x", "  2 y", "  1 x again"];
    expect(displayed(doc)).toEqual(expected);
    expect(displayed(await markdownToDocx(renderToMarkdown(doc)))).toEqual(expected);
    expect(renderToHtml(doc).html.match(/<ol[^>]*>/g)).toEqual(["<ol>", "<ol>"]);
  });
});

describe("Word numbering rules shared by every consumer", () => {
  const withLevels = (
    levels: unknown[],
    instances: unknown[],
    body: DocxDocument["body"]
  ): DocxDocument =>
    ({
      ...makeDoc(body),
      abstractNumberings: [{ abstractNumId: 7, levels }],
      numberingInstances: instances
    }) as unknown as DocxDocument;

  it("w:lvlRestart=0 keeps a deeper level counting across shallower items", () => {
    const doc = withLevels(
      [
        { level: 0, format: "decimal", text: "%1.", start: 1 },
        { level: 1, format: "decimal", text: "(%2)", start: 1, restartAfterLevel: 0 }
      ],
      [{ numId: 1, abstractNumId: 7 }],
      [
        para("one", { numId: 1, level: 0 }),
        para("a", { numId: 1, level: 1 }),
        para("two", { numId: 1, level: 0 }),
        para("b", { numId: 1, level: 1 })
      ]
    );
    expect(layoutMarkers(doc)).toEqual(["1.", "(1)", "2.", "(2)"]);
    expect(displayed(doc)).toEqual(["1 one", "  1 a", "2 two", "  2 b"]);
  });

  it("w:lvlRestart=n restarts only after level n or above", () => {
    const doc = withLevels(
      [
        { level: 0, format: "decimal", text: "%1.", start: 1 },
        { level: 1, format: "decimal", text: "%2.", start: 1 },
        { level: 2, format: "decimal", text: "[%3]", start: 1, restartAfterLevel: 1 }
      ],
      [{ numId: 1, abstractNumId: 7 }],
      [
        para("one", { numId: 1, level: 0 }),
        para("a", { numId: 1, level: 1 }),
        para("x", { numId: 1, level: 2 }),
        para("b", { numId: 1, level: 1 }),
        para("y", { numId: 1, level: 2 }),
        para("two", { numId: 1, level: 0 }),
        para("z", { numId: 1, level: 2 })
      ]
    );
    // Level 1 advancing does not restart level 2; level 0 advancing does.
    expect(layoutMarkers(doc)).toEqual(["1.", "1.", "[1]", "2.", "[2]", "2.", "[1]"]);
  });

  it("instances sharing an abstract definition share one count", () => {
    const doc = withLevels(
      [{ level: 0, format: "decimal", text: "%1.", start: 1 }],
      [
        { numId: 1, abstractNumId: 7 },
        { numId: 2, abstractNumId: 7 },
        { numId: 3, abstractNumId: 7, overrides: [{ level: 0, startOverride: 1 }] }
      ],
      [
        para("one", { numId: 1, level: 0 }),
        para("two", { numId: 2, level: 0 }),
        para("restart", { numId: 3, level: 0 })
      ]
    );
    expect(layoutMarkers(doc)).toEqual(["1.", "2.", "1."]);
    expect(renderToMarkdown(doc)).toMatch(/2\. two/);
    expect(renderToHtml(doc).html).toContain('<ol start="2">');
  });

  it("w:isLgl shows every referenced level as an arabic numeral", () => {
    const doc = withLevels(
      [
        { level: 0, format: "upperRoman", text: "%1.", start: 1 },
        { level: 1, format: "lowerLetter", text: "%1.%2", start: 1, isLegalNumberingStyle: true }
      ],
      [{ numId: 1, abstractNumId: 7 }],
      [
        para("one", { numId: 1, level: 0 }),
        para("two", { numId: 1, level: 0 }),
        para("two-a", { numId: 1, level: 1 })
      ]
    );
    expect(layoutMarkers(doc)).toEqual(["I.", "II.", "2.1"]);
  });
});

/**
 * Expectations here are Microsoft Word's own output: the same numbering was
 * built as a .docx, exported to PDF by Word, and its list markers read back.
 */
describe("instances sharing an abstract definition (matches Word)", () => {
  function shared(instances: DocxDocument["numberingInstances"]): string[] {
    const doc = makeDoc([
      para("a", { numId: 10, level: 0 }),
      para("b", { numId: 10, level: 0 }),
      para("c", { numId: 11, level: 0 }),
      para("d", { numId: 10, level: 0 })
    ]);
    return displayed({ ...doc, numberingInstances: instances } as DocxDocument);
  }

  it("(a) without an override, continues one count across numIds", () => {
    expect(
      shared([
        { numId: 10, abstractNumId: 1 },
        { numId: 11, abstractNumId: 1 }
      ])
    ).toEqual(["1 a", "2 b", "3 c", "4 d"]);
  });

  it("(b) a startOverride restarts the shared count once; later items continue it", () => {
    expect(
      shared([
        { numId: 10, abstractNumId: 1 },
        { numId: 11, abstractNumId: 1, overrides: [{ level: 0, startOverride: 1 }] }
      ])
    ).toEqual(["1 a", "2 b", "1 c", "2 d"]);
  });

  it("(c) a replacement level definition changes the look, not the count", () => {
    expect(
      shared([
        { numId: 10, abstractNumId: 1 },
        {
          numId: 11,
          abstractNumId: 1,
          overrides: [
            { level: 0, levelDef: { level: 0, format: "decimal", text: "%1)", start: 7 } }
          ]
        }
      ])
    ).toEqual(["1 a", "2 b", "3 c", "4 d"]);
  });

  it("reproduces Word's numbering of a mixed document marker for marker", () => {
    const levels = [
      { level: 0, format: "decimal", text: "%1.", start: 1 },
      { level: 1, format: "lowerLetter", text: "%2.", start: 1 }
    ];
    const doc = {
      ...makeDoc([
        para("a1", { numId: 1, level: 0 }),
        para("a2", { numId: 1, level: 0 }),
        para("b-num4", { numId: 4, level: 0 }),
        para("c-num1", { numId: 1, level: 0 }),
        para("d-num4-again", { numId: 4, level: 0 }),
        para("e-num1", { numId: 1, level: 0 }),
        para("f-num5-lvldef", { numId: 5, level: 0 }),
        para("g-num1", { numId: 1, level: 0 }),
        para("h-num5-again", { numId: 5, level: 0 }),
        para("i-num1-sub", { numId: 1, level: 1 }),
        para("j-num6-sub", { numId: 6, level: 1 }),
        para("k-num6-sub", { numId: 6, level: 1 }),
        para("l-num1-sub", { numId: 1, level: 1 }),
        para("m-num6-top", { numId: 6, level: 0 }),
        para("n-num6-sub", { numId: 6, level: 1 })
      ]),
      abstractNumberings: [{ abstractNumId: 0, levels }],
      numberingInstances: [
        { numId: 1, abstractNumId: 0 },
        { numId: 4, abstractNumId: 0, overrides: [{ level: 0, startOverride: 7 }] },
        {
          numId: 5,
          abstractNumId: 0,
          overrides: [
            { level: 0, levelDef: { level: 0, format: "upperRoman", text: "%1)", start: 20 } }
          ]
        },
        { numId: 6, abstractNumId: 0, overrides: [{ level: 1, startOverride: 5 }] }
      ]
    } as unknown as DocxDocument;
    // Word: 1. 2. 7. 8. 9. 10. XI) 12. XIII) a. e. f. g. 14. e.
    expect(layoutMarkers(doc)).toEqual([
      "1.",
      "2.",
      "7.",
      "8.",
      "9.",
      "10.",
      "XI)",
      "12.",
      "XIII)",
      "a.",
      "e.",
      "f.",
      "g.",
      "14.",
      "e."
    ]);
  });
});
