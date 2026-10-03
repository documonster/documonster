/**
 * Every DOCX consumer — HTML, Markdown, the semantic IR and ODT — must agree on
 * what a paragraph *means*: whether it is a heading and at which level, which
 * text the final view shows, and which formatting its styles give it. Each
 * case here renders one model through all four and compares the answers.
 */

import { extractAll } from "@archive/unzip/extract";
import { docxToSemantic } from "@word/convert/docx-to-semantic";
import { renderToHtml } from "@word/convert/html/html-renderer";
import { renderToMarkdown } from "@word/convert/markdown/markdown-renderer";
import { writeOdt } from "@word/convert/odt/odt";
import { resolveHeadingLevel } from "@word/query/heading";
import { resolveRunStyle, resolveStyle } from "@word/query/style-resolve";
import type {
  DocxDocument,
  Paragraph,
  ParagraphChild,
  ParagraphProperties,
  Run,
  RunProperties,
  StyleDef
} from "@word/types";
import { parseXml, textContent, walk } from "@xml/dom";
import type { XmlElement } from "@xml/types";
import { describe, expect, it } from "vitest";

const REV = { author: "A", id: 1 };

function run(text: string, properties?: RunProperties): Run {
  return { properties, content: [{ type: "text", text }] };
}

function doc(
  children: readonly ParagraphChild[],
  props?: ParagraphProperties,
  styles: StyleDef[] = []
): DocxDocument {
  const para: Paragraph = { type: "paragraph", properties: props, children };
  return { body: [para], styles } as DocxDocument;
}

function style(styleId: string, extra: Partial<StyleDef> = {}): StyleDef {
  return { type: "paragraph", styleId, name: styleId, ...extra };
}

async function odtContent(d: DocxDocument): Promise<string> {
  const entries = await extractAll(await writeOdt(d));
  return new TextDecoder().decode(entries.get("content.xml")!.data);
}

function htmlOf(d: DocxDocument): string {
  return renderToHtml(d, { fullDocument: false, includeStyles: false }).html;
}

interface Views {
  readonly html: string;
  readonly md: string;
  readonly semantic: ReturnType<typeof docxToSemantic>["document"]["blocks"];
  readonly odt: string;
}

async function views(d: DocxDocument): Promise<Views> {
  return {
    html: htmlOf(d),
    md: renderToMarkdown(d),
    semantic: docxToSemantic(d).document.blocks,
    odt: await odtContent(d)
  };
}

async function headingLevels(d: DocxDocument): Promise<number[]> {
  const v = await views(d);
  const block = v.semantic[0];
  return [
    Number(/<h(\d)/.exec(v.html)?.[1] ?? 0),
    /^(#+) /.exec(v.md)?.[1].length ?? 0,
    block?.type === "heading" ? block.level : 0,
    Number(/<text:h[^>]*text:outline-level="(\d)"/.exec(v.odt)?.[1] ?? 0)
  ];
}

function semanticText(blocks: Views["semantic"]): string {
  return JSON.stringify(blocks)
    .match(/"text":"[^"]*"/g)!
    .map(t => t.slice(8, -1))
    .join("");
}

/**
 * Text of a markup string, read by a real parser rather than by deleting
 * tag-shaped substrings: a regex cannot know where a tag ends (a `>` inside an
 * attribute value) and leaves entities undecoded.
 */
function markupText(markup: string): string {
  return textContent(parseXml(`<root>${markup}</root>`).root);
}

function odtBodyText(xml: string): string {
  let body: XmlElement | undefined;
  walk(parseXml(xml).root, el => {
    if (el.name === "office:text") {
      body ??= el;
    }
  });
  return body ? textContent(body) : "";
}

async function texts(d: DocxDocument): Promise<string[]> {
  const v = await views(d);
  return [
    markupText(v.html).replace(/\s+/g, ""),
    v.md.trim(),
    semanticText(v.semantic),
    odtBodyText(v.odt)
  ];
}

describe("headings resolve identically in every converter", () => {
  it("lets a direct outline level win over the style", async () => {
    const d = doc([run("H")], { style: "Heading2", outlineLevel: 0 }, [
      style("Heading2", { name: "heading 2" })
    ]);
    expect(resolveHeadingLevel(d, d.body[0] as Paragraph)).toEqual({ level: 1, kind: "heading" });
    expect(await headingLevels(d)).toEqual([1, 1, 1, 1]);
  });

  it("renders Title as a level-1 heading", async () => {
    const d = doc([run("T")], { style: "Title" }, [style("Title")]);
    expect(resolveHeadingLevel(d, d.body[0] as Paragraph)).toEqual({ level: 1, kind: "title" });
    expect(await headingLevels(d)).toEqual([1, 1, 1, 1]);
  });

  it("inherits the outline level through basedOn", async () => {
    const d = doc([run("C")], { style: "Custom" }, [
      style("MyHeading", { name: "My Heading", paragraphProperties: { outlineLevel: 2 } }),
      style("Custom", { name: "Custom", basedOn: "MyHeading" })
    ]);
    expect(await headingLevels(d)).toEqual([3, 3, 3, 3]);
  });

  it("treats a built-in heading style without outlineLvl as a heading, ODT included", async () => {
    const d = doc([run("H")], { style: "Heading1" }, [style("Heading1", { name: "heading 1" })]);
    expect(await headingLevels(d)).toEqual([1, 1, 1, 1]);
  });

  it("recognises a localised style id by its built-in name", async () => {
    const d = doc([run("H")], { style: "berschrift3" }, [
      style("berschrift3", { name: "heading 3" })
    ]);
    expect(await headingLevels(d)).toEqual([3, 3, 3, 3]);
  });

  it("lets outline level 9 (body text) override a heading style", async () => {
    const d = doc([run("B")], { style: "Heading1", outlineLevel: 9 }, [
      style("Heading1", { name: "heading 1" })
    ]);
    expect(resolveHeadingLevel(d, d.body[0] as Paragraph)).toBeUndefined();
    expect(await headingLevels(d)).toEqual([0, 0, 0, 0]);
  });

  it("lets a base style's outline level 9 override a built-in heading name further down", async () => {
    const d = doc([run("B")], { style: "Heading2" }, [
      style("Body", { name: "Body", paragraphProperties: { outlineLevel: 9 } }),
      style("Heading2", { name: "heading 2", basedOn: "Body" })
    ]);
    expect(resolveHeadingLevel(d, d.body[0] as Paragraph)).toBeUndefined();
    expect(await headingLevels(d)).toEqual([0, 0, 0, 0]);
  });

  it("falls back to the nearest built-in name only when no style sets an outline level", () => {
    const d = doc([run("H")], { style: "Custom" }, [
      style("Heading1", { name: "heading 1" }),
      style("Heading3", { name: "heading 3", basedOn: "Heading1" }),
      style("Custom", { name: "Custom", basedOn: "Heading3" })
    ]);
    expect(resolveHeadingLevel(d, d.body[0] as Paragraph)).toEqual({ level: 3, kind: "heading" });
    const withOutline = doc([run("H")], { style: "Heading3" }, [
      style("Base", { name: "Base", paragraphProperties: { outlineLevel: 4 } }),
      style("Heading3", { name: "heading 3", basedOn: "Base" })
    ]);
    expect(resolveHeadingLevel(withOutline, withOutline.body[0] as Paragraph)).toEqual({
      level: 5,
      kind: "heading"
    });
  });

  it("clamps levels 7–9 to six where the format has six, and keeps them in ODT", async () => {
    const d = doc([run("H")], { outlineLevel: 7 });
    expect(await headingLevels(d)).toEqual([6, 6, 6, 8]);
  });
});

describe("final view of tracked changes and hidden text", () => {
  it("shows inserted and moved-to text, hides deleted, moved-from and vanish", async () => {
    const d = doc([
      run("a"),
      { type: "insertedRun", revision: REV, run: run("b") },
      { type: "deletedRun", revision: REV, run: run("X") },
      { type: "movedFromRun", revision: REV, run: run("Y") },
      { type: "movedToRun", revision: REV, run: run("c") },
      run("Z", { vanish: true })
    ]);
    expect(await texts(d)).toEqual(["abc", "abc", "abc", "abc"]);
  });

  it("hides text a character style makes vanish", async () => {
    const d = doc([run("a"), run("Z", { style: "Hidden" })], undefined, [
      { type: "character", styleId: "Hidden", name: "Hidden", runProperties: { vanish: true } }
    ]);
    expect(await texts(d)).toEqual(["a", "a", "a", "a"]);
  });

  it("does not show deleted text in HTML just because comments are included", () => {
    const d = doc([run("a"), { type: "deletedRun", revision: REV, run: run("X") }]);
    expect(renderToHtml(d, { fullDocument: false, includeComments: true }).html).not.toContain("X");
    expect(renderToHtml(d, { fullDocument: false, includeRevisions: true }).html).toContain("<del");
  });
});

describe("run formatting from styles", () => {
  const boldPara = style("Strong", { name: "Strong para", runProperties: { bold: true } });
  const boldChar: StyleDef = {
    type: "character",
    styleId: "BoldChar",
    name: "Bold Char",
    runProperties: { bold: true }
  };

  async function bolds(d: DocxDocument): Promise<boolean[]> {
    const v = await views(d);
    return [
      v.html.includes("<strong>"),
      v.md.includes("**"),
      JSON.stringify(v.semantic).includes('"bold":true')
    ];
  }

  it("applies bold inherited from the paragraph style", async () => {
    expect(await bolds(doc([run("x")], { style: "Strong" }, [boldPara]))).toEqual([
      true,
      true,
      true
    ]);
  });

  it("toggles: bold paragraph style + bold character style is not bold", async () => {
    const d = doc([run("x", { style: "BoldChar" })], { style: "Strong" }, [boldPara, boldChar]);
    expect(await bolds(d)).toEqual([false, false, false]);
  });

  it("direct bold is absolute, not a toggle", async () => {
    const d = doc([run("x", { style: "BoldChar", bold: true })], { style: "Strong" }, [
      boldPara,
      boldChar
    ]);
    expect(await bolds(d)).toEqual([true, true, true]);
  });

  it("ODT keeps the character style as the parent of the run's automatic style", async () => {
    const xml = await odtContent(doc([run("x", { style: "BoldChar" })], undefined, [boldChar]));
    expect(xml).toMatch(/style:family="text" style:parent-style-name="BoldChar"/);
  });
});

describe("toggle properties (ECMA-376 §17.7.3)", () => {
  const charStyle = (styleId: string, rPr: RunProperties, basedOn?: string): StyleDef => ({
    type: "character",
    styleId,
    name: styleId,
    basedOn,
    runProperties: rPr
  });
  function resolved(
    styles: StyleDef[],
    paraStyle: string | undefined,
    runProps: RunProperties | undefined,
    docDefaults?: RunProperties
  ): RunProperties {
    const d = {
      body: [],
      styles,
      docDefaults: docDefaults ? { runProperties: docDefaults } : undefined
    } as unknown as DocxDocument;
    const para: Paragraph = {
      type: "paragraph",
      properties: paraStyle ? { style: paraStyle } : {},
      children: []
    };
    return resolveRunStyle(d, run("x", runProps), resolveStyle(d, para).runProperties)
      .runProperties;
  }

  it("inherits, not toggles, within a paragraph style's basedOn chain", () => {
    const styles = [
      style("Base", { runProperties: { bold: true } }),
      style("Derived", { basedOn: "Base", runProperties: { bold: true } }),
      style("Plain", { basedOn: "Base" })
    ];
    expect(resolved(styles, "Derived", undefined).bold).toBe(true);
    expect(resolved(styles, "Plain", undefined).bold).toBe(true);
  });

  it("inherits, not toggles, within a character style's basedOn chain", () => {
    const styles = [
      charStyle("CBase", { italic: true }),
      charStyle("CDerived", { italic: true }, "CBase"),
      charStyle("COff", { italic: false }, "CBase")
    ];
    expect(resolved(styles, undefined, { style: "CDerived" }).italic).toBe(true);
    expect(resolved(styles, undefined, { style: "COff" }).italic).toBe(false);
  });

  it("toggles between the paragraph-style and character-style levels", () => {
    const styles = [
      style("P", { runProperties: { bold: true, caps: true } }),
      charStyle("CBase", { bold: true }),
      charStyle("C", { caps: false }, "CBase")
    ];
    const rPr = resolved(styles, "P", { style: "C" });
    expect(rPr.bold).toBe(false); // true at both levels
    expect(rPr.caps).toBe(true); // false at a level does not flip a false default
    expect(resolved(styles, "P", undefined).bold).toBe(true);
  });

  it("with a true document default, flips once per level that sets false (MS-OI29500)", () => {
    const styles = [
      style("POn", { runProperties: { bold: true } }),
      style("POff", { runProperties: { bold: false } }),
      charStyle("COn", { bold: true }),
      charStyle("COff", { bold: false })
    ];
    const on = { bold: true };
    expect(resolved(styles, "POn", { style: "COn" }, on).bold).toBe(true);
    expect(resolved(styles, "POff", { style: "COn" }, on).bold).toBe(false);
    expect(resolved(styles, "POff", { style: "COff" }, on).bold).toBe(true);
    expect(resolved(styles, undefined, undefined, on).bold).toBe(true);
  });

  it("keeps direct formatting absolute", () => {
    const styles = [style("P", { runProperties: { bold: true } }), charStyle("C", { bold: true })];
    expect(resolved(styles, "P", { style: "C", bold: true }).bold).toBe(true);
    expect(resolved(styles, "P", { bold: false }).bold).toBe(false);
  });

  it("hides text via styles with vanish toggling like any other", () => {
    const styles = [
      style("Hidden", { runProperties: { vanish: true } }),
      charStyle("AlsoHidden", { vanish: true })
    ];
    expect(resolved(styles, "Hidden", undefined).vanish).toBe(true);
    expect(resolved(styles, "Hidden", { style: "AlsoHidden" }).vanish).toBe(false);
  });

  it("does not toggle dstrike, which is not a toggle property", () => {
    const styles = [
      style("P", { runProperties: { doubleStrike: true } }),
      charStyle("C", { doubleStrike: true })
    ];
    expect(resolved(styles, "P", { style: "C" }).doubleStrike).toBe(true);
  });
});

describe("symbols and fields", () => {
  it("renders a symbol as its character everywhere", async () => {
    const d = doc([{ content: [{ type: "symbol", font: "Symbol", char: "03A9" }] }]);
    expect(await texts(d)).toEqual(["Ω", "Ω", "Ω", "Ω"]);
  });

  it("keeps a field's cached result in every converter, ODT included", async () => {
    const d = doc([{ content: [{ type: "field", instruction: " PAGE ", cachedValue: "7" }] }]);
    expect(await texts(d)).toEqual(["7", "7", "7", "7"]);
  });
});
