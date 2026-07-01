/**
 * DOCX Module - Builder Helpers Tests
 */

import { describe, it, expect } from "vitest";

import {
  Document,
  DocxInvalidStructureError,
  DocxUnsupportedFeatureError,
  DocxWriteError,
  Build,
  Io,
  Query,
  Theme
} from "../index";
import type { DocxDocument, MathContent, Paragraph } from "../types";

// =============================================================================
// Run content helpers
// =============================================================================

describe("Run content helpers", () => {
  it("strikethrough creates a strike run", () => {
    const run = Build.strikethrough("deleted");
    expect(run.properties?.strike).toBe(true);
    expect(run.content[0]).toEqual({ type: "text", text: "deleted" });
  });

  it("columnBreak creates a column break run", () => {
    const run = Build.columnBreak();
    expect(run.content[0]).toEqual({ type: "break", breakType: "column" });
  });

  it("carriageReturn creates a CR run", () => {
    const run = Build.carriageReturn();
    expect(run.content[0]).toEqual({ type: "carriageReturn" });
  });

  it("noBreakHyphen creates a no-break hyphen run", () => {
    const run = Build.noBreakHyphen();
    expect(run.content[0]).toEqual({ type: "noBreakHyphen" });
  });

  it("softHyphen creates a soft hyphen run", () => {
    const run = Build.softHyphen();
    expect(run.content[0]).toEqual({ type: "softHyphen" });
  });

  it("symbol creates a symbol run", () => {
    const run = Build.symbol("Wingdings", "F0FC");
    expect(run.content[0]).toEqual({ type: "symbol", font: "Wingdings", char: "F0FC" });
  });
});

// =============================================================================
// Field helpers
// =============================================================================

describe("Field helpers", () => {
  it("pageNumberField", () => {
    const run = Build.pageNumberField("3");
    expect(run.content[0]).toMatchObject({
      type: "field",
      instruction: " PAGE ",
      cachedValue: "3"
    });
  });

  it("totalPagesField", () => {
    const run = Build.totalPagesField();
    expect(run.content[0]).toMatchObject({ type: "field", instruction: " NUMPAGES " });
  });

  it("sectionPagesField", () => {
    const run = Build.sectionPagesField();
    expect(run.content[0]).toMatchObject({ type: "field", instruction: " SECTIONPAGES " });
  });

  it("sectionField", () => {
    const run = Build.sectionField();
    expect(run.content[0]).toMatchObject({ type: "field", instruction: " SECTION " });
  });

  it("dateField with format", () => {
    const run = Build.dateField("dd/MM/yyyy");
    expect((run.content[0] as any).instruction).toContain("DATE");
    expect((run.content[0] as any).instruction).toContain("dd/MM/yyyy");
  });

  it("sequenceField", () => {
    const run = Build.sequenceField("Figure");
    expect((run.content[0] as any).instruction).toContain("SEQ Figure");
  });

  it("timeField", () => {
    const run = Build.timeField("HH:mm");
    expect((run.content[0] as any).instruction).toContain("TIME");
  });

  it("authorField", () => {
    const run = Build.authorField("John");
    expect(run.content[0]).toMatchObject({
      type: "field",
      instruction: " AUTHOR ",
      cachedValue: "John"
    });
  });

  it("titleField", () => {
    const run = Build.titleField("My Doc");
    expect(run.content[0]).toMatchObject({ type: "field", instruction: " TITLE " });
  });

  it("subjectField", () => {
    const run = Build.subjectField();
    expect((run.content[0] as any).instruction).toContain("SUBJECT");
  });

  it("keywordsField", () => {
    const run = Build.keywordsField();
    expect((run.content[0] as any).instruction).toContain("KEYWORDS");
  });

  it("fileNameField", () => {
    const run = Build.fileNameField({ includePath: true });
    expect((run.content[0] as any).instruction).toContain("FILENAME");
    expect((run.content[0] as any).instruction).toContain("\\p");
  });

  it("fileSizeField", () => {
    const run = Build.fileSizeField();
    expect((run.content[0] as any).instruction).toContain("FILESIZE");
  });

  it("styleRefField", () => {
    const run = Build.styleRefField("Heading 1", { fromBottom: true });
    expect((run.content[0] as any).instruction).toContain("STYLEREF");
    expect((run.content[0] as any).instruction).toContain("\\l");
  });

  it("refField", () => {
    const run = Build.refField("_Ref123", { hyperlink: true });
    expect((run.content[0] as any).instruction).toContain("REF _Ref123");
    expect((run.content[0] as any).instruction).toContain("\\h");
  });

  it("pageRefField", () => {
    const run = Build.pageRefField("bk1");
    expect((run.content[0] as any).instruction).toContain("PAGEREF bk1");
  });

  it("noteRefField", () => {
    const run = Build.noteRefField("fn1", { hyperlink: true });
    expect((run.content[0] as any).instruction).toContain("NOTEREF fn1");
  });

  it("hyperlinkField", () => {
    const run = Build.hyperlinkField("https://example.com", { anchor: "top" });
    expect((run.content[0] as any).instruction).toContain("HYPERLINK");
    expect((run.content[0] as any).instruction).toContain("\\l");
  });

  it("quoteField", () => {
    const run = Build.quoteField("hello");
    expect((run.content[0] as any).instruction).toContain("QUOTE");
  });

  it("tocField", () => {
    const run = Build.tocField({ headingLevels: "1-3", hyperlink: true });
    expect((run.content[0] as any).instruction).toContain("TOC");
    expect((run.content[0] as any).instruction).toContain("\\h");
  });

  it("tocField does not emit \\p for tabLeader (would disable dot leader)", () => {
    // The dotted leader is the TOC default; `\p` sets a single separator char
    // and would DISABLE the leader + right-aligned page number.
    const run = Build.tocField({ headingLevels: "1-3", tabLeader: "." });
    const instr = (run.content[0] as any).instruction as string;
    expect(instr).toContain("TOC");
    expect(instr).not.toContain("\\p");
  });

  it("tocField does not emit \\z for rightAlignedPageNumbers", () => {
    // `\z` hides the leader/page numbers in Web layout view — it is NOT a
    // right-align switch. Right alignment is the TOC default via styles.
    const run = Build.tocField({ headingLevels: "1-3", rightAlignedPageNumbers: true });
    const instr = (run.content[0] as any).instruction as string;
    expect(instr).toContain("TOC");
    expect(instr).not.toContain("\\z");
  });

  it("tcField", () => {
    const run = Build.tcField("Entry", { level: 2 });
    expect((run.content[0] as any).instruction).toContain("TC");
    expect((run.content[0] as any).instruction).toContain("\\l 2");
  });

  it("indexEntryField", () => {
    const run = Build.indexEntryField("Term", { bold: true });
    expect((run.content[0] as any).instruction).toContain("XE");
    expect((run.content[0] as any).instruction).toContain("\\b");
  });

  it("indexField", () => {
    const run = Build.indexField({ columns: 2 });
    expect((run.content[0] as any).instruction).toContain("INDEX");
    expect((run.content[0] as any).instruction).toContain("\\c 2");
  });

  it("ifField", () => {
    const run = Build.ifField("1 = 1", "yes", "no");
    expect((run.content[0] as any).instruction).toContain("IF 1 = 1");
  });

  it("includeTextField", () => {
    const run = Build.includeTextField("C:\\doc.txt");
    expect((run.content[0] as any).instruction).toContain("INCLUDETEXT");
  });

  it("includePictureField", () => {
    const run = Build.includePictureField("logo.png");
    expect((run.content[0] as any).instruction).toContain("INCLUDEPICTURE");
  });

  it("formTextField", () => {
    const run = Build.formTextField({ name: "Name", maxLength: 50 });
    expect((run.content[0] as any).instruction).toContain("FORMTEXT");
    expect((run.content[0] as any).formField.type).toBe("text");
    expect((run.content[0] as any).formField.name).toBe("Name");
  });

  it("formCheckboxField", () => {
    const run = Build.formCheckboxField({ checked: true });
    expect((run.content[0] as any).formField.type).toBe("checkBox");
    expect((run.content[0] as any).formField.checked).toBe(true);
  });

  it("formDropdownField", () => {
    const run = Build.formDropdownField({ entries: ["A", "B"] });
    expect((run.content[0] as any).formField.type).toBe("dropDown");
    expect((run.content[0] as any).formField.entries).toEqual(["A", "B"]);
  });
});

// =============================================================================
// Comment helpers
// =============================================================================

describe("Comment helpers", () => {
  it("commentRangeStart", () => {
    const marker = Build.commentRangeStart(1);
    expect(marker).toEqual({ type: "commentRangeStart", id: 1 });
  });

  it("commentRangeEnd", () => {
    const marker = Build.commentRangeEnd(1);
    expect(marker).toEqual({ type: "commentRangeEnd", id: 1 });
  });

  it("commentReference", () => {
    const ref = Build.commentReference(1);
    expect(ref).toEqual({ type: "commentReference", id: 1 });
  });
});

// =============================================================================
// Track changes helpers
// =============================================================================

describe("Track changes helpers", () => {
  const revision = { id: 1, author: "user", date: "2024-01-01T00:00:00Z" };
  const run = Build.text("moved");

  it("movedFromRun", () => {
    const result = Build.movedFromRun(run, revision);
    expect(result.type).toBe("movedFromRun");
    expect(result.revision).toEqual(revision);
    expect(result.run).toBe(run);
  });

  it("movedToRun", () => {
    const result = Build.movedToRun(run, revision);
    expect(result.type).toBe("movedToRun");
    expect(result.revision).toEqual(revision);
  });

  it("moveFromRangeStart", () => {
    const marker = Build.moveFromRangeStart(1, "author1");
    expect(marker.type).toBe("moveFromRangeStart");
    expect(marker.id).toBe(1);
    expect(marker.author).toBe("author1");
  });

  it("moveFromRangeEnd", () => {
    const marker = Build.moveFromRangeEnd(1);
    expect(marker.type).toBe("moveFromRangeEnd");
    expect(marker.id).toBe(1);
  });

  it("moveToRangeStart", () => {
    const marker = Build.moveToRangeStart(2, "author2");
    expect(marker.type).toBe("moveToRangeStart");
    expect(marker.id).toBe(2);
    expect(marker.author).toBe("author2");
  });

  it("moveToRangeEnd", () => {
    const marker = Build.moveToRangeEnd(2);
    expect(marker.type).toBe("moveToRangeEnd");
    expect(marker.id).toBe(2);
  });
});

// =============================================================================
// Math helpers
// =============================================================================

describe("Math helpers", () => {
  const mc = [{ type: "mathRun", text: "x" }] as any as MathContent[];

  it("mathBlock", () => {
    const block = Build.mathBlock(mc);
    expect(block.type).toBe("math");
    expect(block.content).toBe(mc);
  });

  it("mathFraction", () => {
    const frac = Build.mathFraction(mc, mc);
    expect((frac as any).type).toBe("mathFraction");
    expect((frac as any).numerator).toBe(mc);
    expect((frac as any).denominator).toBe(mc);
  });

  it("mathSqrt", () => {
    const sqrt = Build.mathSqrt(mc);
    expect((sqrt as any).type).toBe("mathRadical");
    expect((sqrt as any).hideDegree).toBe(true);
  });

  it("mathRoot", () => {
    const root = Build.mathRoot(mc, mc);
    expect((root as any).type).toBe("mathRadical");
    expect((root as any).degree).toBe(mc);
  });

  it("mathSum", () => {
    const sum = Build.mathSum(mc, mc, mc);
    expect((sum as any).type).toBe("mathNary");
    expect((sum as any).char).toBe("\u2211");
    expect((sum as any).content).toBe(mc);
  });

  it("mathIntegral", () => {
    const integral = Build.mathIntegral(mc, mc, mc);
    expect((integral as any).type).toBe("mathNary");
    expect((integral as any).char).toBe("\u222B");
  });

  it("mathProduct", () => {
    const product = Build.mathProduct(mc, mc, mc);
    expect((product as any).type).toBe("mathNary");
    expect((product as any).char).toBe("\u220F");
  });

  it("mathSuperScript", () => {
    const sup = Build.mathSuperScript(mc, mc);
    expect((sup as any).type).toBe("mathSuperScript");
  });

  it("mathSubScript", () => {
    const sub = Build.mathSubScript(mc, mc);
    expect((sub as any).type).toBe("mathSubScript");
  });

  it("mathSubSuperScript", () => {
    const ss = Build.mathSubSuperScript(mc, mc, mc);
    expect((ss as any).type).toBe("mathSubSuperScript");
  });

  it("mathDelimiter", () => {
    const delim = Build.mathDelimiter([mc, mc]);
    expect((delim as any).type).toBe("mathDelimiter");
    expect((delim as any).content).toHaveLength(2);
  });

  it("mathNary", () => {
    const nary = Build.mathNary("\u222E", mc, mc, mc);
    expect((nary as any).type).toBe("mathNary");
    expect((nary as any).char).toBe("\u222E");
  });

  it("mathFunction", () => {
    const fn = Build.mathFunction(mc, mc);
    expect((fn as any).type).toBe("mathFunction");
    expect((fn as any).name).toBe(mc);
    expect((fn as any).content).toBe(mc);
  });

  it("mathLimit", () => {
    const lim = Build.mathLimit(mc, mc, "lower");
    expect((lim as any).type).toBe("mathLimit");
    expect((lim as any).limitType).toBe("lower");
  });

  it("mathMatrix", () => {
    const mat = Build.mathMatrix([
      [mc, mc],
      [mc, mc]
    ]);
    expect((mat as any).type).toBe("mathMatrix");
    expect((mat as any).rows).toHaveLength(2);
  });

  it("mathAccent", () => {
    const accent = Build.mathAccent(mc, "\u0302");
    expect((accent as any).type).toBe("mathAccent");
    expect((accent as any).char).toBe("\u0302");
  });

  it("mathBar", () => {
    const bar = Build.mathBar(mc, "top");
    expect((bar as any).type).toBe("mathBar");
    expect((bar as any).position).toBe("top");
  });

  it("mathBox", () => {
    const box = Build.mathBox(mc);
    expect((box as any).type).toBe("mathBox");
    expect((box as any).content).toBe(mc);
  });

  it("mathEquationArray", () => {
    const eq = Build.mathEquationArray([mc, mc]);
    expect((eq as any).type).toBe("mathEquationArray");
    expect((eq as any).rows).toHaveLength(2);
  });
});

// =============================================================================
// structuredDocumentTag
// =============================================================================

describe("structuredDocumentTag", () => {
  it("creates SDT with content and properties", () => {
    const p = Build.textParagraph("Hello") as Paragraph;
    const sdt = Build.structuredDocumentTag([p], { alias: "test" });
    expect(sdt.type).toBe("sdt");
    expect(sdt.content).toHaveLength(1);
    expect(sdt.properties?.alias).toBe("test");
  });

  it("creates SDT with default empty properties", () => {
    const sdt = Build.structuredDocumentTag([]);
    expect(sdt.properties).toEqual({});
  });
});

// =============================================================================
// resolveThemeColor
// =============================================================================

describe("resolveThemeColor", () => {
  it("returns hex string directly if color is a string", () => {
    expect(Theme.resolveColor("FF0000")).toBe("FF0000");
  });

  it("returns undefined for undefined", () => {
    expect(Theme.resolveColor(undefined)).toBeUndefined();
  });

  it("returns val from ColorSpec", () => {
    expect(Theme.resolveColor({ val: "00FF00" })).toBe("00FF00");
  });
});

// =============================================================================
// listSections
// =============================================================================

describe("listSections", () => {
  it("returns sections from document", () => {
    const doc: DocxDocument = {
      body: [Build.textParagraph("para1"), Build.textParagraph("para2")],
      sectionProperties: { pageSize: { width: 12240, height: 15840 } }
    } as any;
    const sections = Query.listSections(doc);
    expect(sections.length).toBeGreaterThanOrEqual(1);
  });
});

// =============================================================================
// toBase64 and fillTemplateFromBuffer
// =============================================================================

describe("toBase64", () => {
  it("produces a base64 string from a document", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, "Hello");
    const b64 = await Io.toBase64(Document.build(doc));
    expect(typeof b64).toBe("string");
    expect(b64.length).toBeGreaterThan(0);
  });
});

describe("fillTemplateFromBuffer", () => {
  it("fills a template buffer with data", async () => {
    const doc = Document.create();
    Document.addParagraphElement(doc, Build.paragraph([Build.text("{{name}}")]));
    const buffer = await Io.package(Document.build(doc));
    const result = await Io.fillTemplateFromBuffer(buffer, { name: "Alice" });
    expect(result).toBeInstanceOf(Uint8Array);
    expect(result.length).toBeGreaterThan(0);
  });
});

// =============================================================================
// Error classes
// =============================================================================

describe("Error classes", () => {
  it("DocxWriteError is an instance of Error", () => {
    const err = new DocxWriteError("write failed");
    expect(err).toBeInstanceOf(Error);
    expect(err.message).toBe("write failed");
    expect(err.name).toBe("DocxWriteError");
  });

  it("DocxInvalidStructureError is an instance of Error", () => {
    const err = new DocxInvalidStructureError("bad structure");
    expect(err).toBeInstanceOf(Error);
    expect(err.name).toBe("DocxInvalidStructureError");
  });

  it("DocxUnsupportedFeatureError includes feature name", () => {
    const err = new DocxUnsupportedFeatureError("VML");
    expect(err.message).toContain("VML");
    expect(err.name).toBe("DocxUnsupportedFeatureError");
  });
});

// =============================================================================
// addBulletList / addNumberedList numbering-instance robustness
//
// Previously the code did `find(...)!.numId` after a separate code path
// inserted the abstract definition. If the abstract was already present
// (e.g. a caller added it manually, or a future code path forgot to push
// the matching instance) the `!` would crash with a TypeError. The helper
// now lazily registers an instance whenever one is missing.
// =============================================================================

describe("addBulletList / addNumberedList: numbering-instance resilience", () => {
  it("creates a numbering instance when an abstract exists without one", () => {
    const h = Document.create();
    // Inject a bullet abstract directly into the handle's state via
    // Document.addStyle is not appropriate here — but `addStyle` is the
    // wrong API. Instead we drive addBulletList twice: the first call
    // registers both abstract and instance; we then strip the instance
    // and call addBulletList again to verify it heals itself.
    Document.addBulletList(h, ["A"]);
    const built1 = Document.build(h);
    expect(built1.numberingInstances?.length).toBe(1);

    // Surgically remove the instance to simulate a drift scenario (this
    // path triggered the bug).
    type State = {
      numberingInstances: Array<{ numId: number; abstractNumId: number }>;
    };
    const state = h as unknown as State;
    state.numberingInstances = [];

    // Must not throw — and must re-register an instance for the existing
    // bullet abstract.
    expect(() => Document.addBulletList(h, ["B"])).not.toThrow();
    const built2 = Document.build(h);
    expect(built2.numberingInstances?.length).toBe(1);
    expect(built2.abstractNumberings?.length).toBe(1);
  });

  it("reuses the same numbering instance across repeated calls", () => {
    const h = Document.create();
    Document.addBulletList(h, ["A", "B"]);
    Document.addBulletList(h, ["C", "D"]);
    const built = Document.build(h);
    // One abstract, one instance — repeated bullet calls share both.
    expect(built.abstractNumberings?.length).toBe(1);
    expect(built.numberingInstances?.length).toBe(1);
  });

  it("creates separate abstracts and instances for bullets vs numbered", () => {
    const h = Document.create();
    Document.addBulletList(h, ["A"]);
    Document.addNumberedList(h, ["1"]);
    const built = Document.build(h);
    expect(built.abstractNumberings?.length).toBe(2);
    expect(built.numberingInstances?.length).toBe(2);
  });
});
