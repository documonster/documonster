/**
 * Deprecated names kept until 1.0, and the new names that replace
 * them. Pure aliases are checked by identity; wrappers by behaviour.
 */
import { describe, it, expect } from "vitest";

import { WordCrypto } from "../crypto";
import * as CryptoEntry from "../crypto";
import { WordExcel } from "../excel";
import * as ExcelEntry from "../excel";
import { WordHtml } from "../html";
import * as HtmlEntry from "../html";
import {
  Build,
  Convert,
  Diff,
  Document,
  Font,
  Glossary,
  Io,
  Layout,
  Package,
  Query,
  RenderContext,
  Security,
  Streaming,
  Template,
  TemplateError,
  Transform,
  Validation
} from "../index";
import type { DocxDocument } from "../index";
import { WordMarkdown } from "../markdown";
import * as MarkdownEntry from "../markdown";

function sample(): DocxDocument {
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  // Fixed timestamps: two writes compared byte-for-byte must not straddle a second
  const at = new Date("2024-01-01T00:00:00Z");
  Document.setCoreProperties(doc, { created: at, modified: at });
  Document.addParagraph(doc, "Hello {{name}}");
  Document.addTable(doc, [["a", "b"]]);
  return doc;
}

describe("Document", () => {
  it("keeps deprecated aliases", () => {
    expect(Document.insertContentAt).toBe(Document.insertContent);
  });

  it("Security.encrypt forwards to WordCrypto.encrypt (it loads the encryption on call, so it is not the same object)", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, "x");
    const zip = await Io.toBuffer(doc);
    const encrypted = await Security.encrypt(zip, "pw", { spinCount: 10 });
    expect(WordCrypto.isEncrypted(encrypted)).toBe(true);
    expect(await WordCrypto.decrypt(encrypted, "pw")).toEqual(zip);
    // Its argument checks are WordCrypto.encrypt's.
    await expect(Security.encrypt(new Uint8Array([1, 2, 3]), "pw")).rejects.toThrow(
      /WordCrypto\.encrypt/
    );
  });

  it("addParagraphElement / addTableElement append the node they are given, as before", () => {
    const doc = Document.create();
    const para = Build.textParagraph("node");
    const table = Build.simpleTable([["x"]]);
    Document.addParagraphElement(doc, para);
    Document.addTableElement(doc, table);
    expect(doc.body[0]).toBe(para);
    expect(doc.body[1]).toBe(table);
  });

  // Observed against the last release (c3b2407c), where `addParagraph` took text only and
  // `addTable` rows only: only a node is appended as it is, anything else is still text or rows.
  it("addParagraph / addTable read a non-node value as text and rows, and refuse what is neither", () => {
    // These used to be put into a run as they were (`text: null`), which no writer can emit; they
    // are the text form's argument, so they are refused as text.
    for (const value of [null, undefined, 5, true, ["x"]]) {
      const doc = Document.create();
      expect(() => Document.addParagraph(doc, value as unknown as string)).toThrow(
        expect.objectContaining({ name: "DocxArgumentError", field: "content" })
      );
      expect(doc.body).toEqual([]);
    }
    for (const value of [null, undefined, 5, true]) {
      expect(() => Document.addTable(Document.create(), value as unknown as string[][])).toThrow(
        expect.objectContaining({ name: "DocxArgumentError", field: "data" })
      );
    }
  });

  it("addParagraph and addTable accept ready-made nodes", () => {
    const doc = Document.create();
    Document.addParagraph(doc, Build.textParagraph("node"));
    Document.addTable(doc, Build.simpleTable([["x"]]));
    expect(doc.body.map(b => b.type)).toEqual(["paragraph", "table"]);
  });

  it("build returns a writable document", () => {
    const doc = sample();
    expect(Document.build(doc).body).toBe(doc.body);
  });

  it("build returns a new top-level object, as it always did", () => {
    const doc = sample();
    const built = Document.build(doc);
    expect(built).not.toBe(doc);
    Document.setSettings(doc, { evenAndOddHeaders: true });
    expect(built.settings).toBeUndefined();
  });

  it("setHeader/setFooter still accept a plain string type", () => {
    const doc = sample();
    const type: string = "default";
    Document.setHeader(doc, type, { children: [] });
    Document.setFooter(doc, type, { children: [] });
    expect(doc.headers?.has("default")).toBe(true);
  });
});

describe("Template", () => {
  it("keeps deprecated aliases", () => {
    expect(Template.fillTemplate).toBeTypeOf("function");
    expect(Template.listTemplateTags).toBeTypeOf("function");
    expect(Template.isTemplateChart).toBe(Template.isChart);
    expect(Template.TemplateError).toBe(TemplateError);
    expect(Io.patchDocument).toBe(Package.patch);
    expect(Io.compileTemplate).toBe(Package.compileTemplate);
    expect(Io.patchTemplate).toBe(Package.patchTemplate);
  });

  it("old fill-from-source/buffer names keep ignoring richPlaceholders", async () => {
    const source = Template.createJsonDataSource({ name: "Ada" });
    const opts = { richPlaceholders: true } as const;
    // The option is not part of the old names' signatures; a caller that
    // passes it anyway (from JavaScript, say) gets the old behaviour.
    // @ts-expect-error -- richPlaceholders is not an option of the deprecated name
    const fromSource = Template.fillTemplateFromSource(sample(), source, opts);
    expect(Query.getText(fromSource)).toBe(
      Query.getText(Template.fillTemplate(sample(), { name: "Ada" }))
    );
    const doc = sample();
    // @ts-expect-error -- as above
    expect(Template.fillTemplateFromSource(doc, source, opts)).toBe(doc);
    expect(Template.fillFromSource(doc, source, opts)).not.toBe(doc);
    const bytes = await Io.toBuffer(sample());
    // @ts-expect-error -- as above
    const oldOut = await Io.read(await Io.fillTemplateFromBuffer(bytes, { name: "Ada" }, opts));
    const newOut = await Io.read(await Package.fillTemplate(bytes, { name: "Ada" }, opts));
    expect(Query.getText(oldOut)).toBe(Query.getText(newOut));
  });

  it("does not advertise richPlaceholders on the deprecated fill names", () => {
    // @ts-expect-error -- fillTemplate never resolved rich placeholders
    Template.fillTemplate(sample(), { name: "Ada" }, { richPlaceholders: true });
    // @ts-expect-error -- fillTemplateEnhanced always does
    Template.fillTemplateEnhanced(sample(), { name: "Ada" }, { richPlaceholders: false });
  });

  it("fill matches fillTemplate, and richPlaceholders matches fillTemplateEnhanced", () => {
    const a = Template.fill(sample(), { name: "Ada" });
    const b = Template.fillTemplate(sample(), { name: "Ada" });
    expect(Query.getText(a)).toBe(Query.getText(b));
    const c = Template.fill(sample(), { name: "Ada" }, { richPlaceholders: true });
    const d = Template.fillTemplateEnhanced(sample(), { name: "Ada" });
    expect(Query.getText(c)).toBe(Query.getText(d));
    expect(Query.getText(c)).toContain("Hello Ada");
  });

  it("fill leaves its input untouched, with or without richPlaceholders", () => {
    for (const options of [undefined, { richPlaceholders: true }]) {
      const doc = sample();
      Document.setHeader(doc, "default", { children: [Build.textParagraph("Head {{name}}")] });
      Document.addFootnote(doc, "Note {{name}}");
      const snapshot = structuredClone(doc);
      const filled = Template.fill(doc, { name: "Ada" }, options);
      expect(doc).toEqual(snapshot);
      expect(filled).not.toBe(doc);
      expect(Query.getText(filled)).toContain("Hello Ada");
    }
  });

  it("fillTemplate keeps filling in place", () => {
    const doc = sample();
    expect(Template.fillTemplate(doc, { name: "Ada" })).toBe(doc);
    expect(Query.getText(doc)).toContain("Hello Ada");
  });
});

describe("Io / Package / Transform", () => {
  it("keeps deprecated aliases", () => {
    expect((Io as Record<string, unknown>).package).toBe(Io.toBuffer);
    expect(Io.editDocxIncremental).toBe(Package.edit);
    expect(Io.listDocxParts).toBe(Package.listParts);
    expect(Io.readDocxPart).toBe(Package.readPart);
    expect(Io.merge).toBe(Transform.merge);
    expect(Io.split).toBe(Transform.split);
    expect(Io.updateFields).toBe(Transform.updateFields);
    expect(Io.updateTableOfContents).toBe(Transform.updateTableOfContents);
    expect(Convert.mapDocument).toBe(Transform.map);
  });

  it("the deprecated Io.package writes what Io.toBuffer writes", async () => {
    const doc = sample();
    Document.setCoreProperties(doc, { created: new Date(0), modified: new Date(0) });
    const packageFn = (Io as unknown as Record<string, typeof Io.toBuffer>).package;
    const options = { compressionLevel: 0 };
    expect(await packageFn(doc, options)).toEqual(await Io.toBuffer(doc, options));
  });

  it("toFlatOpc and the deprecated positional form agree", async () => {
    const doc = sample();
    expect(await Io.toFlatOpcFromDoc(doc, 0)).toBe(
      await Io.toFlatOpc(doc, { compressionLevel: 0 })
    );
  });
});

describe("Convert", () => {
  it("keeps deprecated aliases", () => {
    expect(Convert.parseFlatOpc).toBe(Convert.flatOpcToParts);
    expect(Convert.readOdt).toBe(Convert.fromOdt);
    expect(Convert.writeOdt).toBe(Convert.toOdt);
    expect(Convert.docxToSemantic).toBe(Convert.toSemantic);
    expect(Convert.ommlToMathML).toBe(Convert.ommlToMathMl);
    expect(Convert.mathMLToOmml).toBe(Convert.mathMlToOmml);
    expect(Convert.createConversionContext).toBeTypeOf("function");
    expect(Convert.createConversionContext().warnings).toBeDefined();
  });
});

describe("Layout", () => {
  it("keeps deprecated aliases", () => {
    expect(Layout.document).toBe(Layout.getPageMap);
    expect(Layout.documentFull).toBe(Layout.compute);
    expect(Layout.renderDocumentToSvg).toBe(Layout.toSvg);
    expect(Layout.renderPageToSvg).toBe(Layout.pageToSvg);
    expect(Layout.renderPageFromLayout).toBe(Layout.computedPageToSvg);
  });

  it("renders one page from a document, or from its precomputed layout", () => {
    const doc = sample();
    const fromDoc = Layout.pageToSvg(doc, 1);
    expect(Layout.computedPageToSvg(Layout.compute(doc), 1)).toBe(fromDoc);
  });
});

describe("Build / Streaming / Glossary / Diff / Validation / Security", () => {
  it("keeps deprecated aliases", () => {
    expect(Build.createShape).toBe(Build.shape);
    expect(Build.createRect).toBe(Build.rect);
    expect(Build.createRoundRect).toBe(Build.roundRect);
    expect(Build.createEllipse).toBe(Build.ellipse);
    expect(Build.createLine).toBe(Build.line);
    expect(Build.createArrow).toBe(Build.arrow);
    expect(Build.createFlowchartShape).toBe(Build.flowchartShape);
    expect(Build.createCallout).toBe(Build.callout);
    expect(Build.createStar).toBe(Build.star);
    expect(Build.checkBox).toBe(Build.checkbox);
    expect(Build.mathSuperScript).toBe(Build.mathSuperscript);
    expect(Build.mathSubScript).toBe(Build.mathSubscript);
    expect(Build.mathSubSuperScript).toBe(Build.mathSubSuperscript);
    expect(Build.mathPreSubSuperScript).toBe(Build.mathPreSubSuperscript);
    expect(Document.nextBookmarkId).toBe(Document.allocateBookmarkId);
    expect(Convert.toFlatOpc).toBe(Convert.partsToFlatOpc);
    expect(Streaming.createDocxStream).toBe(Streaming.createWriter);
    expect(Streaming.createDocxStreamReader).toBe(Streaming.createReader);
    expect(Glossary.autoTextEntries).toBe(Glossary.listAutoTextEntries);
    expect(Glossary.quickParts).toBe(Glossary.listQuickParts);
    expect(Diff.documents).toBe(Diff.compare);
    expect(Validation.document).toBe(Validation.validate);
    // `Security.verifyPassword` is the protection-password check again; the
    // unpublished `verifyProtectionPassword` it briefly forwarded to is gone.
    expect(Security.verifyPassword).toBeTypeOf("function");
    expect("verifyProtectionPassword" in Security).toBe(false);
    expect(Security.resolveSecurityPolicy).toBe(Security.resolvePolicy);
    expect(RenderContext.create).toBeTypeOf("function");
    expect(RenderContext.createIds).toBeTypeOf("function");
  });

  // Expected values are what the positional forms returned at the last release (c3b2407c).
  it("border / gridBorders options match the positional forms", () => {
    const b = (style: unknown, size: number, color: string) => ({ style, size, space: 0, color });
    expect(Build.border()).toEqual(b("single", 4, "auto"));
    expect(Build.border({})).toEqual(b("single", 4, "auto"));
    expect(Build.border("nil")).toEqual(b("nil", 4, "auto"));
    expect(Build.border({ style: "nil" })).toEqual(b("nil", 4, "auto"));
    expect(Build.border("dashed", 8, "0070C0")).toEqual(b("dashed", 8, "0070C0"));
    expect(Build.border({ style: "dashed", size: 8, color: "0070C0" })).toEqual(
      b("dashed", 8, "0070C0")
    );
    expect(Build.border(undefined, 12)).toEqual(b("single", 12, "auto"));
    expect(Build.border({ size: 12 })).toEqual(b("single", 12, "auto"));
    // `null` keeps its positional meaning: it is not a default placeholder for a parameter default,
    // so it is the style — which it cannot be, and is refused rather than written as `w:val`.
    expect(() => Build.border(null as unknown as undefined)).toThrow(
      expect.objectContaining({ name: "DocxArgumentError", field: "style" })
    );

    const grid = (size: number, color: string) => {
      const side = b("single", size, color);
      return { top: side, left: side, bottom: side, right: side, insideH: side, insideV: side };
    };
    expect(Build.gridBorders()).toEqual(grid(4, "auto"));
    expect(Build.gridBorders(2, "808080")).toEqual(grid(2, "808080"));
    expect(Build.gridBorders({ size: 2, color: "808080" })).toEqual(grid(2, "808080"));
    expect(Build.gridBorders(undefined, "FF0000")).toEqual(grid(4, "FF0000"));
    expect(Build.gridBorders({ color: "FF0000" })).toEqual(grid(4, "FF0000"));
    const sides = Build.gridBorders({ size: 6 });
    expect(sides.insideV).toBe(sides.top);
  });

  it("dateField / timeField options match the positional forms", () => {
    const f = (instruction: string, cachedValue?: string) => ({
      content: [{ type: "field", instruction, cachedValue }]
    });
    expect(Build.dateField()).toEqual(f(' DATE \\@ "yyyy-MM-dd" '));
    expect(Build.dateField({})).toEqual(f(' DATE \\@ "yyyy-MM-dd" '));
    expect(Build.dateField("d MMMM yyyy")).toEqual(f(' DATE \\@ "d MMMM yyyy" '));
    expect(Build.dateField({ format: "d MMMM yyyy" })).toEqual(f(' DATE \\@ "d MMMM yyyy" '));
    expect(Build.dateField("yyyy", "2024")).toEqual(f(' DATE \\@ "yyyy" ', "2024"));
    expect(Build.dateField({ format: "yyyy", cachedValue: "2024" })).toEqual(
      f(' DATE \\@ "yyyy" ', "2024")
    );
    expect(Build.dateField(undefined, "x")).toEqual(f(' DATE \\@ "yyyy-MM-dd" ', "x"));
    expect(Build.dateField({ cachedValue: "x" })).toEqual(f(' DATE \\@ "yyyy-MM-dd" ', "x"));
    expect(Build.timeField()).toEqual(f(' TIME \\@ "HH:mm:ss" '));
    expect(Build.timeField("HH:mm", "09:30")).toEqual(f(' TIME \\@ "HH:mm" ', "09:30"));
    expect(Build.timeField({ format: "HH:mm", cachedValue: "09:30" })).toEqual(
      f(' TIME \\@ "HH:mm" ', "09:30")
    );
  });

  it("heading defaults to level 1, as Document.addHeading does", () => {
    expect(Build.heading("T")).toEqual(Build.heading("T", 1));
    expect(Build.heading("T").properties?.style).toBe("Heading1");
    const doc = Document.create();
    Document.addHeading(doc, "T");
    expect(doc.body[0]).toEqual(Build.heading("T"));
  });
});

describe("Query", () => {
  it("keeps deprecated aliases", () => {
    expect(Query.paragraphCount).toBe(Query.getParagraphCount);
    expect(Query.tableCount).toBe(Query.getTableCount);
    expect(Query.countWords).toBe(Query.getWordCount);
    expect(Query.countByFormat).toBe(Query.getCountByFormat);
    expect(Query.indexStyles).toBe(Query.createStyleIndex);
    expect(Query.walkBlocks).toBe(Query.eachBlock);
    expect(Query.extractText).toBe(Query.getText);
    expect(Query.findBookmark).toBe(Query.getBookmark);
    expect(Query.findComment).toBe(Query.getComment);
    expect(Query.getHeadings).toBe(Query.listHeadings);
    expect(Query.extractFormFields).toBe(Query.listFormFields);
    expect(Query.getUsedFormats).toBe(Query.listUsedFormats);
    expect(Query.walkDocument).toBe(Query.each);
    expect(Query.collectParagraphs).toBe(Query.listParagraphs);
    expect(Query.collectRuns).toBe(Query.listRuns);
    expect(Query.resolveNumberingLevel).toBe(Query.getNumberingLevel);
  });

  it("keeps forwards to the in-place edits now in Document and the copies now in Transform", () => {
    expect(Query.replaceText).toBe(Document.replaceText);
    expect(Query.mailMerge).toBe(Document.mailMerge);
    expect(Query.acceptAllRevisions).toBe(Document.acceptAllRevisions);
    expect(Query.rejectAllRevisions).toBe(Document.rejectAllRevisions);
    expect(Query.acceptRevision).toBe(Document.acceptRevision);
    expect(Query.rejectRevision).toBe(Document.rejectRevision);
    expect(Query.setCompatibilityMode).toBe(Document.setCompatibilityMode);
    expect(Query.fillFormFields).toBe(Transform.fillFormFields);
    expect(Query.resolveDataBindings).toBe(Transform.resolveDataBindings);
  });

  it("style getters match the positional resolvers", () => {
    const doc = sample();
    const para = doc.body[0];
    if (para.type !== "paragraph") {
      throw new Error("expected a paragraph");
    }
    const styles = Query.createStyleIndex(doc);
    expect(Query.getParagraphStyle(doc, para, { styles })).toEqual(Query.resolveStyle(doc, para));
    const run = { content: [], properties: { bold: true } };
    expect(Query.getRunStyle(doc, run, { paragraphRunProperties: { size: 30 } })).toEqual(
      Query.resolveRunStyle(doc, run, { size: 30 })
    );
    expect(Query.getTableStyle(doc, "TableGrid")).toEqual(
      Query.resolveTableStyle(doc, "TableGrid")
    );
  });

  it("listTables can include other parts, like collectTables", () => {
    const doc = sample();
    Document.setHeader(doc, "default", { children: [Build.simpleTable([["h"]])] });
    expect(Query.listTables(doc)).toHaveLength(1);
    expect(
      Query.listTables(doc, {
        includeHeaders: true,
        includeFooters: true,
        includeFootnotes: true,
        includeEndnotes: true
      })
    ).toHaveLength(Query.collectTables(doc).length);
  });
});

describe("subpath entries", () => {
  it("Html / Markdown / Excel keep deprecated flat names", () => {
    expect(HtmlEntry.renderToHtml).toBe(WordHtml.fromDocx);
    // The flat names keep their old results (body and images; body alone);
    // the namespace members now return a document and a full fragment.
    expect(Object.keys(HtmlEntry.htmlToDocx("<p>x</p>")).sort()).toEqual(["body", "images"]);
    expect(Array.isArray(HtmlEntry.htmlToDocxBody("<p>x</p>"))).toBe(true);
    expect(MarkdownEntry.renderToMarkdown).toBe(WordMarkdown.fromDocx);
    expect(MarkdownEntry.markdownToDocx).toBe(WordMarkdown.toDocx);
    expect(MarkdownEntry.markdownToDocxBody).toBe(WordMarkdown.toDocxBody);
    expect(ExcelEntry.excelToDocx).toBe(WordExcel.toDocx);
    expect(ExcelEntry.extractTablesToExcel).toBe(WordExcel.extractTables);
    expect(ExcelEntry.renderWordChartSvg).toBe(WordExcel.toChartSvg);
    expect(ExcelEntry.buildWordChartExXml).toBe(WordExcel.toChartExXml);
    expect(ExcelEntry.wordChartToChartModel).toBe(WordExcel.toChartModel);
    expect(ExcelEntry.generateChartEmbeddedXlsx).toBe(WordExcel.toChartXlsx);
  });

  it("Crypto keeps deprecated flat names", () => {
    expect(CryptoEntry.isEncryptedDocx).toBe(WordCrypto.isEncrypted);
    expect(CryptoEntry.encryptDocx).toBe(WordCrypto.encrypt);
    expect(CryptoEntry.isWellFormedSignature).toBe(WordCrypto.isWellFormedSignature);
    expect(CryptoEntry.parseSignatureXml).toBe(WordCrypto.parseSignature);
    expect(CryptoEntry.parseEncryptionInfoXml).toBe(WordCrypto.Agile.parseEncryptionInfo);
    expect(CryptoEntry.AGILE_BLOCK_KEYS).toBe(WordCrypto.Agile.BLOCK_KEYS);
    expect(CryptoEntry.writeCfb).toBe(WordCrypto.Cfb.write);
    expect(CryptoEntry.obfuscateFont).toBe(Font.obfuscate);
    expect(CryptoEntry.deobfuscateFont).toBe(Font.deobfuscate);
    expect(CryptoEntry.generateFontKey).toBe(Font.generateObfuscationKey);
  });

  it("WordCrypto.verifyPassword checks an encrypted file without decrypting it", async () => {
    const plain = await Io.toBuffer(sample());
    const encrypted = await WordCrypto.encrypt(plain, "pw", { spinCount: 10 });
    expect(await WordCrypto.verifyPassword(encrypted, "pw")).toBe(true);
    expect(await WordCrypto.verifyPassword(encrypted, "nope")).toBe(false);
    await expect(WordCrypto.verifyPassword(plain, "pw")).rejects.toThrow();
  });

  it("WordCrypto low-level primitives compose into a decrypt", async () => {
    const plain = await Io.toBuffer(sample());
    const encrypted = await WordCrypto.encrypt(plain, "pw", { spinCount: 10 });
    const entries = WordCrypto.Cfb.read(encrypted);
    const infoStream = entries.find(e => e.name === "EncryptionInfo")!.data;
    const pkg = entries.find(e => e.name === "EncryptedPackage")!.data;
    const info = WordCrypto.Agile.parseEncryptionInfo(
      new TextDecoder().decode(infoStream.slice(8))
    );
    expect(await WordCrypto.Agile.verifyPassword(info, "pw")).toBe(true);
    expect(await WordCrypto.Agile.verifyPassword(info, "nope")).toBe(false);
    // The deprecated flat name keeps its (password, info) order.
    expect(await CryptoEntry.verifyPassword("pw", info)).toBe(true);
    expect(await WordCrypto.Agile.decryptPackage(pkg, info, "pw")).toEqual(plain);
    // The deprecated flat names keep their positional cap and unchecked forms;
    // the namespace members check what they are given.
    expect(await CryptoEntry.decryptPackage(pkg, info, "pw")).toEqual(plain);
    await expect(CryptoEntry.decryptPackage(pkg, info, "pw", 10)).rejects.toThrow(/maximum/);
    expect(CryptoEntry.readCfb(encrypted).map(e => e.name)).toEqual(entries.map(e => e.name));
    const block = WordCrypto.Agile.BLOCK_KEYS.encryptedKey;
    expect(await CryptoEntry.deriveEncryptionKey("pw", info, block)).toEqual(
      await WordCrypto.Agile.deriveKey("pw", info, block)
    );
    const key = Font.generateObfuscationKey();
    const font = new Uint8Array(64).map((_, i) => i);
    expect(Font.deobfuscate(Font.obfuscate(font, key), key)).toEqual(font);
  });

  it("WordHtml.toDocx returns a complete document", async () => {
    const doc = WordHtml.toDocx("<h1>T</h1><ul><li>a</li></ul><ol><li>b</li></ol>");
    expect(doc.styles?.some(st => st.styleId === "Heading1")).toBe(true);
    const ids = new Set(doc.numberingInstances?.map(n => n.numId));
    expect(ids.has(HtmlEntry.HTML_BULLET_LIST_NUM_ID)).toBe(true);
    expect(ids.has(HtmlEntry.HTML_NUMBERED_LIST_NUM_ID)).toBe(true);
    expect(doc.body).toEqual(
      WordHtml.toDocxBody("<h1>T</h1><ul><li>a</li></ul><ol><li>b</li></ol>").body
    );
    const back = await Io.read(await Io.toBuffer(doc));
    expect(Query.getText(back)).toContain("b");
  });

  it("WordCrypto.decrypt round-trips and honours maxDecryptedSize", async () => {
    const plain = await Io.toBuffer(sample());
    const encrypted = await WordCrypto.encrypt(plain, "pw", { spinCount: 10 });
    expect(WordCrypto.isEncrypted(encrypted)).toBe(true);
    expect(await WordCrypto.decrypt(encrypted, "pw")).toEqual(plain);
    await expect(WordCrypto.decrypt(encrypted, "pw", { maxDecryptedSize: 10 })).rejects.toThrow();
  });

  it("Crypto signature helpers read the document", () => {
    const doc = sample();
    expect(WordCrypto.hasSignatures(doc)).toBe(false);
    expect(WordCrypto.listSignatures(doc)).toEqual([]);
  });
});

describe("errors a caller can be handed are importable", () => {
  it("exports DocxRawXmlPolicyError from the entry point", async () => {
    const entry = await import("../index");
    const doc = {
      ...Document.create(),
      body: [{ type: "opaqueDrawing", rawXml: "<w:drawing/>", referencedRIds: [] }]
    } as unknown as DocxDocument;
    const error = await Io.toBuffer(doc, { securityPolicy: { rawXmlPolicy: "reject" } }).catch(
      (e: unknown) => e
    );
    expect(error).toBeInstanceOf((entry as Record<string, unknown>).DocxRawXmlPolicyError as never);
  });
});

describe("WordCrypto.hasSignatures", () => {
  it("answers from a package's part names, without reading the document", async () => {
    const plain = await Io.toBuffer(Document.create());
    const signed = await Package.edit(plain, [
      { type: "replacePartText", path: "_xmlsignatures/sig1.xml", text: "<Signature/>" }
    ]);
    expect(WordCrypto.hasSignatures(await Package.listParts(signed))).toBe(true);
    expect(WordCrypto.hasSignatures(await Package.listParts(plain))).toBe(false);
    // …and from a document, as before.
    expect(WordCrypto.hasSignatures(await Io.read(signed))).toBe(true);
    expect(WordCrypto.hasSignatures(Document.create())).toBe(false);
  });
});
