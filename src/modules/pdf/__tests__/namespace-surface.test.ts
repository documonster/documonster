import * as PdfMod from "@pdf/index";
/**
 * Public namespace-surface contract for `documonster/pdf`.
 *
 * Verifies the `Pdf` namespace shape (writing/reading/building/editing/
 * conversion/signatures) and that a representative member produces output.
 */
import { describe, it, expect } from "vitest";

const { Pdf } = PdfMod;

describe("documonster/pdf namespace surface", () => {
  it("exposes the expected members", () => {
    for (const m of [
      "fromTable",
      "extract",
      "fromExcel",
      "fromDocx",
      "fromChart",
      "verifySignatures",
      "parseSvgPath"
    ]) {
      expect(typeof (Pdf as Record<string, unknown>)[m], `Pdf.${m}`).toBe("function");
    }
    expect(typeof Pdf.Builder).toBe("function"); // class
    expect(typeof Pdf.Editor).toBe("function"); // class
    expect(Pdf.PageSizes).toBeDefined();
  });

  it("builds with Pdf.Builder and opens with Pdf.Editor.read", async () => {
    const builder = new Pdf.Builder();
    builder.addPage();
    const bytes = await builder.toBuffer();
    const editor = Pdf.Editor.read(bytes);
    expect(editor).toBeInstanceOf(Pdf.Editor);
    expect(editor.pageCount).toBe(1);
  });

  it("Pdf.fromTable produces a non-trivial PDF byte stream", async () => {
    const bytes = await Pdf.fromTable([
      ["Product", "Revenue"],
      ["Widget", 1000]
    ]);
    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(bytes.length).toBeGreaterThan(100);
    // PDF magic header "%PDF"
    expect(bytes[0]).toBe(0x25);
    expect(bytes[1]).toBe(0x50);
  });

  it("the deprecated verifySignature keeps its old result shape", async () => {
    const result = await Pdf.verifySignature(new Uint8Array(100), "00", [0, 10, 20, 80]);
    expect(result.valid).toBe(false);
    expect(result.reason).toBeDefined();
    expect("intact" in result).toBe(false);
  });

  it("keeps deprecated names as forwards", () => {
    expect(Pdf.create).toBe(Pdf.fromTable);
    expect(Pdf.read).toBe(Pdf.extract);
    expect(Pdf.buildSignatureDictPlaceholder).toBe(Pdf.Signature.buildDictPlaceholder);
    expect(Pdf.asn1Parse).toBe(Pdf.Signature.parseAsn1);
  });

  it("ships no createWordChartRenderer: the default already draws every chart (R4P-S1)", async () => {
    expect("createWordChartRenderer" in Pdf).toBe(false);
    // The deprecated name keeps returning the classic-chart callback it always did.
    const renderer = await Pdf.wordChartRenderer();
    expect(typeof renderer).toBe("function");
    expect(renderer.length).toBe(3);
  });

  it("the deprecated createPdfDrawSurface forwards a positional scale", () => {
    const page = new Pdf.Builder().addPage({ width: 100, height: 100 });
    const viaOptions = Pdf.createDrawSurface(
      page,
      { x: 0, y: 0, width: 10, height: 10 },
      {
        scale: 2
      }
    );
    const viaLegacy = PdfMod.createPdfDrawSurface(page, { x: 0, y: 0, width: 10, height: 10 }, 2);
    viaOptions.rect(1, 1, 2, 2, 0, { fill: { r: 0, g: 0, b: 0, a: 1 } });
    viaLegacy.rect(1, 1, 2, 2, 0, { fill: { r: 0, g: 0, b: 0, a: 1 } });
    const ops =
      page
        .getContentStream()
        .toString()
        .match(/^.* re$/gm) ?? [];
    expect(ops).toHaveLength(2);
    expect(ops[0]).toBe(ops[1]);
  });

  it("Builder.toBuffer and the deprecated build produce a PDF", async () => {
    const doc = new Pdf.Builder();
    doc.addPage();
    expect((await doc.toBuffer())[0]).toBe(0x25);
    expect((await doc.build())[0]).toBe(0x25);
  });

  it("Builder.setTextLanguage and the deprecated textLanguage both chain", () => {
    const doc = new Pdf.Builder();
    expect(doc.setTextLanguage("ja")).toBe(doc);
    expect(doc.textLanguage("ko")).toBe(doc);
  });

  it("Editor.read / toBuffer and the deprecated load / save / saveIncremental agree", async () => {
    const doc = new Pdf.Builder();
    doc.addPage();
    doc.addPage();
    const bytes = await doc.toBuffer();
    expect(Pdf.Editor.load(bytes).pageCount).toBe(Pdf.Editor.read(bytes).pageCount);
    expect((await Pdf.Editor.read(bytes).toBuffer())[0]).toBe(0x25);
    expect((await Pdf.Editor.read(bytes).save())[0]).toBe(0x25);
    const incremental = await Pdf.Editor.read(bytes).toBuffer({ incremental: true });
    const legacy = await Pdf.Editor.read(bytes).saveIncremental();
    // An incremental update appends to the original bytes rather than rewriting them.
    expect(incremental.subarray(0, bytes.length)).toEqual(bytes);
    expect(legacy.length).toBe(incremental.length);
  });

  it("Editor.copyPagesFrom accepts { pages } and the deprecated positional indices", async () => {
    const src = new Pdf.Builder();
    src.addPage();
    src.addPage();
    src.addPage();
    const source = await src.toBuffer();
    const base = new Pdf.Builder();
    base.addPage();
    const target = await base.toBuffer();

    const a = Pdf.Editor.read(target).copyPagesFrom(source, { pageIndices: [0, 2] });
    const b = Pdf.Editor.read(target).copyPagesFrom(source, [0, 2]);
    expect((await Pdf.extract(await a.toBuffer())).pages).toHaveLength(3);
    expect((await Pdf.extract(await b.toBuffer())).pages).toHaveLength(3);
  });

  it("the deprecated positional copyPagesFrom still passes the source password", async () => {
    const src = new Pdf.Builder();
    src.addPage();
    src.addPage();
    src.setEncryption({ userPassword: "open", ownerPassword: "owner" });
    const source = await src.toBuffer();
    const base = new Pdf.Builder();
    base.addPage();
    const target = await base.toBuffer();

    const legacy = Pdf.Editor.read(target).copyPagesFrom(source, undefined, { password: "open" });
    const current = Pdf.Editor.read(target).copyPagesFrom(source, { password: "open" });
    expect((await Pdf.extract(await legacy.toBuffer())).pages).toHaveLength(3);
    expect((await Pdf.extract(await current.toBuffer())).pages).toHaveLength(3);
    expect(() => Pdf.Editor.read(target).copyPagesFrom(source, [0])).toThrow();
  });

  it("a signature-guard error names the method that was called", async () => {
    const { generateSelfSignedCertificate } =
      await import("../examples/utils/self-signed-certificate");
    const { certificate, privateKey } = await generateSelfSignedCertificate("Guard");
    const doc = new Pdf.Builder();
    doc.addPage();
    const signed = await Pdf.Editor.read(await doc.toBuffer()).sign({ certificate, privateKey });

    await expect(Pdf.Editor.read(signed).save()).rejects.toThrow(/save\(\) rewrites/);
    await expect(Pdf.Editor.read(signed).toBuffer()).rejects.toThrow(/toBuffer\(\) rewrites/);
    const legacy = Pdf.Editor.read(signed);
    legacy.addPage();
    await expect(legacy.saveIncremental()).rejects.toThrow(/saveIncremental\(\) must fall back/);
    const current = Pdf.Editor.read(signed);
    current.addPage();
    await expect(current.toBuffer({ incremental: true })).rejects.toThrow(
      /toBuffer\(\{ incremental: true \}\) must fall back/
    );
  });

  it("the deprecated Pdf.sign keeps both forms and its original error", async () => {
    const noPlaceholder = await Pdf.fromTable([["a"]]);
    const key = new Uint8Array([1]);
    await expect(Pdf.sign(noPlaceholder, { certificate: key, privateKey: key })).rejects.toThrow(
      /ByteRange/
    );
    await expect(Pdf.sign(noPlaceholder, key, key)).rejects.toThrow(/ByteRange/);
  });

  it("error classes stay flat", () => {
    expect(typeof PdfMod.PdfError).toBe("function");
    expect(typeof PdfMod.isPdfError).toBe("function");
  });
});
