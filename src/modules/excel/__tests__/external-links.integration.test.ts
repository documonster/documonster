import { ZipParser } from "@archive/unzip/zip-parser";
/**
 * Integration test for external workbook links ([Book]Sheet!Ref formulas).
 *
 * Exercises the full pipeline:
 *   1. Create a workbook, call addExternalLink(), write a formula that
 *      references the external file.
 *   2. Write the workbook to a buffer.
 *   3. Unzip that buffer and verify every required OOXML part is present
 *      and well-formed:
 *        - xl/externalLinks/externalLink1.xml
 *        - xl/externalLinks/_rels/externalLink1.xml.rels (with the
 *          critical TargetMode="External" on a relative Target)
 *        - <externalReferences> in xl/workbook.xml
 *        - externalLink entry in xl/_rels/workbook.xml.rels
 *        - Override in [Content_Types].xml
 *        - Formula in the sheet rewritten to `[N]` numeric form
 *   4. Reload the buffer into a fresh Workbook and confirm the model is
 *      reconstructed identically.
 *
 * This is the end-to-end regression test for external-link relative path
 * resolution: if any of the
 * above pieces regresses, Office/WPS will not resolve relative paths
 * correctly and the file will either open with a "damaged" warning or
 * silently fall back to %USERPROFILE%\Documents.
 */
import { Cell, Workbook } from "@excel/index";
import type { WorkbookData } from "@excel/workbook-core";
import { describe, expect, it } from "vitest";

import { expectValidXlsx } from "./helpers/expect-valid-xlsx";

async function writeAndUnzip(wb: WorkbookData): Promise<Map<string, string>> {
  const buf = await Workbook.toBuffer(wb);
  await expectValidXlsx(buf, { label: "external-links writeAndUnzip" });
  const parser = new ZipParser(buf);
  const files = await parser.extractAll();
  const decoder = new TextDecoder();
  const out = new Map<string, string>();
  for (const [name, data] of files) {
    // Decode only text entries; binary files (none for our assertions here)
    // can be decoded safely as UTF-8 since the library only emits XML here.
    out.set(name, decoder.decode(data));
  }
  return out;
}

describe("external workbook links — end-to-end", () => {
  it("produces a complete externalLink part for a relative filename target", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Main");

    // The user-facing call for external links: declare the external
    // workbook with just its bare filename — Office will resolve this
    // relative to the current workbook's directory.
    Workbook.addExternalLink(wb, {
      target: "测试.xlsx",
      sheetNames: ["Sheet1"],
      cachedValues: { Sheet1: { A1: 42 } }
    });

    // Formula uses the filename form; the library rewrites it to [1].
    Cell.setValue(ws, "A1", {
      formula: "[测试.xlsx]Sheet1!A1",
      result: 42
    });

    const files = await writeAndUnzip(wb);

    // 1. externalLink part exists with cached value and sheet name
    const extLinkXml = files.get("xl/externalLinks/externalLink1.xml");
    expect(extLinkXml).toBeDefined();
    expect(extLinkXml).toContain('<sheetName val="Sheet1"/>');
    expect(extLinkXml).toContain("<v>42</v>");
    expect(extLinkXml).toContain('<externalBook r:id="rId1">');

    // 2. externalLink rels — target must be relative, TargetMode External
    const extLinkRels = files.get("xl/externalLinks/_rels/externalLink1.xml.rels");
    expect(extLinkRels).toBeDefined();
    expect(extLinkRels).toContain('Target="测试.xlsx"');
    expect(extLinkRels).toContain('TargetMode="External"');
    // Must NOT have promoted the relative path to an absolute file:// URL,
    // which would change Office's lookup behaviour back to the Documents
    // folder.
    expect(extLinkRels).not.toContain("file://");

    // 3. workbook.xml has <externalReferences>
    const workbookXml = files.get("xl/workbook.xml");
    expect(workbookXml).toBeDefined();
    expect(workbookXml).toContain("<externalReferences>");
    expect(workbookXml).toContain("<externalReference");

    // 4. workbook.xml.rels references the externalLink part
    const workbookRels = files.get("xl/_rels/workbook.xml.rels");
    expect(workbookRels).toBeDefined();
    expect(workbookRels).toContain("externalLinks/externalLink1.xml");
    expect(workbookRels).toContain(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
    );

    // 5. [Content_Types].xml has the Override
    const contentTypes = files.get("[Content_Types].xml");
    expect(contentTypes).toBeDefined();
    expect(contentTypes).toContain('PartName="/xl/externalLinks/externalLink1.xml"');
    expect(contentTypes).toContain(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"
    );

    // 6. Sheet XML: formula normalised to numeric form.
    const sheetXml = files.get("xl/worksheets/sheet1.xml");
    expect(sheetXml).toBeDefined();
    // The filename form must have been replaced with [1] on disk — that is
    // the canonical OOXML storage form.
    expect(sheetXml).toContain("[1]Sheet1!A1");
    expect(sheetXml).not.toContain("[测试.xlsx]Sheet1!A1");
  });

  it("auto-creates an externalLink when a formula is written without addExternalLink", async () => {
    // Users writing `{ formula: "[測試.xlsx]Sheet1!A1" }` without first
    // calling addExternalLink() should still get a valid file. The writer
    // must auto-upsert the missing link and rewrite the formula.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Main");

    Cell.setValue(ws, "A1", {
      formula: "[data.xlsx]Summary!B2",
      result: 100
    });

    const files = await writeAndUnzip(wb);
    expect(files.get("xl/externalLinks/externalLink1.xml")).toBeDefined();
    expect(files.get("xl/externalLinks/_rels/externalLink1.xml.rels")).toContain(
      'Target="data.xlsx"'
    );
    expect(files.get("xl/worksheets/sheet1.xml")).toContain("[1]Summary!B2");
  });

  it("reuses the same link when multiple formulas reference the same file", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Main");

    Cell.setValue(ws, "A1", { formula: "[src.xlsx]S1!A1", result: 1 });
    Cell.setValue(ws, "A2", { formula: "[src.xlsx]S1!A2", result: 2 });
    Cell.setValue(ws, "A3", { formula: "[src.xlsx]S2!A1", result: 3 });

    const files = await writeAndUnzip(wb);

    // Only one externalLink file — the three formulas share it.
    expect(files.has("xl/externalLinks/externalLink1.xml")).toBe(true);
    expect(files.has("xl/externalLinks/externalLink2.xml")).toBe(false);

    // Both sheet names recorded on the same link.
    const extLinkXml = files.get("xl/externalLinks/externalLink1.xml")!;
    expect(extLinkXml).toContain('<sheetName val="S1"/>');
    expect(extLinkXml).toContain('<sheetName val="S2"/>');
  });

  it("handles multiple external workbooks with stable indices", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Main");

    Cell.setValue(ws, "A1", { formula: "[a.xlsx]S!A1", result: 1 });
    Cell.setValue(ws, "A2", { formula: "[b.xlsx]S!A1", result: 2 });

    const files = await writeAndUnzip(wb);
    expect(files.has("xl/externalLinks/externalLink1.xml")).toBe(true);
    expect(files.has("xl/externalLinks/externalLink2.xml")).toBe(true);

    const sheet = files.get("xl/worksheets/sheet1.xml")!;
    expect(sheet).toContain("[1]S!A1");
    expect(sheet).toContain("[2]S!A1");
  });

  it("preserves quoted-sheet external refs on round trip", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Main");

    Cell.setValue(ws, "A1", {
      formula: "'[a.xlsx]Has Space'!A1",
      result: 1
    });

    const files = await writeAndUnzip(wb);
    const sheetXml = files.get("xl/worksheets/sheet1.xml")!;
    // Sheet still needs quoting; workbook token became numeric.
    expect(sheetXml).toContain("&apos;[1]Has Space&apos;!A1");
  });

  it("round-trips cached values through read → write → read", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Main");
    Workbook.addExternalLink(wb, {
      target: "src.xlsx",
      sheetNames: ["Sheet1"],
      cachedValues: {
        Sheet1: {
          A1: 123,
          B1: "text",
          C1: true,
          D1: false
        }
      }
    });
    Cell.setValue(ws, "A1", { formula: "[src.xlsx]Sheet1!A1", result: 123 });

    const buf = await Workbook.toBuffer(wb);
    await expectValidXlsx(buf, { label: "cached-values roundtrip" });

    const wb2 = Workbook.create();
    await Workbook.read(wb2, buf);

    expect(wb2.externalLinks).toHaveLength(1);
    const link = wb2.externalLinks[0];
    expect(link.target).toBe("src.xlsx");
    expect(link.targetMode).toBe("External");
    expect(link.sheetNames).toEqual(["Sheet1"]);
    expect(link.cachedValues).toEqual({
      Sheet1: { A1: 123, B1: "text", C1: true, D1: false }
    });
  });

  it("emits no externalLinks artifacts when the workbook has no external refs", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Main");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", { formula: "A1+1", result: 2 });

    const files = await writeAndUnzip(wb);
    for (const name of files.keys()) {
      expect(name).not.toMatch(/^xl\/externalLinks\//);
    }
    const workbookXml = files.get("xl/workbook.xml")!;
    expect(workbookXml).not.toContain("<externalReferences");
    const contentTypes = files.get("[Content_Types].xml")!;
    expect(contentTypes).not.toContain("externalLink");
  });

  it("rejects overflow numeric refs gracefully (they stay in the formula, no link invented)", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Main");
    // Formula mentions [3] but the workbook only has no externalLinks.
    // We leave the formula as-is and emit no spurious link; Excel will
    // surface the error at load time (same as before).
    Cell.setValue(ws, "A1", { formula: "[3]Sheet1!A1+1", result: 0 });

    const files = await writeAndUnzip(wb);
    expect(files.has("xl/externalLinks/externalLink1.xml")).toBe(false);
    const sheetXml = files.get("xl/worksheets/sheet1.xml")!;
    expect(sheetXml).toContain("[3]Sheet1!A1");
  });

  // External link with a non-ASCII (Chinese) filename. We leave the
  // resulting xlsx in tmp/ so it can be opened manually in Office/WPS for a
  // final visual check (the test itself only asserts the on-disk structure).
  it("writes a valid external link with Chinese filename", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "XYZ");
    Cell.setValue(ws, "A1", {
      formula: "[测试.xlsx]Sheet1!A1",
      result: 0
    });

    const buf = await Workbook.toBuffer(wb);
    await expectValidXlsx(buf, { label: "external-link-chinese" });

    const { writeFile, mkdir } = await import("node:fs/promises");
    await mkdir("tmp", { recursive: true });
    await writeFile("tmp/external-link-chinese.xlsx", buf);

    // Reload and verify — the file we just wrote must be a valid xlsx
    // that documonster itself can read back.
    const wb2 = Workbook.create();
    await Workbook.read(wb2, buf);

    expect(wb2.externalLinks).toHaveLength(1);
    expect(wb2.externalLinks[0].target).toBe("测试.xlsx");
    expect(wb2.externalLinks[0].targetMode).toBe("External");
    expect(wb2.externalLinks[0].sheetNames).toEqual(["Sheet1"]);
  });

  it("preserves absolute file:// URLs on round trip", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Main");
    Workbook.addExternalLink(wb, {
      target: "file:///Users/test/data.xlsx",
      sheetNames: ["Sheet1"]
    });
    Cell.setValue(ws, "A1", {
      formula: "[file:///Users/test/data.xlsx]Sheet1!A1",
      result: 0
    });

    const buf = await Workbook.toBuffer(wb);
    await expectValidXlsx(buf, { label: "absolute-file-url" });
    const wb2 = Workbook.create();
    await Workbook.read(wb2, buf);

    expect(wb2.externalLinks[0].target).toBe("file:///Users/test/data.xlsx");
  });

  it("supports multiple sheet references within the same external workbook", async () => {
    // Sanity check for the `upsertSheet` logic: three formulas referencing
    // three sheets of the same external file should produce a single link
    // with all three sheet names in declaration order.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Main");

    Cell.setValue(ws, "A1", { formula: "[ref.xlsx]Alpha!A1", result: 1 });
    Cell.setValue(ws, "A2", { formula: "[ref.xlsx]Beta!A1", result: 2 });
    Cell.setValue(ws, "A3", { formula: "[ref.xlsx]Gamma!A1", result: 3 });
    Cell.setValue(ws, "A4", { formula: "[ref.xlsx]Alpha!B1", result: 4 }); // dup

    const buf = await Workbook.toBuffer(wb);
    await expectValidXlsx(buf, { label: "multi-sheet external ref" });
    const wb2 = Workbook.create();
    await Workbook.read(wb2, buf);

    expect(wb2.externalLinks).toHaveLength(1);
    expect(wb2.externalLinks[0].sheetNames).toEqual(["Alpha", "Beta", "Gamma"]);
  });

  it("is stable under repeated writeBuffer() calls (no duplicate links)", async () => {
    // Auto-discovered external refs are stashed on a private per-workbook
    // cache, not on `wb.externalLinks`. This makes subsequent writes
    // fixed-point stable: the second writeBuffer() sees formulas already
    // in `[N]Sheet!` numeric form and resolves them against the same
    // cached entry, producing identical output.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Main");
    Cell.setValue(ws, "A1", { formula: "[src.xlsx]Sheet1!A1", result: 1 });

    const buf1 = await Workbook.toBuffer(wb);
    await expectValidXlsx(buf1, { label: "repeated-writeBuffer pass 1" });
    const buf2 = await Workbook.toBuffer(wb);
    await expectValidXlsx(buf2, { label: "repeated-writeBuffer pass 2" });

    // Public API stays empty (no user-declared links).
    expect(wb.externalLinks).toHaveLength(0);

    // Both buffers produce identical reload results with the expected link.
    const wb1 = Workbook.create();
    await Workbook.read(wb1, buf1);
    const wbb = Workbook.create();
    await Workbook.read(wbb, buf2);
    expect(wb1.externalLinks).toHaveLength(1);
    expect(wbb.externalLinks).toHaveLength(1);
    expect(wb1.externalLinks[0].target).toBe("src.xlsx");
    expect(wbb.externalLinks[0].target).toBe("src.xlsx");
  });

  it("keeps wb.externalLinks clean when links are auto-discovered from formulas", async () => {
    // A formula references a file the user never declared via
    // addExternalLink(). The writer produces a valid xlsx, but
    // wb.externalLinks remains empty because the auto-discovered entry
    // lives on a private cache — the public list only reflects what the
    // user explicitly declared.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Main");
    Cell.setValue(ws, "A1", { formula: "[src.xlsx]Sheet1!A1", result: 0 });

    expect(wb.externalLinks).toHaveLength(0);
    const buf = await Workbook.toBuffer(wb);
    await expectValidXlsx(buf, { label: "auto-discovered external link" });
    expect(wb.externalLinks).toHaveLength(0);

    // Reading it back, though, surfaces the link — the file is complete.
    const wb2 = Workbook.create();
    await Workbook.read(wb2, buf);
    expect(wb2.externalLinks).toHaveLength(1);
    expect(wb2.externalLinks[0].target).toBe("src.xlsx");
  });

  it("does not mutate user-declared externalLinks entries on write", async () => {
    // When the user declared a link, the writer must leave the user's
    // ExternalLinkModel object intact: no rId assignment, no sheetNames
    // growth, no index renumbering that the user can observe.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Main");

    const link = Workbook.addExternalLink(wb, {
      target: "src.xlsx",
      sheetNames: ["Sheet1"]
    });
    const originalSnapshot = JSON.parse(JSON.stringify(link));

    Cell.setValue(ws, "A1", { formula: "[src.xlsx]Sheet1!A1", result: 0 });
    // Reference a sheet the user never declared — auto-discovery must
    // not inflate the user's sheetNames either.
    Cell.setValue(ws, "A2", { formula: "[src.xlsx]OtherSheet!A1", result: 0 });

    const mutatedBuf = await Workbook.toBuffer(wb);
    await expectValidXlsx(mutatedBuf, { label: "no-user-link-mutation" });

    expect(link.rId).toBeUndefined();
    expect(link.sheetNames).toEqual(["Sheet1"]);
    expect(link).toEqual(originalSnapshot);
  });

  // T5: the Node streaming read path (`xlsx.read(stream)` →
  // `loadFromZipEntries`) is a separate code path from the buffer path
  // (`xlsx.load(buffer)` → `loadFromFiles`). Both must reconcile
  // externalLinks correctly; only the buffer path was covered above.
  it("reconstructs externalLinks via the Node stream read path", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Main");
    Workbook.addExternalLink(wb, {
      target: "ref.xlsx",
      sheetNames: ["Data"],
      cachedValues: { Data: { A1: 777 } }
    });
    Cell.setValue(ws, "A1", { formula: "[ref.xlsx]Data!A1", result: 777 });

    const buf = await Workbook.toBuffer(wb);
    await expectValidXlsx(buf, { label: "stream-read-path" });

    // Feed the bytes through a Node Readable to exercise the streaming
    // reader path (loadFromZipEntries), not the in-memory buffer path.
    const { Readable } = await import("node:stream");
    const stream = Readable.from([Buffer.from(buf)]);

    const wb2 = Workbook.create();
    await Workbook.readStream(wb2, stream);

    expect(wb2.externalLinks).toHaveLength(1);
    expect(wb2.externalLinks[0].target).toBe("ref.xlsx");
    expect(wb2.externalLinks[0].targetMode).toBe("External");
    expect(wb2.externalLinks[0].sheetNames).toEqual(["Data"]);
    expect(wb2.externalLinks[0].cachedValues).toEqual({ Data: { A1: 777 } });
  });

  it("works with addExternalLink() + numeric-form formula [1]Sheet!A1", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Main");

    Workbook.addExternalLink(wb, {
      target: "src.xlsx",
      sheetNames: ["Sheet1"],
      cachedValues: { Sheet1: { A1: 99 } }
    });

    // Write the formula using the numeric form directly — the user already
    // knows the index. The writer must accept it, upsert the sheet onto the
    // write-scoped copy (not the user's object), and produce a valid file.
    Cell.setValue(ws, "A1", { formula: "[1]Sheet1!A1", result: 99 });
    // Reference a sheet the user didn't declare — it should appear in the
    // output but NOT on the user's original sheetNames array.
    Cell.setValue(ws, "A2", { formula: "[1]Extra!B2", result: 0 });

    const originalSheetNames = [...wb.externalLinks[0].sheetNames];
    const buf = await Workbook.toBuffer(wb);
    await expectValidXlsx(buf, { label: "numeric-form-formula" });

    // User's object must not have been mutated.
    expect(wb.externalLinks[0].sheetNames).toEqual(originalSheetNames);

    // Reload and verify both sheets are present in the output.
    const wb2 = Workbook.create();
    await Workbook.read(wb2, buf);

    expect(wb2.externalLinks).toHaveLength(1);
    expect(wb2.externalLinks[0].target).toBe("src.xlsx");
    expect(wb2.externalLinks[0].sheetNames).toContain("Sheet1");
    expect(wb2.externalLinks[0].sheetNames).toContain("Extra");
  });

  it("round-trips cached error values (t='e') correctly", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Main");
    Workbook.addExternalLink(wb, {
      target: "errors.xlsx",
      sheetNames: ["Sheet1"],
      cachedValues: {
        Sheet1: {
          A1: "#DIV/0!",
          B1: "#REF!",
          C1: "#N/A",
          D1: "#VALUE!",
          E1: "#NAME?",
          F1: "#NUM!",
          G1: "#NULL!"
        }
      }
    });
    Cell.setValue(ws, "A1", { formula: "[errors.xlsx]Sheet1!A1", result: 0 });

    const buf = await Workbook.toBuffer(wb);
    await expectValidXlsx(buf, { label: "cached-error-values" });

    const wb2 = Workbook.create();
    await Workbook.read(wb2, buf);

    expect(wb2.externalLinks).toHaveLength(1);
    const cached = wb2.externalLinks[0].cachedValues?.Sheet1;
    expect(cached).toBeDefined();
    expect(cached!.A1).toBe("#DIV/0!");
    expect(cached!.B1).toBe("#REF!");
    expect(cached!.C1).toBe("#N/A");
    expect(cached!.D1).toBe("#VALUE!");
    expect(cached!.E1).toBe("#NAME?");
    expect(cached!.F1).toBe("#NUM!");
    expect(cached!.G1).toBe("#NULL!");
  });
});
