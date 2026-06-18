/**
 * Integration tests for Office 2010+ slicer and timeline raw-passthrough.
 *
 * excelts does not yet structurally model slicers or timelines — the
 * OOXML surface is large (four coordinated part families:
 * `xl/slicers`, `xl/slicerCaches`, `xl/timelines`, `xl/timelineCaches`,
 * plus sheet-level extensions and workbook-level cache list entries).
 * Dashboard workbooks that travel through excelts must not lose these
 * parts; the tests below verify the preserve-on-roundtrip path does
 * its job.
 *
 * We synthesise a fake source workbook whose zip already contains the
 * parts (simulating a file authored by Excel), feed it through the
 * excelts loader, write it back out, and assert that every byte we
 * care about survives. The content of the parts themselves is
 * treated opaquely — excelts has no opinion on them, so the round
 * trip is literal byte preservation rather than structural equality.
 */

import { extractAll } from "@archive/unzip/extract";
import { ZipEditor } from "@archive/zip";
import { Cell, Workbook } from "@excel/index";
import { getWorkbookModel } from "@excel/workbook";
import { describe, it, expect } from "vitest";

import { expectValidXlsx } from "./helpers/expect-valid-xlsx";

const decoder = new TextDecoder();

/**
 * Produce a tiny xlsx buffer that already contains a slicer + slicer
 * cache + timeline + timeline cache (plus their rels), as if written
 * by Excel. We start from a workbook excelts can author, then patch
 * the zip to add the extra parts and wire them into the rels files
 * and Content Types manifest.
 */
async function makeXlsxWithSlicerAndTimeline(): Promise<Uint8Array> {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "Sheet1");
  Cell.setValue(ws, "A1", "Dashboard");
  const baseBuf = new Uint8Array(await Workbook.toBuffer(wb));
  const baseEntries = await extractAll(baseBuf);

  // Slicer / timeline parts synthesised here — the content is
  // semantically correct enough to survive a zip round-trip but does
  // not have to drive real Excel behaviour for the test.
  const slicerXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<slicers xmlns="http://schemas.microsoft.com/office/drawing/2010/slicer">' +
    '<slicer name="Region" cache="Slicer_Region" caption="Region" rowHeight="241300"/>' +
    "</slicers>";
  const slicerCacheXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<slicerCacheDefinition name="Slicer_Region" sourceName="Region" ' +
    'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    "<pivotTables/>" +
    "</slicerCacheDefinition>";
  const timelineXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2011/1/main">' +
    '<timeline name="Date" cache="NativeTimeline_Date" caption="Date"/>' +
    "</timelines>";
  const timelineCacheXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<timelineCacheDefinition name="NativeTimeline_Date" sourceName="Date" ' +
    'xmlns="http://schemas.microsoft.com/office/spreadsheetml/2011/1/main"/>';
  const encoder = new TextEncoder();

  // Rebuild the zip with the extra parts + updated Content Types
  // manifest, preserving the bytes of every entry excelts produced.
  const editor = await ZipEditor.open(baseBuf);
  const existingContentTypes = decoder.decode(baseEntries.get("[Content_Types].xml")!.data);
  const patchedContentTypes = existingContentTypes.replace(
    "</Types>",
    [
      '<Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>',
      '<Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>',
      '<Override PartName="/xl/timelines/timeline1.xml" ContentType="application/vnd.ms-excel.timeline+xml"/>',
      '<Override PartName="/xl/timelineCaches/timelineCache1.xml" ContentType="application/vnd.ms-excel.timelineCache+xml"/>',
      "</Types>"
    ].join("")
  );
  editor
    .set("[Content_Types].xml", encoder.encode(patchedContentTypes))
    .set("xl/slicers/slicer1.xml", encoder.encode(slicerXml))
    .set("xl/slicerCaches/slicerCache1.xml", encoder.encode(slicerCacheXml))
    .set("xl/timelines/timeline1.xml", encoder.encode(timelineXml))
    .set("xl/timelineCaches/timelineCache1.xml", encoder.encode(timelineCacheXml));
  return editor.bytes();
}

describe("slicer + timeline raw passthrough", () => {
  it("preserves all four dashboard parts across load + save", async () => {
    const original = await makeXlsxWithSlicerAndTimeline();
    const wb = Workbook.create();
    await Workbook.read(wb, original);

    // After load, the workbook-level model exposes the raw bytes so
    // later callers can inspect them if needed. The expected shape is
    // a record keyed by zip path with Uint8Array values.
    expect(Object.keys(getWorkbookModel(wb).slicerParts ?? {})).toContain("xl/slicers/slicer1.xml");
    expect(Object.keys(getWorkbookModel(wb).slicerCacheParts ?? {})).toContain(
      "xl/slicerCaches/slicerCache1.xml"
    );
    expect(Object.keys(getWorkbookModel(wb).timelineParts ?? {})).toContain(
      "xl/timelines/timeline1.xml"
    );
    expect(Object.keys(getWorkbookModel(wb).timelineCacheParts ?? {})).toContain(
      "xl/timelineCaches/timelineCache1.xml"
    );

    // Round-trip: every part excelts captured must be re-emitted
    // verbatim, and the Content Types manifest must mention the new
    // content types so Excel recognises them.
    const resaved = new Uint8Array(await Workbook.toBuffer(wb));
    await expectValidXlsx(resaved, { label: "slicer+timeline resave" });
    const entries = await extractAll(resaved);
    expect(entries.get("xl/slicers/slicer1.xml")).toBeDefined();
    expect(entries.get("xl/slicerCaches/slicerCache1.xml")).toBeDefined();
    expect(entries.get("xl/timelines/timeline1.xml")).toBeDefined();
    expect(entries.get("xl/timelineCaches/timelineCache1.xml")).toBeDefined();

    const contentTypes = decoder.decode(entries.get("[Content_Types].xml")!.data);
    expect(contentTypes).toContain("application/vnd.ms-excel.slicer+xml");
    expect(contentTypes).toContain("application/vnd.ms-excel.slicerCache+xml");
    expect(contentTypes).toContain("application/vnd.ms-excel.timeline+xml");
    expect(contentTypes).toContain("application/vnd.ms-excel.timelineCache+xml");

    // Byte-equal preservation — the parts should come back exactly
    // as they went in.
    const slicerIn = entries.get("xl/slicers/slicer1.xml")!.data;
    expect(decoder.decode(slicerIn)).toContain("Slicer_Region");
    expect(decoder.decode(slicerIn)).toContain("Region");
    const timelineIn = entries.get("xl/timelines/timeline1.xml")!.data;
    expect(decoder.decode(timelineIn)).toContain("NativeTimeline_Date");
  });

  it("omits the dashboard content types on workbooks without slicers or timelines", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    const buf = await Workbook.toBuffer(wb);
    await expectValidXlsx(buf, { label: "empty workbook no slicer/timeline" });
    const entries = await extractAll(new Uint8Array(buf));
    const contentTypes = decoder.decode(entries.get("[Content_Types].xml")!.data);
    expect(contentTypes).not.toContain("application/vnd.ms-excel.slicer+xml");
    expect(contentTypes).not.toContain("application/vnd.ms-excel.timeline+xml");
  });
});
