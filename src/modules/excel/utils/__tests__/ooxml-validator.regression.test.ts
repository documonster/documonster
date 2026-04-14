import { ZipArchive } from "@archive";
import type { ExtractedFile } from "@archive/unzip/extract";
import { extractAll } from "@archive/unzip/extract";
import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { describe, it, expect } from "vitest";

import { Workbook } from "../../../../index";

const textDecoder = new TextDecoder();
const textEncoder = new TextEncoder();

function setFile(entries: Map<string, ExtractedFile>, path: string, data: Uint8Array): void {
  entries.set(path, {
    path,
    data,
    type: "file",
    size: data.length,
    mode: 0
  });
}

function rebuildZip(entries: Map<string, ExtractedFile>): Uint8Array {
  const zip = new ZipArchive({
    // Keep this deterministic-ish for tests; compression doesn't matter.
    level: 0,
    timestamps: "dos",
    modTime: new Date(1980, 0, 1, 0, 0, 0)
  });

  for (const [p, entry] of entries) {
    if (entry.type === "directory") {
      continue;
    }
    zip.add(p, entry.data);
  }

  return zip.bytesSync();
}

async function makeWorkbookWithSingleCheckbox(): Promise<Uint8Array> {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Sheet1");

  ws.addFormCheckbox("J2:K3", { link: "D6", checked: false, text: "J2:K3" });
  ws.getCell("D6").value = false;

  return wb.xlsx.writeBuffer();
}

describe("OOXML validator regressions (legacy form controls)", () => {
  it("flags legacyDrawing after controls", async () => {
    const buffer = await makeWorkbookWithSingleCheckbox();
    const entries = await extractAll(buffer);

    const sheetPath = "xl/worksheets/sheet1.xml";
    const sheet = entries.get(sheetPath);
    expect(sheet).toBeDefined();

    const xml = textDecoder.decode(sheet!.data);

    const legacyMatch = xml.match(/<legacyDrawing\b[^>]*\/>/);
    expect(legacyMatch).toBeTruthy();

    // Move <legacyDrawing/> to the end of the worksheet so it appears after <controls>.
    const legacyNode = legacyMatch![0];
    let mutated = xml.replace(legacyNode, "");
    mutated = mutated.replace("</worksheet>", `${legacyNode}</worksheet>`);

    setFile(entries, sheetPath, textEncoder.encode(mutated));
    const mutatedZip = rebuildZip(entries);

    const report = await validateXlsxBuffer(mutatedZip, { maxProblems: 50 });
    expect(report.ok).toBe(false);
    expect(report.problems.some(p => p.kind === "sheet-legacyDrawing-after-controls")).toBe(true);
  });

  it("flags controls without a DrawingML <drawing>", async () => {
    const buffer = await makeWorkbookWithSingleCheckbox();
    const entries = await extractAll(buffer);

    const sheetPath = "xl/worksheets/sheet1.xml";
    const sheet = entries.get(sheetPath);
    expect(sheet).toBeDefined();

    const xml = textDecoder.decode(sheet!.data);
    expect(xml.includes("<control")).toBe(true);

    const drawingMatch = xml.match(/<drawing\b[^>]*\/>/);
    expect(drawingMatch).toBeTruthy();

    // Simulate the original broken output: legacy <controls> but no DrawingML <drawing/>.
    const mutated = xml.replace(drawingMatch![0], "");

    setFile(entries, sheetPath, textEncoder.encode(mutated));
    const mutatedZip = rebuildZip(entries);

    const report = await validateXlsxBuffer(mutatedZip, { maxProblems: 50 });
    expect(report.ok).toBe(false);
    expect(report.problems.some(p => p.kind === "sheet-controls-missing-drawing")).toBe(true);
  });
});
