/**
 * VML shape ids are numbered across the workbook: every VML part declares its own 1024-id blocks in
 * `o:idmap`, no two parts share a block, and every shape id falls inside its part's blocks. Each sheet
 * used to number its notes from 1025 and its header/footer pictures from 2049, so two sheets wrote the
 * same ids.
 */
import { extractAll } from "@archive/unzip/extract";
import { Cell, Form, HeaderFooterImage, Image, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

const PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

function build(sheets = 3): Workbook.Handle {
  const wb = Workbook.create();
  const imageId = Image.add(wb, { base64: PNG, extension: "png" });
  for (let s = 1; s <= sheets; s++) {
    const ws = Workbook.addWorksheet(wb, `S${s}`);
    Cell.setNote(ws, "A1", `note ${s}a`);
    Cell.setNote(ws, "A2", `note ${s}b`);
    Form.addCheckbox(ws, "C2:D3", { text: `box ${s}a` });
    Form.addCheckbox(ws, "C5:D6", { text: `box ${s}b` });
    HeaderFooterImage.set(ws, { imageId, position: "CH" });
    HeaderFooterImage.set(ws, { imageId, position: "LF" });
  }
  return wb;
}

interface VmlPart {
  path: string;
  blocks: number[];
  ids: number[];
}

async function vmlParts(
  bytes: Uint8Array
): Promise<{ parts: VmlPart[]; text: Map<string, string> }> {
  const entries = await extractAll(bytes);
  const text = new Map<string, string>();
  const parts: VmlPart[] = [];
  for (const [path, entry] of entries) {
    const xml = new TextDecoder().decode(entry.data);
    text.set(path, xml);
    if (!path.endsWith(".vml")) {
      continue;
    }
    const data = /<o:idmap[^>]*data="([^"]*)"/.exec(xml)![1];
    const ids = [...xml.matchAll(/(?:\bid|o:spid)="_x0000_s(\d+)"/g)].map(match =>
      Number(match[1])
    );
    parts.push({ path, blocks: data.split(",").map(Number), ids });
  }
  return { parts, text };
}

function expectWorkbookWideIds(parts: readonly VmlPart[]): void {
  const allIds = parts.flatMap(part => part.ids);
  expect(new Set(allIds).size).toBe(allIds.length);
  const allBlocks = parts.flatMap(part => part.blocks);
  expect(new Set(allBlocks).size).toBe(allBlocks.length);
  for (const part of parts) {
    expect(part.ids.length).toBeGreaterThan(0);
    for (const id of part.ids) {
      expect(part.blocks).toContain(Math.floor(id / 1024));
    }
  }
}

describe("VML shape ids across sheets", () => {
  for (const format of ["xlsx", "xlsb"] as const) {
    it(`gives every ${format} VML part its own blocks and unique ids`, async () => {
      const { parts } = await vmlParts(await Workbook.toBuffer(build(), { format }));
      // Three sheets, each with a notes+controls part and a header/footer part.
      expect(parts).toHaveLength(6);
      expectWorkbookWideIds(parts);
    });

    it(`keeps the ids unique through a ${format} read and rewrite`, async () => {
      const first = await Workbook.toBuffer(build(), { format });
      const reopened = Workbook.create();
      await Workbook.read(reopened, first);
      const { parts } = await vmlParts(await Workbook.toBuffer(reopened, { format }));
      expect(parts).toHaveLength(6);
      expectWorkbookWideIds(parts);
    });
  }

  it("keeps a checkbox's own id when its block is free, and names it alike in sheet, drawing and VML", async () => {
    const wb = build();
    // Ids a file carried: sheet 2's controls live in block 3, which nothing else claims first.
    const [a, b] = Form.list(Workbook.getWorksheet(wb, "S2")!);
    a.model.shapeId = 3073;
    b.model.shapeId = 3074;
    const { parts, text } = await vmlParts(await Workbook.toBuffer(wb));
    expectWorkbookWideIds(parts);
    const controlIds = (index: number) =>
      [...text.get(`xl/worksheets/sheet${index}.xml`)!.matchAll(/<control shapeId="(\d+)"/g)].map(
        m => Number(m[1])
      );
    expect(controlIds(1)).toEqual([1025, 1026]);
    expect(controlIds(2)).toEqual([3073, 3074]);
    for (const index of [1, 2, 3]) {
      const vml = text.get(`xl/drawings/vmlDrawing${index}.vml`)!;
      const drawing = [...text.keys()].filter(path => /^xl\/drawings\/drawing\d+\.xml$/.test(path));
      for (const id of controlIds(index)) {
        expect(vml).toContain(`id="_x0000_s${id}"`);
        expect(drawing.some(path => text.get(path)!.includes(`spid="_x0000_s${id}"`))).toBe(true);
      }
    }
  });

  it("does not change the live checkbox when the write renumbers it", async () => {
    const wb = build(2);
    const second = Workbook.getWorksheet(wb, "S2")!;
    const live = Form.list(second).map(box => box.model.shapeId);
    // Both sheets number their own checkboxes from 1025; one of them is renumbered in the file.
    expect(live).toEqual([1025, 1026]);
    const { text } = await vmlParts(await Workbook.toBuffer(wb));
    const written = [
      ...text.get("xl/worksheets/sheet2.xml")!.matchAll(/<control shapeId="(\d+)"/g)
    ].map(m => Number(m[1]));
    expect(written).not.toEqual(live);
    expect(Form.list(second).map(box => box.model.shapeId)).toEqual(live);
    // The default name is the position in the block, as Excel names it, not the raw id.
    expect(text.get("xl/worksheets/sheet2.xml")).toContain('name="Check Box 1"');
  });

  it("keeps a checkbox's name when the write renumbers it", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    for (const range of ["C2:D3", "C5:D6", "C8:D9"]) {
      Form.addCheckbox(ws, range, { text: "box" });
    }
    // The survivor keeps id 1027, third in its block: Excel names it "Check Box 3".
    const [first, second] = Form.list(ws);
    Form.remove(ws, first);
    Form.remove(ws, second);
    // A copy carries the same id, so the write renumbers one of the two into a fresh block — and a
    // name derived from the written id made the control's name depend on which sheet came first.
    Workbook.importWorksheet(wb, ws, "Copy");
    const { parts, text } = await vmlParts(await Workbook.toBuffer(wb));
    expectWorkbookWideIds(parts);
    const controls = (sheet: string) =>
      [
        ...text
          .get(`xl/worksheets/${sheet}.xml`)!
          .matchAll(/<control shapeId="(\d+)"[^>]*name="([^"]*)"/g)
      ].map(m => [Number(m[1]), m[2]]);
    expect(controls("sheet1")).toEqual([[1027, "Check Box 3"]]);
    const [[copyId, copyName]] = controls("sheet2");
    expect(copyId).not.toBe(1027);
    expect(copyName).toBe("Check Box 3");
  });
});
