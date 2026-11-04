import { PassThrough } from "node:stream";

import { extractAll } from "@archive/unzip/extract";
import { cellSetNote, cellSetValue } from "@excel/core/cell";
import { WorkbookWriter } from "@excel/stream/workbook-writer";
import { describe, expect, it } from "vitest";

describe("streamed notes — VML shape ids", () => {
  it("gives every sheet's VML part its own block, so no two sheets share an id", async () => {
    const sink = new PassThrough();
    const chunks: Uint8Array[] = [];
    sink.on("data", chunk => chunks.push(chunk));
    const wb = new WorkbookWriter({ stream: sink });
    for (const name of ["A", "B"]) {
      const ws = wb.addWorksheet(name);
      for (const address of ["A1", "A2"]) {
        cellSetValue(ws.getCell(address), address);
        cellSetNote(ws.getCell(address), `${name} ${address}`);
      }
      ws.commit();
    }
    await wb.commit();
    const entries = await extractAll(Buffer.concat(chunks));
    const parts = [1, 2].map(index => {
      const xml = new TextDecoder().decode(entries.get(`xl/drawings/vmlDrawing${index}.vml`)!.data);
      return {
        idmap: /<o:idmap[^>]*data="([^"]*)"/.exec(xml)![1],
        ids: [...xml.matchAll(/id="_x0000_s(\d+)"/g)].map(match => Number(match[1]))
      };
    });
    expect(parts).toEqual([
      { idmap: "1", ids: [1025, 1026] },
      { idmap: "2", ids: [2049, 2050] }
    ]);
  });
});
