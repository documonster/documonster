import * as csvEntry from "@excel/bridge/csv";
import * as csvBridge from "@excel/bridge/csv-bridge";
import * as csvBridgeNode from "@excel/bridge/csv-bridge.node";
import * as csvNodeEntry from "@excel/bridge/csv.node";
import * as formulaEntry from "@excel/bridge/formula";
import * as markdownEntry from "@excel/bridge/markdown";
import * as markdownBridge from "@excel/bridge/markdown-bridge";
import * as markdownBridgeNode from "@excel/bridge/markdown-bridge.node";
import * as markdownNodeEntry from "@excel/bridge/markdown.node";
import { calculateFormulas } from "@excel/core/formula-adapter";
import { Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

describe("ExcelCsv namespace", () => {
  it("names every operation", () => {
    for (const entry of [csvEntry, csvNodeEntry]) {
      expect(entry.ExcelCsv.read).toBe(csvBridge.readCsvWorkbook);
      expect(entry.ExcelCsv.addWorksheet).toBe(csvBridge.readCsv);
      expect(entry.ExcelCsv.format).toBe(csvBridge.writeCsvString);
      expect(entry.ExcelCsv.toBuffer).toBe(csvBridge.writeCsvBuffer);
      expect(entry.ExcelCsv.writeStream).toBe(csvBridge.writeCsvStream);
      expect(entry.ExcelCsv.toBufferSync).toBe(csvBridge.writeCsvBufferSync);
    }
    expect(csvEntry.ExcelCsv.createImportStream).toBe(csvBridge.createCsvWriteStream);
    // The Node entry types it as the `stream.Writable` it is.
    expect(csvNodeEntry.ExcelCsv.createImportStream).toBe(csvBridgeNode.createCsvImportStream);
    expect(csvEntry.ExcelCsv.toStream).toBe(csvBridge.createCsvReadStream);
    expect(csvNodeEntry.ExcelCsv.toStream).toBe(csvBridgeNode.createCsvReadStream);
    expect(csvNodeEntry.ExcelCsv.readFile).toBe(csvBridgeNode.readCsvFileWorkbook);
    expect(csvNodeEntry.ExcelCsv.writeFile).toBe(csvBridgeNode.writeCsvFile);
  });

  it("keeps the deprecated flat exports", () => {
    expect(csvNodeEntry.readCsv).toBe(csvBridge.readCsv);
    expect(csvNodeEntry.writeCsv).toBe(csvBridge.writeCsv);
    expect(csvNodeEntry.writeCsvBuffer).toBe(csvBridge.writeCsvBuffer);
    expect(csvNodeEntry.createCsvWriteStream).toBe(csvBridge.createCsvWriteStream);
    expect(csvNodeEntry.createCsvReadStream).toBe(csvBridgeNode.createCsvReadStream);
    expect(csvNodeEntry.readCsvFile).toBe(csvBridgeNode.readCsvFile);
    expect(csvNodeEntry.writeCsvFile).toBe(csvBridgeNode.writeCsvFile);
    expect(csvEntry.createCsvReadStream).toBe(csvBridge.createCsvReadStream);
  });

  it("format returns the same text as the deprecated writeCsv", async () => {
    const wb = Workbook.create();
    await csvNodeEntry.ExcelCsv.read(wb, "a,b\n1,2");
    expect(csvNodeEntry.ExcelCsv.format(wb)).toBe(csvNodeEntry.writeCsv(wb));
    expect(csvNodeEntry.ExcelCsv.format(wb, { delimiter: ";" })).toBe(
      csvNodeEntry.writeCsv(wb, { delimiter: ";" })
    );
  });

  it("toBufferSync returns the bytes toBuffer resolves to", async () => {
    const wb = Workbook.create();
    await csvEntry.ExcelCsv.read(wb, "a,b\n1,2");
    expect(csvEntry.ExcelCsv.toBufferSync(wb)).toEqual(await csvEntry.ExcelCsv.toBuffer(wb));
    expect(new TextDecoder().decode(csvEntry.ExcelCsv.toBufferSync(wb))).toBe("a,b\n1,2");
  });

  it("createImportStream imports the CSV written into it as a new worksheet", async () => {
    const wb = Workbook.create();
    const sink = csvEntry.ExcelCsv.createImportStream(wb, { sheetName: "In" });
    const done = new Promise<void>((resolve, reject) => {
      (sink as unknown as NodeJS.EventEmitter).on("finish", resolve).on("error", reject);
    });
    sink.write("a,b\n1,2\n");
    sink.end();
    await done;
    expect(csvEntry.ExcelCsv.format(wb, { sheetName: "In" })).toBe("a,b\n1,2");
  });

  it("writeStream writes the same bytes as the deprecated writeCsv stream overload", async () => {
    const { PassThrough } = await import("node:stream");
    const wb = Workbook.create();
    await csvNodeEntry.ExcelCsv.read(wb, "a,b\n1,2");
    const capture = async (write: (out: InstanceType<typeof PassThrough>) => Promise<void>) => {
      const out = new PassThrough();
      const chunks: Buffer[] = [];
      out.on("data", (c: Buffer) => chunks.push(Buffer.from(c)));
      const result = await write(out);
      return { text: Buffer.concat(chunks).toString(), result };
    };
    const viaForward = await capture(out => csvNodeEntry.writeCsv(wb, out, { delimiter: ";" }));
    const viaNamespace = await capture(out =>
      csvNodeEntry.ExcelCsv.writeStream(wb, out, { delimiter: ";" })
    );
    expect(viaNamespace).toEqual(viaForward);
    expect(viaForward.text).toBe("a;b\n1;2");
  });
});

describe("ExcelMarkdown namespace", () => {
  it("names every operation", () => {
    for (const entry of [markdownEntry, markdownNodeEntry]) {
      expect(entry.ExcelMarkdown.read).toBe(markdownBridge.readMarkdownWorkbook);
      // One table or every table is `read`'s `tables` option, not a second name.
      expect("readAll" in entry.ExcelMarkdown).toBe(false);
      expect(entry.ExcelMarkdown.addWorksheet).toBe(markdownBridge.addMarkdownWorksheet);
      expect(entry.ExcelMarkdown.addWorksheets).toBe(markdownBridge.addMarkdownWorksheets);
      expect(entry.ExcelMarkdown.format).toBe(markdownBridge.writeMarkdown);
      expect(entry.ExcelMarkdown.toBuffer).toBe(markdownBridge.markdownToBuffer);
      expect(entry.ExcelMarkdown.toBufferSync).toBe(markdownBridge.writeMarkdownBuffer);
    }
    expect(markdownNodeEntry.ExcelMarkdown.readFile).toBe(
      markdownBridgeNode.readMarkdownFileWorkbook
    );
    expect("readAllFile" in markdownNodeEntry.ExcelMarkdown).toBe(false);
    expect(markdownNodeEntry.ExcelMarkdown.writeFile).toBe(markdownBridgeNode.writeMarkdownFile);
  });

  it("keeps the deprecated flat exports", () => {
    expect(markdownNodeEntry.readMarkdown).toBe(markdownBridge.readMarkdown);
    expect(markdownNodeEntry.readMarkdownAll).toBe(markdownBridge.readMarkdownAll);
    expect(markdownNodeEntry.writeMarkdown).toBe(markdownBridge.writeMarkdown);
    expect(markdownNodeEntry.writeMarkdownBuffer).toBe(markdownBridge.writeMarkdownBuffer);
    expect(markdownNodeEntry.readMarkdownFile).toBe(markdownBridgeNode.readMarkdownFile);
    expect(markdownNodeEntry.readMarkdownAllFile).toBe(markdownBridgeNode.readMarkdownAllFile);
    expect(markdownNodeEntry.writeMarkdownFile).toBe(markdownBridgeNode.writeMarkdownFile);
  });

  it("toBuffer resolves to the bytes the deprecated sync writeMarkdownBuffer returns", async () => {
    const wb = Workbook.create();
    markdownEntry.ExcelMarkdown.read(wb, "| a | b |\n| --- | --- |\n| 1 | 2 |");
    const pending = markdownEntry.ExcelMarkdown.toBuffer(wb);
    expect(pending).toBeInstanceOf(Promise);
    expect(await pending).toEqual(markdownEntry.writeMarkdownBuffer(wb));
    expect(markdownEntry.ExcelMarkdown.toBufferSync(wb)).toEqual(
      markdownEntry.writeMarkdownBuffer(wb)
    );
  });
});

describe("ExcelFormula namespace", () => {
  it("calculate is the adapter, and calculateFormulas forwards to it", () => {
    expect(formulaEntry.ExcelFormula.calculate).toBe(calculateFormulas);
    expect(formulaEntry.calculateFormulas).toBe(calculateFormulas);
  });
});
