/**
 * Bytes that are not a workbook package this library can read are an `ExcelPackageError` whose `reason` says what was
 * found — the same reason from every reader, buffered and streamed — so a caller can tell a legacy `.xls` from a
 * damaged file, from something that is not a workbook, from the other format, without reading the message.
 */
import { mkdtemp, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Readable } from "node:stream";

import { packZip, unpackZip } from "@archive/zip-one-shot";
import { ExcelFileError, ExcelPackageError, isExcelError, Stream, Workbook } from "@excel";
import { encryptWorkbookPackage } from "@excel/utils/container-errors";
import { XmlLimitExceededError } from "@xml/errors";
import { beforeAll, describe, expect, it } from "vitest";

type Reason = ExcelPackageError["reason"];

let xlsx: Uint8Array;
let xlsb: Uint8Array;
let directory: string;

beforeAll(async () => {
  const wb = Workbook.create();
  Workbook.addWorksheet(wb, "S");
  xlsx = await Workbook.toBuffer(wb);
  xlsb = await Workbook.toBuffer(wb, { format: "xlsb" });
  directory = await mkdtemp(join(tmpdir(), "dm-excel-package-"));
});

/** A compound file with no encrypted package in it: what a legacy `.xls` looks like to the reader. */
function compoundFile(): Uint8Array {
  const bytes = new Uint8Array(512);
  bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  return bytes;
}

/** `bytes` with the first part's compressed data broken, so it does not inflate. */
function brokenDeflate(bytes: Uint8Array): Uint8Array {
  const out = bytes.slice();
  const nameLength = out[26]! | (out[27]! << 8);
  const extraLength = out[28]! | (out[29]! << 8);
  out[30 + nameLength + extraLength + 2]! ^= 0xff;
  return out;
}

async function failure(read: () => Promise<unknown>): Promise<unknown> {
  try {
    await read();
  } catch (error) {
    return error;
  }
  throw new Error("expected the read to fail");
}

async function drain(reader: AsyncIterable<AsyncIterable<unknown>>): Promise<void> {
  for await (const sheet of reader) {
    for await (const _row of sheet) {
      // read through
    }
  }
}

function expectReason(error: unknown, reason: Reason): void {
  expect(error).toBeInstanceOf(ExcelPackageError);
  expect(error).toBeInstanceOf(ExcelFileError);
  expect(isExcelError(error)).toBe(true);
  expect((error as ExcelPackageError).reason).toBe(reason);
  expect((error as ExcelPackageError).operation).toBe("read");
}

/** Every reader of an XLSX, given `bytes`. */
async function everyXlsxReader(name: string, bytes: Uint8Array): Promise<[string, unknown][]> {
  const path = join(directory, `${name}.xlsx`);
  await writeFile(path, bytes);
  return [
    ["Workbook.read", await failure(() => Workbook.read(bytes))],
    ["Workbook.readFile", await failure(() => Workbook.readFile(path))],
    ["Workbook.readStream", await failure(() => Workbook.readStream(Readable.from([bytes])))],
    [
      "Stream.WorkbookReader",
      await failure(() => drain(new Stream.WorkbookReader(Readable.from([bytes]))))
    ]
  ];
}

describe("ExcelPackageError", () => {
  it('reports a compound file that is not encrypted as "legacy-format"', async () => {
    for (const [reader, error] of await everyXlsxReader("legacy", compoundFile())) {
      expectReason(error, "legacy-format");
      expect((error as Error).message, reader).toContain("legacy .xls");
    }
  });

  it('reports bytes that are not a ZIP package as "not-a-workbook"', async () => {
    const text = new TextEncoder().encode("name,amount\nwidget,3\n");
    for (const [, error] of await everyXlsxReader("text", text)) {
      expectReason(error, "not-a-workbook");
    }
  });

  it('reports a truncated package as "damaged", whichever reader meets it', async () => {
    for (const [, error] of await everyXlsxReader("truncated", xlsx.slice(0, xlsx.length >> 1))) {
      expectReason(error, "damaged");
    }
  });

  it('reports compressed data that does not inflate as "damaged", not as an archive error', async () => {
    for (const [, error] of await everyXlsxReader("deflate", brokenDeflate(xlsx))) {
      expectReason(error, "damaged");
    }
  });

  // The streamed readers tell damaged from not-a-ZIP by the first bytes of what reaches the ZIP parser. For an
  // encrypted workbook that is the decrypted package — the compound file's own bytes said "legacy .xls".
  it('reports an encrypted package that decrypts to a damaged one as "damaged", not "legacy-format"', async () => {
    const encrypted = await encryptWorkbookPackage(xlsx.slice(0, xlsx.length >> 1), {
      password: "secret"
    });
    const options = { password: "secret" };
    const path = join(directory, "encrypted-truncated.xlsx");
    await writeFile(path, encrypted);
    for (const [reader, read] of [
      ["Workbook.read", () => Workbook.read(encrypted, options)],
      ["Workbook.readFile", () => Workbook.readFile(path, options)],
      ["Workbook.readStream", () => Workbook.readStream(Readable.from([encrypted]), options)],
      [
        "Stream.WorkbookReader",
        () => drain(new Stream.WorkbookReader(Readable.from([encrypted]), options))
      ]
    ] as const) {
      const error = await failure(read);
      expect((error as ExcelPackageError).reason, reader).toBe("damaged");
      expectReason(error, "damaged");
    }
  });

  it('reports an XLSB package read as XLSX as "format-mismatch"', async () => {
    expectReason(await failure(() => Workbook.read(xlsb, { format: "xlsx" })), "format-mismatch");
    const path = join(directory, "binary.xlsx");
    await writeFile(path, xlsb);
    expectReason(await failure(() => Workbook.readFile(path)), "format-mismatch");
    expectReason(
      await failure(() => Workbook.readStream(Readable.from([xlsb]))),
      "format-mismatch"
    );
  });

  it('reports an XLSX package read as XLSB as "format-mismatch"', async () => {
    expectReason(await failure(() => Workbook.read(xlsx, { format: "xlsb" })), "format-mismatch");
    const path = join(directory, "xml.xlsb");
    await writeFile(path, xlsx);
    expectReason(await failure(() => Workbook.readFile(path)), "format-mismatch");
  });

  it('reports a package with neither workbook part, read as XLSB, as "not-a-workbook"', async () => {
    const other = await packZip([["word/document.xml", "<w:document/>"]]);
    expectReason(await failure(() => Workbook.read(other, { format: "xlsb" })), "not-a-workbook");
  });

  it('reports a package without a workbook part as "not-a-workbook", not as an empty workbook', async () => {
    const document = await packZip([
      ["[Content_Types].xml", "<Types/>"],
      ["word/document.xml", "<w:document/>"]
    ]);
    for (const [, error] of await everyXlsxReader("document", document)) {
      expectReason(error, "not-a-workbook");
    }
    // A package that holds a sheet but lost its workbook part: nothing names the sheet.
    const orphanSheet = await packZip([
      ["xl/worksheets/sheet1.xml", "<worksheet><sheetData/></worksheet>"]
    ]);
    for (const [, error] of await everyXlsxReader("orphan", orphanSheet)) {
      expectReason(error, "not-a-workbook");
    }
  });

  it("detects the format when none is given, so a well-formed XLSB reads", async () => {
    await expect(Workbook.read(xlsb)).resolves.toBeDefined();
  });

  // Nesting past the XML parser's depth is a limit, not malformed markup: it is not renamed a parse error of the part.
  it("keeps the XML parser's depth limit as it is, not as a parse error of the part", async () => {
    const deep = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${"<x>".repeat(400)}${"</x>".repeat(400)}<sheetData/></worksheet>`;
    const entries = new Map<string, Uint8Array | string>();
    for (const [path, entry] of await unpackZip(xlsx)) {
      entries.set(path, path === "xl/worksheets/sheet1.xml" ? deep : entry.data);
    }
    for (const [reader, error] of await everyXlsxReader("deep", await packZip(entries))) {
      expect(error, reader).toBeInstanceOf(XmlLimitExceededError);
      expect(error, reader).toMatchObject({ limit: "depth" });
    }
  });

  // Flip a byte or cut the file anywhere and the read either succeeds or fails with an Excel error —
  // never the XML parser's, the archive's or a bare one.
  it("answers every damaged variant of a package with an Excel error", async () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "S");
    const stored = await Workbook.toBuffer(wb, { zip: { level: 0 } });
    let seed = 7;
    const random = () => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
    const variants: Uint8Array[] = [];
    for (const source of [xlsx, stored]) {
      for (let i = 0; i < 40; i++) {
        const copy = source.slice();
        copy[Math.floor(random() * copy.length)]! ^= 1 + Math.floor(random() * 254);
        variants.push(copy);
      }
      for (let i = 1; i < 20; i++) {
        variants.push(source.slice(0, Math.floor((source.length * i) / 20)));
      }
    }
    for (const bytes of variants) {
      for (const read of [
        () => Workbook.read(bytes),
        () => Workbook.readStream(Readable.from([bytes])),
        () => drain(new Stream.WorkbookReader(Readable.from([bytes])))
      ]) {
        const error = await read().then(
          () => undefined,
          (e: unknown) => e
        );
        if (error !== undefined) {
          expect(isExcelError(error), String(error)).toBe(true);
        }
      }
    }
  });
});
