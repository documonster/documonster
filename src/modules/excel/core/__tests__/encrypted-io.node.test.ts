/**
 * Encrypted workbooks, read and written through the Office encryption Word uses too: what Excel 16 saved with
 * "Encrypt with Password" (`data/office-encrypted/excel-encrypted.xlsx` / `.xlsb`, password `Excel-pw1`, made from a workbook this
 * library wrote — see `tmp/office-tests/enc/README.md`) is read by every reader given `{ password }`, and every writer
 * encrypts with it.
 */
import { readFileSync } from "node:fs";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { PassThrough, Readable } from "node:stream";

import {
  Cell,
  ExcelArgumentError,
  ExcelDecryptionError,
  ExcelEncryptedError,
  ExcelFileError,
  Stream,
  Workbook
} from "@excel";
import { rowGetModel } from "@excel/core/row";
import type { WorkbookData } from "@excel/core/workbook-core";
import { beforeAll, describe, expect, it } from "vitest";

const data = (name: string) =>
  new Uint8Array(
    readFileSync(new URL(`../../__tests__/data/office-encrypted/${name}`, import.meta.url))
  );
const OFFICE_XLSX = data("excel-encrypted.xlsx");
const OFFICE_XLSB = data("excel-encrypted.xlsb");
const PASSWORD = "Excel-pw1";

let directory: string;
beforeAll(async () => {
  directory = await mkdtemp(join(tmpdir(), "dm-excel-enc-"));
  await writeFile(join(directory, "office.xlsx"), OFFICE_XLSX);
  await writeFile(join(directory, "office.xlsb"), OFFICE_XLSB);
});

function chunked(bytes: Uint8Array): Readable {
  return Readable.from([bytes.subarray(0, 5), bytes.subarray(5, 700), bytes.subarray(700)]);
}

async function streamedCells(input: unknown, password?: string): Promise<unknown[]> {
  const values: unknown[] = [];
  const reader = new Stream.WorkbookReader(
    input as never,
    password === undefined ? {} : { password }
  );
  for await (const event of reader.parse()) {
    if (event.eventType === "worksheet") {
      for await (const row of event.value as AsyncIterable<Parameters<typeof rowGetModel>[0]>) {
        const model = rowGetModel(row) as { cells?: readonly { value?: unknown }[] } | null;
        values.push(...(model?.cells ?? []).map(cell => cell.value));
      }
    }
  }
  return values;
}

function a1(wb: WorkbookData): unknown {
  return Cell.getValue(Workbook.getWorksheet(wb, "Secret")!, "A1");
}

async function rejection(promise: Promise<unknown>): Promise<unknown> {
  return promise.then(
    () => undefined,
    (e: unknown) => e
  );
}

type Read = (password?: string) => Promise<unknown>;

const READERS: [string, Read][] = [
  ["Workbook.read xlsx", password => Workbook.read(OFFICE_XLSX, { password }).then(a1)],
  ["Workbook.read xlsb", password => Workbook.read(OFFICE_XLSB, { password }).then(a1)],
  [
    "Workbook.read base64",
    password =>
      Workbook.read(Buffer.from(OFFICE_XLSX).toString("base64"), { base64: true, password }).then(
        a1
      )
  ],
  ["Workbook.read Blob", password => Workbook.read(new Blob([OFFICE_XLSX]), { password }).then(a1)],
  [
    "Workbook.readWithDiagnostics xlsx",
    password => Workbook.readWithDiagnostics(OFFICE_XLSX, { password }).then(r => a1(r.workbook))
  ],
  [
    "Workbook.readWithDiagnostics xlsb",
    password => Workbook.readWithDiagnostics(OFFICE_XLSB, { password }).then(r => a1(r.workbook))
  ],
  [
    "Workbook.readFile .xlsx",
    password => Workbook.readFile(join(directory, "office.xlsx"), { password }).then(a1)
  ],
  [
    "Workbook.readFile .xlsb",
    password => Workbook.readFile(join(directory, "office.xlsb"), { password }).then(a1)
  ],
  [
    "Workbook.readStream",
    password => Workbook.readStream(chunked(OFFICE_XLSX), { password }).then(a1)
  ],
  [
    "Workbook.readStream xlsb",
    password => Workbook.readStream(chunked(OFFICE_XLSB), { password, format: "xlsb" }).then(a1)
  ],
  [
    "Stream.WorkbookReader bytes",
    password => streamedCells(OFFICE_XLSX, password).then(values => values[0])
  ],
  [
    "Stream.WorkbookReader stream",
    password => streamedCells(chunked(OFFICE_XLSX), password).then(values => values[0])
  ],
  [
    "Stream.WorkbookReader xlsb",
    password => streamedCells(OFFICE_XLSB, password).then(values => values[0])
  ]
];

describe("a workbook Excel encrypted", () => {
  it.each(READERS)("%s reads it with its password", async (_name, read) => {
    expect(await read(PASSWORD)).toBe("encrypted by Excel");
  });

  it.each(READERS)("%s: no password is ExcelEncryptedError", async (_name, read) => {
    const error = await rejection(read());
    expect(error).toBeInstanceOf(ExcelEncryptedError);
    expect((error as Error).message).toMatch(/pass its password as the read's `password` option/);
  });

  it.each(READERS)("%s: a wrong password is ExcelDecryptionError", async (_name, read) => {
    const error = await rejection(read("wrong"));
    expect(error).toBeInstanceOf(ExcelDecryptionError);
    expect(error).toBeInstanceOf(ExcelFileError);
    expect((error as ExcelDecryptionError).reason).toBe("password");
    expect((error as Error).message).toMatch(/the password is incorrect/);
  });

  it("bounds the decrypted package by zip.maxTotalUncompressedSize", async () => {
    const error = await rejection(
      Workbook.read(OFFICE_XLSX, { password: PASSWORD, zip: { maxTotalUncompressedSize: 1000 } })
    );
    expect(error).toBeInstanceOf(ExcelDecryptionError);
    expect((error as ExcelDecryptionError).reason).toBe("limit");
  });

  it("refuses a password that is not a string", async () => {
    const error = await rejection(Workbook.read(OFFICE_XLSX, { password: 42 as never }));
    expect(error).toBeInstanceOf(ExcelArgumentError);
    expect((error as ExcelArgumentError).field).toBe("options.password");
  });
});

describe("every writer encrypts with { password }", () => {
  const fast = { spinCount: 1000 } as const;
  function sample(): WorkbookData {
    const wb = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(wb, "Secret"), "A1", "written encrypted");
    return wb;
  }
  async function readBack(bytes: Uint8Array): Promise<unknown> {
    expect([...bytes.subarray(0, 4)]).toEqual([0xd0, 0xcf, 0x11, 0xe0]);
    expect(await rejection(Workbook.read(bytes))).toBeInstanceOf(ExcelEncryptedError);
    return a1(await Workbook.read(bytes, { password: "pw" }));
  }

  it.each(["xlsx", "xlsb"] as const)("Workbook.toBuffer %s", async format => {
    const bytes = await Workbook.toBuffer(sample(), { format, password: "pw", encryption: fast });
    expect(await readBack(bytes)).toBe("written encrypted");
    // The decrypted package is the format asked for.
    const report = await Workbook.readWithDiagnostics(bytes, { password: "pw" });
    expect(a1(report.workbook)).toBe("written encrypted");
  });

  it("uses Office's defaults without encryption options", async () => {
    const bytes = await Workbook.toBuffer(sample(), { password: "pw" });
    expect(new TextDecoder("latin1").decode(bytes)).toContain('spinCount="100000"');
  });

  it.each(["report.xlsx", "report.xlsb"])("Workbook.writeFile %s", async name => {
    const path = join(directory, name);
    await Workbook.writeFile(sample(), path, { password: "pw", encryption: fast });
    expect(await readBack(new Uint8Array(await readFile(path)))).toBe("written encrypted");
    expect(a1(await Workbook.readFile(path, { password: "pw" }))).toBe("written encrypted");
  });

  it.each(["xlsx", "xlsb"] as const)("Workbook.writeStream and toStream %s", async format => {
    const options = { format, password: "pw", encryption: fast };
    const sink = new PassThrough();
    const chunks: Uint8Array[] = [];
    sink.on("data", chunk => chunks.push(chunk));
    const ended = new Promise(resolve => sink.on("end", resolve));
    await Workbook.writeStream(sample(), sink, options);
    await ended;
    expect(await readBack(Buffer.concat(chunks))).toBe("written encrypted");
    const streamed: Uint8Array[] = [];
    for await (const chunk of Workbook.toStream(sample(), options)) {
      streamed.push(chunk as Uint8Array);
    }
    expect(await readBack(Buffer.concat(streamed))).toBe("written encrypted");
  });

  it.each(["xlsx", "xlsb"] as const)(
    "Workbook.writeStream %s to a Web WritableStream",
    async format => {
      const chunks: Uint8Array[] = [];
      const sink = new WritableStream<Uint8Array>({
        write(chunk) {
          chunks.push(chunk);
        }
      });
      const written = Workbook.writeStream(sample(), sink, {
        format,
        password: "pw",
        encryption: fast
      });
      const outcome = await Promise.race([
        written.then(() => "finished"),
        new Promise(resolve => setTimeout(() => resolve("hung"), 5000))
      ]);
      expect(outcome).toBe("finished");
      expect(await readBack(Buffer.concat(chunks))).toBe("written encrypted");
    }
  );

  it("writes an unencrypted package without a password, as before", async () => {
    const bytes = await Workbook.toBuffer(sample());
    expect([...bytes.subarray(0, 2)]).toEqual([0x50, 0x4b]);
  });

  it.each([
    [{ password: "" }, "options.password"],
    [{ password: 7 }, "options.password"],
    [{ password: "x".repeat(256) }, "options.password"],
    [{ encryption: fast }, "options.encryption"],
    [{ password: "pw", encryption: { keyBits: 100 } }, "options.encryption.keyBits"],
    [{ password: "pw", encryption: { hashAlgorithm: "MD5" } }, "options.encryption.hashAlgorithm"],
    [{ password: "pw", encryption: { spinCount: 1.5 } }, "options.encryption.spinCount"]
  ])("refuses %j naming %s", async (options, field) => {
    for (const write of [
      () => Workbook.toBuffer(sample(), options as never),
      () => Workbook.writeFile(sample(), join(directory, "never.xlsx"), options as never)
    ]) {
      const error = await rejection(write());
      expect(error).toBeInstanceOf(ExcelArgumentError);
      expect((error as ExcelArgumentError).field).toBe(field);
    }
    expect(() => Workbook.toStream(sample(), options as never)).toThrow(ExcelArgumentError);
  });
});

describe("a workbook encrypted by Word's WordCrypto.encrypt reads too (one implementation)", () => {
  it("decrypts what WordCrypto wrote", async () => {
    const { WordCrypto } = await import("@word/crypto");
    const wb = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(wb, "Secret"), "A1", "via word");
    const bytes = await WordCrypto.encrypt(await Workbook.toBuffer(wb), "pw", { spinCount: 10 });
    expect(a1(await Workbook.read(bytes, { password: "pw" }))).toBe("via word");
  });
});
