/**
 * An encrypted workbook — the OLE compound file (`D0 CF 11 E0 A1 B1 1A E1`) holding `EncryptionInfo` and
 * `EncryptedPackage` that Excel writes for a password-protected file — read without its `password` is reported as
 * `ExcelEncryptedError` by every reader: buffered XLSX and XLSB, `readFile`, `readStream`, `Stream.WorkbookReader`. It was reported as a corrupt ZIP
 * ("End of Central Directory not found"). A compound file without them (a legacy `.xls`) is an `ExcelFileError` that
 * says so, not "encrypted".
 */
import { mkdtemp, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Readable } from "node:stream";

import {
  Cell,
  ExcelEncryptedError,
  ExcelError,
  ExcelFileError,
  isExcelError,
  Stream,
  Workbook
} from "@excel";
import { WordCrypto } from "@word/crypto";
import { beforeAll, describe, expect, it } from "vitest";

let encrypted: Uint8Array;
let legacy: Uint8Array;
let directory: string;

beforeAll(async () => {
  const wb = Workbook.create();
  Cell.setValue(Workbook.addWorksheet(wb, "S"), "A1", "secret");
  encrypted = await WordCrypto.encrypt(await Workbook.toBuffer(wb), "pw");
  // A compound file that is not an encrypted package: the signature, then a `Workbook` stream's name.
  legacy = new Uint8Array(2048);
  legacy.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  legacy.set(
    [..."Workbook"].flatMap(ch => [ch.charCodeAt(0), 0]),
    1024
  );
  directory = await mkdtemp(join(tmpdir(), "dm-encrypted-"));
  await writeFile(join(directory, "enc.xlsx"), encrypted);
  await writeFile(join(directory, "enc.xlsb"), encrypted);
  await writeFile(join(directory, "old.xls"), legacy);
});

function chunked(bytes: Uint8Array): Readable {
  return Readable.from([bytes.subarray(0, 5), bytes.subarray(5, 700), bytes.subarray(700)]);
}

async function streamed(input: unknown): Promise<void> {
  for await (const _event of new Stream.WorkbookReader(input as never).parse()) {
    // Reading is what fails.
  }
}

const READERS: [string, (bytes: Uint8Array, name: string) => Promise<unknown>][] = [
  ["Workbook.read", bytes => Workbook.read(bytes)],
  ["Workbook.read xlsb", bytes => Workbook.read(bytes, { format: "xlsb" })],
  ["Workbook.readFile", (_bytes, name) => Workbook.readFile(join(directory, `${name}`))],
  [
    "Workbook.readFile .xlsb",
    (_bytes, name) => Workbook.readFile(join(directory, name.replace(".xlsx", ".xlsb")))
  ],
  ["Workbook.readStream", bytes => Workbook.readStream(chunked(bytes))],
  ["Workbook.readStream xlsb", bytes => Workbook.readStream(chunked(bytes), { format: "xlsb" })],
  ["Stream.WorkbookReader bytes", bytes => streamed(bytes)],
  ["Stream.WorkbookReader stream", bytes => streamed(chunked(bytes))]
];

describe("an encrypted workbook", () => {
  it.each(READERS)("%s reports ExcelEncryptedError", async (_name, read) => {
    const error = await read(encrypted, "enc.xlsx").then(
      () => undefined,
      (e: unknown) => e
    );
    expect(error).toBeInstanceOf(ExcelEncryptedError);
    expect(error).toBeInstanceOf(ExcelFileError);
    expect(error).toBeInstanceOf(ExcelError);
    expect(isExcelError(error)).toBe(true);
    expect((error as Error).message).toMatch(/encrypted \(password-protected\)/);
    // It pointed at `WordCrypto.decrypt`; Excel now reads an encrypted workbook itself, given its password.
    expect((error as Error).message).toMatch(/pass its password as the read's `password` option/);
  });

  it("decrypted with WordCrypto, reads", async () => {
    const read = await Workbook.read(await WordCrypto.decrypt(encrypted, "pw"));
    expect(Cell.getValue(Workbook.getWorksheet(read, "S")!, "A1")).toBe("secret");
  });
});

describe("a compound file that is not encrypted", () => {
  it.each(READERS.filter(([name]) => !name.includes(".xlsb")))(
    "%s says what it is",
    async (_name, read) => {
      const error = await read(legacy, "old.xls").then(
        () => undefined,
        (e: unknown) => e
      );
      expect(error).toBeInstanceOf(ExcelFileError);
      expect(error).not.toBeInstanceOf(ExcelEncryptedError);
      expect((error as Error).message).toMatch(/OLE compound file/);
      expect((error as Error).message).toMatch(/legacy \.xls/);
    }
  );
});
