/**
 * `ZipEditor` copies an entry it does not touch as it found it: the raw DOS date and time,
 * and the general-purpose flags the payload depends on. A ZipCrypto entry's encryption
 * header carries a check byte taken from the DOS time when the entry has a data descriptor
 * (flag bit 3) and from the CRC otherwise, so dropping bit 3 or moving the time makes the
 * entry unreadable to Info-ZIP, 7-Zip and Windows — documonster's own reader accepts either
 * byte and cannot show it. The output is checked with the system `unzip` when there is one.
 */
import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Archive } from "@archive";
import { concatUint8Arrays } from "@utils/binary";
import { afterAll, afterEach, describe, expect, it } from "vitest";

const dir = mkdtempSync(path.join(tmpdir(), "documonster-zip-editor-preserve-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

const originalTz = process.env.TZ;
afterEach(() => {
  if (originalTz === undefined) {
    delete process.env.TZ;
  } else {
    process.env.TZ = originalTz;
  }
});

/** Info-ZIP's `echo secret-data | zip -P pw out.zip -`: ZipCrypto with a data descriptor. */
const infoZipZipCrypto = new Uint8Array(
  readFileSync(fileURLToPath(new URL("../unzip/data/zipcrypto-info-zip.zip", import.meta.url)))
);

interface CentralRecord {
  name: string;
  flags: number;
  dosTime: number;
  dosDate: number;
}

function centralDirectory(bytes: Uint8Array): CentralRecord[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = bytes.length - 22;
  while (view.getUint32(eocd, true) !== 0x06054b50) {
    eocd--;
  }
  const count = view.getUint16(eocd + 10, true);
  let p = view.getUint32(eocd + 16, true);
  const out: CentralRecord[] = [];
  for (let i = 0; i < count; i++) {
    const nameLength = view.getUint16(p + 28, true);
    out.push({
      name: new TextDecoder().decode(bytes.subarray(p + 46, p + 46 + nameLength)),
      flags: view.getUint16(p + 8, true),
      dosTime: view.getUint16(p + 12, true),
      dosDate: view.getUint16(p + 14, true)
    });
    p += 46 + nameLength + view.getUint16(p + 30, true) + view.getUint16(p + 32, true);
  }
  return out;
}

const record = (bytes: Uint8Array, name: string) =>
  centralDirectory(bytes).find(entry => entry.name === name)!;

let hasUnzip = true;
try {
  execFileSync("unzip", ["-v"], { stdio: "ignore" });
} catch {
  hasUnzip = false;
}

function systemUnzip(bytes: Uint8Array, entry: string, password?: string): string {
  const file = path.join(dir, `check-${Math.random().toString(36).slice(2)}.zip`);
  writeFileSync(file, bytes);
  const args = password ? ["-P", password, "-p", file, entry] : ["-p", file, entry];
  return execFileSync("unzip", args, { stdio: ["ignore", "pipe", "pipe"] }).toString();
}

async function collect(chunks: AsyncIterable<Uint8Array>): Promise<Uint8Array[]> {
  const out: Uint8Array[] = [];
  for await (const chunk of chunks) {
    out.push(chunk);
  }
  return out;
}

async function edited(source: Uint8Array, output: "buffer" | "stream"): Promise<Uint8Array> {
  const editor = await Archive.ZipEditor.open(source);
  editor.set("added.txt", "new");
  if (output === "buffer") {
    return editor.toBuffer();
  }
  return concatUint8Arrays(await collect(editor.stream()));
}

describe.each(["buffer", "stream"] as const)("ZipEditor %s output", output => {
  it("keeps a ZipCrypto entry with a data descriptor readable by other tools", async () => {
    const out = await edited(infoZipZipCrypto, output);
    const before = record(infoZipZipCrypto, "-");
    const after = record(out, "-");
    expect(after.flags & 0x0009).toBe(0x0009);
    expect([after.dosTime, after.dosDate]).toEqual([before.dosTime, before.dosDate]);
    if (hasUnzip) {
      expect(systemUnzip(out, "-", "pw").trim()).toBe("secret-data");
    }
  });

  it("keeps a ZipCrypto entry without a data descriptor readable by other tools", async () => {
    // ArchiveFile's buffered writer encrypts without a data descriptor (CRC check byte).
    const source = new Archive.ArchiveFile({ password: "pw", encryptionMethod: "zipcrypto" })
      .addText("secret-data\n", "plain.txt")
      .toBufferSync();
    const before = record(source, "plain.txt");
    expect(before.flags & 0x0009).toBe(0x0001);
    const out = await edited(source, output);
    const after = record(out, "plain.txt");
    expect(after.flags & 0x0008).toBe(before.flags & 0x0008);
    expect([after.dosTime, after.dosDate]).toEqual([before.dosTime, before.dosDate]);
    if (hasUnzip) {
      expect(systemUnzip(out, "plain.txt", "pw")).toBe("secret-data\n");
    }
  });

  it("keeps the raw DOS date and time of an untouched entry whatever the zone", async () => {
    process.env.TZ = "Asia/Tokyo";
    const date = new Date("2024-06-01T12:00:00Z");
    const source = Archive.createZip({ timestamps: "dos+utc" })
      .add("a.txt", "A", { modTime: date })
      .add("b.txt", "B", { modTime: date })
      .toBufferSync();
    process.env.TZ = "America/New_York";
    const editor = await Archive.ZipEditor.open(source);
    editor.rename("b.txt", "c.txt");
    const out =
      output === "buffer"
        ? await editor.toBuffer()
        : concatUint8Arrays(await collect(editor.stream()));
    for (const [from, to] of [
      ["a.txt", "a.txt"],
      ["b.txt", "c.txt"]
    ]) {
      const a = record(source, from);
      const b = record(out, to);
      expect([b.dosTime, b.dosDate]).toEqual([a.dosTime, a.dosDate]);
    }
    expect(Archive.unpackZipSync(out).get("a.txt")!.info.lastModified.getTime()).toBe(
      date.getTime()
    );
  });

  it("keeps a DOS-only wall time that does not exist in the editing zone", async () => {
    process.env.TZ = "UTC";
    // 02:30 on 2024-03-10 does not exist in New York (clocks go from 02:00 to 03:00).
    const source = Archive.createZip()
      .add("gap.txt", "x", { modTime: new Date(2024, 2, 10, 2, 30, 0) })
      .toBufferSync();
    process.env.TZ = "America/New_York";
    const out = await edited(source, output);
    const a = record(source, "gap.txt");
    const b = record(out, "gap.txt");
    expect([b.dosTime, b.dosDate]).toEqual([a.dosTime, a.dosDate]);
  });
});
