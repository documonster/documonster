import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { Archive, ZipCrc32MismatchError } from "@archive";
import { afterAll, describe, expect, it } from "vitest";

const dir = mkdtempSync(path.join(tmpdir(), "documonster-archive-file-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

describe("ArchiveFile opening", () => {
  it("the constructor builds; fromFile / fromFileSync / fromBuffer read", async () => {
    const zip = new Archive.ArchiveFile();
    expect(zip).toBeInstanceOf(Archive.ArchiveFile);
    zip.addText("alpha", "a.txt");
    const bytes = await zip.toBuffer();
    const file = path.join(dir, "a.zip");
    await zip.writeToFile(file);

    const fromPath = await Archive.ArchiveFile.fromFile(file);
    const fromBytes = Archive.ArchiveFile.fromBuffer(bytes);
    const sync = Archive.ArchiveFile.fromFileSync(file);
    for (const archive of [fromPath, fromBytes, sync]) {
      expect(archive.getEntryNames()).toEqual(["a.txt"]);
    }
    expect(fromPath.sourcePath).toBe(path.resolve(file));
    expect(fromBytes.sourcePath).toBeNull();
  });

  it("opens TAR, gunzipping a .tgz path", async () => {
    const file = path.join(dir, "a.tgz");
    writeFileSync(file, await Archive.packTarGz([["t.txt", "tango"]]));

    const archive = await Archive.ArchiveFile.fromFile(file, { format: "tar" });
    expect(await archive.readAsText("t.txt")).toBe("tango");
    expect(await archive.readEntry("missing")).toBeNull();
  });
});

describe("ArchiveFile lookups", () => {
  it("return null for a missing entry", async () => {
    const zip = new Archive.ArchiveFile({ password: "pw", encryptionMethod: "aes-256" });
    zip.addText("secret", "s.txt", { comment: "c" });
    const bytes = await zip.bytes();
    expect(zip.bytesSync().length).toBeGreaterThan(0);
    const archive = Archive.ArchiveFile.fromBuffer(bytes);

    expect(archive.sourcePath).toBeNull();
    expect(archive.getEntry("missing")).toBeNull();
    expect(await archive.readEntry("missing")).toBeNull();
    expect(archive.readEntrySync("missing")).toBeNull();
    expect(await archive.readAsText("missing")).toBeNull();
    expect(archive.readAsTextSync("missing")).toBeNull();
    expect(archive.getZipEntryComment("missing")).toBeNull();
    expect(archive.getEntry("s.txt")?.comment).toBe("c");
    expect(archive.getZipEntryComment("s.txt")).toBe("c");
    const opened = Archive.ArchiveFile.fromBuffer(bytes, { password: "pw" });
    expect(await opened.readAsText("s.txt")).toBe("secret");
  });

  it("take an entry password and an encoding as options", async () => {
    const zip = new Archive.ArchiveFile({ password: "pw", encryptionMethod: "aes-256" });
    zip.addText("secret", "s.txt");
    const archive = Archive.ArchiveFile.fromBuffer(await zip.toBuffer());
    const decode = (data: Uint8Array | null) => new TextDecoder().decode(data!);

    expect(decode(await archive.readEntry("s.txt", { password: "pw" }))).toBe("secret");
    expect(await archive.readAsText("s.txt", { password: "pw", encoding: "utf-8" })).toBe("secret");

    // The sync readers cannot decrypt AES (Web Crypto is async); ZipCrypto they can.
    const legacy = new Archive.ArchiveFile({ password: "pw", encryptionMethod: "zipcrypto" });
    legacy.addText("old", "z.txt");
    const legacyArchive = Archive.ArchiveFile.fromBuffer(legacy.toBufferSync());
    expect(decode(legacyArchive.readEntrySync("z.txt", { password: "pw" }))).toBe("old");
    expect(legacyArchive.readAsTextSync("z.txt", { password: "pw", encoding: "utf-8" })).toBe(
      "old"
    );
  });

  it("keep the positional password and encoding", async () => {
    const zip = new Archive.ArchiveFile({ password: "pw", encryptionMethod: "aes-256" });
    zip.addText("secret", "s.txt");
    const archive = Archive.ArchiveFile.fromBuffer(await zip.toBuffer());
    expect(new TextDecoder().decode((await archive.readEntry("s.txt", "pw"))!)).toBe("secret");

    const plain = new Archive.ArchiveFile();
    plain.addText("hi", "p.txt");
    const plainFile = Archive.ArchiveFile.fromBuffer(plain.toBufferSync());
    expect(plainFile.readAsTextSync("p.txt", "utf-8")).toBe("hi");
    expect(new TextDecoder().decode(plainFile.readEntrySync("p.txt")!)).toBe("hi");
  });

  it("takes an entry password as an option", async () => {
    const zip = new Archive.ArchiveFile({ password: "pw", encryptionMethod: "aes-256" });
    zip.addText("secret", "s.txt");
    const archive = Archive.ArchiveFile.fromBuffer(await zip.toBuffer());
    expect(await archive.readAsText("s.txt", { password: "pw" })).toBe("secret");
  });

  it("names the public asynchronous alternative when AES cannot be handled synchronously", async () => {
    const message = (run: () => unknown): string => {
      try {
        run();
      } catch (error) {
        return (error as Error).message;
      }
      return "";
    };
    const build = (): Archive.ArchiveFile => {
      const zip = new Archive.ArchiveFile({ password: "pw", encryptionMethod: "aes-256" });
      zip.addText("secret", "s.txt");
      return zip;
    };

    const write = message(() => build().toBufferSync());
    expect(write).toMatch(/toBuffer\(\)/);
    expect(write).not.toMatch(/createZip/);

    const aes = await build().toBuffer();
    const read = message(() => Archive.unpackZipSync(aes, { password: "pw" }));
    expect(read).toMatch(/Archive\.unpackZip\b/);
    expect(read).not.toMatch(/extract(Sync)?\(\)/);
    const fileRead = message(() =>
      Archive.ArchiveFile.fromBuffer(aes, { password: "pw" }).readEntrySync("s.txt")
    );
    expect(fileRead).toMatch(/without the "Sync" suffix/);
  });

  it("verifies CRC-32 by default, with checkCrc32: false as the opt-out", async () => {
    const zip = await Archive.packZip([["a.txt", "payload that is long enough"]], { level: 0 });
    const at = Buffer.from(zip).indexOf("payload");
    zip[at]! ^= 0x20;
    const strict = Archive.ArchiveFile.fromBuffer(zip);
    await expect(strict.readEntry("a.txt")).rejects.toBeInstanceOf(ZipCrc32MismatchError);
    expect(() => strict.readEntrySync("a.txt")).toThrow(ZipCrc32MismatchError);
    const lenient = Archive.ArchiveFile.fromBuffer(zip, { checkCrc32: false });
    expect(await lenient.readAsText("a.txt")).toBe("Payload that is long enough");
    expect(Archive.ArchiveFile.fromBuffer(zip, { checkCrc32: false }).readAsTextSync("a.txt")).toBe(
      "Payload that is long enough"
    );
  });
});
