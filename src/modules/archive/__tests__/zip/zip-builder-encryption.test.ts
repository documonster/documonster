/**
 * `Archive.createZip` and `Archive.packZip` encrypt what they are asked to. Their options used to have
 * no `encryptionMethod` / `password`, and an unrecognised key is not an error, so an archive the
 * caller asked to protect was written in the clear — with a wrong password then opening it.
 */
import { Archive, ArchiveArgumentError } from "@archive/index";
import { describe, expect, it } from "vitest";

const TEXT = "the quick brown fox jumps over the lazy dog";
const decoder = new TextDecoder();
const METHODS = ["zipcrypto", "aes-128", "aes-192", "aes-256"] as const;

function includes(haystack: Uint8Array, needle: string): boolean {
  return decoder.decode(haystack).includes(needle);
}

async function read(bytes: Uint8Array, password?: string): Promise<string> {
  const files = await Archive.unpackZip(bytes, { password });
  return decoder.decode(files.get("a.txt")!.data);
}

type Options = { readonly encryptionMethod: (typeof METHODS)[number]; readonly password: string };

/** Every way of building a ZIP, each given the same options. Stored (`level: 0`), so plain text would show. */
const BUILDS: ReadonlyArray<readonly [string, (options: Options) => Promise<Uint8Array>]> = [
  ["packZip", options => Archive.packZip([["a.txt", TEXT]], { ...options, level: 0 })],
  [
    "createZip#toBuffer, bytes",
    options =>
      Archive.createZip({ ...options, level: 0 })
        .add("a.txt", TEXT)
        .toBuffer()
  ],
  [
    "createZip#toBuffer, a Blob",
    options =>
      Archive.createZip({ ...options, level: 0 })
        .add("a.txt", new Blob([TEXT]))
        .toBuffer()
  ],
  [
    "createZip#stream",
    async options => {
      const archive = Archive.createZip({ ...options, level: 0 }).add("a.txt", TEXT);
      const chunks = await Array.fromAsync(archive.stream());
      const bytes = new Uint8Array(chunks.reduce((size, chunk) => size + chunk.length, 0));
      let offset = 0;
      for (const chunk of chunks) {
        bytes.set(chunk, offset);
        offset += chunk.length;
      }
      return bytes;
    }
  ]
];

describe("ZIP builders encrypt", () => {
  describe.each(METHODS)("%s", method => {
    it.each(BUILDS)("by %s", async (_name, build) => {
      const bytes = await build({ encryptionMethod: method, password: "pw" });
      expect(includes(bytes, TEXT)).toBe(false);
      expect(await read(bytes, "pw")).toBe(TEXT);
      await expect(read(bytes, "wrong")).rejects.toThrow();
    });
  });

  it("writes ZipCrypto synchronously and refuses AES there", async () => {
    const sync = Archive.packZipSync([["a.txt", TEXT]], {
      encryptionMethod: "zipcrypto",
      password: "pw",
      level: 0
    });
    expect(includes(sync, TEXT)).toBe(false);
    expect(await read(sync, "pw")).toBe(TEXT);
    await expect(read(sync, "wrong")).rejects.toThrow();
    expect(() =>
      Archive.packZipSync([["a.txt", TEXT]], { encryptionMethod: "aes-256", password: "pw" })
    ).toThrow(/asynchronous/);
  });

  it.each([
    [{ encryptionMethod: "aes-256" }, /needs options\.password/],
    [{ password: "pw" }, /encrypts nothing without options\.encryptionMethod/],
    [{ encryptionMethod: "rot13", password: "pw" }, /options\.encryptionMethod/],
    [{ encryptionMethod: "aes-256", password: "" }, /options\.password/]
  ])("refuses %j before anything is built", (options, message) => {
    for (const call of [
      () => Archive.createZip(options as never),
      () => Archive.packZipSync([["a.txt", TEXT]], options as never)
    ]) {
      expect(call).toThrow(ArchiveArgumentError);
      expect(call).toThrow(message);
    }
  });
});

describe("the deprecated Archive.zip", () => {
  it("refuses to encrypt a TAR rather than write it in the clear", () => {
    expect(() =>
      // oxlint-disable-next-line typescript/no-deprecated
      Archive.zip({ format: "tar", encryptionMethod: "aes-256", password: "pw" } as never)
    ).toThrow(expect.objectContaining({ field: "options.encryptionMethod" }));
  });
});

describe("ZIP builder options", () => {
  it.each([null, [], 42])("refuses %j as options, naming the call that was made", value => {
    expect(() => Archive.createZip(value as never)).toThrow(/^Archive\.createZip: options must be/);
    expect(() => Archive.packZipSync([], value as never)).toThrow(
      /^Archive\.packZipSync: options must be/
    );
    expect(() => new Archive.ZipArchive(value as never)).toThrow(
      /^new Archive\.ZipArchive: options must be/
    );
  });
});
