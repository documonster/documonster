/**
 * Members whose wrong arguments used to escape as a raw `TypeError` from inside: each is now
 * an `ArchiveArgumentError` naming the public call and the parameter, thrown before anything
 * is read or changed.
 */
import { Archive, ArchiveArgumentError } from "@archive";
import { describe, expect, it } from "vitest";

function caught(call: () => unknown): unknown {
  try {
    call();
  } catch (error) {
    return error;
  }
  return undefined;
}

async function caughtAsync(call: () => unknown): Promise<unknown> {
  try {
    await call();
  } catch (error) {
    return error;
  }
  return undefined;
}

function expectArgumentError(error: unknown, field: string, where: string): void {
  expect(error).toBeInstanceOf(ArchiveArgumentError);
  expect((error as ArchiveArgumentError).field).toBe(field);
  expect((error as Error).message.startsWith(`${where}: `)).toBe(true);
}

const bytes = new Uint8Array([1, 2, 3]);
const random = (n: number): Uint8Array => new Uint8Array(n);

describe("Archive.Crypto ZipCrypto primitives check their arguments", () => {
  it.each<[string, () => unknown]>([
    ["password", () => Archive.Crypto.createZipCryptoKeys(undefined as never)],
    ["password", () => Archive.Crypto.createZipCryptoKeys(42 as never)]
  ])("createZipCryptoKeys refuses %s", (field, call) => {
    expectArgumentError(caught(call), field, "Archive.Crypto.createZipCryptoKeys");
  });

  it.each<[string, () => unknown]>([
    ["data", () => (Archive.Crypto.encryptZipCrypto as (...a: unknown[]) => unknown)()],
    ["data", () => Archive.Crypto.encryptZipCrypto("x" as never, "pw", 0, random)],
    ["password", () => Archive.Crypto.encryptZipCrypto(bytes, null as never, 0, random)],
    ["crc32", () => Archive.Crypto.encryptZipCrypto(bytes, "pw", undefined as never, random)],
    ["crc32", () => Archive.Crypto.encryptZipCrypto(bytes, "pw", -1, random)],
    ["getRandomBytes", () => Archive.Crypto.encryptZipCrypto(bytes, "pw", 0, null as never)],
    ["getRandomBytes", () => Archive.Crypto.encryptZipCrypto(bytes, "pw", 0, () => random(3))]
  ])("encryptZipCrypto refuses %s", (field, call) => {
    expectArgumentError(caught(call), field, "Archive.Crypto.encryptZipCrypto");
  });

  it.each<[string, () => unknown]>([
    ["encryptedData", () => (Archive.Crypto.decryptZipCrypto as (...a: unknown[]) => unknown)()],
    ["encryptedData", () => Archive.Crypto.decryptZipCrypto({} as never, "pw", 0)],
    ["password", () => Archive.Crypto.decryptZipCrypto(bytes, 42 as never, 0)],
    ["crc32", () => Archive.Crypto.decryptZipCrypto(bytes, "pw", "0" as never)],
    ["lastModTime", () => Archive.Crypto.decryptZipCrypto(bytes, "pw", 0, 0x10000)]
  ])("decryptZipCrypto refuses %s", (field, call) => {
    expectArgumentError(caught(call), field, "Archive.Crypto.decryptZipCrypto");
  });

  it("still round-trips valid input", () => {
    const encrypted = Archive.Crypto.encryptZipCrypto(bytes, "pw", 0x12345678, random);
    expect(Archive.Crypto.decryptZipCrypto(encrypted, "pw", 0x12345678)).toEqual(bytes);
    expect(Archive.Crypto.createZipCryptoKeys("pw")).toEqual(
      Archive.Crypto.createZipCryptoKeys(new TextEncoder().encode("pw"))
    );
  });
});

describe("Archive.Crypto AES primitives and lookups check their arguments", () => {
  it.each<[string, string, () => unknown]>([
    [
      "encryptedData",
      "Archive.Crypto.decryptAes",
      () => Archive.Crypto.decryptAes(null as never, "pw", 256)
    ],
    [
      "password",
      "Archive.Crypto.decryptAes",
      () => Archive.Crypto.decryptAes(bytes, 42 as never, 256)
    ],
    ["data", "Archive.Crypto.encryptAes", () => Archive.Crypto.encryptAes("x" as never, "pw", 256)],
    [
      "password",
      "Archive.Crypto.encryptAes",
      () => Archive.Crypto.encryptAes(bytes, {} as never, 256)
    ]
  ])("refuses %s (%s)", async (field, where, call) => {
    expectArgumentError(await caughtAsync(call), field, where);
  });

  it("getEncryptionMethodName refuses a method it has no name for, including inherited keys", () => {
    for (const method of ["bogus", "toString", undefined]) {
      expectArgumentError(
        caught(() => Archive.Crypto.getEncryptionMethodName(method as never)),
        "method",
        "Archive.Crypto.getEncryptionMethodName"
      );
    }
    expect(Archive.Crypto.getEncryptionMethodName("aes-256")).toBe("AES-256");
  });

  it("getEncryptionMethodForAesKeyStrength refuses a strength the format has no method for", () => {
    expectArgumentError(
      caught(() => Archive.Crypto.getEncryptionMethodForAesKeyStrength(512 as never)),
      "strength",
      "Archive.Crypto.getEncryptionMethodForAesKeyStrength"
    );
    expect(Archive.Crypto.getEncryptionMethodForAesKeyStrength(128)).toBe("aes-128");
  });

  it("isAesEncryption / getAesKeyStrength do not answer for inherited keys", () => {
    expect(Archive.Crypto.isAesEncryption("toString" as never)).toBe(false);
    expect(Archive.Crypto.getAesKeyStrength("constructor" as never)).toBeUndefined();
    expect(Archive.Crypto.isAesEncryption("aes-192")).toBe(true);
    expect(Archive.Crypto.getAesKeyStrength("aes-192")).toBe(192);
  });
});

describe("Archive.toAsyncIterable / toReadableStream check their arguments when called", () => {
  it.each<[string, unknown]>([
    ["options", null],
    ["options", 42],
    ["options.signal", { signal: 42 }],
    ["options.onChunk", { onChunk: "x" }]
  ])("toAsyncIterable refuses %s", (field, options) => {
    expectArgumentError(
      caught(() => Archive.toAsyncIterable(bytes, options as never)),
      field,
      "Archive.toAsyncIterable"
    );
  });

  it("toAsyncIterable refuses a source before the first next()", () => {
    expectArgumentError(
      caught(() => Archive.toAsyncIterable(42 as never)),
      "source",
      "Archive.toAsyncIterable"
    );
  });

  it.each<[string, unknown]>([
    ["options", null],
    ["options.signal", { signal: {} }]
  ])("toReadableStream refuses %s", (field, options) => {
    expectArgumentError(
      caught(() => Archive.toReadableStream(bytes, options as never)),
      field,
      "Archive.toReadableStream"
    );
  });

  it("a chunk that is not bytes is an ArchiveArgumentError on source", async () => {
    async function* numbers(): AsyncGenerator<number> {
      yield 1;
    }
    const error = await caughtAsync(async () => {
      for await (const _ of Archive.toAsyncIterable(numbers())) {
        // drain
      }
    });
    expect(error).toBeInstanceOf(ArchiveArgumentError);
    expect((error as ArchiveArgumentError).field).toBe("source");
  });

  it("still yields valid input", async () => {
    const chunks: Uint8Array[] = [];
    for await (const chunk of Archive.toAsyncIterable("ab", {})) {
      chunks.push(chunk);
    }
    expect(chunks).toEqual([new TextEncoder().encode("ab")]);
  });
});

describe("Archive.ZipEditPlan checks the operations it is built from", () => {
  it.each<[string, unknown]>([
    ["ops", null],
    ["ops", {}],
    ["ops", "x"],
    ["ops[0]", [42]],
    ["ops[0]", [{ type: "bogus" }]],
    ["ops[0].name", [{ type: "delete" }]],
    ["ops[0].source", [{ type: "set", name: "a", source: 42 }]],
    [
      "ops[1].to",
      [
        { type: "delete", name: "a" },
        { type: "rename", from: "a" }
      ]
    ],
    ["ops[0].comment", [{ type: "comment", comment: 1 }]]
  ])("new ZipEditPlan refuses %s", (field, ops) => {
    expectArgumentError(
      caught(() => new Archive.ZipEditPlan(ops as never)),
      field,
      "Archive.ZipEditPlan"
    );
  });

  it.each<[string, () => unknown]>([
    ["name", () => new Archive.ZipEditPlan().delete(42 as never)],
    ["prefix", () => new Archive.ZipEditPlan().deleteDirectory(null as never)],
    ["from", () => new Archive.ZipEditPlan().rename(1 as never, "b")],
    ["to", () => new Archive.ZipEditPlan().rename("a", undefined as never)]
  ])("its methods refuse %s", (field, call) => {
    expect(caught(call)).toBeInstanceOf(ArchiveArgumentError);
    expect((caught(call) as ArchiveArgumentError).field).toBe(field);
  });

  it("deserialize refuses set data that is not base64, on every platform", () => {
    // Node's decoder skipped the characters it did not know, a browser's threw a DOMException
    expectArgumentError(
      caught(() =>
        Archive.ZipEditPlan.deserialize({
          version: 1,
          ops: [{ type: "set", name: "a.txt", data: "aGk!!" }]
        })
      ),
      "data.ops[0].data",
      "Archive.ZipEditPlan.deserialize"
    );
  });

  it("deserialize names data.ops in the caller's words", () => {
    expectArgumentError(
      caught(() => Archive.ZipEditPlan.deserialize({ version: 2, ops: [] } as never)),
      "data.version",
      "Archive.ZipEditPlan.deserialize"
    );
    expectArgumentError(
      caught(() => Archive.ZipEditPlan.deserialize({ version: 1, ops: [null as never] })),
      "data.ops[0]",
      "Archive.ZipEditPlan.deserialize"
    );
    expectArgumentError(
      caught(() =>
        Archive.ZipEditPlan.deserialize({ version: 1, ops: [{ type: "delete" } as never] })
      ),
      "data.ops[0].name",
      "Archive.ZipEditPlan.deserialize"
    );
    expectArgumentError(
      caught(() =>
        Archive.ZipEditPlan.deserialize({
          version: 1,
          ops: [{ type: "set", name: "a", data: 42 } as never]
        })
      ),
      "data.ops[0].data",
      "Archive.ZipEditPlan.deserialize"
    );
  });

  it("still accepts valid operations", () => {
    const plan = new Archive.ZipEditPlan([
      { type: "set", name: "a", source: bytes },
      { type: "rename", from: "a", to: "b" },
      { type: "comment" }
    ]);
    expect(plan.getOperations()).toHaveLength(3);
  });
});
