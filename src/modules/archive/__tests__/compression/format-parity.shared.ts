/**
 * What gzip and zlib input decodes to, held identical on every platform. The expectations are
 * zlib's — what Node reads — measured on Node 26 with `zlib.gunzipSync` / `zlib.inflateSync`:
 *
 * - gzip: members follow one another and are concatenated (RFC 1952 §2.2, `cat a.gz b.gz`); zero
 *   bytes after the last are padding and end the file; anything else after a member is an error.
 * - zlib: whatever follows the Adler-32 is not part of the stream and is ignored.
 *
 * Every browser's `DecompressionStream` refuses all of these shapes, so the browser decodes them
 * with the pure-JS codec; this is what keeps the answer from depending on where it runs.
 */
import { ArchiveError, ArchiveLimitExceededError } from "@archive/core/errors";
import { describe, expect, it } from "vitest";

interface Codecs {
  gzipSync(data: Uint8Array): Uint8Array;
  zlibSync(data: Uint8Array): Uint8Array;
  gunzip(data: Uint8Array, options?: { maxOutputLength?: number }): Promise<Uint8Array>;
  gunzipSync(data: Uint8Array, options?: { maxOutputLength?: number }): Uint8Array;
  unzlib(data: Uint8Array, options?: { maxOutputLength?: number }): Promise<Uint8Array>;
  unzlibSync(data: Uint8Array, options?: { maxOutputLength?: number }): Uint8Array;
  decompressAuto(data: Uint8Array): Promise<Uint8Array>;
}

const encode = (text: string): Uint8Array => new TextEncoder().encode(text);
const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

function join(...parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}

export function defineFormatParitySuite(codecs: Codecs): void {
  const gz = (text: string) => codecs.gzipSync(encode(text));
  const zl = (text: string) => codecs.zlibSync(encode(text));

  /** Every async and sync entry for gzip input, so none can disagree with the others. */
  const gunzips = [
    ["gunzip", (d: Uint8Array) => codecs.gunzip(d)],
    ["gunzipSync", async (d: Uint8Array) => codecs.gunzipSync(d)],
    ["decompressAuto", (d: Uint8Array) => codecs.decompressAuto(d)]
  ] as const;
  const unzlibs = [
    ["unzlib", (d: Uint8Array) => codecs.unzlib(d)],
    ["unzlibSync", async (d: Uint8Array) => codecs.unzlibSync(d)],
    ["decompressAuto", (d: Uint8Array) => codecs.decompressAuto(d)]
  ] as const;

  describe("gzip input", () => {
    for (const [name, run] of gunzips) {
      it(`${name} concatenates every member`, async () => {
        expect(decode(await run(join(gz("hello "), gz("world"), gz("!"))))).toBe("hello world!");
      });

      it(`${name} treats zero bytes after the last member as padding`, async () => {
        expect(decode(await run(join(gz("hello "), gz("w"), new Uint8Array(5))))).toBe("hello w");
        // Padding ends the file: a member after it is not read
        expect(decode(await run(join(gz("hello "), new Uint8Array(2), gz("w"))))).toBe("hello ");
      });

      it(`${name} refuses other bytes after a member`, async () => {
        await expect(run(join(gz("hello "), Uint8Array.of(1, 2, 3)))).rejects.toThrow(ArchiveError);
      });

      it(`${name} refuses a truncated member`, async () => {
        const whole = gz("hello world");
        await expect(run(whole.subarray(0, whole.length - 2))).rejects.toThrow(ArchiveError);
      });
    }

    it("bounds the output of every member together", async () => {
      const two = join(gz("hello "), gz("world"));
      expect(decode(await codecs.gunzip(two, { maxOutputLength: 11 }))).toBe("hello world");
      for (const run of [
        () => codecs.gunzip(two, { maxOutputLength: 10 }),
        async () => codecs.gunzipSync(two, { maxOutputLength: 10 })
      ]) {
        const failure = run();
        await expect(failure).rejects.toThrow(ArchiveLimitExceededError);
        // The bound reported is the caller's, not what was left of it for the second member
        await expect(failure).rejects.toMatchObject({ allowed: 10 });
      }
    });
  });

  describe("zlib input", () => {
    for (const [name, run] of unzlibs) {
      it(`${name} ignores bytes after the Adler-32`, async () => {
        expect(decode(await run(join(zl("hello "), Uint8Array.of(1, 2, 3))))).toBe("hello ");
        expect(decode(await run(join(zl("hello "), zl("x"))))).toBe("hello ");
      });

      it(`${name} refuses a truncated stream`, async () => {
        const whole = zl("hello world");
        await expect(run(whole.subarray(0, whole.length - 2))).rejects.toThrow(ArchiveError);
      });
    }
  });
}
