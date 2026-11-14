/**
 * `Bytes` from `documonster/utils`: one answer on every platform, results that own their memory,
 * and wrong arguments refused as `UtilsArgumentError` naming the call and the field.
 */
import { Bytes, UtilsArgumentError, isUtilsError } from "@utils/index";
import { describe, expect, it } from "vitest";

function refusal(run: () => unknown): UtilsArgumentError {
  try {
    run();
  } catch (error) {
    expect(error).toBeInstanceOf(UtilsArgumentError);
    expect(isUtilsError(error)).toBe(true);
    return error as UtilsArgumentError;
  }
  throw new Error("expected a refusal");
}

export function defineBytesSuite(): void {
  describe("Bytes", () => {
    it("fromBase64 reads as Uint8Array.fromBase64 does, and refuses what it refuses", () => {
      expect([...Bytes.fromBase64("YWJj")]).toEqual([97, 98, 99]);
      expect([...Bytes.fromBase64(" YQ\n")]).toEqual([97]);
      expect([...Bytes.fromBase64("-_8", { alphabet: "base64url" })]).toEqual([251, 255]);
      for (const text of ["YQ=", "Y", "YW!Jj", "-_8"]) {
        const error = refusal(() => Bytes.fromBase64(text));
        expect(error.field).toBe("text");
        expect(error.message).toBe("Bytes.fromBase64: text is not base64 text");
      }
      expect(refusal(() => Bytes.fromBase64("+/8", { alphabet: "base64url" })).field).toBe("text");
    });

    it("toBase64 writes as bytes.toBase64 does", () => {
      const bytes = Uint8Array.of(251, 255, 191, 97);
      expect(Bytes.toBase64(bytes)).toBe("+/+/YQ==");
      expect(Bytes.toBase64(bytes, { alphabet: "base64url", omitPadding: true })).toBe("-_-_YQ");
    });

    it("fromHex / toHex read and write as Uint8Array.fromHex / toHex do", () => {
      expect([...Bytes.fromHex("0aBc")]).toEqual([10, 188]);
      expect(Bytes.toHex(Uint8Array.of(10, 188))).toBe("0abc");
      for (const text of ["abc", "0x00", "00 ff", "zz"]) {
        const error = refusal(() => Bytes.fromHex(text));
        expect(error.field).toBe("text");
        expect(error.message).toBe("Bytes.fromHex: text is not hex text");
      }
    });

    it("concat copies, even a single part", () => {
      const part = Uint8Array.of(1, 2, 3);
      const one = Bytes.concat([part]);
      expect(one).toEqual(part);
      one[0] = 9;
      expect(part[0]).toBe(1);
      expect([...Bytes.concat([Uint8Array.of(1), new Uint8Array(0), Uint8Array.of(2, 3)])]).toEqual(
        [1, 2, 3]
      );
      expect(Bytes.concat([])).toEqual(new Uint8Array(0));
    });

    it("fromText / toText are UTF-8 by default, and take TextDecoder's labels", () => {
      expect([...Bytes.fromText("é")]).toEqual([0xc3, 0xa9]);
      expect(Bytes.toText(Bytes.fromText("héllo"))).toBe("héllo");
      expect(Bytes.toText(Uint8Array.of(0xe9), { encoding: "windows-1252" })).toBe("é");
      expect(Bytes.toText(Uint8Array.of(0xff))).toBe("\uFFFD");
      expect(refusal(() => Bytes.toText(Uint8Array.of(0xff), { fatal: true })).field).toBe("bytes");
      expect(refusal(() => Bytes.toText(Uint8Array.of(1), { encoding: "klingon" })).field).toBe(
        "options.encoding"
      );
    });

    it("random fills beyond Web Crypto's per-call quota", () => {
      const bytes = Bytes.random(200_000);
      expect(bytes.length).toBe(200_000);
      // The last quota-sized piece was filled too
      expect(bytes.subarray(196_608).some(b => b !== 0)).toBe(true);
      expect(Bytes.random(0)).toEqual(new Uint8Array(0));
    });

    it("hands out plain Uint8Arrays over memory of their own", () => {
      for (const bytes of [
        Bytes.fromBase64("YWJj"),
        Bytes.fromHex("616263"),
        Bytes.concat([Uint8Array.of(1)]),
        Bytes.fromText("a"),
        Bytes.random(4)
      ]) {
        expect(bytes.constructor).toBe(Uint8Array);
        expect(bytes.byteOffset).toBe(0);
        expect(bytes.buffer.byteLength).toBe(bytes.length);
      }
    });

    it.each<[string, () => unknown, string]>([
      ["fromBase64 text", () => Bytes.fromBase64(42 as never), "text"],
      [
        "fromBase64 alphabet",
        () => Bytes.fromBase64("YQ", { alphabet: "hex" as never }),
        "options.alphabet"
      ],
      ["toBase64 bytes", () => Bytes.toBase64([1, 2] as never), "bytes"],
      [
        "toBase64 omitPadding",
        () => Bytes.toBase64(Uint8Array.of(1), { omitPadding: 1 as never }),
        "options.omitPadding"
      ],
      ["concat parts", () => Bytes.concat(Uint8Array.of(1) as never), "parts"],
      ["concat parts[1]", () => Bytes.concat([Uint8Array.of(1), "x" as never]), "parts[1]"],
      ["fromHex text", () => Bytes.fromHex(255 as never), "text"],
      ["toHex bytes", () => Bytes.toHex("ff" as never), "bytes"],
      ["fromText text", () => Bytes.fromText(undefined as never), "text"],
      ["toText bytes", () => Bytes.toText(new ArrayBuffer(1) as never), "bytes"],
      ["random length", () => Bytes.random(1.5), "length"],
      ["random negative", () => Bytes.random(-1), "length"]
    ])("refuses a wrong %s", (_name, run, field) => {
      const error = refusal(run);
      expect(error.field).toBe(field);
      expect(error.message).toMatch(/^Bytes\.\w+: /);
    });
  });
}
