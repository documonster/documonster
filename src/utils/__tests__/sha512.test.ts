/**
 * SHA-384 / SHA-512 against FIPS 180-4 test vectors. Runs in the browser too,
 * where the pure implementation is the one under test.
 */

import { sha384, sha512 } from "@utils/crypto";
import { describe, expect, it } from "vitest";

const hex = (b: Uint8Array) => Array.from(b, x => x.toString(16).padStart(2, "0")).join("");
const ascii = (s: string) => new TextEncoder().encode(s);
const TWO_BLOCK =
  "abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu";

describe("sha512 / sha384", () => {
  it("matches the FIPS 180-4 vectors", () => {
    expect(hex(sha512(ascii("abc")))).toBe(
      "ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a" +
        "2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f"
    );
    expect(hex(sha512(ascii(TWO_BLOCK)))).toBe(
      "8e959b75dae313da8cf4f72814fc143f8f7779c6eb9f7fa17299aeadb6889018" +
        "501d289e4900f7e4331b99dec4b5433ac7d329eeb6dd26545e96e55b874be909"
    );
    expect(hex(sha512(new Uint8Array(0)))).toBe(
      "cf83e1357eefb8bdf1542850d66d8007d620e4050b5715dc83f4a921d36ce9ce" +
        "47d0d13c5d85f2b0ff8318d2877eec2f63b931bd47417a81a538327af927da3e"
    );
    expect(hex(sha384(ascii("abc")))).toBe(
      "cb00753f45a35e8bb5a03d699ac65007272c32ab0eded1631a8b605a43ff5bed" +
        "8086072ba1e7cc2358baeca134c825a7"
    );
    expect(hex(sha384(ascii(TWO_BLOCK)))).toBe(
      "09330c33f71147e83d192fc782cd1b4753111b173b3b05d22fa08086e3b0f712" +
        "fcc7c71a557e2db966c3e9fa91746039"
    );
  });
});
