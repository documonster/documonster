/**
 * `uuidV4` and its three tiers: `randomUUID`, `getRandomValues`, `Math.random`.
 *
 * Node-only, because the tiers can only be told apart by replacing
 * `globalThis.crypto`.
 */

import { uuidV4 } from "@utils/uuid";
import { afterEach, describe, expect, it, vi } from "vitest";

const V4 = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("uuidV4", () => {
  it("generates RFC 4122 v4 UUIDs", () => {
    expect(uuidV4()).toMatch(V4);
  });

  it("generates different UUIDs across calls", () => {
    expect(uuidV4()).not.toBe(uuidV4());
  });

  describe("tier 1 — crypto.randomUUID", () => {
    it("returns the native value verbatim when available", () => {
      const native = "0f9e8d7c-6b5a-4938-8271-605f4e3d2c1b";
      vi.stubGlobal("crypto", { randomUUID: () => native });
      expect(uuidV4()).toBe(native);
    });

    it("prefers randomUUID over getRandomValues", () => {
      const getRandomValues = vi.fn();
      vi.stubGlobal("crypto", {
        randomUUID: () => "0f9e8d7c-6b5a-4938-8271-605f4e3d2c1b",
        getRandomValues
      });
      uuidV4();
      expect(getRandomValues).not.toHaveBeenCalled();
    });
  });

  describe("tier 2 — crypto.getRandomValues", () => {
    it("is used when randomUUID is missing, as on plain HTTP", () => {
      const getRandomValues = vi.fn((buf: Uint8Array) => buf.fill(0xff));
      vi.stubGlobal("crypto", { getRandomValues });

      // All bits set, so the only zero bits left are the ones the version and
      // variant masks clear: byte 6 becomes 0x4f and byte 8 becomes 0xbf.
      expect(uuidV4()).toBe("ffffffff-ffff-4fff-bfff-ffffffffffff");
      expect(getRandomValues).toHaveBeenCalledTimes(1);
      expect(getRandomValues.mock.calls[0][0]).toHaveLength(16);
    });
  });

  describe("tier 3 — Math.random", () => {
    it("still produces a well-formed v4 with no Web Crypto at all", () => {
      vi.stubGlobal("crypto", undefined);
      expect(uuidV4()).toMatch(V4);
    });

    it("applies the same version and variant bits as the crypto tiers", () => {
      vi.stubGlobal("crypto", undefined);
      vi.spyOn(Math, "random").mockReturnValue(0);
      expect(uuidV4()).toBe("00000000-0000-4000-8000-000000000000");
    });

    it("is reached when crypto exists but exposes neither member", () => {
      vi.stubGlobal("crypto", {});
      expect(uuidV4()).toMatch(V4);
    });
  });
});
