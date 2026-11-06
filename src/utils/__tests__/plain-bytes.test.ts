import { describe, expect, it } from "vitest";

import { asPlainBytes, plainChunks } from "../binary";

describe("asPlainBytes", () => {
  it("returns a plain Uint8Array as it is", () => {
    const bytes = new Uint8Array([1, 2, 3]);
    expect(asPlainBytes(bytes)).toBe(bytes);
  });

  it("turns a subclass into a plain view of the same memory", () => {
    class Sub extends Uint8Array {}
    const sub = new Sub([1, 2, 3, 4]).subarray(1, 3);
    const plain = asPlainBytes(sub);
    expect(plain.constructor).toBe(Uint8Array);
    expect(plain.buffer).toBe(sub.buffer);
    expect([...plain]).toEqual([2, 3]);
  });

  it("copies bytes held in a SharedArrayBuffer", () => {
    if (typeof SharedArrayBuffer === "undefined") {
      return;
    }
    const shared = new Uint8Array(new SharedArrayBuffer(4));
    shared.set([1, 2, 3, 4]);
    const plain = asPlainBytes(shared.subarray(1));
    expect(plain.buffer).toBeInstanceOf(ArrayBuffer);
    expect([...plain]).toEqual([2, 3, 4]);
    shared[1] = 9;
    expect(plain[0]).toBe(2);
  });
});

describe("plainChunks", () => {
  it("maps each chunk and forwards an early exit to the source", async () => {
    let returned = false;
    async function* source(): AsyncGenerator<Uint8Array> {
      try {
        yield new Uint8Array([1]);
        yield new Uint8Array([2]);
      } finally {
        returned = true;
      }
    }
    const seen: number[] = [];
    for await (const chunk of plainChunks(source())) {
      expect(chunk.constructor).toBe(Uint8Array);
      seen.push(chunk[0]!);
      break;
    }
    expect(seen).toEqual([1]);
    expect(returned).toBe(true);
  });
});
