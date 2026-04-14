import { uuidV4 } from "@utils/uuid";
import { describe, expect, it } from "vitest";

describe("uuidV4", () => {
  it("generates RFC 4122 v4 UUIDs", () => {
    const id = uuidV4();

    // xxxxxxxx-xxxx-4xxx-[89ab]xxx-xxxxxxxxxxxx
    expect(id).toMatch(/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i);
  });

  it("generates different UUIDs across calls", () => {
    const a = uuidV4();
    const b = uuidV4();
    expect(a).not.toBe(b);
  });
});
