import { describe, expect, it } from "vitest";

import { escapeRegExp } from "../regexp";

describe("escapeRegExp", () => {
  it("escapes every syntax character", () => {
    for (const c of [".", "*", "+", "?", "^", "$", "{", "}", "(", ")", "|", "[", "]", "\\", "/"]) {
      expect(escapeRegExp(c)).toBe(`\\${c}`);
    }
    expect(escapeRegExp("-")).toBe("\\x2d");
  });

  it("leaves other characters unchanged", () => {
    expect(escapeRegExp("hello")).toBe("hello");
    expect(escapeRegExp("abc123")).toBe("abc123");
    expect(escapeRegExp("")).toBe("");
  });

  it("matches the text literally, with or without the u and v flags and inside a class", () => {
    const text = "[a-z]+ (foo|bar) a.b*c ^$ {1} \\ /";
    for (const flags of ["", "u", "v"]) {
      const re = new RegExp(`^${escapeRegExp(text)}$`, flags);
      expect(re.test(text)).toBe(true);
      expect(re.test(text.replace(".", "_"))).toBe(false);
    }
    expect(new RegExp(`^[${escapeRegExp("a-c")}]+$`, "u").test("-ac")).toBe(true);
    expect(new RegExp(`^[${escapeRegExp("a-c")}]$`, "u").test("b")).toBe(false);
  });
});
