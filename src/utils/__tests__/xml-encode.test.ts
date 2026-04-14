import { xmlEncode } from "@xml/encode";
import { describe, it, expect } from "vitest";

describe("xmlEncode", () => {
  it("encodes xml text", () => {
    expect(xmlEncode("<")).toBe("&lt;");
    expect(xmlEncode(">")).toBe("&gt;");
    expect(xmlEncode("&")).toBe("&amp;");
    expect(xmlEncode('"')).toBe("&quot;");
    expect(xmlEncode("'")).toBe("&apos;");

    expect(
      xmlEncode(
        "abc\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0b\x0e\x0f\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1a\x1b\x1c\x1d\x1e\x1f\x20abc\x7f"
      )
    ).toBe("abc abc");

    expect(xmlEncode('<a href="www.whatever.com">Talk to the H&</a>')).toBe(
      "&lt;a href=&quot;www.whatever.com&quot;&gt;Talk to the H&amp;&lt;/a&gt;"
    );

    // Should preserve newlines
    expect(xmlEncode("new\x0aline")).toBe("new\x0aline");
  });
});
