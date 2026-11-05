import * as PortableIo from "@word/surface/io-portable";
/**
 * The browser build of the `Io` namespace publishes everything the Node one does
 * except the file-path members, as Excel's and Archive's browser entries publish
 * no file-path API.
 */
import * as BrowserIo from "@word/surface/io.browser";
import { describe, expect, it } from "vitest";

describe("Io browser surface", () => {
  it("has no file-path members", () => {
    expect("readFile" in BrowserIo).toBe(false);
    expect("writeFile" in BrowserIo).toBe(false);
  });

  it("publishes every portable member", () => {
    expect(Object.keys(BrowserIo).sort()).toEqual(Object.keys(PortableIo).sort());
    expect(BrowserIo.read).toBeTypeOf("function");
    expect(BrowserIo.toBuffer).toBeTypeOf("function");
  });
});
