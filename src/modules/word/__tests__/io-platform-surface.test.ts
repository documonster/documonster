import * as PortableDocument from "@word/surface/document-portable";
/**
 * The browser builds of the `Document` and the deprecated `Io` namespaces publish
 * everything the Node ones do except the file-path members, as Excel's and
 * Archive's browser entries publish no file-path API.
 */
import * as BrowserDocument from "@word/surface/document.browser";
import * as PortableIo from "@word/surface/io-portable";
import * as BrowserIo from "@word/surface/io.browser";
import { describe, expect, it } from "vitest";

describe("Document browser surface", () => {
  it("has no file-path members", () => {
    expect("readFile" in BrowserDocument).toBe(false);
    expect("writeFile" in BrowserDocument).toBe(false);
  });

  it("publishes every portable member", () => {
    expect(Object.keys(BrowserDocument).sort()).toEqual(Object.keys(PortableDocument).sort());
    expect(BrowserDocument.create).toBeTypeOf("function");
    expect(BrowserDocument.read).toBeTypeOf("function");
    expect(BrowserDocument.toBuffer).toBeTypeOf("function");
  });
});

describe("Io browser surface", () => {
  it("has no file-path members", () => {
    expect("readFile" in BrowserIo).toBe(false);
    expect("writeFile" in BrowserIo).toBe(false);
  });

  it("publishes every portable member", () => {
    expect(Object.keys(BrowserIo).sort()).toEqual(Object.keys(PortableIo).sort());
  });

  it("forwards each reading and writing member to Document", () => {
    expect(BrowserIo.read).toBe(BrowserDocument.read);
    expect(BrowserIo.toBuffer).toBe(BrowserDocument.toBuffer);
    expect(BrowserIo.toBase64).toBe(BrowserDocument.toBase64);
    expect(BrowserIo.toFlatOpc).toBe(BrowserDocument.toFlatOpc);
  });
});
