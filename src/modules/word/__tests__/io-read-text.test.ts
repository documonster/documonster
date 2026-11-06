/**
 * `Io.read` reads back what `Io.toBase64` and `Io.toFlatOpc` write: base64
 * text with `{ base64: true }` (as `Workbook.read` takes it), and Flat OPC
 * XML as a string. Before, both strings were refused as "not a Uint8Array",
 * so the round trip needed `TextEncoder` or a base64 decoder of the caller's.
 */

import { describe, it, expect } from "vitest";

import { Document, DocxArgumentError, DocxParseError, Io, Query } from "../index";

function doc() {
  const d = Document.create();
  Document.addParagraph(d, "round trip");
  return d;
}

describe("Io.read of text", () => {
  it("reads Io.toBase64 output with { base64: true }", async () => {
    const read = await Io.read(await Io.toBase64(doc()), { base64: true });
    expect(Query.getText(read)).toContain("round trip");
  });

  it("reads Io.toFlatOpc output as a string", async () => {
    const read = await Io.read(await Io.toFlatOpc(doc()));
    expect(Query.getText(read)).toContain("round trip");
  });

  it("reads a Flat OPC string with a BOM or leading whitespace, as it reads the bytes", async () => {
    const flat = await Io.toFlatOpc(doc());
    expect(Query.getText(await Io.read(`\uFEFF${flat}`))).toContain("round trip");
    expect(Query.getText(await Io.read(`\n  ${flat}`))).toContain("round trip");
  });

  it("refuses other text, saying what a string may be", async () => {
    const base64 = await Io.toBase64(doc());
    const error = await Io.read(base64).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(DocxArgumentError);
    expect((error as DocxArgumentError).field).toBe("input");
    expect((error as Error).message).toMatch(/base64: true/);
    await expect(Io.read("not base64 !!", { base64: true })).rejects.toBeInstanceOf(
      DocxArgumentError
    );
  });

  it("checks the base64 option", async () => {
    const error = await Io.read("", { base64: "yes" as never }).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(DocxArgumentError);
    expect((error as DocxArgumentError).field).toBe("options.base64");
  });

  it("reads bytes as before when base64 is set (it applies to text)", async () => {
    const bytes = await Io.toBuffer(doc());
    expect(Query.getText(await Io.read(bytes, { base64: true }))).toContain("round trip");
  });

  it("reports base64 that is not a package as a parse error", async () => {
    await expect(Io.read(btoa("hello"), { base64: true })).rejects.toBeInstanceOf(DocxParseError);
  });
});
