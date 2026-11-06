import { describe, expect, it } from "vitest";

import { DocxArgumentError, Document, Io } from "../index";

describe("Document Node surface", () => {
  it("adds the file-path members to the portable ones", () => {
    expect(Document.readFile).toBeTypeOf("function");
    expect(Document.writeFile).toBeTypeOf("function");
    expect(Document.read).toBeTypeOf("function");
  });

  it("Document.writeFile refuses a value that is not its handle", async () => {
    for (const bad of [42, null, {}, "doc"]) {
      const error = await Document.writeFile(bad as never, "unused.docx").catch((e: unknown) => e);
      expect(error, JSON.stringify(bad)).toBeInstanceOf(DocxArgumentError);
      expect((error as DocxArgumentError).field).toBe("doc");
      expect((error as Error).message).toMatch(/^Document\.writeFile: /);
    }
  });

  it("keeps the deprecated Io file-path members as forwards", () => {
    expect(Io.readFile).toBe(Document.readFile);
    expect(Io.writeFile).toBe(Document.writeFile);
  });
});
