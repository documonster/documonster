import { describe, expect, it } from "vitest";

import { DocxArgumentError, Io } from "../index";

describe("Io Node surface", () => {
  it("adds the file-path members to the portable ones", () => {
    expect(Io.readFile).toBeTypeOf("function");
    expect(Io.writeFile).toBeTypeOf("function");
    expect(Io.read).toBeTypeOf("function");
  });

  it("Io.writeFile refuses a value that is not its handle", async () => {
    for (const bad of [42, null, {}, "doc"]) {
      const error = await Io.writeFile(bad as never, "unused.docx").catch((e: unknown) => e);
      expect(error, JSON.stringify(bad)).toBeInstanceOf(DocxArgumentError);
      expect((error as DocxArgumentError).field).toBe("doc");
      expect((error as Error).message).toMatch(/^Io\.writeFile: /);
    }
  });
});
