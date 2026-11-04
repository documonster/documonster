/**
 * The writer's validation of a document model is shared by every writing entry
 * point (`Io.toBuffer`, `Io.writeFile`, `Package.*`, the streaming writer), so it
 * cannot name its caller: its message states the problem without a prefix — no
 * `write:` and no internal name — and `field` is the model path (AGENTS.md, "An
 * argument error says where").
 */

import { describe, it, expect } from "vitest";

import { Build, Document, DocxArgumentError, Io, Streaming } from "../index";
import type { BodyContent, DocxDocument } from "../index";

async function writeError(doc: DocxDocument): Promise<DocxArgumentError> {
  const error = await Io.toBuffer(doc).catch((e: unknown) => e);
  expect(error).toBeInstanceOf(DocxArgumentError);
  return error as DocxArgumentError;
}

function withBlock(block: unknown): DocxDocument {
  return { ...Document.create(), body: [block as BodyContent] };
}

describe("model checks at write time", () => {
  it("an unknown node kind: no prefix, the model path as field", async () => {
    const error = await writeError(withBlock({ type: "bogus" }));
    expect(error.field).toBe("body[0].type");
    expect(error.message).toMatch(/^type must be /);
    expect(error.message).toContain("(at body[0].type)");
  });

  it("a value outside its vocabulary", async () => {
    const para = Build.paragraph([Build.text("x")]);
    const error = await writeError(
      withBlock({ ...para, properties: { alignment: "bogus" as never } })
    );
    expect(error.field).toBe("body[0].properties.alignment");
    expect(error.message).toMatch(/^alignment must be one of /);
    expect(error.message).not.toMatch(/write/i);
  });

  it("a drawing number with no XML form", async () => {
    const shape = Build.rect(1, 1);
    const error = await writeError(withBlock({ ...shape, width: Number.NaN }));
    expect(error.field).toMatch(/^body\[0\]\./);
    expect(error.message).toMatch(/^width must be a finite number, got NaN/);
  });

  it("the streaming writer reports the same model check the same way", () => {
    const writer = Streaming.createWriter();
    let error: unknown;
    try {
      writer.addText("x", { alignment: "bogus" as never });
    } catch (e) {
      error = e;
    }
    // `addText` checks its own argument and names itself …
    expect((error as Error).message).toMatch(/^Streaming\.StreamingDocxWriter#addText: /);
    // … while the model check behind `add` states the problem alone.
    const para = { ...Build.paragraph([]), properties: { alignment: "bogus" } };
    try {
      writer.add(para as never);
    } catch (e) {
      error = e;
    }
    expect((error as DocxArgumentError).field).toBe("element.properties.alignment");
    expect((error as Error).message).toMatch(/^alignment must be one of /);
  });
});
