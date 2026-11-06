/**
 * `documonster/word/crypto` exports the errors its functions throw, so a
 * caller of that entry alone can tell them apart.
 */

import { describe, it, expect } from "vitest";

import { DocxArgumentError, DocxDecryptionError, WordCrypto, isDocxError } from "../crypto";
import { Document, Io } from "../index";

describe("documonster/word/crypto errors", () => {
  it("exports what decrypt and encrypt throw, and the predicate", async () => {
    const d = Document.create();
    Document.addParagraph(d, "x");
    const enc = await WordCrypto.encrypt(await Io.toBuffer(d), "pw", { spinCount: 1 });
    const wrong = await WordCrypto.decrypt(enc, "nope").catch((e: unknown) => e);
    expect(wrong).toBeInstanceOf(DocxDecryptionError);
    expect(isDocxError(wrong)).toBe(true);
    const bad = await WordCrypto.encrypt(new Uint8Array(4), 42 as never).catch((e: unknown) => e);
    expect(bad).toBeInstanceOf(DocxArgumentError);
  });
});
