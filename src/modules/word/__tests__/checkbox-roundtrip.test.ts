/**
 * A `Build.checkbox` block reads back as the `CheckBox` it was, so its state
 * can be read and changed after `Io.read`.
 */

import { describe, it, expect } from "vitest";

import { Build, Document, Io, Package } from "../index";
import type { CheckBox } from "../types";

describe("CheckBox round trip", () => {
  it("reads back the block it was written as", async () => {
    const doc = Document.create();
    Document.addContent(doc, Build.checkbox({ checked: true }));
    Document.addContent(
      doc,
      Build.checkbox({ checked: false, checkedState: { value: "2714", font: "Segoe UI Symbol" } })
    );
    const read = await Io.read(await Io.toBuffer(doc));
    expect(read.body[0]).toMatchObject({ type: "checkBox", checked: true });
    expect(read.body[1]).toMatchObject({
      type: "checkBox",
      checked: false,
      checkedState: { value: "2714", font: "Segoe UI Symbol" }
    });

    const toggled = {
      ...read,
      body: [{ ...(read.body[0] as CheckBox), checked: false }, read.body[1]!]
    };
    const again = await Io.read(await Io.toBuffer(toggled));
    expect(again.body[0]).toMatchObject({ type: "checkBox", checked: false });
    const xml = new TextDecoder().decode(
      (await Package.readPart(await Io.toBuffer(again), "word/document.xml"))!
    );
    expect(xml.match(/<w14:checkbox[ >]/g)).toHaveLength(2);
  });
});
