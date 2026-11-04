/**
 * `Convert.toOdt` finalizes the document like `Io.toBuffer` does, so a built
 * Chinese document states its East Asian language and face in ODF too.
 */
import { unzip } from "@archive/read-archive";
import { describe, it, expect } from "vitest";

import { Convert, Document } from "../../../index";

async function stylesXml(bytes: Uint8Array): Promise<string> {
  for await (const entry of unzip(bytes).entries()) {
    if (entry.path === "styles.xml") {
      return new TextDecoder().decode(await entry.bytes());
    }
  }
  throw new Error("no styles.xml");
}

describe("Convert.toOdt language", () => {
  it("derives the East Asian language and face of a built document", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, "数据校验完成");
    const xml = await stylesXml(await Convert.toOdt(doc));
    expect(xml).toMatch(/style:language-asian="zh"/);
    expect(xml).toMatch(/style:country-asian="CN"/);
    expect(xml).toMatch(/style:font-name-asian="[^"]+"/);
  });

  it("writes a stated language as ODF's separate attributes", async () => {
    const doc = Document.create();
    Document.setDocDefaults(doc, {
      runProperties: { language: { val: "en-US", eastAsia: "zh-Hant-TW", bidi: "ar-SA" } }
    });
    Document.addParagraph(doc, "x");
    const xml = await stylesXml(await Convert.toOdt(doc));
    expect(xml).toContain('fo:language="en"');
    expect(xml).toContain('fo:country="US"');
    expect(xml).toContain('style:language-asian="zh"');
    expect(xml).toContain('style:script-asian="Hant"');
    expect(xml).toContain('style:country-asian="TW"');
    expect(xml).toContain('style:language-complex="ar"');
    expect(xml).toContain('style:country-complex="SA"');
  });

  it("leaves a read document's language alone", async () => {
    const built = Document.create();
    Document.addParagraph(built, "数据校验完成");
    const read = await Convert.fromOdt(
      await Convert.toOdt({ ...built, deriveEastAsianDefaults: false })
    );
    expect(await stylesXml(await Convert.toOdt(read))).not.toMatch(/language-asian/);
  });
});
