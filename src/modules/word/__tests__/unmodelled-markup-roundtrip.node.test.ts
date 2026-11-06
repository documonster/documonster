/**
 * Read → write keeps the markup of regenerated parts the model does not represent.
 *
 * `settings.xml`, `styles.xml`, the font table, the theme and `app.xml` are
 * written from the model, so whatever the model has no field for used to be
 * dropped: `w:latentStyles`, `m:mathPr`, `w:clrSchemeMapping` (which maps the
 * theme's colours onto the document's), `w14:docId`/`w15:docId`, the roots'
 * namespace declarations and `mc:Ignorable`. Every element of the source is
 * now kept — modelled, or verbatim at its schema position.
 */

import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";

import { unzip } from "@archive/read-archive";
import { DocxRawXmlPolicyError } from "@word/errors";
import { Io } from "@word/index";
import { parseXml } from "@xml/dom";
import type { XmlElement } from "@xml/types";
import { describe, expect, it } from "vitest";

const DATA = join(__dirname, "data");
const FIXTURES = readdirSync(DATA, { withFileTypes: true })
  // `office-encrypted` holds password-protected files: compound files, not packages to round-trip.
  .filter(d => d.isDirectory() && d.name !== "office-encrypted")
  .flatMap(d =>
    readdirSync(join(DATA, d.name))
      .filter(f => f.endsWith(".docx"))
      .map(f => join(d.name, f))
  );

const REGENERATED = [
  "word/settings.xml",
  "word/styles.xml",
  "word/fontTable.xml",
  "word/webSettings.xml",
  "word/numbering.xml",
  "word/theme/theme1.xml",
  "word/footnotes.xml",
  "word/endnotes.xml",
  "docProps/app.xml",
  "docProps/core.xml"
];

/**
 * Root attributes the writer drops on purpose: the font table and web settings
 * hold nothing in those namespaces once written, so their declarations — and an
 * `mc:Ignorable` naming them — would declare namespaces no element uses.
 */
const INTENTIONAL_ROOT_DROPS: Readonly<Record<string, RegExp>> = {
  "word/fontTable.xml": /^(xmlns:|mc:Ignorable$)/,
  "word/webSettings.xml": /^(xmlns:|mc:Ignorable$)/
};

async function xmlParts(bytes: Uint8Array): Promise<Map<string, string>> {
  const parts = new Map<string, string>();
  for await (const entry of unzip(bytes).entries()) {
    if (REGENERATED.includes(entry.path)) {
      parts.set(entry.path, new TextDecoder().decode(await entry.bytes()));
    }
  }
  return parts;
}

/** Every element path in the part, and the root's attribute names. */
function inventory(xml: string): Set<string> {
  const out = new Set<string>();
  const walk = (el: XmlElement, path: string): void => {
    for (const child of el.children) {
      if (child.type === "element") {
        const key = `${path}/${child.name}`;
        out.add(key);
        walk(child, key);
      }
    }
  };
  const root = parseXml(xml).root;
  walk(root, root.name);
  for (const name of Object.keys(root.attributes)) {
    out.add(`@${name}`);
  }
  return out;
}

async function roundTrip(file: string): Promise<{
  before: Map<string, string>;
  after: Map<string, string>;
}> {
  const source = readFileSync(join(DATA, file));
  const after = await Io.toBuffer(await Io.read(source));
  return { before: await xmlParts(source), after: await xmlParts(after) };
}

describe("read → write keeps unmodelled markup of regenerated parts", () => {
  it.each(FIXTURES)("%s", async file => {
    const { before, after } = await roundTrip(file);
    const dropped: string[] = [];
    for (const [part, xml] of before) {
      const written = after.get(part);
      expect(written, `${part} is written`).toBeDefined();
      const kept = inventory(written!);
      for (const key of inventory(xml)) {
        const intentional = INTENTIONAL_ROOT_DROPS[part];
        if (!kept.has(key) && !(key.startsWith("@") && intentional?.test(key.slice(1)))) {
          dropped.push(`${part} ${key}`);
        }
      }
    }
    expect(dropped).toEqual([]);
  });

  it("writes the kept settings in CT_Settings order, extensions last", async () => {
    const { after } = await roundTrip("word-native/toc-word.docx");
    const children = parseXml(after.get("word/settings.xml")!)
      .root.children.filter((c): c is XmlElement => c.type === "element")
      .map(c => c.name);
    expect(children).toEqual([
      "w:zoom",
      "w:defaultTabStop",
      "w:characterSpacingControl",
      "w:compat",
      "w:rsids",
      "m:mathPr",
      "w:themeFontLang",
      "w:clrSchemeMapping",
      "w:decimalSymbol",
      "w:listSeparator",
      "w14:docId",
      "w15:docId"
    ]);
    const root = parseXml(after.get("word/settings.xml")!).root;
    expect(root.attributes["mc:Ignorable"]).toContain("w15");
    expect(root.attributes["xmlns:w15"]).toBe(
      "http://schemas.microsoft.com/office/word/2012/wordml"
    );
  });

  it("puts w:latentStyles between w:docDefaults and the styles", async () => {
    const { after } = await roundTrip("word-native/toc-word.docx");
    const children = parseXml(after.get("word/styles.xml")!)
      .root.children.filter((c): c is XmlElement => c.type === "element")
      .map(c => c.name);
    expect(children.slice(0, 3)).toEqual(["w:docDefaults", "w:latentStyles", "w:style"]);
  });

  it("honours rawXmlPolicy for the kept markup", async () => {
    const doc = await Io.read(readFileSync(join(DATA, "word-native/toc-word.docx")));
    const stripped = await xmlParts(
      await Io.toBuffer(doc, { securityPolicy: { rawXmlPolicy: "strip" } })
    );
    expect(stripped.get("word/settings.xml")).not.toContain("m:mathPr");
    expect(stripped.get("word/styles.xml")).not.toContain("w:latentStyles");
    await expect(
      Io.toBuffer({ ...doc, theme: undefined }, { securityPolicy: { rawXmlPolicy: "reject" } })
    ).rejects.toThrow(DocxRawXmlPolicyError);
  });
});
