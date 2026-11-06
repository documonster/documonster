/**
 * `settings.xml` and the theme: what a Word-authored part says survives read → write.
 *
 * - `w:zoom` keeps its `w:val` (bestFit, fullPage, textFit, none) beside `w:percent`.
 * - `w:footnotePr` / `w:endnotePr` in settings keep their references to the
 *   separator notes, rewritten to the ids those notes are written with.
 * - The theme keeps `a:sysClr`, the font typefaces' other attributes and every
 *   element and attribute value of a theme laid out as Word writes one.
 */

import { readFileSync } from "node:fs";
import { join } from "node:path";

import { unzip } from "@archive/read-archive";
import { ZipArchive } from "@archive/zip/zip-archive";
import { Io } from "@word/index";
import { parseXml } from "@xml/dom";
import type { XmlElement } from "@xml/types";
import { describe, expect, it } from "vitest";

const DATA = join(__dirname, "data", "word-native");
const BASE = readFileSync(join(DATA, "toc-word.docx"));

async function readParts(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const parts = new Map<string, Uint8Array>();
  for await (const entry of unzip(bytes).entries()) {
    parts.set(entry.path, await entry.bytes());
  }
  return parts;
}

/** The base document with some parts replaced or added. */
async function withParts(replace: Record<string, string>): Promise<Uint8Array> {
  const parts = await readParts(BASE);
  for (const [path, text] of Object.entries(replace)) {
    parts.set(path, new TextEncoder().encode(text));
  }
  const archive = new ZipArchive();
  for (const [path, data] of parts) {
    archive.add(path, data);
  }
  return archive.toBuffer();
}

async function roundTripPart(input: Uint8Array, path: string): Promise<string> {
  const out = await readParts(await Io.toBuffer(await Io.read(input)));
  return new TextDecoder().decode(out.get(path)!);
}

async function partText(bytes: Uint8Array, path: string): Promise<string> {
  return new TextDecoder().decode((await readParts(bytes)).get(path)!);
}

function setSettingsChild(settings: string, find: RegExp, replacement: string): string {
  expect(settings).toMatch(find);
  return settings.replace(find, replacement);
}

/** Every element path with its attributes (namespace declarations aside), in document order. */
function describeTree(xml: string): string[] {
  const out: string[] = [];
  const walk = (el: XmlElement, path: string): void => {
    const attrs = Object.entries(el.attributes)
      .filter(([k]) => k !== "xmlns" && !k.startsWith("xmlns:"))
      .map(([k, v]) => `${k}=${v}`)
      .sort()
      .join(" ");
    const key = `${path}/${el.name}`;
    out.push(`${key} ${attrs}`);
    for (const child of el.children) {
      if (child.type === "element") {
        walk(child, key);
      }
    }
  };
  walk(parseXml(xml).root, "");
  return out;
}

describe("w:zoom", () => {
  it.each(["bestFit", "fullPage", "textFit", "none"])(
    "keeps w:val=%s and the percent",
    async val => {
      const settings = await partText(BASE, "word/settings.xml");
      const input = await withParts({
        "word/settings.xml": setSettingsChild(
          settings,
          /<w:zoom [^>]*\/>/,
          `<w:zoom w:val="${val}" w:percent="120"/>`
        )
      });
      const written = await roundTripPart(input, "word/settings.xml");
      const zoom = /<w:zoom [^>]*\/>/.exec(written)?.[0] ?? "";
      expect(zoom).toContain(`w:val="${val}"`);
      expect(zoom).toContain('w:percent="120"');
    }
  );
});

describe("separator references in settings", () => {
  /** Notes whose separators use ids 1 and 2 — not the -1 and 0 the writer gives them. */
  const NOTES = (kind: "footnote" | "endnote"): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:${kind}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:${kind} w:type="separator" w:id="1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
    `<w:${kind} w:type="continuationSeparator" w:id="2"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>` +
    `<w:${kind} w:id="3"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:${kind}>` +
    `</w:${kind}s>`;

  async function input(footnotePr: string, endnotePr: string): Promise<Uint8Array> {
    const settings = (await partText(BASE, "word/settings.xml")).replace(
      /<w:compat>/,
      `${footnotePr}${endnotePr}<w:compat>`
    );
    const rels = (await partText(BASE, "word/_rels/document.xml.rels")).replace(
      "</Relationships>",
      '<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>' +
        '<Relationship Id="rIdEn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>'
    );
    const types = (await partText(BASE, "[Content_Types].xml")).replace(
      "</Types>",
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' +
        '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>'
    );
    const document = (await partText(BASE, "word/document.xml")).replace(
      /<w:p\b/,
      '<w:p><w:r><w:footnoteReference w:id="3"/></w:r><w:r><w:endnoteReference w:id="3"/></w:r></w:p><w:p'
    );
    return withParts({
      "word/settings.xml": settings,
      "word/_rels/document.xml.rels": rels,
      "[Content_Types].xml": types,
      "word/document.xml": document,
      "word/footnotes.xml": NOTES("footnote"),
      "word/endnotes.xml": NOTES("endnote")
    });
  }

  /** The ids `kind` notes of a type are written with. */
  function writtenIds(notesXml: string, kind: string, type: string): string[] {
    return [...notesXml.matchAll(new RegExp(`<w:${kind} ([^>]*)>`, "g"))]
      .filter(m => m[1].includes(`w:type="${type}"`))
      .map(m => /w:id="(-?\d+)"/.exec(m[1])![1]);
  }

  it("keeps them, pointing at the separators as written, with the other properties", async () => {
    const bytes = await input(
      '<w:footnotePr><w:pos w:val="beneathText"/><w:numFmt w:val="lowerRoman"/><w:footnote w:id="1"/><w:footnote w:id="2"/></w:footnotePr>',
      '<w:endnotePr><w:footnote w:id="1"/></w:endnotePr>'.replace(/footnote/g, "endnote")
    );
    const out = await readParts(await Io.toBuffer(await Io.read(bytes)));
    const text = (p: string): string => new TextDecoder().decode(out.get(p)!);
    const settings = text("word/settings.xml");
    const footnotePr = /<w:footnotePr>.*?<\/w:footnotePr>/.exec(settings)?.[0] ?? "";
    const endnotePr = /<w:endnotePr>.*?<\/w:endnotePr>/.exec(settings)?.[0] ?? "";

    const fnSep = writtenIds(text("word/footnotes.xml"), "footnote", "separator");
    const fnCont = writtenIds(text("word/footnotes.xml"), "footnote", "continuationSeparator");
    const enSep = writtenIds(text("word/endnotes.xml"), "endnote", "separator");
    expect(fnSep).toHaveLength(1);
    expect(fnCont).toHaveLength(1);
    expect(footnotePr).toBe(
      `<w:footnotePr><w:pos w:val="beneathText"/><w:numFmt w:val="lowerRoman"/>` +
        `<w:footnote w:id="${fnSep[0]}"/><w:footnote w:id="${fnCont[0]}"/></w:footnotePr>`
    );
    expect(endnotePr).toBe(`<w:endnotePr><w:endnote w:id="${enSep[0]}"/></w:endnotePr>`);
  });

  it("writes no reference to a note that is not written", async () => {
    const settings = (await partText(BASE, "word/settings.xml")).replace(
      /<w:compat>/,
      '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:compat>'
    );
    const written = await roundTripPart(
      await withParts({ "word/settings.xml": settings }),
      "word/settings.xml"
    );
    expect(written).not.toContain("<w:footnote ");
  });
});

describe("a full theme1.xml (sysClr, face attributes, script fonts, fmtScheme, objectDefaults, extLst)", () => {
  const theme = readFileSync(join(__dirname, "data", "fixture-theme1.xml"), "utf8");

  it("survives read → write element for element, attribute for attribute", async () => {
    const written = await roundTripPart(
      await withParts({ "word/theme/theme1.xml": theme }),
      "word/theme/theme1.xml"
    );
    expect(describeTree(written)).toEqual(describeTree(theme));
  });

  it("writes an edited system colour as the colour it now is", async () => {
    const doc = await Io.read(await withParts({ "word/theme/theme1.xml": theme }));
    const colorScheme = doc.theme!.colorScheme;
    const edited = {
      ...doc,
      theme: {
        ...doc.theme!,
        colorScheme: { ...colorScheme, colors: { ...colorScheme.colors, dk1: "112233" } }
      }
    };
    const written = await partText(await Io.toBuffer(edited), "word/theme/theme1.xml");
    expect(/<a:dk1>.*?<\/a:dk1>/.exec(written)?.[0]).toBe(
      '<a:dk1><a:srgbClr val="112233"/></a:dk1>'
    );
    expect(/<a:lt1>.*?<\/a:lt1>/.exec(written)?.[0]).toBe(
      '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
    );
  });
});
