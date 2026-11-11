/**
 * `Package.edit` types a part it adds in `[Content_Types].xml`: inside the root element wherever its
 * end is — a self-closing root, a prefixed one, one followed by a comment that spells an end tag — in
 * the root's namespace, the rest of the part kept as it was.
 */
import { Archive } from "@archive/index";
import { describe, expect, it } from "vitest";

import { DocxParseError, Package } from "../index";

const CT = "http://schemas.openxmlformats.org/package/2006/content-types";

async function typesAfterAdding(contentTypes: string): Promise<string> {
  const bytes = await Archive.packZip([["[Content_Types].xml", contentTypes]]);
  const edited = await Package.edit(bytes, [
    { type: "replacePartText", path: "new.dat", text: "hello" }
  ]);
  return new TextDecoder().decode(await Package.readPart(edited, "[Content_Types].xml"));
}

const DEFAULT = 'Default Extension="dat" ContentType="application/octet-stream"/>';

describe("Package.edit adds a part's content type into [Content_Types].xml", () => {
  it.each([
    [`<Types xmlns="${CT}"/>`, `<Types xmlns="${CT}"><${DEFAULT}</Types>`],
    [
      `<ct:Types xmlns:ct="${CT}"></ct:Types>`,
      `<ct:Types xmlns:ct="${CT}"><ct:${DEFAULT}</ct:Types>`
    ],
    [
      `<Types xmlns="${CT}"></Types><!-- </fake> -->`,
      `<Types xmlns="${CT}"><${DEFAULT}</Types><!-- </fake> -->`
    ],
    [
      ` \n<?xml version="1.0"?>\n<Types xmlns="${CT}">\n  <Default Extension="xml" ContentType="x"/>\n</Types>\n`,
      ` \n<?xml version="1.0"?>\n<Types xmlns="${CT}">\n  <Default Extension="xml" ContentType="x"/>\n<${DEFAULT}</Types>\n`
    ]
  ])("%j", async (contentTypes, expected) => {
    expect(await typesAfterAdding(contentTypes)).toBe(expected);
  });

  it("adds nothing for an extension already typed, whatever its prefix", async () => {
    const types = `<ct:Types xmlns:ct="${CT}"><ct:Default Extension="DAT" ContentType="x"/></ct:Types>`;
    expect(await typesAfterAdding(types)).toBe(types);
  });

  it("reports a part it cannot read as the package's error", async () => {
    for (const types of ["<Types>", `<Types xmlns="urn:other"/>`]) {
      await expect(typesAfterAdding(types)).rejects.toBeInstanceOf(DocxParseError);
    }
  });
});
