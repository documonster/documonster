/**
 * An external link keeps what this library does not model: Excel 2021's alternate URLs (`xxl21:alternateUrls`), with
 * the relationships they name, and the declarations that make them readable.
 *
 * `many-defined-names.xlsx` links to a workbook on SharePoint. Its `<externalBook r:id="rId1">` names the
 * `file:///\\FAEDDB10\…` path, and `<xxl21:absoluteUrl r:id="rId3"/>` the `https://…` one; `rId2` is the
 * `externalLinkLongPath`. A write kept only one relationship — the first `externalLinkPath`, which is the https one —
 * and so also pointed the book at the wrong file.
 *
 * `.node` because it reads a fixture off disk.
 */
import { readFile } from "node:fs/promises";

import { extractAll } from "@archive/unzip/extract";
import { Workbook } from "@excel";
import { describe, expect, it } from "vitest";

const FIXTURE = "src/modules/excel/__tests__/data/many-defined-names.xlsx";
const decoder = new TextDecoder();

function relationships(xml: string): Map<string, { type: string; target: string }> {
  return new Map(
    [...xml.matchAll(/<Relationship ([^>]*)\/>/g)].map(([, attrs]) => {
      const attr = (name: string) => new RegExp(`${name}="([^"]*)"`).exec(attrs!)?.[1] ?? "";
      return [attr("Id"), { type: attr("Type"), target: attr("Target") }];
    })
  );
}

describe("external link: unmodelled content", () => {
  it("keeps alternateUrls, its relationships and the book's own path", async () => {
    const source = await extractAll(new Uint8Array(await readFile(FIXTURE)));
    const wb = await Workbook.read(Workbook.create(), new Uint8Array(await readFile(FIXTURE)));
    const out = await extractAll(await Workbook.toBuffer(wb));
    const text = (entries: typeof out, path: string) => decoder.decode(entries.get(path)!.data);

    const xml = text(out, "xl/externalLinks/externalLink1.xml");
    const rels = relationships(text(out, "xl/externalLinks/_rels/externalLink1.xml.rels"));
    const sourceRels = relationships(text(source, "xl/externalLinks/_rels/externalLink1.xml.rels"));

    expect(xml).toMatch(
      /<xxl21:alternateUrls[^>]*><xxl21:absoluteUrl r:id="rId3"\/><\/xxl21:alternateUrls>/
    );
    expect(xml).toMatch(
      /xmlns:xxl21="http:\/\/schemas\.microsoft\.com\/office\/spreadsheetml\/2021\/extlinks2021"/
    );
    expect(xml).toMatch(/mc:Ignorable="[^"]*\bxxl21\b/);
    // Before `sheetNames`, where the source has it.
    expect(xml.indexOf("alternateUrls")).toBeLessThan(xml.indexOf("<sheetNames>"));

    const book = /<externalBook[^>]* r:id="(rId\d+)"/.exec(xml)![1]!;
    expect(rels.get(book)).toEqual(sourceRels.get("rId1"));
    expect(rels.get("rId3")).toEqual(sourceRels.get("rId3"));
    expect(rels.get("rId2")).toEqual(sourceRels.get("rId2"));
    expect(Workbook.getModel(wb).externalLinks?.[0]?.target).toBe(sourceRels.get("rId1")!.target);
  });
});
