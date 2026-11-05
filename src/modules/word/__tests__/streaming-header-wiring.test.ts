/**
 * The streaming writer wires a first-page or even-page header as
 * `Document.setHeader` does: the section gets its title page, the document
 * its even/odd setting, without which Word shows neither header.
 */

import { describe, it, expect } from "vitest";

import { Build, Io, Layout, Stream } from "../index";
import type { HeaderDef } from "../index";

const header = (text: string): HeaderDef => ({
  content: { children: [Build.textParagraph(text)] }
});

describe("streaming writer header wiring", () => {
  it("turns the title page on for a first-page header", async () => {
    const w = new Stream.DocumentWriter({ headers: new Map([["first", header("FIRSTMARK")]]) });
    w.add(Build.textParagraph("body"));
    const read = await Io.read(await w.finalize());
    expect(read.sectionProperties?.titlePage).toBe(true);
    expect(Layout.toSvg(read)[0]).toContain("FIRSTMARK");
  });

  it("turns even/odd headers on for an even-page footer", async () => {
    const w = new Stream.DocumentWriter({ footers: new Map([["even", header("EVENMARK")]]) });
    w.add(Build.textParagraph("body"));
    const read = await Io.read(await w.finalize());
    expect(read.settings?.evenAndOddHeaders).toBe(true);
  });

  it("leaves a title page the caller turned off explicitly alone", async () => {
    const w = new Stream.DocumentWriter({
      headers: new Map([["first", header("FIRSTMARK")]]),
      sectionProperties: { titlePage: false }
    });
    w.add(Build.textParagraph("body"));
    const read = await Io.read(await w.finalize());
    expect(read.sectionProperties?.titlePage ?? false).toBe(false);
  });
});
