/**
 * A read → write of every fixture that once tripped the writer's own self-check. Each of these
 * produced a package the validator refused: a workbook relationship and a content-type override
 * naming `xl/theme/theme1.xml` with no such part (a source file that has no theme), or
 * `<externalReferences>` written after `<definedNames>` against the CT_Workbook sequence.
 */
import fs from "node:fs";

import { extractAll } from "@archive/unzip/extract";
import { expectValidXlsx } from "@excel/__tests__/helpers/expect-valid-xlsx";
import { requireEntryText } from "@excel/__tests__/helpers/zip-text";
import { Workbook } from "@excel/index";
import { makeTestDataPath } from "@test/utils";
import { describe, expect, it } from "vitest";

const dataPath = makeTestDataPath(import.meta.url, "../../__tests__/data");
const streamDataPath = makeTestDataPath(import.meta.url, "../../stream/__tests__/data");

const FIXTURES = [
  dataPath("data-validation-missing-type.xlsx"),
  dataPath("graceful-read-no-throw.xlsx"),
  dataPath("many-defined-names.xlsx"),
  dataPath("missing-cell-address.xlsx"),
  dataPath("unexpected-xml-node.xlsx"),
  streamDataPath("date-styles.xlsx")
];

describe("read → write of fixtures the self-check used to flag", () => {
  for (const file of FIXTURES) {
    it(`${file.split("/").pop()} writes a valid package`, async () => {
      const wb = Workbook.create();
      await Workbook.read(wb, fs.readFileSync(file));
      const out = await Workbook.toBuffer(wb, { validate: false });
      await expectValidXlsx(out, { label: file });
    });
  }

  it("a source with no workbook part can be written", async () => {
    // `missing-bits.xlsx` holds a sheet and the workbook rels, and nothing else.
    const wb = Workbook.create();
    await Workbook.read(wb, fs.readFileSync(streamDataPath("missing-bits.xlsx")));
    await expectValidXlsx(await Workbook.toBuffer(wb, { validate: false }));
  });

  it("a theme-less source gets the default theme its relationship names", async () => {
    const wb = Workbook.create();
    await Workbook.read(wb, fs.readFileSync(dataPath("missing-cell-address.xlsx")));
    const entries = await extractAll(await Workbook.toBuffer(wb, { validate: false }));
    expect(requireEntryText(entries, "xl/theme/theme1.xml")).toContain("<a:theme");
    expect(requireEntryText(entries, "xl/_rels/workbook.xml.rels")).toContain(
      'Target="theme/theme1.xml"'
    );
  });

  it("writes <externalReferences> before <definedNames>", async () => {
    const wb = Workbook.create();
    await Workbook.read(wb, fs.readFileSync(dataPath("many-defined-names.xlsx")));
    const entries = await extractAll(await Workbook.toBuffer(wb, { validate: false }));
    const xml = requireEntryText(entries, "xl/workbook.xml");
    expect(xml.indexOf("<externalReferences")).toBeGreaterThan(-1);
    expect(xml.indexOf("<externalReferences")).toBeLessThan(xml.indexOf("<definedNames"));
  });
});
