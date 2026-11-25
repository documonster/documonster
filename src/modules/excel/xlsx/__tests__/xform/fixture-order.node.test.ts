import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";

import { CT } from "@excel/utils/ooxml-validator/check-content-types";
import { SEQUENCE_PARTS } from "@excel/utils/ooxml-validator/check-sequences";
import { OOXML_SEQUENCES } from "@excel/utils/ooxml-validator/ooxml-sequences";
import { parseXml } from "@xml/dom";
import { checkSequenceOrder, sequenceRanks } from "@xml/sequence-order";
import { describe, expect, it } from "vitest";

/**
 * The expected XML the xform tests compare against is itself in schema order.
 *
 * Those tests now compare child order (`test-xform-helper.ts`), so a fixture is what the writer is held to — and a
 * fixture edited to match a writer's wrong order would make the writer pass. Each fixture is checked here by the
 * validator's own content models; the worksheet's root by `CT_Worksheet`, which `check-worksheet.ts` ranks.
 */

const DATA = "src/modules/excel/xlsx/__tests__/xform";

/** A fixture's part kind, by its root element. */
const BY_ROOT: Readonly<Record<string, string>> = {
  worksheet: CT.worksheet,
  styleSheet: CT.styles,
  sst: CT.sharedStrings,
  wsDr: CT.drawing,
  comments: CT.comments,
  table: CT.tableDef,
  externalLink: CT.externalLink
};

const RANKS = sequenceRanks(OOXML_SEQUENCES);

/** Every fixture whose part the validator ranks, with its root's local name. */
const FIXTURES = readdirSync(DATA, { withFileTypes: true })
  .filter(entry => entry.isDirectory() && existsSync(join(DATA, entry.name, "data")))
  .flatMap(entry => {
    const dir = join(DATA, entry.name, "data");
    return readdirSync(dir)
      .filter(name => name.endsWith(".xml"))
      .map(name => join(dir, name));
  })
  .map(file => {
    const root = parseXml(readFileSync(file, "utf8")).root;
    return { file, root, local: root.name.slice(root.name.lastIndexOf(":") + 1) };
  })
  .filter(({ local }) => BY_ROOT[local] !== undefined);

describe("xform fixtures", () => {
  it("finds the fixtures of the parts the validator ranks", () => {
    expect(FIXTURES.length).toBeGreaterThanOrEqual(10);
  });

  for (const { file, root, local } of FIXTURES) {
    it(`${file} is in schema order`, () => {
      const part = SEQUENCE_PARTS.get(BY_ROOT[local]!)!;
      const tree = local === "worksheet" ? { ...part, type: "x:CT_Worksheet" as const } : part;
      const found: string[] = [];
      checkSequenceOrder(root, tree, RANKS, ({ trail, child, after }) =>
        found.push(`<${trail.join("/")}> <${child}> after <${after}>`)
      );
      expect(found).toEqual([]);
    });
  }
});
