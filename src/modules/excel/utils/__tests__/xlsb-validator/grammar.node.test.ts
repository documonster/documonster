import { readdirSync, readFileSync } from "node:fs";

import { validateXlsbBuffer, validateXlsbPart } from "@excel/utils/xlsb-validator";
import {
  COMMENTS_GRAMMAR,
  COMMON_OPENERS,
  grammarScopes,
  STYLESHEET_GRAMMAR,
  TABLE_GRAMMAR,
  WORKBOOK_GRAMMAR,
  WORKSHEET_GRAMMAR
} from "@excel/utils/xlsb-validator/grammar";
import { MS_XLSB_RECORD_NAMES } from "@excel/xlsb/spec/record-names";
import { biff } from "@test/biff-fixture";
import { describe, expect, it } from "vitest";

/**
 * The top-level record grammars (`grammar.ts`): every record they name is one the dictionary knows, they place
 * the records this library writes where MS-XLSB does, a part out of order is reported, and the Excel-authored
 * workbooks committed here are not. The records below carry no payload: only the order is under test, so only
 * the grammar's messages are read.
 */

const GRAMMARS = [
  WORKSHEET_GRAMMAR,
  WORKBOOK_GRAMMAR,
  STYLESHEET_GRAMMAR,
  TABLE_GRAMMAR,
  COMMENTS_GRAMMAR
];
const KNOWN = new Set(MS_XLSB_RECORD_NAMES.values());

/** Records the grammar names that the record table does not list, so no file can carry them under that name. */
const ABSENT_FROM_DICTIONARY = new Set(["BrtRevisionPtr", "brtBeginModelTimeGroupings"]);

describe("XLSB top-level grammars", () => {
  for (const grammar of GRAMMARS) {
    it(`${grammar.root} names only records the dictionary knows, in every scope`, () => {
      const scopes = grammarScopes(grammar);
      const names = [...scopes].flatMap(([opener, ranks]) => [opener, ...ranks.keys()]);
      expect(names.filter(name => !KNOWN.has(name) && !ABSENT_FROM_DICTIONARY.has(name))).toEqual(
        []
      );
    });
  }

  it("opens every common production with a record the dictionary knows", () => {
    expect(Object.values(COMMON_OPENERS).filter(name => !KNOWN.has(name))).toEqual([]);
  });

  it("ranks every part's root scope", () => {
    for (const [grammar, opener] of [
      [WORKSHEET_GRAMMAR, "BrtBeginSheet"],
      [WORKBOOK_GRAMMAR, "BrtBeginBook"],
      [STYLESHEET_GRAMMAR, "BrtBeginStyleSheet"],
      [TABLE_GRAMMAR, "BrtBeginList"],
      [COMMENTS_GRAMMAR, "BrtBeginComments"]
    ] as const) {
      const root = grammarScopes(grammar).get(opener);
      expect(root, grammar.root).toBeDefined();
      // The scope's own frame is not one of its positions.
      expect(root!.has(opener), grammar.root).toBe(false);
    }
  });

  it("places the records the writer emits where MS-XLSB does", () => {
    const sheet = grammarScopes(WORKSHEET_GRAMMAR).get("BrtBeginSheet")!;
    expect(sheet.get("BrtBeginAFilter")!).toBeLessThan(sheet.get("BrtBeginMergeCells")!);
    expect(sheet.get("BrtBeginListParts")!).toBeLessThan(sheet.get("BrtBeginSparklineGroups")!);
    const book = grammarScopes(WORKBOOK_GRAMMAR).get("BrtBeginBook")!;
    expect(book.get("BrtBeginExternals")!).toBeLessThan(book.get("BrtName")!);
    expect(book.get("BrtBeginPivotCacheIDs")!).toBeLessThan(book.get("BrtFileRecover")!);
  });
});

describe("XLSB record order check", () => {
  const grammarMessages = (names: readonly string[]) =>
    validateXlsbPart(biff(names.map(name => [name] as const)), "xl/worksheets/sheet1.bin")
      .problems.map(p => p.message)
      .filter(message => message.includes("record grammar"));

  it("reports merges written ahead of the auto filter", () => {
    expect(
      grammarMessages([
        "BrtBeginSheet",
        "BrtBeginMergeCells",
        "BrtEndMergeCells",
        "BrtBeginAFilter",
        "BrtEndAFilter",
        "BrtEndSheet"
      ])
    ).toEqual([
      "BrtBeginAFilter appears after BrtBeginMergeCells; the record grammar (WORKSHEET) places it earlier"
    ]);
  });

  it("names a collection behind BrtFRTBegin by the record inside it", () => {
    expect(
      grammarMessages([
        "BrtBeginSheet",
        "BrtFRTBegin",
        "BrtBeginSparklineGroups",
        "BrtEndSparklineGroups",
        "BrtFRTEnd",
        "BrtBeginMergeCells",
        "BrtEndMergeCells",
        "BrtEndSheet"
      ])
    ).toEqual([
      "BrtBeginMergeCells appears after BrtBeginSparklineGroups; the record grammar (WORKSHEET) places it earlier"
    ]);
  });

  it("does not rank the records inside a collection", () => {
    // `BrtWsProp` comes first at the top level; inside a scope it is a level down and not this check's.
    expect(
      grammarMessages([
        "BrtBeginSheet",
        "BrtBeginMergeCells",
        "BrtWsProp",
        "BrtEndMergeCells",
        "BrtEndSheet"
      ])
    ).toEqual([]);
  });

  it("ranks the records inside a scope by its production: a view's pane before its selection", () => {
    expect(
      grammarMessages([
        "BrtBeginSheet",
        "BrtBeginWsViews",
        "BrtBeginWsView",
        "BrtSel",
        "BrtPane",
        "BrtEndWsView",
        "BrtEndWsViews",
        "BrtEndSheet"
      ])
    ).toEqual(["BrtPane appears after BrtSel; the record grammar (WORKSHEET) places it earlier"]);
  });

  it("keeps each scope's ranking to itself", () => {
    // Two views, each pane-then-selection: the second view's pane is not "after" the first view's selection.
    expect(
      grammarMessages([
        "BrtBeginSheet",
        "BrtBeginWsViews",
        "BrtBeginWsView",
        "BrtPane",
        "BrtSel",
        "BrtEndWsView",
        "BrtBeginWsView",
        "BrtPane",
        "BrtSel",
        "BrtEndWsView",
        "BrtEndWsViews",
        "BrtEndSheet"
      ])
    ).toEqual([]);
  });

  it("lets a part open with a wrapper block before its root", () => {
    // MS-XLSB §2.1.7.51: `TABLE = [ACUID] BrtBeginList …`.
    const problems = validateXlsbPart(
      biff(
        [["BrtACBegin"], ["BrtACEnd"], ["BrtBeginList"], ["BrtEndList"]].map(r => r as [string])
      ),
      "xl/tables/table1.bin"
    ).problems;
    expect(problems.filter(p => p.kind === "scope-missing-root")).toEqual([]);
  });

  it("accepts the grammar's order", () => {
    expect(
      grammarMessages([
        "BrtBeginSheet",
        "BrtWsProp",
        "BrtWsDim",
        "BrtBeginSheetData",
        "BrtEndSheetData",
        "BrtBeginAFilter",
        "BrtEndAFilter",
        "BrtBeginMergeCells",
        "BrtEndMergeCells",
        "BrtFRTBegin",
        "BrtBeginSparklineGroups",
        "BrtEndSparklineGroups",
        "BrtFRTEnd",
        "BrtEndSheet"
      ])
    ).toEqual([]);
  });
});

/**
 * Every committed XLSB — Excel's own files and the fixture `scripts/generate-themed-xlsb.ts` builds — passes the
 * whole validator, not only its order check. Holding them to less let two defects sit unseen: the generated
 * fixture's style XF named a parent (Excel answers by removing every named style), and the root-scope check
 * refused Excel's own table parts, which open with an `ACUID` block before `BrtBeginList`.
 */
describe("committed XLSB files", () => {
  const FIXTURES = [
    "src/modules/excel/xlsb/__tests__/data",
    "src/modules/excel/examples/data"
  ].flatMap(dir =>
    readdirSync(dir)
      .filter(f => f.endsWith(".xlsb"))
      .map(f => `${dir}/${f}`)
  );

  it("finds the committed fixtures", () => {
    expect(FIXTURES.length).toBeGreaterThan(5);
  });

  for (const file of FIXTURES) {
    it(`${file} validates cleanly`, async () => {
      const report = await validateXlsbBuffer(new Uint8Array(readFileSync(file)));
      expect(report.problems).toEqual([]);
    });
  }
});

describe("XLSB record order in the table and comments parts", () => {
  const messages = (path: string, names: readonly string[]) =>
    validateXlsbPart(biff(names.map(name => [name] as const)), path)
      .problems.map(p => p.message)
      .filter(message => message.includes("record grammar"));

  it("reports a table's columns ahead of its auto filter", () => {
    expect(
      messages("xl/tables/table1.bin", [
        "BrtBeginList",
        "BrtBeginListCols",
        "BrtEndListCols",
        "BrtBeginAFilter",
        "BrtEndAFilter",
        "BrtEndList"
      ])
    ).toEqual([
      "BrtBeginAFilter appears after BrtBeginListCols; the record grammar (TABLE) places it earlier"
    ]);
  });

  it("reports comments ahead of their authors", () => {
    expect(
      messages("xl/comments1.bin", [
        "BrtBeginComments",
        "BrtBeginCommentList",
        "BrtEndCommentList",
        "BrtBeginCommentAuthors",
        "BrtEndCommentAuthors",
        "BrtEndComments"
      ])
    ).toEqual([
      "BrtBeginCommentAuthors appears after BrtBeginCommentList; the record grammar (COMMENTS) places it earlier"
    ]);
  });
});

describe("grammarScopes", () => {
  const grammar = (productions: Record<string, string>) => ({
    source: "test",
    root: "ROOT",
    productions: { ROOT: "BrtBeginSheet BrtWsProp BrtWsDim BrtEndSheet", ...productions }
  });

  it("throws on a production name that resolves nowhere — a typo in a copied production — in any scope", () => {
    expect(() =>
      grammarScopes(grammar({ ROOT: "BrtBeginSheet BrtWsProp WSFMTINF0 BrtEndSheet" }))
    ).toThrow(/WSFMTINF0/);
    expect(() =>
      grammarScopes(grammar({ A: "BrtBeginMergeCells BrtWsProp WSFMTINF0 BrtEndMergeCells" }))
    ).toThrow(/WSFMTINF0/);
  });

  it("throws on two productions opening one scope with different content, either way round", () => {
    const one = "BrtBeginMergeCells BrtMergeCell BrtWsProp BrtEndMergeCells";
    const more = "BrtBeginMergeCells BrtMergeCell BrtWsProp BrtWsDim BrtEndMergeCells";
    expect(() => grammarScopes(grammar({ A: one, B: more }))).toThrow(/BrtBeginMergeCells/);
    expect(() => grammarScopes(grammar({ A: more, B: one }))).toThrow(/BrtBeginMergeCells/);
    expect(() => grammarScopes(grammar({ A: one, B: one }))).not.toThrow();
  });

  it("ranks a scope whose unresolved name sits inside an inline scope it never looks into", () => {
    const scopes = grammarScopes(
      grammar({
        A: "BrtBeginMergeCells BrtWsProp BrtBeginSparklines WSFMTINF0 BrtEndSparklines BrtWsDim BrtEndMergeCells"
      })
    );
    expect([...scopes.get("BrtBeginMergeCells")!]).toEqual([
      ["BrtWsProp", 0],
      ["BrtBeginSparklines", 1],
      ["BrtWsDim", 2]
    ]);
  });

  it("skips a scope that needs an unpublished production, and ranks the one wrapped in BrtFRTBegin", () => {
    const scopes = grammarScopes(
      grammar({
        A: "BrtBeginMergeCells BrtWsProp PIVOTRULE BrtEndMergeCells",
        B: "BrtFRTBegin BrtBeginSparklineGroups BrtWsProp BrtWsDim BrtEndSparklineGroups BrtFRTEnd"
      })
    );
    expect(scopes.has("BrtBeginMergeCells")).toBe(false);
    expect([...scopes.get("BrtBeginSparklineGroups")!.keys()]).toEqual(["BrtWsProp", "BrtWsDim"]);
  });
});
