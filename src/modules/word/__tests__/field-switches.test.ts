/**
 * Field switches — the general `\*`, `\#` and `\@` switches and MERGEFIELD's
 * own `\b` / `\f` — as mail merge and field update apply them, and fields that
 * depend on merged values (an `IF` testing a `MERGEFIELD`) evaluated again
 * once those values are in.
 */

import { describe, expect, it } from "vitest";

import { formatFieldResult } from "../core/field-format";
import { Document, Io, Query, Transform } from "../index";
import type { DocxDocument } from "../types";
import { buildDocx, packageParts, partText } from "./roundtrip-integrity.helpers";

const R = (t: string): string => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const CH = (type: string): string => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const BEGIN = CH("begin");
const SEP = CH("separate");
const END = CH("end");
const INSTR = (t: string): string =>
  `<w:r><w:instrText xml:space="preserve">${t.replace(/"/g, "&quot;")}</w:instrText></w:r>`;
const P = (inner: string): string => `<w:p>${inner}</w:p>`;
const SIMPLE = (instr: string, result: string): string =>
  `<w:fldSimple w:instr="${instr.replace(/"/g, "&quot;")}">${R(result)}</w:fldSimple>`;

async function read(body: string): Promise<DocxDocument> {
  return Io.read(await buildDocx({ body }));
}

function text(doc: DocxDocument): string {
  return Query.extractText(doc).trim();
}

describe("formatFieldResult — the general switches", () => {
  it.each([
    [" MERGEFIELD N \\* Upper ", "ann <b>", "ANN <B>"],
    [" MERGEFIELD N \\* Lower ", "Ann SMITH", "ann smith"],
    // Word's own results (Word for Mac 16): Caps lowercases the rest of each
    // word, and a hyphen starts a word where an apostrophe does not.
    [" MERGEFIELD N \\* Caps ", "ann smith-jones", "Ann Smith-Jones"],
    [' QUOTE "hELLO bIG wORLD" \\* Caps ', "hELLO bIG wORLD", "Hello Big World"],
    [' QUOTE "o\'neil mc-donald" \\* Caps ', "o'neil mc-donald", "O'neil Mc-Donald"],
    [" = 21 \\* CardText \\* Caps ", "21", "Twenty-One"],
    [" MERGEFIELD N \\* FirstCap ", "ann smith", "Ann smith"],
    [" MERGEFIELD N \\* MERGEFORMAT ", "ann", "ann"],
    [" SEQ F \\* roman ", "4", "iv"],
    [" SEQ F \\* ROMAN ", "14", "XIV"],
    [" SEQ F \\* alphabetic ", "28", "bb"],
    [" SEQ F \\* ALPHABETIC ", "3", "C"],
    [" SEQ F \\* Ordinal ", "22", "22nd"],
    [" SEQ F \\* Hex ", "255", "FF"],
    [' MERGEFIELD A \\# "#,##0.00" ', "1234.5", "1,234.50"],
    [' MERGEFIELD A \\# "0.00" ', "-3.456", "-3.46"],
    [' MERGEFIELD A \\# "$#,##0.00;($#,##0.00)" ', "-1234.5", "($1,234.50)"],
    [" MERGEFIELD A \\# 0 ", "2.5", "3"],
    [' MERGEFIELD A \\# "0%" ', "12", "12%"],
    [' MERGEFIELD A \\# "#,##0.00" ', "abc", "abc"],
    [' MERGEFIELD D \\@ "d MMMM yyyy" ', "2024-03-05", "5 March 2024"],
    [' MERGEFIELD D \\@ "dddd, MMM d, yy" ', "2024-03-05", "Tuesday, Mar 5, 24"],
    [' MERGEFIELD D \\@ "dd/MM/yyyy HH:mm" ', "2024-03-05T07:08", "05/03/2024 07:08"],
    [' MERGEFIELD D \\@ "h:mm am/pm" ', "2024-03-05T13:08", "1:08 pm"],
    [" MERGEFIELD D \\@ \"'Q' M\" ", "2024-11-05", "Q 11"],
    [' MERGEFIELD D \\@ "d MMMM yyyy" ', "not a date", "not a date"],
    [' MERGEFIELD A \\# "0.00" \\* Upper ', "1.5", "1.50"],
    // Word's own results (Word for Mac 16, fields updated in Word): an unused
    // `#` is a space, an unused grouping separator is nothing.
    [' = 5 \\# "#,##0.00" ', "5", "   5.00"],
    [" = 42 \\# $### ", "42", "$ 42"],
    [" = 1234.5 \\# $### ", "1234.5", "$1235"],
    [" = 1234 \\# #,##0 ", "1234", "1,234"],
    [' = 5 \\# "#,###,##0" ', "5", "      5"],
    [' = 12345 \\# "#,###,##0" ', "12345", "  12,345"],
    [" = 2 \\* CardText ", "2", "two"],
    [" = 12 \\* OrdText ", "12", "twelfth"],
    [" = 0 \\* CardText ", "0", "zero"],
    [" = 21 \\* CardText ", "21", "twenty-one"],
    [" = 790 \\* CardText ", "790", "seven hundred ninety"],
    [
      " = 123456 \\* CardText ",
      "123456",
      "one hundred twenty-three thousand four hundred fifty-six"
    ],
    [" = 21 \\* OrdText ", "21", "twenty-first"],
    [" = 40 \\* OrdText ", "40", "fortieth"],
    [" = 100 \\* OrdText ", "100", "one hundredth"],
    [" = 3 \\* CardText \\* Upper ", "3", "THREE"],
    [" = 999.45 \\* DollarText ", "999.45", "nine hundred ninety-nine and 45/100"],
    [" = 2 \\* DollarText ", "2", "two and 00/100"]
  ])("%s on %j", (instruction, value, expected) => {
    expect(formatFieldResult(value, instruction)).toBe(expected);
  });
});

describe("mail merge applies the field's switches", () => {
  it("formats case, numbers and text before/after", async () => {
    const doc = await read(
      P(
        R("Dear ") +
          BEGIN +
          INSTR(" MERGE") +
          INSTR("FIELD  Na") +
          INSTR("me \\* Upper ") +
          SEP +
          R("«Name»") +
          END +
          R(", total ") +
          BEGIN +
          INSTR(' MERGEFIELD Amount \\# "#,##0.00" ') +
          SEP +
          R("«Amount»") +
          END +
          R(" ") +
          SIMPLE(' MERGEFIELD City \\b "in " \\f "!" ', "«City»")
      )
    );
    expect(Query.mailMerge(doc, { Name: "ann <b>", Amount: "1234.5", City: "Oslo" })).toBe(3);
    expect(text(doc)).toBe("Dear ANN <B>, total 1,234.50 in Oslo!");
  });

  it("leaves out the \\b / \\f text when the value is empty", async () => {
    const doc = await read(P(R("x") + SIMPLE(' MERGEFIELD City \\b "in " \\f "!" ', "«City»")));
    Query.mailMerge(doc, { City: "" });
    expect(text(doc)).toBe("x");
  });

  it("evaluates an IF again with the values merged into its condition", async () => {
    const ifField = (vip: string): string =>
      BEGIN +
      INSTR(" IF ") +
      BEGIN +
      INSTR(` MERGEFIELD ${vip} `) +
      SEP +
      R("x") +
      END +
      INSTR(' = "Y" "VIP" "std" ') +
      SEP +
      R("std") +
      END;
    const doc = await read(P(ifField("Vip")) + P(ifField("Other")));
    expect(Query.mailMerge(doc, { Vip: "Y", Other: "N" })).toBe(2);
    expect(text(doc).split("\n")).toEqual(["VIP", "std"]);
  });

  it("evaluates an IF nested in an IF's condition before the outer one", async () => {
    const doc = await read(
      P(
        BEGIN +
          INSTR(" IF ") +
          BEGIN +
          INSTR(" IF ") +
          BEGIN +
          INSTR(" MERGEFIELD A ") +
          SEP +
          R("0") +
          END +
          INSTR(' > 10 "big" "small" ') +
          SEP +
          R("small") +
          END +
          INSTR(' = "big" "yes" "no" ') +
          SEP +
          R("no") +
          END
      )
    );
    Query.mailMerge(doc, { A: "12" });
    expect(text(doc)).toBe("yes");
  });

  // Word's own results for the same document (Word for Mac 16, fields
  // updated in Word over the merge data as literals).
  it("merges fields nested in an IF's branches in place, as Word does", async () => {
    const field = (instr: string, cached = "x"): string =>
      BEGIN + INSTR(instr) + SEP + R(cached) + END;
    const splitMerge = (name: string, sw: string): string =>
      BEGIN +
      INSTR(" MERGE") +
      INSTR(`FIELD ${name.slice(0, 2)}`) +
      INSTR(`${name.slice(2)} ${sw} `) +
      SEP +
      R(`«${name}»`) +
      END;
    const nestedMerge =
      BEGIN +
      INSTR(" IF ") +
      splitMerge("Region", "") +
      INSTR(' = "EU" "') +
      field(' MERGEFIELD Big \\# "#,##0.00" ') +
      INSTR(' EUR" "other" ') +
      SEP +
      R("?") +
      END;
    const nestedIf =
      BEGIN +
      INSTR(" IF ") +
      field(" MERGEFIELD Count ") +
      INSTR(' > 2 "') +
      BEGIN +
      INSTR(" IF ") +
      field(" MERGEFIELD Vip ") +
      INSTR(' = "Y" "many-vip" "many" ') +
      SEP +
      R("?") +
      END +
      INSTR('" "few" ') +
      SEP +
      R("?") +
      END;
    const doc = await read(P(nestedMerge) + P(nestedIf));
    Query.mailMerge(doc, { Big: "1234.5", Vip: "Y", Region: "EU", Count: "3" });
    expect(text(doc).split("\n")).toEqual(["1,234.50 EUR", "many-vip"]);

    // Written, the outer fields keep their instruction in order, so Word —
    // and the library — can update them again.
    const back = await Io.read(await Io.toBuffer(doc));
    expect(
      Query.listFields(back)
        .filter(f => f.type === "IF")
        .map(f => [f.instruction.trim(), f.result])
    ).toEqual([
      ['IF EU = "EU" "1,234.50 EUR" "other"', "1,234.50 EUR"],
      ['IF 3 > 2 "{ IF Y = "Y" "many-vip" "many" }" "few"', "many-vip"],
      ['IF Y = "Y" "many-vip" "many"', "many-vip"]
    ]);
    expect(text(Transform.updateFields(back)).split("\n")).toEqual(["1,234.50 EUR", "many-vip"]);
  });

  it("reads an instruction's text in place between the fields nested in it", async () => {
    const source = P(
      BEGIN +
        INSTR(" IF ") +
        BEGIN +
        INSTR(" MERGEFIELD Region ") +
        SEP +
        R("EU") +
        END +
        INSTR(' = "EU" "') +
        BEGIN +
        INSTR(" MERGEFIELD Big ") +
        SEP +
        R("7") +
        END +
        INSTR(' EUR" "other" ') +
        SEP +
        R("?") +
        END
    );
    const doc = await read(source);
    expect(Query.listFields(doc)[0]!.instruction).toBe(
      ' IF { MERGEFIELD Region } = "EU" "{ MERGEFIELD Big } EUR" "other" '
    );
    expect(text(Transform.updateFields(doc))).toBe("7 EUR");
    const xml = partText(await packageParts(await Io.toBuffer(doc)), "word/document.xml");
    const order = [...xml.matchAll(/w:fldCharType="(\w+)"|<w:instrText[^>]*>([^<]*)/g)]
      .map(m => (m[1] ?? m[2]!).replace(/&quot;/g, '"'))
      .join("|");
    expect(order).toBe(
      'begin| IF |begin|MERGEFIELD Region|separate|end| = "EU" "|begin|MERGEFIELD Big|separate|end| EUR" "other" |separate|end'
    );
  });

  it("leaves an IF alone when nothing in it was merged", async () => {
    const doc = await read(
      P(
        BEGIN +
          INSTR(" IF ") +
          BEGIN +
          INSTR(" MERGEFIELD Missing ") +
          SEP +
          R("Y") +
          END +
          INSTR(' = "Y" "VIP" "std" ') +
          SEP +
          R("stale") +
          END
      )
    );
    expect(Query.mailMerge(doc, { Other: "1" })).toBe(0);
    expect(text(doc)).toBe("stale");
  });
});

describe("field update applies the general switches", () => {
  it("formats SEQ, REF and = results", async () => {
    const doc = await read(
      P(`<w:bookmarkStart w:id="1" w:name="who"/>${R("bob")}<w:bookmarkEnd w:id="1"/>`) +
        P(SIMPLE(" SEQ Fig \\* ROMAN ", "?") + R(" ") + SIMPLE(" SEQ Fig \\* ROMAN ", "?")) +
        P(SIMPLE(" REF who \\* Upper ", "?")) +
        P(SIMPLE(' = 1234.5 * 2 \\# "#,##0.00" ', "?")) +
        P(SIMPLE(' IF 10 > 9 "more" "less" \\* Upper ', "?"))
    );
    expect(text(Transform.updateFields(doc)).split("\n")).toEqual([
      "bob",
      "I II",
      "BOB",
      "2,469.00",
      "MORE"
    ]);
  });

  it("evaluates QUOTE: its text, then its switches", async () => {
    const doc = await read(
      P(SIMPLE(' QUOTE "hello big world" \\* Caps ', "?")) +
        P(SIMPLE(' QUOTE "abc" \\* Upper ', "?")) +
        P(SIMPLE(" QUOTE 21 \\* Ordinal ", "?"))
    );
    expect(text(Transform.updateFields(doc)).split("\n")).toEqual([
      "Hello Big World",
      "ABC",
      "21st"
    ]);
  });

  it("compares IF operands as numbers when both are numbers", async () => {
    const doc = await read(P(SIMPLE(' IF 10 > 9 "more" "less" ', "?")));
    expect(text(Transform.updateFields(doc))).toBe("more");
  });
});

describe("Query.listFields", () => {
  it("shows a nested field by its own code, never by its stale result", async () => {
    const doc = await read(
      P(
        BEGIN +
          INSTR(" IF ") +
          BEGIN +
          INSTR(" MERGEFIELD Vip ") +
          SEP +
          R("x") +
          END +
          INSTR(' = "Y" "VIP" "std" ') +
          SEP +
          R("std") +
          END
      )
    );
    expect(Query.listFields(doc).map(f => [f.type, f.instruction.trim(), f.result])).toEqual([
      ["IF", '{ MERGEFIELD Vip } = "Y" "VIP" "std"'.replace(/^/, "IF "), "std"],
      ["MERGEFIELD", "MERGEFIELD Vip", "x"]
    ]);
  });
});

describe("w:updateFields — Word's prompt to update fields on open", () => {
  const toc = async (): Promise<DocxDocument> => {
    const doc = Document.create();
    Document.addTableOfContents(doc);
    Document.addHeading(doc, "One", 1);
    return doc;
  };

  it("is never set by updating fields: the library computes the page numbers itself", async () => {
    const updated = Transform.updateFields(await toc());
    expect(updated.settings?.updateFieldsOnOpen).toBeUndefined();
    expect(
      Transform.updateTableOfContents(await toc()).settings?.updateFieldsOnOpen
    ).toBeUndefined();
    expect(
      Transform.updateFields(await read(P(SIMPLE(" PAGE ", "9")))).settings?.updateFieldsOnOpen
    ).toBeUndefined();
    const settings = partText(await packageParts(await Io.toBuffer(updated)), "word/settings.xml");
    expect(settings).not.toContain("w:updateFields");
  });

  it("is kept when the caller sets it", async () => {
    const doc = await toc();
    const asked: DocxDocument = { ...doc, settings: { ...doc.settings, updateFieldsOnOpen: true } };
    const updated = Transform.updateFields(asked);
    expect(updated.settings?.updateFieldsOnOpen).toBe(true);
    const settings = partText(await packageParts(await Io.toBuffer(updated)), "word/settings.xml");
    expect(settings).toContain('<w:updateFields w:val="true"/>');
  });

  it("round-trips the setting, including an explicit false", async () => {
    const doc = await toc();
    const on = await Io.read(
      await Io.toBuffer({ ...doc, settings: { ...doc.settings, updateFieldsOnOpen: true } })
    );
    expect(on.settings?.updateFieldsOnOpen).toBe(true);
    const cleared: DocxDocument = {
      ...on,
      settings: { ...on.settings, updateFieldsOnOpen: false }
    };
    const back = await Io.read(await Io.toBuffer(cleared));
    expect(back.settings?.updateFieldsOnOpen).toBeFalsy();
    const off = await Io.read(
      await buildDocx({ body: P(R("x")), settings: '<w:updateFields w:val="false"/>' })
    );
    expect(off.settings?.updateFieldsOnOpen).toBeFalsy();
  });
});
