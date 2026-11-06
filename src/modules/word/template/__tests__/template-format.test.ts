/**
 * Format switches in placeholders are Word's own field switches —
 * `{{amount \# "#,##0.00"}}`, `{{date \@ "d MMMM yyyy"}}`, `{{name \* Upper}}`,
 * `{{phone \b "Tel: "}}` — read as a `MERGEFIELD` reads them, and written in
 * the locale the fill is given, as Word writes in its regional settings.
 */
import { describe, it, expect } from "vitest";

import {
  Build,
  Document,
  DocxArgumentError,
  DocxTemplateError,
  Query,
  Template
} from "../../index";
import type { DocxDocument, TemplateOptions } from "../../index";
import { localeFieldFormat } from "../template-format";

function fillText(text: string, data: Record<string, unknown>, options?: TemplateOptions): string {
  const doc = Document.create();
  Document.addParagraph(doc, text);
  return Query.getText(Template.fill(doc, data, options));
}

describe("numeric pictures (\\#)", () => {
  it("formats a number written as XML and JSON write one", () => {
    expect(fillText('{{a \\# "#,##0.00"}}', { a: "1234567.5" })).toBe("1,234,567.50");
    expect(fillText('{{a \\# "#,##0.00"}}', { a: 1234567.5 })).toBe("1,234,567.50");
    expect(fillText('{{a \\# "0.0"}}', { a: "-0.04" })).toBe("0.0");
  });

  it("takes sections for negative and zero", () => {
    const tag = "{{a \\# \"0.00;(0.00);'nil'\"}}";
    expect(fillText(tag, { a: "-5" })).toBe("(5.00)");
    expect(fillText(tag, { a: "0" })).toBe("nil");
  });

  it("reads a picture in the locale's symbols and writes them", () => {
    expect(fillText('{{a \\# "#.##0,00"}}', { a: "1234567.5" }, { locale: "de-DE" })).toBe(
      "1.234.567,50"
    );
    expect(fillText('{{a \\# "#,##0.00"}}', { a: "1234567.5" }, { locale: "en-IN" })).toBe(
      "12,34,567.50"
    );
  });

  it("takes any space as the grouping symbol of a locale that groups with one", () => {
    const out = fillText('{{a \\# "# ##0,00"}}', { a: "1234567.5" }, { locale: "fr-FR" });
    expect(out.replace(/\s/gu, " ")).toBe("1 234 567,50");
  });

  it("keeps a space that is not between digit places as text in such a locale", () => {
    const fr = { locale: "fr-FR" };
    const norm = (s: string): string => s.replace(/\s/gu, " ");
    expect(norm(fillText('{{a \\# "# ##0,00 €"}}', { a: "1234.5" }, fr))).toBe("1 234,50 €");
    expect(norm(fillText('{{a \\# "€ # ##0,00"}}', { a: "1234.5" }, fr))).toBe("€ 1 234,50");
  });

  it("takes an apostrophe between digit places as the grouping symbol of a locale that groups with one", () => {
    const ch = { locale: "de-CH" };
    const group = new Intl.NumberFormat("de-CH", {
      useGrouping: "always"
    } as Intl.NumberFormatOptions)
      .formatToParts(1234)
      .find(part => part.type === "group")!.value;
    expect(fillText("{{a \\# \"#'##0.00 'CHF';-#'##0.00\"}}", { a: "1234.5" }, ch)).toBe(
      `1${group}234.50 CHF`
    );
    expect(fillText("{{a \\# \"#'##0.00 'CHF';-#'##0.00\"}}", { a: "-1234.5" }, ch)).toBe(
      `-1${group}234.50`
    );
    // As Word types it, turned into `'` like every other typographic quote.
    expect(fillText('{{a \\# "#\u2019##0.00"}}', { a: "1234.5" }, ch)).toBe(`1${group}234.50`);
  });

  it("reads a typographic quote as Word types one", () => {
    expect(fillText("{{a \\# \u201C#,##0.00\u201D}}", { a: "1234.5" })).toBe("1,234.50");
  });
});

describe("date pictures (\\@)", () => {
  it("refuses a date-time with a part out of range, and reads XML Schema's end of day", () => {
    const pic = '{{d \\@ "yyyy-MM-dd HH:mm:ss"}}';
    for (const d of [
      "2026-11-07T99:99:99",
      "2026-11-07T23:60:00",
      "2026-11-07 24:00",
      "2026-02-29",
      "2026-11-07T12:30:00+15:00",
      "2026-11-07T12:30:00+14:30",
      "2026-11-07T12:30:00+01:60",
      "2026-11-07T12:30.123",
      "2026-11-07T24:00:01",
      "20261131"
    ]) {
      expect(() => fillText(pic, { d }), d).toThrow(/is not a date/);
    }
    expect(fillText(pic, { d: "2026-12-31T24:00:00Z" })).toBe("2027-01-01 00:00:00");
    expect(fillText(pic, { d: "2026-11-07T12:30:00.5-14:00" })).toBe("2026-11-07 12:30:00");
  });

  it("names the Gregorian month the data holds, whatever calendar the locale uses", () => {
    const pic = '{{d \\@ "d MMMM yyyy"}}';
    for (const locale of ["fa-IR", "en-US-u-ca-islamic", "he-IL-u-ca-hebrew", "th-TH"]) {
      const month = new Intl.DateTimeFormat(locale, {
        calendar: "gregory",
        day: "numeric",
        month: "long",
        timeZone: "UTC"
      })
        .formatToParts(Date.UTC(2026, 10, 7))
        .find(part => part.type === "month")!.value;
      expect(fillText(pic, { d: "2026-11-07" }, { locale }), locale).toBe(`7 ${month} 2026`);
    }
    expect(fillText(pic, { d: "2026-11-07" }, { locale: "en-US-u-ca-islamic" })).toBe(
      "7 November 2026"
    );
  });

  it("reads a year below 100 as written, and names its weekday", () => {
    expect(fillText('{{d \\@ "yyyy-MM-dd dddd"}}', { d: "0096-02-29" })).toBe(
      "0096-02-29 Wednesday"
    );
    expect(() => fillText('{{d \\@ "yyyy"}}', { d: "0100-02-29" })).toThrow(/is not a date/);
  });

  it("reads ISO, XML Schema and SAP dates", () => {
    expect(fillText('{{d \\@ "dd.MM.yyyy"}}', { d: "2026-11-07" })).toBe("07.11.2026");
    expect(fillText('{{d \\@ "dd.MM.yyyy HH:mm"}}', { d: "2026-11-07T14:05:09.120+05:30" })).toBe(
      "07.11.2026 14:05"
    );
    expect(fillText('{{d \\@ "dd.MM.yyyy"}}', { d: "2026-11-07Z" })).toBe("07.11.2026");
    // Fractional seconds without a zone, as .NET and Java write a local date-time.
    expect(fillText('{{d \\@ "dd.MM.yyyy HH:mm"}}', { d: "2026-11-07T14:05:09.1234567" })).toBe(
      "07.11.2026 14:05"
    );
    expect(fillText('{{d \\@ "dd.MM.yyyy"}}', { d: "20261107" })).toBe("07.11.2026");
  });

  it("reads a Date by its UTC date and time", () => {
    expect(fillText('{{d \\@ "d MMM yyyy"}}', { d: new Date(Date.UTC(2026, 10, 7, 23)) })).toBe(
      "7 Nov 2026"
    );
  });

  it("names months and days in the locale's language, as they stand in a date", () => {
    expect(
      fillText('{{d \\@ "dddd, d. MMMM yyyy"}}', { d: "2026-11-07" }, { locale: "de-DE" })
    ).toBe("Samstag, 7. November 2026");
    expect(fillText('{{d \\@ "d MMMM yyyy"}}', { d: "2026-11-07" }, { locale: "ru-RU" })).toBe(
      "7 ноября 2026"
    );
  });
});

describe("text switches", () => {
  it("applies \\* case and number formats", () => {
    expect(fillText("{{n \\* Upper}}|{{n \\* Caps}}", { n: "acme corp" })).toBe(
      "ACME CORP|Acme Corp"
    );
    expect(fillText("{{n \\* DollarText}}", { n: "12.5" })).toBe("twelve and 50/100");
    expect(fillText('{{n \\# "0" \\* roman}}', { n: "14" })).toBe("xiv");
  });

  it("adds \\b and \\f text around a value that is not empty, and nothing around one that is", () => {
    const tag = 'Tel{{p \\b ": " \\f "."}}|';
    expect(fillText(tag, { p: "123" })).toBe("Tel: 123.|");
    expect(fillText(tag, { p: "" })).toBe("Tel|");
    expect(fillText(tag, { p: null })).toBe("Tel|");
  });

  it("writes an XML element's text, attributes beside it", () => {
    const amount = { "@currency": "EUR", "#text": "1234.5" };
    expect(fillText('{{a \\# "#,##0.00"}} {{a.@currency}}', { a: amount })).toBe("1,234.50 EUR");
  });

  it("still reads a key that holds a space", () => {
    expect(fillText("{{first name}}", { "first name": "Jo" })).toBe("Jo");
  });
});

describe("format errors", () => {
  it("throws for a value a picture cannot take in strict mode, and writes it otherwise", () => {
    expect(() => fillText('{{a \\# "0.00"}}', { a: "abc" })).toThrow(
      /Cannot format \{\{a \\# "0\.00"\}\}: "abc" is not a number/
    );
    expect(() => fillText('{{d \\@ "yyyy"}}', { d: "soon" })).toThrow(/"soon" is not a date/);
    expect(fillText('{{a \\# "0.00"}}', { a: "abc" }, { strict: false })).toBe("abc");
  });

  it("throws for a switch that is not one, whatever the mode", () => {
    for (const tag of [
      '{{a \\# "0.00}}',
      '{{a \\x "0"}}',
      "{{a \\* Bogus}}",
      "{{a \\#}}",
      '{{a \\# "0" junk}}'
    ]) {
      expect(() => fillText(tag, { a: "1" }, { strict: false }), tag).toThrow(DocxTemplateError);
    }
  });

  it("throws for a switch that is not one where it is never written, and changes nothing", () => {
    for (const bad of ["{{a \\* Bogus}}", '{{a \\# "0.00}}']) {
      const templates = [
        [`{{#if no}}${bad}{{/if}}`],
        ["{{#if no}}", bad, "{{/if}}"],
        ["{{#each none}}", bad, "{{/each}}"],
        [`{{#each none}}${bad}{{/each}}`]
      ];
      for (const paragraphs of templates) {
        const doc = Document.create();
        for (const text of paragraphs) {
          Document.addParagraph(doc, text);
        }
        const before = Query.getText(doc);
        expect(
          () => Template.fill(doc, { no: false, none: [] }, { strict: false }),
          paragraphs.join("|")
        ).toThrow(/Invalid placeholder/);
        expect(Query.getText(doc)).toBe(before);
      }
    }
  });

  it("formats a locale tag as the locale it resolves to, and keeps one entry for all its spellings", () => {
    // A private-use subtag is a new tag every time; caching by the tag grew without bound in a long-running server.
    const tags = Array.from({ length: 50 }, (_, i) => `de-DE-x-run${i}`);
    for (const locale of tags) {
      expect(fillText('{{a \\# "#.##0,00"}}', { a: "1234.5" }, { locale })).toBe("1.234,50");
    }
    expect(new Set(tags.map(localeFieldFormat)).size).toBe(1);
    expect(localeFieldFormat("de-DE-x-other")).toBe(localeFieldFormat("de-DE"));
  });

  it("refuses a locale this platform cannot format in", () => {
    expect(() => fillText("{{a}}", { a: "1" }, { locale: "not a locale" })).toThrow(
      DocxArgumentError
    );
    expect(() => fillText("{{a}}", { a: "1" }, { locale: "" })).toThrow(/options\.locale/);
  });
});

describe("Template.listTags", () => {
  it("gives a variable's data path without its switches", () => {
    const doc = Document.create();
    Document.addParagraph(doc, '{{total \\# "#,##0.00"}} {{name}}');
    const tags = Template.listTags(doc).filter(t => t.type === "variable");
    expect(tags.map(t => [t.expression, t.path])).toEqual([
      ['total \\# "#,##0.00"', "total"],
      ["name", "name"]
    ]);
  });
});

describe("loop variables", () => {
  function loop(body: string, items: unknown[]): string {
    const doc: DocxDocument = Document.create();
    Document.addParagraph(doc, Build.paragraph([Build.text(`{{#each l}}${body}{{/each}}`)]));
    return Query.getText(Template.fill(doc, { l: items }));
  }

  it("numbers from 0 and from 1, and marks the first and the last", () => {
    expect(
      loop("{{@index}}/{{@number}}{{#if @first}}F{{/if}}{{#if @last}}L{{/if}} ", ["a", "b", "c"])
    ).toBe("0/1F 1/2 2/3L ");
  });

  it("marks a single item as both first and last", () => {
    expect(loop("{{#if @first}}F{{/if}}{{#if @last}}L{{/if}}", ["a"])).toBe("FL");
  });

  it("win over an item key of the same name, which .name still reads", () => {
    expect(loop("{{@number}}={{.@number}} ", [{ "@number": "10" }, { "@number": "20" }])).toBe(
      "1=10 2=20 "
    );
  });

  it("reads .name from the item, not from the data around the loop", () => {
    const doc = Document.create();
    Document.addParagraph(doc, "{{#each l}}[{{.name}}|{{name}}]{{/each}}");
    expect(
      Query.getText(
        Template.fill(doc, { name: "outer", l: [{ name: "a" }, {}] }, { strict: false })
      )
    ).toBe("[a|a][|outer]");
  });
});
