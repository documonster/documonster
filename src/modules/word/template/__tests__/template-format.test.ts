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

  it("reads a typographic quote as Word types one", () => {
    expect(fillText("{{a \\# \u201C#,##0.00\u201D}}", { a: "1234.5" })).toBe("1,234.50");
  });
});

describe("date pictures (\\@)", () => {
  it("reads ISO, XML Schema and SAP dates", () => {
    expect(fillText('{{d \\@ "dd.MM.yyyy"}}', { d: "2026-11-07" })).toBe("07.11.2026");
    expect(fillText('{{d \\@ "dd.MM.yyyy HH:mm"}}', { d: "2026-11-07T14:05:09.120+05:30" })).toBe(
      "07.11.2026 14:05"
    );
    expect(fillText('{{d \\@ "dd.MM.yyyy"}}', { d: "2026-11-07Z" })).toBe("07.11.2026");
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
    for (const tag of ['{{a \\x "0"}}', "{{a \\* Bogus}}", "{{a \\#}}", '{{a \\# "0" junk}}']) {
      expect(() => fillText(tag, { a: "1" }, { strict: false }), tag).toThrow(DocxTemplateError);
    }
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
