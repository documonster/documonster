/**
 * A loop repeats over a list, and a value that is not one is a list of one: a
 * template must not depend on how many items its data happens to hold. XML is
 * where that bites — it cannot say "a list of one", so a single `<item>` reads
 * as an object where two read as an array, and an invoice with one line used
 * to fail where one with two filled.
 */
import { describe, it, expect } from "vitest";

import { Document, Query, Template } from "../../index";
import type { DocxDocument, Table } from "../../index";

function paragraphs(...texts: string[]): DocxDocument {
  const doc = Document.create();
  for (const text of texts) {
    Document.addParagraph(doc, text);
  }
  return doc;
}

function cell(text: string) {
  return {
    content: [
      { type: "paragraph" as const, children: [{ content: [{ type: "text" as const, text }] }] }
    ]
  };
}

function tableDoc(rows: string[][]): DocxDocument {
  const table: Table = { type: "table", rows: rows.map(r => ({ cells: r.map(cell) })) };
  return { body: [table] };
}

function cellTexts(doc: DocxDocument): string[][] {
  return (doc.body[0] as Table).rows.map(row =>
    row.cells.map(c => Query.getText({ body: [...c.content] }))
  );
}

const INVOICE = paragraphs(
  "Invoice {{number}}",
  "{{#each items.item}}",
  "{{name}}: {{amount}}",
  "{{/each}}",
  "Total: {{total}}"
);

function invoiceXml(...items: [string, string][]): string {
  const lines = items
    .map(([name, amount]) => `<item><name>${name}</name><amount>${amount}</amount></item>`)
    .join("");
  return `<invoice><number>INV-1</number><items>${lines}</items><total>9</total></invoice>`;
}

describe("loops over a value that is not an array", () => {
  it("fill an XML invoice with one line as they fill one with two", () => {
    const one = Template.fillFromSource(
      INVOICE,
      Template.createXmlDataSource(invoiceXml(["Widget", "10"]))
    );
    expect(Query.getText(one)).toBe("Invoice INV-1\nWidget: 10\nTotal: 9");

    const two = Template.fillFromSource(
      INVOICE,
      Template.createXmlDataSource(invoiceXml(["Widget", "10"], ["Gadget", "20"]))
    );
    expect(Query.getText(two)).toBe("Invoice INV-1\nWidget: 10\nGadget: 20\nTotal: 9");
  });

  it("repeat a nested loop once for a single XML child", () => {
    const doc = paragraphs(
      "{{#each order}}",
      "{{id}}",
      "{{#each line}}",
      "- {{sku}} ({{@index}})",
      "{{/each}}",
      "{{/each}}"
    );
    const xml = "<root><order><id>A</id><line><sku>x</sku></line></order></root>";
    const filled = Template.fillFromSource(doc, Template.createXmlDataSource(xml));
    expect(Query.getText(filled)).toBe("A\n- x (0)");
  });

  it("treat a single text element as one item", () => {
    const doc = paragraphs("Tags: {{#each tags.tag}}[{{.}}]{{/each}}");
    const filled = Template.fillFromSource(
      doc,
      Template.createXmlDataSource("<r><tags><tag>solo</tag></tags></r>")
    );
    expect(Query.getText(filled)).toBe("Tags: [solo]");
  });

  it("repeat a block loop once for an object, in strict mode too", () => {
    const doc = paragraphs("{{#each customer}}", "{{name}}", "{{/each}}");
    expect(Query.getText(Template.fill(doc, { customer: { name: "ACME" } }))).toBe("ACME");
    expect(
      Query.getText(Template.fill(doc, { customer: { name: "ACME" } }, { strict: false }))
    ).toBe("ACME");
  });

  it("repeat an inline loop once for a scalar", () => {
    const doc = paragraphs("[{{#each n}}{{.}}{{/each}}]");
    expect(Query.getText(Template.fill(doc, { n: 0 }))).toBe("[0]");
  });

  it("repeat a table-row loop once for an object", () => {
    const doc = tableDoc([["{{#each rows}}{{name}}{{/each}}"]]);
    expect(cellTexts(Template.fill(doc, { rows: { name: "only" } }))).toEqual([["only"]]);
  });

  it("repeat a multi-row table loop once for an object", () => {
    const doc = tableDoc([["{{#each rows}}{{name}}"], ["{{qty}}{{/each}}"], ["end"]]);
    expect(cellTexts(Template.fill(doc, { rows: { name: "a", qty: "1" } }))).toEqual([
      ["a"],
      ["1"],
      ["end"]
    ]);
  });

  it("repeat a #cols loop once for a scalar", () => {
    const doc = tableDoc([["Name", "{{#cols months}}{{.}}{{/cols}}", "Sum"]]);
    const filled = Template.fill(doc, { months: "Jan" }, { richPlaceholders: true });
    expect(cellTexts(filled)).toEqual([["Name", "Jan", "Sum"]]);
  });

  it("repeat nothing for undefined or null", () => {
    const doc = paragraphs("start", "{{#each items}}", "{{.}}", "{{/each}}", "end");
    expect(Query.getText(Template.fill(doc, {}))).toBe("start\nend");
    expect(Query.getText(Template.fill(doc, { items: null }))).toBe("start\nend");
    const rows = tableDoc([["{{#each rows}}{{.}}{{/each}}"], ["end"]]);
    expect(cellTexts(Template.fill(rows, { rows: null }))).toEqual([["end"]]);
  });
});

describe("Template.createXmlDataSource getArray", () => {
  it("returns a single element as a one-item array", () => {
    const source = Template.createXmlDataSource(invoiceXml(["Widget", "10"]));
    expect(source.getArray("items.item")).toEqual([{ name: "Widget", amount: "10" }]);
  });

  it("returns an empty array for a path that names nothing", () => {
    const source = Template.createXmlDataSource(invoiceXml(["Widget", "10"]));
    expect(source.getArray("items.missing")).toEqual([]);
  });

  it("leaves a JSON source's non-array values out", () => {
    const source = Template.createJsonDataSource({ name: "Alice" });
    expect(source.getArray("name")).toEqual([]);
  });
});
