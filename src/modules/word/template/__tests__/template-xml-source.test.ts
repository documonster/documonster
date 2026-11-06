/**
 * Filling a template from XML as an ERP or billing system exports it: values
 * carry attributes (`<amount currency="EUR">`), absent values are `xsi:nil`,
 * and names are namespace-prefixed. A template reads an element that has
 * attributes as its text, so adding an attribute to the data never changes
 * what `{{amount}}` writes or whether `{{#if amount}}` is taken.
 */
import { describe, it, expect } from "vitest";

import { Document, Query, Template } from "../../index";
import type { DocxDocument } from "../../index";

function paragraphs(...texts: string[]): DocxDocument {
  const doc = Document.create();
  for (const text of texts) {
    Document.addParagraph(doc, text);
  }
  return doc;
}

function fillXml(doc: DocxDocument, xml: string): string {
  return Query.getText(Template.fillFromSource(doc, Template.createXmlDataSource(xml)));
}

const XSI = 'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"';

describe("Template.fillFromSource with XML data", () => {
  it("writes an attributed element as its text, and its attributes by name", () => {
    const doc = paragraphs("{{amount}} {{amount.@currency}}");
    expect(fillXml(doc, '<r><amount currency="EUR">10.00</amount></r>')).toBe("10.00 EUR");
  });

  it("tests an attributed element by its text", () => {
    const doc = paragraphs("{{#if note}}note{{else}}none{{/if}}");
    expect(fillXml(doc, '<r><note lang="en"></note></r>')).toBe("none");
    expect(fillXml(doc, '<r><note lang="en">x</note></r>')).toBe("note");
  });

  it("writes mixed content as the element's whole text", () => {
    const doc = paragraphs("{{p}}");
    expect(fillXml(doc, "<r><p>Hello <b>x</b> world</p></r>")).toBe("Hello x world");
  });

  it("writes an xsi:nil element as nothing and does not take its #if", () => {
    const doc = paragraphs("[{{discount}}]{{#if discount}} off{{/if}}");
    expect(fillXml(doc, `<r ${XSI}><discount xsi:nil="true"/></r>`)).toBe("[]");
  });

  it("repeats nothing for an xsi:nil list", () => {
    const doc = paragraphs("start", "{{#each items}}", "{{.}}", "{{/each}}", "end");
    expect(fillXml(doc, `<r ${XSI}><items xsi:nil="true"/></r>`)).toBe("start\nend");
  });

  it("reads loop items with attributes, one or several", () => {
    const doc = paragraphs("{{#each items.item}}[{{.}} {{.@unit}}]{{/each}}");
    expect(fillXml(doc, '<r><items><item unit="PC">5</item></items></r>')).toBe("[5 PC]");
    expect(
      fillXml(doc, '<r><items><item unit="PC">5</item><item unit="KG">2</item></items></r>')
    ).toBe("[5 PC][2 KG]");
  });

  it("fills a namespace-prefixed SAP-style invoice", () => {
    const doc = paragraphs(
      "Invoice {{header.number}} ({{header.@type}})",
      "{{#each lines.line}}",
      "{{@index}}. {{text}}: {{net}} {{net.@currency}}",
      "{{/each}}",
      "Total: {{total}} {{total.@currency}}"
    );
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<n0:Invoice xmlns:n0="urn:sap-com:document:sap:rfc:functions" ${XSI}>
  <n0:header type="F2">
    <n0:number>90000001</n0:number>
  </n0:header>
  <n0:lines>
    <n0:line>
      <n0:text>Widget</n0:text>
      <n0:net currency="EUR">10.00</n0:net>
    </n0:line>
  </n0:lines>
  <n0:total currency="EUR">10.00</n0:total>
</n0:Invoice>`;
    expect(fillXml(doc, xml)).toBe("Invoice 90000001 (F2)\n0. Widget: 10.00 EUR\nTotal: 10.00 EUR");
  });
});
