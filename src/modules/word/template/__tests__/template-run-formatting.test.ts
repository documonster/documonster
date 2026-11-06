/**
 * Filling a paragraph changes only the placeholders in it. A Word template
 * mixes formatting in a line — a bold label, a tab, then the value — and every
 * paragraph used to be merged into one run in the first run's formatting, so
 * the label's bold spread to the value and its tabs and line breaks vanished.
 */
import { describe, it, expect } from "vitest";

import { Build, Document, Template } from "../../index";
import type { DocxDocument, Paragraph, Run, Table } from "../../index";

function docOf(...children: Run[]): DocxDocument {
  const doc = Document.create();
  Document.addParagraph(doc, Build.paragraph(children));
  return doc;
}

/** Each run as `[bold?, content]`, a tab as `⇥`, a line break as `↵`. */
function runs(p: Paragraph): [boolean, string][] {
  return p.children.map(child => {
    const run = child as Run;
    const text = run.content
      .map(c =>
        c.type === "text"
          ? c.text
          : c.type === "tab"
            ? "⇥"
            : c.type === "break"
              ? "↵"
              : `<${c.type}>`
      )
      .join("");
    return [run.properties?.bold === true, text];
  });
}

function filled(
  doc: DocxDocument,
  data: Record<string, unknown>,
  options = {}
): [boolean, string][] {
  return runs(Template.fill(doc, data, options).body[0] as Paragraph);
}

describe("Template.fill keeps a paragraph's runs", () => {
  it("keeps a bold label's formatting off the value after it", () => {
    const doc = docOf(Build.bold("Invoice No: "), Build.text("{{n}}"));
    expect(filled(doc, { n: "9" })).toEqual([
      [true, "Invoice No: "],
      [false, "9"]
    ]);
  });

  it("writes a value in the formatting of the run its placeholder is in", () => {
    const doc = docOf(Build.text("Total: "), Build.bold("{{n}}"), Build.text(" due"));
    expect(filled(doc, { n: "10" })).toEqual([
      [false, "Total: "],
      [true, "10"],
      [false, " due"]
    ]);
  });

  it("keeps the tabs and line breaks between runs", () => {
    const doc = docOf(
      Build.bold("No:"),
      Build.tab(),
      Build.text("{{n}}"),
      Build.lineBreak(),
      Build.bold("Date:"),
      Build.tab(),
      Build.text("{{d}}")
    );
    expect(filled(doc, { n: "9", d: "1 Jan" })).toEqual([
      [true, "No:"],
      [false, "⇥"],
      [false, "9"],
      [false, "↵"],
      [true, "Date:"],
      [false, "⇥"],
      [false, "1 Jan"]
    ]);
  });

  it("writes a line break or a tab in a value as one", () => {
    const doc = docOf(Build.text("To: {{address}}"));
    expect(filled(doc, { address: "1 Main St\nSpringfield\r\nUSA\tX" })).toEqual([
      [false, "To: 1 Main St↵Springfield↵USA⇥X"]
    ]);
    const content = (Template.fill(doc, { address: "a\nb" }).body[0] as Paragraph)
      .children[0] as Run;
    expect(content.content.map(c => c.type)).toEqual(["text", "break", "text"]);
  });

  it("writes a placeholder split across runs where it starts", () => {
    const doc = docOf(Build.text("A {{to"), Build.bold("tal}} B"));
    expect(filled(doc, { total: "10" })).toEqual([
      [false, "A 10"],
      [true, " B"]
    ]);
  });

  it("keeps the runs of a taken inline #if branch", () => {
    const doc = docOf(
      Build.text("Total{{#if vip}} "),
      Build.bold("VIP"),
      Build.text("{{/if}} end")
    );
    expect(filled(doc, { vip: true })).toEqual([
      [false, "Total "],
      [true, "VIP"],
      [false, " end"]
    ]);
  });

  it("removes a branch not taken, with the tabs and breaks inside it", () => {
    const doc = docOf(
      Build.text("A{{#if x}}"),
      Build.tab(),
      Build.bold("{{v}}"),
      Build.lineBreak(),
      Build.text("{{else}}-{{/if}}B")
    );
    expect(filled(doc, { x: false })).toEqual([
      [false, "A"],
      [false, "-B"]
    ]);
    expect(filled(doc, { x: true, v: "1" })).toEqual([
      [false, "A"],
      [false, "⇥"],
      [true, "1"],
      [false, "↵"],
      [false, "B"]
    ]);
  });

  it("resolves nested inline #if branches", () => {
    const doc = docOf(
      Build.text("[{{#if a}}a{{#if b}}"),
      Build.bold("b"),
      Build.text("{{/if}}{{/if}}]")
    );
    expect(filled(doc, { a: true, b: true })).toEqual([
      [false, "[a"],
      [true, "b"],
      [false, "]"]
    ]);
    expect(filled(doc, { a: true, b: false })).toEqual([
      [false, "[a"],
      [false, "]"]
    ]);
    expect(filled(doc, { a: false, b: true })).toEqual([
      [false, "["],
      [false, "]"]
    ]);
  });

  it("still throws on an unclosed inline #if in strict mode", () => {
    const doc = docOf(Build.text("A {{#if x}}"), Build.bold("B"));
    expect(() => Template.fill(doc, { x: true })).toThrow(/Unclosed inline \{\{#if x\}\}/);
  });

  it("keeps a #cols cell's formatting", () => {
    const cell = {
      content: [
        Build.paragraph([
          Build.text("{{#cols months}}"),
          Build.bold("{{.}}"),
          Build.text("{{/cols}}")
        ])
      ]
    };
    const table: Table = { type: "table", rows: [{ cells: [cell] }] };
    const doc: DocxDocument = { body: [table] };
    const out = Template.fill(doc, { months: ["Jan", "Feb"] }, { richPlaceholders: true });
    const cells = (out.body[0] as Table).rows[0]!.cells;
    expect(cells.map(c => runs(c.content[0] as Paragraph).filter(([, t]) => t !== ""))).toEqual([
      [[true, "Jan"]],
      [[true, "Feb"]]
    ]);
  });
});
