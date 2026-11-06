/**
 * `headerRow` is structure (`w:tblHeader`: repeated on every page, read as column headers); how the header looks
 * is `headerRunProperties`, as Word keeps "Repeat as header row" apart from the row's formatting.
 */
import { describe, expect, it } from "vitest";

import { Build, Document, DocxArgumentError } from "../../index";
import type { Paragraph, Table } from "../../index";

function firstRun(table: Table, row: number): unknown {
  const paragraph = table.rows[row]!.cells[0]!.content[0] as Paragraph;
  return (paragraph.children[0] as { properties?: unknown }).properties;
}

describe("Build.simpleTable header row", () => {
  const DATA = [
    ["Name", "Age"],
    ["Ada", "36"]
  ];

  it("marks the header row and leaves its text unformatted by default, as before", () => {
    const table = Build.simpleTable(DATA);
    expect(table.rows[0]!.properties).toEqual({ tableHeader: true });
    expect(firstRun(table, 0)).toBeUndefined();
  });

  it("formats only the header row's text with headerRunProperties", () => {
    const table = Build.simpleTable(DATA, { headerRunProperties: { bold: true } });
    expect(table.rows[0]!.properties).toEqual({ tableHeader: true });
    expect(firstRun(table, 0)).toEqual({ bold: true });
    expect(table.rows[0]!.cells[1]!.content[0]).toMatchObject({
      children: [{ properties: { bold: true } }]
    });
    expect(firstRun(table, 1)).toBeUndefined();
  });

  it("is what Document.addTable builds from a grid", () => {
    const doc = Document.create();
    Document.addTable(doc, DATA, { headerRunProperties: { bold: true, color: { val: "FF0000" } } });
    expect(firstRun(doc.body[0] as Table, 0)).toEqual({ bold: true, color: { val: "FF0000" } });
  });

  it("refuses header formatting for a table without a header row", () => {
    expect(() =>
      Build.simpleTable(DATA, { headerRow: false, headerRunProperties: { bold: true } })
    ).toThrow(DocxArgumentError);
  });

  it("checks the run properties", () => {
    try {
      Build.simpleTable(DATA, { headerRunProperties: { size: Number.NaN } });
      expect.unreachable();
    } catch (error) {
      expect((error as DocxArgumentError).field).toMatch(/^options\.headerRunProperties\.size/);
    }
  });
});
