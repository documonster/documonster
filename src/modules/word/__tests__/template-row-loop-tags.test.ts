/**
 * The tags of a table-row `{{#each}}` loop: where in a row they may sit, the
 * forms they may take, and what removing them leaves of the paragraph that
 * held them.
 */

import { describe, expect, it } from "vitest";

import { Document, Template } from "../index";
import type {
  DocxDocument,
  Hyperlink,
  Paragraph,
  ParagraphChild,
  Run,
  Table,
  TableCell,
  TableRow
} from "../types";

const run = (text: string, bold?: boolean): Run => ({
  ...(bold ? { properties: { bold: true } } : {}),
  content: [{ type: "text", text }]
});
const p = (...children: ParagraphChild[]): Paragraph => ({ type: "paragraph", children });
const t = (text: string): Paragraph => p(run(text));
const cell = (...content: (Paragraph | Table)[]): TableCell => ({ content });
const row = (...cells: TableCell[]): TableRow => ({ cells });
const table = (...rows: TableRow[]): Table => ({ type: "table", rows });

function fill(body: Table, data: Record<string, unknown>): Table {
  const doc: DocxDocument = Document.create();
  Document.addContent(doc, body);
  return Template.fill(doc, data).body[0] as Table;
}

function runText(r: Run): string {
  return r.content.map(c => (c.type === "text" ? c.text : "")).join("");
}

/** The text of each cell's top-level paragraphs, row by row. */
function grid(tbl: Table): string[][] {
  return tbl.rows.map(r =>
    r.cells.map(c =>
      c.content
        .map(b =>
          b.type === "paragraph"
            ? b.children.map(ch => ("content" in ch && !("type" in ch) ? runText(ch) : "")).join("")
            : "[table]"
        )
        .join("|")
    )
  );
}

describe("table-row {{#each}} tags", () => {
  it("are found in any cell of the opening and closing rows", () => {
    const filled = fill(
      table(
        row(cell(t("Name")), cell(t("{{#each xs}}"))),
        row(cell(t("{{.}}")), cell(t("x"))),
        row(cell(t("{{/each}}")), cell(t("end")))
      ),
      { xs: ["a", "b"] }
    );
    expect(grid(filled)).toEqual([
      ["Name", ""],
      ["a", "x"],
      ["", "end"],
      ["Name", ""],
      ["b", "x"],
      ["", "end"]
    ]);
  });

  it("are accepted with spacing inside the delimiters, opening and closing alike", () => {
    const filled = fill(
      table(row(cell(t("{{ #each xs }}"))), row(cell(t("{{.}}"))), row(cell(t("{{ /each }}")))),
      { xs: ["a", "b"] }
    );
    expect(grid(filled)).toEqual([[""], ["a"], [""], [""], ["b"], [""]]);
  });

  it("count a spaced nested loop when finding the closing row", () => {
    const filled = fill(
      table(
        row(cell(t("{{#each xs}}"))),
        row(cell(t("{{ #each ys }}{{.}}{{ /each }}"))),
        row(cell(t("{{/each}}")))
      ),
      { xs: ["a"], ys: ["1", "2"] }
    );
    expect(grid(filled)).toEqual([[""], ["12"], [""]]);
  });

  it("in a nested table belong to that table, not to the row around it", () => {
    const inner = table(
      row(cell(t("{{#each ys}}"))),
      row(cell(t("{{.}}"))),
      row(cell(t("{{/each}}")))
    );
    const filled = fill(
      table(row(cell(t("{{#each xs}}"))), row(cell(t("{{.}}"), inner)), row(cell(t("{{/each}}")))),
      { xs: ["a", "b"], ys: ["1"] }
    );
    expect(grid(filled)).toEqual([[""], ["a|[table]"], [""], [""], ["b|[table]"], [""]]);
    const nested = filled.rows[1]!.cells[0]!.content[1] as Table;
    expect(grid(nested)).toEqual([[""], ["1"], [""]]);
  });

  it("on one row may sit in different cells, and the rest of the row keeps its formatting", () => {
    const filled = fill(
      table(
        row(cell(t("{{#each xs}}"), p(run("Label", true))), cell(t("{{.}}")), cell(t("{{/each}}")))
      ),
      { xs: ["a", "b"] }
    );
    expect(grid(filled)).toEqual([
      ["|Label", "a", ""],
      ["|Label", "b", ""]
    ]);
    expect((filled.rows[1]!.cells[0]!.content[1] as Paragraph).children).toEqual([
      run("Label", true)
    ]);
  });

  it("are removed from their own runs, keeping the paragraph's formatting and hyperlinks", () => {
    const link: Hyperlink = { type: "hyperlink", anchor: "top", children: [run("link")] };
    const opening = p(run("Items "), run("{{#each xs}}"), run("bold", true), link);
    const filled = fill(
      table(row(cell(opening)), row(cell(t("{{.}}"))), row(cell(t("{{/each}}")))),
      {
        xs: ["a"]
      }
    );
    const para = filled.rows[0]!.cells[0]!.content[0] as Paragraph;
    expect(para.children).toEqual([run("Items "), run("bold", true), link]);
  });

  it("are removed from a run they share with other text, keeping its formatting", () => {
    const closing = p(run("Total "), run("{{/each}} done", true));
    const filled = fill(
      table(row(cell(t("{{#each xs}}"))), row(cell(t("{{.}}"))), row(cell(closing))),
      {
        xs: ["a"]
      }
    );
    const para = filled.rows[2]!.cells[0]!.content[0] as Paragraph;
    expect(para.children).toEqual([run("Total "), run(" done", true)]);
  });

  describe("a single-row loop matches its close by nesting depth", () => {
    const data = {
      rows: [{ tags: ["a", "b"] }, { tags: ["c"] }]
    };
    const nestedThenClose = () =>
      table(row(cell(t("{{#each rows}} {{#each tags}}{{.}}{{/each}}"), t("{{/each}}"))));

    it.each([true, false])(
      "treats a nested inline loop before the close in another paragraph as a row loop (strict: %s)",
      strict => {
        const doc: DocxDocument = Document.create();
        Document.addContent(doc, nestedThenClose());
        const out = Template.fill(doc, data, { strict }).body[0] as Table;
        expect(grid(out)).toEqual([[" ab|"], [" c|"]]);
      }
    );

    it("keeps a nested loop wholly in one paragraph inline", () => {
      const filled = fill(
        table(row(cell(t("{{#each rows}}[{{#each tags}}{{.}}{{/each}}]{{/each}}")))),
        data
      );
      // Kept whole and filled per item, like any single-paragraph row loop.
      expect(grid(filled)).toEqual([["[ab][c]"], ["[ab][c]"]]);
    });

    it("repeats the row when a nested inline loop sits in another cell", () => {
      const filled = fill(
        table(
          row(
            cell(t("{{#each rows}}")),
            cell(t("{{#each tags}}{{.}}{{/each}}")),
            cell(t("{{/each}}"))
          )
        ),
        data
      );
      expect(grid(filled)).toEqual([
        ["", "ab", ""],
        ["", "c", ""]
      ]);
    });

    it("closes the row loop with its own tag when a sibling loop follows in a later cell", () => {
      const filled = fill(
        table(
          row(cell(t("{{#each xs}}{{.}}")), cell(t("{{/each}}")), cell(t("{{#each ys}}y{{/each}}")))
        ),
        { xs: ["a", "b"], ys: [1, 2] }
      );
      // The row repeats once per `xs` item; the sibling `ys` loop stays whole and fills
      // inline in every copy rather than losing its close to the row loop.
      expect(grid(filled)).toEqual([
        ["a", "", "yy"],
        ["b", "", "yy"]
      ]);
    });
  });
});
