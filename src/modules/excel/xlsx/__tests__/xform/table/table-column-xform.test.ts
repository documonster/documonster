import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { ListXform } from "@excel/xlsx/xform/list-xform";
import { TableColumnXform } from "@excel/xlsx/xform/table/table-column-xform";
import { PassThrough } from "@stream";
import { parseSax } from "@xml/sax";
import { describe, it, expect } from "vitest";

const expectations = [
  {
    title: "label",
    create() {
      return new TableColumnXform();
    },
    preparedModel: { id: 1, name: "Foo", totalsRowLabel: "Bar" },
    xml: '<tableColumn id="1" name="Foo" totalsRowLabel="Bar" />',
    parsedModel: { name: "Foo", totalsRowLabel: "Bar" },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "function",
    create() {
      return new TableColumnXform();
    },
    preparedModel: { id: 1, name: "Foo", totalsRowFunction: "Baz" },
    xml: '<tableColumn id="1" name="Foo" totalsRowFunction="Baz" />',
    parsedModel: { name: "Foo", totalsRowFunction: "Baz" },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "OOXML escape in name (_x000a_ → newline)",
    create() {
      return new TableColumnXform();
    },
    preparedModel: { id: 1, name: "Col3\nnew line" },
    xml: '<tableColumn id="1" name="Col3_x000A_new line" />',
    parsedModel: { name: "Col3\nnew line" },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "OOXML escape in name (lowercase _x000a_ → newline)",
    create() {
      return new TableColumnXform();
    },
    xml: '<tableColumn id="1" name="Col3_x000a_new line" />',
    parsedModel: { name: "Col3\nnew line" },
    tests: ["parse"]
  },
  {
    title: "OOXML escape in totalsRowLabel",
    create() {
      return new TableColumnXform();
    },
    preparedModel: { id: 1, name: "Totals", totalsRowLabel: "Sum\ntotal" },
    xml: '<tableColumn id="1" name="Totals" totalsRowLabel="Sum_x000A_total" />',
    parsedModel: { name: "Totals", totalsRowLabel: "Sum\ntotal" },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "with calculatedColumnFormula child element",
    create() {
      return new TableColumnXform();
    },
    preparedModel: { id: 2, name: "Calc", calculatedColumnFormula: "[Col1]*2" },
    xml: '<tableColumn id="2" name="Calc"><calculatedColumnFormula>[Col1]*2</calculatedColumnFormula></tableColumn>',
    parsedModel: { name: "Calc", calculatedColumnFormula: "[Col1]*2" },
    tests: ["render", "renderIn", "parse", "parseIn"]
  },
  {
    title: "with totalsRowFormula child element",
    create() {
      return new TableColumnXform();
    },
    preparedModel: {
      id: 3,
      name: "Val",
      totalsRowFunction: "custom",
      totalsRowFormula: "SUM([Val])"
    },
    xml: '<tableColumn id="3" name="Val" totalsRowFunction="custom"><totalsRowFormula>SUM([Val])</totalsRowFormula></tableColumn>',
    parsedModel: { name: "Val", totalsRowFunction: "custom", totalsRowFormula: "SUM([Val])" },
    tests: ["render", "renderIn", "parse", "parseIn"]
  },
  {
    title: "with both calculatedColumnFormula and totalsRowFormula",
    create() {
      return new TableColumnXform();
    },
    preparedModel: {
      id: 4,
      name: "Both",
      totalsRowFunction: "custom",
      calculatedColumnFormula: "[A]*2",
      totalsRowFormula: "MY_SUM([Both])"
    },
    xml: '<tableColumn id="4" name="Both" totalsRowFunction="custom"><calculatedColumnFormula>[A]*2</calculatedColumnFormula><totalsRowFormula>MY_SUM([Both])</totalsRowFormula></tableColumn>',
    parsedModel: {
      name: "Both",
      totalsRowFunction: "custom",
      calculatedColumnFormula: "[A]*2",
      totalsRowFormula: "MY_SUM([Both])"
    },
    tests: ["render", "parse"]
  }
];

describe("TableColumnXform", () => {
  testXformHelper(expectations);

  it("parses all columns when some have child elements", async () => {
    const xml =
      '<tableColumns count="3">' +
      '<tableColumn id="1" name="Col1"/>' +
      '<tableColumn id="2" name="Calc"><calculatedColumnFormula>[Col1]*2</calculatedColumnFormula></tableColumn>' +
      '<tableColumn id="3" name="Col3"/>' +
      "</tableColumns>";

    const listXform = new ListXform({
      tag: "tableColumns",
      count: true,
      empty: true,
      childXform: new TableColumnXform()
    });

    const stream = new PassThrough();
    stream.write(xml);
    stream.end();
    const model = await listXform.parse(parseSax(stream));

    expect(model).toHaveLength(3);
    expect(model![0].name).toBe("Col1");
    expect(model![1].name).toBe("Calc");
    expect(model![1].calculatedColumnFormula).toBe("[Col1]*2");
    expect(model![2].name).toBe("Col3");
  });
});
