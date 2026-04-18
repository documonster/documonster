/**
 * Tests for ExternalLinkXform. Uses the standard xform test harness to
 * exercise both rendering (model → XML) and parsing (XML → model) round
 * trips. Focuses on the shapes that show up in real Excel output:
 *
 *   - sheetNames only (no cache)
 *   - cache with numbers, strings, booleans, blanks
 *   - numeric-valued cells (the common case)
 *   - mixed quoted / unquoted addresses
 */

import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { ExternalLinkXform } from "@excel/xlsx/xform/book/external-link-xform";
import { describe, expect, it } from "vitest";

const RENDER_XMLNS =
  'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const expectations = [
  {
    title: "ExternalLink — sheetNames only, no cache",
    create() {
      return new ExternalLinkXform();
    },
    preparedModel: {
      index: 1,
      target: "测试.xlsx",
      targetMode: "External",
      sheetNames: ["Sheet1"],
      cachedValues: {}
    },
    xml:
      `<externalLink ${RENDER_XMLNS}>` +
      `<externalBook r:id="rId1">` +
      `<sheetNames><sheetName val="Sheet1"/></sheetNames>` +
      `</externalBook>` +
      `</externalLink>`,
    tests: ["render"]
  },
  {
    title: "ExternalLink — cached numeric value",
    create() {
      return new ExternalLinkXform();
    },
    preparedModel: {
      index: 1,
      target: "src.xlsx",
      targetMode: "External",
      sheetNames: ["Sheet1"],
      cachedValues: {
        Sheet1: { A1: 42 }
      }
    },
    xml:
      `<externalLink ${RENDER_XMLNS}>` +
      `<externalBook r:id="rId1">` +
      `<sheetNames><sheetName val="Sheet1"/></sheetNames>` +
      `<sheetDataSet>` +
      `<sheetData sheetId="0">` +
      `<row r="1"><cell r="A1"><v>42</v></cell></row>` +
      `</sheetData>` +
      `</sheetDataSet>` +
      `</externalBook>` +
      `</externalLink>`,
    tests: ["render"]
  },
  {
    title: "ExternalLink — mixed value types",
    create() {
      return new ExternalLinkXform();
    },
    preparedModel: {
      index: 1,
      target: "src.xlsx",
      targetMode: "External",
      sheetNames: ["S"],
      cachedValues: {
        S: {
          A1: 1,
          B1: "hello",
          C1: true,
          D1: false,
          E1: null
        }
      }
    },
    xml:
      `<externalLink ${RENDER_XMLNS}>` +
      `<externalBook r:id="rId1">` +
      `<sheetNames><sheetName val="S"/></sheetNames>` +
      `<sheetDataSet>` +
      `<sheetData sheetId="0">` +
      `<row r="1">` +
      `<cell r="A1"><v>1</v></cell>` +
      `<cell r="B1" t="str"><v>hello</v></cell>` +
      `<cell r="C1" t="b"><v>1</v></cell>` +
      `<cell r="D1" t="b"><v>0</v></cell>` +
      `<cell r="E1"/>` +
      `</row>` +
      `</sheetData>` +
      `</sheetDataSet>` +
      `</externalBook>` +
      `</externalLink>`,
    tests: ["render"]
  },
  {
    title: "ExternalLink — parse cached numeric + string",
    create() {
      return new ExternalLinkXform();
    },
    xml:
      `<externalLink ${RENDER_XMLNS}>` +
      `<externalBook r:id="rId99">` +
      `<sheetNames><sheetName val="Data"/></sheetNames>` +
      `<sheetDataSet>` +
      `<sheetData sheetId="0">` +
      `<row r="1">` +
      `<cell r="A1"><v>3.14</v></cell>` +
      `<cell r="B1" t="str"><v>text value</v></cell>` +
      `</row>` +
      `</sheetData>` +
      `</sheetDataSet>` +
      `</externalBook>` +
      `</externalLink>`,
    parsedModel: {
      externalBookRId: "rId99",
      sheetNames: ["Data"],
      cachedValues: {
        Data: { A1: 3.14, B1: "text value" }
      }
    },
    tests: ["parse"]
  },
  {
    title: "ExternalLink — parse boolean and error values",
    create() {
      return new ExternalLinkXform();
    },
    xml:
      `<externalLink ${RENDER_XMLNS}>` +
      `<externalBook r:id="rId1">` +
      `<sheetNames><sheetName val="Sheet1"/></sheetNames>` +
      `<sheetDataSet>` +
      `<sheetData sheetId="0">` +
      `<row r="1">` +
      `<cell r="A1" t="b"><v>1</v></cell>` +
      `<cell r="B1" t="b"><v>0</v></cell>` +
      `<cell r="C1" t="e"><v>#DIV/0!</v></cell>` +
      `</row>` +
      `</sheetData>` +
      `</sheetDataSet>` +
      `</externalBook>` +
      `</externalLink>`,
    parsedModel: {
      externalBookRId: "rId1",
      sheetNames: ["Sheet1"],
      cachedValues: {
        Sheet1: { A1: true, B1: false, C1: "#DIV/0!" }
      }
    },
    tests: ["parse"]
  },
  {
    title: "ExternalLink — parse document with no cache",
    create() {
      return new ExternalLinkXform();
    },
    xml:
      `<externalLink ${RENDER_XMLNS}>` +
      `<externalBook r:id="rId1">` +
      `<sheetNames>` +
      `<sheetName val="Sheet1"/>` +
      `<sheetName val="Sheet2"/>` +
      `</sheetNames>` +
      `</externalBook>` +
      `</externalLink>`,
    parsedModel: {
      externalBookRId: "rId1",
      sheetNames: ["Sheet1", "Sheet2"],
      cachedValues: {}
    },
    tests: ["parse"]
  },
  {
    title: "ExternalLink — render error cached values with t='e'",
    create() {
      return new ExternalLinkXform();
    },
    preparedModel: {
      index: 1,
      target: "src.xlsx",
      targetMode: "External",
      sheetNames: ["S"],
      cachedValues: {
        S: {
          A1: "#DIV/0!",
          B1: "#REF!"
        }
      }
    },
    xml:
      `<externalLink ${RENDER_XMLNS}>` +
      `<externalBook r:id="rId1">` +
      `<sheetNames><sheetName val="S"/></sheetNames>` +
      `<sheetDataSet>` +
      `<sheetData sheetId="0">` +
      `<row r="1">` +
      `<cell r="A1" t="e"><v>#DIV/0!</v></cell>` +
      `<cell r="B1" t="e"><v>#REF!</v></cell>` +
      `</row>` +
      `</sheetData>` +
      `</sheetDataSet>` +
      `</externalBook>` +
      `</externalLink>`,
    tests: ["render"]
  }
];

describe("ExternalLinkXform", () => {
  testXformHelper(expectations);

  it("maps cache sheetId back to the correct sheet name (multi-sheet)", async () => {
    // Ensures `sheetId="1"` resolves to sheetNames[1], not always sheetNames[0].
    const xml =
      `<externalLink ${RENDER_XMLNS}>` +
      `<externalBook r:id="rId1">` +
      `<sheetNames>` +
      `<sheetName val="First"/>` +
      `<sheetName val="Second"/>` +
      `</sheetNames>` +
      `<sheetDataSet>` +
      `<sheetData sheetId="1">` +
      `<row r="2"><cell r="B2"><v>99</v></cell></row>` +
      `</sheetData>` +
      `</sheetDataSet>` +
      `</externalBook>` +
      `</externalLink>`;

    const xform = new ExternalLinkXform();
    const { parseSax } = await import("@xml/sax");
    const { PassThrough } = await import("@stream");
    const stream = new PassThrough();
    stream.write(xml);
    stream.end();
    const model = await xform.parse(parseSax(stream));

    expect(model).toBeDefined();
    expect(model!.cachedValues).toEqual({ Second: { B2: 99 } });
  });
});
