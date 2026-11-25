import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { parseXformEvents } from "@excel/xlsx/xform/parse-xform";
import { SharedStringXform } from "@excel/xlsx/xform/strings/shared-string-xform";
import { PassThrough } from "@stream";
import { parseSax } from "@xml/sax";
import { describe, expect, it } from "vitest";

const expectations = [
  {
    title: "text",
    create() {
      return new SharedStringXform();
    },
    preparedModel: "Hello, World!",
    xml: "<si><t>Hello, World!</t></si>",
    parsedModel: "Hello, World!",
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "rich text",
    create() {
      return new SharedStringXform();
    },
    preparedModel: {
      richText: [
        {
          font: {
            size: 11,
            bold: true,
            color: { theme: 1 },
            name: "Calibri",
            family: 2,
            scheme: "minor"
          },
          text: "Bold,"
        }
      ]
    },
    xml: '<si><r><rPr><b/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>Bold,</t></r></si>',
    parsedModel: {
      richText: [
        {
          font: {
            size: 11,
            bold: true,
            color: { theme: 1 },
            name: "Calibri",
            family: 2,
            scheme: "minor"
          },
          text: "Bold,"
        }
      ]
    },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "richText is empty",
    create() {
      return new SharedStringXform();
    },
    preparedModel: {
      richText: []
    },
    xml: "<si><t></t></si>",
    parsedModel: "",
    tests: ["parse"]
  },
  {
    title: "text + phonetic",
    create() {
      return new SharedStringXform();
    },
    preparedModel: {
      text: "Hello, World!",
      phoneticText: { text: "Helow woruld" }
    },
    xml: '<si><t>Hello, World!</t><rPh eb="0" sb="0"><t>Helow woruld</t></rPh></si>',
    parsedModel: "Hello, World!",
    tests: ["parse"]
  },
  {
    title: "Kanji + Katakana",
    create() {
      return new SharedStringXform();
    },
    preparedModel: {
      text: "役割",
      phoneticText: {
        sb: 0,
        eb: 2,
        text: "ヤクワリ",
        properties: { fontId: 1 }
      }
    },
    xml: '<si><t>役割</t><rPh sb="0" eb="2"><t>ヤクワリ</t></rPh><phoneticPr fontId="1" /></si>',
    parsedModel: "役割",
    tests: ["parse"]
  },
  {
    title: "text with newline",
    create() {
      return new SharedStringXform();
    },
    preparedModel: "Hello,\nWorld!",
    xml: '<si><t xml:space="preserve">Hello,\nWorld!</t></si>',
    parsedModel: "Hello,\nWorld!",
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "text with lowercase OOXML escape (e.g. _x000a_)",
    create() {
      return new SharedStringXform();
    },
    xml: "<si><t>Col3_x000a_new line</t></si>",
    parsedModel: "Col3\nnew line",
    tests: ["parse"]
  },
  {
    title: "text with uppercase OOXML escape (_x000A_)",
    create() {
      return new SharedStringXform();
    },
    xml: "<si><t>Hello_x000A_World</t></si>",
    parsedModel: "Hello\nWorld",
    tests: ["parse"]
  },
  {
    title: "text with escaped underscore (_x005F_x000D_) roundtrips to literal _x000D_",
    create() {
      return new SharedStringXform();
    },
    preparedModel: "_x000D_",
    xml: "<si><t>_x005F_x000D_</t></si>",
    parsedModel: "_x000D_",
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "text before runs, as the first run",
    create() {
      return new SharedStringXform();
    },
    xml: "<si><t>a</t><r><t>b</t></r></si>",
    parsedModel: { richText: [{ text: "a" }, { text: "b" }] },
    tests: ["parse"]
  },
  {
    title: "empty",
    create() {
      return new SharedStringXform();
    },
    xml: "<si/>",
    parsedModel: "",
    tests: ["parse"]
  }
];

describe("SharedStringXform", () => {
  testXformHelper(expectations);

  // `CT_Rst` is `t?`, `r*`: a second `<t>` used to replace the first, and one after a run replaced the runs.
  it.each([
    ["more than one <t>", "<si><t>a</t><t>b</t></si>"],
    ["a <t> after an <r>", "<si><r><t>a</t></r><t>b</t></si>"],
    ["a run with more than one <t>", "<si><r><t>a</t><t>b</t></r></si>"]
  ])("refuses a string with %s", async (detail, xml) => {
    const stream = new PassThrough();
    stream.write(xml);
    stream.end();
    await expect(parseXformEvents(new SharedStringXform(), parseSax(stream))).rejects.toThrow(
      `<si> has ${detail}`
    );
  });
});
