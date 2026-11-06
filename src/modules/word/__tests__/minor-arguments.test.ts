/**
 * The last entry points that answered a wrong argument with a plain
 * `DocxError` or an untyped `TypeError`: each now says which argument.
 */

import { describe, it, expect } from "vitest";

import { DocxArgumentError } from "../errors";
import { WordExcel } from "../excel";
import { Document, Font, Glossary, Styles, Template } from "../index";

function refused(label: string, run: () => unknown, field: string): void {
  let error: unknown;
  try {
    run();
  } catch (e) {
    error = e;
  }
  expect(error, label).toBeInstanceOf(DocxArgumentError);
  expect((error as DocxArgumentError).field, label).toBe(field);
}

describe("wrong arguments are named", () => {
  it("Font.obfuscate / deobfuscate", () => {
    const key = Font.generateObfuscationKey();
    refused("null data", () => Font.obfuscate(null as never, key), "data");
    refused("bad key", () => Font.deobfuscate(new Uint8Array(40), "bad key"), "fontKey");
    refused(
      "non-hex key",
      () => Font.obfuscate(new Uint8Array(40), "{zzzzzzzz-zzzz-zzzz-zzzz-zzzzzzzzzzzz}"),
      "fontKey"
    );
    const data = Uint8Array.from({ length: 40 }, (_, i) => i);
    expect(Font.deobfuscate(Font.obfuscate(data, key), key)).toEqual(data);
  });

  it("null where an object belongs", () => {
    const d = Document.create();
    refused("fillFromSource", () => Template.fillFromSource(d, null as never), "source");
    refused("Styles.match", () => Styles.match(null as never, "p" as never, {}), "map");
    refused("Glossary.findBlock", () => Glossary.findBlock(null as never, "x"), "glossary");
    refused("WordExcel.toChartSvg", () => WordExcel.toChartSvg(null as never), "chart");
    refused("WordExcel.toChartModel", () => WordExcel.toChartModel(null as never), "chart");
  });
});
