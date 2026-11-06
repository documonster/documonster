import { parseXformEvents } from "@excel/xlsx/xform/parse-xform";
/**
 * Malformed parts that used to escape the reader or the writer as a raw `TypeError`. Each was found by fuzzing a real
 * workbook; the case here is the smallest input reaching the same line.
 */
import { FontXform } from "@excel/xlsx/xform/style/font-xform";
import { StylesXform } from "@excel/xlsx/xform/style/styles-xform";
import { parseSax } from "@xml/sax";
import { describe, expect, it } from "vitest";

async function parse(xform: { parseOpen: unknown }, xml: string): Promise<unknown> {
  async function* chunks() {
    yield xml;
  }
  return parseXformEvents(xform as never, parseSax(chunks()));
}

describe("malformed style parts", () => {
  it("reads a font whose child holds an unknown element", async () => {
    // The unknown element's close reached the font while its `<b>` parser was active, and the font looked the
    // closing name up in its child map — `undefined.prop`.
    const model = await parse(new FontXform(), '<font><b><x/></b><sz val="11"/></font>');
    expect(model).toMatchObject({ bold: true, size: 11 });
  });

  it("answers no style for a styles part with no cellXfs", () => {
    const styles = new StylesXform();
    styles.model = {} as never;
    expect(styles.getStyleModel(1)).toBeNull();
  });
});
