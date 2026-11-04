/**
 * A part that is not well-formed fails with the parser's typed error, whatever the xforms were in the middle of.
 *
 * The SAX parser reports an error and keeps going, so it goes on to deliver close tags for elements whose open tags
 * it never delivered. An xform handed one of those reads state it never set up: `<font0<sz …/>` in `styles.xml`
 * reached `FontXform.parseClose` with no font and threw `TypeError: Cannot set properties of undefined`, which a
 * caller cannot tell from a bug in this library.
 */
import { parseXformStream } from "@excel/xlsx/xform/parse-xform";
import { StylesXform } from "@excel/xlsx/xform/style/styles-xform";
import { XmlParseError } from "@xml/errors";
import { describe, expect, it } from "vitest";

const MALFORMED_STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<fonts count="1"><font0<sz val="11"/><color theme="1"/><name val="Aptos Narrow"/></font></fonts>' +
  "</styleSheet>";

describe("parseXformStream on a part that is not well-formed", () => {
  it("rejects with XmlParseError and runs no xform handler after the error", async () => {
    const xform = new StylesXform();
    await expect(parseXformStream(xform, [MALFORMED_STYLES])).rejects.toBeInstanceOf(XmlParseError);
  });
});
