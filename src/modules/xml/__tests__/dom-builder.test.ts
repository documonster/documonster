/**
 * `createDomBuilder` is `parseXml` fed a piece at a time: wherever the pieces end — inside a tag, a name, an
 * attribute value, a reference or a surrogate pair — the tree is the one a single `parseXml` call builds.
 */
import { createDomBuilder, parseXml } from "@xml/dom";
import { XmlParseError } from "@xml/errors";
import { describe, expect, it } from "vitest";

const XML =
  '<?xml version="1.0" encoding="UTF-8"?>\n<!-- lead -->\n' +
  '<w:root xmlns:w="urn:w" a="1 &amp; 2"><w:p x="é😀">text &lt;here&gt; 数据</w:p>' +
  "<![CDATA[raw <b>]]><?pi body?><w:empty/><w:q>tail</w:q></w:root>\n<!-- trail -->";

function built(pieces: readonly string[], options?: Parameters<typeof parseXml>[1]) {
  const builder = createDomBuilder(options);
  for (const piece of pieces) {
    builder.write(piece);
  }
  return builder.close();
}

describe("createDomBuilder", () => {
  it("builds the tree parseXml builds, cut at every position", () => {
    const options = { comments: true, processingInstructions: true, cdataAsNodes: true };
    const whole = parseXml(XML, options);
    for (let at = 1; at < XML.length; at++) {
      expect(built([XML.slice(0, at), XML.slice(at)], options), `cut at ${at}`).toEqual(whole);
    }
  });

  it("builds the same tree from one-character pieces", () => {
    expect(built([...XML])).toEqual(parseXml(XML));
  });

  it.each([
    ["a mismatched close tag", ["<a><b>", "</a>"]],
    ["a second root", ["<a/>", "<b/>"]],
    ["no root", [" ", " "]]
  ])("refuses %s as parseXml does", (_name, pieces) => {
    let expected: unknown;
    try {
      parseXml(pieces.join(""));
    } catch (error) {
      expected = error;
    }
    expect(expected).toBeInstanceOf(XmlParseError);
    expect(() => built(pieces)).toThrow((expected as Error).message);
  });
});
