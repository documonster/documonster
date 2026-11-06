import {
  attr,
  eachElementOf,
  findChildOf,
  findChildren,
  getAttrOf,
  getTextOf,
  listChildrenOf,
  textContent,
  walk
} from "@xml/dom";
import { encodeCDataSection, validateXmlName } from "@xml/encode";
import { Xml } from "@xml/index";
import { SaxParser, saxStream } from "@xml/sax";
import { XmlStreamWriter } from "@xml/stream-writer";
import { XmlWriter } from "@xml/writer";
import { describe, expect, it } from "vitest";

describe("Xml namespace", () => {
  it("publishes the renamed members", () => {
    // `encodeCData` keeps its published name, matching `writeCData` and `XmlCData`; the
    // `encodeCdata` rename was undone before it shipped.
    expect(Xml.encodeCData).toBe(encodeCDataSection);
    // The live names check their arguments; the deprecated forwards below keep the old forms.
    expect(Xml.getText).toBe(getTextOf);
    expect(Xml.getAttr).toBe(getAttrOf);
    expect(Xml.eachElement).toBe(eachElementOf);
    expect(Xml.validateName).toBe(validateXmlName);
    expect(Xml.feedSax).toBe(saxStream);
    expect(Xml.listChildren).toBe(listChildrenOf);
  });

  it("keeps the deprecated forwards", () => {
    expect(Xml.textContent).toBe(textContent);
    expect(Xml.attr).toBe(attr);
    expect(Xml.walk).toBe(walk);
    expect(Xml.validateXmlName).toBe(validateXmlName);
    expect(Xml.saxStream).toBe(saxStream);
    expect(Xml.SaxParser).toBe(SaxParser);
    expect(Xml.findChildren).toBe(findChildren);
  });

  it("Writer / StreamWriter are the public writer classes", () => {
    const w = new Xml.Writer();
    expect(w).toBeInstanceOf(Xml.Writer);
    expect(Xml.Writer).toBe(XmlWriter);
    w.leafNode("a", undefined, "x");
    expect(w.toString()).toBe("<a>x</a>");
    const chunks: string[] = [];
    const sw = new Xml.StreamWriter({ write: (c: string) => void chunks.push(c) });
    expect(sw).toBeInstanceOf(Xml.StreamWriter);
    expect(Xml.StreamWriter).toBe(XmlStreamWriter);
    sw.leafNode("b");
    expect(chunks.join("")).toBe("<b/>");
  });

  it("findChild / listChildren find children by name", () => {
    const doc = Xml.parse("<r><i n='1'/><j/><i n='2'/></r>");
    expect(Xml.findChild).toBe(findChildOf);
    expect(Xml.getAttr(Xml.findChild(doc.root, "i")!, "n")).toBe("1");
    expect(Xml.listChildren(doc.root, "i").map(el => Xml.getAttr(el, "n"))).toEqual(["1", "2"]);
    // Same results as the baseline name
    expect(Xml.listChildren(doc.root, "i")).toEqual(Xml.findChildren(doc.root, "i"));
    expect(Xml.findChild(doc.root, "missing")).toBeUndefined();
    expect("getChild" in Xml).toBe(false);
  });

  it("SaxParser is a parser feedSax can drive", async () => {
    const parser = new Xml.SaxParser();
    expect(parser).toBeInstanceOf(SaxParser);
    const names: string[] = [];
    parser.on("opentag", tag => names.push(tag.name));
    await Xml.feedSax(parser, ["<a><b/>", "</a>"]);
    expect(names).toEqual(["a", "b"]);
  });
});
