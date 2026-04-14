import { XmlWriter, StdDocAttributes } from "@xml/writer";
import { describe, it, expect } from "vitest";

describe("XmlWriter", () => {
  it("Writes simple XML doc", () => {
    const xmlStream = new XmlWriter();

    xmlStream.openXml(StdDocAttributes);
    xmlStream.openNode("root", {
      attr1: "attr1-value",
      attr2: "attr2-value"
    });
    xmlStream.openNode("l1");
    xmlStream.openNode("l2");
    xmlStream.addAttribute("l2a1", "v1");
    xmlStream.addAttribute("l2a2", "v2");
    xmlStream.closeNode();
    xmlStream.closeNode();
    xmlStream.closeNode();
    expect(xmlStream.xml).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<root attr1="attr1-value" attr2="attr2-value"><l1><l2 l2a1="v1" l2a2="v2"/></l1></root>'
    );
  });

  it("Writes text in XML doc", () => {
    const xmlStream = new XmlWriter();

    xmlStream.openNode("root");
    xmlStream.openNode("l1");
    xmlStream.openNode("l2");
    xmlStream.addAttribute("l2a1", "v1");
    xmlStream.writeText("Hello, World!");
    xmlStream.closeNode();
    xmlStream.openNode("l2");
    xmlStream.addAttribute("l2a1", "v2");
    xmlStream.writeText("See ya later, Alligator!");
    xmlStream.closeNode();
    xmlStream.closeNode();
    xmlStream.closeNode();
    expect(xmlStream.xml).toBe(
      '<root><l1><l2 l2a1="v1">Hello, World!</l2><l2 l2a1="v2">See ya later, Alligator!</l2></l1></root>'
    );
  });
  it("text is escaped", () => {
    const xmlStream = new XmlWriter();

    xmlStream.openNode("root");
    xmlStream.openNode("l1");
    xmlStream.writeText("<escape this!>");
    xmlStream.closeNode();
    xmlStream.closeNode();
    expect(xmlStream.xml).toBe("<root><l1>&lt;escape this!&gt;</l1></root>");
  });
  it("attributes are escaped", () => {
    const xmlStream = new XmlWriter();

    xmlStream.openNode("root");
    xmlStream.openNode("l1");
    xmlStream.addAttribute("stuff", "this & that");
    xmlStream.openNode("l2", { foo: "<bar>" });
    xmlStream.closeNode();
    xmlStream.leafNode("l2", { quote: '"this"' });
    xmlStream.closeNode();
    xmlStream.closeNode();
    expect(xmlStream.xml).toBe(
      '<root><l1 stuff="this &amp; that"><l2 foo="&lt;bar&gt;"/><l2 quote="&quot;this&quot;"/></l1></root>'
    );
  });

  it("rolls back", () => {
    const xmlStream = new XmlWriter();

    xmlStream.openNode("root");
    xmlStream.addAttribute("in", "1");
    xmlStream.save();
    xmlStream.addAttribute("not", "1");
    xmlStream.openNode("invalid");
    xmlStream.rollback();
    xmlStream.addAttribute("also", "2");
    xmlStream.openNode("valid");
    xmlStream.closeNode();
    xmlStream.closeNode();
    expect(xmlStream.xml).toBe('<root in="1" also="2"><valid/></root>');
  });

  it("throws when adding attributes without an open node", () => {
    const xmlStream = new XmlWriter();

    expect(() => xmlStream.addAttribute("a", "1")).toThrow("no element is open");
    expect(() => xmlStream.addAttributes({ a: "1" })).toThrow("no element is open");
  });

  it("throws when adding attributes after closing a node", () => {
    const xmlStream = new XmlWriter();

    xmlStream.openNode("root");
    xmlStream.closeNode();

    expect(() => xmlStream.addAttribute("a", "1")).toThrow("no element is open");
  });
});
