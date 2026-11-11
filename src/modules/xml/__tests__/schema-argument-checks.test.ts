/** Refusals that changed when the `Xml` argument checks moved to `@utils/schema`. */
import { Xml } from "@xml/index";
import { describe, expect, it } from "vitest";

function refusal(call: () => unknown): { name: string; field: string; message: string } {
  try {
    call();
  } catch (err) {
    const e = err as { name: string; field: string; message: string };
    return { name: e.name, field: e.field, message: e.message };
  }
  throw new Error("expected a refusal");
}

describe("Xml argument checks declared as schemas", () => {
  it("describes a refused text by its value", () => {
    expect(refusal(() => Xml.encode(5 as never))).toEqual({
      name: "XmlArgumentError",
      field: "text",
      message: "Xml.encode: text must be a string, got 5"
    });
  });

  it("says what a wrong node, parser or fileName was", async () => {
    expect(refusal(() => Xml.getText(5 as never)).message).toBe(
      "Xml.getText: node must be an XmlNode (pass doc.root for a document), got 5"
    );
    expect(refusal(() => new Xml.SaxParser({ fileName: 5 as never })).message).toBe(
      "Xml.SaxParser: options.fileName must be a string, got 5"
    );
    await expect(Xml.feedSax(5 as never, [])).rejects.toThrow(
      "Xml.feedSax: parser must be an Xml.SaxParser, got 5"
    );
  });

  it("describes a refused isArray, attributePrefix or query path by its value", () => {
    expect(refusal(() => Xml.parseToObject("<a/>", { isArray: 5 as never })).message).toBe(
      "Xml.parseToObject: options.isArray must be a function, got 5"
    );
    const root = Xml.parse("<a/>").root;
    expect(refusal(() => Xml.toPlainObject(root, { attributePrefix: 5 as never })).message).toBe(
      "Xml.toPlainObject: options.attributePrefix must be a string, got 5"
    );
    expect(refusal(() => Xml.query(root, 5 as never)).message).toBe(
      "Xml.query: path must be a string, got 5"
    );
  });
});

describe("DOM helpers refuse a node missing what they read", () => {
  it("getAttr needs attributes", () => {
    const r = refusal(() => Xml.getAttr({ type: "element", children: [] } as never, "id"));
    expect(r.name).toBe("XmlArgumentError");
    expect(r.message).toBe("Xml.getAttr: element.attributes must be an object, got undefined");
  });

  it("getText needs a text node's value to be a string, at any depth", () => {
    expect(refusal(() => Xml.getText({ type: "text", value: 42 } as never)).message).toBe(
      "Xml.getText: node.value must be a string, got 42"
    );
    const deep = {
      type: "element",
      name: "a",
      attributes: {},
      children: [{ type: "text", value: 42 }]
    };
    expect(refusal(() => Xml.getText(deep as never)).name).toBe("XmlArgumentError");
    expect(refusal(() => Xml.toPlainObject(deep as never)).name).toBe("XmlArgumentError");
  });
});

describe("Xml.toPlainObject reads its options once", () => {
  it("uses the value it checked", () => {
    let reads = 0;
    const options = {
      get textKey() {
        return reads++ === 0 ? "v" : (42 as never);
      }
    };
    const root = Xml.parse("<r a='1'>x</r>").root;
    expect(Xml.toPlainObject(root, options)).toEqual({ r: { "@_a": "1", v: "x" } });
    expect(reads).toBe(1);
  });
});
