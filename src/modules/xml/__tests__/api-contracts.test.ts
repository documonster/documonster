/**
 * Contracts of the `Xml` entry points that the API review found broken: a writer whose
 * rollback lost the names of elements closed after `save()`, names that opened a comment or
 * a processing instruction, query paths that searched the wrong depth or returned a node
 * twice, safety limits switched off by `NaN`, attributes dropped for their name, and inputs
 * that surfaced as a `TypeError` instead of an `XmlError`.
 */
import { Xml, XmlArgumentError, XmlError, XmlParseError, XmlWriteError } from "@xml/index";
import { describe, expect, it } from "vitest";

const names = (els: { name: string; attributes: Record<string, string> }[]) =>
  els.map(e => e.name + (e.attributes.id ? `#${e.attributes.id}` : ""));

describe("Xml.Writer#rollback", () => {
  it("restores elements closed after snapshot()", () => {
    const w = new Xml.Writer();
    w.openNode("a");
    w.snapshot();
    w.closeNode();
    w.rollback();
    w.leafNode("b");
    w.closeNode();
    expect(w.toString()).toBe("<a><b/></a>");
  });

  it("restores a nested stack", () => {
    const w = new Xml.Writer();
    w.openNode("a");
    w.openNode("b");
    w.writeText("x");
    w.snapshot();
    w.closeNode();
    w.closeNode();
    w.rollback();
    w.closeAll();
    expect(w.toString()).toBe("<a><b>x</b></a>");
  });
});

describe("Xml.Writer#save (deprecated)", () => {
  it("takes the same snapshot as snapshot()", () => {
    const viaSave = new Xml.Writer();
    const viaSnapshot = new Xml.Writer();
    for (const [w, mark] of [
      [viaSave, () => viaSave.save()],
      [viaSnapshot, () => viaSnapshot.snapshot()]
    ] as const) {
      w.openNode("a");
      mark();
      w.leafNode("dropped");
      w.rollback();
      mark();
      w.leafNode("kept");
      w.commit();
      w.closeAll();
    }
    expect(viaSave.toString()).toBe("<a><kept/></a>");
    expect(viaSave.toString()).toBe(viaSnapshot.toString());
  });
});

describe("Xml.Writer#cursor", () => {
  it("never goes down on a write, including one that consolidates", () => {
    const w = new Xml.Writer();
    w.openNode("root");
    w.openNode("sec");
    for (let i = 0; i < 9997; i++) {
      w.leafNode("c");
    }
    const before = w.cursor;
    w.closeNode();
    expect(w.cursor).toBeGreaterThan(before);
  });
});

describe("Xml.validateName and the writers refuse what is not an XML name", () => {
  it.each(["!--x", "?pi", "![CDATA[x", "x\u0000y", ":", ":a", "a:", "a:b:c", "a;b", "a(b)", "a]b"])(
    "%j",
    name => {
      expect(() => Xml.validateName(name)).toThrow(XmlError);
      expect(() => new Xml.Writer().openNode(name)).toThrow(XmlError);
      expect(() => new Xml.Writer().leafNode("a", { [name]: "1" })).toThrow(XmlError);
    }
  );

  it.each(["a", "x:row", "_a", "é", "a-b.c", "a\u00B7b", "\u4E2D\u6587", "mc:AlternateContent"])(
    "accepts %j",
    name => {
      expect(() => Xml.validateName(name)).not.toThrow();
    }
  );

  it("refuses a non-string with XmlError", () => {
    expect(() => Xml.validateName(42 as never)).toThrow(XmlError);
    expect(() => Xml.validateCommentText(42 as never)).toThrow(XmlError);
  });
});

describe("the declaration is not written into an open start tag", () => {
  it("in Xml.Writer", () => {
    const w = new Xml.Writer();
    w.openNode("a");
    expect(() => w.openXml()).toThrow(XmlWriteError);
    w.closeNode();
    expect(w.toString()).toBe("<a/>");
  });

  it("in Xml.StreamWriter", () => {
    let out = "";
    const w = new Xml.StreamWriter({ write: c => void (out += c) });
    w.openNode("a");
    expect(() => w.openXml()).toThrow(XmlWriteError);
    w.closeNode();
    expect(out).toBe("<a/>");
  });
});

describe("text round-trips through the writer and the parser", () => {
  it("keeps a carriage return", () => {
    const w = new Xml.Writer();
    w.leafNode("r", undefined, "a\rb\r\nc");
    expect(Xml.getText(Xml.parse(w.toString()).root)).toBe("a\rb\r\nc");
    expect(Xml.decode(Xml.encode("a\rb"))).toBe("a\rb");
  });
});

describe("Xml.query / queryAll", () => {
  const doc = Xml.parse(
    '<r><a id="1"><b id="2"><c id="3"/><b id="4"><c id="5"/></b></b></a><a id="6"><c id="7"/></a></r>'
  );

  it("returns each node once, in document order, for overlapping // steps", () => {
    expect(names(Xml.queryAll(doc.root, "//b//c"))).toEqual(["c#3", "c#5"]);
    expect(names(Xml.queryAll(doc.root, "//b"))).toEqual(["b#2", "b#4"]);
  });

  it("refuses an absolute path instead of searching every depth", () => {
    expect(() => Xml.queryAll(doc.root, "/c")).toThrow(XmlError);
  });

  it.each(["", "a/", "a//", "a///b", "a[@id]", "a[last()]", ".", "..", "a/../b"])(
    "refuses the unsupported path %j",
    path => {
      expect(() => Xml.queryAll(doc.root, path)).toThrow(XmlError);
    }
  );

  it("accepts a double-quoted attribute value", () => {
    expect(names(Xml.queryAll(doc.root, 'a/b[@id="2"]'))).toEqual(["b#2"]);
  });

  it("keeps the 0-based index", () => {
    expect(names(Xml.queryAll(doc.root, "a[1]"))).toEqual(["a#6"]);
  });

  it("refuses arguments that are not an element and a path", () => {
    expect(() => Xml.query(undefined as never, "a")).toThrow(XmlError);
    expect(() => Xml.queryAll(doc.root, 42 as never)).toThrow(XmlError);
  });
});

describe("Xml.parse", () => {
  it("keeps every attribute, whatever its name", () => {
    const { attributes } = Xml.parse('<a constructor="1" __proto__="2" prototype="3" x="4"/>').root;
    expect(Object.keys(attributes)).toEqual(["constructor", "__proto__", "prototype", "x"]);
    expect(({} as Record<string, unknown>).x).toBeUndefined();
  });

  it("reports comments after the root as the epilogue", () => {
    const doc = Xml.parse("<!--a--><r/><!--b-->", { comments: true });
    expect(doc.prologue.map(n => (n as { value: string }).value)).toEqual(["a"]);
    expect(doc.epilogue?.map(n => (n as { value: string }).value)).toEqual(["b"]);
  });

  it.each([
    ["maxDepth NaN", { maxDepth: NaN }],
    ["maxDepth -1", { maxDepth: -1 }],
    ["maxDepth as a string", { maxDepth: "10" }],
    ["maxEntityExpansions NaN", { maxEntityExpansions: NaN }],
    ["an unknown invalidCharHandling", { invalidCharHandling: "bogus" }],
    ["fragment as a string", { fragment: "yes" }]
  ])("refuses %s instead of switching the limit off", (_label, options) => {
    expect(() => Xml.parse("<a/>", options as never)).toThrow(XmlError);
    expect(() => Xml.parseToObject("<a/>", options as never)).toThrow(XmlError);
    expect(() => new Xml.SaxParser(options as never)).toThrow(XmlError);
  });

  it("refuses a parse-only option of the wrong type", () => {
    expect(() => Xml.parse("<a/>", { comments: "yes" as never })).toThrow(XmlError);
  });

  it("still accepts 0 as the documented way to switch a limit off", () => {
    const deep = "<a>".repeat(300) + "</a>".repeat(300);
    expect(Xml.parse(deep, { maxDepth: 0 }).root.name).toBe("a");
  });

  it.each([undefined, null, 42, new Uint8Array(1)])("refuses the input %s", input => {
    expect(() => Xml.parse(input as never)).toThrow(XmlError);
    expect(() => Xml.parseToObject(input as never)).toThrow(XmlError);
  });
});

describe("argument checks of the other entry points", () => {
  it("encode / decode / encodeAttr / encodeCData", () => {
    expect(() => Xml.encode(42 as never)).toThrow(XmlError);
    expect(() => Xml.encode(null as never)).toThrow(XmlError);
    expect(() => Xml.decode(42 as never)).toThrow(XmlError);
    expect(() => Xml.encodeAttr(undefined as never)).toThrow(XmlError);
    expect(() => Xml.encodeCData({} as never)).toThrow(XmlError);
    expect(Xml.encodeCData("a]]>b")).toBe("<![CDATA[a]]]]><![CDATA[>b]]>");
  });

  it("DOM helpers", () => {
    const root = Xml.parse("<r><a n='1'/></r>").root;
    expect(() => Xml.findChild(undefined as never, "a")).toThrow(XmlError);
    expect(() => Xml.listChildren(root, 42 as never)).toThrow(XmlError);
    expect(() => Xml.getAttr({} as never, "n")).toThrow(XmlError);
    expect(() => Xml.getText(null as never)).toThrow(XmlError);
    expect(() => Xml.eachElement(root, 42 as never)).toThrow(XmlError);
    expect(() => Xml.toPlainObject(root, 42 as never)).toThrow(XmlError);
    expect(() => Xml.toPlainObject(root, { isArray: true as never })).toThrow(XmlError);
  });

  it("StreamWriter target", () => {
    expect(() => new Xml.StreamWriter(null as never)).toThrow(XmlError);
    expect(() => new Xml.StreamWriter({} as never)).toThrow(XmlError);
  });

  it("SaxParser handlers and chunks", () => {
    const p = new Xml.SaxParser();
    expect(() => p.on("opentag", 42 as never)).toThrow(XmlError);
    expect(() => p.on("bogus" as never, () => {})).toThrow(XmlError);
    expect(() => p.write(new Uint8Array(1) as never)).toThrow(XmlError);
  });

  it("parseSax / feedSax sources and bytes", async () => {
    await expect(Xml.feedSax({} as never, ["<a/>"])).rejects.toThrow(XmlError);
    await expect(Xml.feedSax(new Xml.SaxParser(), 42 as never)).rejects.toThrow(XmlError);
    const read = async () => {
      for await (const _ of Xml.parseSax([new Uint8Array([0xff, 0xfe, 0x3c, 0])])) {
        // drain
      }
    };
    await expect(read()).rejects.toThrow(XmlParseError);
  });
});

describe("r3: plain objects have no prototype (TX-S1)", () => {
  it("Xml.toPlainObject's wrapper, like its nested objects and parseToObject", () => {
    const obj = Xml.toPlainObject(Xml.parse("<r><a>1</a></r>").root);
    expect(Object.getPrototypeOf(obj)).toBeNull();
    expect("constructor" in obj).toBe(false);
    expect(Object.getPrototypeOf(Xml.parseToObject("<r/>"))).toBeNull();
  });
});

describe("r3: a refused comment names the rule it broke (TX-D1)", () => {
  it.each([
    ["a--b", /contains "--"/],
    ["a-", /ends with "-"/],
    ["a\u0001-", /ends with "-" once the characters XML cannot hold are removed/]
  ])("%j", (text, message) => {
    expect(() => new Xml.Writer().writeComment(text)).toThrow(message);
    expect(() => Xml.validateCommentText(text.replace("\u0001", ""))).toThrow(XmlError);
  });
});

describe("X12 / X10: argument errors are XmlArgumentError with a field", () => {
  const root = Xml.parse("<r/>").root;
  it.each<[string, () => unknown, string]>([
    ["parse input", () => Xml.parse(42 as never), "input"],
    ["parse option", () => Xml.parse("<a/>", { maxDepth: NaN }), "options.maxDepth"],
    ["SaxParser option", () => new Xml.SaxParser({ fragment: "x" as never }), "options.fragment"],
    ["SaxParser#on", () => new Xml.SaxParser().on("text", 1 as never), "handler"],
    ["SaxParser#write", () => new Xml.SaxParser().write(1 as never), "chunk"],
    ["element handle", () => Xml.findChild({} as never, "a"), "element"],
    ["node handle", () => Xml.getText(null as never), "node"],
    ["query path", () => Xml.query(root, "/a"), "path"],
    [
      "toPlainObject option",
      () => Xml.toPlainObject(root, { textKey: 1 as never }),
      "options.textKey"
    ],
    ["writer name", () => new Xml.Writer().openNode("!x"), "name"],
    // `field` is the path from the parameter name: the attribute map is `attributes`.
    ["writer attribute value", () => new Xml.Writer().leafNode("a", { v: NaN }), "attributes.v"],
    ["writer attribute name", () => new Xml.Writer().leafNode("a", { "1x": "y" }), "attributes.1x"],
    [
      "writer addAttribute value",
      () => {
        const w = new Xml.Writer();
        w.openNode("a");
        w.addAttribute("v", NaN);
      },
      "value"
    ],
    [
      "writer duplicate attribute",
      () => {
        const w = new Xml.Writer();
        w.openNode("a", { v: "1" });
        w.addAttributes({ v: "2" });
      },
      "attributes.v"
    ],
    ["writer raw markup", () => new Xml.StreamWriter({ write() {} }).writeRaw(1 as never), "xml"],
    ["writer text", () => new Xml.Writer().leafNode("a", undefined, NaN), "text"],
    ["comment", () => new Xml.Writer().writeComment("a--b"), "text"],
    ["encode", () => Xml.encode(1 as never), "text"],
    ["StreamWriter target", () => new Xml.StreamWriter({} as never), "target"]
  ])("%s", (_label, fn, field) => {
    let error: unknown;
    try {
      fn();
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(XmlArgumentError);
    expect(error).toBeInstanceOf(XmlError);
    expect((error as XmlArgumentError).field).toBe(field);
    expect((error as XmlArgumentError).toJSON()).toMatchObject({ field });
  });

  it("checks the parser handed to feedSax", async () => {
    await expect(Xml.feedSax({} as never, ["<a/>"])).rejects.toMatchObject({ field: "parser" });
  });

  it("is not used for malformed input", () => {
    expect(() => Xml.parse("<a>")).toThrow(XmlParseError);
    expect(() => Xml.parse("<a>")).not.toThrow(XmlArgumentError);
  });
});
