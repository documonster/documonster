/**
 * What the DOM helpers, the query engine and the writers do with a tree or a target a caller built:
 * only an element's own attributes count, a node without the fields read is refused, an element among
 * its own descendants is refused rather than walked for ever, and a target that calls back into its
 * writer is refused rather than given what was already written.
 */
import { Xml, XmlArgumentError, XmlWriteError } from "@xml/index";
import type { XmlElement, XmlNode } from "@xml/index";
import { describe, expect, it } from "vitest";

function element(name: string, children: XmlNode[] = [], attributes = {}): XmlElement {
  return { type: "element", name, attributes, children };
}

describe("an element's attributes are its own", () => {
  it("in Xml.getAttr, Xml.query and the writers", () => {
    const inherited = element("a", [], Object.create({ id: "ghost" }) as Record<string, string>);
    const root = element("r", [inherited]);
    expect(Xml.getAttr(root, "constructor")).toBeUndefined();
    expect(Xml.getAttr(inherited, "id")).toBeUndefined();
    expect(Xml.queryAll(root, "a[@id='ghost']")).toEqual([]);
    const attributes = Object.assign(Object.create({ ghost: "p" }) as object, { own: "y" });
    const writer = new Xml.Writer();
    writer.leafNode("r", attributes);
    expect(writer.toString()).toBe('<r own="y"/>');
    // An own attribute named like an object's property is an attribute.
    expect(Xml.getAttr(element("r", [], { constructor: "c" }), "constructor")).toBe("c");
  });
});

describe("a tree a caller built is checked where it is read", () => {
  it.each<[string, () => unknown, string]>([
    ["Xml.findChild", () => Xml.findChild(element("r", [null as never]), "a"), "element"],
    ["Xml.listChildren", () => Xml.listChildren(element("r", [null as never]), "a"), "element"],
    ["Xml.query", () => Xml.query(element("r", [null as never]), "a"), "element"],
    [
      "Xml.query, attributes",
      () =>
        Xml.query(
          element("r", [{ type: "element", name: "a", children: [] } as never]),
          "a[@id='1']"
        ),
      "element"
    ],
    ["Xml.getAttr, a value", () => Xml.getAttr(element("r", [], { id: 42 }), "id"), "element"],
    ["Xml.toPlainObject, a value", () => Xml.toPlainObject(element("r", [], { id: {} })), "element"]
  ])("%s", (_label, call, field) => {
    expect(call).toThrow(XmlArgumentError);
    // `field` is the path to what was refused, from the parameter.
    expect(call).toThrow(expect.objectContaining({ field: expect.stringMatching(`^${field}`) }));
  });
});

describe("an element among its own descendants", () => {
  const cyclic = (): XmlElement => {
    const root = Xml.parse("<r><a/></r>").root;
    (root.children[0] as XmlElement).children.push(root);
    return root;
  };

  it.each<[string, (root: XmlElement) => unknown]>([
    ["Xml.getText", root => Xml.getText(root)],
    ["Xml.eachElement", root => Xml.eachElement(root, () => {})],
    ["Xml.toPlainObject", root => Xml.toPlainObject(root)],
    ["Xml.queryAll", root => Xml.queryAll(root, "//x")],
    ["Xml.query", root => Xml.query(root, "a//x")]
  ])("is refused by %s", (_label, call) => {
    expect(() => call(cyclic())).toThrow(/contains itself/);
  });

  it("is told from one element shared by two parents", () => {
    const shared = Xml.parse("<s>t</s>").root;
    const root = element("r", [shared, shared]);
    expect(Xml.getText(root)).toBe("tt");
    expect(Xml.toPlainObject(root)).toEqual({ r: { s: ["t", "t"] } });
  });
});

describe("Xml.query and Xml.queryAll walk each element once", () => {
  /** How many times `children` is read in `root`'s tree while `call` runs. */
  function childReads(root: XmlElement, call: () => unknown): number {
    let reads = 0;
    const count = (el: XmlElement): void => {
      const children = el.children;
      Object.defineProperty(el, "children", {
        get() {
          reads++;
          return children;
        }
      });
      for (const child of children) {
        if (child.type === "element") {
          count(child);
        }
      }
    };
    count(root);
    call();
    return reads;
  }

  it("for overlapping // steps, in time that grows with the tree", () => {
    const reads = [100, 200, 400].map(depth => {
      const root = Xml.parse("<a>".repeat(depth) + "</a>".repeat(depth), { maxDepth: 0 }).root;
      let found: XmlElement[] = [];
      const n = childReads(root, () => {
        found = Xml.queryAll(root, "//a//a");
      });
      expect(found).toHaveLength(depth - 2);
      return n;
    });
    // Doubling the tree about doubles the work, where it quadrupled.
    expect(reads[2] / reads[1]).toBeLessThan(2.5);
    expect(reads[1] / reads[0]).toBeLessThan(2.5);
  });

  it("stops at the first match for Xml.query, and finds the same as Xml.queryAll", () => {
    const xml = `<r>${"<a><b/></a>".repeat(200)}</r>`;
    for (const path of ["//b", "a//b", "//a[0]", "//b[0]", "a/b"]) {
      const root = Xml.parse(xml).root;
      expect(Xml.query(root, path), path).toBe(Xml.queryAll(root, path)[0]);
    }
    const root = Xml.parse(xml).root;
    const reads = childReads(root, () => Xml.query(root, "a//b"));
    // The root's children are checked, then the first <a> walked: nothing like all 200.
    expect(reads).toBeLessThan(230);
  });
});

describe("Xml.StreamWriter and its target", () => {
  it("refuses a write from inside the target's write()", () => {
    let out = "";
    let refused: unknown;
    const writer: InstanceType<typeof Xml.StreamWriter> = new Xml.StreamWriter({
      write(chunk) {
        out += chunk;
        if (refused === undefined) {
          try {
            writer.leafNode("b");
          } catch (err) {
            refused = err;
          }
        }
      }
    });
    writer.openNode("r");
    writer.writeText("x");
    writer.closeAll();
    expect(refused).toBeInstanceOf(XmlWriteError);
    expect(out).toBe("<r>x</r>");
  });

  it("calls the target's write as a method, and stops once it fails", () => {
    const target: { written: string[]; write(chunk: string): void } = {
      written: [],
      write(chunk: string) {
        this.written.push(chunk);
      }
    };
    const writer = new Xml.StreamWriter(target);
    writer.leafNode("r");
    expect(target.written).toEqual(["<r/>"]);
    // A target that loses its write fails as one whose write throws: nothing more is written.
    (target as { write: unknown }).write = 42;
    expect(() => writer.leafNode("s")).toThrow(TypeError);
    expect(() => writer.leafNode("t")).toThrow(XmlWriteError);
  });
});

describe("the writers' attributes", () => {
  it("are checked against the start tag's in time that grows with their number", () => {
    const time = (count: number): number => {
      const writer = new Xml.Writer();
      writer.openNode("r");
      const started = performance.now();
      for (let i = 0; i < count; i++) {
        writer.addAttribute(`a${i}`, i);
      }
      writer.closeAll();
      return performance.now() - started;
    };
    // The fastest of a few runs: one run's time is the machine's load as much as the work.
    const best = (count: number): number => Math.min(time(count), time(count), time(count));
    best(2000);
    // 16× the attributes in far less than 16² the time: linear work is near 16×, quadratic near 256×.
    expect(best(32000)).toBeLessThan(Math.max(1, best(2000)) * 64);
  });

  it("refuse a repeated name after a rollback as before it", () => {
    for (const make of [
      () => new Xml.Writer(),
      () => new Xml.StreamWriter({ write() {} })
    ] as const) {
      const writer = make();
      writer.openNode("r", { a: '1 b="x"' });
      writer.addAttributes({ c: "3" });
      expect(() => writer.addAttribute("a", "2")).toThrow(XmlArgumentError);
      expect(() => writer.addAttributes({ c: "4" })).toThrow(XmlArgumentError);
      // A refused call adds no name.
      expect(() => writer.addAttributes({ d: "4", c: "5" })).toThrow(XmlArgumentError);
      writer.addAttribute("d", "6");
      // `b="x"` inside a value is no attribute.
      writer.addAttribute("b", "2");
    }
    const writer = new Xml.Writer();
    writer.openNode("r", { a: "1" });
    writer.snapshot();
    writer.addAttribute("b", "2");
    writer.rollback();
    writer.addAttribute("b", "3");
    expect(() => writer.addAttribute("a", "x")).toThrow(XmlArgumentError);
    writer.closeAll();
    expect(writer.toString()).toBe('<r a="1" b="3"/>');
  });

  it("take booleans as text, as they always have", () => {
    const writer = new Xml.Writer();
    writer.openNode("r");
    writer.writeText(true);
    writer.leafNode("f", undefined, false);
    writer.closeAll();
    expect(writer.toString()).toBe("<r>true<f>false</f></r>");
  });
});

describe("Xml.encodeAttr", () => {
  it("names its parameter when it refuses it", () => {
    expect(() => Xml.encodeAttr(42 as never)).toThrow(expect.objectContaining({ field: "value" }));
  });
});

describe("the writers' namespace declarations", () => {
  it.each([
    ["xmlns:xml", "urn:wrong"],
    ["xmlns:xmlns", "urn:x"],
    ["xmlns:p", ""],
    ["xmlns:p", "http://www.w3.org/XML/1998/namespace"],
    ["xmlns:p", "http://www.w3.org/2000/xmlns/"],
    ["xmlns", "http://www.w3.org/2000/xmlns/"]
  ])("refuse %s=%j, which no document can hold", (name, value) => {
    for (const make of [() => new Xml.Writer(), () => new Xml.StreamWriter({ write() {} })]) {
      expect(() => make().leafNode("r", { [name]: value })).toThrow(XmlArgumentError);
      const writer = make();
      writer.openNode("r");
      expect(() => writer.addAttribute(name, value)).toThrow(XmlArgumentError);
    }
  });

  it("let a prefix be whose binding a document around it may declare", () => {
    const writer = new Xml.Writer();
    writer.leafNode("p:r", {
      "xml:lang": "en",
      "xmlns:xml": "http://www.w3.org/XML/1998/namespace"
    });
    expect(writer.toString()).toBe(
      '<p:r xml:lang="en" xmlns:xml="http://www.w3.org/XML/1998/namespace"/>'
    );
  });
});

describe("a walk of a tree a caller built", () => {
  it("refuses only an element among its own descendants, however many it has passed", () => {
    // Far more elements walked than the depth at which ancestors are looked up, and one of them under
    // two parents: shared, not a cycle.
    const shared = element("s", [{ type: "text", value: "x" }]);
    // The first parent before the walk starts looking up ancestors, the second long after.
    const children: XmlNode[] = [element("p", [shared])];
    for (let i = 0; i < 3000; i++) {
      children.push(element("a", [element("b")]));
    }
    children.push(element("q", [shared]));
    const root = element("r", children);
    expect(Xml.getText(root)).toBe("xx");
    let count = 0;
    Xml.eachElement(root, () => void count++);
    expect(count).toBe(6000 + 5);
    expect(Xml.queryAll(root, "//s")).toHaveLength(2);
    // A cycle below the depth looked up at first is still found.
    const looped = element("l");
    looped.children = [looped];
    for (const walk of [
      () => Xml.getText(looped),
      () => Xml.eachElement(looped, () => {}),
      () => Xml.queryAll(looped, "//x")
    ]) {
      expect(walk).toThrow(XmlArgumentError);
    }
  });
});

describe("Xml.query and Xml.queryAll on trees a caller built", () => {
  it("refuse a cycle through an element already walked", () => {
    const a = element("a");
    a.children.push(a);
    const root = element("r", [a]);
    for (const run of [() => Xml.queryAll(root, "//x"), () => Xml.query(root, "//x")]) {
      expect(run).toThrow(XmlArgumentError);
      try {
        run();
      } catch (error) {
        expect((error as XmlArgumentError).field).toBe("element");
      }
    }
  });

  it("find each nested context's own indexed descendant, once", () => {
    let chain = element("a", [element("b")]);
    for (let i = 0; i < 399; i++) {
      chain = element("a", [chain]);
    }
    // Every `a` names the one `b` as its first: one result.
    expect(Xml.queryAll(element("r", [chain]), "//a//b[0]")).toHaveLength(1);
    // Each nested context's own Nth match, as walking each context would find it.
    const tree = Xml.parse("<r><a><b i='1'/><a><b i='2'/><b i='3'/></a><b i='4'/></a></r>").root;
    expect(Xml.queryAll(tree, "//a//b[1]").map(b => b.attributes.i)).toEqual(["2", "3"]);
    expect(Xml.query(tree, "//a//b[1]")?.attributes.i).toBe("2");
  });

  it("reach a shared element along each path, and refuse one shared along exponentially many", () => {
    // One element under two parents is reached through each, as in the tree the paths spell out.
    const z = element("z");
    const shared = element("r", [element("d", [z]), element("d", [z])]);
    expect(Xml.queryAll(shared, "//d/z")).toEqual([z, z]);
    // Each level sharing the one below twice: linear in size, 2^40 paths. Refused, not walked.
    let node = element("z");
    for (let i = 0; i < 40; i++) {
      node = element("d", [node, node]);
    }
    expect(() => Xml.queryAll(element("r", [node]), "//d/z")).toThrow(XmlArgumentError);
  });

  it("Xml.query reads no further than its first match", () => {
    const children: XmlNode[] = [element("x"), null as unknown as XmlNode];
    for (let i = 0; i < 1000; i++) {
      children.push(element("y"));
    }
    const root = element("r", children);
    expect(Xml.query(root, "x")?.name).toBe("x");
    expect(() => Xml.queryAll(root, "x")).toThrow(XmlArgumentError);
  });
});

describe("the DOM helpers use the value they checked", () => {
  it("reads each field once", () => {
    let reads = 0;
    const flip = {
      type: "element",
      name: "r",
      attributes: {},
      get children() {
        return reads++ === 0 ? [{ type: "text", value: "t" }] : 42;
      }
    } as unknown as XmlElement;
    expect(Xml.getText(flip)).toBe("t");
    let valueReads = 0;
    const text = {
      type: "text",
      get value() {
        return valueReads++ === 0 ? "s" : 42;
      }
    } as unknown as XmlNode;
    expect(Xml.getText(text)).toBe("s");
    let attributeReads = 0;
    const attrs = {
      type: "element",
      name: "r",
      children: [],
      get attributes() {
        return attributeReads++ === 0 ? { a: "1" } : 42;
      }
    } as unknown as XmlElement;
    expect(Xml.getAttr(attrs, "a")).toBe("1");
    attributeReads = 0;
    expect(Xml.query(element("w", [attrs]), "r[@a='1']")).toBe(attrs);
  });
});

function writers(): Array<{ w: Xml.Writer | Xml.StreamWriter; out: () => string }> {
  const w = new Xml.Writer();
  const chunks: Array<string | Uint8Array> = [];
  const sw = new Xml.StreamWriter({ write: chunk => void chunks.push(chunk) });
  return [
    { w, out: () => w.toString() },
    { w: sw, out: () => chunks.join("") }
  ];
}

describe("the writers' attributes", () => {
  it("refuse a write from an attribute getter, and are left unchanged", () => {
    for (const { w, out } of writers()) {
      w.openNode("r");
      expect(() =>
        w.addAttributes({
          get a() {
            w.writeText("X");
            return "Y";
          }
        })
      ).toThrow(XmlWriteError);
      expect(() =>
        w.addAttributes({
          get b() {
            w.addAttribute("b", "inner");
            return "outer";
          }
        })
      ).toThrow(XmlWriteError);
      w.addAttribute("b", "1");
      w.closeNode();
      expect(out()).toBe('<r b="1"/>');
    }
  });

  it("check a namespace declaration as it is written", () => {
    for (const { w } of writers()) {
      expect(() => w.leafNode("r", { "xmlns:p": "\u0001" })).toThrow(XmlArgumentError);
      expect(() =>
        w.leafNode("r", { "xmlns:p": "http://www.w3.org/XML/1998/namespace\u0001" })
      ).toThrow(XmlArgumentError);
      w.openNode("r");
      expect(() => w.addAttribute("xmlns:p", "\u0001")).toThrow(XmlArgumentError);
    }
  });

  it("refuse two names for one expanded name where both bindings are known", () => {
    for (const { w, out } of writers()) {
      expect(() =>
        w.leafNode("r", { "xmlns:p": "u", "xmlns:q": "u", "p:a": "1", "q:a": "2" })
      ).toThrow(XmlArgumentError);
      // A declaration after the attributes it binds.
      w.openNode("r", { "p:a": "1", "q:a": "2", "xmlns:p": "u" });
      expect(() => w.addAttribute("xmlns:q", "u")).toThrow(XmlArgumentError);
      // An ancestor's declarations are in scope.
      w.openNode("s", { "xmlns:q": "v", "p:a": "1" });
      expect(() => w.addAttributes({ "xmlns:z": "u", "z:a": "2" })).toThrow(XmlArgumentError);
      // Unbound prefixes are a fragment's, not an error.
      w.leafNode("t", { "x:a": "1", "y:a": "2" });
      w.closeAll();
      const xml = out();
      expect(xml).toBe(
        '<r p:a="1" q:a="2" xmlns:p="u"><s xmlns:q="v" p:a="1"><t x:a="1" y:a="2"/></s></r>'
      );
    }
  });

  it("forget a declaration a rollback undid", () => {
    const w = new Xml.Writer();
    w.openNode("r", { "p:a": "1" });
    w.snapshot();
    w.addAttribute("xmlns:p", "u");
    w.addAttribute("xmlns:q", "v");
    w.rollback();
    w.addAttributes({ "xmlns:p": "u", "xmlns:q": "w", "q:a": "2" });
    w.closeNode();
    expect(w.toString()).toBe('<r p:a="1" xmlns:p="u" xmlns:q="w" q:a="2"/>');
  });

  it("take a very large attribute map", () => {
    const attributes: Record<string, string> = {};
    for (let i = 0; i < 150_000; i++) {
      attributes[`a${i}`] = "1";
    }
    const w = new Xml.Writer();
    w.openNode("r");
    expect(() => w.addAttributes(attributes)).not.toThrow();
  });
});
