/**
 * Both XML writers check a call's arguments before they change anything, so a refused
 * attribute, comment or CDATA section leaves the output well-formed and the writer where it
 * was. They used to flush a pending `>` — or half of a start tag — before validating.
 */
import { Xml, XmlError } from "@xml/index";
import { describe, expect, it } from "vitest";

interface WriterLike {
  openNode(name: string, attributes?: Record<string, unknown>): void;
  leafNode(name: string, attributes?: Record<string, unknown>, text?: unknown): void;
  addAttribute(name: string, value: unknown): void;
  addAttributes(attributes: Record<string, unknown>): void;
  writeText(text: unknown): void;
  writeCData(text: unknown): void;
  writeComment(text: unknown): void;
  writeRaw(xml: unknown): void;
  openXml(attributes?: Record<string, unknown>): void;
  closeAll(): void;
}

function writers(): Array<[string, () => { w: WriterLike; out: () => string }]> {
  return [
    [
      "Writer",
      () => {
        const w = new Xml.Writer();
        return { w: w as unknown as WriterLike, out: () => w.toString() };
      }
    ],
    [
      "StreamWriter",
      () => {
        const chunks: string[] = [];
        const w = new Xml.StreamWriter({ write: (chunk: string) => void chunks.push(chunk) });
        return { w: w as unknown as WriterLike, out: () => chunks.join("") };
      }
    ]
  ];
}

const refusals: Array<[string, (w: WriterLike) => void]> = [
  ["an attribute that is null", w => w.leafNode("b", { ok: "1", bad: null })],
  ["an attribute that is an object", w => w.openNode("b", { ok: "1", bad: { x: 1 } })],
  ["an attribute that is NaN", w => w.addAttributes({ ok: "1", bad: NaN })],
  ["an attribute named badly", w => w.leafNode("b", { ok: "1", "a b": "x" })],
  ["leaf text that is NaN", w => w.leafNode("b", {}, NaN)],
  ["text that is undefined", w => w.writeText(undefined)],
  ["a comment containing --", w => w.writeComment("a -- b")],
  ["a comment ending in -", w => w.writeComment("a-")],
  ["a comment that is not a string", w => w.writeComment(5)],
  ["CDATA that is not a string", w => w.writeCData(undefined)],
  ["raw markup that is undefined", w => w.writeRaw(undefined)],
  ["raw markup that is a number", w => w.writeRaw(42)]
];

describe.each(writers())("Xml.%s", (_name, make) => {
  it.each(refusals)("refuses %s without writing anything", (_case, call) => {
    const { w, out } = make();
    w.openNode("root", { id: "1" });
    expect(() => call(w)).toThrow(XmlError);
    w.closeAll();
    // The pending start tag is still the writer's to close, and nothing half-written leaked.
    expect(out()).toBe('<root id="1"/>');
  });

  it("splits ]]> across CDATA sections and strips characters XML cannot carry", () => {
    const { w, out } = make();
    w.openNode("r");
    w.writeCData("a]]>b");
    w.writeComment("x\u0000y");
    w.closeAll();
    expect(out()).toBe("<r><![CDATA[a]]]]><![CDATA[>b]]><!--xy--></r>");
  });

  it("keeps writing valid values as before", () => {
    const { w, out } = make();
    w.openNode("r", { a: 0, b: false, c: "", d: undefined });
    w.addAttributes({ e: "<&>" });
    w.leafNode("x", undefined, 1.5);
    w.writeComment(" ok ");
    w.closeAll();
    expect(out()).toBe('<r a="0" b="false" c="" e="&lt;&amp;&gt;"><x>1.5</x><!-- ok --></r>');
  });
});

describe.each(writers())("Xml.%s duplicate attributes", (_name, make) => {
  it.each<[string, (w: WriterLike) => void]>([
    ["addAttribute after openNode", w => w.addAttribute("id", "2")],
    ["addAttributes after openNode", w => w.addAttributes({ other: "x", id: "2" })],
    [
      "addAttribute twice",
      w => {
        w.addAttribute("k", "1");
        w.addAttribute("k", "2");
      }
    ]
  ])("refuses %s, naming the attribute", (_case, call) => {
    const { w, out } = make();
    w.openNode("root", { id: "1" });
    expect(() => call(w)).toThrow(XmlError);
    expect(() => call(w)).toThrow(/"(id|k)"/);
    w.closeAll();
    // Well-formed: the refused attribute was not written.
    expect(out()).not.toMatch(/ id="1"[^>]* id=/);
  });

  it("allows the same name again on the next element", () => {
    const { w, out } = make();
    w.openNode("a", { id: "1" });
    w.openNode("b");
    w.addAttribute("id", "2");
    w.closeAll();
    w.openNode("c");
    w.addAttribute("id", "3");
    w.closeAll();
    expect(out()).toBe('<a id="1"><b id="2"/></a><c id="3"/>');
  });
});

describe.each(writers())("Xml.%s — the XML declaration", (_name, make) => {
  it.each<[string, (w: WriterLike) => void]>([
    [
      "after text inside an element",
      w => {
        w.openNode("a");
        w.writeText("x");
      }
    ],
    ["after a closed root", w => w.leafNode("a")],
    ["after another declaration", w => w.openXml()],
    ["after a comment", w => w.writeComment("c")],
    ["after raw markup", w => w.writeRaw("<a/>")]
  ])("is refused %s", (_case, before) => {
    const { w, out } = make();
    before(w);
    const written = out();
    expect(() => w.openXml()).toThrow(/declaration must come first/);
    expect(out()).toBe(written);
  });

  it("is written at the start", () => {
    const { w, out } = make();
    w.openXml();
    w.leafNode("a");
    expect(out()).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a/>');
  });
});
