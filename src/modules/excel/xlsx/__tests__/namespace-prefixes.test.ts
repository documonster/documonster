/**
 * A part means the same thing whatever prefixes it binds its namespaces to. The reader used to match
 * the literal prefixes Excel writes, so a chart whose namespace was bound to another prefix — or made
 * the default namespace — crashed the chart parser.
 */
import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { expandedNames, kindOf, walk } from "@excel/__tests__/helpers/expanded-names";
import { Cell, Chart, Image, Workbook } from "@excel/index";
import {
  canonicalizeXml,
  createPrefixNormalizer,
  needsPrefixNormalization
} from "@excel/xlsx/xform/namespace-prefixes";
import { SaxParser } from "@xml/sax";
import { describe, expect, it } from "vitest";

const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const XSI = "http://www.w3.org/2001/XMLSchema-instance";
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

function book(): Workbook.Handle {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", 1);
  Cell.setValue(ws, "A2", 2);
  Chart.place(ws, "C1:J10", {
    type: "bar",
    title: "Totals",
    series: [{ name: "Values", values: "S!$A$1:$A$2" }]
  });
  Image.place(ws, Image.add(wb, { base64: PNG, extension: "png" }), "L1:M2");
  return wb;
}

/** `xml` with every `from:` element and attribute moved to `to` (`""`: the default namespace). */
function rebind(xml: string, from: string, uri: string, to: string): string {
  const prefixed = to === "" ? "" : `${to}:`;
  return xml
    .replace(new RegExp(`(</?)${from}:`, "g"), `$1${prefixed}`)
    .replace(` xmlns:${from}="${uri}"`, to === "" ? ` xmlns="${uri}"` : ` xmlns:${to}="${uri}"`);
}

async function rewrite(
  bytes: Uint8Array,
  path: string,
  edit: (xml: string) => string
): Promise<Uint8Array> {
  const entries = await extractAll(bytes);
  return createZip(
    [...entries].map(([name, file]) => ({
      name,
      data:
        name === path
          ? new TextEncoder().encode(edit(new TextDecoder().decode(file.data)))
          : file.data
    }))
  );
}

async function read(bytes: Uint8Array): Promise<Workbook.Handle> {
  const wb = Workbook.create();
  await Workbook.read(wb, bytes);
  return wb;
}

function chartSummary(wb: Workbook.Handle) {
  const [chart] = Chart.list(Workbook.getWorksheet(wb, "S")!);
  return {
    title: Chart.getTitle(chart),
    series: Chart.getSeriesCount(chart, 0),
    first: Chart.getSeries(chart, 0)
  };
}

describe("namespace prefixes are read by URI", () => {
  const variants = {
    "another prefix": (xml: string) => rebind(xml, "c", C, "ch"),
    "the default namespace": (xml: string) => rebind(xml, "c", C, "")
  };

  for (const [label, edit] of Object.entries(variants)) {
    it(`reads a chart whose namespace is bound to ${label}`, async () => {
      const original = await Workbook.toBuffer(book());
      const expected = chartSummary(await read(original));
      const edited = await rewrite(original, "xl/charts/chart1.xml", edit);
      expect(
        new TextDecoder().decode((await extractAll(edited)).get("xl/charts/chart1.xml")!.data)
      ).not.toContain("<c:chartSpace");

      const back = await read(edited);
      expect(chartSummary(back)).toEqual(expected);
      // And it writes back as a chart that reads the same again.
      expect(chartSummary(await read(await Workbook.toBuffer(back)))).toEqual(expected);
    });
  }

  it("reads a drawing whose namespace is bound to another prefix", async () => {
    const original = await Workbook.toBuffer(book());
    const edited = await rewrite(original, "xl/drawings/drawing1.xml", xml =>
      rebind(xml, "xdr", XDR, "sd")
    );
    const back = await read(edited);
    const ws = Workbook.getWorksheet(back, "S")!;
    expect(Image.list(ws)).toHaveLength(1);
    expect(Chart.list(ws)).toHaveLength(1);
  });

  it("writes an unedited part with Excel's prefixes", async () => {
    const edited = await rewrite(await Workbook.toBuffer(book()), "xl/charts/chart1.xml", xml =>
      rebind(xml, "c", C, "ch")
    );
    const back = await read(edited);
    Chart.setTitle(Chart.list(Workbook.getWorksheet(back, "S")!)[0], "Changed");
    const xml = new TextDecoder().decode(
      (await extractAll(await Workbook.toBuffer(back))).get("xl/charts/chart1.xml")!.data
    );
    expect(xml).toContain("<c:chartSpace");
    expect(xml).not.toContain("ch:");
  });
});

describe("a chart's kept bytes use Excel's prefixes", () => {
  // The bytes are patched in place on some edits, by expressions naming `c:` elements; kept with
  // another prefix, such an edit would miss, or insert `c:` elements into a part that never binds `c`.
  it("so even an unedited chart is written back with them", async () => {
    const edited = await rewrite(await Workbook.toBuffer(book()), "xl/charts/chart1.xml", xml =>
      rebind(xml, "c", C, "ch")
    );
    const written = await Workbook.toBuffer(await read(edited));
    const xml = new TextDecoder().decode(
      (await extractAll(written)).get("xl/charts/chart1.xml")!.data
    );
    expect(xml).toContain("<c:chartSpace");
    expect(xml).toContain(`xmlns:c="${C}"`);
    expect(xml).not.toMatch(/<\/?ch:/);
  });

  it("canonicalizeXml leaves a canonical part untouched and rewrites the rest", () => {
    const canonical = `<?xml version="1.0"?><c:chartSpace xmlns:c="${C}"><c:v>1 &amp; 2</c:v><!--n--></c:chartSpace>`;
    expect(canonicalizeXml(canonical)).toBe(canonical);
    expect(
      canonicalizeXml(
        `<?xml version="1.0"?><chartSpace xmlns="${C}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="q"><v>1 &amp; 2</v><e/></chartSpace>`
      )
    ).toBe(
      `<?xml version="1.0"?><c:chartSpace xmlns:c="${C}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="q"><c:v>1 &amp; 2</c:v><c:e/></c:chartSpace>`
    );
  });
});

describe("chartEx parts", () => {
  it("reads a chartEx whose namespace is the default namespace", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", 2);
    Chart.place(ws, "C1:J10", { type: "funnel", title: "F", series: [{ values: "S!$A$1:$A$2" }] });
    const original = await Workbook.toBuffer(wb);
    const CX = "http://schemas.microsoft.com/office/drawing/2014/chartex";
    const path = [...(await extractAll(original)).keys()].find(p => /chartEx\d+\.xml$/.test(p))!;
    const expected = Chart.getTitle(
      Chart.list(Workbook.getWorksheet(await read(original), "S")!)[0]
    );
    expect(expected).toBe("F");
    const back = await read(await rewrite(original, path, xml => rebind(xml, "cx", CX, "")));
    expect(Chart.getTitle(Chart.list(Workbook.getWorksheet(back, "S")!)[0])).toEqual(expected);
  });
});

describe("canonicalizeXml keeps every prefix declared", () => {
  it("declares an inherited unknown namespace renamed because a nested scope took its prefix", () => {
    const xml =
      `<root xmlns:c="urn:other"><c:a/><k:inner xmlns:k="${C}" c:attr="1">` +
      `<c:b/><k:chart/></k:inner></root>`;
    const out = canonicalizeXml(xml);
    expect(out).not.toBe(xml);
    expect(expandedNames(out)).toEqual(expandedNames(xml));
    expect(out).toContain(`<c:chart/>`);
  });

  it("names a namespace that became the default one in a prefix list through a declared prefix", () => {
    const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const xml = `<s:worksheet xmlns:s="${MAIN}" xmlns:m="${MC}" xmlns:q="urn:q" m:Ignorable="s q"/>`;
    const out = canonicalizeXml(xml);
    expect(out).toBe(
      `<worksheet xmlns="${MAIN}" xmlns:mc="${MC}" xmlns:q="urn:q" xmlns:s="${MAIN}" mc:Ignorable="s q"/>`
    );
    expect(expandedNames(out)).toEqual(expandedNames(xml));
  });
});

function escapeXml(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
}

/**
 * `xml` with every prefix renamed by `map` (`""`: the default namespace) — in names, declarations and the
 * values that name prefixes — so it means the same thing with other prefixes. An attribute, or a value
 * entry, cannot be moved to the default namespace; asking for that throws.
 */
function rebindPrefixes(xml: string, map: Readonly<Record<string, string>>): string {
  const to = (prefix: string) => map[prefix] ?? prefix;
  const prefixed = (prefix: string, rest: string, where: string) => {
    const target = to(prefix);
    if (target === "") {
      throw new Error(`${where} cannot move to the default namespace`);
    }
    return `${target}:${rest}`;
  };
  const name = (qname: string) => {
    const colon = qname.indexOf(":");
    if (colon === -1 && map[""] === undefined && Object.values(map).includes("")) {
      throw new Error(`${qname} is in no namespace, which a default namespace would capture`);
    }
    const target = to(colon === -1 ? "" : qname.slice(0, colon));
    const local = colon === -1 ? qname : qname.slice(colon + 1);
    return target === "" ? local : `${target}:${local}`;
  };
  let out = xml.startsWith("<?xml") ? xml.slice(0, xml.indexOf("?>") + 2) : "";
  let open = false;
  const parser = new SaxParser({ position: false });
  const flush = () => {
    if (open) {
      out += ">";
      open = false;
    }
  };
  const stack: Map<string, string>[] = [new Map()];
  parser.on("opentag", tag => {
    flush();
    const scope = new Map(stack[stack.length - 1]);
    for (const key in tag.attributes) {
      if (key === "xmlns") {
        scope.set("", tag.attributes[key]);
      } else if (key.startsWith("xmlns:")) {
        scope.set(key.slice(6), tag.attributes[key]);
      }
    }
    stack.push(scope);
    out += `<${name(tag.name)}`;
    for (const key in tag.attributes) {
      let attribute: string;
      let value = tag.attributes[key];
      if (key === "xmlns" || key.startsWith("xmlns:")) {
        const target = to(key === "xmlns" ? "" : key.slice(6));
        attribute = target === "" ? "xmlns" : `xmlns:${target}`;
      } else {
        const colon = key.indexOf(":");
        attribute = colon === -1 ? key : prefixed(key.slice(0, colon), key.slice(colon + 1), key);
        const kind = kindOf(tag.name, key, scope);
        const entries = value.split(/\s+/).filter(Boolean);
        if (kind === "prefixes") {
          value = entries.map(prefix => prefixed(prefix, "", key).slice(0, -1)).join(" ");
        } else if (kind === "qnames") {
          value = entries
            .map(entry => {
              const at = entry.indexOf(":");
              return prefixed(entry.slice(0, at), entry.slice(at + 1), key);
            })
            .join(" ");
        } else if (kind === "qname") {
          value = name(value.trim());
        }
      }
      out += ` ${attribute}="${escapeXml(value)}"`;
    }
    open = true;
  });
  parser.on("closetag", tag => {
    stack.pop();
    if (open) {
      out += "/>";
      open = false;
    } else {
      out += `</${name(tag.name)}>`;
    }
  });
  parser.on("text", text => {
    flush();
    out += escapeXml(text);
  });
  parser.write(xml);
  parser.close();
  return out;
}

/** `canonicalizeXml(xml)`, checked to mean exactly what `xml` means and to name only declared prefixes. */
function canonicalizedEquivalently(xml: string): string {
  const out = canonicalizeXml(xml);
  expect(expandedNames(out)).toEqual(expandedNames(xml));
  return out;
}

describe("canonicalizeXml rewrites the prefixes named in attribute values", () => {
  const CH = `xmlns:ch="${C}"`;

  it("resolves an MCE attribute by its namespace, not by the prefix mc", () => {
    // `m` is bound to MCE: its attributes are MCE lists. `mc` is bound to something else: its are not.
    const xml =
      `<root xmlns:m="${MC}" xmlns:mc="urn:not-mce" ${CH} m:Ignorable="ch" mc:Ignorable="ch">` +
      `<ch:foo/></root>`;
    const out = canonicalizedEquivalently(xml);
    expect(out).toContain(`mc:Ignorable="c"`);
    expect(out).toMatch(/ ns\d+:Ignorable="ch"/);
    // Without MCE in scope at all, an `mc:Ignorable` is just an attribute named that.
    const foreign = `<root xmlns:mc="urn:not-mce" ${CH} mc:Ignorable="ch"><ch:foo/></root>`;
    expect(canonicalizedEquivalently(foreign)).toContain(`mc:Ignorable="ch"`);
  });

  it("renames the prefix of every entry of mc:ProcessContent, wildcards included", () => {
    const xml =
      `<root xmlns:mc="${MC}" ${CH} xmlns:q="urn:q" mc:Ignorable="ch q" ` +
      `mc:ProcessContent="ch:foo  q:bar ch:*"><ch:foo/></root>`;
    const out = canonicalizedEquivalently(xml);
    expect(out).toContain(`mc:ProcessContent="c:foo q:bar c:*"`);
    expect(out).toContain(`mc:Ignorable="c q"`);
  });

  it("renames mc:PreserveElements and mc:PreserveAttributes the same way", () => {
    const xml =
      `<root xmlns:mc="${MC}" ${CH} mc:Ignorable="ch" ` +
      `mc:PreserveElements="ch:* ch:e" mc:PreserveAttributes="ch:a ch:*"><ch:e ch:a="1"/></root>`;
    const out = canonicalizedEquivalently(xml);
    expect(out).toContain(`mc:PreserveElements="c:* c:e"`);
    expect(out).toContain(`mc:PreserveAttributes="c:a c:*"`);
    expect(out).toContain(`<c:e c:a="1"/>`);
  });

  it("renames mc:MustUnderstand", () => {
    const out = canonicalizedEquivalently(
      `<root xmlns:mc="${MC}" ${CH} mc:MustUnderstand="ch"><ch:foo/></root>`
    );
    expect(out).toContain(`mc:MustUnderstand="c"`);
  });

  it("renames Requires on mc:Choice, and only there", () => {
    const xml =
      `<m:AlternateContent xmlns:m="${MC}"><m:Choice ${CH} Requires="ch"><ch:x/></m:Choice>` +
      `<m:Fallback/></m:AlternateContent>`;
    expect(canonicalizedEquivalently(xml)).toContain(`<mc:Choice xmlns:c="${C}" Requires="c">`);
    // An unprefixed `Requires` anywhere else is in no namespace and names nothing.
    const other = `<root ${CH}><ch:x Requires="ch"/></root>`;
    expect(canonicalizedEquivalently(other)).toContain(`<c:x Requires="ch"/>`);
  });

  it("gives a namespace that became the default one a declared prefix in every list kind", () => {
    const xml =
      `<s:worksheet xmlns:s="${MAIN}" xmlns:mc="${MC}" mc:Ignorable="s" mc:MustUnderstand="s" ` +
      `mc:ProcessContent="s:a s:*" mc:PreserveAttributes="s:b"><mc:AlternateContent>` +
      `<mc:Choice Requires="s"><s:a/></mc:Choice></mc:AlternateContent></s:worksheet>`;
    const out = canonicalizedEquivalently(xml);
    expect(out).toContain(`<worksheet xmlns="${MAIN}"`);
    expect(out).toContain(`mc:ProcessContent="s:a s:*"`);
    // Declared on the element that names it, the scope's own rename being unable to.
    expect(out).toContain(`<mc:Choice xmlns:s="${MAIN}" Requires="s">`);
  });

  it("picks another alias when the document's prefix is bound to something else in the output", () => {
    // `c` names the main namespace here, which becomes the default; the output binds `c` to the chart.
    const xml = `<c:root xmlns:c="${MAIN}" xmlns:mc="${MC}" xmlns:k="${C}" mc:Ignorable="c"><k:x/></c:root>`;
    const out = canonicalizedEquivalently(xml);
    expect(out).toMatch(/mc:Ignorable="ns\d+"/);
  });

  it("follows the default namespace in an unprefixed xsi:type", () => {
    // The default namespace is the chart one, renamed to `c`: an unprefixed type is a chart type.
    const xml = `<chartSpace xmlns="${C}" xmlns:i="${XSI}" i:type="T"><t i:type="ch:U" ${CH}/></chartSpace>`;
    const out = canonicalizedEquivalently(xml);
    expect(out).toContain(`xsi:type="c:T"`);
    expect(out).toContain(`xsi:type="c:U"`);
    // And a type in a namespace that became the default one loses its prefix.
    const main = `<s:worksheet xmlns:s="${MAIN}" xmlns:xsi="${XSI}" xsi:type="s:T"/>`;
    expect(canonicalizedEquivalently(main)).toContain(`xsi:type="T"`);
  });

  it("keeps values that hold URIs, and prefixes nothing declares", () => {
    const xml =
      `<ch:root ${CH} xmlns:xsi="${XSI}" xmlns:mc="${MC}" xsi:schemaLocation="${C} c.xsd" ` +
      `mc:Ignorable="zz"><ch:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"/></ch:root>`;
    const out = canonicalizeXml(xml);
    expect(out).toContain(`xsi:schemaLocation="${C} c.xsd"`);
    expect(out).toContain(`uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"`);
    expect(out).toContain(`mc:Ignorable="zz"`);
  });

  it("reproduces the reported case: mc:ProcessContent keeps naming a declared prefix", () => {
    const xml = `<root xmlns:mc="${MC}" ${CH} mc:ProcessContent="ch:foo"><ch:foo/></root>`;
    expect(canonicalizedEquivalently(xml)).toBe(
      `<root xmlns:mc="${MC}" xmlns:c="${C}" mc:ProcessContent="c:foo"><c:foo/></root>`
    );
  });
});

describe("canonicalizeXml keeps the default namespace right where a prefixed one became it", () => {
  // `s:` → the default namespace, so the output declares one where the document declared none.
  const WS = `xmlns:s="${MAIN}"`;

  it("keeps an element in no namespace out of it", () => {
    const xml = `<s:worksheet ${WS}><foo/><s:sheetData/></s:worksheet>`;
    expect(canonicalizedEquivalently(xml)).toBe(
      `<worksheet xmlns="${MAIN}"><foo xmlns=""/><sheetData/></worksheet>`
    );
  });

  it("re-declares as often as the nesting alternates, and only then", () => {
    const xml =
      `<s:worksheet ${WS}><foo><bar/><s:a><baz/><s:b/></s:a></foo>` + `<s:c/></s:worksheet>`;
    expect(canonicalizedEquivalently(xml)).toBe(
      `<worksheet xmlns="${MAIN}"><foo xmlns=""><bar/><a xmlns="${MAIN}"><baz xmlns=""/><b/></a>` +
        `</foo><c/></worksheet>`
    );
  });

  it("follows the document's own default declarations, and a scope that renames nothing", () => {
    const xml =
      `<s:worksheet ${WS}><d xmlns="urn:d"><e/><n xmlns=""><m/></n></d>` +
      // `s` re-bound to an unknown namespace: the scope renames nothing, but `MAIN` is still in force.
      `<x xmlns:s="urn:other"><y/><s:z/></x><s:after><w/></s:after></s:worksheet>`;
    const out = canonicalizedEquivalently(xml);
    // An unknown default namespace yields `""` to the known one, as everywhere: it moves to a prefix.
    expect(out).toContain(`<ns1:d xmlns:ns1="urn:d"><ns1:e/><n xmlns=""><m/></n></ns1:d>`);
    expect(out).toContain(`<x xmlns:s="urn:other" xmlns=""><y/><s:z/></x>`);
    expect(out).toContain(`<after><w xmlns=""/></after>`);
  });

  it("names a type in a namespace through a prefix where the default namespace is not its", () => {
    const xml = `<s:worksheet ${WS} xmlns:xsi="${XSI}"><foo xsi:type="s:T"/><s:a xsi:type="s:U"/></s:worksheet>`;
    const out = canonicalizedEquivalently(xml);
    expect(out).toContain(`<foo xsi:type="s:T" xmlns="" xmlns:s="${MAIN}"/>`);
    expect(out).toContain(`<a xsi:type="U"/>`);
  });

  it("writes an element with a prefix when it carries a type in no namespace", () => {
    const xml = `<s:worksheet ${WS} xmlns:xsi="${XSI}" xsi:type="T"><s:a/></s:worksheet>`;
    const out = canonicalizedEquivalently(xml);
    expect(out).toBe(
      `<s:worksheet xmlns="" xmlns:xsi="${XSI}" xsi:type="T" xmlns:s="${MAIN}"><a xmlns="${MAIN}"/></s:worksheet>`
    );
  });

  it("leaves the names an xform sees as they are, with no declaration added", () => {
    const prefixes = createPrefixNormalizer();
    prefixes.open({ name: "s:worksheet", attributes: { "xmlns:s": MAIN }, isSelfClosing: false });
    const foo = prefixes.open({ name: "foo", attributes: {}, isSelfClosing: true });
    expect(foo).toEqual({ name: "foo", attributes: {}, isSelfClosing: true });
    expect(prefixes.close("foo")).toBe("foo");
    expect(prefixes.close("s:worksheet")).toBe("worksheet");
  });
});

/** Every XML part of a workbook carrying a chart, a chartEx, an image and a note. */
async function writtenParts(): Promise<(readonly [string, string])[]> {
  const wb = book();
  const ws = Workbook.getWorksheet(wb, "S")!;
  Cell.setNote(ws, "A1", "note");
  Chart.place(ws, "C12:J20", { type: "funnel", title: "F", series: [{ values: "S!$A$1:$A$2" }] });
  return [...(await extractAll(await Workbook.toBuffer(wb)))]
    .filter(([path]) => /\.(xml|vml|rels)$/.test(path))
    .map(([path, file]) => [path, new TextDecoder().decode(file.data)] as const);
}

const parts = await writtenParts();

describe("every part this library writes, rebound to other prefixes, canonicalises equivalently", () => {
  /** The prefixes `xml` declares, in document order (`""`: the default namespace). */
  const declared = (xml: string) => [
    ...new Set([...xml.matchAll(/xmlns(?::([\w.-]+))?=/g)].map(m => m[1] ?? ""))
  ];

  it("covers the parts that matter", () => {
    const paths = parts.map(([path]) => path);
    for (const want of [
      /chart\d+\.xml$/,
      /chartEx\d+\.xml$/,
      /drawing\d+\.xml$/,
      /\.vml$/,
      /core\.xml$/
    ]) {
      expect(paths.some(path => want.test(path))).toBe(true);
    }
  });

  for (const [path, xml] of parts) {
    const prefixes = declared(xml);
    const variants: Record<string, Record<string, string>> = {
      // Every namespace under a prefix nothing knows, the default one included.
      "foreign prefixes": Object.fromEntries(prefixes.map(p => [p, `q${p}`])),
      // Each known namespace under another's canonical prefix: every rename collides with one.
      "rotated prefixes": Object.fromEntries(
        prefixes.filter(Boolean).map((p, i, all) => [p, all[(i + 1) % all.length]])
      )
    };
    // A namespace as the default one: the first whose prefix no attribute and no value names, in a part
    // with no element in no namespace for the default namespace to capture.
    const movable = prefixes.includes("")
      ? undefined
      : prefixes.find(p => {
          try {
            rebindPrefixes(xml, { [p]: "" });
            return true;
          } catch {
            return false;
          }
        });
    if (movable !== undefined) {
      variants["a namespace as the default"] = { [movable]: "" };
    }
    for (const [label, map] of Object.entries(variants)) {
      it(`${path}: ${label}`, () => {
        const input = rebindPrefixes(xml, map);
        expect(expandedNames(input)).toEqual(expandedNames(xml));
        const out = canonicalizedEquivalently(input);
        if (Object.entries(map).some(([from, to]) => from !== to)) {
          expect(out).not.toBe(input);
        }
        if (label === "foreign prefixes") {
          // Back in Excel's prefixes: its names are the original part's, but for the namespaces the
          // normaliser does not know, which keep the foreign prefix they were given.
          const names: string[] = [];
          walk(
            out,
            tag => names.push(tag.name.replace(/^q([\w.-]*:)/, "$1")),
            name => names.push(`/${name.replace(/^q([\w.-]*:)/, "$1")}`)
          );
          const original: string[] = [];
          walk(
            xml,
            tag => original.push(tag.name),
            name => original.push(`/${name}`)
          );
          expect(names).toEqual(original);
        }
      });
    }
  }
});

describe("canonicalizeXml re-serialises a rewritten part equivalently", () => {
  it("keeps text, comments, CDATA and encoded whitespace; drops XML-illegal characters", () => {
    const xml =
      `<?xml version="1.0"?><k:a xmlns:k="${C}" v="a\tb&#10;c\u0001">` +
      `<!--c--><![CDATA[<x>]]>t&amp;\u0001&#65;'</k:a>`;
    expect(canonicalizeXml(xml)).toBe(
      `<?xml version="1.0"?><c:a xmlns:c="${C}" v="a b&#xA;c">` +
        `<!--c--><![CDATA[<x>]]>t&amp;A&apos;</c:a>`
    );
  });
});

describe("a lazy normaliser enables itself from the raw text", () => {
  const S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const xml =
    `<worksheet xmlns="${S}"><sheetData><row r="1"/></sheetData>` +
    `<x:extra xmlns:x="${S}"><x:inner/></x:extra><tail/></worksheet>`;

  function names(chunks: readonly string[], lazy: boolean): string[] {
    const prefixes = createPrefixNormalizer({ lazy });
    const parser = new SaxParser({ position: false });
    const seen: string[] = [];
    parser.on("opentag", tag => seen.push(prefixes.open(tag).name));
    parser.on("closetag", tag => seen.push(`/${prefixes.close(tag.name)}`));
    for (const chunk of chunks) {
      prefixes.scan(chunk);
      parser.write(chunk);
    }
    parser.close();
    return seen;
  }

  it("matches an eager normaliser at every split point, including through a declaration", () => {
    const expected = names([xml], false);
    expect(expected).toContain("extra");
    expect(expected).toContain("/inner");
    for (let i = 1; i < xml.length; i++) {
      expect(names([xml.slice(0, i), xml.slice(i)], true)).toEqual(expected);
    }
  });

  it("skips no attribute and no declaration it has to rewrite, wherever the text is cut", () => {
    // Once active, a lazy normaliser counts in the text the declarations and the attributes that can
    // change, and leaves a tag alone while none is outstanding. Every way to fool that count is here:
    // a renamed main namespace and `r`, prefixed attributes on most tags, a declaration below the root,
    // and `xmlns` and `rel:` in text content and in an attribute value.
    const doc =
      `<?xml version="1.0"?><s:worksheet xmlns:s="${S}" xmlns:rel="${R}" xmlns:mc="${MC}" ` +
      `xmlns:x14ac="urn:x14ac" mc:Ignorable="x14ac"><s:sheetData><s:row r="0" rel:id="rId0"/>` +
      `<s:row r="00"/><s:row r="000" rel:id="rId00"/><s:row r="1" x14ac:dyDescent="1">` +
      `<s:c r="A1" t="inlineStr"><s:is><s:t>xmlns rel:id=" s:x</s:t></s:is></s:c></s:row>` +
      `<s:row r="2"><s:c r="A2" s="xmlns:q"><s:v>1</s:v></s:c></s:row></s:sheetData>` +
      `<s:drawing rel:id="rId1"/><s:extLst><s:ext uri="{X}" xmlns:q="urn:q" xmlns:sp="${S}">` +
      `<sp:inner rel:id="rId2" q:a="1"/><q:b rel:id="rId3"/></s:ext></s:extLst></s:worksheet>`;
    // `hinted` is the protocol the readers use: the parser counts prefixed attribute names, and the normaliser is
    // told per tag whether the count moved instead of searching the text — counting from the chunk that enables it,
    // so the tag the parser is part-way through then is reported as having one.
    const tags = (chunks: readonly string[], mode: "eager" | "lazy" | "hinted"): string[] => {
      const prefixes = createPrefixNormalizer({
        lazy: mode !== "eager",
        hinted: mode === "hinted"
      });
      const parser = new SaxParser({ position: false });
      const seen: string[] = [];
      let active = false;
      let counted = 0;
      parser.on("opentag", tag => {
        let opened;
        if (mode === "hinted") {
          opened = prefixes.open(tag, parser.prefixedAttributes !== counted);
          counted = parser.prefixedAttributes;
        } else {
          opened = prefixes.open(tag);
        }
        seen.push(`${opened.name} ${JSON.stringify({ ...opened.attributes })}`);
      });
      parser.on("closetag", tag => seen.push(`/${prefixes.close(tag.name)}`));
      for (const chunk of chunks) {
        if (prefixes.scan(chunk) && !active && mode === "hinted") {
          active = true;
          parser.countPrefixedAttributes = true;
          counted = -1;
        }
        parser.write(chunk);
      }
      parser.close();
      return seen;
    };
    const expected = tags([doc], "eager");
    expect(expected).toContain(`drawing {"r:id":"rId1"}`);
    expect(expected).toContain(`inner {"r:id":"rId2","q:a":"1"}`);
    for (const mode of ["lazy", "hinted"] as const) {
      expect(tags([doc], mode)).toEqual(expected);
      for (let i = 1; i < doc.length; i++) {
        expect(tags([doc.slice(0, i), doc.slice(i)], mode)).toEqual(expected);
      }
      // And cut into many small chunks.
      for (const size of [1, 3, 7, 16]) {
        const chunks: string[] = [];
        for (let at = 0; at < doc.length; at += size) {
          chunks.push(doc.slice(at, at + size));
        }
        expect(tags(chunks, mode)).toEqual(expected);
      }
    }
  });

  it("does no work for a part with only canonical declarations", () => {
    const canonical = `<worksheet xmlns="${S}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a r:id="x"/></worksheet>`;
    const prefixes = createPrefixNormalizer({ lazy: true });
    prefixes.scan(canonical);
    const tag = { name: "x:odd", attributes: { "xmlns:x": S }, isSelfClosing: true };
    expect(prefixes.open(tag)).toBe(tag);
    expect(needsPrefixNormalization(canonical)).toBe(false);
    expect(needsPrefixNormalization(xml)).toBe(true);
  });
});
