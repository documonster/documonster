/**
 * Builds a `.docx` whose chart part is a given `c:chartSpace` — with its
 * relationships and the parts they point at — and compares chart parts the
 * way a reader would: as element trees, not as text.
 */

import { packZipSync, unpackZipSync } from "@archive/zip-one-shot";
import { Build, Document, Io } from "@word/index";
import type { DocxDocument } from "@word/index";
import { parsePartXml } from "@word/reader/part-xml";
import type { XmlElement } from "@xml/types";

export interface ChartPackage {
  readonly chartXml: string;
  readonly relsXml?: string;
  readonly parts?: ReadonlyMap<string, { readonly data: string; readonly contentType: string }>;
}

/** A document with one chart, whose `word/charts/chart1.xml` is `pkg.chartXml`. */
export async function docxWithChartPart(pkg: ChartPackage): Promise<Uint8Array> {
  const h = Document.create();
  Document.addContent(
    h,
    Build.chart({ type: "column", series: [{ name: "x", categories: ["a"], values: [1] }] })
  );
  const files = unpackZipSync(await Io.toBuffer(h));
  const entries = new Map<string, Uint8Array | string>();
  for (const [path, entry] of files) {
    entries.set(path, entry.data);
  }
  entries.set("word/charts/chart1.xml", pkg.chartXml);
  if (pkg.relsXml !== undefined) {
    entries.set("word/charts/_rels/chart1.xml.rels", pkg.relsXml);
  }
  if (pkg.parts) {
    let types = new TextDecoder().decode(entries.get("[Content_Types].xml") as Uint8Array);
    for (const [path, part] of pkg.parts) {
      entries.set(path, part.data);
      types = types.replace(
        "</Types>",
        `<Override PartName="/${path}" ContentType="${part.contentType}"/></Types>`
      );
    }
    entries.set("[Content_Types].xml", types);
  }
  return packZipSync(entries);
}

/** The package's entries as text, by path. */
export async function packageText(doc: DocxDocument, options?: Parameters<typeof Io.toBuffer>[1]) {
  const files = unpackZipSync(await Io.toBuffer(doc, options));
  const out = new Map<string, string>();
  for (const [path, entry] of files) {
    out.set(path, new TextDecoder().decode(entry.data));
  }
  return out;
}

function canonicalElement(el: XmlElement): string {
  const attrs = Object.entries(el.attributes)
    .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
    .map(([k, v]) => ` ${k}=${JSON.stringify(v)}`)
    .join("");
  let inner = "";
  for (const child of el.children) {
    if (child.type === "element") {
      inner += canonicalElement(child);
    } else if (child.type === "text" && child.value.trim() !== "") {
      inner += JSON.stringify(child.value);
    }
  }
  return `<${el.name}${attrs}>${inner}</${el.name}>`;
}

/** An XML part as a canonical string: attributes sorted, whitespace-only text dropped. */
export function canonicalXml(xml: string): string {
  return canonicalElement(parsePartXml(xml, undefined).root);
}

function localName(name: string): string {
  return name.slice(name.indexOf(":") + 1);
}

function collectPaths(el: XmlElement, parent: string, out: Map<string, number>): void {
  const attrs = Object.entries(el.attributes)
    .filter(([k]) => !k.startsWith("xmlns"))
    .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
    .map(([k, v]) => `${k}=${v}`)
    .join(",");
  const text = el.children
    .filter(c => c.type === "text")
    .map(c => (c as { value: string }).value)
    .join("")
    .trim();
  const path = `${parent}/${localName(el.name)}`;
  const key = `${path}${attrs ? `[${attrs}]` : ""}${text ? ` "${text}"` : ""}`;
  out.set(key, (out.get(key) ?? 0) + 1);
  for (const child of el.children) {
    if (child.type === "element") {
      collectPaths(child, path, out);
    }
  }
}

/**
 * What `after` lost from `before` and what it gained, as element paths with
 * their attributes and text (`chartSpace/chart/plotArea/…/numFmt[formatCode=@]`).
 */
export function chartPartDiff(before: string, after: string): { lost: string[]; added: string[] } {
  const a = new Map<string, number>();
  const b = new Map<string, number>();
  collectPaths(parsePartXml(before, undefined).root, "", a);
  collectPaths(parsePartXml(after, undefined).root, "", b);
  const lost: string[] = [];
  const added: string[] = [];
  for (const [key, n] of a) {
    for (let i = b.get(key) ?? 0; i < n; i++) {
      lost.push(key);
    }
  }
  for (const [key, n] of b) {
    for (let i = a.get(key) ?? 0; i < n; i++) {
      added.push(key);
    }
  }
  return { lost, added };
}
