/**
 * Shared helpers for the round-trip integrity tests: build a package from raw
 * part XML, read a package's parts, and check a written package for the
 * structural defects a lossy round trip produces.
 */

import { zip } from "@archive/create-archive";
import { unzip } from "@archive/read-archive";
import { utf8Decoder, utf8Encoder } from "@word/core/internal-utils";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS = `xmlns:w="${W_NS}" xmlns:r="${R_NS}"`;

/** Every part of a package, by path without a leading slash. */
export async function packageParts(buffer: Uint8Array): Promise<Map<string, Uint8Array>> {
  const out = new Map<string, Uint8Array>();
  for await (const entry of unzip(buffer).entries()) {
    out.set(entry.path.replace(/^\/+/, ""), await entry.bytes());
  }
  return out;
}

/** The text of one part. */
export function partText(parts: Map<string, Uint8Array>, path: string): string {
  const bytes = parts.get(path);
  return bytes ? utf8Decoder.decode(bytes) : "";
}

/** Options for {@link buildDocx}. */
export interface BuildDocxParts {
  /** Content of `w:body`, before the final `w:sectPr`. */
  readonly body: string;
  /** Content of the final `w:sectPr` (a header reference is added for `header`). */
  readonly sectPr?: string;
  /** Content of `word/header1.xml`'s `w:hdr`, referenced as `rIdHdr`. */
  readonly header?: string;
  /** Content of `word/settings.xml`'s `w:settings`, when it should have one. */
  readonly settings?: string;
  /** Styles added to `word/styles.xml` after its `Normal`. */
  readonly styles?: string;
}

/** A minimal hand-written package (no library writer involved) around `parts`. */
export function buildDocx(parts: BuildDocxParts): Promise<Uint8Array> {
  const archive = zip();
  const add = (path: string, xml: string): void => {
    archive.add(path, utf8Encoder.encode(xml));
  };
  const decl = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
  let rels = `<Relationship Id="rIdStyles" Type="${REL}/styles" Target="styles.xml"/>`;
  let overrides =
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`;
  let headerRef = "";
  if (parts.header !== undefined) {
    add("word/header1.xml", `${decl}<w:hdr ${NS}>${parts.header}</w:hdr>`);
    rels += `<Relationship Id="rIdHdr" Type="${REL}/header" Target="header1.xml"/>`;
    overrides += `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`;
    headerRef = `<w:headerReference w:type="default" r:id="rIdHdr"/>`;
  }
  if (parts.settings !== undefined) {
    add("word/settings.xml", `${decl}<w:settings ${NS}>${parts.settings}</w:settings>`);
    rels += `<Relationship Id="rIdSettings" Type="${REL}/settings" Target="settings.xml"/>`;
    overrides += `<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>`;
  }
  add(
    "[Content_Types].xml",
    `${decl}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>${overrides}</Types>`
  );
  add(
    "_rels/.rels",
    `${decl}<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  add(
    "word/_rels/document.xml.rels",
    `${decl}<Relationships xmlns="${REL_NS}">${rels}</Relationships>`
  );
  add(
    "word/styles.xml",
    `${decl}<w:styles ${NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${parts.styles ?? ""}</w:styles>`
  );
  add(
    "word/document.xml",
    `${decl}<w:document ${NS}><w:body>${parts.body}<w:sectPr>${headerRef}${
      parts.sectPr ?? `<w:pgSz w:w="12240" w:h="15840"/>`
    }</w:sectPr></w:body></w:document>`
  );
  return archive.toBuffer();
}

/** Attributes of every `<prefix:tag …>` in `xml`. */
function elementAttrs(xml: string, tag: string): Record<string, string>[] {
  const out: Record<string, string>[] = [];
  for (const m of xml.matchAll(new RegExp(`<${tag}\\b([^>]*?)/?>`, "g"))) {
    const attrs: Record<string, string> = {};
    for (const a of m[1]!.matchAll(/([\w:]+)="([^"]*)"/g)) {
      attrs[a[1]!] = a[2]!;
    }
    out.push(attrs);
  }
  return out;
}

/** Range kinds whose start and end must pair by `w:id` within a part. */
const PAIRED_RANGES = [
  ["w:bookmarkStart", "w:bookmarkEnd"],
  ["w:permStart", "w:permEnd"],
  ["w:commentRangeStart", "w:commentRangeEnd"]
] as const;

/**
 * Structural defects in a written package:
 *
 * - a `w:p` with more than one `w:pPr`;
 * - an `r:id` (or `r:embed`/`r:link`) in a part that its relationships do not define;
 * - a `wp:docPr/@id` used twice anywhere in the package;
 * - a range start without its end (or the reverse), a start id used twice, or
 *   a bookmark id or name used twice in the package.
 */
export function packageIssues(parts: Map<string, Uint8Array>): string[] {
  const issues: string[] = [];
  const docPrIds = new Map<string, string>();
  const bookmarkIds = new Map<string, string>();
  const bookmarkNames = new Map<string, string>();
  for (const [path] of parts) {
    if (!path.startsWith("word/") || !path.endsWith(".xml") || path.includes("_rels/")) {
      continue;
    }
    const xml = partText(parts, path);
    // `w:pPr` is the first child of a paragraph and appears once, so a second
    // one always follows the first directly.
    const repeated = xml.match(/(?:<\/w:pPr>|<w:pPr\s*\/>)\s*<w:pPr\b/g)?.length ?? 0;
    if (repeated > 0) {
      issues.push(`${path}: ${repeated} paragraph(s) with a repeated w:pPr`);
    }
    const relsPath = path.replace(/([^/]+)$/, "_rels/$1.rels");
    const relIds = new Set(
      [...partText(parts, relsPath).matchAll(/\bId="([^"]+)"/g)].map(m => m[1]!)
    );
    for (const m of xml.matchAll(/\br:(?:id|embed|link|pict|dm|lo|qs|cs)="([^"]*)"/g)) {
      if (m[1] && !relIds.has(m[1])) {
        issues.push(`${path}: dangling r:id ${m[1]}`);
      }
    }
    for (const m of xml.matchAll(/<wp:docPr\b[^>]*?\bid="([^"]*)"/g)) {
      const seen = docPrIds.get(m[1]!);
      if (seen) {
        issues.push(`duplicate wp:docPr id ${m[1]} (${seen}, ${path})`);
      } else {
        docPrIds.set(m[1]!, path);
      }
    }
    for (const [startTag, endTag] of PAIRED_RANGES) {
      const starts = elementAttrs(xml, startTag).map(a => a["w:id"]!);
      const ends = elementAttrs(xml, endTag).map(a => a["w:id"]!);
      const open = new Set<string>();
      for (const id of starts) {
        if (open.has(id)) {
          issues.push(`${path}: ${startTag} id ${id} twice`);
        }
        open.add(id);
      }
      const closed = new Set<string>();
      for (const id of ends) {
        if (closed.has(id)) {
          issues.push(`${path}: ${endTag} id ${id} twice`);
        }
        closed.add(id);
        if (!open.has(id)) {
          issues.push(`${path}: ${endTag} id ${id} without start`);
        }
      }
      for (const id of open) {
        if (!closed.has(id)) {
          issues.push(`${path}: ${startTag} id ${id} without end`);
        }
      }
    }
    for (const a of elementAttrs(xml, "w:bookmarkStart")) {
      const id = a["w:id"]!;
      const name = a["w:name"]!;
      const idSeen = bookmarkIds.get(id);
      if (idSeen && idSeen !== path) {
        issues.push(`bookmark id ${id} in ${idSeen} and ${path}`);
      }
      bookmarkIds.set(id, path);
      if (bookmarkNames.has(name)) {
        issues.push(`bookmark name ${name} twice`);
      }
      bookmarkNames.set(name, path);
    }
  }
  return issues;
}

/** Parts that differ between two packages, ignoring the core properties' timestamps. */
export function differingParts(a: Map<string, Uint8Array>, b: Map<string, Uint8Array>): string[] {
  const normalize = (path: string, bytes: Uint8Array | undefined): string | undefined => {
    if (!bytes) {
      return undefined;
    }
    if (path !== "docProps/core.xml") {
      // Compare bytes, losslessly, as a string of code units.
      let s = "";
      for (const byte of bytes) {
        s += String.fromCharCode(byte);
      }
      return s;
    }
    return utf8Decoder
      .decode(bytes)
      .replace(/<dcterms:(created|modified)\b[^>]*>[^<]*<\/dcterms:\1>/g, "");
  };
  const out: string[] = [];
  for (const path of new Set([...a.keys(), ...b.keys()])) {
    if (normalize(path, a.get(path)) !== normalize(path, b.get(path))) {
      out.push(path);
    }
  }
  return out;
}

/** The instruction of every TOC block's own field, decoded, in document order. */
export function tocBlockInstructions(xml: string): string[] {
  const out: string[] = [];
  for (const m of xml.matchAll(/<w:sdt>([\s\S]*?)<\/w:sdt>/g)) {
    if (!/<w:docPartGallery w:val="Table of Contents"\/>/.test(m[1]!)) {
      continue;
    }
    // The block's own field is the first one; its instruction runs up to its separator.
    const head =
      /<w:fldChar w:fldCharType="begin"\/>([\s\S]*?)<w:fldChar w:fldCharType="separate"\/>/.exec(
        m[1]!
      );
    out.push(
      [...(head?.[1] ?? "").matchAll(/<w:instrText[^>]*>([^<]*)<\/w:instrText>/g)]
        .map(i => i[1]!)
        .join("")
        .replace(/&quot;/g, '"')
        .replace(/&amp;/g, "&")
    );
  }
  return out;
}
