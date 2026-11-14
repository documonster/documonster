/**
 * DOCX Module - Flat OPC Format Support
 *
 * Implements reading and writing of Flat OPC (Office Open XML Package) format,
 * which represents an entire DOCX package as a single XML file.
 * This is useful for debugging, XML transformations, and enterprise integration scenarios.
 *
 * Flat OPC uses the `pkg:` namespace to wrap all package parts in a single XML document.
 * Each part is represented as a `<pkg:part>` element with its content either as XML or base64.
 *
 * References:
 * - [MS-OFFCRYPTO] §2.1.5 (pkgStart namespace)
 * - ECMA-376 Part 2 (OPC) Annex C (non-normative)
 */

import { xmlEncodeAttr } from "@utils/xml-encode";
import { paceChunks, runSteps, runStepsPaced } from "@utils/yield";
import type { Steps, YieldPoint } from "@utils/yield";
import { NS_CONTENT_TYPES, NS_PKG } from "@word/constants";
import { checkBytes, invalidArgument } from "@word/core/argument-checks";
import { utf8Decoder, utf8Encoder } from "@word/core/internal-utils";
import { getFileExt } from "@word/core/opc-paths";
import { DocxParseError } from "@word/errors";
import { XmlParseError } from "@xml/errors";
import { SaxParser } from "@xml/sax";
import type { XmlSaxTag } from "@xml/types";
import { XmlWriter } from "@xml/writer";

// =============================================================================
// Types
// =============================================================================

/** A part extracted from Flat OPC XML. */
export interface FlatOpcPart {
  /** Part name (e.g. "/word/document.xml"). */
  readonly name: string;
  /** Content type of the part. */
  readonly contentType: string;
  /** Raw data (decoded from base64 for binary, or UTF-8 encoded XML). */
  readonly data: Uint8Array;
}

// =============================================================================
// Reading Flat OPC
// =============================================================================

/**
 * Parse a Flat OPC XML document into individual package parts.
 * Returns a Map keyed by part path (without leading slash) to binary data,
 * compatible with the archive entry format used by `readDocx`.
 *
 * @param xmlContent - The Flat OPC XML string or UTF-8 bytes.
 * @returns A map of part paths to their binary content.
 */
export function parseFlatOpc(
  xmlContent: string | Uint8Array | ArrayBuffer | ArrayBufferView
): Map<string, Uint8Array> {
  const given = checkTextOrBytes(xmlContent, "xmlContent", "Convert.flatOpcToParts");
  const bytes = typeof given === "string" ? utf8Encoder.encode(given) : given;
  const read = startFlatOpcRead(bytes, work =>
    typeof work === "function" ? work() : runSteps(work)
  );
  try {
    read.parser.write(read.body).close();
  } catch (error) {
    throw flatOpcError(error);
  }
  return read.parts();
}

/**
 * {@link parseFlatOpc} a slice of the input at a time, asking `yieldPoint` (a `createYieldBudget`) between the slices —
 * for the streaming reader, which must give the event loop a turn while it reads a Flat OPC document: the whole
 * package is one XML file, and parsing it in one call blocked for as long as it took.
 */
export async function parseFlatOpcPaced(
  bytes: Uint8Array,
  yieldPoint: YieldPoint
): Promise<Map<string, Uint8Array>> {
  // A part's content is finished — its XML made to stand alone, its base64 decoded — between slices rather than in the
  // event that ends it: one large part is one event, and finishing it in there held the event loop for as long as the
  // part took. What is queued lies before where the parser has got to, so when the parser fails, the queue is finished
  // first: a failure in it comes first in the document, and is the one a read that finishes each part at once reports.
  const queued: FinishWork[] = [];
  const read = startFlatOpcRead(bytes, work => queued.push(work));
  const finishQueued = async (): Promise<void> => {
    for (const work of queued.splice(0)) {
      if (typeof work === "function") {
        work();
      } else {
        await runStepsPaced(work, yieldPoint);
      }
    }
  };
  try {
    // The body written as the whole read writes it, an empty one included: the parser's error then names the byte.
    for await (const slice of paceChunks([read.body], yieldPoint)) {
      await whenParsed(() => read.parser.write(slice), finishQueued);
    }
    if (read.body.length === 0) {
      read.parser.write(read.body);
    }
    await whenParsed(() => read.parser.close(), finishQueued);
  } catch (error) {
    throw flatOpcError(error);
  }
  return read.parts();
}

/** `parse`, then `finishQueued` — which, when `parse` fails, runs before its failure is thrown and may replace it. */
async function whenParsed(parse: () => void, finishQueued: () => Promise<void>): Promise<void> {
  try {
    parse();
  } catch (error) {
    await finishQueued();
    throw error;
  }
  await finishQueued();
}

/** Work finishing a part: its content's steps, or recording it once they have run. */
type FinishWork = Steps<void> | (() => void);

/**
 * A Flat OPC read: the parser, the input to give it from its first `<`, and the parts its events make. `finish` is
 * given the work of finishing each part's content, and of recording each part, in document order, to run now or
 * later — in that order, before `parts` is asked for.
 */
function startFlatOpcRead(
  bytes: Uint8Array,
  finish: (work: FinishWork) => void
): {
  readonly parser: SaxParser;
  readonly body: Uint8Array;
  parts(): Map<string, Uint8Array>;
} {
  // Text pasted or saved by an editor often begins with a byte order mark or white space before the
  // XML declaration, where XML allows neither: a Flat OPC file is read from its first `<`.
  let lead = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  while (lead < bytes.length && isSpaceByte(bytes[lead]!)) {
    lead++;
  }

  const reader = new FlatOpcReader(bytes, finish);
  // With namespaces: a package element is one in the package namespace, whatever its prefix.
  const parser = new SaxParser({ position: false, xmlns: true, byteOffset: lead });
  parser.on("opentag", tag => reader.open(parser, tag));
  parser.on("closetag", tag => reader.close(parser, tag));
  parser.on("text", text => reader.text(parser, text, false));
  parser.on("cdata", text => reader.text(parser, text, true));
  return {
    parser,
    body: bytes.subarray(lead),
    parts() {
      const { parts, contentTypes } = reader;
      // Ensure [Content_Types].xml exists
      if (!parts.has("[Content_Types].xml")) {
        parts.set("[Content_Types].xml", buildContentTypesXml(contentTypes));
      }
      return parts;
    }
  };
}

/** What a Flat OPC read threw, as it reports it: malformed XML is the reader's `DocxParseError`. */
function flatOpcError(error: unknown): unknown {
  if (error instanceof XmlParseError) {
    return new DocxParseError(`Flat OPC: not well-formed XML: ${error.message}`, {
      cause: error
    });
  }
  return error;
}

function isSpaceByte(byte: number): boolean {
  return byte === 0x20 || byte === 0x09 || byte === 0x0a || byte === 0x0d;
}

/** A part being read: where its `xmlData` content lies, or the text of its `binaryData`. */
interface OpenPart {
  readonly path: string;
  readonly contentType: string;
  /** Whether its `xmlData` or `binaryData` has ended: what follows is outside it. */
  ended?: boolean;
  /** `xmlData`: where its content begins in the input, and the namespaces in force there. */
  xml?: {
    readonly from: number;
    readonly scope: ReadonlyMap<string, string>;
    root?: XmlSaxTag;
    rootStart?: number;
  };
  /** `binaryData`: its base64 text so far. */
  base64?: string;
  /** Its content, once finished. */
  data?: Uint8Array;
}

/**
 * Reads a Flat OPC document's parts off SAX events. A part's XML is taken from the input as it is —
 * its comments, processing instructions, CDATA sections and references with it — and its root is given
 * the namespace declarations it used from the elements around it, so the part stands on its own.
 */
class FlatOpcReader {
  readonly parts = new Map<string, Uint8Array>();
  readonly contentTypes = new Map<string, string>();
  /** The namespaces in force at each open element: prefix (`""` the default) → URI. */
  private readonly scopes: Array<ReadonlyMap<string, string>> = [new Map()];
  private part: OpenPart | undefined;
  /** Whether the document's DOCTYPE declares entities or attribute lists. */
  private doctype = false;

  constructor(
    private readonly bytes: Uint8Array,
    private readonly finish: (work: FinishWork) => void
  ) {}

  open(parser: SaxParser, tag: XmlSaxTag): void {
    const outer = this.scopes[this.scopes.length - 1]!;
    let scope = outer;
    if (tag.ns !== undefined) {
      const inner = new Map(outer);
      for (const prefix of Object.keys(tag.ns)) {
        inner.set(prefix, tag.ns[prefix]!);
      }
      scope = inner;
    }
    const depth = parser.depth;
    if (!tag.isSelfClosing) {
      this.scopes.push(scope);
    }
    const part = this.part;
    if (part?.xml !== undefined && part.ended !== true) {
      // Inside a part's XML: only its root is looked at, to declare on it what it inherits.
      if (depth === 3) {
        if (part.xml.root !== undefined) {
          throw new DocxParseError(
            `Flat OPC: the xmlData of part "${part.path}" has more than one root element`
          );
        }
        part.xml.root = tag;
        part.xml.rootStart = parser.startByte!;
      }
      return;
    }
    if (part?.base64 !== undefined && part.ended !== true) {
      throw new DocxParseError(
        `Flat OPC: the binaryData of part "${part.path}" contains an element <${tag.name}>`
      );
    }
    if (depth === 0) {
      if (!isPackageElement(tag, "package")) {
        throw new DocxParseError(
          `Flat OPC: expected root element <pkg:package> of ${NS_PKG}, got <${tag.name}>`
        );
      }
      this.doctype = scanPrologue(this.bytes, 0, parser.startByte!)?.doctype ?? false;
    } else if (depth === 1 && isPackageElement(tag, "part")) {
      // Get part name (with leading /)
      const name = packageAttribute(tag, scope, "name");
      this.part = {
        // Normalize: remove leading slash for map key
        path: name.startsWith("/") ? name.substring(1) : name,
        contentType: packageAttribute(tag, scope, "contentType")
      };
    } else if (depth === 2 && part !== undefined) {
      const xmlData = isPackageElement(tag, "xmlData");
      if (!xmlData && !isPackageElement(tag, "binaryData")) {
        return;
      }
      if (part.xml !== undefined || part.base64 !== undefined || part.ended === true) {
        throw new DocxParseError(
          `Flat OPC: part "${part.path}" has more than one xmlData or binaryData element`
        );
      }
      if (xmlData) {
        if (this.doctype) {
          // Its defaulted attributes and entities would be lost: the part is copied, not re-serialised.
          throw new DocxParseError(
            `Flat OPC: part "${part.path}" has xmlData in a document whose DOCTYPE declares entities or attribute lists, which cannot be carried into the part`
          );
        }
        part.xml = { from: parser.endByte!, scope };
        if (tag.isSelfClosing) {
          this.endXmlData(part, parser.endByte!);
        }
      } else {
        part.base64 = "";
        if (tag.isSelfClosing) {
          this.endBinaryData(part);
        }
      }
    }
  }

  close(parser: SaxParser, tag: XmlSaxTag): void {
    if (!tag.isSelfClosing) {
      this.scopes.pop();
    }
    const depth = parser.depth;
    const part = this.part;
    if (part === undefined || tag.isSelfClosing) {
      return;
    }
    if (depth === 2 && part.xml !== undefined && part.ended !== true) {
      this.endXmlData(part, parser.startByte!);
    } else if (depth === 2 && part.base64 !== undefined && part.ended !== true) {
      this.endBinaryData(part);
    } else if (depth === 1) {
      this.finish(() => this.recordPart(part));
      this.part = undefined;
    }
  }

  text(parser: SaxParser, text: string, cdata: boolean): void {
    const part = this.part;
    if (part === undefined || part.ended === true) {
      return;
    }
    if (part.base64 !== undefined) {
      part.base64 += text;
    } else if (part.xml !== undefined && parser.depth === 3 && (cdata || /[^ \t\r\n]/.test(text))) {
      throw new DocxParseError(
        `Flat OPC: the xmlData of part "${part.path}" has character data outside its root element`
      );
    }
  }

  /** The end of a part's `xmlData`, at `to` in the input: checked to have had a root, and its content queued. */
  private endXmlData(part: OpenPart, to: number): void {
    const { from, scope, root, rootStart } = part.xml!;
    if (root === undefined) {
      throw new DocxParseError(`Flat OPC: the xmlData of part "${part.path}" has no root element`);
    }
    part.ended = true;
    this.finish(
      contentSteps(
        part,
        standaloneXml(this.bytes, { from, to, scope, root, rootStart: rootStart! })
      )
    );
  }

  /** The end of a part's `binaryData`: its text queued to be decoded. */
  private endBinaryData(part: OpenPart): void {
    const text = part.base64!;
    part.base64 = undefined;
    part.ended = true;
    this.finish(contentSteps(part, base64DecodeSteps(text, part.path)));
  }

  /** A part recorded, once its content is finished: its bytes, then its content type for `[Content_Types].xml`. */
  private recordPart(part: OpenPart): void {
    if (part.path !== "" && part.data !== undefined) {
      this.parts.set(part.path, part.data);
    }
    // Collect content type info for [Content_Types].xml reconstruction
    if (part.path !== "" && part.contentType && part.path !== "[Content_Types].xml") {
      this.contentTypes.set(part.path, part.contentType);
    }
  }
}

/** `steps`, its result kept as the content of `part`. */
function* contentSteps(part: OpenPart, steps: Steps<Uint8Array>): Steps<void> {
  part.data = yield* steps;
}

/** How many bytes of a part's XML {@link standaloneXml} takes in one step. */
const PART_STEP = 1 << 20;

/**
 * How many characters of a part's base64 {@link base64DecodeSteps} takes in one step: a mebibyte of it took 4–30 ms to
 * check and decode where the platform has no native decoder.
 */
const BASE64_STEP = 256 * 1024;

/**
 * A part's XML: its content as the input holds it, made to stand alone — given the namespace declarations it uses from
 * the elements around it — in steps of {@link PART_STEP}.
 */
function* standaloneXml(
  bytes: Uint8Array,
  xml: {
    readonly from: number;
    readonly to: number;
    readonly scope: ReadonlyMap<string, string>;
    readonly root: XmlSaxTag;
    readonly rootStart: number;
  }
): Steps<Uint8Array> {
  const content = bytes.subarray(xml.from, xml.to);
  const declaration = utf8Encoder.encode(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
  );
  const own = xml.root.ns ?? {};
  // A binding the root makes itself, or one nothing in the part could use, is not repeated: a prefix is declared when
  // `prefix:` appears in the part's text, which is searched once for all of them.
  const inherited = [...xml.scope].filter(
    ([prefix]) => !Object.hasOwn(own, prefix) && prefix !== "xml"
  );
  const sought = new Set(inherited.map(([prefix]) => prefix).filter(prefix => prefix !== ""));
  const found = sought.size === 0 ? sought : yield* prefixesUsed(content, sought);
  let declarations = "";
  for (const [prefix, uri] of inherited) {
    if (prefix === "" ? uri !== "" : found.has(prefix)) {
      declarations += ` ${prefix === "" ? "xmlns" : `xmlns:${prefix}`}="${xmlEncodeAttr(uri)}"`;
    }
  }
  if (declarations === "") {
    return yield* concatSteps([declaration, trimLeadingSpace(content)]);
  }
  // After the root's name: `<name` is its start tag's first bytes.
  const at = utf8Encoder.encode(xml.root.name).length + 1 + (xml.rootStart - xml.from);
  return yield* concatSteps([
    declaration,
    trimLeadingSpace(content.subarray(0, at)),
    utf8Encoder.encode(declarations),
    content.subarray(at)
  ]);
}

/**
 * Which of `prefixes` appear as `prefix:` in `content`'s text, read in steps of {@link PART_STEP} bytes. The text
 * between steps overlaps by the longest prefix, so a name cut by a step boundary is still found.
 */
function* prefixesUsed(content: Uint8Array, prefixes: ReadonlySet<string>): Steps<Set<string>> {
  const found = new Set<string>();
  const decoder = new TextDecoder("utf-8");
  let longest = 0;
  for (const prefix of prefixes) {
    longest = Math.max(longest, prefix.length + 1);
  }
  let tail = "";
  for (let start = 0; start < content.length && found.size < prefixes.size; start += PART_STEP) {
    const end = Math.min(start + PART_STEP, content.length);
    const text =
      tail + decoder.decode(content.subarray(start, end), { stream: end < content.length });
    for (const prefix of prefixes) {
      if (!found.has(prefix) && text.includes(`${prefix}:`)) {
        found.add(prefix);
      }
    }
    tail = text.slice(-longest);
    yield;
  }
  return found;
}

/** `parts` joined, copied in steps of {@link PART_STEP} bytes. */
function* concatSteps(parts: readonly Uint8Array[]): Steps<Uint8Array> {
  let length = 0;
  for (const part of parts) {
    length += part.length;
  }
  const out = new Uint8Array(length);
  let at = 0;
  for (const part of parts) {
    for (let start = 0; start < part.length; start += PART_STEP) {
      const piece = part.subarray(start, Math.min(start + PART_STEP, part.length));
      out.set(piece, at);
      at += piece.length;
      yield;
    }
  }
  return out;
}

/**
 * A part's `binaryData` decoded, in steps of {@link BASE64_STEP} characters: XML white space anywhere is skipped,
 * and the text must otherwise be base64 — four characters at a time, `=` padding only at its very end. Each step's
 * whole groups are decoded on their own and the pieces joined at the end, each piece and the result allocated at
 * their size.
 */
function* base64DecodeSteps(text: string, path: string): Steps<Uint8Array> {
  const pieces: Uint8Array[] = [];
  let carry = "";
  let padded = false;
  for (let start = 0; start < text.length; start += BASE64_STEP) {
    const piece = carry + text.slice(start, start + BASE64_STEP).replace(XML_SPACE, "");
    if (padded && piece.length > 0) {
      // Padding ended the data; anything after it is not base64.
      throw notBase64(path);
    }
    const usable = piece.length - (piece.length % 4);
    if (usable > 0) {
      pieces.push(base64Decode(piece.slice(0, usable), path));
      padded = piece.charCodeAt(usable - 1) === 0x3d;
    }
    carry = piece.slice(usable);
    yield;
  }
  if (carry.length > 0) {
    throw notBase64(path);
  }
  return pieces.length === 1 ? pieces[0]! : yield* concatSteps(pieces);
}

/** XML's white space: the characters base64 text in a document may be broken with. */
const XML_SPACE = /[ \t\r\n]+/g;

function notBase64(path: string): DocxParseError {
  return new DocxParseError(`Flat OPC: the binaryData of part "${path}" is not base64`);
}

/** Whether `tag` is the package element `local`: in the package namespace, or — as some writers leave it — in none. */
function isPackageElement(tag: XmlSaxTag, local: string): boolean {
  return tag.local === local && (tag.uri === NS_PKG || tag.uri === "");
}

/** A package attribute of `tag`: in the package namespace, or unprefixed. */
function packageAttribute(
  tag: XmlSaxTag,
  scope: ReadonlyMap<string, string>,
  local: string
): string {
  for (const name of Object.keys(tag.attributes)) {
    const colon = name.indexOf(":");
    if (
      colon < 0
        ? name === local
        : name.slice(colon + 1) === local && scope.get(name.slice(0, colon)) === NS_PKG
    ) {
      return tag.attributes[name]!;
    }
  }
  return "";
}

function trimLeadingSpace(bytes: Uint8Array): Uint8Array {
  let i = 0;
  while (i < bytes.length && isSpaceByte(bytes[i]!)) {
    i++;
  }
  return bytes.subarray(i);
}

/**
 * Check if content is a Flat OPC XML document: its root element is `package` in the package namespace.
 * A ZIP is never one, whatever text it stores; the prologue (declaration, comments, processing
 * instructions, DOCTYPE) is skipped without parsing the rest of the document.
 */
export function isFlatOpc(content: string | Uint8Array | ArrayBuffer | ArrayBufferView): boolean {
  const given = checkTextOrBytes(content, "content", "Convert.isFlatOpc");
  const bytes = typeof given === "string" ? utf8Encoder.encode(given) : given;
  if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
    return false;
  }
  let at = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  const prologue = scanPrologue(bytes, at, bytes.length);
  if (prologue === undefined || prologue.rootStart === undefined) {
    return false;
  }
  at = prologue.rootStart;
  const end = findTagEnd(bytes, at, Math.min(bytes.length, at + MAX_ROOT_TAG_BYTES));
  if (end < 0) {
    return false;
  }
  const tag = utf8Decoder.decode(bytes.subarray(at + 1, end));
  const name = /^[^\s/>]+/.exec(tag)?.[0] ?? "";
  const colon = name.indexOf(":");
  const prefix = colon < 0 ? "" : name.slice(0, colon);
  if (name.slice(colon + 1) !== "package") {
    return false;
  }
  const attr = prefix === "" ? "xmlns" : `xmlns:${prefix}`;
  for (const match of tag.matchAll(/([^\s=/>]+)\s*=\s*("[^"]*"|'[^']*')/g)) {
    if (match[1] === attr) {
      return match[2]!.slice(1, -1) === NS_PKG;
    }
  }
  return false;
}

/** How far a root start tag is read to find its namespace declarations. */
const MAX_ROOT_TAG_BYTES = 64 * 1024;

function startsWithAscii(bytes: Uint8Array, at: number, text: string): boolean {
  for (let i = 0; i < text.length; i++) {
    if (bytes[at + i] !== text.charCodeAt(i)) {
      return false;
    }
  }
  return true;
}

function indexOfAscii(bytes: Uint8Array, text: string, from: number, to: number): number {
  for (let i = from; i + text.length <= to; i++) {
    if (startsWithAscii(bytes, i, text)) {
      return i;
    }
  }
  return -1;
}

/** The index of the `>` closing the tag that starts at `at`, outside quoted values; -1 before `to`. */
function findTagEnd(bytes: Uint8Array, at: number, to: number): number {
  let quote = 0;
  for (let i = at; i < to; i++) {
    const c = bytes[i]!;
    if (quote !== 0) {
      if (c === quote) {
        quote = 0;
      }
    } else if (c === 0x22 || c === 0x27) {
      quote = c;
    } else if (c === 0x3e) {
      return i;
    }
  }
  return -1;
}

/**
 * Walk the prologue in `[from, to)`: white space, the declaration, comments, processing
 * instructions and a DOCTYPE (internal subset included). Returns where the first start tag
 * begins (`undefined` if `to` comes first) and whether a DOCTYPE was passed, or `undefined`
 * if something else stands there.
 */
function scanPrologue(
  bytes: Uint8Array,
  from: number,
  to: number
): { rootStart: number | undefined; doctype: boolean } | undefined {
  // `doctype`: whether a DOCTYPE declares entities or attribute lists — what a part cut from the
  // document would lose.
  let at = from;
  let doctype = false;
  while (at < to) {
    const c = bytes[at]!;
    if (isSpaceByte(c)) {
      at++;
      continue;
    }
    if (c !== 0x3c) {
      return undefined;
    }
    let end: number;
    if (startsWithAscii(bytes, at, "<?")) {
      end = indexOfAscii(bytes, "?>", at + 2, to);
      end = end < 0 ? -1 : end + 2;
    } else if (startsWithAscii(bytes, at, "<!--")) {
      end = indexOfAscii(bytes, "-->", at + 4, to);
      end = end < 0 ? -1 : end + 3;
    } else if (startsWithAscii(bytes, at, "<!DOCTYPE")) {
      const skipped = skipDoctype(bytes, at + 9, to);
      doctype ||= skipped.declares;
      end = skipped.end;
    } else if (at + 1 < to && bytes[at + 1] !== 0x21 && bytes[at + 1] !== 0x3f) {
      return { rootStart: at, doctype };
    } else {
      return undefined;
    }
    if (end < 0) {
      return { rootStart: undefined, doctype };
    }
    at = end;
  }
  return { rootStart: undefined, doctype };
}

/**
 * The index after a DOCTYPE whose name starts at `at` — past its `>`, skipping quotes, comments,
 * processing instructions and its internal subset, or `-1` before `to` — and whether that subset
 * declares an entity or an attribute list: the declarations that change what an element reads as.
 */
function skipDoctype(
  bytes: Uint8Array,
  at: number,
  to: number
): { end: number; declares: boolean } {
  let quote = 0;
  let subset = false;
  let declares = false;
  for (let i = at; i < to; i++) {
    const c = bytes[i]!;
    if (quote !== 0) {
      if (c === quote) {
        quote = 0;
      }
    } else if (c === 0x22 || c === 0x27) {
      quote = c;
    } else if (subset && (startsWithAscii(bytes, i, "<!--") || startsWithAscii(bytes, i, "<?"))) {
      const comment = bytes[i + 1] === 0x21;
      const end = indexOfAscii(bytes, comment ? "-->" : "?>", i + (comment ? 4 : 2), to);
      if (end < 0) {
        return { end: -1, declares };
      }
      i = end + (comment ? 2 : 1);
    } else if (
      subset &&
      (startsWithAscii(bytes, i, "<!ENTITY") || startsWithAscii(bytes, i, "<!ATTLIST"))
    ) {
      declares = true;
    } else if (c === 0x5b) {
      subset = true;
    } else if (c === 0x5d) {
      subset = false;
    } else if (c === 0x3e && !subset) {
      return { end: i + 1, declares };
    }
  }
  return { end: -1, declares };
}

// =============================================================================
// Writing Flat OPC
// =============================================================================

/**
 * Convert a ZIP archive's entries (as a map of path → data) to Flat OPC XML string.
 *
 * @param entries - Map of part paths to their binary content.
 * @param contentTypes - Map of part paths to content types (from [Content_Types].xml).
 * @returns The Flat OPC XML string.
 */
export function toFlatOpc(
  entries: ReadonlyMap<string, Uint8Array>,
  contentTypes?: ReadonlyMap<string, string>
): string {
  if (!(entries instanceof Map)) {
    invalidArgument("Convert.partsToFlatOpc", "entries", "a Map of part name to bytes", entries);
  }
  if (contentTypes !== undefined && !(contentTypes instanceof Map)) {
    invalidArgument("Convert.partsToFlatOpc", "contentTypes", "a Map", contentTypes);
  }
  const resolvedCT = contentTypes ?? extractContentTypes(entries);
  const xml = new XmlWriter();

  xml.openXml({ version: "1.0", encoding: "UTF-8", standalone: "yes" });
  xml.openNode("pkg:package", {
    "xmlns:pkg": NS_PKG
  });

  // Write [Content_Types].xml first if present
  const ctEntry = entries.get("[Content_Types].xml");
  if (ctEntry) {
    writePartElement(
      xml,
      "/[Content_Types].xml",
      "application/vnd.openxmlformats-package.content-types+xml",
      ctEntry,
      true
    );
  }

  // Write all other parts
  for (const [path, data] of entries) {
    if (path === "[Content_Types].xml") {
      continue;
    }

    const partName = path.startsWith("/") ? path : `/${path}`;
    const ct = resolvedCT.get(path) ?? inferContentTypeFromPath(path);
    const isXml = isXmlContentType(ct) || path.endsWith(".xml") || path.endsWith(".rels");

    writePartElement(xml, partName, ct, data, isXml);
  }

  xml.closeNode(); // </pkg:package>
  return xml.xml;
}

// toFlatOpcFromDoc is exported from document-io.ts (uses dynamic import of this module)

// =============================================================================
// Internal Helpers
// =============================================================================

function writePartElement(
  xml: XmlWriter,
  partName: string,
  contentType: string,
  data: Uint8Array,
  isXml: boolean
): void {
  xml.openNode("pkg:part", {
    "pkg:name": partName,
    "pkg:contentType": contentType
  });

  if (isXml) {
    xml.openNode("pkg:xmlData", {});
    // Write raw XML content (already has XML declaration usually)
    const xmlStr = utf8Decoder.decode(data);
    // Strip XML declaration if present (Flat OPC embeds it inline)
    const stripped = xmlStr.replace(/<\?xml[^?]*\?>\s*/, "");
    xml.writeRaw(stripped);
    xml.closeNode();
  } else {
    xml.openNode("pkg:binaryData", {});
    xml.writeText(base64Encode(data));
    xml.closeNode();
  }

  xml.closeNode(); // </pkg:part>
}

/**
 * `b64` — whole groups of four, white space removed — decoded. Refused as this module's error, whichever decoder the
 * platform offers: the native one (`Uint8Array.fromBase64`, ES2026) checks the text itself, the others are checked
 * here first.
 */
function base64Decode(b64: string, path: string): Uint8Array {
  const native = (Uint8Array as NativeBase64).fromBase64;
  if (typeof native === "function") {
    try {
      return native.call(Uint8Array, b64);
    } catch {
      throw notBase64(path);
    }
  }
  if (b64.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(b64)) {
    throw notBase64(path);
  }
  if (typeof globalThis.atob === "function") {
    const binary = globalThis.atob(b64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  }
  return manualBase64Decode(b64);
}

/** `Uint8Array.fromBase64`, where the platform has it. */
interface NativeBase64 {
  readonly fromBase64?: (this: typeof Uint8Array, text: string) => Uint8Array;
}

function base64Encode(data: Uint8Array): string {
  if (typeof globalThis.btoa === "function") {
    let binary = "";
    for (let i = 0; i < data.length; i++) {
      binary += String.fromCharCode(data[i]!);
    }
    return globalThis.btoa(binary);
  }
  return manualBase64Encode(data);
}

const B64_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

function manualBase64Encode(data: Uint8Array): string {
  let result = "";
  const len = data.length;
  for (let i = 0; i < len; i += 3) {
    const b0 = data[i]!;
    const b1 = i + 1 < len ? data[i + 1]! : 0;
    const b2 = i + 2 < len ? data[i + 2]! : 0;
    result += B64_CHARS[(b0 >> 2) & 0x3f];
    result += B64_CHARS[((b0 << 4) | (b1 >> 4)) & 0x3f];
    result += i + 1 < len ? B64_CHARS[((b1 << 2) | (b2 >> 6)) & 0x3f]! : "=";
    result += i + 2 < len ? B64_CHARS[b2 & 0x3f]! : "=";
  }
  return result;
}

function manualBase64Decode(b64: string): Uint8Array {
  const lookup = new Uint8Array(128);
  for (let i = 0; i < B64_CHARS.length; i++) {
    lookup[B64_CHARS.charCodeAt(i)] = i;
  }
  const len = b64.length;
  const padLen = b64.endsWith("==") ? 2 : b64.endsWith("=") ? 1 : 0;
  const byteLen = (len * 3) / 4 - padLen;
  const result = new Uint8Array(byteLen);
  let j = 0;
  for (let i = 0; i < len; i += 4) {
    const a = lookup[b64.charCodeAt(i)]!;
    const b = lookup[b64.charCodeAt(i + 1)]!;
    const c = lookup[b64.charCodeAt(i + 2)]!;
    const d = lookup[b64.charCodeAt(i + 3)]!;
    result[j++] = (a << 2) | (b >> 4);
    if (j < byteLen) {
      result[j++] = ((b << 4) | (c >> 2)) & 0xff;
    }
    if (j < byteLen) {
      result[j++] = ((c << 6) | d) & 0xff;
    }
  }
  return result;
}

function isXmlContentType(ct: string): boolean {
  return ct.endsWith("+xml") || ct === "application/xml" || ct === "text/xml";
}

function inferContentTypeFromPath(path: string): string {
  const ext = getFileExt(path);
  return INFERRED_CONTENT_TYPES.get(ext) ?? "application/octet-stream";
}

const INFERRED_CONTENT_TYPES: ReadonlyMap<string, string> = new Map([
  ["xml", "application/xml"],
  ["rels", "application/vnd.openxmlformats-package.relationships+xml"],
  ["png", "image/png"],
  ["jpeg", "image/jpeg"],
  ["jpg", "image/jpeg"],
  ["gif", "image/gif"],
  ["svg", "image/svg+xml"],
  ["emf", "image/x-emf"],
  ["wmf", "image/x-wmf"],
  ["bin", "application/vnd.openxmlformats-officedocument.oleObject"],
  ["odttf", "application/vnd.openxmlformats-officedocument.obfuscatedFont"]
]);

/** Extract content types from [Content_Types].xml if present in entries. */
function extractContentTypes(entries: ReadonlyMap<string, Uint8Array>): Map<string, string> {
  const ctMap = new Map<string, string>();
  const ctXml = entries.get("[Content_Types].xml");
  if (!ctXml) {
    return ctMap;
  }

  const defaults = new Map<string, string>();
  // From the first `<`, as a part is read; elements are matched by namespace and local name.
  let lead = ctXml[0] === 0xef && ctXml[1] === 0xbb && ctXml[2] === 0xbf ? 3 : 0;
  while (lead < ctXml.length && isSpaceByte(ctXml[lead]!)) {
    lead++;
  }
  const parser = new SaxParser({ position: false, xmlns: true, byteOffset: lead });
  parser.on("opentag", tag => {
    if (parser.depth === 0) {
      if (tag.uri !== NS_CONTENT_TYPES || tag.local !== "Types") {
        parser.fail(`expected the root <Types> of ${NS_CONTENT_TYPES}, got <${tag.name}>`);
      }
      return;
    }
    if (parser.depth !== 1 || tag.uri !== NS_CONTENT_TYPES) {
      return;
    }
    const ct = tag.attributes.ContentType ?? "";
    if (tag.local === "Default") {
      const ext = tag.attributes.Extension ?? "";
      if (ext && ct) {
        defaults.set(ext.toLowerCase(), ct);
      }
    } else if (tag.local === "Override") {
      const partName = tag.attributes.PartName ?? "";
      if (partName && ct) {
        ctMap.set(partName.startsWith("/") ? partName.substring(1) : partName, ct);
      }
    }
  });
  try {
    parser.write(ctXml.subarray(lead)).close();
  } catch (error) {
    if (error instanceof XmlParseError) {
      throw new DocxParseError(`Flat OPC: [Content_Types].xml cannot be read: ${error.message}`, {
        cause: error
      });
    }
    throw error;
  }

  // Apply defaults to entries without explicit overrides
  for (const path of entries.keys()) {
    if (path === "[Content_Types].xml") {
      continue;
    }
    if (ctMap.has(path)) {
      continue;
    }
    const ext = getFileExt(path);
    const ct = defaults.get(ext);
    if (ct) {
      ctMap.set(path, ct);
    }
  }

  return ctMap;
}

/** Build [Content_Types].xml from collected content type info. */
function buildContentTypesXml(ctMap: Map<string, string>): Uint8Array {
  // Collect extension defaults
  const defaults = new Map<string, string>();
  defaults.set("rels", "application/vnd.openxmlformats-package.relationships+xml");
  defaults.set("xml", "application/xml");

  for (const [path, ct] of ctMap) {
    const ext = getFileExt(path);
    if (isImageExtension(ext) && !defaults.has(ext)) {
      defaults.set(ext, ct);
    }
  }

  const xml = new XmlWriter();
  xml.openXml({ version: "1.0", encoding: "UTF-8", standalone: "yes" });
  xml.openNode("Types", {
    xmlns: "http://schemas.openxmlformats.org/package/2006/content-types"
  });

  // Write Defaults first (sorted)
  const sortedDefaults = [...defaults.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  for (const [ext, ct] of sortedDefaults) {
    xml.leafNode("Default", { Extension: ext, ContentType: ct });
  }

  // Write Overrides for parts that don't match their extension default
  for (const [path, ct] of ctMap) {
    const ext = getFileExt(path);
    const defaultCt = defaults.get(ext);
    if (!defaultCt || defaultCt !== ct) {
      xml.leafNode("Override", { PartName: `/${path}`, ContentType: ct });
    }
  }

  xml.closeNode();
  return utf8Encoder.encode(xml.xml);
}

function isImageExtension(ext: string): boolean {
  return ["png", "jpeg", "jpg", "gif", "bmp", "tiff", "tif", "svg", "webp", "emf", "wmf"].includes(
    ext
  );
}

function checkTextOrBytes(value: unknown, field: string, where: string): string | Uint8Array {
  return typeof value === "string" ? value : checkBytes(value, field, where);
}
