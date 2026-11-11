/**
 * XML Encoding / Decoding Utilities
 *
 * The character-level encode/decode primitives now live in
 * `@utils/xml-encode` (Layer 0) so dependency-restricted modules — notably
 * `pdf`, which may not import from the `xml` module — can share the exact
 * same canonical implementation. This file re-exports them under the stable
 * `@xml/encode` surface and adds the name / comment validators that depend on
 * `XmlError`.
 */

import { fail } from "@utils/schema";
import {
  encodeCData,
  stripXmlIllegalChars,
  xmlDecode,
  xmlEncode,
  xmlEncodeAttr
} from "@utils/xml-encode";
import { checkStringArgument, xmlArgument } from "@xml/arguments";
import { XmlArgumentError } from "@xml/errors";
import type { XmlAttributes } from "@xml/types";

export {
  isLegalXmlChar,
  xmlDecode,
  xmlEncode,
  xmlEncodeAttr,
  stripXmlIllegalChars,
  encodeCData
} from "@utils/xml-encode";

// =============================================================================
// Name Validation
// =============================================================================

/**
 * Characters that must NEVER appear in XML element or attribute names.
 * This is a fast security check to prevent markup injection via names,
 * not a full XML NameChar validation (which would require Unicode tables).
 */
const INVALID_NAME_CHARS = /[\s<>"'/=&]/;

/** XML 1.0 (5th ed.) §2.3 NameStartChar, without ":" (Namespaces in XML: an NCName's). */
const NC_NAME_START =
  "A-Z_a-z\\u00C0-\\u00D6\\u00D8-\\u00F6\\u00F8-\\u02FF\\u0370-\\u037D\\u037F-\\u1FFF\\u200C-\\u200D" +
  "\\u2070-\\u218F\\u2C00-\\u2FEF\\u3001-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFFD\\u{10000}-\\u{EFFFF}";
/** §2.3 NameChar, without ":". */
const NC_NAME_CHAR = `${NC_NAME_START}\\-.0-9\\u00B7\\u0300-\\u036F\\u203F-\\u2040`;
const NC_NAME = `[${NC_NAME_START}][${NC_NAME_CHAR}]*`;
/**
 * A QName (Namespaces in XML 1.0 §4): an NCName, or two joined by one ":". Every name these
 * writers produce is read by a namespace-aware consumer, which refuses `a:b:c` or `:a`.
 */
const QNAME_RE = new RegExp(`^${NC_NAME}(?::${NC_NAME})?$`, "u");

/** Names that have passed {@link validateXmlName}, so the common ones are tested once. */
const VALID_NAMES = new Set<string>();
const VALID_NAMES_CAP = 4096;

/**
 * Validate an XML element or attribute name.
 *
 * Accepts a QName: an XML 1.0 `Name` (§2.3) with at most one `:`, between two non-empty
 * parts, as Namespaces in XML requires. Anything else is refused with an {@link XmlError} —
 * including `!` and `?`, which would let a "name" open a comment, a CDATA section or a
 * processing instruction (`<!--x>`), and control characters.
 */
export function validateXmlName(name: string): void {
  validateXmlNameAt(name, "name");
}

/**
 * {@link validateXmlName}, reporting `field` on refusal — an attribute name is checked as
 * `attributes.<key>` of the call that passed the attribute map.
 */
export function validateXmlNameAt(name: string, field: string): void {
  // A writer names the same few elements and attributes over and over — `c`, `r`, `s`, `t`, `v` for every cell of a
  // sheet — so a name that has passed once is not tested again. Capped, so names a caller invents cannot grow it
  // without bound; past the cap a name is simply checked every time, as before.
  // The last two names first: a writer alternates between few (a row and its cell), and comparing to
  // them is cheaper than a set lookup.
  if (name === recentName || name === otherRecentName) {
    return;
  }
  if (!VALID_NAMES.has(name)) {
    checkXmlName(name, field);
    if (VALID_NAMES.size < VALID_NAMES_CAP) {
      VALID_NAMES.add(name);
    }
  }
  otherRecentName = recentName;
  recentName = name;
}

/** The two names {@link validateXmlNameAt} passed last; at first a name that would pass. */
let recentName = "_";
let otherRecentName = "_";

function checkXmlName(name: string, field: string): void {
  if (typeof name !== "string") {
    throw new XmlArgumentError(
      field,
      `XML name must be a string, got ${name === null ? "null" : typeof name}`
    );
  }
  if (!name) {
    throw new XmlArgumentError(field, "XML name must not be empty");
  }
  if (INVALID_NAME_CHARS.test(name)) {
    throw new XmlArgumentError(
      field,
      `Invalid XML name: contains forbidden character in "${name}"`
    );
  }
  // XML names cannot start with a digit, hyphen, or dot
  const first = name.charCodeAt(0);
  if ((first >= 0x30 && first <= 0x39) || first === 0x2d || first === 0x2e) {
    throw new XmlArgumentError(
      field,
      `Invalid XML name: "${name}" starts with forbidden character`
    );
  }
  // The full production: a `!` or `?` let a name open a comment, a CDATA section or a
  // processing instruction (`<!--x>`), and a control character made the document unreadable.
  if (!QNAME_RE.test(name)) {
    throw new XmlArgumentError(
      field,
      `Invalid XML name: "${name}" is not an XML name (XML 1.0 Name, with at most one ":" between two parts)`
    );
  }
}

/**
 * Stringify an attribute value or text node, refusing a non-finite number and
 * anything that is not a string, number or boolean.
 *
 * `String(NaN)` is `"NaN"` and `String(Infinity)` is `"Infinity"`, and neither is a
 * legal lexical form for any numeric XSD type — so a non-finite number reaching a
 * writer always produces a document its consumer rejects. It was reached through
 * `IntegerXform`, whose `if (model || this.zero)` treats `Infinity` as a value worth
 * writing, and it emitted `<sz val="Infinity"/>` into a stylesheet Excel then refused
 * to open. Guarding here rather than in that one xform covers every other raw numeric
 * attribute the same way — a colour's `theme`/`tint`, a gradient's `degree`, a stop's
 * `position` — instead of waiting for each to be reported separately.
 *
 * Throwing, rather than omitting the attribute, matches {@link validateXmlName}: this
 * writer already refuses input it cannot represent instead of quietly emitting
 * something broken. The caller gets a stack trace at the point of the write, which is
 * more use than a file that opens to a repair dialog.
 */
export function xmlValue(
  value: string | number | boolean,
  name: string,
  field: string = name === "#text" ? "text" : name
): string {
  // The type admits only these three, but a model built by hand or from JSON is
  // not checked by the compiler: `null` became `"null"` and an object
  // `"[object Object]"` — Word refused a drawing written with `cx="[object Object]"`.
  switch (typeof value) {
    case "string":
      return value;
    case "boolean":
      return String(value);
    case "number":
      if (Number.isFinite(value)) {
        return String(value);
      }
      throw new XmlArgumentError(
        field,
        `Invalid XML value for "${name}": ${String(value)} has no valid XML representation`
      );
    default:
      throw new XmlArgumentError(
        field,
        `Invalid XML value for "${name}": ${value === null ? "null" : `a value of type ${typeof value}`} has no XML representation`
      );
  }
}

// =============================================================================
// Comment Validation
// =============================================================================

/** Which XML comment rule `text` breaks (XML 1.0 §2.5), if any. */
function commentTextProblem(text: string): string | undefined {
  if (text.includes("--")) {
    return 'it contains "--"';
  }
  if (text.endsWith("-")) {
    return 'it ends with "-"';
  }
  return undefined;
}

/**
 * Validate that text is legal for an XML comment.
 *
 * XML spec: comments must not contain `--` and must not end with `-`.
 * @throws {XmlError} if the text is invalid.
 */
export function validateCommentText(text: string): void {
  if (typeof text !== "string") {
    throw new XmlArgumentError(
      "text",
      `Invalid comment: expected a string, got ${text === null ? "null" : typeof text}`
    );
  }
  const problem = commentTextProblem(text);
  if (problem !== undefined) {
    throw new XmlArgumentError("text", `Invalid comment: ${problem}`);
  }
}

function entryNameField(name: string, inMap: boolean): string {
  return inMap ? `attributes.${name}` : "name";
}

function entryValueField(name: string, inMap: boolean): string {
  return inMap ? `attributes.${name}` : "value";
}

/**
 * An attribute's value for a writer: the name checked, the value stringified and — for a namespace
 * declaration — cleaned of the characters XML cannot hold and only then checked, so what is checked is
 * what is written. `inMap`: whether it came in an attribute map, and is reported as `attributes.<name>`.
 */
export function xmlAttributeText(
  name: string,
  value: string | number | boolean,
  inMap: boolean
): string {
  if (!VALID_NAMES.has(name)) {
    validateXmlNameAt(name, entryNameField(name, inMap));
  }
  let text =
    typeof value === "string" ? value : xmlValue(value, name, entryValueField(name, inMap));
  if (name.charCodeAt(0) === 0x78 && (name === "xmlns" || name.startsWith("xmlns:"))) {
    text = stripXmlIllegalChars(text);
    checkNamespaceDeclaration(name, text, entryValueField(name, inMap));
  }
  return text;
}

const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE = "http://www.w3.org/2000/xmlns/";

/**
 * Refuse a namespace declaration no document can hold, wherever it is put (Namespaces in XML 1.0 §3):
 * `xml` bound elsewhere than its namespace, `xmlns` declared, either namespace bound to another
 * prefix or made the default, or a prefix bound to no namespace. A prefix the writer is not told the
 * binding of is let be: what it writes may be part of a document that declares it.
 */
function checkNamespaceDeclaration(name: string, value: string, field: string): void {
  let problem: string | undefined;
  if (name === "xmlns") {
    if (value === XML_NAMESPACE || value === XMLNS_NAMESPACE) {
      problem = `the namespace ${value} cannot be the default namespace`;
    }
  } else if (name.startsWith("xmlns:")) {
    const prefix = name.slice(6);
    if (prefix === "xmlns") {
      problem = "the xmlns prefix cannot be declared";
    } else if (prefix === "xml" ? value !== XML_NAMESPACE : value === XML_NAMESPACE) {
      problem = `the xml prefix and the namespace ${XML_NAMESPACE} are bound to each other only`;
    } else if (value === XMLNS_NAMESPACE) {
      problem = `the namespace ${XMLNS_NAMESPACE} cannot be bound to a prefix`;
    } else if (value === "") {
      problem = `the prefix ${prefix} cannot be bound to an empty namespace name`;
    }
  }
  if (problem !== undefined) {
    throw new XmlArgumentError(
      field,
      `Invalid namespace declaration ${name}="${value}": ${problem}`
    );
  }
}

/** The prefix of an attribute that is not a namespace declaration, or `undefined` when it has none. */
function attributePrefix(name: string): string | undefined {
  const colon = name.indexOf(":");
  return colon > 0 && !name.startsWith("xmlns:") ? name.slice(0, colon) : undefined;
}

const hasOwnProperty = Object.prototype.hasOwnProperty;

/** A writer's `busy` flag: it is reading the caller's attribute map, whose getters must not write. */
export const BUSY_READING = 1;

/** What the attributes a writer read hold, as far as namespaces go. */
const DECLARES = 1;
const PREFIXED = 2;

/** A name in markup this writer wrote: its values hold no `"`, so each name is the text before `="`. */
const WRITTEN_NAME = / ([^\s="]+)="/g;

interface DeclarationScope {
  /** The depth of the element that declares, counting the document element as 1. */
  readonly depth: number;
  readonly declarations: Map<string, string>;
}

/** What {@link XmlTagStack.snapshot} keeps, for a rollback. */
export interface XmlTagStackSnapshot {
  readonly elements: string[];
  readonly scopes: DeclarationScope[];
  readonly openMarkup: string;
}

/**
 * The open elements of a writer, and what the open start tag holds — the state a writer checks
 * attributes against. A start tag refuses an attribute it already carries by name ("Unique Att Spec")
 * and, where the writer knows both prefixes' bindings, by expanded name (Namespaces in XML 1.0 §6.3:
 * `p:a` and `q:a` with `p` and `q` bound to one namespace). A prefix bound by no declaration the writer
 * saw is let be: what it writes may be part of a document that declares it.
 *
 * Built for the common case, which is most of a writer's work: a tag whose attributes are unprefixed
 * and declare nothing is read, checked and serialised in one pass, with nothing allocated beyond its
 * markup. An attribute map cannot repeat a name, so a new tag needs no names kept; they are read back
 * from the tag's own markup only when a later call adds to it. Namespace declarations are kept only for
 * the elements that make them.
 */
export class XmlTagStack {
  /** The open elements' names, outermost first. */
  elements: string[] = [];
  /**
   * The writer the stack belongs to, whose `busy` flags record what it is in the middle of: a read of
   * the caller's attribute map sets {@link BUSY_READING} there, so the check the writer makes on every
   * call reads one field of its own.
   */
  constructor(private readonly owner: { busy: number }) {}
  private scopes: DeclarationScope[] = [];
  /*
   * What is kept about an element is kept with its depth (the document element's is 1), and is let
   * go as the element at that depth closes: opening a tag with no attributes — most tags — touches
   * none of it, and closing one compares one number, `held`.
   */
  /** The depth of the element that made the innermost scope, or `-1` when there is none. */
  private scopeDepth = -1;
  /**
   * The attributes of the start tag at `markupDepth`, as markup: the open tag's when that is the
   * innermost element's depth, and none otherwise. Its names are read back from it when needed. A tag
   * that is closed, or has an element opened inside it, takes no more.
   */
  private openMarkup = "";
  private markupDepth = -1;
  /**
   * What a later call adding to the open tag has had to know about it — its names, and its prefixed
   * names by expanded name — built the first time one does, for the tag at depth `of`.
   */
  private openTag: { of: number; names: Set<string>; expanded?: Map<string, string> } | undefined;
  /** The deepest depth anything above is kept for, or `-1`. */
  private held = -1;
  /**
   * The attributes the last read held, name then text, in the first `readLength` slots: the array is
   * reused and never shrunk, so a read allocates none.
   */
  private readonly read: string[] = [];
  private readLength = 0;
  private readFlags = 0;

  get depth(): number {
    return this.elements.length;
  }

  get current(): string | undefined {
    return this.elements[this.elements.length - 1];
  }

  /**
   * Read, check and serialise an attribute map, with `BUSY_READING` set while its getters run. Nothing of
   * the stack changes: what it read is kept for the call that follows.
   */
  private readMap(
    attributes: Readonly<Record<string, string | number | boolean | undefined>> | undefined
  ): string {
    this.readLength = 0;
    this.readFlags = 0;
    if (attributes === undefined || attributes === null) {
      return "";
    }
    if (typeof attributes !== "object") {
      throw new XmlArgumentError(
        "attributes",
        `Invalid XML attributes: expected an object, got ${typeof attributes}`
      );
    }
    const read = this.read;
    let s = "";
    let n = 0;
    // Every call that reads has found its writer not busy, so the flag is set and cleared outright.
    const owner = this.owner;
    owner.busy = BUSY_READING;
    try {
      for (const key in attributes) {
        // Its own attributes only: one a map inherits is not the caller's. V8 answers this form in a
        // for-in from the object's enum cache, where `Object.hasOwn` is a full lookup.
        if (!hasOwnProperty.call(attributes, key)) {
          continue;
        }
        // Each value is read once: a getter read again could answer otherwise.
        const value = attributes[key];
        if (value === undefined) {
          continue;
        }
        if (!VALID_NAMES.has(key)) {
          validateXmlNameAt(key, `attributes.${key}`);
        }
        read[n] = key;
        n += 2;
        // The common attribute — unprefixed, no declaration, a string or a number — is written here;
        // the rest by `readOther`, so this loop stays small enough to be compiled well.
        if (key.includes(":") || key === "xmlns") {
          s += this.readOther(key, value, n - 1);
        } else if (typeof value === "string") {
          read[n - 1] = value;
          s += ` ${key}="${xmlEncodeAttr(value)}"`;
        } else if (typeof value === "number" && Number.isFinite(value)) {
          // A number's text holds nothing to escape.
          const text = String(value);
          read[n - 1] = text;
          s += ` ${key}="${text}"`;
        } else {
          s += this.readOther(key, value, n - 1);
        }
      }
    } catch (error) {
      owner.busy = 0;
      throw error;
    }
    owner.busy = 0;
    this.readLength = n;
    return s;
  }

  /** {@link readMap} for an attribute it does not write itself: its text goes to `read[at]`. */
  private readOther(key: string, value: string | number | boolean, at: number): string {
    let text = xmlValue(value, key, `attributes.${key}`);
    if (key.includes(":")) {
      if (key.startsWith("xmlns:")) {
        text = stripXmlIllegalChars(text);
        checkNamespaceDeclaration(key, text, `attributes.${key}`);
        this.readFlags |= DECLARES;
      } else {
        this.readFlags |= PREFIXED;
      }
    } else if (key === "xmlns") {
      text = stripXmlIllegalChars(text);
      checkNamespaceDeclaration(key, text, `attributes.${key}`);
    }
    this.read[at] = text;
    return ` ${key}="${xmlEncodeAttr(text)}"`;
  }

  /** The declarations among what was read last, or `undefined` for none. */
  private readDeclarations(base: Map<string, string> | undefined): Map<string, string> | undefined {
    if ((this.readFlags & DECLARES) === 0) {
      return base;
    }
    const declarations = new Map(base);
    const read = this.read;
    for (let i = 0; i < this.readLength; i += 2) {
      if (read[i].startsWith("xmlns:")) {
        declarations.set(read[i].slice(6), read[i + 1]);
      }
    }
    return declarations;
  }

  private resolve(prefix: string, own: Map<string, string> | undefined): string | undefined {
    const bound = own?.get(prefix);
    if (bound !== undefined) {
      return bound;
    }
    for (let i = this.scopes.length - 1; i >= 0; i--) {
      const outer = this.scopes[i].declarations.get(prefix);
      if (outer !== undefined) {
        return outer;
      }
    }
    return prefix === "xml" ? XML_NAMESPACE : undefined;
  }

  /** Record `name`'s expanded name in `expanded`, refusing one already there. */
  private checkExpanded(
    element: string,
    name: string,
    field: string,
    own: Map<string, string> | undefined,
    expanded: Map<string, string>
  ): void {
    const prefix = attributePrefix(name);
    if (prefix === undefined) {
      return;
    }
    const namespace = this.resolve(prefix, own);
    if (namespace === undefined) {
      return;
    }
    const key = `{${namespace}}${name.slice(prefix.length + 1)}`;
    const other = expanded.get(key);
    if (other !== undefined) {
      throw new XmlArgumentError(
        field,
        `Duplicate attribute ${key} on element <${element}>: "${other}" and "${name}" name the same attribute`
      );
    }
    expanded.set(key, name);
  }

  /**
   * The attribute markup of a new start tag `name` — one to open or a leaf — checked against the open
   * elements' namespaces. Nothing changes until {@link open}.
   */
  startTag(
    name: string,
    attributes: Readonly<Record<string, string | number | boolean | undefined>> | undefined
  ): string {
    if (attributes === undefined) {
      // A tag with no attributes — a good part of what a writer writes — reads nothing. Only the flags
      // are read before the next read: by `open`.
      this.readFlags = 0;
      return "";
    }
    const markup = this.readMap(attributes);
    if ((this.readFlags & PREFIXED) !== 0) {
      this.checkReadPrefixes(name);
    }
    return markup;
  }

  /** Refuse two of the attributes read last that have one expanded name. */
  private checkReadPrefixes(name: string): void {
    const read = this.read;
    const length = this.readLength;
    // Two names of one map differ, so two prefixed ones can name one attribute only with one local
    // part under two prefixes: only such a pair has its prefixes resolved. A tag has a handful of
    // prefixed attributes, and nearly every one has none of these pairs.
    if (length > 128) {
      // A long list is compared by expanded name in one pass, not pair by pair.
      const declarations = this.readDeclarations(undefined);
      const expanded = new Map<string, string>();
      for (let i = 0; i < length; i += 2) {
        this.checkExpanded(name, read[i], `attributes.${read[i]}`, declarations, expanded);
      }
      return;
    }
    let own: Map<string, string> | undefined;
    let resolved = false;
    for (let i = 0; i < length; i += 2) {
      const a = read[i];
      const colonA = a.indexOf(":");
      if (colonA <= 0 || a.startsWith("xmlns:")) {
        continue;
      }
      for (let j = i + 2; j < length; j += 2) {
        const b = read[j];
        const colonB = b.indexOf(":");
        if (
          colonB <= 0 ||
          b.length - colonB !== a.length - colonA ||
          !b.endsWith(a.slice(colonA)) ||
          b.startsWith("xmlns:")
        ) {
          continue;
        }
        if (!resolved) {
          own = this.readDeclarations(undefined);
          resolved = true;
        }
        const namespace = this.resolve(a.slice(0, colonA), own);
        if (namespace !== undefined && namespace === this.resolve(b.slice(0, colonB), own)) {
          throw new XmlArgumentError(
            `attributes.${b}`,
            `Duplicate attribute {${namespace}}${a.slice(colonA + 1)} on element <${name}>: "${a}" and "${b}" name the same attribute`
          );
        }
      }
    }
  }

  /** Open element `name` with the attribute markup {@link startTag} returned for it. */
  open(name: string, markup: string): void {
    const depth = this.elements.push(name);
    if (markup !== "") {
      this.openMarkup = markup;
      this.markupDepth = depth;
      this.held = depth;
    }
    if ((this.readFlags & DECLARES) !== 0) {
      this.scopes.push({ depth, declarations: this.readDeclarations(undefined)! });
      this.scopeDepth = depth;
      this.held = depth;
    }
  }

  /** {@link open} for a tag with no attributes, which {@link startTag} need not be asked about. */
  openPlain(name: string): void {
    this.elements.push(name);
  }

  /** Close the innermost element and return its name, or `undefined` when none is open. */
  close(): string | undefined {
    if (this.held >= this.elements.length) {
      this.release(this.elements.length);
    }
    return this.elements.pop();
  }

  /** Let go of what is kept for the element at `depth`, which is closing. */
  private release(depth: number): void {
    if (this.scopeDepth === depth) {
      const scopes = this.scopes;
      scopes.pop();
      this.scopeDepth = scopes.length > 0 ? scopes[scopes.length - 1].depth : -1;
    }
    if (this.markupDepth === depth) {
      this.markupDepth = -1;
    }
    if (this.openTag?.of === depth) {
      this.openTag = undefined;
    }
    this.held = Math.max(this.scopeDepth, this.markupDepth, this.openTag?.of ?? -1);
  }

  /** The open tag's attributes, as markup. */
  private currentMarkup(): string {
    return this.markupDepth === this.elements.length ? this.openMarkup : "";
  }

  /** The open tag's own declarations: the innermost scope, when the open tag made it. */
  private ownDeclarations(): Map<string, string> | undefined {
    return this.scopeDepth === this.elements.length
      ? this.scopes[this.scopes.length - 1].declarations
      : undefined;
  }

  /** Add an attribute map to the open start tag; returns its markup. Refused, it adds nothing. */
  addMap(
    attributes: Readonly<Record<string, string | number | boolean | undefined>> | undefined
  ): string {
    const markup = this.readMap(attributes);
    this.addRead(markup, true);
    return markup;
  }

  /** Add one attribute to the open start tag; returns its markup. Refused, it adds nothing. */
  addOne(name: string, value: string | number | boolean): string {
    const text = xmlAttributeText(name, value, false);
    const read = this.read;
    read[0] = name;
    read[1] = text;
    this.readLength = 2;
    this.readFlags = name.includes(":") ? (name.startsWith("xmlns:") ? DECLARES : PREFIXED) : 0;
    const markup = ` ${name}="${xmlEncodeAttr(text)}"`;
    this.addRead(markup, false);
    return markup;
  }

  /** The open tag's names, read back from its markup the first time a call adds to it. */
  private names(): Set<string> {
    const depth = this.elements.length;
    let tag = this.openTag;
    if (tag === undefined || tag.of !== depth) {
      const names = new Set<string>();
      for (const match of this.currentMarkup().matchAll(WRITTEN_NAME)) {
        names.add(match[1]);
      }
      tag = this.openTag = { of: depth, names };
      this.held = Math.max(this.held, depth);
    }
    return tag.names;
  }

  /** Check what was read last as additions to the open tag, then record it. */
  private addRead(markup: string, inMap: boolean): void {
    const read = this.read;
    if (this.readLength === 0) {
      return;
    }
    const element = this.elements[this.elements.length - 1];
    const names = this.names();
    for (let i = 0; i < this.readLength; i += 2) {
      // An attribute map cannot repeat a name, so only the tag's are compared.
      if (names.has(read[i]) || (!inMap && i > 0 && read[i] === read[0])) {
        throw new XmlArgumentError(
          entryNameField(read[i], inMap),
          `Duplicate attribute "${read[i]}" on element <${element}>`
        );
      }
    }
    const flags = this.readFlags;
    const own = this.ownDeclarations();
    let declarations = own;
    let expanded = this.openTag!.expanded;
    if ((flags & DECLARES) !== 0) {
      declarations = this.readDeclarations(declarations);
      // A declaration binds the prefixes of attributes written before it too: check the whole tag again.
      expanded = new Map();
      for (const name of names) {
        this.checkExpanded(element, name, "attributes", declarations, expanded);
      }
      for (let i = 0; i < this.readLength; i += 2) {
        this.checkExpanded(
          element,
          read[i],
          entryNameField(read[i], inMap),
          declarations,
          expanded
        );
      }
    } else if ((flags & PREFIXED) !== 0) {
      if (expanded === undefined) {
        expanded = new Map();
        for (const name of names) {
          this.checkExpanded(element, name, "attributes", declarations, expanded);
        }
      } else {
        expanded = new Map(expanded);
      }
      for (let i = 0; i < this.readLength; i += 2) {
        this.checkExpanded(
          element,
          read[i],
          entryNameField(read[i], inMap),
          declarations,
          expanded
        );
      }
    }
    // Checked: now recorded.
    for (let i = 0; i < this.readLength; i += 2) {
      names.add(read[i]);
    }
    this.openMarkup = this.currentMarkup() + markup;
    this.markupDepth = this.elements.length;
    this.held = Math.max(this.held, this.markupDepth);
    this.openTag!.expanded = expanded;
    if (declarations !== own) {
      const scopes = this.scopes;
      const depth = this.elements.length;
      if (own !== undefined) {
        scopes[scopes.length - 1] = { depth, declarations: declarations! };
      } else {
        scopes.push({ depth, declarations: declarations! });
        this.scopeDepth = depth;
        this.held = Math.max(this.held, depth);
      }
    }
  }

  /** The state a rollback returns to. */
  snapshot(): XmlTagStackSnapshot {
    return {
      elements: this.elements.slice(),
      scopes: this.scopes.slice(),
      openMarkup: this.currentMarkup()
    };
  }

  restore(snapshot: XmlTagStackSnapshot): void {
    this.elements = snapshot.elements.slice();
    const scopes = (this.scopes = snapshot.scopes.slice());
    this.scopeDepth = scopes.length > 0 ? scopes[scopes.length - 1].depth : -1;
    // The open tag, if one is, as it was; what was kept about it since is built again when asked for.
    this.openMarkup = snapshot.openMarkup;
    this.markupDepth = snapshot.openMarkup === "" ? -1 : this.elements.length;
    this.openTag = undefined;
    this.held = Math.max(this.scopeDepth, this.markupDepth);
  }

  reset(): void {
    this.elements = [];
    this.scopes = [];
    this.scopeDepth = -1;
    this.openMarkup = "";
    this.markupDepth = -1;
    this.openTag = undefined;
    this.held = -1;
  }
}

/**
 * A comment's markup. Characters XML cannot carry are removed, as they are from text and
 * CDATA; what remains must not contain `--` or end with `-`.
 */
export function commentMarkup(text: string): string {
  if (typeof text !== "string") {
    throw new XmlArgumentError(
      "text",
      `Invalid comment: expected a string, got ${text === null ? "null" : typeof text}`
    );
  }
  const clean = stripXmlIllegalChars(text);
  const problem = commentTextProblem(clean);
  if (problem !== undefined) {
    // Name the rule that failed, and say when it failed only after characters XML cannot hold
    // were removed (`"a\u0001-"` ends with "-" once the control character is gone).
    throw new XmlArgumentError(
      "text",
      `Invalid comment: ${problem}${clean === text ? "" : " once the characters XML cannot hold are removed"}`
    );
  }
  return `<!--${clean}-->`;
}

/** A CDATA section's markup; `]]>` inside the text is split across two sections. */
export function cdataMarkup(text: string): string {
  if (typeof text !== "string") {
    throw new XmlArgumentError(
      "text",
      `Invalid CDATA: expected a string, got ${text === null ? "null" : typeof text}`
    );
  }
  return encodeCData(text);
}

/** Raw markup for `writeRaw`, refused unless it is a string — `undefined` was written as nothing, `42` as `42`. */
export function rawMarkup(xml: string): string {
  if (typeof xml !== "string") {
    throw new XmlArgumentError(
      "xml",
      `Invalid raw XML: expected a string, got ${xml === null ? "null" : typeof xml}`
    );
  }
  return xml;
}

// =============================================================================
// Standard XML Declaration
// =============================================================================

/** Default XML declaration attributes (`version`, `encoding`, `standalone`). */
export const StdDocAttributes: Readonly<Record<string, string>> = {
  version: "1.0",
  encoding: "UTF-8",
  standalone: "yes"
};

const XML_VERSION_RE = /^1\.[0-9]+$/;
/** XML 1.0 §4.3.3 EncName. */
const ENC_NAME_RE = /^[A-Za-z][A-Za-z0-9._-]*$/;

/**
 * The XML declaration's markup for `openXml`: `attributes` over {@link StdDocAttributes}, in
 * the order XML 1.0 §2.8 requires. Only `version` (an XML 1.x `VersionNum`), `encoding` (an
 * `EncName`) and `standalone` (`"yes"`, `"no"` or a boolean) are accepted; an omitted version
 * takes the default, an `undefined` encoding or standalone leaves it out.
 */
export function xmlDeclarationMarkup(where: string, attributes: XmlAttributes | undefined): string {
  const options = xmlArgument(where, "attributes");
  if (attributes === undefined) {
    attributes = StdDocAttributes;
  } else if (typeof attributes !== "object" || attributes === null) {
    throw fail(options, "", "an object of version, encoding and standalone", attributes);
  }
  for (const key in attributes) {
    if (key !== "version" && key !== "encoding" && key !== "standalone") {
      throw fail(
        options,
        `.${key}`,
        "absent (an XML declaration has only version, encoding and standalone)",
        attributes[key]
      );
    }
  }
  const merged =
    attributes === StdDocAttributes ? attributes : { ...StdDocAttributes, ...attributes };
  const version = merged.version ?? StdDocAttributes.version;
  if (typeof version !== "string" || !XML_VERSION_RE.test(version)) {
    throw fail(options, ".version", 'an XML 1.x version string such as "1.0"', version);
  }
  let s = `<?xml version="${version}"`;
  const encoding = merged.encoding;
  if (encoding !== undefined) {
    if (typeof encoding !== "string" || !ENC_NAME_RE.test(encoding)) {
      throw fail(options, ".encoding", 'an encoding name such as "UTF-8"', encoding);
    }
    s += ` encoding="${encoding}"`;
  }
  const standalone = merged.standalone;
  if (standalone !== undefined) {
    const text =
      standalone === true || standalone === "yes"
        ? "yes"
        : standalone === false || standalone === "no"
          ? "no"
          : undefined;
    if (text === undefined) {
      throw fail(options, ".standalone", '"yes", "no" or a boolean', standalone);
    }
    s += ` standalone="${text}"`;
  }
  return `${s}?>\n`;
}

// =============================================================================
// Checked public forms
// =============================================================================

function checkText(where: string, text: unknown): asserts text is string {
  checkStringArgument(where, "text", text);
}

/**
 * `Xml.encode`: {@link xmlEncode}, refusing a value that is not a string. The unchecked form
 * returned a number unchanged and threw a `TypeError` on `null`.
 */
export function encodeText(text: string): string {
  checkText("Xml.encode", text);
  return xmlEncode(text);
}

/** `Xml.decode`: {@link xmlDecode}, refusing a value that is not a string. */
export function decodeText(text: string): string {
  checkText("Xml.decode", text);
  return xmlDecode(text);
}

/** `Xml.encodeAttr`: {@link xmlEncodeAttr}, refusing a value that is not a string. */
export function encodeAttributeValue(value: string): string {
  checkStringArgument("Xml.encodeAttr", "value", value);
  return xmlEncodeAttr(value);
}

/** `Xml.encodeCData`: {@link encodeCData}, refusing a value that is not a string. */
export function encodeCDataSection(text: string): string {
  checkText("Xml.encodeCData", text);
  return encodeCData(text);
}
