/**
 * Namespace prefixes, normalised on read — for a format whose readers match names by prefix.
 *
 * An Office reader matches element names with the prefix its format's producer writes — `c:chartSpace`,
 * `w:p`, an unprefixed `worksheet` — but a prefix is only a local alias for a namespace URI. A producer
 * may bind the chart namespace to `cs`, or make it the default namespace and write `<chartSpace>`, and
 * the part means exactly the same thing. Matching on the literal prefix made such a part unreadable —
 * and let an element of another namespace pass for one of the format's under the prefix it borrowed.
 *
 * So the reader rewrites names before it matches them: each element and prefixed attribute is renamed
 * to the **canonical prefix of its namespace URI** (the one the reader expects, and the writer emits),
 * as the format's {@link PrefixVocabulary} gives it, its `xmlns` declarations are rewritten to bind
 * those prefixes, and the attribute values that name prefixes are rewritten with them: the MCE lists
 * (`mc:Ignorable`, `mc:MustUnderstand`, `mc:Choice/@Requires`, and the QName lists
 * `mc:ProcessContent`, `mc:PreserveElements`, `mc:PreserveAttributes`) and `xsi:type`, each recognised
 * by the namespace its name resolves to. No other attribute in these parts names a prefix —
 * `xsi:schemaLocation` and the `uri` of an `ext` or a `graphicData` hold URIs — so every other value is
 * kept. A part that already uses the canonical prefixes — every part Office or this library writes —
 * passes through untouched, so the common case costs a scan of each element's attribute names for an
 * `xmlns` declaration.
 *
 * The tags a reader matches are the **name** view. Anything written back out as text uses the **text**
 * view, which also keeps every unprefixed name in the namespace it was in where a prefixed namespace
 * becoming the default one would otherwise capture it — `canonicalizeXml` for a whole part
 * (`forText`), and a raw-captured subtree through `textOpened`.
 * A namespace the vocabulary does not know keeps the prefix the document gave it, unless a known
 * namespace in the same scope is being given that prefix — or the vocabulary reserves its prefixes —
 * in which case it is renamed.
 */
import { escapeRegExp } from "@utils/regexp";
import { xmlEncode, xmlEncodeAttr } from "@utils/xml-encode";
import { SaxParser } from "@xml/sax";
import type { ParseOpenTag, XmlParseOptions } from "@xml/types";

/**
 * The namespaces a document format's readers match by prefix: each known namespace URI and the prefix
 * its readers expect (`""` is "unprefixed"). Several URIs may share a prefix — a Transitional namespace
 * and its Strict counterpart, say. Made once per format with {@link createPrefixVocabulary}.
 */
export interface PrefixVocabulary {
  /** Namespace URI → canonical prefix. */
  readonly prefixOf: ReadonlyMap<string, string>;
  /** Every canonical prefix. */
  readonly prefixes: ReadonlySet<string>;
  /** Canonical prefix (other than `""`) → the namespace it stands for, the first URI given it. */
  readonly namespaceOf: ReadonlyMap<string, string>;
  /**
   * Whether a namespace the vocabulary does not know is renamed when it uses a prefix the vocabulary
   * gives a known one, even with that known one nowhere in scope — so a reader matching that prefix is
   * never shown an element of another namespace under it. The default namespace is never renamed for
   * this. Off, an unknown namespace is renamed only when a known one in scope takes its prefix.
   */
  readonly reservePrefixes: boolean;
}

export function createPrefixVocabulary(
  prefixOf: ReadonlyMap<string, string>,
  options?: { readonly reservePrefixes?: boolean }
): PrefixVocabulary {
  const namespaceOf = new Map<string, string>();
  for (const [uri, prefix] of prefixOf) {
    if (prefix !== "" && !namespaceOf.has(prefix)) {
      namespaceOf.set(prefix, uri);
    }
  }
  return {
    prefixOf,
    prefixes: new Set(prefixOf.values()),
    namespaceOf,
    reservePrefixes: options?.reservePrefixes ?? false
  };
}

const MCE_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const XSI_NAMESPACE = "http://www.w3.org/2001/XMLSchema-instance";

/** How an attribute's value names namespaces — and so how it is rewritten with them. */
const enum ValueKind {
  /** Names none: `xsi:schemaLocation`, `ext/@uri`, `graphicData/@uri`, VML's `v:ext`, … are kept. */
  None,
  /** A whitespace-separated list of prefixes. */
  Prefixes,
  /** A whitespace-separated list of prefixed QNames, a local name of `*` meaning "every name". */
  QNames,
  /** One QName, an unprefixed one in the default namespace. */
  QName
}

/**
 * The MCE attributes (ECMA-376 Part 3 §10), keyed by local name: `Ignorable` and `MustUnderstand` list
 * prefixes; `ProcessContent`, `PreserveElements` and `PreserveAttributes` list `prefix:local` or
 * `prefix:*`. Each is one only when its prefix is bound to the MCE namespace, whatever the prefix is.
 */
const MCE_ATTRIBUTES: ReadonlyMap<string, ValueKind> = new Map([
  ["Ignorable", ValueKind.Prefixes],
  ["MustUnderstand", ValueKind.Prefixes],
  ["ProcessContent", ValueKind.QNames],
  ["PreserveElements", ValueKind.QNames],
  ["PreserveAttributes", ValueKind.QNames]
]);

interface Scope {
  /** The format's namespaces: see {@link PrefixVocabulary}. */
  readonly vocab: PrefixVocabulary;
  /** Prefix → URI in force (`""` is the default namespace). */
  readonly bindings: ReadonlyMap<string, string>;
  /** Document prefix → canonical prefix, for every binding in force. */
  readonly rename: ReadonlyMap<string, string>;
  /** Whether `rename` maps every prefix to itself. */
  readonly identity: boolean;
  /** Canonical prefix → URI bound in the rewritten output. */
  readonly output: ReadonlyMap<string, string>;
  /**
   * The `xmlns` declarations the element that opened this scope must carry in the rewritten output: its
   * own, renamed, plus one for every inherited binding whose canonical prefix changed here (an unknown
   * namespace pushed to `nsN` because a known one took its prefix). Without those, a descendant would
   * name a prefix nothing declares.
   */
  readonly declarations: Readonly<Record<string, string>>;
  /** Whether `declarations` is exactly the element's own declarations, so its tag can pass through. */
  readonly ownDeclarationsSuffice: boolean;
  /** The document's default namespace in force (`""`: none). */
  readonly inputDefault: string;
  /**
   * Filled as names are met, so a document that renames every tag resolves each distinct name once:
   * an element's QName → its renamed name and namespace, and a prefixed attribute's QName → its
   * renamed name and what its value names.
   */
  readonly elements: Map<string, ElementName>;
  readonly attributes: Map<string, AttributeName>;
  /**
   * A small direct-mapped cache in front of `elements`, by length and last character: a part names a
   * handful of elements over and over (`c`, `v`, `row`), and comparing one string beats hashing it.
   */
  readonly hotNames: string[];
  readonly hotEntries: ElementName[];
  /** The prefixes in force whose attributes can change (`affectsAttributes`), for a hinted `open`. */
  readonly affecting: Set<string>;
}

interface ElementName {
  readonly name: string;
  /** Whether `name` has no prefix, so the default namespace in force decides it. */
  readonly unprefixed: boolean;
  /** The namespace the document puts it in (`""`: none), `undefined` for an undeclared prefix. */
  readonly uri: string | undefined;
  /** Whether it is an `mc:Choice`, whose unprefixed `Requires` lists prefixes. */
  readonly choice: boolean;
}

interface AttributeName {
  readonly name: string;
  readonly kind: ValueKind;
}

function rootScope(vocab: PrefixVocabulary): Scope {
  return {
    vocab,
    bindings: new Map([["xml", "http://www.w3.org/XML/1998/namespace"]]),
    rename: new Map([["xml", "xml"]]),
    identity: true,
    output: new Map([["xml", "http://www.w3.org/XML/1998/namespace"]]),
    declarations: attributeRecord(),
    ownDeclarationsSuffice: true,
    inputDefault: "",
    elements: new Map(),
    attributes: new Map(),
    hotNames: new Array<string>(HOT_NAMES).fill(""),
    hotEntries: new Array<ElementName>(HOT_NAMES),
    affecting: new Set()
  };
}

const HOT_NAMES = 16;

/**
 * Most distinct names a scope remembers. The caches live until the scope closes — for the document
 * element's, the end of the part — so a part naming ever new elements under one scope would otherwise
 * hold one entry per name; past this a name is resolved again each time it is met.
 */
const MAX_CACHED_NAMES = 1024;

/**
 * An attribute map with no prototype, so a name such as `__proto__` or `constructor` is an attribute
 * like any other rather than a property every object inherits.
 */
function attributeRecord(): Record<string, string> {
  return Object.create(null) as Record<string, string>;
}

function elementName(qname: string, scope: Scope): ElementName {
  const slot = (qname.length + qname.charCodeAt(qname.length - 1)) & (HOT_NAMES - 1);
  if (scope.hotNames[slot] === qname) {
    return scope.hotEntries[slot];
  }
  let entry = scope.elements.get(qname);
  if (entry === undefined) {
    const name = renameQName(qname, scope);
    const uri = elementNamespace(qname, scope);
    entry = {
      name,
      unprefixed: name.indexOf(":") === -1,
      uri,
      choice: uri === MCE_NAMESPACE && qname.slice(qname.indexOf(":") + 1) === "Choice"
    };
    if (scope.elements.size < MAX_CACHED_NAMES) {
      scope.elements.set(qname, entry);
    }
  }
  scope.hotNames[slot] = qname;
  scope.hotEntries[slot] = entry;
  return entry;
}

/** A prefixed attribute's renamed name and value kind; `qname` must have a prefix. */
function attributeName(qname: string, scope: Scope): AttributeName {
  let entry = scope.attributes.get(qname);
  if (entry === undefined) {
    entry = { name: renameQName(qname, scope), kind: valueKind("", qname, scope) };
    if (scope.attributes.size < MAX_CACHED_NAMES) {
      scope.attributes.set(qname, entry);
    }
  }
  return entry;
}

/** How many namespace declarations `attributes` holds. */
function declarationCount(attributes: Record<string, string>): number {
  let count = 0;
  for (const key in attributes) {
    if (key.charCodeAt(0) === 120 && (key === "xmlns" || key.startsWith("xmlns:"))) {
      count++;
    }
  }
  return count;
}

/** How many times `xmlns` occurs in `text`. */
function xmlnsCount(text: string): number {
  let count = 0;
  for (let at = text.indexOf("xmlns"); at !== -1; at = text.indexOf("xmlns", at + 5)) {
    count++;
  }
  return count;
}

function childScope(parent: Scope, attributes: Record<string, string>): Scope {
  const vocab = parent.vocab;
  const bindings = new Map(parent.bindings);
  for (const key in attributes) {
    if (key === "xmlns") {
      bindings.set("", attributes[key]);
    } else if (key.startsWith("xmlns:")) {
      bindings.set(key.slice(6), attributes[key]);
    }
  }
  const rename = new Map<string, string>();
  const taken = new Set<string>();
  // Known namespaces first, so they get their canonical prefix whoever else wanted it.
  for (const [prefix, uri] of bindings) {
    const canonical = prefix === "xml" ? "xml" : vocab.prefixOf.get(uri);
    if (canonical !== undefined) {
      rename.set(prefix, canonical);
      taken.add(canonical);
    }
  }
  let counter = 0;
  for (const [prefix, uri] of bindings) {
    if (rename.has(prefix) || uri === "") {
      continue;
    }
    let target = prefix;
    // A prefix the vocabulary gives a known namespace is not left to an unknown one, when it reserves them.
    if (
      taken.has(target) ||
      (vocab.reservePrefixes && target !== "" && vocab.prefixes.has(target))
    ) {
      do {
        target = `ns${++counter}`;
      } while (taken.has(target) || bindings.has(target) || vocab.prefixes.has(target));
    }
    rename.set(prefix, target);
    taken.add(target);
  }
  let identity = true;
  for (const [prefix, target] of rename) {
    if (prefix !== target) {
      identity = false;
      break;
    }
  }
  const output = new Map(parent.output);
  const declarations = attributeRecord();
  let ownDeclarationsSuffice = true;
  for (const key in attributes) {
    if (key === "xmlns" || key.startsWith("xmlns:")) {
      const prefix = key === "xmlns" ? "" : key.slice(6);
      const target = rename.get(prefix) ?? prefix;
      if (target !== prefix) {
        ownDeclarationsSuffice = false;
      }
      declarations[target === "" ? "xmlns" : `xmlns:${target}`] = attributes[key];
      output.set(target, attributes[key]);
    }
  }
  for (const [prefix, target] of rename) {
    const uri = bindings.get(prefix)!;
    if (output.get(target) !== uri) {
      ownDeclarationsSuffice = false;
      declarations[target === "" ? "xmlns" : `xmlns:${target}`] = uri;
      output.set(target, uri);
    }
  }
  const affecting = new Set<string>();
  for (const [prefix, uri] of bindings) {
    if (affectsAttributes(prefix, uri, vocab)) {
      affecting.add(prefix);
    }
  }
  return {
    vocab,
    bindings,
    rename,
    identity,
    output,
    declarations,
    ownDeclarationsSuffice,
    inputDefault: bindings.get("") ?? "",
    elements: new Map(),
    attributes: new Map(),
    hotNames: new Array<string>(HOT_NAMES).fill(""),
    hotEntries: new Array<ElementName>(HOT_NAMES),
    affecting
  };
}

function renameQName(name: string, scope: Scope): string {
  const colon = name.indexOf(":");
  const prefix = colon === -1 ? "" : name.slice(0, colon);
  const target = scope.rename.get(prefix);
  if (target === undefined || target === prefix) {
    return name;
  }
  const local = colon === -1 ? name : name.slice(colon + 1);
  return target === "" ? local : `${target}:${local}`;
}

function renameAttribute(name: string, scope: Scope): string {
  // An unprefixed attribute is in no namespace, whatever the default namespace is.
  return name.indexOf(":") === -1 ? name : renameQName(name, scope);
}

/** The namespace URI `prefix` is bound to in the document (`""` is the default namespace). */
function namespaceOf(qname: string, scope: Scope): string | undefined {
  const colon = qname.indexOf(":");
  return scope.bindings.get(colon === -1 ? "" : qname.slice(0, colon));
}

/**
 * What the value of attribute `key` on element `element` names, both resolved by namespace URI from the
 * document's own bindings — so `m:Ignorable` with `m` bound to MCE is one, and `mc:Ignorable` with `mc`
 * bound to anything else is not.
 */
function valueKind(element: string, key: string, scope: Scope): ValueKind {
  const colon = key.indexOf(":");
  if (colon === -1) {
    // `Requires` is in no namespace, and is the MCE one only on an `mc:Choice`.
    return key === "Requires" &&
      element.slice(element.indexOf(":") + 1) === "Choice" &&
      namespaceOf(element, scope) === MCE_NAMESPACE
      ? ValueKind.Prefixes
      : ValueKind.None;
  }
  const uri = scope.bindings.get(key.slice(0, colon));
  const local = key.slice(colon + 1);
  if (uri === MCE_NAMESPACE) {
    return MCE_ATTRIBUTES.get(local) ?? ValueKind.None;
  }
  return uri === XSI_NAMESPACE && local === "type" ? ValueKind.QName : ValueKind.None;
}

/**
 * A prefix bound to `uri` on the element being written, declared in `out` the first time it is asked
 * for (`aliases` records them, URI → prefix). `preferred` when it is one and the output does not bind it
 * here, else `nsN`: nothing in the output names a prefix `output` does not hold, so a free one captures
 * no name.
 */
function declareAlias(
  uri: string,
  preferred: string,
  scope: Scope,
  out: Record<string, string>,
  aliases: Map<string, string>
): string {
  let alias = aliases.get(uri);
  if (alias === undefined) {
    alias = preferred;
    for (
      let n = 1;
      alias === "" || scope.output.has(alias) || out[`xmlns:${alias}`] !== undefined;
      n++
    ) {
      alias = `ns${n}`;
    }
    aliases.set(uri, alias);
    out[`xmlns:${alias}`] = uri;
  }
  return alias;
}

/**
 * `prefix`, named in an attribute value, as the output names it; one the document never declared is kept
 * as written. A prefix whose namespace became the default one is the one case the rename cannot express —
 * a list entry, or the prefix of a `p:*`, has to be a prefix — so the element is given a declaration that
 * binds one to that namespace, and the value names that: it keeps naming exactly the namespaces it named.
 */
function valuePrefix(
  prefix: string,
  scope: Scope,
  out: Record<string, string>,
  aliases: Map<string, string>
): string {
  const target = scope.rename.get(prefix);
  return target !== ""
    ? (target ?? prefix)
    : declareAlias(scope.bindings.get(prefix)!, prefix, scope, out, aliases);
}

function rewriteValue(
  kind: ValueKind,
  value: string,
  scope: Scope,
  out: Record<string, string>,
  aliases: Map<string, string>
): string {
  if (kind === ValueKind.QName) {
    const qname = value.trim();
    // Unprefixed, it is in the default namespace, which `renameQName` follows as it does an element's.
    return qname === "" ? value : renameQName(qname, scope);
  }
  return mapEntries(value, entry => {
    if (kind === ValueKind.Prefixes) {
      return valuePrefix(entry, scope, out, aliases);
    }
    const colon = entry.indexOf(":");
    // MCE requires a prefix on every QName entry; one without is malformed and kept as written.
    return colon <= 0
      ? entry
      : `${valuePrefix(entry.slice(0, colon), scope, out, aliases)}${entry.slice(colon)}`;
  });
}

function mapEntries(value: string, map: (entry: string) => string): string {
  let out = "";
  for (const entry of value.split(/\s+/)) {
    if (entry !== "") {
      out = out === "" ? map(entry) : `${out} ${map(entry)}`;
    }
  }
  return out;
}

function rewriteAttributes(
  element: string,
  attributes: Record<string, string>,
  scope: Scope,
  opensScope: boolean,
  aliases: Map<string, string>
): Record<string, string> {
  // Declarations are emitted only by the element that opened the scope; others carry none. Each of the
  // element's own takes the place it had, renamed, so the rewritten tag lists its attributes in the
  // document's order; the declarations it gains (inherited bindings renamed here) follow them.
  const out = attributeRecord();
  for (const key in attributes) {
    if (key.charCodeAt(0) === 120 && (key === "xmlns" || key.startsWith("xmlns:"))) {
      const prefix = key === "xmlns" ? "" : key.slice(6);
      const target = scope.rename.get(prefix) ?? prefix;
      const written = target === "" ? "xmlns" : `xmlns:${target}`;
      if (scope.declarations[written] !== undefined) {
        out[written] = scope.declarations[written];
      }
      continue;
    }
    let name = key;
    let kind = ValueKind.None;
    if (key.indexOf(":") !== -1) {
      const entry = attributeName(key, scope);
      name = writtenAttribute(key, entry.name, scope, out, aliases);
      kind = entry.kind;
    } else if (key === "Requires") {
      // An unprefixed attribute is in no namespace, whatever the default namespace is — and names
      // prefixes only as an `mc:Choice`'s `Requires`.
      kind = valueKind(element, key, scope);
    }
    out[name] =
      kind === ValueKind.None
        ? attributes[key]
        : rewriteValue(kind, attributes[key], scope, out, aliases);
  }
  if (opensScope) {
    for (const key in scope.declarations) {
      if (out[key] === undefined) {
        out[key] = scope.declarations[key];
      }
    }
  }
  return out;
}

/**
 * The name prefixed attribute `key`, renamed to `renamed`, is written with. An attribute name is in a
 * namespace only through a prefix, so one whose namespace became the default one cannot lose it — it
 * would leave its namespace, and take the place of the unprefixed attribute of the same local name. It
 * keeps one bound to that namespace on the element instead, the document's own where it is free.
 */
function writtenAttribute(
  key: string,
  renamed: string,
  scope: Scope,
  out: Record<string, string>,
  aliases: Map<string, string>
): string {
  if (renamed.indexOf(":") !== -1) {
    return renamed;
  }
  const uri = namespaceOf(key, scope);
  return uri === undefined
    ? key
    : `${declareAlias(uri, prefixOf(key), scope, out, aliases)}:${renamed}`;
}

/** Whether no attribute name in `attributes` has a prefix. */
function unprefixedAttributes(attributes: Record<string, string>): boolean {
  for (const key in attributes) {
    if (key.indexOf(":") !== -1) {
      return false;
    }
  }
  return true;
}

/** How many attribute names in `attributes` have a prefix in `prefixes`. */
function prefixedAttributeCount(
  attributes: Record<string, string>,
  prefixes: ReadonlySet<string>
): number {
  let count = 0;
  for (const key in attributes) {
    const colon = key.indexOf(":");
    if (colon > 0 && prefixes.has(key.slice(0, colon))) {
      count++;
    }
  }
  return count;
}

/** Whether the declaration `prefix = uri` can change an attribute: by being renamed, or by naming prefixes. */
function affectsAttributes(prefix: string, uri: string, vocab: PrefixVocabulary): boolean {
  return (
    prefix !== "" &&
    (!isCanonicalDeclaration(prefix, uri, vocab) || uri === MCE_NAMESPACE || uri === XSI_NAMESPACE)
  );
}

/** How many matches of the global `pattern` in `text` end after `from`. */
function matchCount(pattern: RegExp | undefined, text: string, from: number): number {
  if (pattern === undefined) {
    return 0;
  }
  let count = 0;
  pattern.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(text)) !== null) {
    if (match.index + match[0].length > from) {
      count++;
    }
  }
  return count;
}

/** Whether a prefixed attribute of `attributes` is an `xsi:type`, whose value the text view may need. */
function hasQNameValue(attributes: Record<string, string>, scope: Scope): boolean {
  for (const key in attributes) {
    if (
      key.indexOf(":") !== -1 &&
      !key.startsWith("xmlns:") &&
      attributeName(key, scope).kind === ValueKind.QName
    ) {
      return true;
    }
  }
  return false;
}

/** The namespace `qname` is in (`""`: none), as an element name or a QName value resolves it. */
function elementNamespace(qname: string, scope: Scope): string | undefined {
  const colon = qname.indexOf(":");
  return colon === -1 ? (scope.bindings.get("") ?? "") : scope.bindings.get(qname.slice(0, colon));
}

/**
 * Make the rewritten tag mean in **text** what the input did, for the one thing the rename cannot see:
 * what an unprefixed name resolves to depends on the default namespace in force in the output, which is
 * not the document's once a prefixed namespace has become the default one. An element in no namespace
 * (`<foo/>` under `<s:worksheet xmlns:s="…main">`) would be captured by it, and an unprefixed `xsi:type`
 * would follow it too. So the default namespace actually written is tracked (`inherited` is the one in
 * force at the parent), and corrected where a name needs another:
 *
 * - an unprefixed element whose namespace is not the default in force declares its own — `xmlns=""` for
 *   none — which its descendants then inherit, and are corrected against in turn;
 * - an unprefixed `xsi:type` in another namespace names it through a declared prefix, and one in **no**
 *   namespace needs `xmlns=""`, so an element in a namespace carrying one is written with a prefix.
 *
 * Returns the default namespace in force inside the element, for its children.
 */
function settleDefaultNamespace(
  input: ParseOpenTag,
  tag: { name: string; attributes: Record<string, string> },
  scope: Scope,
  inherited: string,
  aliases: Map<string, string>
): string {
  const out = tag.attributes;
  // The output's unprefixed `xsi:type` values: attribute, the namespace it must resolve to, and the
  // prefix the document gave it (preferred for a declaration).
  let types: [key: string, uri: string, prefix: string][] | undefined;
  let noNamespaceType = false;
  for (const key in input.attributes) {
    if (key.indexOf(":") !== -1 && valueKind(input.name, key, scope) === ValueKind.QName) {
      const written = writtenAttribute(key, renameAttribute(key, scope), scope, out, aliases);
      const original = input.attributes[key].trim();
      const uri = elementNamespace(original, scope);
      if (out[written] !== "" && out[written].indexOf(":") === -1 && uri !== undefined) {
        (types ??= []).push([written, uri, prefixOf(original)]);
        noNamespaceType ||= uri === "";
      }
    }
  }
  let inForce = out.xmlns ?? inherited;
  if (tag.name.indexOf(":") === -1) {
    const uri = elementNamespace(input.name, scope)!;
    if (uri !== "" && noNamespaceType) {
      const alias = declareAlias(uri, prefixOf(input.name), scope, out, aliases);
      tag.name = `${alias}:${tag.name}`;
    } else if (uri !== inForce) {
      out.xmlns = inForce = uri;
    }
  }
  if (noNamespaceType && inForce !== "") {
    out.xmlns = inForce = "";
  }
  for (const [key, uri, prefix] of types ?? []) {
    if (uri !== inForce) {
      out[key] = `${declareAlias(uri, prefix, scope, out, aliases)}:${out[key]}`;
    }
  }
  return inForce;
}

/** The prefix of `qname`, `""` for none. */
function prefixOf(qname: string): string {
  const colon = qname.indexOf(":");
  return colon === -1 ? "" : qname.slice(0, colon);
}

/**
 * A stateful prefix normaliser for one document: `open` each start tag and `close` each end tag, in
 * document order (a self-closing tag is opened and closed), and hand the results to the xform.
 *
 * What it hands the xform is the **name** view: names an xform can match, which is not always text that
 * means the same — an element in no namespace keeps its unprefixed name even where the output's default
 * namespace is one a prefixed namespace was renamed to. Whoever writes tags back out as text (a capture
 * of a subtree the reader keeps verbatim) asks for the **text** view of the tag just opened or closed,
 * and for the bindings in force there — see `fragment-namespaces.ts`.
 */
export interface PrefixNormalizer {
  /**
   * @param prefixedAttributes - Whether any of the tag's attribute names has a prefix, from a caller that knows
   *   without looking (the parser counts them: `SaxParser.countPrefixedAttributes`). Given, a tag reported to have
   *   none keeps its attributes without their being enumerated; see `PrefixNormalizerOptions.hinted`.
   */
  open<T extends ParseOpenTag>(tag: T, prefixedAttributes?: boolean): T | ParseOpenTag;
  close(name: string): string;
  /**
   * Show the normaliser the raw text **before** it is handed to the parser. Only a normaliser created
   * with `{ lazy: true }` uses it: such a normaliser does no per-tag work at all until the text it has
   * been shown declares a namespace with a non-canonical prefix, or declares one anywhere but on the
   * document element (whose scope a capture then needs tracked). Returns whether the normaliser is
   * active: until it is, `open` and `close` return their input, so a hot loop may skip calling them.
   */
  scan(text: string): boolean;
  /**
   * The tag `open` returned last, as **text** must write it: with an `xmlns` where the output's default
   * namespace would capture an unprefixed name, and a prefix where a name needs one. `undefined` when
   * that is the tag `open` returned.
   */
  textOpened(): ParseOpenTag | undefined;
  /** The name `close` returned last, as text must write it; `undefined` when that is the same name. */
  textClosed(): string | undefined;
  /**
   * The namespace bindings in force inside the element `open` returned last, as its text view binds
   * them (`""`: the default namespace, `""` for none). A capture starting at that element declares from
   * these the prefixes its subtree uses without declaring.
   */
  bindings(): ReadonlyMap<string, string>;
  /** The declarations on the document element, as its text view writes them; empty until it is seen. */
  documentBindings(): ReadonlyMap<string, string>;
}

export interface PrefixNormalizerOptions {
  /**
   * Stay dormant — `open` and `close` return their input — until `scan` sees a declaration that the
   * normaliser would not map to itself, or any declaration below the document element. Excel and this
   * library declare canonical prefixes, and a large part (a worksheet) declares them on its root, so
   * reading one then costs an `indexOf("xmlns")` per chunk rather than work per tag; the small parts
   * that declare a namespace further down (a drawing's `a:ext`) are tracked from there, so a capture
   * knows which declarations are in force.
   *
   * Enabling part-way through a document is exact: until then every declaration in force was canonical
   * and on the document element, so each one maps to itself and none can collide with a canonical
   * prefix — the scopes computed from the declarations seen afterwards are the ones a normaliser
   * running from the start would compute, and the document element's are known from the scan.
   */
  readonly lazy?: boolean;
  /**
   * `open` and `close` return the **text** view rather than the name view (`canonicalizeXml`, which
   * writes every tag out). Text resolves an unprefixed name against the default namespace the output
   * declares, so this view keeps that one right — see `settleDefaultNamespace` — at the cost of an
   * `xmlns` attribute on an element that needs one. An xform matches names and must not see that
   * attribute, so the reading paths leave this off and ask for the text view only where they write one
   * (`textOpened`). Implies an eager normaliser.
   */
  readonly forText?: boolean;
  /**
   * Every `open` will say whether the tag has a prefixed attribute name, so a lazy normaliser need not search the
   * text for the ones that can change. That search is a regular expression over the whole part — often a single
   * chunk of tens of megabytes — and was the largest cost of reading a document with non-canonical prefixes; the
   * parser learns the same thing while reading the names, for nothing.
   */
  readonly hinted?: boolean;
}

/** URI → whether the declaration `prefix = uri` is one the normaliser maps to itself, unconditionally. */
function isCanonicalDeclaration(prefix: string, rawUri: string, vocab: PrefixVocabulary): boolean {
  if (rawUri.includes("&")) {
    // An escaped URI may name a known namespace; do not try to decode it here.
    return false;
  }
  if (prefix === "xml") {
    return rawUri === "http://www.w3.org/XML/1998/namespace";
  }
  const canonical = vocab.prefixOf.get(rawUri);
  if (canonical !== undefined) {
    return canonical === prefix;
  }
  // An unknown namespace keeps its prefix — unless that prefix is one a known namespace may claim, the
  // default namespace, or the shape of a prefix the renamer invents.
  return prefix !== "" && !vocab.prefixes.has(prefix) && !/^ns\d+$/.test(prefix);
}

/**
 * The namespace `prefix` stands for in every part this library reads and writes — the reader renames a
 * known namespace to it — or `undefined` for a prefix that is no namespace's canonical one.
 */
export function canonicalNamespaceOf(vocab: PrefixVocabulary, prefix: string): string | undefined {
  return vocab.namespaceOf.get(prefix);
}

/** A namespace declaration, matched only at `lastIndex` — see `eachDeclaration`. */
const DECLARATION_AT = /xmlns(?::([^\s=/>]+))?\s*=\s*(?:"([^"]*)"|'([^']*)')/y;

/**
 * Hand `visit` every declaration in `text` from `from` on, as a global search for one would find them, and
 * return where the last one ends (`from` for none) — or `-1` as soon as `visit` returns false.
 *
 * Matched only where `xmlns` occurs. A whole part is often a single chunk, and a global search walks all forty
 * megabytes of a worksheet's cells for the four declarations on its root; `indexOf` skips them natively.
 */
function eachDeclaration(
  text: string,
  from: number,
  visit: (prefix: string, rawUri: string) => boolean
): number {
  let end = from;
  for (let at = text.indexOf("xmlns", from); at !== -1;) {
    DECLARATION_AT.lastIndex = at;
    const match = DECLARATION_AT.exec(text);
    if (match === null) {
      at = text.indexOf("xmlns", at + 5);
      continue;
    }
    if (!visit(match[1] ?? "", match[2] ?? match[3])) {
      return -1;
    }
    end = DECLARATION_AT.lastIndex;
    at = text.indexOf("xmlns", end);
  }
  return end;
}

/** Longest unfinished declaration carried into the next chunk before giving up and enabling. */
const MAX_CARRY = 65536;

/**
 * Whether every declaration in `text` is one `accept` takes. Returns the unscanned tail (a declaration,
 * or the start of one, cut by the end of the text) as `carry`, or `undefined` once one is not accepted.
 */
function scanDeclarations(
  text: string,
  accept: (prefix: string, rawUri: string) => boolean
): string | undefined {
  let from = text.indexOf("xmlns");
  if (from === -1) {
    // Only an `xmlns` cut short can start in the last four characters.
    return text.slice(Math.max(0, text.length - 4));
  }
  const end = eachDeclaration(text, from, accept);
  if (end === -1) {
    return undefined;
  }
  from = text.indexOf("xmlns", end);
  if (from === -1) {
    return text.slice(Math.max(end, text.length - 4));
  }
  // An `xmlns` the pattern did not match: either cut short by the end of the text, or one in text
  // content that never will. Keep it for the next chunk, bounded.
  const carry = text.slice(from);
  return carry.length > MAX_CARRY ? undefined : carry;
}

const acceptNone = (): boolean => false;

/**
 * Whether `prolog` holds a DOCTYPE with an internal subset, which can declare namespaces no scan of
 * the tags sees: an `<!ATTLIST>` default gives an element an `xmlns` its tag does not write, and an
 * entity expands to markup that may carry one. The parser applies both, so the tags the normaliser is
 * handed declare what the text never spelt as a declaration — every one of them has to be looked at.
 */
function prologDeclaresAttributes(prolog: string): boolean {
  const doctype = prolog.indexOf("<!DOCTYPE");
  return doctype !== -1 && prolog.indexOf("[", doctype) !== -1;
}

/**
 * Where the document element's start tag is in `text` — its `<` and the index after its `>` — or
 * `undefined` when `text` does not reach its end. Skips the prolog: the XML declaration, processing
 * instructions, comments and a DOCTYPE.
 */
function documentElementTag(text: string): [start: number, end: number] | undefined {
  let from = 0;
  for (;;) {
    const lt = text.indexOf("<", from);
    if (lt === -1 || lt + 1 >= text.length) {
      return undefined;
    }
    const next = text.charCodeAt(lt + 1);
    if (next === 63 /* ? */ || next === 33 /* ! */) {
      const close = text.startsWith("<!--", lt)
        ? text.indexOf("-->", lt + 4) + 3
        : next === 63
          ? text.indexOf("?>", lt + 2) + 2
          : text.indexOf(">", lt + 2) + 1;
      if (close <= 0) {
        return undefined;
      }
      from = close;
      continue;
    }
    let quote = 0;
    for (let i = lt + 1; i < text.length; i++) {
      const c = text.charCodeAt(i);
      if (quote !== 0) {
        if (c === quote) {
          quote = 0;
        }
      } else if (c === 34 || c === 39) {
        quote = c;
      } else if (c === 62 /* > */) {
        return [lt, i + 1];
      }
    }
    return undefined;
  }
}

export function createPrefixNormalizer(
  vocab: PrefixVocabulary,
  options?: PrefixNormalizerOptions
): PrefixNormalizer {
  // Scopes are pushed only by elements that declare a namespace, with the depth they were opened
  // at; an element's closing name is renamed under the scope it opened in, so no name is stored.
  const scopes: Scope[] = [];
  const depths: number[] = [];
  let scope = rootScope(vocab);
  let depth = 0;
  const forText = options?.forText === true;
  let active = forText || options?.lazy !== true;
  let carry = "";
  // Dormant: the text before the document element's start tag is complete, its declarations once it
  // is (all canonical, or the normaliser is active), and whether the parser can have opened it yet.
  let head = "";
  let headDone = false;
  let documentDeclarations: Map<string, string> | undefined;
  let documentPending = false;
  /** Whether the document element was opened while dormant, so no scope holds its declarations. */
  let documentDormant = false;
  // The text view: the default namespace in force inside each open element (`""`: none) and the name
  // each was written with, which may carry a prefix its rename does not give it. An element opened
  // while dormant has neither; its close finds `names` empty.
  const defaults: string[] = [""];
  // The name each open element was handed to the xform with — and, for the few whose text view is
  // written with another, that name and the element's index in `names`.
  const names: string[] = [];
  const textNames: string[] = [];
  const textAt: number[] = [];
  // A lazy normaliser keeps scanning once active, counting the `xmlns` the text holds that no opened
  // tag has accounted for yet: while that is none, no tag can declare anything, and `open` need not
  // look. A count that cannot be trusted (it went below zero) turns the shortcut off.
  let counting = false;
  let pending = 0;
  let scannedTail = "";
  // Likewise for attributes: the prefixes whose attributes can change (`affectsAttributes`) and how many
  // of their attribute names (whitespace, prefix, colon) the text holds that no opened tag accounted
  // for. While that is none, a tag's attributes are kept without being looked at. A declaration seen
  // after the count started may add a prefix, so it turns this off; so does a tag that took the count
  // below zero.
  let attributeCounting = false;
  let attributePending = 0;
  let attributeTail = "";
  let attributeTailLength = 0;
  const affecting = new Set<string>();
  let affectingPattern: RegExp | undefined;
  // Dormant: the text since the last `<` — a start tag the parser may not have opened yet — or
  // `undefined` once it outgrows `MAX_CARRY`.
  let openTagText: string | undefined = "";
  let lastOpened: ParseOpenTag | undefined;
  let lastClosed: string | undefined;
  // The prefixes declared for attribute values on the element being rewritten; reused, as most need none.
  const aliases = new Map<string, string>();

  /**
   * Become active; `scanned` is the text no tag has been opened with yet — everything since the last
   * tag the parser opened, at the least.
   */
  const activate = (scanned: string | undefined): void => {
    active = true;
    carry = "";
    head = "";
    if (options?.hinted === true) {
      // The document element's declarations, when it was opened dormant, are in no scope; the root's stands in.
      for (const [prefix, uri] of documentDeclarations ?? []) {
        if (affectsAttributes(prefix, uri, vocab)) {
          scope.affecting.add(prefix);
        }
      }
    }
    if (options?.lazy === true && scanned !== undefined) {
      counting = true;
      pending = xmlnsCount(scanned);
      scannedTail = scanned.slice(-4);
      let declarations = 0;
      for (const [prefix, uri] of documentDeclarations ?? []) {
        if (affectsAttributes(prefix, uri, vocab)) {
          affecting.add(prefix);
        }
      }
      eachDeclaration(scanned, 0, (prefix, uri) => {
        declarations++;
        if (affectsAttributes(prefix, uri, vocab)) {
          affecting.add(prefix);
        }
        return true;
      });
      // A declaration cut by the end of the text cannot be read, so its prefix is not known. A hinted normaliser
      // is told per tag instead.
      attributeCounting = declarations === pending && options?.hinted !== true;
      if (attributeCounting && affecting.size > 0) {
        const alternatives = [...affecting].map(escapeRegExp);
        affectingPattern = new RegExp(`[ \t\r\n](?:${alternatives.join("|")}):`, "g");
        attributeTailLength = Math.max(...[...affecting].map(prefix => prefix.length)) + 1;
      }
      attributePending = matchCount(affectingPattern, scanned, 0);
      attributeTail = scanned.slice(-attributeTailLength);
    }
    if (documentPending) {
      // The parser has yet to open the document element: it will be opened active, and its own view
      // supplies the declarations.
      documentDeclarations = undefined;
    } else if (documentDeclarations !== undefined) {
      // Opened while dormant, with canonical declarations only — so the text view's default namespace
      // in force is still the one it declared.
      documentDormant = true;
      defaults[0] = documentDeclarations.get("") ?? "";
    }
  };

  const declarationsOf = (attributes: Record<string, string>): Map<string, string> => {
    const out = new Map<string, string>();
    for (const key in attributes) {
      if (key === "xmlns") {
        out.set("", attributes[key]);
      } else if (key.startsWith("xmlns:")) {
        out.set(key.slice(6), attributes[key]);
      }
    }
    return out;
  };

  return {
    scan(text) {
      if (active) {
        if (counting) {
          // An `xmlns` cut by the previous chunk starts in its last four characters.
          const joined = scannedTail + text;
          const declarations = xmlnsCount(joined);
          pending += declarations;
          scannedTail = joined.slice(-4);
          if (declarations > 0) {
            attributeCounting = false;
          }
        }
        if (attributeCounting && affectingPattern !== undefined) {
          const joined = attributeTail + text;
          attributePending += matchCount(affectingPattern, joined, attributeTail.length);
          attributeTail = joined.slice(-attributeTailLength);
        }
        return true;
      }
      documentPending = false;
      // When the document element completes in this text, it too is opened only after it is scanned.
      let whole: string | undefined;
      if (!headDone) {
        head += text;
        const tag = documentElementTag(head);
        if (tag === undefined) {
          if (head.length > MAX_CARRY) {
            activate(head);
          }
          return active;
        }
        headDone = true;
        if (prologDeclaresAttributes(head.slice(0, tag[0]))) {
          activate(head);
          return true;
        }
        documentPending = true;
        documentDeclarations = new Map();
        const declarations = documentDeclarations;
        const tail = scanDeclarations(head.slice(tag[0], tag[1]), (prefix, uri) => {
          declarations.set(prefix, uri);
          return isCanonicalDeclaration(prefix, uri, vocab);
        });
        if (tail === undefined) {
          activate(head);
          return true;
        }
        whole = head;
        text = head.slice(tag[1]);
        head = "";
      }
      // Past the document element, any declaration opens a scope a capture may need to know.
      const scanned = carry === "" ? text : carry + text;
      const tail = scanDeclarations(scanned, acceptNone);
      if (tail === undefined) {
        // `carry` and `openTagText` both end where the previous text did: the longer holds the other.
        const before =
          whole !== undefined || openTagText === undefined || openTagText.length <= carry.length
            ? carry
            : openTagText;
        activate(whole ?? (openTagText === undefined ? undefined : before + text));
      } else {
        carry = tail;
        const lt = text.lastIndexOf("<");
        openTagText =
          lt !== -1
            ? text.slice(lt)
            : openTagText === undefined || openTagText.length + text.length > MAX_CARRY
              ? undefined
              : openTagText + text;
      }
      return active;
    },
    open(tag, prefixedAttributes) {
      lastOpened = undefined;
      if (!active) {
        return tag;
      }
      depth++;
      let opensScope = false;
      if (!counting || pending !== 0) {
        const count = declarationCount(tag.attributes);
        opensScope = count > 0;
        pending -= count;
        if (pending < 0) {
          counting = false;
        }
      }
      if (opensScope) {
        scopes.push(scope);
        depths.push(depth);
        scope = childScope(scope, tag.attributes);
      }
      const inherited = defaults[names.length];
      // Only a declaring tag can carry `xmlns`; looking it up costs a hash lookup per tag otherwise.
      const own = opensScope ? (tag.attributes.xmlns ?? inherited) : inherited;
      if (
        scope.identity &&
        (!opensScope || scope.ownDeclarationsSuffice) &&
        // Written as it is, the tag declares what the document does: right, unless an ancestor's
        // correction left another default namespace in force than the document's.
        own === scope.inputDefault
      ) {
        if (documentDeclarations === undefined) {
          documentDeclarations = opensScope ? declarationsOf(tag.attributes) : new Map();
        }
        defaults[names.length + 1] = own;
        names.push(tag.name);
        return tag;
      }
      const element = elementName(tag.name, scope);
      // Outside a scope's opening element, the attributes are kept as they are unless one of them can
      // change — the common case of a renamed document, where only the element's name does.
      let plain: boolean;
      if (prefixedAttributes !== undefined) {
        plain =
          !opensScope &&
          (!prefixedAttributes || prefixedAttributeCount(tag.attributes, scope.affecting) === 0);
      } else if (attributeCounting && attributePending === 0 && !opensScope) {
        plain = true;
      } else if (attributeCounting) {
        const count = prefixedAttributeCount(tag.attributes, affecting);
        attributePending -= count;
        if (attributePending < 0) {
          attributeCounting = false;
        }
        plain = count === 0 && !opensScope;
      } else {
        plain = !opensScope && unprefixedAttributes(tag.attributes);
      }
      if (element.choice && tag.attributes.Requires !== undefined) {
        plain = false;
      }
      let attributes = tag.attributes;
      let typed = false;
      if (!plain) {
        aliases.clear();
        attributes = rewriteAttributes(tag.name, tag.attributes, scope, opensScope, aliases);
        typed = hasQNameValue(tag.attributes, scope);
      }
      const renamed = { name: element.name, attributes, isSelfClosing: tag.isSelfClosing };
      const inForce = opensScope ? (attributes.xmlns ?? inherited) : inherited;
      if (documentDeclarations === undefined) {
        documentDeclarations = declarationsOf(attributes);
      }
      if (!typed && !(element.unprefixed && element.uri !== inForce)) {
        // The text view is the name view.
        defaults[names.length + 1] = inForce;
        names.push(element.name);
        return renamed;
      }
      if (attributes === tag.attributes) {
        aliases.clear();
      }
      const text = {
        name: renamed.name,
        attributes: Object.assign(attributeRecord(), attributes),
        isSelfClosing: tag.isSelfClosing
      };
      defaults[names.length + 1] = settleDefaultNamespace(tag, text, scope, inherited, aliases);
      names.push(renamed.name);
      if (text.name !== renamed.name) {
        textNames.push(text.name);
        textAt.push(names.length - 1);
      }
      if (forText) {
        return text;
      }
      if (text.name !== renamed.name || !sameAttributes(text.attributes, renamed.attributes)) {
        lastOpened = text;
      }
      return renamed;
    },
    close(name) {
      lastClosed = undefined;
      if (!active) {
        return name;
      }
      let renamed: string;
      let written: string;
      if (names.length > 0) {
        renamed = written = names.pop()!;
        if (textAt.length > 0 && textAt[textAt.length - 1] === names.length) {
          textAt.pop();
          written = textNames.pop()!;
        }
      } else {
        // Opened while the normaliser was dormant: its name is canonical, and so is its view.
        renamed = written = scope.identity ? name : elementName(name, scope).name;
      }
      if (depths.length > 0 && depths[depths.length - 1] === depth) {
        depths.pop();
        scope = scopes.pop()!;
      }
      depth--;
      if (forText) {
        return written;
      }
      if (written !== renamed) {
        lastClosed = written;
      }
      return renamed;
    },
    textOpened() {
      return lastOpened;
    },
    textClosed() {
      return lastClosed;
    },
    bindings() {
      const out = new Map(active ? scope.output : undefined);
      if (documentDeclarations !== undefined && (!active || documentDormant)) {
        // Declared on a document element opened while dormant: canonical, so as written.
        for (const [prefix, uri] of documentDeclarations) {
          if (!out.has(prefix)) {
            out.set(prefix, uri);
          }
        }
      }
      out.delete("xml");
      out.set("", active ? defaults[names.length] : (documentDeclarations?.get("") ?? ""));
      return out;
    },
    documentBindings() {
      return documentDeclarations ?? new Map();
    }
  };
}

function sameAttributes(a: Record<string, string>, b: Record<string, string>): boolean {
  let count = 0;
  for (const key in a) {
    if (a[key] !== b[key]) {
      return false;
    }
    count++;
  }
  for (const _key in b) {
    count--;
  }
  return count === 0;
}

/**
 * Whether `xml` declares a namespace prefix the normaliser would rewrite. When it does not, normalising
 * it is the identity, so a caller holding the whole text can skip the per-tag work outright.
 */
export function needsPrefixNormalization(vocab: PrefixVocabulary, xml: string): boolean {
  const tag = documentElementTag(xml);
  return (
    prologDeclaresAttributes(tag === undefined ? xml : xml.slice(0, tag[0])) ||
    scanDeclarations(xml, (prefix, uri) => isCanonicalDeclaration(prefix, uri, vocab)) === undefined
  );
}

/**
 * A parsed element tree, normalised in place — for the readers that parse a part into a DOM. The tree
 * gets the name view; the returned map holds the text view of every element whose differs, for a reader
 * that writes a subtree back out (see `fragment-namespaces.ts`).
 */
export function normalizeElementPrefixes(
  vocab: PrefixVocabulary,
  root: NormalizableElement
): ReadonlyMap<NormalizableElement, ParseOpenTag> {
  const normalizer = createPrefixNormalizer(vocab);
  const texts = new Map<NormalizableElement, ParseOpenTag>();
  const open = (element: NormalizableElement): void => {
    const opened = normalizer.open(element);
    const text = normalizer.textOpened();
    if (text !== undefined) {
      texts.set(element, text);
    }
    if (opened !== element) {
      element.name = opened.name;
      element.attributes = opened.attributes;
    }
  };
  // An explicit stack, not recursion: the depth is the reader's policy, which may be set far past
  // what the call stack holds.
  open(root);
  const stack: Array<{ element: NormalizableElement; next: number }> = [{ element: root, next: 0 }];
  while (stack.length > 0) {
    const frame = stack[stack.length - 1];
    const children = frame.element.children ?? [];
    if (frame.next < children.length) {
      const child = children[frame.next++];
      if (child.type === "element") {
        open(child as NormalizableElement);
        stack.push({ element: child as NormalizableElement, next: 0 });
      }
      continue;
    }
    normalizer.close(frame.element.name);
    stack.pop();
  }
  return texts;
}

interface NormalizableElement {
  readonly type?: string;
  name: string;
  attributes: Record<string, string>;
  readonly children?: readonly { readonly type: string }[];
}

/**
 * `xml` rewritten with the canonical prefixes — or `xml` itself, unchanged, when it already uses them.
 *
 * For a part whose **bytes** are kept and later edited in place: a chart's `rawData` is patched with
 * expressions that name Excel's prefixes (`<c:title>…</c:title>`), so it has to be stored in that
 * form, or an edit would miss and leave the old content standing. Parts kept only to be written back
 * verbatim do not need this.
 *
 * A part already in canonical form is returned as the same string, byte for byte. Any other is parsed
 * under `options` — the parse policy of the reader that will consume the result: its depth and entity
 * limits, and what it does with a character XML forbids — and one that does not parse under it throws
 * the parser's `XmlParseError`, as the reader would have. A part that is rewritten is **re-serialised** from parse events, so it is equivalent rather
 * than identical, in these exact ways:
 * - Characters XML 1.0 forbids (C0 controls other than tab/LF/CR, lone surrogates, U+FFFE/U+FFFF) are
 *   handled as `options.invalidCharHandling` says — refused by default, as the parser does — and any
 *   that survive are dropped by the encoder, as it does on every write.
 * - Character data is decoded and re-encoded: entity and character references become literal
 *   characters, and `&`, `<`, `>`, `"` and `'` are written as entities. The text itself is unchanged.
 * - Attribute values carry the parser's attribute-value normalisation — a literal tab or line break
 *   becomes a space, as XML requires every consumer to read it, while one written as a character
 *   reference survives as `&#xA;`/`&#x9;`/`&#xD;` — and are re-quoted with `"`.
 * - Whitespace between attributes and inside tags is not kept, and a DOCTYPE is not emitted
 *   (OOXML parts have none). Comments, CDATA sections and processing instructions are kept.
 * - The MCE lists are re-joined with single spaces. One that names a namespace that became the default
 *   namespace names it through a prefix declared for it on the same element, since a list entry has to
 *   be a prefix.
 * - Where a prefixed namespace became the default one, an element in no namespace under it declares
 *   `xmlns=""` (and its descendants back into it re-declare it), and an unprefixed `xsi:type` that the
 *   default namespace in force would misread names its namespace through a declared prefix.
 */
export function canonicalizeXml(
  vocab: PrefixVocabulary,
  xml: string,
  options?: CanonicalizeXmlOptions
): string {
  if (!needsPrefixNormalization(vocab, xml)) {
    return xml;
  }
  const parser = new SaxParser({
    position: false,
    maxDepth: options?.maxDepth,
    maxEntityExpansions: options?.maxEntityExpansions,
    maxEntityExpansionLength: options?.maxEntityExpansionLength,
    invalidCharHandling: options?.invalidCharHandling,
    internalEntities: options?.internalEntities
  });
  const prefixes = createPrefixNormalizer(vocab, { forText: true });
  const out: string[] = [];
  let changed = false;
  parser.on("pi", (target, body) => out.push(`<?${target}${body ? ` ${body}` : ""}?>`));
  parser.on("comment", text => out.push(`<!--${text}-->`));
  parser.on("cdata", text => out.push(`<![CDATA[${text}]]>`));
  parser.on("text", text => out.push(xmlEncode(text)));
  parser.on("opentag", raw => {
    const tag = prefixes.open(raw);
    changed ||= tag !== raw;
    let s = `<${tag.name}`;
    for (const key in tag.attributes) {
      s += ` ${key}="${xmlEncodeAttr(tag.attributes[key])}"`;
    }
    out.push(raw.isSelfClosing ? `${s}/>` : `${s}>`);
  });
  parser.on("closetag", raw => {
    const name = prefixes.close(raw.name);
    if (!raw.isSelfClosing) {
      out.push(`</${name}>`);
    }
  });
  parser.write(xml);
  parser.close();
  return changed ? out.join("") : xml;
}

/** The parse policy {@link canonicalizeXml} reads a part under: the consuming reader's. */
export type CanonicalizeXmlOptions = Pick<
  XmlParseOptions,
  | "maxDepth"
  | "maxEntityExpansions"
  | "maxEntityExpansionLength"
  | "invalidCharHandling"
  | "internalEntities"
>;
