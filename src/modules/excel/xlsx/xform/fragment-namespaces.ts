/**
 * Namespace-correct text for a subtree the reader keeps verbatim.
 *
 * Several xforms keep a subtree they do not model as XML text and write it back unchanged — an `extLst`,
 * a shape's `spPr`, a group shape, a filter. Such a fragment is cut out of its document: the declarations
 * of its ancestors stay behind, and the element it is written into belongs to whatever the writer emits.
 * Two things then go wrong unless the fragment carries them itself:
 *
 * - **A prefix declared on an ancestor.** `<x14:sparklineGroup>` captured under an `<ext xmlns:x14>` the
 *   reader models, or a `r:embed` inside a group shape whose drawing declared `r` on its root, names a
 *   prefix nothing in the written part declares — the part is not well-formed.
 * - **The name view is not text.** The tags an xform sees carry the names it matches (see
 *   `namespace-prefixes.ts`); where a prefixed namespace was renamed to the default one, an element in no
 *   namespace keeps its unprefixed name, which as text would put it in that namespace.
 *
 * So a capture records each tag through a {@link FragmentNamespaces}: it gets the **text view** of the tag
 * from the normaliser driving the parse, and notes every prefix the subtree names — in element and
 * attribute names, and in the attribute values that name prefixes — that the subtree does not declare. When
 * the capture is complete, those are declared on its root element, from the bindings in force where it
 * started. A binding the part's writer puts on its own root anyway is left out, so a fragment written by
 * Excel, which relies on the root's declarations, is kept byte for byte.
 *
 * The normaliser and the writer's root declarations are found through the parse in progress
 * ({@link withNamespaceContext}): a capture is fed synchronously from inside the callbacks the driver
 * dispatches, so the context in force when it starts is the one of the document it is reading. With no
 * parse in progress — an xform driven by hand — a capture records nothing and returns its text as built.
 */
import type { PrefixNormalizer } from "@excel/xlsx/xform/namespace-prefixes";
import { xmlEncodeAttr } from "@utils/xml-encode";
import type { ParseOpenTag } from "@xml/types";

const MCE_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const XSI_NAMESPACE = "http://www.w3.org/2001/XMLSchema-instance";
const MCE_PREFIX_LISTS = new Set(["Ignorable", "MustUnderstand"]);
const MCE_QNAME_LISTS = new Set(["ProcessContent", "PreserveElements", "PreserveAttributes"]);
const NO_BINDINGS: ReadonlyMap<string, string> = new Map();

/** The declarations a part's writer puts on its root element, whatever the source declared. */
export interface WriterNamespaces {
  /** Always written: `{ xmlns: …, "xmlns:r": … }`; other attributes in it are ignored. */
  readonly fixed: Readonly<Record<string, string>>;
  /**
   * Whether the writer also writes back the `xmlns:*` declarations the source's root element carried
   * (where `fixed` does not already bind that prefix). The default namespace is never taken from it.
   */
  readonly preservesDocumentPrefixes?: boolean;
}

/** An xform that reads a whole part, and so can say what that part's writer declares. */
export interface PartXform {
  readonly writerNamespaces?: WriterNamespaces;
}

/** The parse in progress, as captures see it — see {@link createNamespaceContext}. */
export interface NamespaceContext {
  readonly normalizer: PrefixNormalizer;
  readonly writer: WriterNamespaces | undefined;
  /** `writer`'s bindings, prefix → URI, computed by the first capture. */
  provided?: ReadonlyMap<string, string>;
}

let current: NamespaceContext | undefined;

/** The context for a parse of `xform`'s part, normalised by `normalizer`. */
export function createNamespaceContext(
  normalizer: PrefixNormalizer,
  xform: object
): NamespaceContext {
  return { normalizer, writer: (xform as PartXform).writerNamespaces };
}

/**
 * Run `dispatch` — which feeds the parser, and so the xform's callbacks — with `context` as the source of
 * text views for any capture it starts.
 */
export function withNamespaceContext<T>(context: NamespaceContext, dispatch: () => T): T {
  const previous = current;
  current = context;
  try {
    return dispatch();
  } finally {
    current = previous;
  }
}

/**
 * The namespace bindings in force inside the element the parse in progress opened last (`""`: the
 * default), as text binds them — or `undefined` with no parse in progress.
 */
export function namespacesInForce(): ReadonlyMap<string, string> | undefined {
  return current?.normalizer.bindings();
}

/**
 * What the writer declares on its root, prefix → URI. Asked for by a capture, which starts inside the
 * document element, so the source's own declarations are known by then. A part xform that does not say
 * is taken to write its root's declarations back as they were — the assumption every capture made before
 * this was written down.
 */
function providedBy(context: NamespaceContext): ReadonlyMap<string, string> {
  if (context.provided !== undefined) {
    return context.provided;
  }
  const document = context.normalizer.documentBindings();
  if (context.writer === undefined) {
    context.provided = document;
    return document;
  }
  const out = new Map<string, string>();
  for (const key in context.writer.fixed) {
    if (key === "xmlns" || key.startsWith("xmlns:")) {
      out.set(key === "xmlns" ? "" : key.slice(6), context.writer.fixed[key]);
    }
  }
  if (context.writer.preservesDocumentPrefixes) {
    for (const [prefix, uri] of document) {
      if (prefix !== "" && !out.has(prefix)) {
        out.set(prefix, uri);
      }
    }
  }
  context.provided = out;
  return out;
}

function prefixOf(qname: string): string {
  const colon = qname.indexOf(":");
  return colon === -1 ? "" : qname.slice(0, colon);
}

export interface FragmentOptions {
  /**
   * Declarations the element the fragment is written back into carries itself (`{ "xmlns:x14": … }`),
   * when the xform keeps them: the fragment need not repeat those.
   */
  readonly enclosing?: Readonly<Record<string, string>>;
  /**
   * For a fragment cut from a parsed tree rather than a parse in progress: the bindings in force inside
   * its root (prefix → URI, `""` the default), what the part's writer declares, and the default
   * namespace of the document element. The tags passed to `open` are then the text views.
   */
  readonly scope?: {
    readonly inForce: ReadonlyMap<string, string>;
    readonly provided: ReadonlyMap<string, string>;
    readonly documentDefault: string;
  };
}

/**
 * Records one captured subtree's tags. Create it while the subtree's root element is being opened, pass
 * every start tag to {@link open} and every end tag to {@link close} — they return what to write — and
 * hand the finished text to {@link finish}.
 */
export class FragmentNamespaces {
  private readonly normalizer: PrefixNormalizer | undefined;
  /** Bindings in force inside the root element, as text binds them. */
  private readonly inForce: ReadonlyMap<string, string>;
  /** What the writer declares on its own root (and the enclosing element, if the xform keeps it). */
  private readonly provided: ReadonlyMap<string, string>;
  /** The fragment's own declarations in force, per open element. */
  private readonly scopes: ReadonlyMap<string, string>[] = [];
  /** Prefix → URI the root must declare. */
  private readonly needed = new Map<string, string>();
  /** The source's default namespace on its document element (`""`: none). */
  private readonly documentDefault: string;
  private rootName: string | undefined;

  constructor(options?: FragmentOptions) {
    const context = current;
    const explicit = options?.scope;
    this.normalizer = explicit === undefined ? context?.normalizer : undefined;
    this.inForce = explicit?.inForce ?? context?.normalizer.bindings() ?? NO_BINDINGS;
    this.documentDefault =
      explicit?.documentDefault ?? context?.normalizer.documentBindings().get("") ?? "";
    let provided =
      explicit?.provided ?? (context === undefined ? NO_BINDINGS : providedBy(context));
    const enclosing = options?.enclosing;
    if (enclosing !== undefined) {
      const merged = new Map(provided);
      for (const key in enclosing) {
        if (key === "xmlns" || key.startsWith("xmlns:")) {
          merged.set(key === "xmlns" ? "" : key.slice(6), enclosing[key]);
        }
      }
      provided = merged;
    }
    this.provided = provided;
  }

  /** `tag`, as the fragment's text must write it. */
  open(tag: ParseOpenTag): ParseOpenTag {
    const text = this.normalizer?.textOpened() ?? tag;
    const parent = this.scopes[this.scopes.length - 1] ?? NO_BINDINGS;
    let local = parent;
    for (const key in text.attributes) {
      if (key === "xmlns" || key.startsWith("xmlns:")) {
        if (local === parent) {
          local = new Map(parent);
        }
        (local as Map<string, string>).set(
          key === "xmlns" ? "" : key.slice(6),
          text.attributes[key]
        );
      }
    }
    this.scopes.push(local);
    this.rootName ??= text.name;
    if (this.inForce === NO_BINDINGS) {
      return text;
    }
    this.use(prefixOf(text.name), local);
    for (const key in text.attributes) {
      const colon = key.indexOf(":");
      if (colon === -1 || key.startsWith("xmlns")) {
        continue;
      }
      const prefix = key.slice(0, colon);
      this.use(prefix, local);
      const uri = local.get(prefix) ?? this.inForce.get(prefix);
      const localName = key.slice(colon + 1);
      const value = text.attributes[key];
      if (uri === MCE_NAMESPACE && MCE_PREFIX_LISTS.has(localName)) {
        for (const entry of value.split(/\s+/)) {
          if (entry !== "") {
            this.use(entry, local);
          }
        }
      } else if (uri === MCE_NAMESPACE && MCE_QNAME_LISTS.has(localName)) {
        for (const entry of value.split(/\s+/)) {
          if (entry.indexOf(":") > 0) {
            this.use(prefixOf(entry), local);
          }
        }
      } else if (uri === XSI_NAMESPACE && localName === "type") {
        this.use(prefixOf(value.trim()), local);
      }
    }
    if (
      text.attributes.Requires !== undefined &&
      text.name.slice(text.name.indexOf(":") + 1) === "Choice" &&
      (local.get(prefixOf(text.name)) ?? this.inForce.get(prefixOf(text.name))) === MCE_NAMESPACE
    ) {
      for (const entry of text.attributes.Requires.split(/\s+/)) {
        if (entry !== "") {
          this.use(entry, local);
        }
      }
    }
    return text;
  }

  /** The end tag's name, as the fragment's text must write it. */
  close(name: string): string {
    this.scopes.pop();
    return this.normalizer?.textClosed() ?? name;
  }

  /** `xml` — the fragment, its root element first — with the declarations it needs on that root. */
  finish(xml: string): string {
    const open = `<${this.rootName}`;
    if (this.needed.size === 0 || !xml.startsWith(open)) {
      return xml;
    }
    return open + this.declarations() + xml.slice(open.length);
  }

  /** `attributes` of the root element with the declarations it needs — for a capture kept as a tree. */
  finishAttributes<T extends Record<string, unknown>>(attributes: T): T {
    if (this.needed.size === 0) {
      return attributes;
    }
    const out: Record<string, unknown> = { ...attributes };
    for (const [prefix, uri] of this.needed) {
      out[prefix === "" ? "xmlns" : `xmlns:${prefix}`] = uri;
    }
    return out as T;
  }

  private declarations(): string {
    let s = "";
    for (const [prefix, uri] of this.needed) {
      s += ` ${prefix === "" ? "xmlns" : `xmlns:${prefix}`}="${xmlEncodeAttr(uri)}"`;
    }
    return s;
  }

  private use(prefix: string, local: ReadonlyMap<string, string>): void {
    if (prefix === "xml" || local.has(prefix) || this.needed.has(prefix)) {
      return;
    }
    const uri = this.inForce.get(prefix);
    if (uri === undefined) {
      // Not declared where the capture started either: the source names it unbound, and so does this.
      return;
    }
    if (prefix === "" && uri === this.documentDefault) {
      // The document's own default namespace: the writer re-expresses the part in its default, and the
      // fragment goes with the rest of it.
      return;
    }
    const written = this.provided.get(prefix) ?? (prefix === "" ? "" : undefined);
    if (written !== uri) {
      this.needed.set(prefix, uri);
    }
  }
}
