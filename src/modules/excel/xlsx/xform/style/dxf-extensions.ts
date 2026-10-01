import { xmlEncodeAttr } from "@xml/encode";

/**
 * A differential format's `<extLst>`, carried through a round trip as XML.
 *
 * `CT_Dxf` ends in an `extLst`: the schema's extension point, where a later version of the format — or another
 * producer — puts what this one does not define. Nothing here can interpret it, and dropping it is a silent loss
 * of content the file legitimately carried. So the reader keeps the element as written and the writer puts it
 * back, after the six facets it does understand, which is where the schema requires it.
 *
 * **Held under a symbol, not a field.** It is not a formatting property a caller sets or reads, and a `Style`
 * field would publish it as one. A symbol-keyed property is invisible to `Object.keys` and `JSON.stringify` —
 * so it neither shows up in a style's contents nor changes how two styles compare — yet it is copied by object
 * spread, which is how the writer takes its own copy of a format. `structuredClone` does not copy it; a caller
 * who clones a style is building a new one, and a new one has no extensions.
 */
export const DXF_EXTENSIONS: unique symbol = Symbol("documonster.dxfExtensions");

/** A style that may carry a preserved `<extLst>`. */
interface WithExtensions {
  [DXF_EXTENSIONS]?: string;
}

/** The preserved `<extLst>` of a differential format, if it has one. */
export function dxfExtensionsOf(style: object | undefined): string | undefined {
  return (style as WithExtensions | undefined)?.[DXF_EXTENSIONS];
}

/** Attach a preserved `<extLst>` to a differential format. */
export function setDxfExtensions(style: object, xml: string): void {
  (style as WithExtensions)[DXF_EXTENSIONS] = xml;
}

/**
 * Make a preserved fragment's namespace prefixes resolvable wherever it is written.
 *
 * An extension's elements are prefixed (`x14:…`), and the prefix may be declared on an ancestor — the source's
 * `<styleSheet>` — rather than inside the fragment. This writer emits its own `<styleSheet>` with its own
 * declarations, so a prefix the source declared there would be written unbound, and the part would not be
 * well-formed XML. Each prefix the fragment uses and its root element does not declare is therefore declared on
 * that root, from `inherited`. A redundant declaration is harmless; a missing one is not.
 */
export function withInheritedNamespaces(
  xml: string,
  inherited: Readonly<Record<string, string>>
): string {
  const rootEnd = xml.search(/\/?>/);
  if (rootEnd < 0) {
    return xml;
  }
  const root = xml.slice(0, rootEnd);
  const used = new Set<string>();
  for (const match of xml.matchAll(
    /<\/?([A-Za-z_][\w.-]*):[\w.-]+|\s([A-Za-z_][\w.-]*):[\w.-]+\s*=/g
  )) {
    const prefix = match[1] ?? match[2]!;
    if (prefix !== "xmlns" && prefix !== "xml") {
      used.add(prefix);
    }
  }
  let declarations = "";
  for (const prefix of used) {
    const uri = inherited[prefix];
    if (uri !== undefined && !new RegExp(`\\sxmlns:${prefix}\\s*=`).test(root)) {
      declarations += ` xmlns:${prefix}="${xmlEncodeAttr(uri)}"`;
    }
  }
  return declarations === "" ? xml : root + declarations + xml.slice(rootEnd);
}
