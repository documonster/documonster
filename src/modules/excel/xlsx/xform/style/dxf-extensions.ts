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
