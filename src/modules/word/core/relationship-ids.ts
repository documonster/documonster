/**
 * Relationship-id references inside part markup.
 *
 * A relationship id is local to the part whose `.rels` declares it, so moving
 * markup between id spaces — reading a header into the document-wide image
 * table, merging one document's body into another's — has to rewrite the ids
 * the markup carries. Structured content holds them in fields; markup kept
 * verbatim (`rawXml`) holds them in attributes of the relationships
 * namespace, which is what these helpers find and rewrite.
 */

/** The relationships namespace, transitional and strict. */
const RELATIONSHIP_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships"
]);

const NAMESPACE_DECLARATION = /\bxmlns:([A-Za-z_][\w.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
const PREFIXED_ATTRIBUTE =
  /(\s)([A-Za-z_][\w.-]*):([A-Za-z_][\w.-]*)(\s*=\s*)(?:"([^"]*)"|'([^']*)')/g;

/**
 * Prefixes bound to the relationships namespace in `xml`, plus `r`, which a
 * fragment serialised out of its part uses without redeclaring it.
 */
function relationshipPrefixes(xml: string): Set<string> {
  const prefixes = new Set(["r"]);
  for (const m of xml.matchAll(NAMESPACE_DECLARATION)) {
    if (RELATIONSHIP_NAMESPACES.has(m[2] ?? m[3] ?? "")) {
      prefixes.add(m[1]!);
    }
  }
  return prefixes;
}

/** Every relationship id `xml` refers to (`r:embed`, `r:id`, `r:link`, …). */
export function findRelationshipIdReferences(xml: string): string[] {
  if (!xml.includes(":")) {
    return [];
  }
  const prefixes = relationshipPrefixes(xml);
  const out: string[] = [];
  for (const m of xml.matchAll(PREFIXED_ATTRIBUTE)) {
    if (prefixes.has(m[2]!)) {
      out.push(m[5] ?? m[6] ?? "");
    }
  }
  return out;
}

/**
 * `xml` with every relationship-id reference found in `remap` replaced by its
 * mapped id. Returns `xml` itself when nothing changes.
 */
export function rebindRelationshipIds(
  xml: string,
  remap: ReadonlyMap<string, string> | undefined
): string {
  if (!remap || remap.size === 0 || !xml.includes(":")) {
    return xml;
  }
  const prefixes = relationshipPrefixes(xml);
  return xml.replace(
    PREFIXED_ATTRIBUTE,
    (whole, space: string, prefix: string, local: string, eq: string, dq?: string, sq?: string) => {
      if (!prefixes.has(prefix)) {
        return whole;
      }
      const value = dq ?? sq ?? "";
      const next = remap.get(value);
      if (next === undefined) {
        return whole;
      }
      const quote = dq !== undefined ? '"' : "'";
      return `${space}${prefix}:${local}${eq}${quote}${next}${quote}`;
    }
  );
}
