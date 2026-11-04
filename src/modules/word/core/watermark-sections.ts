/**
 * Which headers a document's watermark lives in.
 *
 * `DocxDocument.watermark` is the document's watermark, so it is drawn on
 * every page, as Word's Design › Watermark does: in every header a section
 * shows — its default header, its first-page header when it has a distinct
 * first page (`titlePage`), and its even-page header when the document has
 * distinct even and odd headers (`evenAndOddHeaders`). A section that names
 * no header of a type shows the previous section's (ECMA-376 §17.10.5); the
 * first section to show none at all is given a header holding only the
 * watermark, which later sections then inherit in the same way.
 *
 * The reader applies the converse: a watermark is the document's only when
 * every header shown anywhere draws it, the same in each. Otherwise it is
 * header content and is kept as such.
 */

import type { HeaderFooterRef, HeaderFooterType, SectionProperties } from "@word/types";

/** The header-related part of a section. */
export type SectionHeaders = Pick<SectionProperties, "headers" | "titlePage">;

/** The header types a section shows, in a fixed order. */
export function shownHeaderTypes(
  section: SectionHeaders,
  evenAndOddHeaders: boolean
): HeaderFooterType[] {
  const types: HeaderFooterType[] = ["default"];
  if (section.titlePage) {
    types.push("first");
  }
  if (evenAndOddHeaders) {
    types.push("even");
  }
  return types;
}

/** A header a section shows. */
export interface HeaderSlot {
  /** Index of the section, in document order. */
  readonly section: number;
  readonly type: HeaderFooterType;
  /** The header shown, or `undefined` when the section shows none. */
  readonly rId: string | undefined;
  /** Whether the section names it itself, rather than inheriting it. */
  readonly own: boolean;
  /** The header of this type the previous sections leave in effect. */
  readonly inherited: string | undefined;
}

/** Every header every section shows, in document order. */
export function headerSlots(
  sections: readonly SectionHeaders[],
  evenAndOddHeaders: boolean
): HeaderSlot[] {
  const inEffect = new Map<HeaderFooterType, string>();
  const slots: HeaderSlot[] = [];
  sections.forEach((section, index) => {
    const before = new Map(inEffect);
    const own = ownReferences(section.headers);
    for (const [type, rId] of own) {
      inEffect.set(type, rId);
    }
    for (const type of shownHeaderTypes(section, evenAndOddHeaders)) {
      slots.push({
        section: index,
        type,
        rId: inEffect.get(type),
        own: own.has(type),
        inherited: before.get(type)
      });
    }
  });
  return slots;
}

/** A section's own references by type; the first of a type counts. */
function ownReferences(
  headers: readonly HeaderFooterRef[] | undefined
): Map<HeaderFooterType, string> {
  const own = new Map<HeaderFooterType, string>();
  for (const ref of headers ?? []) {
    if (!own.has(ref.type)) {
      own.set(ref.type, ref.rId);
    }
  }
  return own;
}

/**
 * Plans where the watermark goes, one section at a time in document order —
 * the streaming writer sees its sections one by one.
 */
export interface WatermarkPlan {
  /**
   * The header types `section` must name the watermark-only header for:
   * those it shows but neither names nor inherits.
   */
  section(section: SectionHeaders): HeaderFooterType[];
  /** The headers shown somewhere so far; each draws the watermark. */
  readonly hosts: ReadonlySet<string>;
  /** Whether some section needed the watermark-only header. */
  readonly needsWatermarkHeader: boolean;
}

/** A {@link WatermarkPlan}; `watermarkHeader` stands for the watermark-only header. */
export function planWatermark(evenAndOddHeaders: boolean, watermarkHeader: string): WatermarkPlan {
  const inEffect = new Map<HeaderFooterType, string>();
  const hosts = new Set<string>();
  let needed = false;
  return {
    hosts,
    get needsWatermarkHeader() {
      return needed;
    },
    section(section) {
      for (const [type, rId] of ownReferences(section.headers)) {
        inEffect.set(type, rId);
      }
      const missing: HeaderFooterType[] = [];
      for (const type of shownHeaderTypes(section, evenAndOddHeaders)) {
        const rId = inEffect.get(type);
        if (rId === undefined) {
          missing.push(type);
          inEffect.set(type, watermarkHeader);
          needed = true;
        } else if (rId !== watermarkHeader) {
          hosts.add(rId);
        }
      }
      return missing;
    }
  };
}
