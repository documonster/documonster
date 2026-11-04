/**
 * Package parts the workbook model carries but does not interpret.
 *
 * These types live in `core/` rather than beside the XLSX policy that produces
 * them because they describe *workbook state*, not a serialisation detail: a
 * `WorkbookData` holds them for as long as it exists, `getWorkbookModel` hands
 * them out, and every OPC-based format has the same need. XLSB is the immediate
 * reason — it is also a ZIP of OPC parts with a `[Content_Types].xml` and `.rels`
 * files, so it needs to express "a part I did not model" without importing from
 * the XLSX serialiser to say it.
 *
 * The policy that decides *which* parts to keep, resolves relationship targets
 * and prunes unreachable ones is format-specific and stays in
 * `xlsx/opaque-parts.ts`.
 */

import { ImageError } from "@excel/errors";

/**
 * A relationship carried through verbatim.
 *
 * `targetMode` is typed as `string` rather than `"Internal" | "External"` so a
 * non-standard value in the source package survives instead of being normalised
 * into something the producer never wrote.
 */
export interface OpaqueRelationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode?: string;
}

/** An inbound relationship, together with the part that declared it. */
export interface OpaqueSourceRelationship extends OpaqueRelationship {
  /** Part whose `.rels` carried this relationship, e.g. `xl/workbook.xml`. */
  readonly source: string;
}

/**
 * Whether a preserved part's inbound edge came from the package root or from the workbook part.
 *
 * **The container's own spelling of "the workbook" is not the answer, and using it lost data.** Both writers had this
 * test inline, each comparing against its own path — `xl/workbook.xml` in the XLSX writer, `xl/workbook.bin` in the
 * XLSB one. A part read from one container and written to the other therefore had no recognised inbound edge, was
 * judged unreachable, and was dropped.
 *
 * Measured on `cal-any_sheets.xlsb` → XLSX: the workbook's theme carries `accent1 = 5B9BD5` and came out as the
 * built-in `4472C4`, because the preserved theme was dropped as an orphan and the default written in its place. Every
 * cell whose colour is a `{ theme: n }` index therefore rendered in a different colour, silently. The same applies to
 * any workbook-sourced preserved part — a VBA project, workbook-level custom XML — in either direction.
 *
 * Both spellings are accepted because both are "the workbook", and the edge is re-emitted wholesale by
 * `appendOpaqueSourceRelationships` regardless of which container is being written.
 */
/**
 * Whether a content type belongs to one container's own part family.
 *
 * A preserved part is normally container-neutral — a chart, a drawing, a theme, a VBA project are the same bytes with
 * the same content type in both containers, which is why preservation works across a conversion at all. The
 * SpreadsheetML parts and their BIFF12 counterparts are the exception: a sheet, table, comments, pivot, query-table,
 * connections, external-link, metadata or calc-chain part of one kind is meaningless in a package of the other.
 *
 * This used to list only the sheet, styles and shared-strings types, so every other part of a family crossed: an XLSX
 * conversion of an XLSB carried `xl/tables/table1.bin` with `application/vnd.ms-excel.table` beside the `table1.xml`
 * written from the model. The families are recognised by their prefix instead of by a list, because a list is what
 * left them out.
 *
 * - BIFF12: `application/vnd.ms-excel.*` without an `+xml` suffix. The `+xml` members of that namespace (threaded
 *   comments, persons, rich data, slicers, timelines, control properties) are XML in both containers.
 * - SpreadsheetML: `application/vnd.openxmlformats-officedocument.spreadsheetml.*` and the macro-enabled main parts,
 *   except printer settings — an opaque blob that both containers carry under that type.
 *
 * Keyed on the content type rather than the extension, because the extension does not decide it. `.bin` is perfectly
 * legitimate in an XLSX package — `printerSettings.bin` and `vbaProject.bin` both are — so dropping preserved parts by
 * extension would throw away a workbook's macros.
 */
function containerFamily(contentType: string): "xlsx" | "xlsb" | undefined {
  if (contentType.startsWith("application/vnd.ms-excel.")) {
    if (!contentType.endsWith("+xml")) {
      return "xlsb";
    }
    return /\.macroenabled\.main\+xml$/.test(contentType) ? "xlsx" : undefined;
  }
  if (
    contentType.startsWith("application/vnd.openxmlformats-officedocument.spreadsheetml.") &&
    contentType !== "application/vnd.openxmlformats-officedocument.spreadsheetml.printersettings"
  ) {
    return "xlsx";
  }
  return undefined;
}

/**
 * Whether a preserved part belongs to the *other* container's part family, and so cannot be written into this one.
 *
 * **This is the other half of `isRootOrWorkbookSourced`, and without it that fix would have made things worse.**
 * Recognising a workbook-sourced edge across containers is what stops a preserved theme being dropped — and it also
 * kept alive parts that genuinely must not travel: reading `cal-any_sheets.xlsb` and writing XLSX then placed
 * `xl/chartsheets/sheet1.bin`, content type `application/vnd.ms-excel.chartsheet`, inside a SpreadsheetML package.
 *
 * The chart, its style and colour parts, the drawing and the theme are all container-neutral and do travel. A part
 * this library models is never preserved in the first place, so what is filtered here is what it could not model —
 * a real loss, and a caller is told through the ordinary drop report rather than left to find it in the package.
 */
export function isForeignContainerPart(part: OpaquePart, writing: "xlsx" | "xlsb"): boolean {
  const contentType = part.contentType?.toLowerCase();
  if (contentType === undefined) {
    return false;
  }
  const family = containerFamily(contentType);
  return family !== undefined && family !== writing;
}

export function isRootOrWorkbookSourced(relationship: OpaqueSourceRelationship): boolean {
  const source = relationship.source.toLowerCase();
  return source === "" || source === "xl/workbook.xml" || source === "xl/workbook.bin";
}

/**
 * A package part this library does not model, preserved for round-trip fidelity.
 *
 * Bytes alone are not enough to keep a part alive. A reader reaches a part through
 * a relationship and decides how to interpret it from its content type, and the
 * writer regenerates both `[Content_Types].xml` and every `.rels` from the model.
 * So a preserved part needs three things travelling with it:
 *
 *  * `contentType` — the `Override` it had, if any, or nothing knows what it is;
 *  * `relationships` — its *own* `.rels`, so anything it points at resolves;
 *  * `sourceRelationships` — the relationships that pointed *at* it. This is the
 *    one that is easy to forget, and forgetting it produces a package that
 *    contains a VBA project no application will ever look for.
 */
export interface OpaquePart {
  readonly path: string;
  readonly data: Uint8Array;
  readonly contentType?: string;
  readonly relationships?: readonly OpaqueRelationship[];
  readonly sourceRelationships?: readonly OpaqueSourceRelationship[];
}

/** Why a preserved part was not written back, for reporting to the caller. */
/**
 * Why a preserved part was not written back.
 *
 * **This is a *read-time* vocabulary, and that is worth saying because two of its members looked like write-time ones.**
 * `OpaqueDrop`s reach a caller through `WorkbookModel.opaqueDrops`, which the reader populates; a write never touches
 * the workbook's copy, so a reason a writer discovers has nowhere to go here. `"unreachable"` is produced by
 * `resolveReachableOpaqueParts` during a *write* and was assigned onto the throwaway model that write was given — a
 * report that had no reader. A `"foreign-sheet-part"` member was added beside it and was dead for the same reason, so it
 * is gone again.
 *
 * **Not every write-time drop is a loss, which is the other half of why this vocabulary is read-time.** An unreachable
 * part is one nothing in the written package points at, and removing it is the documented purpose of the filter — a
 * deleted sheet taking its drawing with it is correct, not damage. Reporting it as a loss made the *default*
 * `unsupported: "error"` refuse a perfectly good workbook. A part from the other container's sheet family *is* a loss —
 * an XLSB chartsheet has no form in a SpreadsheetML package — and `writeXlsbPackage` reports that one in `unsupported`,
 * beside everything else it cannot carry. An XLSX write has no such channel; see `resolveOpaqueReachability`.
 */
export type OpaqueDropReason = "stale-cache" | "invalidated-signature" | "unreachable";

/**
 * A part that was deliberately not written back, and why.
 *
 * Recorded rather than discarded because two of the three reasons are things a
 * caller may need to know about: a digital signature is removed on every write
 * that re-serialises a modelled part, and a part can become unreachable as a
 * side effect of deleting the sheet that referenced it. Neither is an error, and
 * neither should be silent.
 */
export interface OpaqueDrop {
  readonly path: string;
  readonly reason: OpaqueDropReason;
  readonly description: string;
}

/**
 * The OPC content type for an image extension, for whichever container is being written.
 *
 * **One answer, because there were three and they disagreed.** `image/${extension}` is right for the extensions
 * `ImageData` documents — `png`, `gif` and `jpeg`, which is why `jpeg` needs no case of its own — and wrong for the two
 * below. `jpg` is what people write anyway and `image/jpg` is not a registered type; SVG's IANA type is
 * `image/svg+xml`.
 *
 * The XLSB writer normalised `jpg` and said so in a comment; the XLSX writer had the rule twice — once where it emits
 * the `Default` and once where it reserves the extension — and normalised neither. So the same image, added with
 * `extension: "jpg"`, was declared `image/jpeg` in one container and `image/jpg` in the other. Measured on the same
 * workbook written both ways.
 *
 * Lower-cased by the caller, not here: OPC matches an extension case-insensitively, and the callers key maps by it.
 *
 * Every writer types a media part through here, so this is where an extension that cannot name one is refused:
 * `image/png` passed as an extension used to become a part `xl/media/image1.image/png` and a `Default` for
 * `image/png`, which no consumer can resolve. `Image.add` rejects that at the boundary; this catches a model
 * built by hand.
 */
export function imageContentTypeFor(extension: string): string {
  if (!/^[a-z0-9]+$/i.test(extension)) {
    throw new ImageError(
      `Cannot write an image with extension ${JSON.stringify(extension)}: a media part's extension must be letters and digits (png, jpeg, gif, …)`
    );
  }
  switch (extension) {
    case "svg":
      return "image/svg+xml";
    // `jpg` is not a registered type. Excel writes `image/jpeg` for both spellings.
    case "jpg":
      return "image/jpeg";
    default:
      return `image/${extension}`;
  }
}
