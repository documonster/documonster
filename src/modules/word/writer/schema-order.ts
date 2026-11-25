/**
 * Emitting a container's children in the order its content model requires.
 *
 * A writer used to emit each container's children as a run of `if` blocks, so their order was the order the blocks
 * happened to be written in — and for `w:pPr` that was wrong in four places at once (`numPr` after the spacing, the
 * East Asian switches after the justification, the section before the mark's run properties, …). Moving the blocks
 * would fix this release and not the next property someone adds. Here a writer names each child it has and how to
 * render it, in whatever order is convenient, and the position comes from the schema (`ooxml-sequences.ts`).
 */

import { DocxWriteError } from "@word/errors";
import { OOXML_SEQUENCES, type OoxmlSequenceType } from "@word/writer/ooxml-sequences";
import { sequenceRanks } from "@xml/sequence-order";

/** A child to emit: its name (`w:spacing`), and how to write it. `false`/`undefined` stand for "nothing to write". */
export type OrderedChild = readonly [localName: string, render: () => void] | false | undefined;

/** Each content model's positions by child local name — shared with the order check (`order-check.ts`). */
export const SEQUENCE_RANKS = sequenceRanks(OOXML_SEQUENCES);

/**
 * Word's extension namespaces: `w14:`, `w15:` and the `w16…:` family (`w16cid:`, `w16se:`, `w16du:`, …). Not `w10:`,
 * which is VML's word namespace rather than an extension.
 */
const EXTENSION_PREFIX = /^w1[4-9][a-z0-9]*:/;

/**
 * Render `children` sorted by their position in `type`. The sort is stable, so children that share a position (the
 * members of an `xsd:choice`) keep the order given.
 *
 * An extension element (`w14:`, `w15:`, `w16…:`) inside a container of another namespace is not in that
 * container's content model, and Word ignores it wherever it sits (`mc:Ignorable`); it keeps the position of the
 * nearest standard child present before it in the list, or comes first when there is none. Any other child must be
 * in the type's namespace — an extension container (`w14:CT_SdtCheckbox`) ranks its own `w14:` children — and named by its content model:
 * a name that is not is a writer bug — there is no right place for it — and throws rather than landing somewhere
 * plausible.
 */
export function emitInSchemaOrder(
  type: OoxmlSequenceType,
  children: readonly OrderedChild[]
): void {
  const ranks = prefixedRanks(type);
  const renders: (() => void)[] = [];
  const positions: number[] = [];
  let previousRank = -1;
  let inOrder = true;
  for (const child of children) {
    if (!child) {
      continue;
    }
    const [name, render] = child;
    let rank = ranks.get(name);
    if (rank === undefined) {
      if (!EXTENSION_PREFIX.test(name) || name.startsWith(ranks.prefix)) {
        throw new DocxWriteError(`internal: <${name}> is not a child of ${type}`);
      }
      rank = previousRank;
    }
    if (rank < previousRank) {
      inOrder = false;
    }
    previousRank = rank;
    renders.push(render);
    positions.push(rank);
  }
  // A writer mostly lists its children in schema order already: then there is nothing to sort, and writing a
  // document does not pay for one per property container.
  if (!inOrder) {
    const order = renders.map((_, index) => index);
    order.sort((a, b) => positions[a]! - positions[b]! || a - b);
    for (const index of order) {
      renders[index]!();
    }
    return;
  }
  for (const render of renders) {
    render();
  }
}

/** A content model's positions by prefixed child name (`w:spacing`), and the prefix itself — built once per type. */
type PrefixedRanks = ReadonlyMap<string, number> & { readonly prefix: string };

const PREFIXED_RANKS = new Map<OoxmlSequenceType, PrefixedRanks>();

function prefixedRanks(type: OoxmlSequenceType): PrefixedRanks {
  let ranks = PREFIXED_RANKS.get(type);
  if (ranks === undefined) {
    const prefix = type.slice(0, type.indexOf(":") + 1);
    const map = new Map<string, number>();
    const model = SEQUENCE_RANKS.get(type)!;
    for (const [name, rank] of model) {
      const declaredIn = model.declaredIn.get(name);
      map.set(declaredIn === undefined ? prefix + name : `${declaredIn}:${name}`, rank);
    }
    ranks = Object.assign(map, { prefix });
    PREFIXED_RANKS.set(type, ranks);
  }
  return ranks;
}
