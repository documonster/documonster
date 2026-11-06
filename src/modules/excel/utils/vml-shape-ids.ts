/**
 * Workbook-wide allocation of legacy VML shape ids.
 *
 * Excel numbers VML shapes (notes, form controls, header/footer pictures) in blocks of 1024: block `b`
 * holds `_x0000_s{b * 1024 + 1}` … `_x0000_s{b * 1024 + 1023}`, and each VML part declares the blocks it
 * uses in `<o:idmap data="…">`. The blocks belong to the **workbook**, not to the part: Excel gives the
 * first sheet's drawing block 1, the next block 2, and so on. Two parts claiming the same block — every
 * sheet numbering its notes from 1025 — leaves shapes with the same id in one workbook, and Excel hands
 * an id it believes free to the next shape it creates on either sheet.
 *
 * One space is created per write. A part asks for its ids once; ids a file carried (a checkbox's
 * `shapeId`, which the sheet XML and the drawing anchor also name) are kept when their blocks are still
 * free, and renumbered together otherwise. Blocks a later part's preserved ids live in can be reserved
 * up front, so an earlier part that only needs fresh ids does not take them.
 */

/** Shape ids per block. */
export const VML_BLOCK_SIZE = 1024;

export interface VmlShapeIdSpace {
  /** Blocks a part has been given. */
  readonly claimed: Set<number>;
  /** Blocks kept for the owner whose preserved ids live in them. */
  readonly reserved: Map<number, object>;
}

/** The ids one VML part was given. */
export interface VmlPartShapeIds {
  /** The requested `fixed` ids, in order — the same values where they could be kept. */
  readonly fixed: number[];
  /** `count` fresh ids. */
  readonly free: number[];
}

export function createVmlShapeIdSpace(): VmlShapeIdSpace {
  return { claimed: new Set(), reserved: new Map() };
}

/** The block `shapeId` falls in. */
export function vmlBlockOf(shapeId: number): number {
  return Math.floor(shapeId / VML_BLOCK_SIZE);
}

/** Whether `shapeId` is one Excel would number: inside a block, never its first or last slot. */
function isBlockId(shapeId: number): boolean {
  const local = shapeId % VML_BLOCK_SIZE;
  return Number.isInteger(shapeId) && shapeId > VML_BLOCK_SIZE && local >= 1;
}

/** Keep the blocks `shapeIds` fall in for `owner`; the first owner to reserve a block keeps it. */
export function reserveVmlShapeIds(
  space: VmlShapeIdSpace,
  owner: object,
  shapeIds: Iterable<number>
): void {
  for (const id of shapeIds) {
    if (isBlockId(id) && !space.reserved.has(vmlBlockOf(id))) {
      space.reserved.set(vmlBlockOf(id), owner);
    }
  }
}

function canClaim(space: VmlShapeIdSpace, block: number, owner: object | undefined): boolean {
  if (block < 1 || space.claimed.has(block)) {
    return false;
  }
  const holder = space.reserved.get(block);
  return holder === undefined || holder === owner;
}

/** Claim the lowest block nothing has claimed or reserved for someone else. */
export function claimVmlBlock(space: VmlShapeIdSpace, owner?: object): number {
  let block = 1;
  while (!canClaim(space, block, owner)) {
    block++;
  }
  space.claimed.add(block);
  return block;
}

/**
 * Give one VML part its ids: `fixed` kept where possible, plus `count` fresh ones. `owner` is what
 * {@link reserveVmlShapeIds} was called with for this part, if anything.
 */
export function allocateVmlShapeIds(
  space: VmlShapeIdSpace,
  request: { readonly owner?: object; readonly fixed?: readonly number[]; readonly count: number }
): VmlPartShapeIds {
  const { owner, fixed = [], count } = request;
  const blocks: number[] = [];
  const used = new Set<number>();
  const keep =
    fixed.every(isBlockId) &&
    new Set(fixed).size === fixed.length &&
    [...new Set(fixed.map(vmlBlockOf))].every(block => canClaim(space, block, owner));
  if (keep) {
    for (const id of fixed) {
      const block = vmlBlockOf(id);
      if (!space.claimed.has(block)) {
        space.claimed.add(block);
        blocks.push(block);
      }
      used.add(id);
    }
  }

  let blockIndex = 0;
  let local = 1;
  const next = (): number => {
    for (;;) {
      if (blockIndex === blocks.length) {
        blocks.push(claimVmlBlock(space, owner));
        local = 1;
      }
      const block = blocks[blockIndex];
      while (local < VML_BLOCK_SIZE) {
        const id = block * VML_BLOCK_SIZE + local++;
        if (!used.has(id)) {
          used.add(id);
          return id;
        }
      }
      blockIndex++;
      local = 1;
    }
  };

  const keptFixed = keep ? [...fixed] : fixed.map(() => next());
  const free = Array.from({ length: count }, () => next());
  return { fixed: keptFixed, free };
}

/**
 * The `o:idmap` `data` of a VML part holding `shapeIds`: every block one of them falls in
 * (`_x0000_s1025` is in block 1, `_x0000_s2049` in block 2), comma-separated. `1` when the part holds
 * no shape.
 */
export function vmlIdMap(shapeIds: Iterable<number>): string {
  const blocks = new Set<number>();
  for (const id of shapeIds) {
    blocks.add(vmlBlockOf(id));
  }
  return blocks.size ? [...blocks].sort((a, b) => a - b).join(",") : "1";
}
