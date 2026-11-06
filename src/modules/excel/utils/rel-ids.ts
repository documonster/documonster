/**
 * Relationship-id allocation for the package writers.
 *
 * Both allocators are amortised O(1) per id. The obvious forms are not: scanning the
 * whole relationship list for its highest id, or probing `rId1`, `rId2`, … from the
 * start, costs O(n) per id, so a sheet with n hyperlinks took O(n²) to write — 8000
 * links spent most of 1.7 s numbering them.
 */

/**
 * Allocate `rId{max + 1}` for a relationship list, where `max` is the highest numeric
 * `rIdN` the list holds.
 *
 * Highest-plus-one rather than `length + 1`, because a list may keep a subset of ids
 * read from a file (a preserved drawing group's rels), and its length then says nothing
 * about which ids are free. Ids that are not of the `rIdN` form are ignored.
 *
 * Each list's highest id is remembered with how far it has been read, so a call reads
 * only the entries appended since the last one. A list that shrank is read again in full;
 * a list is never edited in place by the writers, which is what makes appending the only
 * other case.
 */
export function createMaxRelIds(): (rels: readonly unknown[]) => string {
  const seen = new WeakMap<readonly unknown[], { read: number; max: number }>();
  return rels => {
    let state = seen.get(rels);
    if (!state || state.read > rels.length) {
      state = { read: 0, max: 0 };
      seen.set(rels, state);
    }
    for (let i = state.read; i < rels.length; i++) {
      const id = (rels[i] as { Id?: unknown } | undefined)?.Id;
      const m = typeof id === "string" ? /^rId(\d+)$/.exec(id) : null;
      if (m) {
        state.max = Math.max(state.max, Number(m[1]));
      }
    }
    state.read = rels.length;
    return `rId${state.max + 1}`;
  };
}

/**
 * Allocate the lowest free `rIdN` (from `rId1`) not in `taken`, and add it to `taken`.
 *
 * The search resumes where the previous one stopped: `taken` only grows, so every id
 * below that point is still taken and need not be probed again. The ids are the same as
 * a fresh search from `rId1` would give.
 */
export function createFirstFreeRelIds(taken: Set<string>): () => string {
  let n = 1;
  return () => {
    while (taken.has(`rId${n}`)) {
      n++;
    }
    const id = `rId${n++}`;
    taken.add(id);
    return id;
  };
}
