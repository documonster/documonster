/**
 * Signatures as Markdown table cells. A table is padded to its widest cell, so one 5,000-character object type turned
 * every row of its table into a line of spaces. A line longer than {@link MAX} is clipped: alone, to its start; beside
 * the signature it is compared with, to the stretch where the two differ, so what changed is still what is shown.
 */
const MAX = 160;
const CONTEXT = 40;

const esc = (s: string): string => "`" + s.replace(/`/g, "'").replace(/\|/g, "\\|") + "`";

/** One signature line, at most {@link MAX} characters of it: the window around `[from, to)`, marked `…` where cut. */
function line(text: string, from = 0, to = 0): string {
  if (text.length <= MAX) {
    return text;
  }
  const span = to - from;
  let start = span + 2 * CONTEXT >= MAX ? from - CONTEXT : from - Math.floor((MAX - span) / 2);
  start = Math.max(0, Math.min(start, text.length - MAX));
  const end = start + MAX;
  return (start > 0 ? "…" : "") + text.slice(start, end) + (end < text.length ? "…" : "");
}

/** Where `a` and `b` stop agreeing at the start, and from the end. */
function differing(a: string, b: string): { start: number; endA: number; endB: number } {
  let start = 0;
  while (start < a.length && start < b.length && a[start] === b[start]) {
    start++;
  }
  let tail = 0;
  while (
    tail < a.length - start &&
    tail < b.length - start &&
    a[a.length - 1 - tail] === b[b.length - 1 - tail]
  ) {
    tail++;
  }
  return { start, endA: a.length - tail, endB: b.length - tail };
}

/** At most this many lines (overloads) to a cell: a table is as wide as its widest cell, `<br>`s and all. */
const MAX_LINES = 4;

/**
 * A signature as a cell; `other`, when given, is the signature it is compared with. Lines (overloads) both have
 * alike are left out of a comparison, and counted, so the cell shows what changed.
 */
export function sigCell(sig: string, other?: string): string {
  if (!sig) {
    return "";
  }
  let lines = sig.split("\n");
  const others = other?.split("\n");
  let same = 0;
  if (others && lines.length > 1) {
    const kept = lines.filter(l => !others.includes(l));
    same = lines.length - kept.length;
    lines = kept;
  }
  const shown = lines.slice(0, MAX_LINES).map(text => {
    if (!others || text.length <= MAX) {
      return esc(line(text));
    }
    // Compare with the line of the other signature that shares the most of its start.
    const peer = others.reduce((best, o) =>
      differing(text, o).start > differing(text, best).start ? o : best
    );
    const { start, endA } = differing(text, peer);
    return esc(line(text, start, Math.max(endA, start + 1)));
  });
  const more = lines.length - shown.length;
  const notes = [more > 0 ? `${more} more` : "", same > 0 ? `${same} unchanged` : ""].filter(
    Boolean
  );
  return [...shown, ...(notes.length ? [`_(${notes.join(", ")})_`] : [])].join("<br>");
}
