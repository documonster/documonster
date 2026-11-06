/**
 * The `TOC` field instruction: reading the switches the model names into a
 * {@link TableOfContents}, and writing one back.
 *
 * A TOC read from a document keeps its instruction verbatim
 * ({@link TableOfContents.instruction}); the modelled options are a reading
 * of it. The writer emits the verbatim text while the options still say what
 * it says, and otherwise writes the options followed by every switch they do
 * not model (`\z`, `\u`, `\n`, `\w`, …), so changing the heading range does
 * not drop the rest of the field.
 */

import type { Mutable } from "@word/core/internal-utils";
import { isRun } from "@word/core/text-utils";
import type { FieldContent, Paragraph, TableOfContents } from "@word/types";

/** One switch of an instruction: its letter, its argument, and its source text. */
export interface TocSwitch {
  readonly letter: string;
  readonly value: string | undefined;
  readonly raw: string;
}

/** The options a TOC instruction's switches set. */
type TocOptions = Pick<
  Mutable<TableOfContents>,
  | "headingStyleRange"
  | "hyperlink"
  | "captionLabel"
  | "sequenceFieldIdentifier"
  | "leader"
  | "stylesWithLevels"
>;

/** Switches the writer builds from options; every other switch is carried through. */
const BUILT_LETTERS = new Set(["o", "h", "t", "c", "s"]);

/**
 * A switch: `\` and a letter, then optionally a quoted argument or an unquoted
 * one. An unquoted argument cannot start with `\`, which is the next switch.
 */
const SWITCH_RE = /\\(\w)(?:\s+"([^"]*)"|\s+([^\\\s][^\s]*))?/g;

/** The switches of a `TOC` instruction, in order (letters lower-cased). */
export function tocSwitches(instr: string): TocSwitch[] {
  const out: TocSwitch[] = [];
  for (const m of instr.matchAll(SWITCH_RE)) {
    out.push({ letter: m[1]!.toLowerCase(), value: m[2] ?? m[3], raw: m[0] });
  }
  return out;
}

/** Apply one switch to `toc`; `false` when the switch sets no option. */
function applySwitch(sw: TocSwitch, toc: TocOptions): boolean {
  const { letter, value } = sw;
  switch (letter) {
    case "o": // Heading level range, e.g. "1-3"
      if (value) {
        toc.headingStyleRange = value;
        return true;
      }
      return false;
    case "h": // Hyperlinks
      toc.hyperlink = true;
      return true;
    case "c": // Caption label (table of figures)
      if (value) {
        toc.captionLabel = value;
        return true;
      }
      return false;
    case "s": // Sequence field identifier
      if (value) {
        toc.sequenceFieldIdentifier = value;
        return true;
      }
      return false;
    case "p": // Page-number leader or style separator
      if (value === "." || value === "-" || value === "_") {
        toc.leader = value === "-" ? "hyphen" : value === "_" ? "underscore" : "dot";
      }
      // The writer does not emit `\p`, so it is always carried through.
      return false;
    case "t": {
      // Styles with levels: "StyleName1,Level1;StyleName2,Level2;..."
      if (!value) {
        return false;
      }
      const items: { styleName: string; level: number }[] = [];
      for (const part of value.split(";")) {
        const [styleName, levelStr] = part.split(",");
        if (styleName && levelStr) {
          items.push({ styleName: styleName.trim(), level: parseInt(levelStr, 10) });
        }
      }
      if (items.length > 0) {
        toc.stylesWithLevels = items;
        return true;
      }
      return false;
    }
    default:
      return false;
  }
}

/**
 * Parse a TOC field instruction and populate `toc` with the switches the
 * model names:
 * - `\o "1-3"` — heading style range
 * - `\h` — hyperlinks
 * - `\c "label"` — caption label (table of figures)
 * - `\s "id"` — sequence field identifier
 * - `\p "."`/`"-"`/`"_"` — leader character
 * - `\t "Style,Level;..."` — custom styles with levels
 *
 * Does nothing when the instruction is not a `TOC` field.
 */
export function parseTocInstruction(instr: string, toc: TocOptions): void {
  if (!/^\s*TOC\b/i.test(instr)) {
    return;
  }
  for (const sw of tocSwitches(instr)) {
    applySwitch(sw, toc);
  }
}

/** The switches built from `toc`'s options, in the writer's fixed order. */
function builtSwitches(toc: TableOfContents): string[] {
  const parts: string[] = [];
  if (toc.headingStyleRange) {
    parts.push(`\\o "${toc.headingStyleRange}"`);
  }
  if (toc.hyperlink) {
    parts.push("\\h");
  }
  if (toc.stylesWithLevels && toc.stylesWithLevels.length > 0) {
    const entries = toc.stylesWithLevels.map(s => `${s.styleName},${s.level}`).join(";");
    parts.push(`\\t "${entries}"`);
  }
  if (toc.captionLabel) {
    parts.push(`\\c "${toc.captionLabel}"`);
  }
  if (toc.sequenceFieldIdentifier) {
    parts.push(`\\s "${toc.sequenceFieldIdentifier}"`);
  }
  return parts;
}

/**
 * The instruction to write for `toc`: its verbatim
 * {@link TableOfContents.instruction} when the options it writes still match
 * it, and otherwise the options' switches followed by the instruction's
 * remaining ones in their original order.
 */
export function buildTocInstruction(toc: TableOfContents): string {
  const built = builtSwitches(toc);
  const verbatim = toc.instruction;
  if (verbatim === undefined || !/^\s*TOC\b/i.test(verbatim)) {
    return ` ${["TOC", ...built].join(" ")} `;
  }
  const switches = tocSwitches(verbatim);
  const read: TocOptions = {};
  const sets = switches.map(sw => applySwitch(sw, read));
  if (builtSwitches(read as TableOfContents).join(" ") === built.join(" ")) {
    return verbatim;
  }
  const builtLetters = new Set(built.map(s => s[1]!));
  const kept = switches.filter(
    (sw, i) => !(BUILT_LETTERS.has(sw.letter) && (sets[i] || builtLetters.has(sw.letter)))
  );
  return ` ${["TOC", ...built, ...kept.map(sw => sw.raw)].join(" ")} `;
}

/** A `TOC` field held in pieces by a block's entries (`TableOfContents.field === "entries"`). */
export interface TocFieldPieces {
  /** Its instruction: its begin piece's, then that of its next piece. */
  readonly instruction: string;
  readonly begin: FieldContent;
  /** Its separate piece, or its end when it has none; `undefined` when neither follows. */
  readonly next: FieldContent | undefined;
  /** Whether its end piece is among the paragraphs. */
  readonly closed: boolean;
}

/**
 * The `TOC` field `paragraphs` hold in pieces: the first field their runs
 * open, when it is one. `undefined` otherwise.
 */
export function findTocFieldPieces(paragraphs: readonly Paragraph[]): TocFieldPieces | undefined {
  let depth = 0;
  let begin: FieldContent | undefined;
  let next: FieldContent | undefined;
  for (const para of paragraphs) {
    for (const child of para.children) {
      if (!isRun(child)) {
        continue;
      }
      for (const c of child.content) {
        if (c.type !== "field" || !c.fieldChar) {
          continue;
        }
        if (c.fieldChar === "begin") {
          depth++;
          begin ??= c;
          continue;
        }
        if (depth === 1) {
          next ??= c;
          if (c.fieldChar === "end") {
            return pieces(begin!, next, true);
          }
        }
        if (c.fieldChar === "end") {
          depth = Math.max(0, depth - 1);
        }
      }
    }
  }
  return begin ? pieces(begin, next, false) : undefined;
}

function pieces(
  begin: FieldContent,
  next: FieldContent | undefined,
  closed: boolean
): TocFieldPieces | undefined {
  const instruction = begin.instruction + (next?.instruction ?? "");
  return /^\s*TOC\b/i.test(instruction) ? { instruction, begin, next, closed } : undefined;
}
