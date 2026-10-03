/**
 * List numbering shared by every consumer of list markers (HTML, Markdown,
 * semantic IR, layout), so a document numbers the same way in every format.
 *
 * Rules, each checked against Microsoft Word's own output (see
 * `numbering-continuation.test.ts`, "matches Word"):
 * - every `w:num` that references one `w:abstractNum` shares one count per
 *   level — with or without a `w:lvlOverride`; an interrupting paragraph does
 *   not reset it;
 * - a `w:startOverride` on (numId, level) restarts that shared count at its
 *   value the first time that numId is used at that level, and is the value
 *   the level restarts at whenever an item of that numId opens it again;
 * - a replacement `w:lvl` changes how the level is drawn, not its count;
 * - a deeper level restarts when a shallower one appears, unless its
 *   `w:lvlRestart` says otherwise (`0` = never, `n` = only when a level at or
 *   above level `n`, 1-based, appears);
 * - `w:isLgl` is a display rule; see {@link placeholderFormat}.
 */

import { resolveNumberingLevelDef } from "@word/query/style-resolve";
import type { DocxDocument, NumberingLevel } from "@word/types";

/** Tracks the current number of each list level during one document walk. */
export interface NumberingCounter {
  /**
   * Record one list item — ordered or bulleted — at `level` and return the
   * number it displays. Deeper levels restart afterwards.
   */
  next(numId: number, level: number): number;
  /**
   * The number `level` currently displays for `numId`'s list, for multi-level
   * marker templates such as `"%1.%2."`. A level with no item yet shows its start.
   */
  current(numId: number, level: number): number;
  /** Effective level definition — see {@link resolveNumberingLevelDef}. */
  levelDef(numId: number, level: number): NumberingLevel | undefined;
}

/**
 * Create a counter bound to a document's numbering definitions.
 *
 * Definitions are memoised for the counter's lifetime (one walk over an
 * unchanging document).
 */
export function createNumberingCounter(doc: DocxDocument): NumberingCounter {
  const values = new Map<string, Map<number, number>>();
  const defs = new Map<string, NumberingLevel | undefined>();
  const instances = new Map(doc.numberingInstances?.map(n => [n.numId, n]));
  /** (numId, level) pairs whose `w:startOverride` has already taken effect. */
  const overridesApplied = new Set<string>();

  const levelDef = (numId: number, level: number): NumberingLevel | undefined => {
    const key = `${numId}:${level}`;
    if (defs.has(key)) {
      return defs.get(key);
    }
    const def = resolveNumberingLevelDef(doc, numId, level);
    defs.set(key, def);
    return def;
  };

  /** The count a `numId` advances: its abstract definition's. */
  const counts = (numId: number): Map<number, number> => {
    const instance = instances.get(numId);
    const key = instance ? `a:${instance.abstractNumId}` : `n:${numId}`;
    let levels = values.get(key);
    if (!levels) {
      levels = new Map();
      values.set(key, levels);
    }
    return levels;
  };

  const startOverride = (numId: number, level: number): number | undefined =>
    instances.get(numId)?.overrides?.find(o => o.level === level)?.startOverride;

  /** Whether `deeper` restarts when an item at `level` appears (`w:lvlRestart`). */
  const restartsAfter = (numId: number, deeper: number, level: number): boolean => {
    const restart = levelDef(numId, deeper)?.restartAfterLevel;
    if (restart === undefined) {
      return true;
    }
    return restart > 0 && level <= restart - 1;
  };

  return {
    next(numId: number, level: number): number {
      const levels = counts(numId);
      const previous = levels.get(level);
      const override = startOverride(numId, level);
      const key = `${numId}:${level}`;
      let value: number;
      if (override !== undefined && !overridesApplied.has(key)) {
        overridesApplied.add(key);
        value = override;
      } else if (previous === undefined) {
        value = override ?? levelDef(numId, level)?.start ?? 1;
      } else {
        value = previous + 1;
      }
      levels.set(level, value);
      for (const deeper of levels.keys()) {
        if (deeper > level && restartsAfter(numId, deeper, level)) {
          levels.delete(deeper);
        }
      }
      return value;
    },
    current(numId: number, level: number): number {
      return (
        counts(numId).get(level) ??
        startOverride(numId, level) ??
        levelDef(numId, level)?.start ??
        1
      );
    },
    levelDef
  };
}

/**
 * Number format for each `%N` placeholder of `def`'s level text: legal
 * numbering (`w:isLgl`) shows every level as an arabic numeral, so "1.2"
 * stays "1.2" under a level whose parent counts in roman numerals.
 */
export function placeholderFormat(
  counter: NumberingCounter,
  numId: number,
  def: NumberingLevel,
  referencedLevel: number
): NumberingLevel["format"] {
  if (def.isLegalNumberingStyle) {
    return "decimal";
  }
  if (referencedLevel === def.level) {
    return def.format;
  }
  return counter.levelDef(numId, referencedLevel)?.format ?? "decimal";
}
