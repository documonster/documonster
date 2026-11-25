/**
 * Ordering check: does each part have the root scope its name implies, and are its records where the record grammar
 * puts them?
 *
 *  * a part must be wrapped in the scope its role requires — a worksheet in
 *    `BrtBeginSheet`, the shared strings in `BrtBeginSst`;
 *  * a part with a grammar (`grammar.ts`, MS-XLSB's productions copied verbatim) has the records of every scope it
 *    can rank in the grammar's order; a part without one keeps the few `precedes` pairs its role lists;
 *  * a cell record has to be inside `BrtBeginSheetData`, and a row's cells after a
 *    `BrtRowHdr` — a cell emitted outside either has no row to belong to.
 */

import type { FramedPart } from "@excel/utils/xlsb-validator/check-framing";
import {
  grammarScopes,
  isDelimiterOnly,
  isWrapper,
  isWrapperEnd,
  type XlsbGrammar
} from "@excel/utils/xlsb-validator/grammar";
import type { XlsbReporter } from "@excel/utils/xlsb-validator/reporter";
import {
  CELL_RECORD_NAMES,
  ROW_RECORD_NAMES,
  XLSB_PART_ROLES,
  type XlsbPartRole
} from "@excel/utils/xlsb-validator/roles";
import { MS_XLSB_RECORD_NAMES } from "@excel/xlsb/spec/record-names";
import { recordSpec } from "@excel/xlsb/spec/records";

export function checkOrdering(
  framed: FramedPart,
  role: XlsbPartRole,
  reporter: XlsbReporter
): void {
  const expectations = XLSB_PART_ROLES[role];
  const names = framed.records.map(record => recordSpec(record.id)?.name);

  if (expectations.rootScope) {
    // Past any leading wrapper block: a part may open with one before its root — the table part's grammar is
    // `[ACUID] BrtBeginList …`, and every Excel-authored table part carries that `BrtACBegin … BrtACEnd` — so
    // the first record proper is the first one outside them. The scope check treats the wrappers the same way.
    let wrapperDepth = 0;
    const first = names.find(name => {
      if (name === undefined) {
        return false;
      }
      if (isWrapper(name)) {
        wrapperDepth++;
        return false;
      }
      if (isWrapperEnd(name)) {
        wrapperDepth = Math.max(0, wrapperDepth - 1);
        return false;
      }
      return wrapperDepth === 0;
    });
    if (first !== expectations.rootScope) {
      reporter.error(
        "scope-missing-root",
        `a ${role} part must open with ${expectations.rootScope}, not ${first ?? "an unknown record"}`,
        { part: framed.part, offset: framed.records[0]?.offset ?? 0 }
      );
    }
  }

  // Absence, which every other check here is blind to. A part can be perfectly framed, correctly
  // scoped and consistently ordered while missing a record Excel writes into every file — and the
  // result is coherent and unopenable, which is the worst combination for a validator to pass.
  //
  // A **warning**, not an error, and the distinction is load-bearing in both directions. A file
  // without these is still readable, and this validator is pointed at input as well as output: two
  // workbooks in the reference corpus are hand-reduced bug reports rather than Excel's own output,
  // and refusing them would be refusing files this library reads correctly. What the warning says is
  // narrower and true — Excel writes these into every file it produces, so a package lacking them
  // has no evidence of being openable.
  for (const required of expectations.requiredRecords ?? []) {
    if (!names.includes(required)) {
      reporter.warning(
        "record-missing-required",
        `a ${role} part has no ${required}; every Excel-authored workbook has one, so a package without it has no evidence of being openable`,
        { part: framed.part, offset: 0 }
      );
    }
  }

  for (const [before, after] of expectations.precedes ?? []) {
    const beforeIndex = names.indexOf(before);
    const afterIndex = names.indexOf(after);
    if (beforeIndex !== -1 && afterIndex !== -1 && beforeIndex > afterIndex) {
      reporter.error("ordering-out-of-place", `${before} must appear before ${after}`, {
        part: framed.part,
        offset: framed.records[beforeIndex]!.offset
      });
    }
  }

  if (expectations.grammar) {
    checkGrammarOrder(framed, expectations.grammar, reporter);
  }

  if (!expectations.cellsInSheetData) {
    return;
  }

  let inSheetData = false;
  let sawRowHeader = false;
  for (const record of framed.records) {
    if (reporter.capped) {
      return;
    }
    const name = recordSpec(record.id)?.name;
    if (name === "BrtBeginSheetData") {
      inSheetData = true;
      sawRowHeader = false;
      continue;
    }
    if (name === "BrtEndSheetData") {
      inSheetData = false;
      continue;
    }
    if (name !== undefined && ROW_RECORD_NAMES.has(name)) {
      if (!inSheetData) {
        reporter.error("ordering-outside-scope", `${name} appears outside BrtBeginSheetData`, {
          part: framed.part,
          offset: record.offset
        });
      }
      sawRowHeader = true;
      continue;
    }
    if (name && CELL_RECORD_NAMES.has(name)) {
      if (!inSheetData) {
        reporter.error("ordering-outside-scope", `${name} appears outside BrtBeginSheetData`, {
          part: framed.part,
          offset: record.offset
        });
      } else if (!sawRowHeader) {
        reporter.error(
          "ordering-outside-scope",
          `${name} appears before any BrtRowHdr, so it belongs to no row`,
          { part: framed.part, offset: record.offset }
        );
      }
    }
  }
}

const SCOPES = new Map<XlsbGrammar, ReadonlyMap<string, ReadonlyMap<string, number>>>();

/** One open scope: how its direct children are ranked (when the grammar knows), and how far they have got. */
interface Frame {
  readonly ranks: ReadonlyMap<string, number> | undefined;
  /** A `BrtFRTBegin`/`BrtACBegin` wrapper, whose first record stands for it in the enclosing scope. */
  readonly wrapper: boolean;
  /** Whether the wrapper's first record has not been seen yet. */
  pending: boolean;
  maxRank: number;
  maxName: string;
}

/**
 * Rank the records of every scope by the part's grammar (`grammar.ts`): the root's children, and the children of
 * each scope whose production is published (a view's pane before its selections, the supporting books before
 * `BrtExternSheet`, …).
 *
 * Depth comes from record names, because the scope table knows only the scopes this library reads: `BrtBegin…`
 * opens, `BrtEnd…` closes, and so do the two wrappers. A collection behind a wrapper is named, in the scope that
 * holds it, by the record after the wrapper — of what a wrapper holds, the walk (`iterateInterpretableRecords`)
 * passes on only the future records this library models, so that is all that is ranked there. A record a scope's
 * grammar does not place is skipped rather than reported — it may be a newer feature's.
 */
function checkGrammarOrder(framed: FramedPart, grammar: XlsbGrammar, reporter: XlsbReporter): void {
  let scopes = SCOPES.get(grammar);
  if (!scopes) {
    scopes = grammarScopes(grammar);
    SCOPES.set(grammar, scopes);
  }
  const stack: Frame[] = [];
  for (const record of framed.records) {
    if (reporter.capped) {
      return;
    }
    const name = MS_XLSB_RECORD_NAMES.get(record.id);
    if (name === undefined) {
      continue;
    }
    if (isDelimiterOnly(name)) {
      stack.pop();
      continue;
    }
    const top = stack[stack.length - 1];
    // The scope this record is a child of, for ranking: a wrapper's first record belongs to the scope around it.
    let owner: Frame | undefined;
    if (top?.wrapper) {
      if (top.pending) {
        top.pending = false;
        for (let i = stack.length - 1; i >= 0 && owner === undefined; i--) {
          owner = stack[i]!.wrapper ? undefined : stack[i];
        }
      }
    } else {
      owner = top;
    }
    if (owner?.ranks) {
      const rank = owner.ranks.get(name);
      if (rank !== undefined) {
        if (rank < owner.maxRank) {
          reporter.error(
            "ordering-out-of-place",
            `${name} appears after ${owner.maxName}; the record grammar (${grammar.root}) places it earlier`,
            { part: framed.part, offset: record.offset }
          );
        } else {
          owner.maxRank = rank;
          owner.maxName = name;
        }
      }
    }
    if (isWrapper(name) || /^[Bb]rtBegin/.test(name)) {
      stack.push({
        ranks: isWrapper(name) ? undefined : scopes.get(name),
        wrapper: isWrapper(name),
        pending: isWrapper(name),
        maxRank: -1,
        maxName: ""
      });
    }
  }
}
