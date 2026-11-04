import {
  derivedRuleFormula,
  missingRuleFormula,
  textRuleHasSearchString,
  textRuleKind,
  textRuleReadShape,
  type TextRuleKind
} from "@excel/core/conditional-formula";
import { XlsbParseError } from "@excel/errors";
/**
 * Conditional formatting: `BrtBeginConditionalFormatting` and the `BrtBeginCFRule` records inside it.
 *
 * ```text
 * BrtBeginConditionalFormatting   ccf (rule count), fPivot, sqrfx
 *   BrtBeginCFRule                iType, iTemplate, dxfId, iPri, iParam, 2 reserved,
 *                                 flags, cbFmla1..3, strParam, rgce1..3
 *   BrtEndCFRule
 * BrtEndConditionalFormatting
 * ```
 *
 * **The rule's shape is decided by a *pair* of enumerations, not one.** `iType` says how the formatting
 * is drawn — a comparison, an expression, a colour scale, a data bar, an icon set — and `iTemplate` says
 * what the condition is. MS-XLSB lists the legal combinations explicitly, and "other combinations MUST NOT
 * be used": a `cellIs` rule is `CF_TYPE_CELLIS` + `CF_TEMPLATE_EXPR`, while everything from
 * `containsText` to `aboveAverage` is `CF_TYPE_EXPRIS` with a different template. Writing the type alone
 * and leaving the template at 0 produces a rule Excel reads as a value comparison whatever it was.
 *
 * **The formatting itself is a `dxfId`,** an index into a workbook-level `BrtBeginDXFs` collection — not a
 * style inline in the rule. So a rule that fills a cell red needs an entry in that table, and the two
 * have to be written together. `collectDxfs` is that table.
 *
 * What is *not* written: the child records for the three graphical types. A colour scale needs a
 * `BrtBeginColorScale`, a data bar a `BrtBeginDatabar`, an icon set a `BrtBeginIconSet` — the
 * specification makes each a MUST for its type. Those three are reported rather than written with the
 * rule alone, because a `CF_TYPE_GRADIENT` rule with no `BrtBeginColorScale` is a record stream Excel
 * rejects rather than a rule that renders plainly.
 */
import type { Cfvo, Color, Style } from "@excel/types";
import {
  type BiffRange,
  boundingTopLeft,
  encodeNullableWideString,
  encodeRange,
  rangeReference,
  readNullableWideString,
  readSheetRange,
  tryDecodeRange,
  XlsbReader
} from "@excel/xlsb/binary";
import { encodeColor } from "@excel/xlsb/color";
import { encodeDxf } from "@excel/xlsb/dxf";
import {
  encodeFormula,
  encodeParsedFormula,
  formulaText,
  type EncodedFormula,
  type PtgContext
} from "@excel/xlsb/formula/ptg";
import { record, type Emitted } from "@excel/xlsb/write/emit";
import type { BinaryReader } from "@utils/binary";
import { BinaryWriter, concatUint8Arrays } from "@utils/binary";

/** `CFType`, MS-XLSB 2.5.18. */
const CF_TYPE = {
  cellIs: 1,
  expris: 2,
  gradient: 3,
  databar: 4,
  filter: 5,
  multistate: 6
} as const;

/** `CFTemp`, MS-XLSB 2.5.16 — the subset this writer emits. */
const CF_TEMPLATE = {
  expr: 0x00,
  fmla: 0x01,
  // The three graphical rules, read off an XLSB Excel wrote: a colour scale is `iType` 3 with template 2, a
  // data bar 4 with 3, an icon set 6 with 4.
  gradient: 0x02,
  databar: 0x03,
  multistate: 0x04,
  filter: 0x05,
  uniqueValues: 0x07,
  containsText: 0x08,
  containsBlanks: 0x09,
  notContainsBlanks: 0x0a,
  containsErrors: 0x0b,
  notContainsErrors: 0x0c,
  aboveAverage: 0x19,
  belowAverage: 0x1a,
  duplicateValues: 0x1b
} as const;

/** `CF_TEMPLATE_EQUALABOVEAVERAGE` and `CF_TEMPLATE_EQUALBELOWAVERAGE`: the average itself included. */
const CF_TEMPLATE_EQUAL_ABOVE_AVERAGE = 0x1d;
const CF_TEMPLATE_EQUAL_BELOW_AVERAGE = 0x1e;

/**
 * `CFTextOper`, MS-XLSB 2.5.17 — the `iParam` of a `CF_TEMPLATE_CONTAINSTEXT` rule.
 *
 * **One template, four conditions, and `iParam` says which.** This wrote 0 for all of them, so a "does not contain"
 * rule was stored as "contains": Excel would highlight exactly the cells the author had asked to leave alone, and
 * the reader — which never looked at `iParam` — gave back `containsText` and wrote `NOT(ISERROR(SEARCH(…)))` for it.
 * The specification's description of `CF_TEXTOPER_ENDSWITH` repeats "does not begin with"; the name and the value
 * are the reliable half.
 */
const CF_TEXT_OPER: Readonly<Partial<Record<TextRuleKind, number>>> = {
  containsText: 0,
  notContainsText: 1,
  beginsWith: 2,
  endsWith: 3
};

/** The `CFTemp` of each text condition without a search string. */
const CF_TEXT_TEMPLATE: Readonly<Partial<Record<TextRuleKind, number>>> = {
  containsBlanks: CF_TEMPLATE.containsBlanks,
  notContainsBlanks: CF_TEMPLATE.notContainsBlanks,
  containsErrors: CF_TEMPLATE.containsErrors,
  notContainsErrors: CF_TEMPLATE.notContainsErrors
};

/** `CFTemp` values for the twelve time periods, keyed by the model's name. */
const CF_TIME_PERIOD: Readonly<Record<string, number>> = {
  today: 0x0f,
  tomorrow: 0x10,
  yesterday: 0x11,
  last7Days: 0x12,
  lastMonth: 0x13,
  nextMonth: 0x14,
  thisWeek: 0x15,
  nextWeek: 0x16,
  lastWeek: 0x17,
  thisMonth: 0x18
};

/**
 * `CFDateOper`, MS-XLSB 2.5.12 — the `iParam` a time-period rule MUST carry, keyed by the model's name.
 *
 * Its own enumeration, in its own order, starting at 0. This wrote the `CFTemp` value here as well (`0x0F` for
 * today), which BrtBeginCFRule's table rules out: "a CFDateOper that MUST be CF_TIMEPERIOD_TODAY".
 */
const CF_DATE_OPER: Readonly<Record<string, number>> = {
  today: 0,
  yesterday: 1,
  last7Days: 2,
  thisWeek: 3,
  lastWeek: 4,
  lastMonth: 5,
  tomorrow: 6,
  nextWeek: 7,
  nextMonth: 8,
  thisMonth: 9
};

/** `CFOper`, MS-XLSB 2.5.15. Note it starts at 1, not 0. */
const CF_OPER: Readonly<Record<string, number>> = {
  between: 1,
  notBetween: 2,
  equal: 3,
  notEqual: 4,
  greaterThan: 5,
  lessThan: 6,
  greaterThanOrEqual: 7,
  lessThanOrEqual: 8
};

/** `DXFId` meaning "no differential formatting". */
const NO_DXF = 0xffffffff;

/** A rule, in the shape the model holds it. */
export interface SheetCfRule {
  readonly type?: string;
  readonly priority?: number;
  readonly formulae?: readonly (string | number)[];
  readonly operator?: string;
  readonly text?: string;
  readonly timePeriod?: string;
  readonly rank?: number;
  readonly percent?: boolean;
  readonly bottom?: boolean;
  readonly aboveAverage?: boolean;
  readonly equalAverage?: boolean;
  readonly stdDev?: number;
  readonly stopIfTrue?: boolean;
  readonly style?: Partial<Style>;
  /**
   * The range the rule applies to, copied down from its block.
   *
   * Not part of a rule in the model — the block owns it — but a *derived* formula is written relative to the
   * range's top-left cell, so the encoder needs it. Copied rather than threaded as a second parameter because
   * `derivedRuleFormula` takes one object and this is the field it looks for.
   */
  readonly ref?: string;
}

/** A conditional-formatting block: the range, and the rules on it. */
export interface SheetConditionalFormatting {
  readonly ref?: string;
  readonly rules?: readonly SheetCfRule[];
}

/** A rule this writer cannot express, and why. */
export interface CfLoss {
  readonly reason: string;
}

/**
 * The `BrtBeginDXFs` collection, and where each format lands in it.
 *
 * A `dxfId` is an index into this, so it has to be built across every sheet before any of them is written
 * — which is why this is a separate pass rather than something the sheet writer does. Deduplicated by
 * serialised form, because two rules with the same fill share one entry in Excel's own output.
 *
 * **The table the workbook was read with comes first** (`options.source`). Not every holder of a `dxfId` is
 * modelled: a `<colorFilter>` is preserved as XML with the number it was read with, and `fromSource` is how
 * that number is translated. A modelled rule read from the same file holds the very object its `dxfId`
 * resolved to, so it is matched by identity before its serialised form is ever computed.
 *
 * **A format this writer cannot encode gets no index at all** (`options.describe`), rather than an index the
 * styles writer then leaves out. Handing it an index and dropping it later moved every format after it down by
 * one, so each later rule — and every preserved reference — pointed at its neighbour's formatting. A format that
 * encodes only *in part* is kept, and what it lost is returned in `lost`, once per distinct format.
 */
export function collectDxfs(
  blocks: readonly SheetConditionalFormatting[],
  options: DxfCollectorOptions = {}
): DxfTable {
  const table = createDxfTable(options);
  table.add(blocks);
  return table;
}

/** Where the table starts, and how to tell a format this writer can encode — see {@link collectDxfs}. */
export interface DxfCollectorOptions {
  readonly source?: readonly (Partial<Style> | undefined)[];
  readonly describe?: (style: Partial<Style>) => {
    readonly writable: boolean;
    readonly dropped: readonly string[];
  };
}

/** The differential-format table, open to more blocks — see {@link createDxfTable}. */
export interface DxfTable {
  readonly styles: readonly Partial<Style>[];
  readonly indexOf: (style: unknown) => number;
  readonly fromSource: (sourceId: number) => number | undefined;
  readonly lost: readonly string[];
  /** Add the formats of more blocks. A format already in the table keeps its index. */
  readonly add: (blocks: readonly SheetConditionalFormatting[]) => void;
}

/**
 * {@link collectDxfs} for a writer that meets its blocks one sheet at a time.
 *
 * A streamed sheet's part is closed before the next sheet exists, so its rules need their `dxfId`s then; the
 * table grows as each sheet ends and the styles part, written last, carries all of it. Indices are only ever
 * appended, so an id a closed sheet wrote still names the same format.
 */
export function createDxfTable(options: DxfCollectorOptions = {}): DxfTable {
  const { source = [], describe = () => ({ writable: true, dropped: [] }) } = options;
  const styles: Partial<Style>[] = [];
  const lost: string[] = [];
  const indexByKey = new Map<string, number>();
  const indexByIdentity = new Map<unknown, number>();
  const addStyle = (style: Partial<Style>): number => {
    const known = indexByIdentity.get(style);
    if (known !== undefined) {
      return known;
    }
    const key = JSON.stringify(style);
    let index = indexByKey.get(key);
    if (index === undefined) {
      const { writable, dropped } = describe(style);
      index = writable ? styles.push(style) - 1 : NO_DXF;
      if (writable && dropped.length > 0) {
        lost.push(`differential format ${index}: ${dropped.join(", ")}`);
      }
      indexByKey.set(key, index);
    }
    indexByIdentity.set(style, index);
    return index;
  };
  const sourceIndex = source.map(style => (style === undefined ? NO_DXF : addStyle(style)));
  return {
    styles,
    indexOf: style =>
      style === undefined
        ? NO_DXF
        : (indexByIdentity.get(style) ?? indexByKey.get(JSON.stringify(style)) ?? NO_DXF),
    fromSource: sourceId => {
      // With no table read from a file there is nothing to translate from: a stated id can only mean the table
      // being written, which is how the XLSX writer treats the same XML — it writes the number as it is.
      if (source.length === 0) {
        return sourceId < styles.length ? sourceId : undefined;
      }
      const index = sourceIndex[sourceId];
      return index === undefined || index === NO_DXF ? undefined : index;
    },
    lost,
    add: blocks => {
      for (const block of blocks) {
        for (const rule of block.rules ?? []) {
          if (rule.style !== undefined) {
            addStyle(rule.style);
          }
        }
      }
    }
  };
}

/**
 * The workbook's differential-format table as the XLSB package writes it: after the table it was read with, and holding
 * only formats the styles part can encode. One definition for the buffered and the streamed writer, so a `dxfId` a
 * streamed sheet wrote is the one the styles part states.
 */
export function createWorkbookDxfTable(source?: readonly (Partial<Style> | undefined)[]): DxfTable {
  return createDxfTable({
    ...(source === undefined ? {} : { source }),
    // The number-format id does not decide whether a format is writable, only which number it is written under —
    // that is allocated with the styles part, so any id serves for the question asked here.
    describe: style => {
      const { payload, dropped } = encodeDxf(style, () => 0);
      return { writable: payload !== undefined, dropped };
    }
  });
}

/**
 * Hands out workbook-unique rule priorities (`iPri`, which "MUST NOT duplicate" another rule's).
 *
 * The model's priorities are per block and routinely collide, so they cannot simply be copied — but they also cannot
 * simply be replaced, which is what a running counter did: Excel keeps the numbers a file states, and comparing this
 * library's output against an XLSB Excel wrote for the same workbook showed every rule renumbered. So a stated priority
 * is honoured when it is a positive integer nobody has claimed, and the counter fills in for the rest. Workbook-wide,
 * which is stricter than the record requires and costs nothing — and is what a streamed workbook, whose sheets are
 * written one at a time, shares across them.
 */
export function createCfPriorities(): (stated: number | undefined) => number {
  const claimed = new Set<number>();
  let counter = 0;
  return stated => {
    if (stated !== undefined && Number.isInteger(stated) && stated > 0 && !claimed.has(stated)) {
      claimed.add(stated);
      return stated;
    }
    do {
      counter += 1;
    } while (claimed.has(counter));
    claimed.add(counter);
    return counter;
  };
}

/**
 * The records for one sheet's conditional formatting, and what could not be written.
 *
 * `priority` is workbook-unique in the record — `iPri` "MUST NOT duplicate" another rule's — so it is
 * taken from the model where it can be — Excel keeps the numbers a file states, and this handed out a running
 * counter that overwrote them. A stated priority is used when it is a positive integer no other rule on the
 * sheet has claimed; the counter is the fallback for the collisions the model does routinely produce.
 */
export function conditionalFormattingRecords(
  blocks: readonly SheetConditionalFormatting[],
  dxf: { readonly indexOf: (style: unknown) => number },
  context: PtgContext,
  resolvePriority: (stated: number | undefined) => number
): { readonly records: readonly Emitted[]; readonly lost: readonly string[] } {
  const records: Emitted[] = [];
  const lost: string[] = [];
  // **Blocks sharing a range are written as one.** Every `addConditionalFormatting` call arrives here as its
  // own block, so four rules added to `B2:B4` one at a time produced four collections over the same cells.
  // Excel writes one collection per range and puts all four rules in it — that is what it produced when it
  // converted this library's XLSX, which had the four separate. Both are legal; merging is what makes a
  // record-level comparison against Excel readable, and it is what a reader gets from Excel's own files.
  for (const block of mergeByRange(blocks)) {
    const ranges = sqrefRanges(block.ref);
    if (ranges === undefined) {
      lost.push(`conditional formatting on ${block.ref ?? "an unnamed range"}`);
      continue;
    }
    const encoded = (block.rules ?? [])
      .map(rule => {
        // A style `BrtDXF` cannot express gets no `dxfId` (see `collectDxfs`): the rule is still written and
        // still applies, but without its formatting — which is a loss, and was silent.
        if (rule.style !== undefined && dxf.indexOf(rule.style) === NO_DXF) {
          lost.push(`conditional formatting format on ${block.ref ?? "an unnamed range"}`);
        }
        const x14 = dataBarExtensionSettings(rule);
        if (x14.length > 0) {
          lost.push(
            `conditional formatting data bar on ${block.ref ?? "an unnamed range"}: ${x14.join(", ")}`
          );
        }
        // The block's range travels with the rule, because a derived formula is written relative to the
        // range's top-left cell and a single rule does not carry the range itself.
        const payload = encodeRule(
          {
            ...rule,
            ...(rule.ref === undefined && block.ref !== undefined ? { ref: block.ref } : {})
          },
          dxf,
          context,
          resolvePriority
        );
        return payload === undefined || !(payload instanceof Uint8Array)
          ? payload
          : { payload, children: graphicalChildren(rule) };
      })
      .filter((entry): entry is EncodedRule | CfLoss => entry !== undefined);
    const usable = encoded.filter((entry): entry is EncodedRule => !("reason" in entry));
    for (const entry of encoded) {
      if ("reason" in entry) {
        lost.push(entry.reason);
      }
    }
    if (usable.length === 0) {
      continue;
    }
    records.push(
      record(
        "BrtBeginConditionalFormatting",
        concatUint8Arrays([
          // `ccf` counts the rules that follow, so it is the *usable* count — a header promising more
          // rules than follow leaves a reader walking off the end of the collection.
          // `sqrfx` is an `UncheckedSqRfX` (MS-XLSB 2.5.156): a count, then that many ranges. This wrote a count
          // of 1 and the first range whatever the block covered, so `B2:B4 D6:D8` lost its second half.
          new BinaryWriter()
            .writeUint32(usable.length)
            .writeUint32(0)
            .writeInt32(ranges.length)
            .toUint8Array(),
          ...ranges.map(range => encodeRange(range))
        ])
      )
    );
    for (const rule of usable) {
      records.push(record("BrtBeginCFRule", rule.payload));
      if (rule.children !== undefined) {
        records.push(record(rule.children.open[0], rule.children.open[1]));
        for (const [name, payload] of rule.children.body) {
          records.push(record(name, payload));
        }
        records.push(record(rule.children.close));
      }
      records.push(record("BrtEndCFRule"));
    }
    records.push(record("BrtEndConditionalFormatting"));
  }
  return { records, lost };
}

/**
 * The ranges of a space-separated `sqref`, or `undefined` when any of them does not decode.
 *
 * All or nothing: writing the ranges that decode and dropping the rest would apply the rule to part of the cells
 * it was given, with nothing said.
 */
function sqrefRanges(reference: string | undefined): BiffRange[] | undefined {
  const parts = (reference ?? "")
    .trim()
    .split(/\s+/)
    .filter(part => part !== "");
  const ranges = parts.map(part => tryDecodeRange(part));
  return ranges.length === 0 || ranges.some(range => range === undefined)
    ? undefined
    : (ranges as BiffRange[]);
}

/**
 * Blocks with the same range, combined — keeping the order the ranges first appeared in and, within a range,
 * the order the rules were added.
 *
 * Compared as written rather than as decoded: two spellings of one range (`B2:B4` and `$B$2:$B$4`) stay apart,
 * which is conservative. Merging them would need the decoded rectangle as the key, and a range that fails to
 * decode has to keep its own block so that it can be reported as its own loss.
 */
function mergeByRange(
  blocks: readonly SheetConditionalFormatting[]
): readonly SheetConditionalFormatting[] {
  const byRange = new Map<string, SheetConditionalFormatting & { rules: SheetCfRule[] }>();
  const merged: (SheetConditionalFormatting & { rules: SheetCfRule[] })[] = [];
  for (const block of blocks) {
    const key = block.ref ?? "";
    const existing = byRange.get(key);
    if (existing === undefined) {
      const copy = { ...block, rules: [...(block.rules ?? [])] };
      byRange.set(key, copy);
      merged.push(copy);
      continue;
    }
    existing.rules.push(...(block.rules ?? []));
  }
  return merged;
}

/** One `BrtBeginCFRule`, or a loss when the rule needs a child record this writer does not emit. */
function encodeRule(
  rule: SheetCfRule,
  dxf: { readonly indexOf: (style: unknown) => number },
  context: PtgContext,
  resolvePriority: (stated: number | undefined) => number
): Uint8Array | CfLoss | undefined {
  const shape = ruleShape(rule);
  if (shape === undefined) {
    return {
      reason: `conditional formatting rule of type ${JSON.stringify(rule.type ?? "unknown")}`
    };
  }
  const missing = missingRuleFormula(rule);
  if (missing !== undefined) {
    // Written, this was a rule with `cbFmla1 = 0`: a comparison against nothing. The XLSX writer skips the same
    // rules for the same reason, and both report them.
    return { reason: `conditional formatting on ${rule.ref ?? "an unnamed range"}: ${missing}` };
  }
  // `dxfId` MUST be 0xFFFFFFFF for the three graphical types, and those are refused above — so any rule
  // reaching here may carry one.
  const dxfId = dxf.indexOf(rule.style);

  let flags = 0;
  // Bit 0 is reserved. `fStopTrue` at bit 1, `fAbove` at 2, `fBottom` at 3, `fPercent` at 4.
  flags |= rule.stopIfTrue === true ? 1 << 1 : 0;
  // `fAbove` MUST be 1 for the two above-average templates and 0 otherwise — derived from the template
  // rather than from the model, so the two cannot contradict each other.
  // `CF_TEMPLATE_EQUALABOVEAVERAGE` (0x1D) counts too. Written as the literal rather than as arithmetic
  // on another constant, which the first version of this did — `CF_TEMPLATE.expr + 0x1d - CF_TEMPLATE.expr`
  // is 0x1D by a route that says nothing.
  flags |=
    shape.template === CF_TEMPLATE.aboveAverage ||
    shape.template === CF_TEMPLATE_EQUAL_ABOVE_AVERAGE
      ? 1 << 2
      : 0;
  if (shape.type === CF_TYPE.filter) {
    flags |= rule.bottom === true ? 1 << 3 : 0;
    flags |= rule.percent === true ? 1 << 4 : 0;
  }

  // `rule.formulae` when the caller wrote one, and the derived formula otherwise. Several rule types express
  // their condition through other fields — `containsText` through `text`, `timePeriod` through the period — and
  // the file still has to carry the formula those fields stand for.
  const stated = rule.formulae ?? [];
  const derived = stated.length > 0 ? undefined : derivedRuleFormula(rule);
  const formulae =
    shape.formulaCount === 0
      ? []
      : (derived === undefined ? stated : [derived]).slice(0, shape.formulaCount);
  // `strParam` "MUST be greater than 0 and less than 256 characters long" for the text template. A rule without
  // one is not a rule Excel can read — and its formula, derived from the same string, searches for nothing.
  if (
    shape.template === CF_TEMPLATE.containsText &&
    (rule.text === undefined || rule.text.length === 0 || rule.text.length > 255)
  ) {
    return {
      reason: `conditional formatting text rule on ${rule.ref ?? "an unnamed range"}: the search string must be 1 to 255 characters`
    };
  }
  const tokens: EncodedFormula[] = [];
  // **The top-left of the bounding box of the whole sqref**, which every relative reference is an offset from — as
  // Excel measures it: its own `H6>A6` on `G2:G4 H6:H8` is `4c 04 00 00 00 01 c0`, row 4 and column 1 from G2. The
  // first range's corner was used instead, so a rule whose first range was not its top-left one was shifted —
  // `H6>A6` written on `H6:H8 G2:G4` showed as `G2>XFD2`.
  const ranges = sqrefRanges(rule.ref);
  const origin = ranges === undefined ? undefined : boundingTopLeft(ranges);
  for (const formula of formulae) {
    const encoded = encodeCfFormula(formula, context, origin);
    if (encoded === undefined) {
      return { reason: `conditional formatting formula ${JSON.stringify(String(formula))}` };
    }
    tokens.push(encoded);
  }

  return concatUint8Arrays([
    new BinaryWriter()
      .writeUint32(shape.type)
      .writeUint32(shape.template)
      .writeUint32(dxfId)
      .writeInt32(resolvePriority(rule.priority))
      .writeUint32(shape.param)
      .writeUint32(0) // reserved1
      .writeUint32(0) // reserved2
      // A `WORD`, not a `DWORD`. This was four bytes, and the two spare ones shifted `cbfmla1` — Excel read
      // the rule's first formula length as 196608 and answered with `Removed Feature: Conditional
      // formatting`. Section 3.1.2's worked example fixes both the width and the total: five flag bits and
      // eleven reserved make sixteen, and the record there is 0x50 bytes exactly.
      .writeUint16(flags)
      // Three lengths, then the strings and token streams. A zero length means the stream is *absent*
      // rather than empty — writing a zero-length `rgce` after a zero `cbFmla` shifts everything after it.
      //
      // **The non-zero value itself is not read.** The specification says so in as many words: "If this value
      // is nonzero, `rgce1` MUST exist and the value of `cbFmla1` MUST be ignored." So this field is a
      // presence flag with a length-shaped value, and Excel's own choice — the token length plus its two-byte
      // `cce`, `0x1A` where this writes `0x18` — is one of several that satisfy it. Worth stating because the
      // difference shows up in a byte comparison against Excel and is not a defect.
      .writeUint32(tokens[0]?.rgce.length ?? 0)
      .writeUint32(tokens[1]?.rgce.length ?? 0)
      .writeUint32(0)
      .toUint8Array(),
    // `strParam` MUST be NULL unless the template is `containsText`.
    encodeNullableWideString(shape.template === CF_TEMPLATE.containsText ? rule.text : undefined),
    // A `CFParsedFormula` is `cce`, then that many bytes of `rgce`, then **`cb`** — the size of a trailing
    // `rgcb` blob that a conditional formatting formula never has. The final `cb` was missing, so the record
    // ended four bytes early. Its absence very nearly cancelled the two surplus flag bytes above, which is
    // how a 55-byte record passed for one that should be 57: two defects, and any check that only compared
    // a total would have called it a near miss rather than two separate faults.
    ...tokens.map(encoded => encodeParsedFormula(encoded.rgce, encoded.extra))
  ]);
}

/**
 * The Excel 2010 data-bar settings a rule states that `BrtBeginDatabar` cannot carry.
 *
 * Those live in the `x14:dataBar` extension — in XLSB a `BrtCFRuleExt` and its `BrtBeginCFRule14` /
 * `BrtBeginDatabar14` future records — which this writer does not emit. Only settings away from their defaults are
 * named: the XLSX reader gives every extended data bar the full set, and a gradient-filled, border-less bar with
 * matching negative colours is exactly what the base record produces.
 */
function dataBarExtensionSettings(rule: unknown): string[] {
  const bar = rule as {
    type?: string;
    border?: boolean;
    gradient?: boolean;
    negativeBarColorSameAsPositive?: boolean;
    negativeBarBorderColorSameAsPositive?: boolean;
    axisPosition?: string;
    direction?: string;
  };
  if (bar.type !== "dataBar") {
    return [];
  }
  return [
    bar.border === true ? "border" : undefined,
    bar.gradient === false ? "solid fill" : undefined,
    bar.negativeBarColorSameAsPositive === false ? "negative fill colour" : undefined,
    bar.negativeBarBorderColorSameAsPositive === false ? "negative border colour" : undefined,
    bar.axisPosition !== undefined && bar.axisPosition !== "auto"
      ? `axis ${bar.axisPosition}`
      : undefined,
    bar.direction !== undefined && bar.direction !== "leftToRight"
      ? `direction ${bar.direction}`
      : undefined
  ]
    .filter((entry): entry is string => entry !== undefined)
    .map(entry => `${entry} (an x14 extension XLSB does not write)`);
}

/**
 * The `iType`/`iTemplate`/`iParam` triple a model rule maps to, or `undefined` when it has no legal one.
 *
 * The pairing is the specification's own table. Getting it wrong is silent: a `containsText` rule written
 * as `CF_TYPE_CELLIS` is read as a value comparison against the search string.
 */
function ruleShape(
  rule: SheetCfRule
): { type: number; template: number; param: number; formulaCount: number } | undefined {
  // **The text family, in either of the model's spellings** — `{ type: "beginsWith" }` or `{ type: "containsText",
  // operator: "beginsWith" }`. Only the second was recognised, and of its operators only the five this switch had
  // templates for: `notContains`, `beginsWith` and `endsWith` were all written as "contains".
  //
  // **One formula, not none.** These rules are *specified* in terms of a formula the file has to carry: Excel's
  // own `BrtBeginCFRule` for `containsText "AP"` is 82 bytes against this writer's 50 at the time, and the
  // difference is `NOT(ISERROR(SEARCH("AP",A2)))`. `derivedRuleFormula` supplies it — the same function the XML
  // writer uses, so the two containers cannot disagree about what a condition means.
  const kind = textRuleKind(rule);
  if (kind !== undefined) {
    const textOperator = CF_TEXT_OPER[kind];
    return textOperator === undefined
      ? { type: CF_TYPE.expris, template: CF_TEXT_TEMPLATE[kind]!, param: 0, formulaCount: 1 }
      : {
          type: CF_TYPE.expris,
          template: CF_TEMPLATE.containsText,
          param: textOperator,
          formulaCount: 1
        };
  }
  switch (rule.type) {
    case "cellIs":
      return {
        type: CF_TYPE.cellIs,
        template: CF_TEMPLATE.expr,
        param: CF_OPER[rule.operator ?? "equal"] ?? CF_OPER.equal,
        // `between` and `notBetween` take two operands; every other operator takes one.
        formulaCount: rule.operator === "between" || rule.operator === "notBetween" ? 2 : 1
      };
    case "expression":
      return { type: CF_TYPE.expris, template: CF_TEMPLATE.fmla, param: 0, formulaCount: 1 };
    case "timePeriod": {
      const template = CF_TIME_PERIOD[rule.timePeriod ?? "today"];
      // "Last week" is arithmetic over `TODAY()`, and the file is where that arithmetic lives — so one formula.
      return template === undefined
        ? undefined
        : {
            type: CF_TYPE.expris,
            template,
            param: CF_DATE_OPER[rule.timePeriod ?? "today"]!,
            formulaCount: 1
          };
    }
    case "top10":
      return {
        type: CF_TYPE.filter,
        template: CF_TEMPLATE.filter,
        // How many cells, or what percentage — `fPercent` says which.
        param: Math.max(1, Math.trunc(rule.rank ?? 10)),
        formulaCount: 0
      };
    case "aboveAverage": {
      const above = rule.aboveAverage !== false;
      // `equalAverage` has templates of its own, whose `iParam` MUST be 0; otherwise `iParam` is the number of
      // standard deviations, 0 to 3. Both were dropped, so "at or above average" came back as "above".
      if (rule.equalAverage === true) {
        return {
          type: CF_TYPE.expris,
          template: above ? CF_TEMPLATE_EQUAL_ABOVE_AVERAGE : CF_TEMPLATE_EQUAL_BELOW_AVERAGE,
          param: 0,
          formulaCount: 0
        };
      }
      return {
        type: CF_TYPE.expris,
        template: above ? CF_TEMPLATE.aboveAverage : CF_TEMPLATE.belowAverage,
        param: Math.max(0, Math.min(3, Math.trunc(rule.stdDev ?? 0))),
        formulaCount: 0
      };
    }
    case "uniqueValues":
      return {
        type: CF_TYPE.expris,
        template: CF_TEMPLATE.uniqueValues,
        param: 0,
        formulaCount: 0
      };
    case "duplicateValues":
      return {
        type: CF_TYPE.expris,
        template: CF_TEMPLATE.duplicateValues,
        param: 0,
        formulaCount: 0
      };
    // The three graphical rules. Each is followed by its own collection — see `graphicalChildren` — and none
    // of them carries a formula or an operator of its own: the thresholds do that work. `dxfId` is -1 for all
    // three, because the colours are in the collection rather than in a differential format.
    case "colorScale":
      return { type: CF_TYPE.gradient, template: CF_TEMPLATE.gradient, param: 0, formulaCount: 0 };
    case "dataBar":
      return { type: CF_TYPE.databar, template: CF_TEMPLATE.databar, param: 0, formulaCount: 0 };
    case "iconSet":
      return {
        type: CF_TYPE.multistate,
        template: CF_TEMPLATE.multistate,
        param: 0,
        formulaCount: 0
      };
    default:
      return undefined;
  }
}

/**
 * `CFVOtype`, MS-XLSB 2.5.19 — how one threshold of a graphical rule is worked out.
 *
 * `autoMin`/`autoMax` are XLSX-only spellings of "the minimum/maximum of the range, chosen automatically", and
 * they map onto the same two values as `min`/`max`: the binary form has no separate "auto" notion.
 */
const CFVO_TYPE: Readonly<Record<string, number>> = {
  num: 1,
  min: 2,
  autoMin: 2,
  max: 3,
  autoMax: 3,
  percent: 4,
  percentile: 5,
  formula: 7
};

/**
 * `KPISets`, MS-XLSB 2.5.86 — which icon set an icon-set rule draws from.
 *
 * Transcribed from the specification's table rather than derived from the order of the model's union, which is
 * a different order and would silently pick the wrong pictures. `NoIcons` is `KPINIL`, -1.
 *
 * Two of the model's names have no counterpart at all: `3Stars`, `3Triangles` and `5Boxes` were added to XLSX
 * by a later extension and are carried there in an `extLst`, not in this enumeration. They fall back to the
 * three-symbol set rather than being refused, because an icon set with the wrong pictures still shows the
 * thresholds the author set, and refusing the rule loses those too.
 */
const ICON_SET: Readonly<Record<string, number>> = {
  "3Arrows": 0,
  "3ArrowsGray": 1,
  "3Flags": 2,
  "3TrafficLights1": 3,
  "3TrafficLights2": 4,
  "3Signs": 5,
  "3Symbols": 6,
  "3Symbols2": 7,
  "4Arrows": 8,
  "4ArrowsGray": 9,
  "4RedToBlack": 10,
  "4Rating": 11,
  "4TrafficLights": 12,
  "5Arrows": 13,
  "5ArrowsGray": 14,
  "5Rating": 15,
  "5Quarters": 16,
  NoIcons: 0xffffffff,
  // Later additions with no `KPISets` value; see the note above.
  "3Stars": 6,
  "3Triangles": 6,
  "5Boxes": 16
};

/** Excel's data bar defaults: a bar between 10% and 90% of the cell, with the number shown beside it. */
const DATABAR_MIN_LENGTH = 10;
const DATABAR_MAX_LENGTH = 90;

/**
 * `BrtCFVO` — one threshold. Twenty-four bytes with no formula.
 *
 * `fSaveGTE` and `fGTE` are ignored unless the collection belongs to an icon set, where the specification
 * requires `fSaveGTE` to be 1; Excel writes 1 for both. Elsewhere it writes zeros, and so does this.
 */
function encodeCfvo(threshold: Cfvo, forIconSet: boolean): Uint8Array {
  const value =
    typeof threshold.value === "number" ? threshold.value : Number(threshold.value ?? 0);
  return (
    new BinaryWriter()
      .writeUint32(CFVO_TYPE[threshold.type] ?? CFVO_TYPE.num!)
      .writeFloat64(Number.isFinite(value) ? value : 0)
      .writeUint32(forIconSet ? 1 : 0)
      .writeUint32(forIconSet ? 1 : 0)
      // `cbFmla`: no formula. A `formula`-typed threshold whose value is an expression is written as its
      // numeric fallback rather than as tokens, which is a loss and is reported by the caller.
      .writeUint32(0)
      .toUint8Array()
  );
}

/** The records that have to sit inside a graphical rule, or `undefined` when the rule is not graphical. */
function graphicalChildren(rule: SheetCfRule): PivotChildRecords | undefined {
  const thresholds = (rule as { cfvo?: readonly Cfvo[] }).cfvo ?? [];
  const colors = colorList(rule);
  if (rule.type === "colorScale") {
    return {
      open: ["BrtBeginColorScale", undefined],
      // Empty: the number of colours and thresholds is what the collection's contents say.
      body: [
        ...thresholds.map(threshold => ["BrtCFVO", encodeCfvo(threshold, false)] as const),
        ...colors.map(color => ["BrtColor", encodeColor(color)] as const)
      ],
      close: "BrtEndColorScale"
    };
  }
  if (rule.type === "dataBar") {
    const bar = rule as { minLength?: number; maxLength?: number; showValue?: boolean };
    return {
      open: [
        "BrtBeginDatabar",
        new BinaryWriter()
          .writeUint8(clampPercent(bar.minLength, DATABAR_MIN_LENGTH))
          .writeUint8(clampPercent(bar.maxLength, DATABAR_MAX_LENGTH))
          .writeUint8(bar.showValue === false ? 0 : 1)
          .toUint8Array()
      ],
      body: [
        ...thresholds.map(threshold => ["BrtCFVO", encodeCfvo(threshold, false)] as const),
        // One colour, the bar's. A model carrying several keeps the first.
        ["BrtColor", encodeColor(colors[0])] as const
      ],
      close: "BrtEndDatabar"
    };
  }
  if (rule.type === "iconSet") {
    const icons = rule as { iconSet?: string; showValue?: boolean; reverse?: boolean };
    let flags = 0;
    if (icons.showValue === false) {
      flags |= 1 << 1; // fIcon: the icon alone, without the value
    }
    if (icons.reverse === true) {
      flags |= 1 << 2; // fReverse
    }
    return {
      open: [
        "BrtBeginIconSet",
        new BinaryWriter()
          .writeUint32(ICON_SET[icons.iconSet ?? "3TrafficLights1"] ?? ICON_SET["3TrafficLights1"]!)
          .writeUint16(flags)
          .toUint8Array()
      ],
      // An icon set has thresholds and no colours: the pictures come from `iSet`.
      body: thresholds.map(threshold => ["BrtCFVO", encodeCfvo(threshold, true)] as const),
      close: "BrtEndIconSet"
    };
  }
  return undefined;
}

/** A rule's colours, which the model spells either as one or as a list. */
function colorList(rule: SheetCfRule): (Partial<Color> | undefined)[] {
  const color = (rule as { color?: unknown }).color;
  if (Array.isArray(color)) {
    return color as Partial<Color>[];
  }
  return color === undefined ? [] : [color as Partial<Color>];
}

/** A data bar length, which the record holds as a percentage in one byte. */
function clampPercent(value: number | undefined, fallback: number): number {
  if (value === undefined || !Number.isFinite(value)) {
    return fallback;
  }
  return Math.max(0, Math.min(100, Math.round(value)));
}

/** A rule ready to emit: its own payload and, for a graphical rule, the collection that has to follow it. */
interface EncodedRule {
  readonly payload: Uint8Array;
  readonly children: PivotChildRecords | undefined;
}

/** A graphical rule's child records: the opening record, its contents, and the closing name. */
interface PivotChildRecords {
  readonly open: readonly [string, Uint8Array | undefined];
  readonly body: readonly (readonly [string, Uint8Array])[];
  readonly close: string;
}

/** A rule formula as tokens. A bare number is a literal; anything else is parsed. */
/**
 * A rule's formula as tokens, written **relative to the rule's range**.
 *
 * Excel stores a conditional-formatting formula once and shifts it across the range, so its references are
 * `PtgRefN` offsets from the top-left cell rather than positions. This passed the plain cell-formula context,
 * which emits `PtgRef`; the two agree for the first cell of the range and differ for every other one, so a
 * rule on `A2:A4` tested `A2` three times.
 */
function encodeCfFormula(
  formula: string | number,
  context: PtgContext,
  origin: { readonly row: number; readonly column: number } | undefined
): EncodedFormula | undefined {
  try {
    const text = typeof formula === "number" ? String(formula) : formula.replace(/^=/, "");
    return text === ""
      ? undefined
      : encodeFormula(
          text,
          origin === undefined
            ? { ...context, role: "conditional" }
            : { ...context, origin, relativeToOrigin: true, role: "conditional" },
          "conditional format"
        );
  } catch {
    return undefined;
  }
}

/** Colours a `BrtDXF` carries, for the styles table. */
export type CfDxfColor = Partial<Color>;

// =============================================================================
// Reading
// =============================================================================

/**
 * The inverse of {@link CF_OPER}, {@link CF_TIME_PERIOD} and the template table.
 *
 * Built by inverting the write-side tables rather than by listing the values again. That is the whole point:
 * a reader with its own copy of `CFOper` is a second place for "1 means between" to be wrong, and it would be
 * wrong in a way no test comparing the reader against the writer could see — both would agree.
 */
const OPER_BY_VALUE = new Map(Object.entries(CF_OPER).map(([name, value]) => [value, name]));
const TIME_PERIOD_BY_VALUE = new Map(
  Object.entries(CF_TIME_PERIOD).map(([name, value]) => [value, name])
);

/** The text conditions without a search string, back from their templates. */
const TEXT_KIND_BY_TEMPLATE = new Map(
  Object.entries(CF_TEXT_TEMPLATE).map(([kind, template]) => [template, kind as TextRuleKind])
);

/** `CFTextOper` back to the condition, inverted from the table the writer uses. */
const TEXT_KIND_BY_OPER = new Map(
  Object.entries(CF_TEXT_OPER).map(([kind, oper]) => [oper, kind as TextRuleKind])
);

/** One block of conditional formatting, as the model holds it. */
export interface ReadConditionalFormatting {
  ref: string;
  rules: SheetCfRule[];
}

/** `CFVOtype` back to the model's spelling. `autoMin`/`autoMax` share their values with `min`/`max`. */
const CFVO_TYPE_NAME: Readonly<Record<number, string>> = {
  1: "num",
  2: "min",
  3: "max",
  4: "percent",
  5: "percentile",
  7: "formula"
};

/** `KPISets` back to the model's icon-set name, inverted from the table the writer uses. */
const ICON_SET_NAME: ReadonlyMap<number, string> = new Map(
  // The three later additions map onto the same value as `3Symbols`/`5Quarters` and would overwrite the real
  // names, so they are excluded: a file carrying `iSet` 6 came from `3Symbols`.
  Object.entries(ICON_SET)
    .filter(([name]) => !["3Stars", "3Triangles", "5Boxes"].includes(name))
    .map(([name, value]) => [value, name])
);

/** Read a `BrtCFVO` into a threshold. Twenty-four bytes; a trailing formula is not modelled. */
export function readCfvo(payload: Uint8Array): Cfvo | undefined {
  if (payload.length < 24) {
    return undefined;
  }
  const view = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const name = CFVO_TYPE_NAME[view.getUint32(0, true)];
  if (name === undefined) {
    return undefined;
  }
  const value = view.getFloat64(4, true);
  // `min` and `max` take no number — the specification says `numParam` is undefined for them — so reporting one
  // would invent a threshold the file does not state.
  return name === "min" || name === "max"
    ? ({ type: name } as Cfvo)
    : ({ type: name, value } as Cfvo);
}

/**
 * Read a `BrtBeginDatabar`: the two lengths and whether the value shows beside the bar.
 *
 * **A value at its schema default is left out**, as it is from the XLSX Excel writes and therefore from what the
 * XLSX reader returns: `minLength` 10, `maxLength` 90 and `showValue` on are what an absent attribute means. The
 * record has no way to say "absent", so reporting them made every data bar read from XLSB carry three settings the
 * same rule read from XLSX does not — and a model that differs by container for one document.
 */
export function readDatabar(payload: Uint8Array):
  | {
      readonly minLength?: number;
      readonly maxLength?: number;
      readonly showValue?: boolean;
    }
  | undefined {
  if (payload.length < 3) {
    return undefined;
  }
  return {
    ...(payload[0] === DATABAR_MIN_LENGTH ? {} : { minLength: payload[0]! }),
    ...(payload[1] === DATABAR_MAX_LENGTH ? {} : { maxLength: payload[1]! }),
    ...(payload[2] !== 0 ? {} : { showValue: false })
  };
}

/** Read a `BrtBeginIconSet`: which set, and the two display bits. */
export function readIconSet(payload: Uint8Array):
  | {
      readonly iconSet: string;
      readonly showValue?: boolean;
      readonly reverse?: boolean;
    }
  | undefined {
  if (payload.length < 6) {
    return undefined;
  }
  const view = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const flags = view.getUint16(4, true);
  // As for a data bar, the defaults — values shown, not reversed — are what an absent attribute means.
  return {
    iconSet: ICON_SET_NAME.get(view.getUint32(0, true)) ?? "3TrafficLights1",
    ...((flags & (1 << 1)) === 0 ? {} : { showValue: false }),
    ...((flags & (1 << 2)) !== 0 ? { reverse: true } : {})
  };
}

/**
 * Read a `BrtBeginConditionalFormatting`: the rule count, two unused fields, and the range.
 *
 * The count is deliberately *not* trusted to bound the read — the rules are collected from the records that
 * actually follow, up to `BrtEndConditionalFormatting`. A header that disagrees with its collection is a file
 * this reader should survive rather than one it should follow off the end of.
 */
export function readConditionalFormattingBlock(
  payload: Uint8Array,
  part: string
): ReadConditionalFormatting | undefined {
  if (payload.length < 12 + 16) {
    return undefined;
  }
  // `sqrfx` is an `UncheckedSqRfX`: a count and that many ranges, of which this read only the first.
  const reader = new XlsbReader(payload, 8, part);
  const count = reader.readInt32();
  if (count < 1 || count > Math.floor(reader.remaining / 16)) {
    return undefined;
  }
  const ranges: string[] = [];
  for (let index = 0; index < count; index++) {
    ranges.push(rangeReference(readSheetRange(reader, "BrtBeginConditionalFormatting")));
  }
  return { ref: ranges.join(" "), rules: [] };
}

/**
 * Read a `BrtBeginCFRule` back into the model's rule shape.
 *
 * The `iType`/`iTemplate` pair is what decides the shape — the same pairing the writer uses, read in the
 * other direction. `iParam` means a different thing for each: an operator for a comparison, a time period for
 * a date rule, a rank for a top-N.
 *
 * **The formulas are decoded back to text,** through the same `decodePtg`/`printAst` pair the cell and the
 * data-validation readers use, relative to the top-left of `ref` — the origin the writer encoded them from.
 * This returned `formulae: []` once, on the grounds that the reverse of the encoder was missing; it was not,
 * and the narrowing it accepted was not harmless: a `cellIs` rule with no operand compares against nothing,
 * and writing it to XLSX produced a `<cfRule>` with no `<formula>`, which the XLSX writer then failed on when
 * the file was read back and saved. Only the rules whose condition *is* a formula carry one here — a
 * comparison and an expression. The others state theirs in other fields, and their formula is derived from
 * those again on write.
 *
 * @throws XlsbParseError when a formula the rule depends on cannot be decoded. Not returned as a rule without
 *         one, because that rule would look complete and evaluate differently.
 */
export function readCfRule(
  payload: Uint8Array,
  part: string,
  context: PtgContext = {},
  ref?: string
): SheetCfRule | undefined {
  // 28 fixed bytes, a two-byte flag word, three formula lengths and a null `strParam` — the shortest legal
  // rule. This said 44, which is what the four-byte flag field made it look like.
  if (payload.length < 46) {
    return undefined;
  }
  const reader = new XlsbReader(payload, 0, part);
  const type = reader.readUint32();
  const template = reader.readUint32();
  const dxfId = reader.readUint32();
  const priority = reader.readInt32();
  const param = reader.readUint32();
  reader.readUint32();
  reader.readUint32();
  // A `WORD`. The writer had this as a `DWORD` and so did this reader, which is why the two agreed with each
  // other while disagreeing with Excel: a symmetry test cannot see a field both sides get wrong.
  const flags = reader.readUint16();
  const stopIfTrue = (flags & (1 << 1)) !== 0;
  const bottom = (flags & (1 << 3)) !== 0;
  const percent = (flags & (1 << 4)) !== 0;
  const formulaLengths = [reader.readUint32(), reader.readUint32(), reader.readUint32()];
  const text = readNullableWideString(reader, part);
  const formulas = (): string[] => readRuleFormulas(reader, formulaLengths, part, context, ref);
  const base: SheetCfRule = {
    priority,
    ...(stopIfTrue ? { stopIfTrue: true } : {}),
    ...(dxfId === NO_DXF ? {} : { dxfId })
  };
  if (type === CF_TYPE.cellIs) {
    const operator = OPER_BY_VALUE.get(param);
    return {
      ...base,
      type: "cellIs",
      ...(operator === undefined ? {} : { operator }),
      formulae: formulas()
    };
  }
  if (type === CF_TYPE.filter) {
    return {
      ...base,
      type: "top10",
      rank: param,
      ...(bottom ? { bottom: true } : {}),
      ...(percent ? { percent: true } : {})
    };
  }
  if (type === CF_TYPE.gradient || type === CF_TYPE.databar || type === CF_TYPE.multistate) {
    // A graphical rule. Its thresholds and colours live in the collection that follows, which this record does
    // not contain — the caller fills them in from the `BrtCFVO` and `BrtColor` records as they arrive. Returning
    // the shell rather than `undefined` is what makes the round trip whole: the writer emits these now, and
    // dropping them here would lose on read what was written correctly.
    return {
      ...base,
      type:
        type === CF_TYPE.gradient ? "colorScale" : type === CF_TYPE.databar ? "dataBar" : "iconSet",
      cfvo: []
    } as SheetCfRule;
  }
  if (type !== CF_TYPE.expris) {
    // Some other type this reader does not model, dropped rather than returned as a bare `expression`, which
    // is what a caller would then write back.
    return undefined;
  }
  if (template === CF_TEMPLATE.fmla) {
    return { ...base, type: "expression", formulae: formulas() };
  }
  // The text family, in the shape the XLSX reader gives it — see `textRuleReadShape`. For the text template
  // `iParam` is the condition (`CFTextOper`); this ignored it, so every text rule read back as "contains".
  const textKind =
    template === CF_TEMPLATE.containsText
      ? TEXT_KIND_BY_OPER.get(param)
      : TEXT_KIND_BY_TEMPLATE.get(template);
  if (textKind !== undefined) {
    return {
      ...base,
      ...textRuleReadShape(textKind),
      ...(text === undefined || !textRuleHasSearchString(textKind) ? {} : { text }),
      ...derivedFormulae(formulas)
    };
  }
  if (template === CF_TEMPLATE.containsText) {
    // A `CFTextOper` outside the enumeration: no condition this reader can name.
    return undefined;
  }
  const timePeriod = TIME_PERIOD_BY_VALUE.get(template);
  if (timePeriod !== undefined) {
    return { ...base, type: "timePeriod", timePeriod, ...derivedFormulae(formulas) };
  }
  if (template === CF_TEMPLATE.aboveAverage || template === CF_TEMPLATE.belowAverage) {
    return {
      ...base,
      type: "aboveAverage",
      // Only the non-default states — as `readDatabar` explains, a default is what an absent attribute means.
      ...(template === CF_TEMPLATE.aboveAverage ? {} : { aboveAverage: false }),
      ...(param > 0 ? { stdDev: param } : {})
    };
  }
  if (
    template === CF_TEMPLATE_EQUAL_ABOVE_AVERAGE ||
    template === CF_TEMPLATE_EQUAL_BELOW_AVERAGE
  ) {
    // `CF_TEMPLATE_EQUALBELOWAVERAGE` was not recognised at all, so such a rule was dropped on read.
    return {
      ...base,
      type: "aboveAverage",
      ...(template === CF_TEMPLATE_EQUAL_ABOVE_AVERAGE ? {} : { aboveAverage: false }),
      equalAverage: true
    };
  }
  if (template === CF_TEMPLATE.uniqueValues) {
    return { ...base, type: "uniqueValues" };
  }
  if (template === CF_TEMPLATE.duplicateValues) {
    return { ...base, type: "duplicateValues" };
  }
  return undefined;
}

/**
 * The formula of a rule that states its condition in other fields, as the XLSX reader reports it.
 *
 * Excel writes the formula for these rules too — it is what the condition evaluates — and the XLSX reader returns
 * it in `formulae`, so this does as well. Unlike a comparison's operand it is not what the rule *means*: a writer
 * derives it again from `text` or `timePeriod` when it is absent. So a formula that cannot be decoded leaves the
 * rule without one, rather than costing the rule.
 */
function derivedFormulae(formulas: () => string[]): { formulae?: string[] } {
  try {
    const formulae = formulas();
    return formulae.length === 0 ? {} : { formulae };
  } catch {
    return {};
  }
}

/**
 * The `CFParsedFormula`s after a rule's fixed fields, as text: `cce`, that many bytes of `rgce`, then `cb` and that
 * many bytes of `rgcb`. A zero `cbFmla` means the formula is absent, not empty, and nothing is stored for it.
 */
function readRuleFormulas(
  reader: BinaryReader,
  lengths: readonly number[],
  part: string,
  context: PtgContext,
  ref: string | undefined
): string[] {
  // Relative to the top-left of the bounding box of the whole block, as Excel writes it — see `encodeRule`. Excel's
  // own `SUM($A:$A)>C5` on `F1 C5:D9` stores C5 as row 4, column 0 from C1; reading it from F1, the first range's
  // corner, gave `$A:$A>F5`.
  const ranges = sqrefRanges(ref);
  const origin = ranges === undefined ? undefined : boundingTopLeft(ranges);
  const where = `${part} conditional formatting formula`;
  const formulae: string[] = [];
  for (const length of lengths) {
    if (length === 0) {
      continue;
    }
    const tokens = reader.readBytes(reader.readUint32());
    const extra = reader.readBytes(reader.readUint32());
    const decoded = formulaText(
      tokens,
      origin === undefined
        ? { ...context, role: "conditional" }
        : { ...context, origin, role: "conditional" },
      where,
      extra
    );
    if (typeof decoded !== "string") {
      // `CFParsedFormula` may not contain `PtgExp`; reading one as a reference would invent a condition.
      throw new XlsbParseError(where, "PtgExp is not allowed in a conditional formatting formula");
    }
    formulae.push(decoded);
  }
  return formulae;
}
