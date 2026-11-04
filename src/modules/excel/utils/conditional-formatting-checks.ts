/**
 * The checks a block of conditional-formatting rules passes before a sheet takes it — shared by
 * `ConditionalFormatting.add` and the streaming worksheet writer, so a rule the one refuses the other cannot write.
 */
import { missingRuleFormula, textRuleKind } from "@excel/core/conditional-formula";
import { ExcelArgumentError } from "@excel/errors";
import type { ArgumentChecks } from "@excel/utils/arguments";
import { argumentChecks } from "@excel/utils/arguments";
import { checkColor, checkStyle } from "@excel/utils/style-checks";

// The published union, plus the rule types readers return and both writers write (see `ContainsTextRuleType`), and
// the text conditions spelled as their own type, which `textRuleKind` reads.
const RULE_TYPES = [
  "expression",
  "cellIs",
  "top10",
  "aboveAverage",
  "colorScale",
  "iconSet",
  "containsText",
  "timePeriod",
  "dataBar",
  "notContainsText",
  "beginsWith",
  "endsWith",
  "containsBlanks",
  "notContainsBlanks",
  "containsErrors",
  "notContainsErrors",
  "uniqueValues",
  "duplicateValues"
];
// ST_ConditionalFormattingOperator — wider than `CellIsOperators`, and Excel writes every one of them.
const CELL_IS_OPERATORS = [
  "lessThan",
  "lessThanOrEqual",
  "equal",
  "notEqual",
  "greaterThanOrEqual",
  "greaterThan",
  "between",
  "notBetween"
];
// The text family's operators in either spelling: the condition, or XLSX's operator attribute.
const TEXT_OPERATORS = [
  "containsText",
  "containsBlanks",
  "notContainsBlanks",
  "containsErrors",
  "notContainsErrors",
  "notContainsText",
  "notContains",
  "beginsWith",
  "endsWith"
];
const TIME_PERIODS = [
  "lastWeek",
  "thisWeek",
  "nextWeek",
  "yesterday",
  "today",
  "tomorrow",
  "last7Days",
  "lastMonth",
  "thisMonth",
  "nextMonth"
];
const ICON_SETS = [
  "5Arrows",
  "5ArrowsGray",
  "5Boxes",
  "5Quarters",
  "5Rating",
  "4Arrows",
  "4ArrowsGray",
  "4Rating",
  "4RedToBlack",
  "4TrafficLights",
  "NoIcons",
  "3Arrows",
  "3ArrowsGray",
  "3Flags",
  "3Signs",
  "3Stars",
  "3Symbols",
  "3Symbols2",
  "3TrafficLights1",
  "3TrafficLights2",
  "3Triangles"
];
const CFVO_TYPES = ["percentile", "percent", "num", "min", "max", "formula", "autoMin", "autoMax"];

function checkCfvos(
  check: ArgumentChecks,
  field: string,
  cfvo: unknown,
  length: { min: number; max?: number }
): void {
  check.required(field, cfvo);
  if (!check.array(field, cfvo, length)) {
    return;
  }
  cfvo.forEach((entry, i) => {
    check.required(`${field}[${i}]`, entry);
    if (check.record(`${field}[${i}]`, entry)) {
      check.required(`${field}[${i}].type`, entry.type);
      check.oneOf(`${field}[${i}].type`, entry.type, CFVO_TYPES);
      if (typeof entry.value === "number") {
        check.number(`${field}[${i}].value`, entry.value);
      } else {
        check.string(`${field}[${i}].value`, entry.value);
      }
    }
  });
}

function checkRule(
  where: string,
  check: ArgumentChecks,
  field: string,
  rule: unknown,
  ref: string
): void {
  check.required(field, rule);
  if (!check.record(field, rule)) {
    return;
  }
  check.required(`${field}.type`, rule.type);
  check.oneOf(`${field}.type`, rule.type, RULE_TYPES);
  // ST_Priority is a positive int; `undefined` lets the writer number the rule.
  check.number(`${field}.priority`, rule.priority, { min: 1, integer: true });
  if (rule.formulae !== undefined && check.array(`${field}.formulae`, rule.formulae, { max: 2 })) {
    rule.formulae.forEach((formula, i) => {
      if (typeof formula === "number") {
        check.number(`${field}.formulae[${i}]`, formula);
      } else {
        check.string(`${field}.formulae[${i}]`, formula);
      }
    });
  }
  switch (rule.type) {
    case "cellIs":
      // The comparison is the condition: a rule without one says nothing Excel can evaluate.
      check.required(`${field}.operator`, rule.operator);
      check.oneOf(`${field}.operator`, rule.operator, CELL_IS_OPERATORS);
      break;
    case "top10":
      // Excel's own bounds: 1–100 for a percentage, 1–1000 for a count.
      check.number(`${field}.rank`, rule.rank, {
        min: 1,
        max: rule.percent === true ? 100 : 1000,
        integer: true
      });
      check.boolean(`${field}.percent`, rule.percent);
      check.boolean(`${field}.bottom`, rule.bottom);
      break;
    case "aboveAverage":
      check.boolean(`${field}.aboveAverage`, rule.aboveAverage);
      check.boolean(`${field}.equalAverage`, rule.equalAverage);
      check.number(`${field}.stdDev`, rule.stdDev, { min: 0, integer: true });
      break;
    case "colorScale":
      checkCfvos(check, `${field}.cfvo`, rule.cfvo, { min: 2 });
      check.required(`${field}.color`, rule.color);
      if (
        check.array(`${field}.color`, rule.color, {
          min: (rule.cfvo as unknown[]).length,
          max: (rule.cfvo as unknown[]).length
        })
      ) {
        // A colour that is not one was written as `auto`, so the scale showed colours nobody chose.
        rule.color.forEach((color, i) => {
          check.required(`${field}.color[${i}]`, color);
          checkColor(where, `${field}.color[${i}]`, color);
        });
      }
      break;
    case "iconSet":
      check.oneOf(`${field}.iconSet`, rule.iconSet, ICON_SETS);
      checkCfvos(check, `${field}.cfvo`, rule.cfvo, { min: 2 });
      for (const flag of ["showValue", "reverse", "custom"]) {
        check.boolean(`${field}.${flag}`, rule[flag]);
      }
      break;
    case "dataBar":
      if (rule.cfvo !== undefined) {
        checkCfvos(check, `${field}.cfvo`, rule.cfvo, { min: 2, max: 2 });
      }
      check.number(`${field}.minLength`, rule.minLength, { min: 0, max: 100, integer: true });
      check.number(`${field}.maxLength`, rule.maxLength, { min: 0, max: 100, integer: true });
      check.oneOf(`${field}.axisPosition`, rule.axisPosition, ["auto", "middle", "none"]);
      check.oneOf(`${field}.direction`, rule.direction, ["context", "leftToRight", "rightToLeft"]);
      for (const color of [
        "color",
        "negativeFillColor",
        "borderColor",
        "negativeBorderColor",
        "axisColor"
      ]) {
        const value = rule[color];
        // The binary writer also takes a list for the bar colour, as a colour scale spells its colours.
        if (Array.isArray(value)) {
          value.forEach((entry, i) => checkColor(where, `${field}.${color}[${i}]`, entry));
        } else {
          checkColor(where, `${field}.${color}`, value);
        }
      }
      for (const flag of [
        "gradient",
        "showValue",
        "border",
        "negativeBarColorSameAsPositive",
        "negativeBarBorderColorSameAsPositive"
      ]) {
        check.boolean(`${field}.${flag}`, rule[flag]);
      }
      break;
    case "timePeriod":
      // The period is the condition: the writer derives the formula from it, and the file states it.
      check.required(`${field}.timePeriod`, rule.timePeriod);
      check.oneOf(`${field}.timePeriod`, rule.timePeriod, TIME_PERIODS);
      break;
    case "containsText":
    case "notContainsText":
    case "beginsWith":
    case "endsWith":
      check.oneOf(`${field}.operator`, rule.operator, TEXT_OPERATORS);
      check.string(`${field}.text`, rule.text);
      break;
  }
  const kind = textRuleKind(rule);
  if (
    kind !== undefined &&
    (kind === "containsText" ||
      kind === "notContainsText" ||
      kind === "beginsWith" ||
      kind === "endsWith")
  ) {
    // The search string is the condition, and Excel reads it from the `text` attribute: a rule without one is
    // discarded when the file is opened, whatever formula it carries.
    check.required(`${field}.text`, rule.text);
  }
  const missing = missingRuleFormula({ ...rule, ref } as Parameters<typeof missingRuleFormula>[0]);
  if (missing !== undefined) {
    throw new ExcelArgumentError(
      `${field}.formulae`,
      `${where}: ${field}.formulae is missing — the ${missing}.`
    );
  }
  checkStyle(where, rule.style, `${field}.style`);
}

/**
 * Check a block of conditional-formatting rules: `ref` is one or more ranges separated by spaces, and each rule is
 * checked against its type — an unknown type, a comparison without its operator or operands, a colour scale without
 * its thresholds or with colours that are not colours, a priority below 1 — and its `style` as a cell style is.
 */
export function checkConditionalFormatting(where: string, cf: unknown, argument = "cf"): void {
  const check = argumentChecks(where);
  check.required(argument, cf);
  if (check.record(argument, cf)) {
    check.required(`${argument}.ref`, cf.ref);
    check.reference(`${argument}.ref`, cf.ref, { range: true, lines: true, list: true });
    check.required(`${argument}.rules`, cf.rules);
    if (check.array(`${argument}.rules`, cf.rules, { min: 1 })) {
      cf.rules.forEach((rule, i) =>
        checkRule(where, check, `${argument}.rules[${i}]`, rule, cf.ref as string)
      );
    }
  }
}
