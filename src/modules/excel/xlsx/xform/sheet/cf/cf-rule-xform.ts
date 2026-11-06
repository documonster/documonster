// The derived formulas live in `core/conditional-formula.ts` so the XLSB writer can reach them too — it
// emitted these rules with no formula at all while this one had them, which is a rule that never fires.
import {
  textRuleFormula as getTextFormula,
  textRuleHasSearchString,
  textRuleKind,
  timePeriodRuleFormula as getTimePeriodFormula,
  type TextRuleKind
} from "@excel/core/conditional-formula";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { CompositeXform } from "@excel/xlsx/xform/composite-xform";
import { ColorScaleXform } from "@excel/xlsx/xform/sheet/cf/color-scale-xform";
import { DatabarXform } from "@excel/xlsx/xform/sheet/cf/databar-xform";
import { ExtLstRefXform } from "@excel/xlsx/xform/sheet/cf/ext-lst-ref-xform";
import { FormulaXform } from "@excel/xlsx/xform/sheet/cf/formula-xform";
import { IconSetXform } from "@excel/xlsx/xform/sheet/cf/icon-set-xform";
import type { ParseOpenTag } from "@xml/types";

const extIcons = {
  "3Triangles": true,
  "3Stars": true,
  "5Boxes": true
};

function opType(attributes) {
  const { type, operator } = attributes;
  switch (type) {
    case "containsText":
    case "containsBlanks":
    case "notContainsBlanks":
    case "containsErrors":
    case "notContainsErrors":
      return {
        type: "containsText",
        operator: type
      };

    default:
      // `notContainsText`, `beginsWith` and `endsWith` keep the XLSX spelling — type and operator attribute —
      // which is the shape readers have always returned for them (see `ContainsTextRuleType`).
      return { type, operator };
  }
}

/** CT_CfRule's `operator` attribute for a text condition: only the four with a search string have one. */
const TEXT_OPERATOR_ATTRIBUTE: Partial<Record<TextRuleKind, string>> = {
  containsText: "containsText",
  notContainsText: "notContains",
  beginsWith: "beginsWith",
  endsWith: "endsWith"
};

class CfRuleXform extends CompositeXform {
  databarXform: DatabarXform;
  extLstRefXform: ExtLstRefXform;
  formulaXform: FormulaXform;
  colorScaleXform: ColorScaleXform;
  iconSetXform: IconSetXform;

  constructor() {
    super();

    this.map = {
      dataBar: (this.databarXform = new DatabarXform()),
      extLst: (this.extLstRefXform = new ExtLstRefXform()),
      formula: (this.formulaXform = new FormulaXform()),
      colorScale: (this.colorScaleXform = new ColorScaleXform()),
      iconSet: (this.iconSetXform = new IconSetXform())
    };
  }

  get tag() {
    return "cfRule";
  }

  static isPrimitive(rule) {
    // is this rule primitive?
    if (rule.type === "iconSet") {
      if (rule.custom || extIcons[rule.iconSet]) {
        return false;
      }
    }
    return true;
  }

  render(xmlStream, model) {
    // The text family in either spelling — `{ type: "beginsWith" }` or `{ type: "containsText", operator:
    // "beginsWith" }`. Only the second was rendered; the first wrote nothing at all.
    if (textRuleKind(model) !== undefined) {
      this.renderText(xmlStream, model);
      return;
    }
    switch (model.type) {
      case "expression":
        this.renderExpression(xmlStream, model);
        break;
      case "cellIs":
        this.renderCellIs(xmlStream, model);
        break;
      case "top10":
        this.renderTop10(xmlStream, model);
        break;
      case "aboveAverage":
        this.renderAboveAverage(xmlStream, model);
        break;
      case "dataBar":
        this.renderDataBar(xmlStream, model);
        break;
      case "colorScale":
        this.renderColorScale(xmlStream, model);
        break;
      case "iconSet":
        this.renderIconSet(xmlStream, model);
        break;
      case "uniqueValues":
      case "duplicateValues":
        // No formula and no attributes beyond the formatting: the condition is the type. Neither was rendered, so
        // both vanished from every XLSX written here.
        xmlStream.leafNode(this.tag, {
          type: model.type,
          dxfId: model.dxfId,
          priority: model.priority
        });
        break;
      case "timePeriod":
        this.renderTimePeriod(xmlStream, model);
        break;
    }
  }

  renderExpression(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      type: "expression",
      dxfId: model.dxfId,
      priority: model.priority
    });

    this.formulaXform.render(xmlStream, model.formulae?.[0]);

    xmlStream.closeNode();
  }

  renderCellIs(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      type: "cellIs",
      dxfId: model.dxfId,
      priority: model.priority,
      operator: model.operator
    });

    (model.formulae ?? []).forEach(formula => {
      this.formulaXform.render(xmlStream, formula);
    });

    xmlStream.closeNode();
  }

  renderTop10(xmlStream, model) {
    xmlStream.leafNode(this.tag, {
      type: "top10",
      dxfId: model.dxfId,
      priority: model.priority,
      percent: BaseXform.toBoolAttribute(model.percent, false),
      bottom: BaseXform.toBoolAttribute(model.bottom, false),
      rank: BaseXform.toIntValue(model.rank, 10)
    });
  }

  renderAboveAverage(xmlStream, model) {
    xmlStream.leafNode(this.tag, {
      type: "aboveAverage",
      dxfId: model.dxfId,
      priority: model.priority,
      // Written whenever the model states it, `true` included: CT_CfRule defaults the attribute to true, but a
      // model that says so explicitly reads back the same way rather than as `undefined`.
      aboveAverage: BaseXform.toBoolAttribute(model.aboveAverage),
      equalAverage: BaseXform.toBoolAttribute(model.equalAverage, false),
      stdDev: model.stdDev
    });
  }

  renderDataBar(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      type: "dataBar",
      priority: model.priority
    });

    this.databarXform.render(xmlStream, model);
    this.extLstRefXform.render(xmlStream, model);

    xmlStream.closeNode();
  }

  renderColorScale(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      type: "colorScale",
      priority: model.priority
    });

    this.colorScaleXform.render(xmlStream, model);

    xmlStream.closeNode();
  }

  renderIconSet(xmlStream, model) {
    // iconset is all primitive or all extLst
    if (!CfRuleXform.isPrimitive(model)) {
      return;
    }

    xmlStream.openNode(this.tag, {
      type: "iconSet",
      priority: model.priority
    });

    this.iconSetXform.render(xmlStream, model);

    xmlStream.closeNode();
  }

  renderText(xmlStream, model) {
    // Per ECMA-376 (CT_CfRule / ST_ConditionalFormattingOperator), only the four conditions with a search string
    // carry an `operator` attribute, and each has its own value — `notContainsText` is `notContains`. The blank and
    // error tests never take one: `type="containsBlanks" operator="containsBlanks"` names an operator that is not
    // in the enumeration, and strict readers reject the whole worksheet for it.
    const kind = textRuleKind(model)!;
    xmlStream.openNode(this.tag, {
      type: kind,
      dxfId: model.dxfId,
      priority: model.priority,
      operator: TEXT_OPERATOR_ATTRIBUTE[kind],
      // **The search string itself.** CT_CfRule carries it on `text`; the `<formula>` written below is the
      // *evaluation* of the rule, not its statement, and Excel reads the attribute rather than the formula.
      // Omitting it made Excel discard the rule outright — converting a workbook written here to XLSB kept
      // four of five rules and dropped this one — so the rule looked fine in this library and did not exist
      // once Excel had seen it.
      text: textRuleHasSearchString(kind) ? model.text : undefined
    });

    const formula = getTextFormula(model);
    if (formula) {
      this.formulaXform.render(xmlStream, formula);
    }

    xmlStream.closeNode();
  }

  renderTimePeriod(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      type: "timePeriod",
      dxfId: model.dxfId,
      priority: model.priority,
      timePeriod: model.timePeriod
    });

    const formula = getTimePeriodFormula(model);
    if (formula) {
      this.formulaXform.render(xmlStream, formula);
    }

    xmlStream.closeNode();
  }

  createNewModel({ attributes }: ParseOpenTag) {
    return {
      ...opType(attributes),
      dxfId: BaseXform.toIntValue(attributes.dxfId),
      priority: BaseXform.toIntValue(attributes.priority),
      timePeriod: attributes.timePeriod,
      percent: BaseXform.toBoolValue(attributes.percent),
      bottom: BaseXform.toBoolValue(attributes.bottom),
      rank: BaseXform.toIntValue(attributes.rank),
      aboveAverage: BaseXform.toBoolValue(attributes.aboveAverage),
      // Three attributes this did not read, so a text rule came back without its search string and an average
      // rule without `equalAverage` or `stdDev`. Added only when the markup has them, so a rule without them has
      // exactly the fields it always had.
      ...(attributes.text === undefined ? {} : { text: attributes.text }),
      ...(attributes.equalAverage === undefined
        ? {}
        : { equalAverage: BaseXform.toBoolValue(attributes.equalAverage) }),
      ...(attributes.stdDev === undefined
        ? {}
        : { stdDev: BaseXform.toIntValue(attributes.stdDev) })
    };
  }

  onParserClose(name, parser) {
    switch (name) {
      case "dataBar":
      case "extLst":
      case "colorScale":
      case "iconSet":
        // merge parser model with ours
        Object.assign(this.model, parser.model);
        break;

      case "formula":
        // except - formula is a string and appends to formulae
        this.model.formulae = this.model.formulae ?? [];
        this.model.formulae.push(parser.model);
        break;
    }
  }
}

export { CfRuleXform };
