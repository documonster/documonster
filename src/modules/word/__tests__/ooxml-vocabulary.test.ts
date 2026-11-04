/**
 * Every vocabulary the Word writers enforce is the schema's, held to the ECMA-376 schemas themselves rather than to
 * a second hand-written list. A value the schema allows but a vocabulary here lacks is read into no field and kept
 * as read, so nothing is lost — but the field reads as absent, and that is a bug this test exists to catch.
 */
import { OOXML_ENUMERATIONS } from "@test/fixtures/ooxml-enumerations";
import { WRITTEN } from "@word/core/argument-checks";
import { describe, expect, it } from "vitest";

const {
  wml,
  "shared-commonSimpleTypes": shared,
  "dml-wordprocessingDrawing": wpDrawing,
  "dml-chart": chart
} = OOXML_ENUMERATIONS;

/**
 * Each vocabulary, the type its attribute is declared with, and anything it accepts beyond that type — each with
 * the reason.
 */
const SCHEMA: Record<keyof typeof WRITTEN, [type: readonly string[], extra?: readonly string[]]> = {
  underline: [wml.ST_Underline],
  highlight: [wml.ST_HighlightColor],
  vertAlign: [shared.ST_VerticalAlignRun],
  effect: [wml.ST_TextEffect],
  emphasisMark: [wml.ST_Em],
  combineBrackets: [wml.ST_CombineBrackets],
  // ECMA-376 1st edition (2006) listed `cs` as well, and Word still writes it.
  fontHint: [wml.ST_Hint, ["cs"]],
  shading: [wml.ST_Shd],
  alignment: [wml.ST_Jc],
  tableAlignment: [wml.ST_JcTable],
  lineRule: [wml.ST_LineSpacingRule],
  tabType: [wml.ST_TabJc],
  tabLeader: [wml.ST_TabTlc],
  textAlignment: [wml.ST_TextAlignment],
  textDirection: [wml.ST_TextDirection],
  dropCap: [wml.ST_DropCap],
  frameWrap: [wml.ST_Wrap],
  frameAnchor: [wml.ST_HAnchor],
  frameXAlign: [shared.ST_XAlign],
  frameYAlign: [shared.ST_YAlign],
  tableWidthType: [wml.ST_TblWidth],
  tableLayout: [wml.ST_TblLayoutType],
  rowHeightRule: [wml.ST_HeightRule],
  cellVerticalAlign: [wml.ST_VerticalJc],
  sectionBreak: [wml.ST_SectionMark],
  pageVerticalAlign: [wml.ST_VerticalJc],
  docGrid: [wml.ST_DocGrid],
  lineNumberRestart: [wml.ST_LineNumberRestart],
  pageBorderDisplay: [wml.ST_PageBorderDisplay],
  pageBorderOffset: [wml.ST_PageBorderOffset],
  pageBorderZOrder: [wml.ST_PageBorderZOrder],
  horizontalRelative: [wpDrawing.ST_RelFromH],
  verticalRelative: [wpDrawing.ST_RelFromV],
  horizontalAlign: [wpDrawing.ST_AlignH],
  verticalAlign: [wpDrawing.ST_AlignV],
  wrapSide: [wpDrawing.ST_WrapText],
  legendPosition: [chart.ST_LegendPos],
  dataLabelPosition: [chart.ST_DLblPos]
};

describe("the vocabularies the Word writers enforce are the schema's", () => {
  it.each(Object.entries(SCHEMA))("%s", (name, [type, extra = []]) => {
    const vocabulary = WRITTEN[name as keyof typeof WRITTEN];
    expect([...vocabulary].sort()).toEqual([...type, ...extra].sort());
  });

  it("ST_HAnchor and ST_VAnchor are one vocabulary", () => {
    expect([...wml.ST_VAnchor].sort()).toEqual([...wml.ST_HAnchor].sort());
  });
});
