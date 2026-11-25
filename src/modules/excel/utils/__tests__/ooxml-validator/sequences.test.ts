/**
 * Child order below each part's root (`check-sequences.ts`): every part kind is wired, a reorder of any ranked
 * container is reported, and what the check must not report — extension elements, `x14` lookalikes in `extLst`,
 * `mc:AlternateContent` — is not.
 */

import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { SEQUENCE_PARTS, type SequenceNode } from "@excel/utils/ooxml-validator/check-sequences";
import { OOXML_SEQUENCES } from "@excel/utils/ooxml-validator/ooxml-sequences";
import { CHART_ORDER_RULES, CHART_SERIES_ORDER_RULES } from "@xml/ooxml-chart-order";
import { describe, expect, it } from "vitest";

import { baseParts, buildPackage, type PackageParts } from "./fixtures";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const CT = "application/vnd.openxmlformats-officedocument";

/** Add a part with a content-type override, leaving the base package otherwise as it is. */
function withPart(path: string, contentType: string, xml: string): PackageParts {
  const parts = baseParts();
  parts["[Content_Types].xml"] = (parts["[Content_Types].xml"] as string).replace(
    "</Types>",
    `<Override PartName="/${path}" ContentType="${contentType}"/></Types>`
  );
  parts[path] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`;
  return parts;
}

async function orderProblems(parts: PackageParts): Promise<string[]> {
  const report = await validateXlsxBuffer(buildPackage(parts), { maxProblems: 100 });
  return report.problems.filter(p => p.kind === "part-child-out-of-order").map(p => p.message);
}

const sheet = (body: string) =>
  withPart(
    "xl/worksheets/sheet1.xml",
    `${CT}.spreadsheetml.worksheet+xml`,
    `<worksheet xmlns="${MAIN}">${body}</worksheet>`
  );
const styles = (body: string) =>
  withPart(
    "xl/styles.xml",
    `${CT}.spreadsheetml.styles+xml`,
    `<styleSheet xmlns="${MAIN}">${body}</styleSheet>`
  );
const drawing = (body: string) =>
  withPart(
    "xl/drawings/drawing1.xml",
    `${CT}.drawing+xml`,
    `<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}">${body}</xdr:wsDr>`
  );

describe("ooxml-validator / element order below the root", () => {
  it("passes the base package", async () => {
    expect(await orderProblems(baseParts())).toEqual([]);
  });

  /** Each case is one part written correctly, then with two children of one container swapped. */
  const cases: { name: string; good: PackageParts; bad: PackageParts; message: string }[] = [
    {
      name: "worksheet <sheetPr> (issue #269)",
      good: sheet(`<sheetPr><outlinePr/><pageSetUpPr fitToPage="1"/></sheetPr><sheetData/>`),
      bad: sheet(`<sheetPr><pageSetUpPr fitToPage="1"/><outlinePr/></sheetPr><sheetData/>`),
      message: "<worksheet/sheetPr> child <outlinePr> appears after <pageSetUpPr>"
    },
    {
      name: "worksheet cell",
      good: sheet(`<sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row></sheetData>`),
      bad: sheet(`<sheetData><row r="1"><c r="A1"><v>2</v><f>1+1</f></c></row></sheetData>`),
      message: "<worksheet/sheetData/row/c> child <f> appears after <v>"
    },
    {
      name: "worksheet inline-string rich run",
      good: sheet(
        `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>x</t></r></is></c></row></sheetData>`
      ),
      bad: sheet(
        `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><t>x</t><rPr><b/></rPr></r></is></c></row></sheetData>`
      ),
      message: "<worksheet/sheetData/row/c/is/r> child <rPr> appears after <t>"
    },
    {
      name: "worksheet conditional-format data bar",
      good: sheet(
        `<sheetData/><conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>`
      ),
      bad: sheet(
        `<sheetData/><conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><color rgb="FF638EC6"/><cfvo type="max"/></dataBar></cfRule></conditionalFormatting>`
      ),
      message: "<worksheet/conditionalFormatting/cfRule/dataBar> child <cfvo> appears after <color>"
    },
    {
      name: "worksheet sheetView",
      good: sheet(
        `<sheetViews><sheetView workbookViewId="0"><pane ySplit="1"/><selection pane="bottomLeft"/></sheetView></sheetViews><sheetData/>`
      ),
      bad: sheet(
        `<sheetViews><sheetView workbookViewId="0"><selection pane="bottomLeft"/><pane ySplit="1"/></sheetView></sheetViews><sheetData/>`
      ),
      message: "<worksheet/sheetViews/sheetView> child <pane> appears after <selection>"
    },
    {
      name: "worksheet headerFooter",
      good: sheet(
        `<sheetData/><headerFooter><oddHeader>a</oddHeader><oddFooter>b</oddFooter></headerFooter>`
      ),
      bad: sheet(
        `<sheetData/><headerFooter><oddFooter>b</oddFooter><oddHeader>a</oddHeader></headerFooter>`
      ),
      message: "<worksheet/headerFooter> child <oddHeader> appears after <oddFooter>"
    },
    {
      name: "styles <xf>",
      good: styles(
        `<cellXfs count="1"><xf><alignment horizontal="center"/><protection locked="0"/></xf></cellXfs>`
      ),
      bad: styles(
        `<cellXfs count="1"><xf><protection locked="0"/><alignment horizontal="center"/></xf></cellXfs>`
      ),
      message: "<styleSheet/cellXfs/xf> child <alignment> appears after <protection>"
    },
    {
      name: "styles <dxf>",
      good: styles(
        `<dxfs count="1"><dxf><font><b/></font><numFmt numFmtId="164" formatCode="0"/><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>`
      ),
      bad: styles(
        `<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill><font><b/></font></dxf></dxfs>`
      ),
      message: "<styleSheet/dxfs/dxf> child <font> appears after <fill>"
    },
    {
      name: "styles <border>",
      good: styles(
        `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>`
      ),
      bad: styles(`<borders count="1"><border><top/><left/></border></borders>`),
      message: "<styleSheet/borders/border> child <left> appears after <top>"
    },
    {
      name: "styles root",
      good: styles(`<fonts count="0"/><fills count="0"/>`),
      bad: styles(`<fills count="0"/><fonts count="0"/>`),
      message: "<styleSheet> child <fonts> appears after <fills>"
    },
    {
      name: "shared strings rich run",
      good: withPart(
        "xl/sharedStrings.xml",
        `${CT}.spreadsheetml.sharedStrings+xml`,
        `<sst xmlns="${MAIN}"><si><r><rPr><b/></rPr><t>x</t></r></si></sst>`
      ),
      bad: withPart(
        "xl/sharedStrings.xml",
        `${CT}.spreadsheetml.sharedStrings+xml`,
        `<sst xmlns="${MAIN}"><si><r><t>x</t><rPr><b/></rPr></r></si></sst>`
      ),
      message: "<sst/si/r> child <rPr> appears after <t>"
    },
    {
      name: "comments",
      good: withPart(
        "xl/comments1.xml",
        `${CT}.spreadsheetml.comments+xml`,
        `<comments xmlns="${MAIN}"><authors/><commentList/></comments>`
      ),
      bad: withPart(
        "xl/comments1.xml",
        `${CT}.spreadsheetml.comments+xml`,
        `<comments xmlns="${MAIN}"><commentList/><authors/></comments>`
      ),
      message: "<comments> child <authors> appears after <commentList>"
    },
    {
      name: "table",
      good: withPart(
        "xl/tables/table1.xml",
        `${CT}.spreadsheetml.table+xml`,
        `<table xmlns="${MAIN}" id="1" ref="A1:A2"><autoFilter ref="A1:A2"/><tableColumns count="1"><tableColumn id="1" name="a"/></tableColumns><tableStyleInfo name="TableStyleMedium2"/></table>`
      ),
      bad: withPart(
        "xl/tables/table1.xml",
        `${CT}.spreadsheetml.table+xml`,
        `<table xmlns="${MAIN}" id="1" ref="A1:A2"><tableColumns count="1"><tableColumn id="1" name="a"/></tableColumns><autoFilter ref="A1:A2"/></table>`
      ),
      message: "<table> child <autoFilter> appears after <tableColumns>"
    },
    {
      name: "external link book",
      good: withPart(
        "xl/externalLinks/externalLink1.xml",
        `${CT}.spreadsheetml.externalLink+xml`,
        `<externalLink xmlns="${MAIN}"><externalBook><sheetNames/><sheetDataSet/></externalBook></externalLink>`
      ),
      bad: withPart(
        "xl/externalLinks/externalLink1.xml",
        `${CT}.spreadsheetml.externalLink+xml`,
        `<externalLink xmlns="${MAIN}"><externalBook><sheetDataSet/><sheetNames/></externalBook></externalLink>`
      ),
      message: "<externalLink/externalBook> child <sheetNames> appears after <sheetDataSet>"
    },
    {
      name: "drawing one-cell anchor",
      good: drawing(
        `<xdr:oneCellAnchor><xdr:from/><xdr:ext cx="1" cy="1"/><xdr:pic/><xdr:clientData/></xdr:oneCellAnchor>`
      ),
      bad: drawing(
        `<xdr:oneCellAnchor><xdr:from/><xdr:pic/><xdr:clientData/><xdr:ext cx="1" cy="1"/></xdr:oneCellAnchor>`
      ),
      message: "<wsDr/oneCellAnchor> child <ext> appears after <clientData>"
    },
    {
      name: "drawing marker",
      good: drawing(
        `<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to/><xdr:clientData/></xdr:twoCellAnchor>`
      ),
      bad: drawing(
        `<xdr:twoCellAnchor><xdr:from><xdr:row>0</xdr:row><xdr:col>0</xdr:col></xdr:from><xdr:to/><xdr:clientData/></xdr:twoCellAnchor>`
      ),
      message: "<wsDr/twoCellAnchor/from> child <col> appears after <row>"
    },
    {
      name: "drawing picture shape properties",
      good: drawing(
        `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="1" cy="1"/><xdr:pic><xdr:nvPicPr/><xdr:blipFill/><xdr:spPr><a:xfrm/><a:prstGeom prst="rect"/><a:noFill/><a:ln/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:absoluteAnchor>`
      ),
      bad: drawing(
        `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="1" cy="1"/><xdr:pic><xdr:nvPicPr/><xdr:blipFill/><xdr:spPr><a:prstGeom prst="rect"/><a:xfrm/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:absoluteAnchor>`
      ),
      message: "<wsDr/absoluteAnchor/pic/spPr> child <xfrm> appears after <prstGeom>"
    }
  ];

  for (const { name, good, bad, message } of cases) {
    it(`${name}: accepts the schema order and reports a swap`, async () => {
      expect(await orderProblems(good)).toEqual([]);
      const problems = await orderProblems(bad);
      expect(problems).toHaveLength(1);
      expect(problems[0]).toContain(message);
    });
  }

  it("ranks the members of an xsd:choice as one position", async () => {
    // CT_ShapeProperties: geometry (custGeom|prstGeom) before fill (noFill|solidFill|…), in either member.
    const problems = await orderProblems(
      drawing(
        `<xdr:twoCellAnchor><xdr:from/><xdr:to/><xdr:sp><xdr:spPr><a:prstGeom prst="rect"/><a:solidFill/><a:ln/></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`
      )
    );
    expect(problems).toEqual([]);
  });

  it("unwraps mc:AlternateContent to its mc:Choice", async () => {
    const wrap = (shape: string) =>
      drawing(
        `<xdr:twoCellAnchor><xdr:from/><xdr:to/><mc:AlternateContent xmlns:mc="${MC}"><mc:Choice Requires="a14">${shape}</mc:Choice><mc:Fallback/></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`
      );
    expect(await orderProblems(wrap(`<xdr:sp/>`))).toEqual([]);
    // The wrapped shape stands in the anchor's content slot, so a `from` inside the Choice is out of place.
    expect(await orderProblems(wrap(`<xdr:sp/><xdr:from/>`))).toHaveLength(1);
  });

  it("skips elements the content model does not name", async () => {
    expect(
      await orderProblems(
        sheet(
          `<sheetPr><outlinePr/><x14ac:unknown xmlns:x14ac="urn:x"/><tabColor rgb="FFFF0000"/></sheetPr><sheetData/>`
        )
      )
    ).toHaveLength(1);
    expect(
      await orderProblems(
        sheet(
          `<sheetPr><tabColor rgb="FFFF0000"/><x14ac:unknown xmlns:x14ac="urn:x"/><outlinePr/></sheetPr><sheetData/>`
        )
      )
    ).toEqual([]);
  });

  it("does not hold an x14:cfRule in extLst to CT_CfRule", async () => {
    expect(
      await orderProblems(
        sheet(
          `<sheetData/><extLst><ext uri="x"><x14:conditionalFormattings xmlns:x14="urn:x14"><x14:conditionalFormatting><x14:cfRule type="dataBar"><x14:dataBar/><formula>1</formula></x14:cfRule></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`
        )
      )
    ).toEqual([]);
  });

  it("can be turned off", async () => {
    const report = await validateXlsxBuffer(
      buildPackage(sheet(`<sheetPr><pageSetUpPr/><outlinePr/></sheetPr><sheetData/>`)),
      { checkElementOrder: false }
    );
    expect(report.problems.some(p => p.kind === "part-child-out-of-order")).toBe(false);
  });

  it("uses every generated content model", () => {
    // The roots of a worksheet and a workbook are ranked by their own checks; the chart part by check-chart.
    const used = new Set<string>([
      "x:CT_Worksheet",
      "x:CT_Workbook",
      ...CHART_ORDER_RULES.map(rule => rule.type),
      ...CHART_SERIES_ORDER_RULES.map(rule => rule.type)
    ]);
    const visit = (node: SequenceNode) => {
      if (node.type) {
        used.add(node.type);
      }
      Object.values(node.children ?? {}).forEach(visit);
    };
    SEQUENCE_PARTS.forEach(visit);
    expect(Object.keys(OOXML_SEQUENCES).filter(type => !used.has(type))).toEqual([]);
  });
});

describe("ooxml-validator / x14 extension order", () => {
  const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
  const XM = "http://schemas.microsoft.com/office/excel/2006/main";
  const extension = (inner: string) =>
    sheet(
      `<sheetData/><extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14}" xmlns:xm="${XM}">${inner}</ext></extLst>`
    );

  it("accepts a sparkline group in [MS-XLSX] order and reports one out of it", async () => {
    const group = (body: string) =>
      extension(
        `<x14:sparklineGroups><x14:sparklineGroup>${body}</x14:sparklineGroup></x14:sparklineGroups>`
      );
    expect(
      await orderProblems(
        group(`<x14:colorSeries rgb="FF000000"/><xm:f>S!A1</xm:f><x14:sparklines/>`)
      )
    ).toEqual([]);
    expect(
      await orderProblems(group(`<x14:sparklines/><x14:colorSeries rgb="FF000000"/>`))
    ).toEqual([expect.stringContaining("child <colorSeries> appears after <sparklines>")]);
  });

  it("reports a data bar's negative fill after its negative border", async () => {
    const bar = (body: string) =>
      extension(
        `<x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar"><x14:dataBar>${body}</x14:dataBar></x14:cfRule><xm:sqref>A1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings>`
      );
    expect(
      await orderProblems(
        bar(
          `<x14:cfvo type="min"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFF0000"/>`
        )
      )
    ).toEqual([]);
    expect(
      await orderProblems(
        bar(
          `<x14:cfvo type="min"/><x14:negativeBorderColor rgb="FFFF0000"/><x14:negativeFillColor rgb="FFFF0000"/>`
        )
      )
    ).toEqual([
      expect.stringContaining("child <negativeFillColor> appears after <negativeBorderColor>")
    ]);
  });
});

describe("ooxml-validator / root content models", () => {
  // `check-worksheet.ts` and `check-workbook.ts` rank a root's children by their index in the flattened table, which
  // is the schema's order only while every position holds one element.
  it.each(["x:CT_Worksheet", "x:CT_Workbook"] as const)("%s has one element per position", type => {
    expect(OOXML_SEQUENCES[type].filter(slot => slot.length !== 1)).toEqual([]);
  });
});
