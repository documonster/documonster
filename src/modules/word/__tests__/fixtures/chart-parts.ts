/**
 * Chart parts for the chart round-trip tests: one built by hand to carry every
 * element of `CT_ChartSpace`, `CT_Chart`, `CT_PlotArea`, the chart groups, the
 * series, the data labels, the axes and the legend that Word's `Chart` model
 * does not type, and the relationships such a part has.
 */

const NS =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"';

const TX_PR = (size: number) =>
  `<c:txPr><a:bodyPr rot="-60000000" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${size}" b="1"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:pPr><a:endParaRPr lang="de-DE"/></a:p></c:txPr>`;
const SP_PR = (scheme: string) =>
  `<c:spPr><a:solidFill><a:schemeClr val="${scheme}"><a:lumMod val="75000"/></a:schemeClr></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>`;
const LAYOUT =
  '<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.1"/><c:y val="0.15"/><c:w val="0.7"/><c:h val="0.6"/></c:manualLayout></c:layout>';
const EXT = (uri: string) =>
  `<c:extLst><c:ext uri="{${uri}}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{0000000${uri.length % 10}-0000-0000-0000-000000000000}"/></c:ext></c:extLst>`;
const RICH = (text: string) =>
  `<c:tx><c:rich><a:bodyPr rot="0" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1600" b="1"/></a:pPr><a:r><a:rPr lang="de-DE" sz="1600" b="1"/><a:t>${text}</a:t></a:r></a:p></c:rich></c:tx>`;

function strRef(formula: string, values: readonly string[]): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
  return `<c:strRef><c:f>${formula}</c:f><c:strCache><c:ptCount val="${values.length}"/>${pts}</c:strCache></c:strRef>`;
}
function numRef(formula: string, values: readonly number[]): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
  return `<c:numRef><c:f>${formula}</c:f><c:numCache><c:formatCode>#,##0.00</c:formatCode><c:ptCount val="${values.length}"/>${pts}</c:numCache></c:numRef>`;
}

const DATA_LABELS =
  "<c:dLbls>" +
  '<c:dLbl><c:idx val="1"/><c:layout><c:manualLayout><c:x val="-0.02"/><c:y val="0.03"/></c:manualLayout></c:layout><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="de-DE"/><a:t>Peak</a:t></a:r></a:p></c:rich></c:tx><c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbl>' +
  '<c:numFmt formatCode="0.0%" sourceLinked="0"/>' +
  SP_PR("bg1") +
  TX_PR(800) +
  '<c:dLblPos val="inEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:separator>; </c:separator><c:showLeaderLines val="1"/>' +
  EXT("CE6537A1-D6FC-4f65-9D91-7224C49458BB") +
  "</c:dLbls>";

const BAR_SERIES =
  '<c:ser><c:idx val="5"/><c:order val="0"/>' +
  `<c:tx>${strRef("Tabelle1!$D$1", ["Umsatz"])}</c:tx>` +
  SP_PR("accent1") +
  '<c:invertIfNegative val="1"/>' +
  '<c:pictureOptions><c:applyToFront val="1"/></c:pictureOptions>' +
  '<c:dPt><c:idx val="3"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="FF8800"/></a:solidFill></c:spPr></c:dPt>' +
  DATA_LABELS +
  '<c:trendline><c:name>Trend</c:name><c:spPr><a:ln w="19050"><a:prstDash val="sysDot"/></a:ln></c:spPr><c:trendlineType val="poly"/><c:order val="3"/><c:forward val="1"/><c:backward val="0.5"/><c:dispRSqr val="1"/><c:dispEq val="1"/><c:trendlineLbl><c:numFmt formatCode="General" sourceLinked="0"/></c:trendlineLbl></c:trendline>' +
  '<c:errBars><c:errDir val="y"/><c:errBarType val="plus"/><c:errValType val="percentage"/><c:noEndCap val="1"/><c:val val="7"/><c:spPr><a:ln><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln></c:spPr></c:errBars>' +
  `<c:cat>${strRef("Tabelle1!$C$2:$C$5", ["Nord", "Süd", "Ost", "West"])}</c:cat>` +
  `<c:val>${numRef("Tabelle1!$D$2:$D$5", [12.5, 7, 3.25, 9])}</c:val>` +
  '<c:shape val="cylinder"/>' +
  EXT("C3380CC4-5D6E-409C-BE32-E72D297353CC") +
  "</c:ser>";

const LINE_SERIES =
  '<c:ser><c:idx val="9"/><c:order val="1"/>' +
  `<c:tx>${strRef("Tabelle1!$E$1", ["Marge"])}</c:tx>` +
  '<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>' +
  '<c:marker><c:symbol val="diamond"/><c:size val="7"/><c:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></c:spPr></c:marker>' +
  '<c:dLbls><c:delete val="1"/></c:dLbls>' +
  `<c:cat>${strRef("Tabelle1!$C$2:$C$5", ["Nord", "Süd", "Ost", "West"])}</c:cat>` +
  `<c:val>${numRef("Tabelle1!$E$2:$E$5", [0.1, 0.2, 0.15, 0.3])}</c:val>` +
  '<c:smooth val="1"/>' +
  "</c:ser>";

const CAT_AX =
  '<c:catAx><c:axId val="501"/><c:scaling><c:orientation val="maxMin"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>' +
  `<c:majorGridlines>${SP_PR("bg2")}</c:majorGridlines><c:minorGridlines/>` +
  `<c:title>${RICH("Region")}${LAYOUT}<c:overlay val="1"/>${SP_PR("bg1")}${TX_PR(1000)}</c:title>` +
  '<c:numFmt formatCode="@" sourceLinked="1"/><c:majorTickMark val="cross"/><c:minorTickMark val="in"/><c:tickLblPos val="low"/>' +
  SP_PR("tx2") +
  TX_PR(900) +
  '<c:crossAx val="502"/><c:crossesAt val="2.5"/><c:auto val="0"/><c:lblAlgn val="l"/><c:lblOffset val="300"/><c:tickLblSkip val="2"/><c:tickMarkSkip val="3"/><c:noMultiLvlLbl val="1"/>' +
  EXT("C3380CC4-0001") +
  "</c:catAx>";
const VAL_AX =
  '<c:valAx><c:axId val="502"/><c:scaling><c:logBase val="10"/><c:orientation val="minMax"/><c:max val="1000"/><c:min val="1"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>' +
  '<c:numFmt formatCode="#,##0 &quot;€&quot;" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="high"/>' +
  '<c:crossAx val="501"/><c:crosses val="max"/><c:crossBetween val="midCat"/><c:majorUnit val="10"/><c:minorUnit val="2"/>' +
  '<c:dispUnits><c:builtInUnit val="thousands"/><c:dispUnitsLbl/></c:dispUnits>' +
  "</c:valAx>";
const SER_AX =
  '<c:serAx><c:axId val="503"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="502"/><c:crosses val="autoZero"/><c:tickLblSkip val="1"/></c:serAx>';
const CAT_AX_2 =
  '<c:catAx><c:axId val="601"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="602"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>';
const VAL_AX_2 =
  '<c:valAx><c:axId val="602"/><c:scaling><c:orientation val="minMax"/><c:max val="0.5"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:numFmt formatCode="0%" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="601"/><c:crosses val="max"/><c:crossBetween val="between"/></c:valAx>';

/** `word/charts/chart1.xml`: every element `Chart` does not type, and a few it does. */
export const HAND_BUILT_CHART =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  `<c:chartSpace ${NS} mc:Ignorable="c16r2">` +
  '<c:date1904 val="1"/><c:lang val="de-DE"/><c:roundedCorners val="1"/>' +
  '<mc:AlternateContent><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent>' +
  '<c:clrMapOvr bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
  '<c:protection><c:chartObject val="0"/><c:data val="1"/><c:formatting val="0"/><c:selection val="0"/><c:userInterface val="0"/></c:protection>' +
  "<c:chart>" +
  `<c:title>${RICH("Umsatz und Marge")}${LAYOUT}<c:overlay val="0"/>${SP_PR("bg1")}${TX_PR(1600)}</c:title>` +
  '<c:autoTitleDeleted val="0"/>' +
  '<c:view3D><c:rotX val="10"/><c:rotY val="30"/><c:depthPercent val="150"/><c:rAngAx val="0"/><c:perspective val="20"/></c:view3D>' +
  `<c:floor><c:thickness val="0"/>${SP_PR("bg2")}</c:floor><c:sideWall><c:thickness val="0"/><c:spPr><a:noFill/></c:spPr></c:sideWall><c:backWall><c:thickness val="0"/><c:spPr><a:noFill/></c:spPr></c:backWall>` +
  `<c:plotArea>${LAYOUT}` +
  '<c:bar3DChart><c:barDir val="bar"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
  BAR_SERIES +
  '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>' +
  '<c:gapWidth val="80"/><c:gapDepth val="120"/><c:shape val="box"/><c:axId val="501"/><c:axId val="502"/><c:axId val="503"/></c:bar3DChart>' +
  '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' +
  LINE_SERIES +
  '<c:dropLines/><c:hiLowLines/><c:upDownBars><c:gapWidth val="150"/><c:upBars/><c:downBars/></c:upDownBars><c:marker val="1"/><c:smooth val="0"/><c:axId val="601"/><c:axId val="602"/></c:lineChart>' +
  CAT_AX +
  VAL_AX +
  SER_AX +
  CAT_AX_2 +
  VAL_AX_2 +
  '<c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="0"/><c:showOutline val="1"/><c:showKeys val="1"/></c:dTable>' +
  SP_PR("bg1") +
  EXT("C3380CC4-0002") +
  "</c:plotArea>" +
  '<c:legend><c:legendPos val="t"/><c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry>' +
  `${LAYOUT}<c:overlay val="1"/>${SP_PR("bg2")}${TX_PR(1100)}</c:legend>` +
  '<c:plotVisOnly val="0"/><c:dispBlanksAs val="zero"/><c:showDLblsOverMax val="1"/>' +
  '<c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst>' +
  "</c:chart>" +
  SP_PR("bg1") +
  TX_PR(1000) +
  '<c:externalData r:id="rId3"><c:autoUpdate val="0"/></c:externalData>' +
  '<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings>' +
  '<c:userShapes r:id="rId4"/>' +
  '<c:extLst><c:ext uri="{781A3756-C4B2-4CAC-9D66-4F8BD8637D16}" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:pivotOptions><c14:dropZoneFilter val="1"/></c14:pivotOptions></c:ext></c:extLst>' +
  "</c:chartSpace>";

/** The hand-built part's relationships, and the parts they point at. */
export const HAND_BUILT_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>' +
  '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/>' +
  '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes" Target="../drawings/drawing1.xml"/>' +
  "</Relationships>";

export const HAND_BUILT_RELATED_PARTS: ReadonlyMap<string, { data: string; contentType: string }> =
  new Map([
    [
      "word/charts/style1.xml",
      {
        data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="201"/>',
        contentType: "application/vnd.ms-office.chartstyle+xml"
      }
    ],
    [
      "word/charts/colors1.xml",
      {
        data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10"><a:schemeClr val="accent1"/></cs:colorStyle>',
        contentType: "application/vnd.ms-office.chartcolorstyle+xml"
      }
    ],
    [
      "word/embeddings/Microsoft_Excel_Worksheet.xlsx",
      {
        data: "PK-not-really-a-workbook",
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      }
    ],
    [
      "word/drawings/drawing1.xml",
      {
        data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>',
        contentType: "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml"
      }
    ]
  ]);
