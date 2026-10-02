/**
 * The drawing that places a chart across a whole chartsheet, shared by the XLSX and XLSB writers.
 */
import { uuidV4 } from "@utils/uuid";
import { xmlEncodeAttr } from "@xml/encode";

/**
 * A chartsheet drawing's extent, in EMU.
 *
 * ≈ 10.84″ × 6.67″ — A4 landscape minus default margins, which is what Excel writes. Both the
 * anchor-level and the frame-level size have to be non-zero or the chart renders as a blank canvas, so
 * one constant serves both.
 */
export const CHARTSHEET_DRAWING_EMU = { cx: 9906000, cy: 6096000 } as const;

/**
 * Build the chartsheet-drawing XML that wraps a single classic or
 * ChartEx chart occupying the entire chartsheet canvas.
 *
 * Chartsheets have no cell grid — `sheetData` is empty and there are
 * no `<cols>` / `<row>` sizing entries for Excel to lay an anchor
 * against. A cell-based `<xdr:twoCellAnchor from="A1" to="R31"/>`
 * (what the generic `DrawingXform` emits) therefore resolves to a
 * 0×0 bounding box on a chartsheet, and Excel renders a blank
 * white canvas with no chart inside. Using `<xdr:absoluteAnchor>`
 * with concrete EMU coordinates is how Excel itself writes
 * chartsheet drawings — the anchor's `pos`/`ext` pair gives the
 * engine something real to lay the graphic against, while the
 * inner `<xdr:graphicFrame>/<xdr:xfrm>` stays at zero, exactly as
 * Excel writes it and as every worksheet chart here already does.
 *
 * That inner extent used to repeat the anchor's, on the reasoning
 * that the graphic then filled it. It does the opposite: the frame
 * re-sized itself against a transform it should have inherited and
 * the chart was never drawn, so a chartsheet showed a blank canvas
 * in both containers.
 *
 * ChartEx drawings additionally need an `<mc:AlternateContent>`
 * wrapper around the `<xdr:graphicFrame>` — the `cx` namespace is
 * a Microsoft extension that legacy-Excel loaders don't understand,
 * so the Fallback branch emits a placeholder shape (the same
 * "This chart isn't available in your version of Excel" message
 * Office uses).
 *
 * Exported so the XLSB writer can produce the same drawing: the reasoning applies identically to a `.bin`
 * chartsheet.
 */
export function renderChartsheetDrawingXml(options: {
  chartRId: string;
  chartName: string;
  isChartEx: boolean;
  extCx: number;
  extCy: number;
}): string {
  const { chartRId, chartName, isChartEx, extCx, extCy } = options;
  const escName = xmlEncodeAttr(chartName);
  const escRId = xmlEncodeAttr(chartRId);
  const cNvPrExtLst = isChartEx
    ? `<a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{${uuidV4().toUpperCase()}}"/></a:ext></a:extLst>`
    : "";
  const graphicFrame =
    `<xdr:graphicFrame macro="">` +
    `<xdr:nvGraphicFramePr>` +
    (cNvPrExtLst
      ? // **`id="2"`, which is what Excel writes.** A `cNvPr` id must be unique within the drawing and non-zero;
        // 1 satisfies both, so this is not a validity fix but a conformance one — Excel numbers the first shape
        // in a drawing 2, reserving 1, and its re-save of this library's chartsheet drawing does the same.
        `<xdr:cNvPr id="2" name="${escName}">${cNvPrExtLst}</xdr:cNvPr>`
      : `<xdr:cNvPr id="2" name="${escName}"/>`) +
    // **`<a:graphicFrameLocks/>`, not an empty element.** Excel writes the child in its own chartsheet
    // drawings — both in the corpus and in its re-save of this library's `financial-report.xlsb` — and an empty
    // `cNvGraphicFramePr` is the one remaining structural difference between the two once the frame's transform
    // is right. The element with no attributes is what Excel writes: `noGrp` is not set on a chartsheet's frame,
    // which occupies the sheet alone and has nothing to be grouped with.
    `<xdr:cNvGraphicFramePr><a:graphicFrameLocks/></xdr:cNvGraphicFramePr>` +
    `</xdr:nvGraphicFramePr>` +
    // **Zero, not the anchor's extent.** The frame's own transform is inherited from the anchor; Excel writes
    // `<a:ext cx="0" cy="0"/>` here in its own chartsheet drawings, and so does the generic `GraphicFrameXform`
    // that serves every worksheet chart in this library — whose comment says "position/size handled by anchor,
    // so use zeros". Repeating the extent here produced a chartsheet that rendered as an empty canvas: the
    // anchor sized the frame and the frame then re-sized itself, and the chart inside it was never drawn.
    //
    // Worth being precise about which half of the original reasoning was wrong, because the other half is
    // right and must stay: a chartsheet has no cell grid, so a `twoCellAnchor` does resolve to 0×0 there and
    // `absoluteAnchor` with real EMU on the *anchor* is what Excel writes. It was only the inner repeat that
    // was invented, and it was written into the comment as though it had been observed.
    `<xdr:xfrm>` +
    `<a:off x="0" y="0"/>` +
    `<a:ext cx="0" cy="0"/>` +
    `</xdr:xfrm>` +
    `<a:graphic>` +
    (isChartEx
      ? `<a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">` +
        `<cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${escRId}"/>` +
        `</a:graphicData>`
      : `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
        `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${escRId}"/>` +
        `</a:graphicData>`) +
    `</a:graphic>` +
    `</xdr:graphicFrame>`;

  const fallbackShape =
    `<xdr:sp macro="" textlink="">` +
    `<xdr:nvSpPr>` +
    `<xdr:cNvPr id="0" name=""/>` +
    `<xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr>` +
    `</xdr:nvSpPr>` +
    `<xdr:spPr>` +
    `<a:xfrm>` +
    `<a:off x="0" y="0"/>` +
    `<a:ext cx="${extCx}" cy="${extCy}"/>` +
    `</a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:prstClr val="white"/></a:solidFill>` +
    `<a:ln w="1"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln>` +
    `</xdr:spPr>` +
    `<xdr:txBody>` +
    `<a:bodyPr vertOverflow="clip" horzOverflow="clip"/>` +
    `<a:lstStyle/>` +
    `<a:p><a:r><a:rPr lang="en-US" sz="1100"/>` +
    `<a:t>This chart isn&apos;t available in your version of Excel.\n\n` +
    `Editing this shape or saving this workbook into a different file format will permanently break the chart.</a:t>` +
    `</a:r></a:p>` +
    `</xdr:txBody>` +
    `</xdr:sp>`;

  const anchorBody = isChartEx
    ? `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
      `<mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1">` +
      graphicFrame +
      `</mc:Choice>` +
      `<mc:Fallback>` +
      fallbackShape +
      `</mc:Fallback>` +
      `</mc:AlternateContent>`
    : graphicFrame;

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<xdr:absoluteAnchor>` +
    `<xdr:pos x="0" y="0"/>` +
    `<xdr:ext cx="${extCx}" cy="${extCy}"/>` +
    anchorBody +
    `<xdr:clientData/>` +
    `</xdr:absoluteAnchor>` +
    `</xdr:wsDr>`
  );
}
