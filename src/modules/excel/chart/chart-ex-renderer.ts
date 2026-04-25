/**
 * ChartEx renderer — serialises a ChartExModel to `cx:chart` XML.
 *
 * This is a standalone renderer (not a full SAX parser): it produces byte
 * output for a programmatically-built chartEx. Round-trip of existing cx:chart
 * files is handled by raw byte passthrough (model.rawXml is preferred when set).
 */

import type { ChartExAxis, ChartExDataEntry, ChartExModel, ChartExSeries } from "./chart-ex-types";
import type { ChartColor, ChartTitle, ShapeProperties } from "./types";

/**
 * Render a ChartExModel to the full XML string representation of cx:chart.
 * If the model was loaded from a file (has `rawXml`), returns that unchanged.
 */
export function renderChartEx(model: ChartExModel): string {
  // Prefer raw XML for existing round-tripped charts
  if (model.rawXml) {
    return model.rawXml;
  }

  const parts: string[] = [];
  parts.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  parts.push(
    [
      "<cx:chartSpace",
      '  xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"',
      '  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
      '  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
      '  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    ].join("\n")
  );
  const space = model.chartSpace;
  parts.push(renderChartData(space.chartData));
  parts.push(renderChart(space.chart));
  if (space.clrMapOvr) {
    parts.push(space.clrMapOvr);
  }
  if (space.extLst) {
    parts.push(space.extLst);
  }
  parts.push("</cx:chartSpace>");
  return parts.join("\n");
}

function renderChartData(data: ChartExModel["chartSpace"]["chartData"]): string {
  const parts: string[] = [];
  parts.push("  <cx:chartData>");
  if (data.externalData) {
    for (const ed of data.externalData) {
      const attrs = ed.autoUpdate === undefined ? "" : ` autoUpdate="${ed.autoUpdate ? "1" : "0"}"`;
      parts.push(`    <cx:externalData r:id="${ed.id}"${attrs}/>`);
    }
  }
  for (const entry of data.data) {
    parts.push(renderDataEntry(entry));
  }
  parts.push("  </cx:chartData>");
  return parts.join("\n");
}

function renderDataEntry(entry: ChartExDataEntry): string {
  const parts: string[] = [];
  parts.push(`    <cx:data id="${entry.id}">`);
  if (entry.strDim) {
    const d = entry.strDim;
    parts.push(`      <cx:strDim type="${d.type}">`);
    if (d.formula) {
      parts.push(`        <cx:f>${escapeXml(d.formula)}</cx:f>`);
    }
    if (d.levels) {
      for (const lvl of d.levels) {
        const ptAttr = lvl.ptCount !== undefined ? ` ptCount="${lvl.ptCount}"` : "";
        if (lvl.points.length === 0) {
          parts.push(`        <cx:lvl${ptAttr}/>`);
        } else {
          parts.push(`        <cx:lvl${ptAttr}>`);
          for (const p of lvl.points) {
            parts.push(`          <cx:pt idx="${p.index}">${escapeXml(p.value)}</cx:pt>`);
          }
          parts.push("        </cx:lvl>");
        }
      }
    }
    parts.push("      </cx:strDim>");
  }
  if (entry.numDim) {
    const d = entry.numDim;
    parts.push(`      <cx:numDim type="${d.type}">`);
    if (d.formula) {
      parts.push(`        <cx:f>${escapeXml(d.formula)}</cx:f>`);
    }
    if (d.levels) {
      for (const lvl of d.levels) {
        const fmtAttr = lvl.formatCode ? ` formatCode="${escapeAttr(lvl.formatCode)}"` : "";
        const ptAttr = lvl.ptCount !== undefined ? ` ptCount="${lvl.ptCount}"` : "";
        if (lvl.points.length === 0) {
          parts.push(`        <cx:lvl${ptAttr}${fmtAttr}/>`);
        } else {
          parts.push(`        <cx:lvl${ptAttr}${fmtAttr}>`);
          for (const p of lvl.points) {
            parts.push(`          <cx:pt idx="${p.index}">${p.value}</cx:pt>`);
          }
          parts.push("        </cx:lvl>");
        }
      }
    }
    parts.push("      </cx:numDim>");
  }
  parts.push("    </cx:data>");
  return parts.join("\n");
}

function renderChart(chart: ChartExModel["chartSpace"]["chart"]): string {
  const parts: string[] = [];
  parts.push("  <cx:chart>");
  if (chart.title) {
    parts.push(renderTitle(chart.title));
  }
  if (chart.autoTitleDeleted !== undefined && !chart.title) {
    parts.push(`    <cx:autoTitleDeleted val="${chart.autoTitleDeleted ? "1" : "0"}"/>`);
  }
  parts.push(renderPlotArea(chart.plotArea));
  if (chart.legend) {
    parts.push(renderLegend(chart.legend));
  }
  parts.push("  </cx:chart>");
  return parts.join("\n");
}

function renderTitle(title: ChartTitle): string {
  const parts: string[] = [];
  parts.push("    <cx:title>");
  if (title.text) {
    parts.push("      <cx:tx>");
    parts.push("        <cx:rich>");
    parts.push("          <a:bodyPr/>");
    parts.push("          <a:lstStyle/>");
    for (const p of title.text.paragraphs) {
      parts.push("          <a:p>");
      for (const run of p.runs ?? []) {
        parts.push("            <a:r>");
        parts.push(`              <a:t>${escapeXml(run.text)}</a:t>`);
        parts.push("            </a:r>");
      }
      parts.push("          </a:p>");
    }
    parts.push("        </cx:rich>");
    parts.push("      </cx:tx>");
  }
  parts.push(`      <cx:overlay val="${title.overlay ? "1" : "0"}"/>`);
  parts.push("    </cx:title>");
  return parts.join("\n");
}

function renderPlotArea(pa: ChartExModel["chartSpace"]["chart"]["plotArea"]): string {
  const parts: string[] = [];
  parts.push("    <cx:plotArea>");
  const region = pa.plotAreaRegion;
  if (region) {
    parts.push("      <cx:plotAreaRegion>");
    if (region.plotSurface) {
      parts.push(renderSpPr(region.plotSurface, "      "));
    }
    for (const s of region.series) {
      parts.push(renderSeries(s));
    }
    parts.push("      </cx:plotAreaRegion>");
  } else if (pa.series) {
    for (const s of pa.series) {
      parts.push(renderSeries(s));
    }
  }
  if (pa.axis) {
    for (const axis of pa.axis) {
      parts.push(renderAxis(axis));
    }
  }
  parts.push("    </cx:plotArea>");
  return parts.join("\n");
}

function renderSeries(s: ChartExSeries): string {
  const parts: string[] = [];
  const attrs = [`layoutId="${s.layoutId}"`];
  if (s.hidden) {
    attrs.push('hidden="1"');
  }
  if (s.ownerIdx !== undefined) {
    attrs.push(`ownerIdx="${s.ownerIdx}"`);
  }
  parts.push(`        <cx:series ${attrs.join(" ")}>`);
  if (s.tx) {
    if (s.tx.value !== undefined) {
      parts.push(
        `          <cx:tx><cx:txData><cx:v>${escapeXml(s.tx.value)}</cx:v></cx:txData></cx:tx>`
      );
    } else if (s.tx.strRef) {
      parts.push(
        `          <cx:tx><cx:txData><cx:f>${escapeXml(s.tx.strRef)}</cx:f></cx:txData></cx:tx>`
      );
    }
  }
  if (s.spPr) {
    parts.push(renderSpPr(s.spPr, "          "));
  }
  if (s.dataRefs) {
    for (const ref of s.dataRefs) {
      if (ref.dataId !== undefined) {
        parts.push(`          <cx:dataId val="${ref.dataId}"/>`);
      }
    }
  }
  if (s.layoutPr) {
    parts.push(renderLayoutProperties(s.layoutId, s.layoutPr));
  }
  if (s.axisId) {
    for (const id of s.axisId) {
      parts.push(`          <cx:axisId val="${id}"/>`);
    }
  }
  if (s.dataLabels) {
    parts.push(renderDataLabels(s.dataLabels));
  }
  if (s.dataPt) {
    for (const dp of s.dataPt) {
      parts.push(`          <cx:dataPt idx="${dp.idx}">`);
      if (dp.spPr) {
        parts.push(renderSpPr(dp.spPr, "            "));
      }
      parts.push("          </cx:dataPt>");
    }
  }
  if (s.extLst) {
    parts.push(s.extLst);
  }
  parts.push("        </cx:series>");
  return parts.join("\n");
}

function renderLayoutProperties(
  layoutId: string,
  lp: NonNullable<ChartExSeries["layoutPr"]>
): string {
  const parts: string[] = [];
  parts.push("          <cx:layoutPr>");
  if (lp.parentLabelLayout && (layoutId === "sunburst" || layoutId === "treemap")) {
    parts.push(`            <cx:parentLabelLayout val="${lp.parentLabelLayout}"/>`);
  }
  if (lp.subtotals && layoutId === "waterfall") {
    parts.push("            <cx:subtotals>");
    for (const st of lp.subtotals) {
      parts.push(`              <cx:subtotal idx="${st.idx}"/>`);
    }
    parts.push("            </cx:subtotals>");
  }
  if (lp.binning) {
    const b = lp.binning;
    const attrs: string[] = [];
    if (b.intervalClosed) {
      attrs.push(`intervalClosed="${b.intervalClosed}"`);
    }
    if (b.underflow !== undefined) {
      attrs.push(`underflow="${b.underflow}"`);
    }
    if (b.overflow !== undefined) {
      attrs.push(`overflow="${b.overflow}"`);
    }
    const attrStr = attrs.length > 0 ? " " + attrs.join(" ") : "";
    parts.push(`            <cx:binning${attrStr}>`);
    if (b.binSize !== undefined) {
      parts.push(`              <cx:binSize val="${b.binSize}"/>`);
    }
    if (b.binCount !== undefined) {
      parts.push(`              <cx:binCount val="${b.binCount}"/>`);
    }
    if (b.binType) {
      // binType is represented differently in XML — actually a binCountOrSize element pick
    }
    parts.push("            </cx:binning>");
  }
  if (layoutId === "boxWhisker") {
    if (lp.quartileMethod) {
      parts.push(`            <cx:quartileMethod val="${lp.quartileMethod}"/>`);
    }
    if (lp.showMeanLine !== undefined) {
      parts.push(`            <cx:showMeanLine val="${lp.showMeanLine ? "1" : "0"}"/>`);
    }
    if (lp.showMeanMarker !== undefined) {
      parts.push(`            <cx:showMeanMarker val="${lp.showMeanMarker ? "1" : "0"}"/>`);
    }
    if (lp.showInnerPoints !== undefined) {
      parts.push(`            <cx:showInnerPoints val="${lp.showInnerPoints ? "1" : "0"}"/>`);
    }
    if (lp.showOutlierPoints !== undefined) {
      parts.push(`            <cx:showOutlierPoints val="${lp.showOutlierPoints ? "1" : "0"}"/>`);
    }
  }
  if (layoutId === "regionMap") {
    if (lp.projection) {
      parts.push(`            <cx:projection val="${lp.projection}"/>`);
    }
    if (lp.regionLabels) {
      parts.push(`            <cx:regionLabels val="${lp.regionLabels}"/>`);
    }
    if (lp.geoMappingLevel) {
      parts.push(`            <cx:geoMappingLevel val="${lp.geoMappingLevel}"/>`);
    }
  }
  parts.push("          </cx:layoutPr>");
  return parts.join("\n");
}

function renderDataLabels(dl: NonNullable<ChartExSeries["dataLabels"]>): string {
  const parts: string[] = [];
  parts.push("          <cx:dataLabels>");
  if (dl.visibility) {
    const v = dl.visibility;
    const attrs: string[] = [];
    if (v.seriesName !== undefined) {
      attrs.push(`seriesName="${v.seriesName ? "1" : "0"}"`);
    }
    if (v.categoryName !== undefined) {
      attrs.push(`categoryName="${v.categoryName ? "1" : "0"}"`);
    }
    if (v.value !== undefined) {
      attrs.push(`value="${v.value ? "1" : "0"}"`);
    }
    parts.push(`            <cx:visibility ${attrs.join(" ")}/>`);
  }
  if (dl.position) {
    parts.push(`            <cx:dataLabel pos="${dl.position}"/>`);
  }
  if (dl.separator) {
    parts.push(`            <cx:separator>${escapeXml(dl.separator)}</cx:separator>`);
  }
  if (dl.numFmt) {
    parts.push(`            <cx:numFmt formatCode="${escapeAttr(dl.numFmt)}"/>`);
  }
  parts.push("          </cx:dataLabels>");
  return parts.join("\n");
}

function renderAxis(axis: ChartExAxis): string {
  const parts: string[] = [];
  parts.push(`      <cx:axis id="${axis.axisId}">`);
  if (axis.hidden) {
    parts.push('        <cx:hidden val="1"/>');
  }
  if (axis.majorTickMark) {
    parts.push(`        <cx:majorTickMark val="${axis.majorTickMark}"/>`);
  }
  if (axis.minorTickMark) {
    parts.push(`        <cx:minorTickMark val="${axis.minorTickMark}"/>`);
  }
  if (axis.numFmt) {
    const attrs = [`formatCode="${escapeAttr(axis.numFmt.formatCode)}"`];
    if (axis.numFmt.sourceLinked !== undefined) {
      attrs.push(`sourceLinked="${axis.numFmt.sourceLinked ? "1" : "0"}"`);
    }
    parts.push(`        <cx:numFmt ${attrs.join(" ")}/>`);
  }
  if (axis.title) {
    parts.push(renderTitle(axis.title));
  }
  if (axis.valScaling) {
    const vs = axis.valScaling;
    const attrs: string[] = [];
    if (vs.min !== undefined) {
      attrs.push(`min="${vs.min}"`);
    }
    if (vs.max !== undefined) {
      attrs.push(`max="${vs.max}"`);
    }
    if (vs.majorUnit !== undefined) {
      attrs.push(`majorUnit="${vs.majorUnit}"`);
    }
    if (vs.minorUnit !== undefined) {
      attrs.push(`minorUnit="${vs.minorUnit}"`);
    }
    parts.push(`        <cx:valScaling ${attrs.join(" ")}/>`);
  }
  if (axis.catScaling) {
    const cs = axis.catScaling;
    const attrs: string[] = [];
    if (cs.gapWidth !== undefined) {
      attrs.push(`gapWidth="${cs.gapWidth}"`);
    }
    parts.push(`        <cx:catScaling ${attrs.join(" ")}/>`);
  }
  if (axis.spPr) {
    parts.push(renderSpPr(axis.spPr, "        "));
  }
  parts.push("      </cx:axis>");
  return parts.join("\n");
}

function renderLegend(l: NonNullable<ChartExModel["chartSpace"]["chart"]["legend"]>): string {
  const parts: string[] = [];
  const attrs: string[] = [];
  if (l.legendPos) {
    attrs.push(`pos="${l.legendPos}"`);
  }
  if (l.overlay !== undefined) {
    attrs.push(`align="ctr" overlay="${l.overlay ? "1" : "0"}"`);
  }
  const hasChildren = !!(l.spPr || l.legendEntries);
  if (hasChildren) {
    parts.push(`    <cx:legend ${attrs.join(" ")}>`);
    if (l.spPr) {
      parts.push(renderSpPr(l.spPr, "      "));
    }
    if (l.legendEntries) {
      for (const entry of l.legendEntries) {
        parts.push(`      <cx:legendEntry idx="${entry.index}"/>`);
      }
    }
    parts.push("    </cx:legend>");
  } else {
    parts.push(`    <cx:legend ${attrs.join(" ")}/>`);
  }
  return parts.join("\n");
}

function renderSpPr(spPr: ShapeProperties, indent: string): string {
  if (spPr._rawXml) {
    return indent + spPr._rawXml;
  }
  const parts: string[] = [];
  parts.push(`${indent}<cx:spPr>`);
  if (spPr.fill) {
    if (spPr.fill.noFill) {
      parts.push(`${indent}  <a:noFill/>`);
    } else if (spPr.fill.solid) {
      parts.push(`${indent}  <a:solidFill>${renderColor(spPr.fill.solid)}</a:solidFill>`);
    } else if (spPr.fill.gradient) {
      const g = spPr.fill.gradient;
      if (g.stops.length >= 2) {
        parts.push(`${indent}  <a:gradFill>`);
        parts.push(`${indent}    <a:gsLst>`);
        for (const stop of g.stops) {
          parts.push(
            `${indent}      <a:gs pos="${Math.round(stop.position * 1000)}">${renderColor(stop.color)}</a:gs>`
          );
        }
        parts.push(`${indent}    </a:gsLst>`);
        if (g.type === "linear" || g.type === undefined) {
          parts.push(`${indent}    <a:lin ang="${(g.angle ?? 0) * 60000}" scaled="1"/>`);
        } else {
          parts.push(
            `${indent}    <a:path path="${g.type}"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>`
          );
        }
        parts.push(`${indent}  </a:gradFill>`);
      }
    } else if (spPr.fill.pattern) {
      const p = spPr.fill.pattern;
      parts.push(`${indent}  <a:pattFill prst="${p.preset}">`);
      if (p.foreground) {
        parts.push(`${indent}    <a:fgClr>${renderColor(p.foreground)}</a:fgClr>`);
      }
      if (p.background) {
        parts.push(`${indent}    <a:bgClr>${renderColor(p.background)}</a:bgClr>`);
      }
      parts.push(`${indent}  </a:pattFill>`);
    }
  }
  if (spPr.line) {
    const widthAttr = spPr.line.width ? ` w="${spPr.line.width}"` : "";
    if (spPr.line.noFill) {
      parts.push(`${indent}  <a:ln${widthAttr}><a:noFill/></a:ln>`);
    } else if (spPr.line.color) {
      const dashPart = spPr.line.dash ? `<a:prstDash val="${spPr.line.dash}"/>` : "";
      parts.push(
        `${indent}  <a:ln${widthAttr}><a:solidFill>${renderColor(spPr.line.color)}</a:solidFill>${dashPart}</a:ln>`
      );
    } else {
      const dashPart = spPr.line.dash ? `<a:prstDash val="${spPr.line.dash}"/>` : "";
      if (dashPart) {
        parts.push(`${indent}  <a:ln${widthAttr}>${dashPart}</a:ln>`);
      } else {
        parts.push(`${indent}  <a:ln${widthAttr}/>`);
      }
    }
  }
  parts.push(`${indent}</cx:spPr>`);
  // NOTE: effectList (a:effectLst), scene3d (a:scene3d), and sp3d (a:sp3d) are not
  // rendered for ChartEx shapes. These DrawingML elements are supported in standard
  // c:spPr via _renderSpPr in chart-space-xform.ts but are omitted here because
  // ChartEx usage of cx:spPr rarely includes them. If round-trip fidelity is needed,
  // the raw XML passthrough (_rawXml) path at the top of this function preserves them.
  return parts.join("\n");
}

function renderColor(c: ChartColor): string {
  const modifiers = renderColorModifiers(c);
  if (c.srgb) {
    if (modifiers) {
      return `<a:srgbClr val="${c.srgb}">${modifiers}</a:srgbClr>`;
    }
    return `<a:srgbClr val="${c.srgb}"/>`;
  }
  if (c.theme !== undefined) {
    const themeNames = [
      "dk1",
      "lt1",
      "dk2",
      "lt2",
      "accent1",
      "accent2",
      "accent3",
      "accent4",
      "accent5",
      "accent6",
      "hlink",
      "folHlink"
    ];
    const name = themeNames[c.theme] ?? "dk1";
    if (modifiers) {
      return `<a:schemeClr val="${name}">${modifiers}</a:schemeClr>`;
    }
    return `<a:schemeClr val="${name}"/>`;
  }
  if (c.sysClr) {
    if (modifiers) {
      return `<a:sysClr val="${c.sysClr}">${modifiers}</a:sysClr>`;
    }
    return `<a:sysClr val="${c.sysClr}"/>`;
  }
  if (c.prstClr) {
    if (modifiers) {
      return `<a:prstClr val="${c.prstClr}">${modifiers}</a:prstClr>`;
    }
    return `<a:prstClr val="${c.prstClr}"/>`;
  }
  return "";
}

function renderColorModifiers(c: ChartColor): string {
  const parts: string[] = [];
  if (c.alpha !== undefined) {
    parts.push(`<a:alpha val="${c.alpha}"/>`);
  }
  if (c.tint !== undefined) {
    parts.push(`<a:tint val="${Math.round(c.tint * 100000)}"/>`);
  }
  if (c.shade !== undefined) {
    parts.push(`<a:shade val="${c.shade}"/>`);
  }
  if (c.satMod !== undefined) {
    parts.push(`<a:satMod val="${c.satMod}"/>`);
  }
  if (c.lumMod !== undefined) {
    parts.push(`<a:lumMod val="${c.lumMod}"/>`);
  }
  if (c.lumOff !== undefined) {
    parts.push(`<a:lumOff val="${c.lumOff}"/>`);
  }
  return parts.join("");
}

function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeAttr(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
