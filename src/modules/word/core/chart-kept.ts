/**
 * DOCX Module - chart part markup kept beside the chart model's fields.
 *
 * The reader keeps every element of a chart part (see `ChartKeptNode`), and
 * records beside each child the field it is written from, as a *field key*.
 * The writer and the renderer compute the same keys from the chart as it is
 * now: a child whose key is unchanged is written (or drawn) as read, one whose
 * key changed is written from the field. One table of keys, used by all three,
 * so a chart cannot be read one way, written another and drawn a third.
 *
 * A key of `undefined` means no field governs the child: it is always kept.
 */

import { unrecognizedValue } from "@utils/unrecognized-values";
import type {
  Chart,
  ChartAxis,
  ChartDataLabels,
  ChartKeptAxis,
  ChartKeptElement,
  ChartKeptNode,
  ChartSeries,
  ChartType
} from "@word/types";

/** Field keys by child local name; `""` is the element itself (whether it is written). */
export type ChartFieldKeys = Readonly<Record<string, string | undefined>>;

/** A field value as a key: equal keys mean the field holds what it held. */
export function fieldKey(value: unknown): string {
  return JSON.stringify(value) ?? "undefined";
}

const isScatter = (type: ChartType): boolean => type === "scatter" || type === "scatterSmooth";

/** The identity of a chart group: series of one type on one pair of axes. */
export function chartGroupKey(type: ChartType, secondary: boolean): string {
  return `${type}|${secondary ? 1 : 0}`;
}

/** `c:chartSpace`. */
export function chartSpaceKeys(chart: Chart): ChartFieldKeys {
  const style = fieldKey(chart.style);
  return {
    AlternateContent: style,
    style,
    spPr: fieldKey(chart.chartAreaColor),
    externalData: fieldKey(!!chart.embedWorkbook)
  };
}

/** `c:chart`. */
export function chartKeys(chart: Chart): ChartFieldKeys {
  return { autoTitleDeleted: fieldKey(chart.title), view3D: fieldKey(!!chart.view3d) };
}

/** A `c:title`, of the chart or of an axis. */
export function titleKeys(text: string | undefined): ChartFieldKeys {
  return { "": fieldKey(text), tx: fieldKey(text) };
}

/** `c:plotArea`. */
export function plotAreaKeys(chart: Chart): ChartFieldKeys {
  return { spPr: fieldKey(chart.plotAreaColor) };
}

/** `c:legend`. */
export function legendKeys(chart: Chart): ChartFieldKeys {
  return { legendPos: fieldKey([chart.legend, unrecognizedValue(chart, "legend")]) };
}

/** A chart group's `c:dLbls`: the chart's `dataLabels`. */
export function groupDataLabelsKeys(dataLabels: ChartDataLabels | undefined): ChartFieldKeys {
  const position = dataLabels
    ? [dataLabels.position, unrecognizedValue(dataLabels, "position")]
    : undefined;
  return {
    "": fieldKey(dataLabels),
    delete: fieldKey(dataLabels),
    dLblPos: fieldKey(position),
    showVal: fieldKey(!!dataLabels?.showValue),
    showCatName: fieldKey(!!dataLabels?.showCategory),
    showSerName: fieldKey(!!dataLabels?.showSerName),
    showPercent: fieldKey(!!dataLabels?.showPercent)
  };
}

/** A series' `c:dLbls`: its `showDataLabels`. */
export function seriesDataLabelsKeys(series: ChartSeries): ChartFieldKeys {
  const shown = fieldKey(!!series.showDataLabels);
  return { "": shown, delete: shown, showVal: shown };
}

/**
 * A `c:ser`, at `index` among the chart's series and drawn as `type`. `c:dPt`
 * is keyed per point (`ChartKeptElement.read`), not here.
 */
export function seriesKeys(series: ChartSeries, index: number, type: ChartType): ChartFieldKeys {
  const position = fieldKey(index);
  const categories = fieldKey([series.categories, isScatter(type)]);
  const values = fieldKey([series.values, isScatter(type)]);
  return {
    idx: position,
    order: position,
    tx: fieldKey(series.name),
    spPr: fieldKey([series.color, type]),
    marker: fieldKey(type),
    trendline: fieldKey(series.trendline),
    errBars: fieldKey(series.errorBars),
    cat: categories,
    xVal: categories,
    val: values,
    yVal: values,
    smooth: fieldKey(type)
  };
}

/** The key of point `index`'s colour. */
export function pointKey(series: ChartSeries, index: number): string {
  return fieldKey(series.pointColors?.[index]);
}

/** An axis (`c:catAx`, `c:valAx`, `c:dateAx`), read into `axis`. */
export function axisKeys(axis: ChartAxis | undefined, role: ChartKeptAxis["role"]): ChartFieldKeys {
  return {
    delete: role === "secondaryCategory" ? undefined : fieldKey(!!axis?.hidden),
    numFmt: fieldKey(axis?.numberFormat),
    majorUnit: fieldKey(axis?.majorUnit)
  };
}

/** An axis' `c:scaling`. */
export function scalingKeys(axis: ChartAxis | undefined): ChartFieldKeys {
  return { max: fieldKey(axis?.max), min: fieldKey(axis?.min) };
}

/** The kept child `name` of `node` while its field key is still `now`; else `undefined`. */
export function keptChild(
  node: ChartKeptNode | undefined,
  name: string,
  keys: ChartFieldKeys
): ChartKeptElement | undefined {
  if (node === undefined || (keys[name] !== undefined && node.read[name] !== keys[name])) {
    return undefined;
  }
  return node.children.find(child => child.name === name && child.xml !== "");
}

/** A series' kept `c:dPt` elements whose colour is still what was read. */
export function keptPoints(series: ChartSeries): ChartKeptElement[] {
  return (series._kept?.node.children ?? []).filter(
    child =>
      child.name === "dPt" &&
      child.index !== undefined &&
      child.read === pointKey(series, child.index)
  );
}
