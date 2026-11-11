/**
 * The colour scheme of a workbook's theme, for the chart renderers.
 *
 * A chart names most of its colours by scheme slot — `<a:schemeClr val="accent1"/>`,
 * `tx1`, `bg1` — and Excel resolves them against the workbook's theme part
 * (`xl/theme/theme1.xml`, `a:theme/a:themeElements/a:clrScheme`). A workbook saved by a
 * recent Excel carries the 2023 Office theme (`accent1` = `156082`), not the 2013 one the
 * renderers fall back to (`4472C4`), so without this every scheme colour of such a chart came
 * out in the wrong theme.
 */

import type { ChartThemeColors } from "@excel/chart/shared/chart-utils";
import type { WorkbookData } from "@excel/core/workbook-core";
import type { OoxmlThemeColorName } from "@utils/theme-colors";
import { trimXmlPartLead } from "@utils/xml-part";
import { parseXml } from "@xml/dom";
import type { XmlElement, XmlNode } from "@xml/types";

const DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

const SLOTS: readonly OoxmlThemeColorName[] = [
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

/** Parsed schemes, by the theme XML they came from: a workbook's charts share one. */
const cache = new Map<string, ChartThemeColors | undefined>();
const CACHE_LIMIT = 16;

/**
 * The colour scheme of `workbook`'s theme, or `undefined` when it has none that can be read
 * (a workbook created in memory, which is written with the default Office theme).
 */
export function workbookThemeColors(
  workbook: WorkbookData | undefined
): ChartThemeColors | undefined {
  const themes = (workbook as { _themes?: unknown } | undefined)?._themes;
  if (themes === null || typeof themes !== "object") {
    return undefined;
  }
  const record = themes as Record<string, unknown>;
  const xml = record.theme1 ?? Object.values(record)[0];
  return typeof xml === "string" ? themeXmlColors(xml) : undefined;
}

/** The `a:clrScheme` of a theme part, slot by slot; `undefined` when there is none. */
export function themeXmlColors(xml: string): ChartThemeColors | undefined {
  if (cache.has(xml)) {
    return cache.get(xml);
  }
  let colors: ChartThemeColors | undefined;
  try {
    const root = parseXml(trimXmlPartLead(xml), { xmlns: true }).root;
    const scheme = isDrawingMl(root, "theme")
      ? drawingMlChild(drawingMlChild(root, "themeElements"), "clrScheme")
      : undefined;
    if (scheme) {
      const out: Partial<Record<OoxmlThemeColorName, string>> = {};
      for (const slot of elementChildren(scheme)) {
        if (slot.uri !== DRAWINGML_NS) {
          continue;
        }
        const name = localName(slot) as OoxmlThemeColorName;
        if (!SLOTS.includes(name)) {
          continue;
        }
        const value = slotColor(elementChildren(slot)[0]);
        if (value !== undefined) {
          out[name] = value;
        }
      }
      colors = Object.keys(out).length > 0 ? out : undefined;
    }
  } catch {
    // A theme part that is not XML is a theme with nothing to say: the defaults apply.
    colors = undefined;
  }
  if (cache.size >= CACHE_LIMIT) {
    cache.delete(cache.keys().next().value!);
  }
  cache.set(xml, colors);
  return colors;
}

/** `a:srgbClr/@val`, or a system colour's last computed value (`a:sysClr/@lastClr`). */
function slotColor(element: XmlElement | undefined): string | undefined {
  if (!element || element.uri !== DRAWINGML_NS) {
    return undefined;
  }
  const kind = localName(element);
  const value =
    kind === "srgbClr"
      ? element.attributes.val
      : kind === "sysClr"
        ? (element.attributes.lastClr ??
          (element.attributes.val === "windowText"
            ? "000000"
            : element.attributes.val === "window"
              ? "FFFFFF"
              : undefined))
        : undefined;
  return value !== undefined && /^[0-9A-Fa-f]{6}$/.test(value) ? value.toUpperCase() : undefined;
}

function localName(element: XmlElement): string {
  return element.local ?? element.name.slice(element.name.indexOf(":") + 1);
}

function elementChildren(element: XmlElement): XmlElement[] {
  return element.children.filter((child: XmlNode): child is XmlElement => child.type === "element");
}

function isDrawingMl(element: XmlElement, name: string): boolean {
  return element.uri === DRAWINGML_NS && localName(element) === name;
}

function drawingMlChild(element: XmlElement | undefined, name: string): XmlElement | undefined {
  return element && elementChildren(element).find(child => isDrawingMl(child, name));
}
