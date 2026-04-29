/**
 * Picture-fill (`<a:blipFill>`) registration for chart parts.
 *
 * Callers supply an image payload via `AddChartSeriesOptions.pictureFill.image`;
 * the builder stores it on the chart model's `spPr.fill.blip._pendingImage`
 * field. This module consumes those staged payloads at chart-registration
 * time, pushes the image into the workbook media collection, allocates a
 * fresh chart-part relationship, and rewrites the blip with the resolved
 * `r:embed` id.
 *
 * The worksheet path calls {@link resolvePendingChartImages} right after
 * `fillChartCaches` and before the chart entry is stored — at that point
 * we have both the built model and the owning workbook, so we can do the
 * binding without requiring callers to juggle rel ids manually.
 */

import type { ChartEntry } from "./chart.ts";
import type {
  AddChartPictureFillImage,
  ChartBlipFill,
  ChartFill,
  ChartModel,
  ChartPictureFillImageData,
  ChartTypeGroup,
  SeriesBase,
  ShapeProperties
} from "./types.ts";

/**
 * Minimal workbook surface used by {@link resolvePendingChartImages}. The
 * concrete worksheet/workbook types pull in Cell/Row/Column machinery
 * this helper has no business touching, so we work against a structural
 * slice instead to keep the chart module import-light.
 */
export interface PictureFillHostWorkbook {
  addImage(image: ImageData): number;
  getImage(id: number | string): WorkbookMediaLike | undefined;
}

interface WorkbookMediaLike {
  extension?: string;
  buffer?: Uint8Array | undefined;
  base64?: string | undefined;
  filename?: string | undefined;
}

interface ImageData {
  extension: "jpeg" | "png" | "gif";
  base64?: string;
  filename?: string;
  buffer?: Uint8Array;
}

/**
 * Rel information that {@link resolvePendingChartImages} appends to the
 * chart entry so the writer can emit the `chart{N}.xml.rels` file and
 * the correct `r:embed` value inside the chart XML.
 *
 * The concrete shape matches the rel entry format used elsewhere
 * (xlsx-writer relationships XForm) — Id + Type + Target triple. We
 * hardcode the Type URI in the literal emitted at the call site so
 * callers don't need to import the workbook-level `RelType` enum.
 */

/**
 * Walk a chart model, resolve every `_pendingImage` blip into a real
 * workbook media entry + chart relationship, and mutate the model so
 * the blip references the allocated `rIdN`.
 *
 * Safe to call multiple times on the same entry: once a blip has a
 * `relationshipId` and no `_pendingImage`, it is skipped.
 *
 * @param entry   The chart entry being registered. Its `rels` array is
 *                populated (created if absent) with any freshly
 *                allocated image relationships.
 * @param workbook Host workbook exposing `addImage`.
 */
export function resolvePendingChartImages(
  entry: ChartEntry,
  workbook: PictureFillHostWorkbook
): void {
  const blips = collectBlipFills(entry.model);
  if (blips.length === 0) {
    return;
  }
  const rels = (entry.rels ??= []);
  // Scan pre-existing rels so freshly allocated ids don't collide with
  // rels the user may already have set (e.g. from a partial migration).
  const usedIds = new Set<string>();
  for (const rel of rels) {
    if (rel?.Id) {
      usedIds.add(rel.Id);
    }
  }
  let counter = 1;
  const allocRelId = (): string => {
    let id = `rId${counter++}`;
    while (usedIds.has(id)) {
      id = `rId${counter++}`;
    }
    usedIds.add(id);
    return id;
  };

  for (const blip of blips) {
    if (blip.relationshipId || blip._pendingImage === undefined) {
      delete blip._pendingImage;
      continue;
    }
    const normalised = normaliseImage(blip._pendingImage, workbook);
    if (!normalised) {
      // Bad payload — drop the blip so we don't emit a broken rel. The
      // series keeps its `pictureOptions` so the chart still opens.
      delete blip._pendingImage;
      continue;
    }
    const { mediaId, extension } = normalised;
    const relId = allocRelId();
    blip.relationshipId = relId;
    delete blip._pendingImage;
    rels.push({
      Id: relId,
      Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      // Chart parts live at `xl/charts/chart{N}.xml`; images live at
      // `xl/media/image{M}.{ext}`. The relative Target is therefore
      // `../media/image{M}.{ext}`. Media ids are 0-based in the workbook
      // collection but the on-disk naming is 1-based, hence `mediaId + 1`.
      Target: `../media/image${mediaId + 1}.${extension}`
    });
  }
}

/**
 * Traverse every `spPr` slot in a chart model and return the list of
 * {@link ChartBlipFill} objects that carry picture fills.
 *
 * The traversal is intentionally narrow: we only visit the locations
 * where `a:blipFill` is actually legal (series, data points, chart
 * frame, plot area, walls, axis, legend, gridlines, drop/hi-low lines,
 * trendline labels…). It walks both the primary (`chart.plotArea`) and
 * pivot-format (`chart.pivotFormats`) branches.
 */
function collectBlipFills(model: ChartModel): ChartBlipFill[] {
  const out: ChartBlipFill[] = [];
  const visitSpPr = (spPr: ShapeProperties | undefined): void => {
    const blip = (spPr?.fill as ChartFill | undefined)?.blip;
    if (blip && (blip.relationshipId || blip._pendingImage)) {
      out.push(blip);
    }
  };
  visitSpPr(model.spPr);
  visitSpPr(model.chart?.floor);
  visitSpPr(model.chart?.sideWall);
  visitSpPr(model.chart?.backWall);
  visitSpPr(model.chart?.plotArea?.spPr);
  for (const group of model.chart?.plotArea?.chartTypes ?? []) {
    visitGroup(group, visitSpPr);
  }
  for (const axis of model.chart?.plotArea?.axes ?? []) {
    visitSpPr(axis.spPr);
    // `majorGridlines` and `minorGridlines` are themselves `ShapeProperties`
    // (the grid-line style uses the same DrawingML fill/line/effect
    // vocabulary as any other shape); pass them through directly rather
    // than treating them as containers.
    visitSpPr(axis.majorGridlines);
    visitSpPr(axis.minorGridlines);
  }
  visitSpPr(model.chart?.legend?.spPr);
  for (const pf of model.chart?.pivotFormats ?? []) {
    visitSpPr(pf.spPr);
  }
  return out;
}

function visitGroup(
  group: ChartTypeGroup,
  visit: (spPr: ShapeProperties | undefined) => void
): void {
  const groupWithSeries = group as {
    series?: SeriesBase[];
    dropLines?: ShapeProperties;
    hiLowLines?: ShapeProperties;
    serLines?: ShapeProperties;
    upDownBars?: { upBars?: ShapeProperties; downBars?: ShapeProperties };
  };
  if (groupWithSeries.series) {
    for (const s of groupWithSeries.series) {
      visit(s.spPr);
      // Per-point overrides live on specific series subtypes (BarSeries,
      // LineSeries, …) as `dataPoints?: DataPoint[]`. They all share the
      // same shape — walk them generically without narrowing to each
      // discriminated union variant.
      const dps = (s as { dataPoints?: Array<{ spPr?: ShapeProperties }> }).dataPoints;
      if (dps) {
        for (const dp of dps) {
          visit(dp.spPr);
        }
      }
    }
  }
  visit(groupWithSeries.dropLines);
  visit(groupWithSeries.hiLowLines);
  visit(groupWithSeries.serLines);
  visit(groupWithSeries.upDownBars?.upBars);
  visit(groupWithSeries.upDownBars?.downBars);
}

/**
 * Coerce a user-facing {@link AddChartPictureFillImage} into a workbook
 * media entry. Returns `{mediaId, extension}` or `undefined` when the
 * payload is unusable (empty / unknown format).
 *
 * Side-effect: when the caller passes raw data (not `workbookImageId`),
 * a new media entry is created on the workbook.
 */
function normaliseImage(
  image: AddChartPictureFillImage,
  workbook: PictureFillHostWorkbook
): { mediaId: number; extension: "png" | "jpeg" | "gif" } | undefined {
  // Re-use an already-registered workbook image.
  if (typeof image === "object" && !Array.isArray(image) && "workbookImageId" in image) {
    const existing = workbook.getImage(image.workbookImageId);
    if (!existing) {
      return undefined;
    }
    const ext = (existing.extension ?? "png") as "png" | "jpeg" | "gif";
    return { mediaId: image.workbookImageId, extension: ext };
  }

  // Raw binary payload.
  if (image instanceof Uint8Array) {
    const ext = sniffImageExtension(image);
    const id = workbook.addImage({ extension: ext, buffer: image });
    return { mediaId: id, extension: ext };
  }

  // String — either a data URL or a bare base64 string.
  if (typeof image === "string") {
    const parsed = parseDataUrlOrBase64(image);
    if (!parsed) {
      return undefined;
    }
    const id = workbook.addImage({ extension: parsed.extension, base64: parsed.base64 });
    return { mediaId: id, extension: parsed.extension };
  }

  // Structured ChartPictureFillImageData.
  const data = image as ChartPictureFillImageData;
  if (data.buffer instanceof Uint8Array) {
    const id = workbook.addImage({ extension: data.extension, buffer: data.buffer });
    return { mediaId: id, extension: data.extension };
  }
  if (typeof data.base64 === "string" && data.base64.length > 0) {
    const id = workbook.addImage({ extension: data.extension, base64: data.base64 });
    return { mediaId: id, extension: data.extension };
  }
  return undefined;
}

/**
 * Parse a `data:image/<type>;base64,…` URL or a bare base64 string into
 * `{ extension, base64 }`. Returns `undefined` for empty or unrecognised
 * input. Unknown data-URL content types default to PNG.
 */
function parseDataUrlOrBase64(
  input: string
): { extension: "png" | "jpeg" | "gif"; base64: string } | undefined {
  if (!input) {
    return undefined;
  }
  const dataUrl = /^data:image\/([a-z]+)(?:;[^,]*)?;base64,(.+)$/i.exec(input.trim());
  if (dataUrl) {
    const rawType = dataUrl[1].toLowerCase();
    const ext =
      rawType === "jpg"
        ? "jpeg"
        : rawType === "png" || rawType === "jpeg" || rawType === "gif"
          ? (rawType as "png" | "jpeg" | "gif")
          : "png";
    return { extension: ext, base64: dataUrl[2] };
  }
  // Bare base64 — extension unknown, default to PNG (the most common
  // Excel bar-fill choice). Callers who need precision should pass a
  // structured ChartPictureFillImageData.
  return { extension: "png", base64: input };
}

/**
 * Identify the image format from the first few magic bytes.
 *
 * Supports the three extensions Excel accepts in `<a:blipFill>` — PNG,
 * JPEG, GIF. Anything else falls back to PNG (which is the safest
 * default: Excel will try to decode and most viewers treat unrecognised
 * binary as PNG).
 */
function sniffImageExtension(data: Uint8Array): "png" | "jpeg" | "gif" {
  if (
    data.length >= 8 &&
    data[0] === 0x89 &&
    data[1] === 0x50 &&
    data[2] === 0x4e &&
    data[3] === 0x47 &&
    data[4] === 0x0d &&
    data[5] === 0x0a &&
    data[6] === 0x1a &&
    data[7] === 0x0a
  ) {
    return "png";
  }
  if (data.length >= 3 && data[0] === 0xff && data[1] === 0xd8 && data[2] === 0xff) {
    return "jpeg";
  }
  if (
    data.length >= 6 &&
    data[0] === 0x47 &&
    data[1] === 0x49 &&
    data[2] === 0x46 &&
    data[3] === 0x38 &&
    (data[4] === 0x37 || data[4] === 0x39) &&
    data[5] === 0x61
  ) {
    return "gif";
  }
  return "png";
}
