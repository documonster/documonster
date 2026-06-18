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

import type {
  AddChartPictureFillImage,
  ChartBlipFill,
  ChartEntry,
  ChartFill,
  ChartModel,
  ChartPictureFillImageData,
  ChartTypeGroup,
  SeriesBase,
  ShapeProperties
} from "@excel/chart/model/types";
import { type WorkbookData, addWorkbookImage, getImage } from "@excel/workbook-core";

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
 * @param workbook Host workbook exposing `addImage`. Also consulted for
 *                its internal `_chartRels[n]` bag — those rels are
 *                loaded from disk during round-trip and carry canonical
 *                rIds that must NOT collide with freshly allocated
 *                image rels (the xlsx writer merges both lists and
 *                silently drops any duplicate Id, which would otherwise
 *                orphan the image relationship).
 * @param chartNumber 1-based chart index. Used to look up
 *                `workbook._chartRels[chartNumber]` so pre-existing
 *                rels seed the collision set.
 */
export function resolvePendingChartImages(
  entry: ChartEntry,
  workbook: WorkbookData,
  chartNumber: number
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
  // Also seed from the workbook-level `_chartRels` bag. Those entries
  // come from the xlsx reader (round-tripped rels: style / colors /
  // userShapes / existing images) and are merged BEFORE `entry.rels`
  // by the writer — the writer drops any `entry.rels` entry whose Id
  // collides. Without seeding here, a fresh image rel allocated as
  // `rId1` would be silently dropped by the writer on a round-tripped
  // chart whose original style sidecar already took `rId1`, leaving
  // the chart's `r:embed="rId1"` pointing at the wrong target.
  const chartRelsBag = (
    workbook as unknown as {
      _chartRels?: Record<number, Array<{ Id?: string }>>;
    }
  )._chartRels;
  const existingChartRels = chartRelsBag?.[chartNumber];
  if (Array.isArray(existingChartRels)) {
    for (const rel of existingChartRels) {
      if (rel?.Id) {
        usedIds.add(rel.Id);
      }
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
  visitSpPr(model.chart?.title?.spPr);
  visitSpPr(model.chart?.floor);
  visitSpPr(model.chart?.sideWall);
  visitSpPr(model.chart?.backWall);
  visitSpPr(model.chart?.plotArea?.spPr);
  for (const group of model.chart?.plotArea?.chartTypes ?? []) {
    visitGroup(group, visitSpPr);
  }
  for (const axis of model.chart?.plotArea?.axes ?? []) {
    visitSpPr(axis.spPr);
    visitSpPr(axis.title?.spPr);
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
      // Series-level data labels and trendlines are DrawingML shapes too;
      // both can carry picture fills (e.g. a branded label background).
      const dataLabels = (s as { dataLabels?: { spPr?: ShapeProperties } }).dataLabels;
      visit(dataLabels?.spPr);
      const trendlines = (s as { trendlines?: Array<{ spPr?: ShapeProperties }> }).trendlines;
      if (trendlines) {
        for (const tl of trendlines) {
          visit(tl.spPr);
        }
      }
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
  workbook: WorkbookData
): { mediaId: number; extension: "png" | "jpeg" | "gif" } | undefined {
  // Re-use an already-registered workbook image.
  if (typeof image === "object" && !Array.isArray(image) && "workbookImageId" in image) {
    const existing = getImage(workbook, image.workbookImageId);
    if (!existing) {
      return undefined;
    }
    // DrawingML's `<a:blipFill>` technically accepts more formats (BMP,
    // TIFF, EMF/WMF) but the workbook media pipeline and our
    // magic-byte sniffer only cover PNG / JPEG / GIF. Silently
    // dropping the blip is preferable to emitting a rel entry whose
    // media bytes would be mislabelled and fail to decode in Excel.
    // Callers needing the exotic formats should add them to the
    // workbook media collection separately and use `workbookImageId`.
    const raw = existing.extension;
    if (raw !== "png" && raw !== "jpeg" && raw !== "gif") {
      return undefined;
    }
    return { mediaId: image.workbookImageId, extension: raw };
  }

  // Raw binary payload. `instanceof Uint8Array` is realm-sensitive:
  // a buffer produced inside a Web Worker or another iframe has a
  // different `Uint8Array` prototype, so the operator returns false
  // even though the object is a genuine byte array. Duck-typing via
  // `ArrayBuffer.isView` + `BYTES_PER_ELEMENT === 1` covers both the
  // same-realm path and cross-realm buffers (workers, SharedWorkers,
  // MessagePort transfers, `structuredClone` round-trips), and still
  // excludes `Uint16Array` / `DataView` etc. — `DataView` has
  // no `BYTES_PER_ELEMENT` on instances (the property is on the
  // constructor), so the check rejects it naturally. `Int8Array` /
  // `Uint8ClampedArray` DO pass the check; that's intentional —
  // both carry byte-granular binary data and the sniffer below
  // operates on the first few bytes regardless of typed-array
  // signedness. The workbook image store only needs a
  // `Uint8Array`, so we coerce through a shared underlying buffer.
  if (
    ArrayBuffer.isView(image) &&
    (image as unknown as { BYTES_PER_ELEMENT?: number }).BYTES_PER_ELEMENT === 1
  ) {
    const view = image as ArrayBufferView;
    const bytes =
      view instanceof Uint8Array
        ? view
        : new Uint8Array(view.buffer, view.byteOffset, view.byteLength);
    const ext = sniffImageExtension(bytes);
    if (!ext) {
      // Unknown format (WebP / AVIF / TIFF / …) — drop the blip rather
      // than mislabel it as `png`. The caller keeps `pictureOptions`
      // so the rest of the chart still renders.
      return undefined;
    }
    const id = addWorkbookImage(workbook, { extension: ext, buffer: bytes });
    return { mediaId: id, extension: ext };
  }

  // String — either a data URL or a bare base64 string.
  if (typeof image === "string") {
    const parsed = parseDataUrlOrBase64(image);
    if (!parsed) {
      return undefined;
    }
    const id = addWorkbookImage(workbook, { extension: parsed.extension, base64: parsed.base64 });
    return { mediaId: id, extension: parsed.extension };
  }

  // Structured ChartPictureFillImageData.
  const data = image as ChartPictureFillImageData;
  if (data.buffer != null) {
    // Use ArrayBuffer.isView to handle cross-realm Uint8Array instances
    // that fail `instanceof` checks but are structurally identical.
    const raw = data.buffer as unknown;
    if (ArrayBuffer.isView(raw)) {
      const view = raw as ArrayBufferView;
      const bytes =
        view instanceof Uint8Array
          ? view
          : new Uint8Array(view.buffer, view.byteOffset, view.byteLength);
      const id = addWorkbookImage(workbook, { extension: data.extension, buffer: bytes });
      return { mediaId: id, extension: data.extension };
    }
  }
  if (typeof data.base64 === "string" && data.base64.length > 0) {
    const id = addWorkbookImage(workbook, { extension: data.extension, base64: data.base64 });
    return { mediaId: id, extension: data.extension };
  }
  return undefined;
}

/**
 * Parse a `data:image/<type>;base64,…` URL or a bare base64 string into
 * `{ extension, base64 }`. Returns `undefined` for empty or
 * unrecognised input (including data URLs whose content-type is not one
 * of the three `<a:blipFill>`-supported formats — e.g. `image/svg+xml`
 * or `image/bmp`). Callers interpret `undefined` as "drop the blip"
 * rather than emitting a corrupted `.png` file.
 */
function parseDataUrlOrBase64(
  input: string
): { extension: "png" | "jpeg" | "gif"; base64: string } | undefined {
  if (!input) {
    return undefined;
  }
  // `data:image/<subtype>[;param=value]*;base64,<payload>`. Support
  // multiple parameter segments (e.g. `;charset=utf-8;base64,...`).
  // The previous `(?:;[^,]*)?` admitted only a single parameter AND
  // used a greedy class that could swallow the required `;base64`
  // token, forcing backtracking that silently failed for common
  // multi-param forms. Allow any number of `;token[=value]` groups
  // before the required trailing `;base64,`.
  const dataUrl = /^data:image\/([a-z+]+)(?:;[^,;=]+(?:=[^,;]*)?)*;base64,(.+)$/i.exec(
    input.trim()
  );
  if (dataUrl) {
    const rawType = dataUrl[1].toLowerCase();
    const ext =
      rawType === "jpg"
        ? "jpeg"
        : rawType === "png" || rawType === "jpeg" || rawType === "gif"
          ? (rawType as "png" | "jpeg" | "gif")
          : undefined;
    if (ext === undefined) {
      // Content-type outside the `<a:blipFill>` white-list (e.g.
      // `svg+xml`, `bmp`, `webp`, `tiff`). Dropping is better than
      // silently relabelling an SVG payload as `image/png` — Excel
      // would try to decode SVG bytes with the PNG parser and either
      // error out or leave an invisible shape.
      return undefined;
    }
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
/**
 * Inspect the first few bytes of `data` and return the matching
 * `<a:blipFill>`-supported image extension, or `undefined` when the
 * payload doesn't look like one of the three supported formats
 * (PNG / JPEG / GIF). Callers treat `undefined` as "drop the blip" —
 * preferable to emitting a `.png` file whose decoded bytes are
 * actually WebP / AVIF / TIFF, which Excel would fail to render.
 */
function sniffImageExtension(data: Uint8Array): "png" | "jpeg" | "gif" | undefined {
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
  return undefined;
}
