/**
 * Unified rendering context for DOCX XML serialization.
 * Passed to all writer functions to provide shared state.
 */

import type { PartName } from "../core/opc-package";
import type { WordSecurityPolicy } from "../security/policy";
import { DEFAULT_SECURITY_POLICY } from "../security/policy";

/** ID generators for various document elements. */
export interface IdGenerators {
  nextDrawingId(): number;
  nextSdtId(): number;
  nextBookmarkId(): number;
  nextDocPrId(): number;
  nextChartId(): number;
  nextImagePartId(): number;
}

/**
 * Lightweight, content-renderer-facing subset of WordRenderContext.
 *
 * Most leaf renderers (run/paragraph/table/image) only need the rId remap and
 * the hyperlink rId map; passing the full WordRenderContext would couple them
 * to chartRIds/securityPolicy/etc. They take this `RenderHelpers` instead.
 */
export interface RenderHelpers {
  /** See WordRenderContext.imageRIdRemap. */
  readonly imageRemap?: ReadonlyMap<string, string>;
  /** See WordRenderContext.hyperlinkRIds. */
  readonly hyperlinkRIds?: ReadonlyWeakMap<object, string>;
}

/**
 * Read-only WeakMap projection.
 * (TS lib doesn't ship one; we mirror the readonly subset of WeakMap.)
 */
interface ReadonlyWeakMap<K extends WeakKey, V> {
  get(key: K): V | undefined;
  has(key: K): boolean;
}

/** Render context passed to all writer functions. */
export interface WordRenderContext {
  /** Current part being rendered. */
  readonly partName: PartName;
  /** Security policy controlling raw XML, external targets, etc. */
  readonly securityPolicy: WordSecurityPolicy;
  /** ID generators (shared across all parts for a single document render). */
  readonly ids: IdGenerators;
  /** Raw XML output policy. If "reject", opaque content throws instead of being written. */
  readonly rawXmlPolicy: "preserve" | "strip" | "reject";
  /** Chart/ChartEx object → rId mapping. Populated by packager before rendering. */
  readonly chartRIds: Map<object, string>;
  /**
   * Optional remap for image relationship IDs that the packager had to rewrite
   * (e.g. when the model's image rId collided with a non-image relationship in
   * the same .rels file). Writers that emit `r:embed` look up this map first
   * and fall back to the original rId stored on the model. The map is keyed by
   * the original (model-side) rId.
   */
  readonly imageRIdRemap: Map<string, string>;
  /**
   * External hyperlink object → rId mapping. The packager registers a
   * relationship for each hyperlink with a `url` and stores the assigned rId
   * here, keyed by the hyperlink object identity. The paragraph writer reads
   * this when emitting `<w:hyperlink r:id="...">` so the model itself is
   * never mutated.
   */
  readonly hyperlinkRIds: WeakMap<object, string>;
}

/** Create ID generators with given starting values. */
export function createIdGenerators(startValues?: {
  drawingId?: number;
  sdtId?: number;
  bookmarkId?: number;
  docPrId?: number;
  chartId?: number;
  imagePartId?: number;
}): IdGenerators {
  let drawingId = startValues?.drawingId ?? 1;
  let sdtId = startValues?.sdtId ?? 1;
  let bookmarkId = startValues?.bookmarkId ?? 0;
  let docPrId = startValues?.docPrId ?? 1;
  let chartId = startValues?.chartId ?? 1;
  let imagePartId = startValues?.imagePartId ?? 1;

  return {
    nextDrawingId: () => drawingId++,
    nextSdtId: () => sdtId++,
    nextBookmarkId: () => bookmarkId++,
    nextDocPrId: () => docPrId++,
    nextChartId: () => chartId++,
    nextImagePartId: () => imagePartId++
  };
}

/** Create a default render context. */
export function createRenderContext(options?: {
  partName?: PartName;
  securityPolicy?: WordSecurityPolicy;
  chartRIds?: Map<object, string>;
  imageRIdRemap?: Map<string, string>;
  hyperlinkRIds?: WeakMap<object, string>;
}): WordRenderContext {
  const securityPolicy = options?.securityPolicy ?? DEFAULT_SECURITY_POLICY;

  return {
    partName: options?.partName ?? "/word/document.xml",
    securityPolicy,
    ids: createIdGenerators(),
    rawXmlPolicy: securityPolicy.rawXmlPolicy ?? "preserve",
    chartRIds: options?.chartRIds ?? new Map(),
    imageRIdRemap: options?.imageRIdRemap ?? new Map(),
    hyperlinkRIds: options?.hyperlinkRIds ?? new WeakMap()
  };
}
