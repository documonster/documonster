/**
 * The reader's last step: join the parts collected while walking the package — relationships, external
 * links, pivot caches, chart user shapes, preserved opaque parts — into the model handed to the workbook.
 */
import type { ChartEntry } from "@excel/chart/model/types";
import type {
  PivotTable,
  PivotTableSubtotal,
  ParsedCacheDefinition
} from "@excel/core/pivot-table";
import type { ExternalLinkModel } from "@excel/core/workbook.browser";
import { pivotTablePathFromName, worksheetPath } from "@excel/utils/ooxml-paths";
import { collectOpaqueParts, groupOpaqueRelationshipsBySource } from "@excel/xlsx/opaque-parts";
import { RelType } from "@excel/xlsx/rel-type";
import type { XlsxOptions } from "@excel/xlsx/types";
import type { ParsedExternalLink } from "@excel/xlsx/xform/book/external-link-xform";
import { WorkbookXform } from "@excel/xlsx/xform/book/workbook-xform";
import type { ParsedPivotTableModel } from "@excel/xlsx/xform/pivot-table/pivot-table-xform";
import { WorkSheetXform } from "@excel/xlsx/xform/sheet/worksheet-xform";

/**
 * Extract the trailing integer from a workbook-rels Target like
 * `"externalLinks/externalLink12.xml"` (a path relative to `xl/`). Mirror
 * of {@link getExternalLinkIndexFromPath} which takes the full
 * `xl/externalLinks/…` form. Used during reconcile to bridge the
 * workbook.xml.rels entry to the parsed externalLinkN.xml part.
 */
export function externalLinkIndexFromRelTarget(target: string): number | undefined {
  const match = /(?:^|\/)externalLink(\d+)[.]xml$/.exec(target);
  return match ? parseInt(match[1], 10) : undefined;
}

/**
 * Shape of a parsed externalLinkN.xml.rels entry. Kept broad so any
 * relationship type passes through verbatim — the writer only cares about
 * the ExternalLinkPath entry, but we preserve the rest for round-trip.
 */
export type ExternalLinkRelsEntry = {
  Id: string;
  Type: string;
  Target: string;
  TargetMode?: string;
};

export async function reconcile(model: any, options?: XlsxOptions): Promise<void> {
  // Join the opaque staging maps first: everything else in reconcile works on
  // modelled state, and the preserved parts must be settled before the model
  // is handed to the workbook.
  const opaque = collectOpaqueParts({
    unknownEntries: model.opaqueUnknownEntries,
    contentTypeOverrides: model.opaqueContentTypeOverrides,
    relationshipsBySource: model.opaqueRelationshipsBySource
  });
  model.opaqueParts = opaque.parts;
  model.opaqueDrops = opaque.drops;
  // Distribute sheet-sourced inbound relationships onto the sheets that
  // declared them, so they move with the sheet rather than with its position.
  const inboundBySource = groupOpaqueRelationshipsBySource(opaque.parts);
  for (const worksheet of model.worksheets) {
    const rels = inboundBySource.get(worksheetPath(worksheet.sheetNo));
    if (rels) {
      worksheet.opaqueRels = rels;
    }
  }
  // The staging maps hold every unrecognised entry's bytes. They are load-time
  // scratch, and the assembled parts now own that memory, so release them
  // rather than carry a second reference to the whole set into the workbook.
  delete model.opaqueUnknownEntries;
  delete model.opaqueContentTypeOverrides;
  delete model.opaqueRelationshipsBySource;

  const workbookXform = new WorkbookXform();
  const worksheetXform = new WorkSheetXform(options);

  workbookXform.reconcile(model);

  // reconcile drawings with their rels — DrawingXform (~34 KB) is loaded
  // lazily so workbooks without drawings never pull it into the bundle.
  const drawingNames = Object.keys(model.drawings);
  if (drawingNames.length > 0) {
    const { DrawingXform } = await import("@excel/xlsx/xform/drawing/drawing-xform");
    const drawingXform = new DrawingXform();
    const drawingOptions: any = {
      media: model.media,
      mediaIndex: model.mediaIndex
    };
    drawingNames.forEach(name => {
      const drawing = model.drawings[name];
      const drawingRel = model.drawingRels[name];
      if (drawingRel) {
        drawingOptions.rels = drawingRel.reduce((o: any, rel: any) => {
          o[rel.Id] = rel;
          return o;
        }, {});
        (drawing.anchors ?? []).forEach((anchor: any) => {
          const hyperlinks = anchor.picture && anchor.picture.hyperlinks;
          if (hyperlinks && drawingOptions.rels[hyperlinks.rId]) {
            hyperlinks.hyperlink = drawingOptions.rels[hyperlinks.rId].Target;
            delete hyperlinks.rId;
          }
        });
        drawingXform.reconcile(drawing, drawingOptions);
      }
    });
  }

  // Reconcile chart references in drawing anchors
  Object.keys(model.drawings).forEach(name => {
    const drawing = model.drawings[name];
    const drawingRel = model.drawingRels[name];
    if (!drawingRel) {
      return;
    }
    const relMap: Record<string, any> = {};
    for (const rel of drawingRel) {
      relMap[rel.Id] = rel;
    }
    for (const anchor of drawing.anchors ?? []) {
      if (anchor.graphicFrame?.rId) {
        const rel = relMap[anchor.graphicFrame.rId];
        if (rel?.Target) {
          // Extract chart number from target like "../charts/chart1.xml"
          const match = /chart(\d+)\.xml/.exec(rel.Target);
          if (match) {
            anchor.chartNumber = parseInt(match[1], 10);
          }
          // Extract chartEx number from target like "../charts/chartEx1.xml"
          const matchEx = /chartEx(\d+)\.xml/.exec(rel.Target);
          if (matchEx) {
            anchor.chartExNumber = parseInt(matchEx[1], 10);
          }
        }
      }
    }
  });

  // reconcile tables with the default styles — TableXform (~14 KB) loaded
  // lazily so table-free workbooks don't pull it in.
  const tables = Object.values(model.tables);
  if (tables.length > 0) {
    const { TableXform } = await import("@excel/xlsx/xform/table/table-xform");
    const tableXform = new TableXform();
    const tableOptions = {
      styles: model.styles
    };
    tables.forEach((table: any) => {
      tableXform.reconcile(table, tableOptions);
    });
  }

  // Reconcile pivot tables
  reconcilePivotTables(model);

  const sheetOptions = {
    styles: model.styles,
    sharedStrings: model.sharedStrings,
    media: model.media,
    mediaIndex: model.mediaIndex,
    date1904: model.properties?.date1904,
    drawings: model.drawings,
    drawingRels: model.drawingRels,
    comments: model.comments,
    tables: model.tables,
    vmlDrawings: model.vmlDrawings,
    vmlDrawingHF: model.vmlDrawingHF,
    vmlDrawingHFRels: model.vmlDrawingHFRels,
    pivotTables: model.pivotTablesIndexed,
    hasDynamicArrayMetadata: !!model.metadata?.hasDynamicArrays,
    dynamicArrayCmIndices: model.metadata?.dynamicArrayCmIndices as Set<number> | undefined
  };
  model.worksheets.forEach((worksheet: any) => {
    worksheet.relationships = model.worksheetRels[worksheet.sheetNo];
    worksheetXform.reconcile(worksheet, sheetOptions);
    // Attach any threaded comments that arrived in a separate
    // `xl/threadedComments/threadedComment{N}.xml` part. The sheet
    // index in that path maps to `worksheet.sheetNo`, not
    // `worksheet.id` — Excel uses the package-relative file number,
    // same as classic `xl/comments{N}.xml`.
    const threaded = model.threadedCommentsByIndex?.[worksheet.sheetNo];
    if (threaded) {
      worksheet.threadedComments = threaded;
    }
  });

  // Reconcile chartsheets — link their drawing references and
  // preserve every relationship so the writer can round-trip
  // every r:id referenced by raw-captured children (legacyDrawing,
  // picture, legacyDrawingHF, drawingHF, etc.). Previously only
  // the drawing rel was hooked up and everything else was
  // silently discarded on save, leaving any raw-captured child
  // with a dangling r:id pointing at a now-missing part.
  const chartsheetsList = model.chartsheetsList || [];
  for (const cs of chartsheetsList) {
    const csRels = model.chartsheetRels[cs.sheetNo];
    if (csRels) {
      // Keep the full rels list attached to the model so
      // `addChartsheets` can re-emit it. Copy so downstream
      // mutations don't leak back into `model.chartsheetRels`.
      cs.relationships = [...csRels];
    }
    if (cs.drawing && csRels) {
      const drawingRel = csRels.find((r: any) => r.Id === cs.drawing.rId);
      if (drawingRel) {
        const match = drawingRel.Target.match(/\/drawings\/([a-zA-Z0-9]+)[.][a-zA-Z]{3,4}$/);
        if (match) {
          cs.drawingName = match[1];
          // Resolve drawing → chart number from drawing rels
          const drawingRelArr = model.drawingRels[cs.drawingName];
          if (drawingRelArr) {
            for (const dr of drawingRelArr) {
              const chartMatch = /chart(\d+)\.xml/.exec(dr.Target);
              if (chartMatch) {
                cs.chartNumber = parseInt(chartMatch[1], 10);
                break;
              }
              const chartExMatch = /chartEx(\d+)\.xml/.exec(dr.Target);
              if (chartExMatch) {
                cs.chartExNumber = parseInt(chartExMatch[1], 10);
                break;
              }
            }
          }
        }
      }
    }
    if (cs.legacyDrawingHF && csRels) {
      const vmlRel = csRels.find((r: any) => r.Id === cs.legacyDrawingHF.rId);
      const match = vmlRel?.Target && /(vmlDrawingHF\d+)[.]vml$/.exec(vmlRel.Target);
      const name = match?.[1];
      const images = name && model.vmlDrawingHF?.[name];
      const imageRels = name && model.vmlDrawingHFRels?.[name];
      if (images && imageRels) {
        const relMap = Object.fromEntries(imageRels.map((rel: any) => [rel.Id, rel]));
        cs.headerImages = images.flatMap((image: any) => {
          const target = relMap[image.imageRelId]?.Target?.split("/media/")[1];
          const imageId = target && model.mediaIndex?.[target];
          return imageId !== undefined && /^(?:LH|CH|RH|LF|CF|RF)$/.test(image.position ?? "")
            ? [
                {
                  imageId,
                  width: image.width,
                  height: image.height,
                  position: image.position
                }
              ]
            : [];
        });
      }
      if (!cs.headerImages?.length) {
        // The source relationship cannot be reproduced safely (missing VML
        // rels/media or malformed positions). Drop both ends instead of
        // emitting a chartsheet relationship to a non-existent part.
        cs.relationships = cs.relationships?.filter(
          (rel: any) => rel.Id !== cs.legacyDrawingHF.rId
        );
        cs.legacyDrawingHF = undefined;
      }
    }
  }
  model.chartsheets = chartsheetsList;

  // Reconcile external workbook links before workbookRels / externalReferences
  // are dropped. Joins 3 sources:
  //   1. model.externalReferences  — ordered list of { rId } from workbook.xml
  //   2. model.workbookRels        — maps rId → target path (inside xl/)
  //   3. model.externalLinksByIndex — parsed externalLinkN.xml parts
  //   4. model.externalLinkRelsByIndex — parsed externalLinkN.xml.rels parts
  reconcileExternalLinks(model);

  // Preserve parsed chart data through to the workbook model.
  // chartEntries, chartRels, chartStyles, chartColors are kept as-is.

  // Reconcile chart user-shapes drawing parts onto their owning
  // ChartEntry. Each chart rels file may reference an overlay drawing
  // via `RelType.ChartUserShapes`; we copy those bytes from
  // `model.drawingRaw` (populated by `processDrawingEntry`) onto the
  // chart entry so writers can emit them back, and so the Chart API
  // can expose them via `Chart.userShapesXml`. Regular worksheet
  // drawings are untouched — this reconcile only moves bytes for
  // chart-overlay parts.
  reconcileChartUserShapes(model);

  // delete unnecessary parts
  delete model.worksheetHash;
  delete model.worksheetRels;
  delete model.globalRels;
  delete model.sharedStrings;
  delete model.workbookRels;
  delete model.sheetDefs;
  // Preserve default font before deleting styles
  model.defaultFont = model.styles?.defaultFont;
  // Preserve named cell styles (OOXML cellStyles) for round-trip fidelity
  if (model.styles?.getNamedStyles) {
    const namedStyles = model.styles.getNamedStyles();
    if (namedStyles && namedStyles.length) {
      model.cellStyles = namedStyles;
    }
  }
  // The `<dxfs>` table, at its source indices. Rules and table columns reconciled above hold these very
  // objects; preserved XML (pivot `<formats>`, `<colorFilter>`) holds only the index. See `seedDxfs`.
  const dxfs = model.styles?.model?.dxfs;
  if (dxfs?.length) {
    model.dxfs = dxfs;
  }
  delete model.styles;
  delete model.mediaIndex;
  delete model.drawings;
  delete model.drawingRels;
  delete model.drawingRaw;
  delete model.vmlDrawings;
  delete model.pivotTableRels;
  delete model.metadata;
  // Internal-only scratch fields consumed by reconcileExternalLinks.
  delete model.externalReferences;
  delete model.externalLinksByIndex;
  delete model.externalLinkRelsByIndex;
  delete model.chartsheetRels;
  delete model.chartsheetsList;
}

/**
 * Copy the raw bytes of each chart's user-shapes drawing part onto
 * the owning `ChartEntry.userShapesXml` so the writer can emit them
 * back verbatim (and so {@link Chart.userShapesXml} can surface them
 * to user code). Runs after all ZIP entries have been processed
 * because chart rels and drawing bytes stream in independent order.
 *
 * Skips charts that have no `ChartUserShapes` rel. The bytes stay
 * keyed by drawing name (e.g. `drawing3`) inside `model.drawingRaw`
 * since a workbook may have many user-shape drawings across
 * different charts; we look up each chart's target through its
 * rels file.
 */
function reconcileChartUserShapes(model: any): void {
  const chartRelsMap = model.chartRels as Record<string, any[]> | undefined;
  const drawingRaw = model.drawingRaw as Record<string, Uint8Array> | undefined;
  const chartEntries = model.chartEntries as Record<string, ChartEntry> | undefined;
  if (!chartRelsMap || !drawingRaw || !chartEntries) {
    return;
  }
  for (const [chartNum, rels] of Object.entries(chartRelsMap)) {
    if (!Array.isArray(rels)) {
      continue;
    }
    const entry = chartEntries[chartNum];
    if (!entry) {
      continue;
    }
    const userShapesRel = rels.find(
      rel => rel && typeof rel === "object" && rel.Type === RelType.ChartUserShapes
    );
    if (!userShapesRel?.Target) {
      continue;
    }
    // Target like `../drawings/drawing3.xml` or `../drawings/chartUserShape2.xml`.
    const match = /drawings\/([^/]+)\.xml$/i.exec(String(userShapesRel.Target));
    if (!match) {
      continue;
    }
    const drawingName = match[1];
    const bytes = drawingRaw[drawingName];
    if (bytes) {
      entry.userShapesXml = bytes;
      // Make sure the chart model carries the r:id so subsequent reads
      // via Chart.userShapesXml can round-trip without extra setup.
      entry.model.userShapesRelId ??= userShapesRel.Id;
    }
  }
}

/**
 * Join the three on-disk sources that together describe external workbook
 * references into a single dense `model.externalLinks: ExternalLinkModel[]`.
 *
 * Sources:
 *   - `<externalReferences>` list in workbook.xml (declaration order)
 *   - `xl/_rels/workbook.xml.rels` (rId → internal path)
 *   - `xl/externalLinks/externalLink{N}.xml` (sheet names, cached values)
 *   - `xl/externalLinks/_rels/externalLink{N}.xml.rels` (target, TargetMode)
 *
 * The 1-based index of each resulting ExternalLinkModel matches the `[N]`
 * used in formula strings — this is the single source of truth formula
 * code should rely on.
 */
function reconcileExternalLinks(model: any): void {
  const refs = model.externalReferences as Array<{ rId: string }> | undefined;
  if (!refs || refs.length === 0) {
    // Even when workbook.xml has no <externalReferences>, we may still
    // have parsed externalLink parts (e.g. orphaned files); those are
    // dropped silently rather than generating synthesised indices.
    if (!model.externalLinks) {
      model.externalLinks = [];
    }
    return;
  }

  const rels = (model.workbookRels ?? []) as Array<{
    Id: string;
    Type: string;
    Target: string;
  }>;
  const relById = new Map<string, { Target: string }>();
  for (const rel of rels) {
    if (rel.Type === RelType.ExternalLink) {
      relById.set(rel.Id, rel);
    }
  }

  const externalLinks: ExternalLinkModel[] = [];
  for (let i = 0; i < refs.length; i++) {
    const ref = refs[i];
    const rel = relById.get(ref.rId);
    if (!rel) {
      // Broken reference — <externalReference> points at an rId that is
      // not of type ExternalLink. We skip silently; the formula engine
      // will see the now-missing index and fall back to #REF! as before.
      continue;
    }

    // The rel Target is a path inside xl/ like "externalLinks/externalLink1.xml".
    // Extract the trailing index to look up the parsed part.
    const partIndex = externalLinkIndexFromRelTarget(rel.Target);
    if (partIndex === undefined) {
      continue;
    }

    const parsed = model.externalLinksByIndex[partIndex] as ParsedExternalLink | undefined;
    const partRels = (model.externalLinkRelsByIndex[partIndex] ?? []) as ExternalLinkRelsEntry[];

    // Locate the externalLinkPath rel (should be unique within a part).
    const pathRel =
      partRels.find(r => r.Type === RelType.ExternalLinkPath) ??
      partRels.find(r => r.TargetMode === "External");

    externalLinks.push({
      index: i + 1,
      rId: ref.rId,
      target: pathRel?.Target ?? "",
      targetMode: (pathRel?.TargetMode as "External" | "Internal" | undefined) ?? "External",
      sheetNames: parsed?.sheetNames ?? [],
      cachedValues: parsed?.cachedValues ?? {}
    });
  }

  model.externalLinks = externalLinks;
}

/**
 * Reconcile pivot tables by linking them to worksheets and their cache data.
 */
function reconcilePivotTables(model: any): void {
  const rawPivotTables = (model.pivotTables || {}) as Record<string, ParsedPivotTableModel>;
  if (typeof rawPivotTables !== "object" || Object.keys(rawPivotTables).length === 0) {
    model.pivotTables = [];
    model.pivotTablesIndexed = {};
    return;
  }

  const definitionToCacheId = buildDefinitionToCacheIdMap(model);

  const cacheMap = new Map<
    number,
    {
      definition: ParsedCacheDefinition;
      records: any;
      definitionName: string;
    }
  >();

  Object.entries(model.pivotCacheDefinitions || {}).forEach(([name, definition]: [string, any]) => {
    const cacheId = definitionToCacheId.get(name);
    if (cacheId !== undefined) {
      const recordsName = name.replace("Definition", "Records");
      cacheMap.set(cacheId, {
        definition,
        records: model.pivotCacheRecords?.[recordsName],
        definitionName: name
      });
    }
  });

  const loadedPivotTables: PivotTable[] = [];
  const pivotTablesIndexed: Record<string, PivotTable> = {};

  Object.entries(rawPivotTables).forEach(([pivotName, pt]) => {
    const tableNumber = extractTableNumber(pivotName);
    const cacheData = cacheMap.get(pt.cacheId);

    const defaultMetric = determineMetric(pt.dataFields);
    const completePivotTable: PivotTable = {
      ...pt,
      name: pt.name ?? `PivotTable${tableNumber}`,
      tableNumber,
      cacheId: String(pt.cacheId),
      cacheDefinition: cacheData?.definition,
      cacheRecords: cacheData?.records,
      cacheFields: cacheData?.definition?.cacheFields ?? [],
      rows: pt.rowFields.filter(f => f >= 0),
      columns: pt.colFields.filter(f => f >= 0 && f !== -2),
      values: pt.dataFields.map(df => df.fld),
      pages: pt.pageFields.map(pf => pf.fld),
      metric: defaultMetric,
      valueMetrics: determineValueMetrics(pt.dataFields, defaultMetric),
      applyWidthHeightFormats: pt.applyWidthHeightFormats === "1" ? "1" : "0"
    };

    loadedPivotTables.push(completePivotTable);
    // Key by absolute zip path so reconcile can match any rel target layout.
    pivotTablesIndexed[pivotTablePathFromName(pivotName)] = completePivotTable;
  });

  loadedPivotTables.sort((a, b) => a.tableNumber - b.tableNumber);
  model.pivotTables = loadedPivotTables;
  model.pivotTablesIndexed = pivotTablesIndexed;
}

function extractTableNumber(name: string): number {
  const match = name.match(/pivotTable(\d+)/);
  return match ? parseInt(match[1], 10) : 1;
}

function buildCacheIdMap(model: any): Map<string, number> {
  const rIdToCacheId = new Map<string, number>();
  const pivotCaches = model.pivotCaches ?? [];
  for (const cache of pivotCaches) {
    if (cache.cacheId && cache.rId) {
      rIdToCacheId.set(cache.rId, parseInt(cache.cacheId, 10));
    }
  }
  return rIdToCacheId;
}

function buildDefinitionToCacheIdMap(model: any): Map<string, number> {
  const definitionToCacheId = new Map<string, number>();
  const rIdToCacheId = buildCacheIdMap(model);
  const workbookRels = model.workbookRels ?? [];

  for (const rel of workbookRels) {
    if (rel.Type === RelType.PivotCacheDefinition && rel.Target) {
      const match = rel.Target.match(/pivotCacheDefinition(\d+)\.xml/);
      if (match) {
        const defName = `pivotCacheDefinition${match[1]}`;
        const cacheId = rIdToCacheId.get(rel.Id);
        if (cacheId !== undefined) {
          definitionToCacheId.set(defName, cacheId);
        }
      }
    }
  }

  return definitionToCacheId;
}

function determineMetric(dataFields: Array<{ subtotal?: string }>): PivotTableSubtotal {
  if (dataFields.length > 0 && dataFields[0].subtotal) {
    return dataFields[0].subtotal as PivotTableSubtotal;
  }
  return "sum";
}

function determineValueMetrics(
  dataFields: Array<{ subtotal?: string }>,
  defaultMetric: PivotTableSubtotal
): PivotTableSubtotal[] {
  return dataFields.map(df => (df.subtotal as PivotTableSubtotal) || defaultMetric);
}
