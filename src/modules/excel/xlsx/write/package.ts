/**
 * The XLSX writer: a workbook in, a package of parts out through a ZIP adapter.
 *
 * Plain functions over the workbook model, like the XLSB writer beside it. Charts are written by
 * `write/charts`, which is loaded only for a package that has one, and nothing here imports the reader.
 */
import { missingRuleFormula } from "@excel/core/conditional-formula";
import {
  imageContentTypeFor,
  isForeignContainerPart,
  isRootOrWorkbookSourced
} from "@excel/core/opaque-part";
import type { PivotTable } from "@excel/core/pivot-table";
import { refuseUnsupported } from "@excel/core/unsupported";
import { namedStyleMap } from "@excel/core/workbook-core";
import type { Workbook, ExternalLinkModel, NamedStyleEntry } from "@excel/core/workbook.browser";
import {
  _collectExternalLinksForWrite,
  _recordAutoExternalLink,
  getWorkbookModel
} from "@excel/core/workbook.browser";
import type { XlsxWritable } from "@excel/core/xlsx-io-types";
import { ImageError, TableError } from "@excel/errors";
import type { ThreadedCommentPerson } from "@excel/types";
import {
  claimArrayRange,
  overlappingArrayReason,
  type ArrayRange
} from "@excel/utils/array-ranges";
import { filterDrawingAnchors, isExternalImage } from "@excel/utils/drawing-utils";
import { rewriteExternalRefs } from "@excel/utils/external-link-formula";
import { formulaNameSet, hasUnnamedThisRow } from "@excel/utils/future-functions";
import { MediaReferenceTracker } from "@excel/utils/media-references";
import {
  commentsPath,
  chartsheetPath,
  chartsheetRelsPath,
  ctrlPropPath,
  drawingPath,
  drawingRelsPath,
  externalLinkPath,
  externalLinkRelsPath,
  externalLinkRelTargetFromWorkbook,
  OOXML_REL_TARGETS,
  pivotCacheDefinitionRelTargetFromWorkbook,
  chartRelTargetFromDrawing,
  chartExRelTargetFromDrawing,
  mediaPath,
  pivotCacheDefinitionPath,
  pivotCacheDefinitionRelsPath,
  pivotCacheDefinitionRelTargetFromPivotTable,
  pivotCacheRecordsPath,
  pivotCacheRecordsRelTarget,
  pivotTablePath,
  pivotTableRelsPath,
  tablePath,
  themePath,
  OOXML_PATHS,
  resolveRelTarget,
  vmlDrawingPath,
  vmlDrawingHFPath,
  vmlDrawingHFRelsPath,
  worksheetPath,
  worksheetRelsPath,
  worksheetRelTarget
} from "@excel/utils/ooxml-paths";
import { isStoredOrdinaryFormula } from "@excel/utils/stored-formulas";
import { StreamBuf } from "@excel/utils/stream-buf";
import { tableNameAt } from "@excel/utils/table-at";
import {
  writableConditionalFormattings,
  writableDataValidations,
  writableDefinedNames,
  type FormulaCheck
} from "@excel/utils/unwritable-formulas";
import { createVmlShapeIdSpace, reserveVmlShapeIds } from "@excel/utils/vml-shape-ids";
import type {
  OpaquePart,
  OpaqueRelationship,
  OpaqueSourceRelationship
} from "@excel/xlsx/opaque-parts";
import {
  appendOpaqueSourceRelationships,
  opaqueContentTypeDeclarations,
  relationshipsPathFor,
  relationshipStillResolves,
  resolveReachableOpaqueParts
} from "@excel/xlsx/opaque-parts";
import { RelType } from "@excel/xlsx/rel-type";
import type {
  IZipWriter,
  MediaModel,
  WorkbookMediaLike,
  XlsxWriteOptions
} from "@excel/xlsx/types";
import {
  CHARTSHEET_DRAWING_EMU,
  renderChartsheetDrawingXml
} from "@excel/xlsx/write/chartsheet-drawing";
import { readMediaFile } from "@excel/xlsx/write/media-file";
import { ExternalLinkXform } from "@excel/xlsx/xform/book/external-link-xform";
import { WorkbookXform } from "@excel/xlsx/xform/book/workbook-xform";
import {
  renderPersonList,
  renderThreadedComments
} from "@excel/xlsx/xform/comment/threaded-comments-render";
import { AppXform } from "@excel/xlsx/xform/core/app-xform";
import { ContentTypesXform } from "@excel/xlsx/xform/core/content-types-xform";
import { CoreXform } from "@excel/xlsx/xform/core/core-xform";
import { FeaturePropertyBagXform } from "@excel/xlsx/xform/core/feature-property-bag-xform";
import { MetadataXform } from "@excel/xlsx/xform/core/metadata-xform";
import { RelationshipsXform } from "@excel/xlsx/xform/core/relationships-xform";
import { writesDynamicArray } from "@excel/xlsx/xform/sheet/cell-xform";
import { WorkSheetXform } from "@excel/xlsx/xform/sheet/worksheet-xform";
import { SharedStringsXform } from "@excel/xlsx/xform/strings/shared-strings-xform";
import { StylesXform } from "@excel/xlsx/xform/style/styles-xform";
import { theme1Xml } from "@excel/xlsx/xml/theme1";
import { appendToZip, createZipWriterAdapter, renderToZip } from "@excel/xlsx/zip-writer";
import { excelNameKey } from "@utils/name-key";

/**
 * Add `sheetName` to an ExternalLinkModel's `sheetNames` list if it isn't
 * already present. Ordering is preserved — the first-seen sheet wins
 * position 0, which matches what Excel does when writing externalLinks
 * itself.
 */
function upsertSheet(link: { sheetNames: string[] }, sheetName: string): void {
  if (!sheetName) {
    return;
  }
  if (!link.sheetNames.includes(sheetName)) {
    link.sheetNames.push(sheetName);
  }
}

/**
 * Scratch state used during `normaliseExternalLinks`. `links` is the
 * write-scoped ExternalLinkModel array (user-declared + auto-discovered)
 * the writer will consume; `byTarget` indexes it by lower-cased target
 * for O(1) lookup during formula rewriting; `workbook` is used to
 * persist auto-discoveries to the Workbook's private cache so subsequent
 * writes stay consistent without mutating `wb.externalLinks`.
 */
interface NormaliseScratch {
  links: ExternalLinkModel[];
  byTarget: Map<string, ExternalLinkModel>;
  workbook: Workbook;
}

function isStrictTemplateMode(options?: XlsxWriteOptions): boolean {
  return options?.templateMode === "strict" || options?.strictTemplateMode === true;
}

/**
 * Decide whether `writeXlsxBytes` should run the OOXML self-check after
 * producing bytes. Resolves the `validate` option against the current
 * environment:
 *
 *   - Explicit `true` / `false` → honoured as-is.
 *   - `undefined` (default)     → `true` in non-production Node.js
 *                                 when NOT running under vitest. We
 *                                 detect vitest via `process.env.VITEST`
 *                                 to avoid adding multi-second
 *                                 validation overhead to fixture
 *                                 `beforeAll` hooks that produce
 *                                 hundreds of workbooks (the chartEx
 *                                 preset corpus alone builds ~100
 *                                 fixtures per run — at ~450 ms each
 *                                 that is a 45 s penalty on every full
 *                                 suite execution). Vitest tests that
 *                                 need validation call
 *                                 `expectValidXlsx()` explicitly.
 *                                 `false` in production and in the
 *                                 browser where `process` is absent.
 *
 * Kept as a helper so the resolution rule is testable in isolation; a
 * caller overrides it by passing an explicit `validate` flag.
 */
function shouldAutoValidate(explicit: boolean | undefined): boolean {
  if (explicit !== undefined) {
    return explicit;
  }
  // In the browser `process` is undefined; skip the overhead there.
  if (typeof process === "undefined" || !process.env) {
    return false;
  }
  if (process.env.NODE_ENV === "production") {
    return false;
  }
  // Vitest sets VITEST=true automatically in its worker processes.
  // Skip the auto-check there; tests opt-in via `expectValidXlsx`.
  if (process.env.VITEST === "true") {
    return false;
  }
  return true;
}

/**
 * Run `validateXlsxBuffer` on writer output and emit a consolidated
 * `console.warn` for every detected problem. Never throws: a validator
 * exception is degraded to a warning so writers that intentionally
 * produce non-conformant xlsx (e.g. for negative-path tests) keep
 * working. The message includes the actionable opt-out so downstream
 * consumers know how to silence it without grepping docs.
 */
async function runWriteBufferSelfCheck(bytes: Uint8Array): Promise<void> {
  try {
    // Dynamic import: the OOXML validator (~66 KB) is a development-only
    // self-check that never runs in production (see `shouldAutoValidate`).
    // Loading it lazily keeps it out of consumer bundles entirely.
    const { validateXlsxBuffer } = await import("@excel/utils/ooxml-validator");
    const report = await validateXlsxBuffer(bytes, { maxProblems: 20 });
    if (report.ok) {
      return;
    }
    const summary = report.problems
      .map((p, i) => `  ${i + 1}. [${p.kind}] ${p.file ?? "<package>"}: ${p.message}`)
      .join("\n");
    // eslint-disable-next-line no-console
    console.warn(
      `[documonster] toBuffer() produced xlsx with ${report.problems.length} OOXML issue(s):\n` +
        `${summary}\n` +
        `Pass \`{ validate: false }\` to silence this self-check, or set NODE_ENV=production.`
    );
  } catch (err) {
    // eslint-disable-next-line no-console
    console.warn(
      `[documonster] toBuffer() self-check threw unexpectedly and was skipped: ${String(err)}`
    );
  }
}

/** Whether any sheet has an ordinary formula that was not read from a file — see `stored-formulas.ts`. */
function hasAuthoredFormula(model: any): boolean {
  return (model.worksheets ?? []).some((worksheet: any) =>
    (worksheet.rows ?? []).some((row: any) =>
      (row.cells ?? []).some(
        (cell: any) =>
          typeof cell?.formula === "string" &&
          cell.shareType === undefined &&
          !isStoredOrdinaryFormula(cell)
      )
    )
  );
}

/** Whether any part of the workbook may hold a formula — a cell, a conditional format, a validation, a name. */
function mayHoldFormula(model: any): boolean {
  return (
    (model.definedNames ?? []).length > 0 ||
    (model.worksheets ?? []).some(
      (worksheet: any) =>
        (worksheet.conditionalFormattings ?? []).length > 0 ||
        Object.keys(worksheet.dataValidations ?? {}).length > 0 ||
        (worksheet.rows ?? []).some((row: any) =>
          (row.cells ?? []).some((cell: any) => typeof cell?.formula === "string")
        )
    )
  );
}

/**
 * Leave out what holds a formula Excel will not open, and say what it was, as `Sheet!A1: reason`.
 *
 * Conditional formats, validations and defined names are replaced on the write model by copies without them. A cell is
 * written with its cached value by `CellXform.prepare`, which asks the same `check`; a shared formula's dependents go
 * with their master, as they do for an unnamed this-row reference.
 */
function dropUnwritableFormulas(model: any, check: FormulaCheck): string[] {
  const lost: string[] = [];
  for (const worksheet of model.worksheets ?? []) {
    const masters = new Set<string>();
    for (const row of worksheet.rows ?? []) {
      for (const cell of row.cells ?? []) {
        const reason = typeof cell?.formula === "string" ? check(cell.formula) : undefined;
        if (reason !== undefined) {
          masters.add(cell.address);
          lost.push(`${worksheet.name}!${cell.address}: ${reason}`);
        } else if (typeof cell?.sharedFormula === "string" && masters.has(cell.sharedFormula)) {
          lost.push(
            `${worksheet.name}!${cell.address}: formula (its shared master ${cell.sharedFormula} cannot be written)`
          );
        }
      }
    }
    if (worksheet.conditionalFormattings !== undefined) {
      worksheet.conditionalFormattings = writableConditionalFormattings(
        worksheet.conditionalFormattings,
        check,
        `${worksheet.name}!`,
        lost
      );
    }
    if (worksheet.dataValidations !== undefined) {
      worksheet.dataValidations = writableDataValidations(
        worksheet.dataValidations,
        check,
        `${worksheet.name}!`,
        lost
      );
    }
  }
  if (model.definedNames !== undefined) {
    model.definedNames = writableDefinedNames(model.definedNames, check, lost);
  }
  return lost;
}

function hasChartParts(model: any): boolean {
  const nonEmpty = (record: Record<string, unknown> | undefined) =>
    !!record && Object.keys(record).length > 0;
  return (
    nonEmpty(model.chartEntries) ||
    nonEmpty(model.chartExEntries) ||
    nonEmpty(model.chartExStructuredEntries)
  );
}

/** `test` asked with the workbook's registered functions, when it has any. */
function withRegisteredFunctions(
  test: (formula: string, functions?: Workbook["userFunctions"]) => boolean,
  functions: Workbook["userFunctions"]
): (formula: string) => boolean {
  return functions === undefined || functions.size === 0
    ? test
    : formula => test(formula, functions);
}

/**
 * Write every part of the package to a ZIP writer — the body shared by
 * {@link writeXlsxBytes} and {@link writeXlsxToStream}.
 */
async function writeXlsxPackage(
  workbook: Workbook,
  zip: IZipWriter,
  options?: XlsxWriteOptions
): Promise<void> {
  const model = getWorkbookModel(workbook);
  // The test for a formula that has to be written as a dynamic array parses formulas, so it is loaded only when a
  // sheet has a formula a caller set — see `CellXform.prepare`.
  // The workbook's registered functions declare whether they return arrays, which decides for a formula calling one.
  const needsArrayEvaluation = hasAuthoredFormula(model)
    ? withRegisteredFunctions(
        (await import("@excel/xlsx/write/array-formulas")).needsArrayEvaluation,
        workbook.userFunctions
      )
    : undefined;
  // A formula Excel refuses to open — a constant where only a reference is accepted, `RANK(2,{2,2,1})` — is written
  // nowhere: a cell keeps its cached value, and a conditional format, validation or defined name holding one is left
  // out. Checked for every formula, read or set: the check parses only those that call such a function.
  const unwritableFormula = mayHoldFormula(model)
    ? (await import("@excel/xlsx/write/array-formulas")).unwritableFormulaReason
    : undefined;
  const formulaLosses =
    unwritableFormula === undefined ? [] : dropUnwritableFormulas(model, unwritableFormula);
  prepareModel(workbook, model, { ...options, needsArrayEvaluation, unwritableFormula });
  // Media are written last, and only those some written part references — see `media-references.ts`.
  const mediaReferences = new MediaReferenceTracker();
  zip = observeParts(zip, mediaReferences);
  // **`unsupported` is honoured here, and used to be inert for this container.**
  //
  // `WorkbookWriteOptions.unsupported` is declared for both formats and was consulted only by the XLSB writer, so
  // `{ format: "xlsx", unsupported: "error" }` was a guarantee that could never fire — measured against the same
  // workbook, XLSB refused and XLSX wrote successfully. The comment justifying that said XLSX "expresses everything
  // this library models", which is true of everything except one thing: a preserved sheet part from the *other*
  // container. `resolveOpaqueReachability` drops those, because a BIFF12 record stream has no form in a
  // SpreadsheetML package, and the tab it described is gone.
  //
  // Applied here rather than at the five public entry points because this is the one place both platforms' writes
  // pass through with the options in hand. `prepareModel` has already run, so `model.opaqueForeignSheetParts` holds
  // what was dropped.
  // `options` is typed `XlsxWriteOptions` here while `unsupported` is declared on the format-aware
  // `WorkbookWriteOptions` that extends it — the object at runtime is the caller's, so the field is present when it
  // was passed. Read through the narrower shape rather than widening this function's signature, which every
  // caller shares.
  refuseUnsupported(
    [
      ...(model.opaqueForeignSheetParts ?? []),
      ...unwritableConditionalRules(model),
      ...unnamedThisRowCells(model),
      ...overlappingArrayCells(model),
      ...formulaLosses
    ],
    options as { readonly unsupported?: "error" | "ignore" } | undefined,
    "XLSX"
  );
  prepareChartsheets(model);
  // **Charts are loaded only when the package has one.** `write/charts` is the largest part of the writer —
  // the chart and chartEx renderers plus the raw-XML patcher that edits a loaded chart in place — and a
  // workbook of cells and images reaches none of it. Every part it writes is keyed by an entry in these three
  // maps, so an empty set of them is exactly "nothing to write", not an approximation of it.
  const charts = hasChartParts(model) ? await import("@excel/xlsx/write/charts") : undefined;
  charts?.prepareChartExSidecars(model);

  await addContentTypes(zip, model);
  await addOfficeRels(zip, model);
  await addWorkbookRels(zip, model);
  // Write workbook.xml before worksheets so that streaming readers can
  // resolve worksheet names/ids/state from workbook metadata before
  // processing worksheet entries.
  await addWorkbook(zip, model);
  await addWorksheets(zip, model);
  await addSharedStrings(zip, model);
  await addDrawings(zip, model);
  await addChartsheets(zip, model);
  const strictTemplateMode = isStrictTemplateMode(options);
  if (charts) {
    await charts.addCharts(zip, model, strictTemplateMode);
    await charts.addChartExEntries(zip, model, strictTemplateMode);
  }
  await addTables(zip, model);
  await addPivotTables(zip, model);
  await addExternalLinks(zip, model);
  await addThemes(zip, model);
  await addStyles(zip, model);
  await addFeaturePropertyBag(zip, model);
  await addMetadata(zip, model);
  await addApp(zip, model);
  await addCore(zip, model);
  await addPersons(zip, model);
  await addSlicerAndTimelineParts(zip, model);
  await addOpaqueParts(zip, model);
  await addMedia(zip, model, mediaReferences, workbook._releasedMedia);
}

/**
 * `zip`, reporting every part written through it to `tracker`.
 *
 * An explicit decorator over the two write paths `IZipWriter` has — `append` (`appendToZip`) and
 * `createEntry` (`renderToZip`) — rather than an interception of arbitrary members: every other
 * member is forwarded by name, so a new write path added to the interface is a compile error here
 * instead of a part the tracker silently never sees. Base64 parts are media bytes, never references.
 */
function observeParts(zip: IZipWriter, tracker: MediaReferenceTracker): IZipWriter {
  const observed: IZipWriter = {
    append(data, options) {
      if (!options.base64) {
        tracker.observe(options.name, data);
      }
      zip.append(data, options);
    },
    createEntry(name) {
      const entry = zip.createEntry(name);
      if (!tracker.wants(name)) {
        return entry;
      }
      const chunks: Array<string | Uint8Array> = [];
      return {
        write(chunk) {
          chunks.push(chunk);
          entry.write(chunk);
        },
        end() {
          tracker.observeChunks(name, chunks);
          entry.end();
        }
      };
    },
    pipe: stream => zip.pipe(stream),
    finalize: () => zip.finalize(),
    waitForDrain: () => zip.waitForDrain(),
    on(event, listener) {
      zip.on(event, listener);
      return observed;
    },
    once(event, listener) {
      zip.once(event, listener);
      return observed;
    },
    off(event, listener) {
      zip.off(event, listener);
      return observed;
    }
  };
  return observed;
}

/**
 * Emit the raw slicer/timeline parts captured on load. Pure
 * byte-copy — documonster does not modify these parts. The partner
 * Content-Types and rels are covered separately (content types in
 * `addContentTypes`, sheet/workbook rels by the corresponding
 * xforms consuming the existing `xl/_rels/*.rels` captured on
 * load).
 */
async function addSlicerAndTimelineParts(zip: IZipWriter, model: any): Promise<void> {
  for (const source of [
    model.slicerParts,
    model.slicerCacheParts,
    model.timelineParts,
    model.timelineCacheParts
  ] as Array<Record<string, Uint8Array> | undefined>) {
    if (!source) {
      continue;
    }
    for (const [path, bytes] of Object.entries(source)) {
      await appendToZip(zip, bytes, { name: path });
    }
  }
}

/**
 * Write the workbook-level `xl/persons/person.xml` part when the
 * model carries Office 365 threaded-comment authors. No-op when the
 * persons list is empty so legacy files without threaded comments
 * stay byte-identical.
 */
async function addPersons(zip: IZipWriter, model: any): Promise<void> {
  const persons = model.persons as ThreadedCommentPerson[] | undefined;
  if (!persons || persons.length === 0) {
    return;
  }
  await appendToZip(zip, renderPersonList(persons), { name: "xl/persons/person.xml" });
}

/**
 * Write workbook to a stream.
 *
 * The returned promise resolves only after the sink has accepted the whole
 * package. This writer **respects downstream backpressure**: when
 * `stream.write()` returns `false` it waits for the sink's `'drain'` event at
 * the next zip-entry boundary before producing more bytes. A sink that errors,
 * or that closes before the package is complete, rejects the promise instead
 * of parking on a `'drain'` that can never arrive.
 *
 * ⚠️ The sink must already be consumed (or be a real terminal sink such as
 * `fs.createWriteStream`, an HTTP response, or an upload body) *before* this
 * call — an unconsumed intermediate stream deadlocks once its buffers fill.
 * `Workbook.writeStream` documents that contract with worked examples, and
 * `Workbook.toStream` removes it altogether by inverting the flow.
 */
export async function writeXlsxToStream(
  workbook: Workbook,
  stream: XlsxWritable,
  options: XlsxWriteOptions = {}
): Promise<void> {
  const zip = createZipWriterAdapter(zipOptionsFor(workbook, options));
  zip.pipe(stream);
  await writeXlsxToZip(workbook, zip, options);
}

/**
 * Write the whole package into `zip` and finalize it — the step both entry points share, and the seam a
 * caller that needs its own adapter (an instrumented one, in the backpressure tests) writes through.
 */
export async function writeXlsxToZip(
  workbook: Workbook,
  zip: IZipWriter,
  options: XlsxWriteOptions = {}
): Promise<void> {
  await writeXlsxPackage(workbook, zip, options);
  await finalizeZip(zip);
}

/**
 * ZIP options for one write, with the entry timestamp defaulted from the workbook.
 *
 * A copy, not an assignment into `options.zip`: this used to write `modTime` into the caller's object, so
 * an options object reused for a second workbook stamped it with the first one's date.
 */
function zipOptionsFor(workbook: Workbook, options: XlsxWriteOptions): XlsxWriteOptions["zip"] {
  return {
    ...options.zip,
    modTime: options.zip?.modTime ?? workbook.modified ?? workbook.created
  };
}

/**
 * Write workbook to buffer
 */
export async function writeXlsxBytes(
  workbook: Workbook,
  options: XlsxWriteOptions = {}
): Promise<Uint8Array> {
  const stream = new StreamBuf();
  const zip = createZipWriterAdapter(zipOptionsFor(workbook, options));
  zip.pipe(stream);
  await writeXlsxToZip(workbook, zip, options);
  const bytes = stream.read() || new Uint8Array(0);

  // Optional OOXML self-check. Enabled by default in non-production
  // Node.js environments; disabled in the browser and in production.
  // See `XlsxWriteOptions.validate` for the resolution rules.
  if (shouldAutoValidate(options.validate)) {
    await runWriteBufferSelfCheck(bytes);
  }

  return bytes;
}

/**
 * Add media files to ZIP
 * Supports buffer, base64, and filename (Node only — see `write/media-file`)
 */
async function addMedia(
  zip: IZipWriter,
  model: MediaModel,
  references: MediaReferenceTracker,
  /** Media ids an API call stopped using — the only ones a write may leave out. */
  releasedMedia: ReadonlySet<number> | undefined
): Promise<void> {
  // Keep filename reads concurrent as before, but append the resolved data in
  // model order. ZIP output itself is strictly sequential, and only a
  // sequential append makes each completion promise mean that entry's bytes
  // have reached the sink before backpressure is inspected.
  const resolved = await Promise.all(
    (model.media as WorkbookMediaLike[]).map(async (medium, index) => {
      if (medium.type !== "image") {
        throw new ImageError("Unsupported media");
      }

      // External (linked) images carry only a `link` target — no bytes are
      // written into the package; the relationship (TargetMode="External")
      // references the image in place.
      if (isExternalImage(medium)) {
        return null;
      }

      // Preserve legacy behavior: `${undefined}` becomes "undefined" in template strings
      const mediaName = medium.name ?? "undefined";
      const fileName = `${mediaName}.${medium.extension}`;
      // Still registered, no longer used — an API call removed its last use. The model keeps it,
      // so its id stays valid; only this package goes without the bytes. Media that were never
      // released (unplaced, or unreferenced in the file that was read) are written as before.
      const name = mediaPath(fileName);
      if (releasedMedia?.has(index) && !references.isReferenced(name)) {
        return null;
      }

      if (medium.filename) {
        return { data: await readMediaFile(medium.filename), options: { name } };
      }

      if (medium.buffer) {
        return { data: medium.buffer, options: { name } };
      }

      if (medium.base64) {
        return {
          data: medium.base64.substring(medium.base64.indexOf(",") + 1),
          options: { name, base64: true }
        };
      }

      throw new ImageError("Unsupported media");
    })
  );

  for (const media of resolved) {
    if (media) {
      await appendToZip(zip, media.data, media.options);
    }
  }
}

/**
 * Write-time pass that brings the workbook model into a shape the writer
 * can serialise cleanly. Two concerns:
 *
 *   1. Build the final external-link list for this write, combining
 *      user-declared links (`wb.externalLinks`) with auto-discovered
 *      ones from previous writes (cached on the Workbook). The result
 *      is assigned to `model.externalLinks` and consumed by the writer;
 *      `wb.externalLinks` is **not** modified.
 *
 *   2. Scan every formula cell for `[Book]Sheet!` prefixes. Filename-form
 *      references that don't match an existing link trigger
 *      `_recordAutoExternalLink()` on the Workbook, which adds the target
 *      to the private writer cache (so subsequent writes are fixed-point
 *      stable) but leaves `wb.externalLinks` untouched.
 *
 *   3. Rewrite every external-ref formula so it uses the numeric `[N]`
 *      form, the canonical OOXML storage form. This mutation lands on
 *      the cell's model object — matching the library's existing
 *      write-time pattern for `ssId`, `styleId`, `si`, and `cm`.
 *      Subsequent writes see the `[N]` form directly and resolve it
 *      against `model.externalLinks`, giving idempotent output.
 */
function normaliseExternalLinks(workbook: Workbook, model: any): void {
  // Start from user-declared links, honouring their declaration order.
  const links = _collectExternalLinksForWrite(workbook);

  // Fast lookup: case-insensitive target → link object in `links`.
  const byTarget = new Map<string, ExternalLinkModel>();
  for (const link of links) {
    if (link.target) {
      byTarget.set(link.target.toLowerCase(), link);
    }
  }

  const scratch: NormaliseScratch = { links, byTarget, workbook: workbook };
  for (const worksheet of model.worksheets ?? []) {
    for (const row of worksheet.rows ?? []) {
      for (const cell of row.cells ?? []) {
        if (typeof cell?.formula === "string" && cell.formula.length > 0) {
          cell.formula = normaliseFormulaExternalRefs(cell.formula, scratch);
        }
        if (typeof cell?.sharedFormula === "string" && cell.sharedFormula.length > 0) {
          // Shared-formula clones typically carry the master's *address*
          // here, not a formula body — they won't match the ref regex.
          // Masters carry the formula on `.formula` (handled above).
          // We rewrite defensively in case a caller stored an actual
          // formula string here.
          cell.sharedFormula = normaliseFormulaExternalRefs(cell.sharedFormula, scratch);
        }
      }
    }
  }

  model.externalLinks = links;
}

/**
 * Rewrite a single formula so every external-ref prefix uses the numeric
 * `[N]` form. When an unknown filename-form reference is found we record
 * it on the workbook's private writer cache (so the next write can still
 * resolve it) and append a local link to `scratch.links` so subsequent
 * refs in the same formula see the freshly-assigned index.
 */
function normaliseFormulaExternalRefs(formula: string, scratch: NormaliseScratch): string {
  // rewriteExternalRefs internally calls findExternalRefs and returns
  // the original string unchanged when there are no matches — no need
  // for a separate guard scan here.
  return rewriteExternalRefs(formula, ref => {
    // Numeric ref: accept if it resolves, otherwise preserve verbatim so
    // Excel surfaces `#REF!` at load time — same as the old behaviour
    // for truly broken references.
    if (ref.numeric) {
      if (ref.index !== null && ref.index >= 1 && ref.index <= scratch.links.length) {
        upsertSheet(scratch.links[ref.index - 1], ref.sheet);
        return ref.index;
      }
      return null;
    }

    // Filename form — look up or auto-register.
    const key = ref.workbook.toLowerCase();
    let link = scratch.byTarget.get(key);
    if (!link) {
      const index = _recordAutoExternalLink(scratch.workbook, ref.workbook, ref.sheet);
      link = {
        index,
        target: ref.workbook,
        targetMode: "External",
        sheetNames: ref.sheet ? [ref.sheet] : [],
        cachedValues: {}
      };
      // Keep the local writer list dense: insert at its future position.
      // `_recordAutoExternalLink` guarantees `index` is user.length + cache.size
      // at the time of insertion, which equals `scratch.links.length + 1`
      // whenever we walk formulas sequentially.
      scratch.links.push(link);
      scratch.byTarget.set(key, link);
    } else {
      upsertSheet(link, ref.sheet);
      // Keep the workbook cache's sheetNames in sync so subsequent
      // writes see the accumulated set.
      if (ref.sheet) {
        _recordAutoExternalLink(scratch.workbook, ref.workbook, ref.sheet);
      }
    }
    return link.index;
  });
}

async function addContentTypes(zip: IZipWriter, model: any): Promise<void> {
  // Extensions this writer declares itself. A preserved part that relied on a
  // conflicting Default for one of these needs an Override instead, or the
  // writer's value silently reclassifies it.
  const reserved = new Map<string, string>([
    ["rels", "application/vnd.openxmlformats-package.relationships+xml"],
    ["xml", "application/xml"],
    // **`vml`, because `ContentTypesXform` declares it and this map decides who may.**
    //
    // A preserved VML whose source declared it through a `Default` would otherwise have that `Default` re-emitted
    // beside this writer's, and OPC allows one per extension. The XLSB writer reserves it for exactly this reason and
    // says so; the XLSX writer emits the same `Default` — for comments, form controls, header watermarks and
    // chartsheet VML — and did not reserve it. A defect fixed on one side of a pair is not fixed.
    ["vml", "application/vnd.openxmlformats-officedocument.vmlDrawing"]
  ]);
  for (const medium of model.media ?? []) {
    if (medium?.type === "image" && typeof medium.extension === "string") {
      // Skipped for a linked image, which adds no part to the package and so reserves nothing — the condition
      // `ContentTypesXform` already applies where it emits the `Default`. Reserving an extension for a part that is
      // never written turns a preserved part's own `Default` into an `Override` for no reason.
      if (isExternalImage(medium)) {
        continue;
      }
      const extension = medium.extension.toLowerCase();
      reserved.set(extension, imageContentTypeFor(extension));
    }
  }

  const declarations = opaqueContentTypeDeclarations(
    model.opaqueParts,
    model.opaqueContentTypeDefaults,
    reserved,
    // The two `Override`s `ContentTypesXform` writes unconditionally — see `reservedPaths`.
    new Set([OOXML_PATHS.xlTheme1.toLowerCase(), OOXML_PATHS.xlStyles.toLowerCase()])
  );
  await renderToZip(zip, OOXML_PATHS.contentTypes, new ContentTypesXform(), {
    ...model,
    opaqueContentTypes: declarations.overrides,
    opaqueContentTypeDefaults: declarations.defaults
  });
}

async function addApp(zip: IZipWriter, model: any): Promise<void> {
  await renderToZip(zip, OOXML_PATHS.docPropsApp, new AppXform(), model);
}

async function addCore(zip: IZipWriter, model: any): Promise<void> {
  await renderToZip(zip, OOXML_PATHS.docPropsCore, new CoreXform(), model);
}

async function addThemes(zip: IZipWriter, model: any): Promise<void> {
  // **The built-in fallback stands down for a preserved theme.**
  //
  // A theme read from XLSX is modelled, so `model.themes` is set and the fallback never applies. A theme read from
  // *XLSB* is a preserved part instead — this reader does not parse one — so `model.themes` is empty and the default
  // was written to `xl/theme/theme1.xml`, which `addOpaqueParts` then wrote again. One ZIP entry survived that, but
  // `[Content_Types].xml` declared the `Override` twice and the package's own validator refused it.
  //
  // The XLSB writer has the same guard for the same reason, in `themeParts`: a fallback exists to stop a package
  // having *no* theme, and a preserved one means it has one. Suppressing it on the *path* rather than on the
  // existence of any preserved theme is what keeps a preserved `theme2.xml` from silencing the default for `theme1`.
  //
  // **`theme1` is ensured, not merely defaulted when nothing is modelled.** The workbook relationship and the
  // content-type override always name `xl/theme/theme1.xml`, so a package without that part dangles both. An XLSX
  // read of a file that carries no theme models `themes` as `{}` — set, so a `model.themes || default` fallback
  // never fired and the package declared a part it did not contain.
  const preserved = new Set(
    ((model.opaqueParts ?? []) as readonly OpaquePart[]).map(part => part.path.toLowerCase())
  );
  const themes: Record<string, string> = { ...model.themes };
  if (themes.theme1 === undefined && !preserved.has(themePath("theme1").toLowerCase())) {
    themes.theme1 = theme1Xml;
  }
  for (const name of Object.keys(themes)) {
    await appendToZip(zip, themes[name], { name: themePath(name) });
  }
}

async function addOfficeRels(zip: IZipWriter, model: any): Promise<void> {
  const relationships: any[] = [
    { Id: "rId1", Type: RelType.OfficeDocument, Target: OOXML_PATHS.xlWorkbook },
    { Id: "rId2", Type: RelType.CoreProperties, Target: OOXML_PATHS.docPropsCore },
    { Id: "rId3", Type: RelType.ExtenderProperties, Target: OOXML_PATHS.docPropsApp }
  ];
  appendOpaqueSourceRelationships(relationships, model.opaqueParts, "");
  await renderToZip(zip, OOXML_PATHS.rootRels, new RelationshipsXform(), relationships);
}

/**
 * The conditional-formatting rules this write leaves out, as `Sheet!range: reason`.
 *
 * A comparison or an expression rule without its formula is skipped by the worksheet writer (see
 * `missingRuleFormula`) — there is no `<cfRule>` that says "compare against nothing" — and this is how the skip
 * reaches the caller: through `unsupported`, like every other thing a writer cannot carry.
 */
function unwritableConditionalRules(model: any): string[] {
  const lost: string[] = [];
  for (const worksheet of model.worksheets ?? []) {
    for (const block of worksheet.conditionalFormattings ?? []) {
      for (const rule of block.rules ?? []) {
        const reason = missingRuleFormula(rule);
        if (reason !== undefined) {
          lost.push(`${worksheet.name}!${block.ref}: conditional formatting ${reason}`);
        }
      }
    }
  }
  return lost;
}

/**
 * Cells whose formula holds an unnamed this-row reference (`[@x]`) and that sit in no table, as `Sheet!A1: reason`.
 *
 * A file stores a this-row reference with its table's name, and there is no table to name: Excel reads the unnamed
 * form as `#REF!`. Such a cell is written with its cached value and no formula (see `CellXform.prepare`), and this is
 * how that reaches the caller.
 */
function unnamedThisRowCells(model: any): string[] {
  const lost: string[] = [];
  for (const worksheet of model.worksheets ?? []) {
    const tableAt = tableNameAt(worksheet.tables);
    // A shared master lost this way takes its dependents with it — see `CellXform.prepare`.
    const masters = new Set<string>();
    for (const row of worksheet.rows ?? []) {
      for (const cell of row.cells ?? []) {
        if (
          typeof cell.formula === "string" &&
          hasUnnamedThisRow(cell.formula) &&
          tableAt(cell.address) === undefined
        ) {
          masters.add(cell.address);
          lost.push(
            `${worksheet.name}!${cell.address}: formula (an unnamed this-row reference outside any table)`
          );
        } else if (typeof cell.sharedFormula === "string" && masters.has(cell.sharedFormula)) {
          lost.push(
            `${worksheet.name}!${cell.address}: formula (its shared master ${cell.sharedFormula} has an unnamed this-row reference outside any table)`
          );
        }
      }
    }
  }
  return lost;
}

/**
 * Array formulas whose master sits inside an earlier array's range, as `Sheet!A1: reason`.
 *
 * Excel refuses a cell that is both a member of one array and the master of another. The cell is written with its
 * cached result (see `CellXform.prepare`, which claims the ranges in the same order through the same rule), and this is
 * how that reaches the caller.
 */
function overlappingArrayCells(model: any): string[] {
  const lost: string[] = [];
  for (const worksheet of model.worksheets ?? []) {
    const ranges: ArrayRange[] = [];
    for (const row of worksheet.rows ?? []) {
      for (const cell of row.cells ?? []) {
        if (
          cell?.shareType === "array" &&
          typeof cell.ref === "string" &&
          typeof cell.formula === "string"
        ) {
          const outer = claimArrayRange(ranges, cell.address, cell.ref);
          if (outer !== undefined) {
            lost.push(`${worksheet.name}!${cell.address}: ${overlappingArrayReason(outer)}`);
          }
        }
      }
    }
  }
  return lost;
}

/**
 * Drop preserved parts this write would leave unreachable, before anything is
 * emitted.
 *
 * Runs at the head of `prepareModel` because four writers consume the result —
 * the content types, the root rels, the workbook rels and the parts themselves —
 * and a part dropped after one of them had already declared it would leave the
 * package inconsistent in a different way than the one being fixed.
 */
function resolveOpaqueReachability(model: any): void {
  const all: OpaquePart[] = model.opaqueParts ?? [];
  // **A BIFF12 sheet part cannot sit in a SpreadsheetML package**, and recognising workbook-sourced edges across
  // containers is what made this reachable in the first place: reading `cal-any_sheets.xlsb` and writing XLSX put
  // `xl/chartsheets/sheet1.bin` — content type `application/vnd.ms-excel.chartsheet` — inside the output. Reported as
  // a drop with its own reason, because "could not cross the container" is a different fact from "nothing points at
  // it any more". See `isForeignContainerPart`.
  const foreign = all.filter(part => isForeignContainerPart(part, "xlsx"));
  const parts =
    foreign.length === 0 ? all : all.filter(part => !isForeignContainerPart(part, "xlsx"));
  if (foreign.length > 0) {
    model.opaqueParts = parts;
    // Recorded on the model so `writeXlsxPackage` can apply `unsupported` to it — the one loss this writer has, and the
    // reason the option is no longer inert for XLSX. Unlike `opaqueDrops`, this is read back in the same call.
    model.opaqueForeignSheetParts = foreign.map(
      part => `${part.path}: preserved part from the other container`
    );
  }
  if (parts.length === 0) {
    pruneSheetRelationships(model, parts);
    return;
  }

  // The edges this write will actually emit: the root and workbook rels are
  // rebuilt with the preserved entries appended, and each sheet re-emits its
  // own. Sheet-sourced targets are resolved against a canonical worksheet path
  // because every sheet lives in `xl/worksheets/`, so a relative target
  // resolves identically whichever file index the sheet ends up with — which
  // is also why deleting a sheet is what removes its edges, not renumbering.
  const emitted: OpaqueSourceRelationship[] = [];
  for (const part of parts) {
    for (const inbound of part.sourceRelationships ?? []) {
      // Both containers' spellings of "the workbook" — see `isRootOrWorkbookSourced`. Comparing against this
      // writer's own path dropped every workbook-sourced part that arrived from an XLSB read.
      if (isRootOrWorkbookSourced(inbound)) {
        emitted.push(inbound);
      }
    }
  }
  const canonicalSheet = worksheetPath(1);
  for (const worksheet of model.worksheets ?? []) {
    for (const relationship of worksheet.opaqueRels ?? []) {
      emitted.push({ ...relationship, source: canonicalSheet });
    }
  }

  const resolved = resolveReachableOpaqueParts(parts, emitted);
  if (resolved.drops.length === 0) {
    if (foreign.length > 0) {
      pruneSheetRelationships(model, parts);
    }
    return;
  }

  model.opaqueParts = resolved.parts;
  // **Deliberately not recorded on the model, because nothing would read it.** `model` here is the object
  // `getWorkbookModel` returned for this write, so assigning `opaqueDrops` on it reaches no caller —
  // `WorkbookModel.opaqueDrops` is a *read-time* report, populated by the reader from the parts it declined to keep,
  // and a write never touches the workbook's copy. The assignment that used to sit here looked like a report and was
  // dead from the moment it was written; a `"foreign-sheet-part"` entry added beside it was dead for the same reason.
  //
  // The XLSB writer surfaces both decisions through its `unsupported` list, which is where everything that writer
  // cannot carry goes. There is no equivalent here: an XLSX write has no loss channel at all — the same gap
  // `WorkbookReadOptions.blankCells` already documents for the collapsed-blank case. Stating that is better than an
  // assignment implying otherwise.

  pruneSheetRelationships(model, resolved.parts);
}

/**
 * A sheet must not keep pointing at a part that is no longer written — whether it became unreachable or could not
 * cross the container (a foreign part's sheet edge would otherwise dangle).
 */
function pruneSheetRelationships(model: any, writtenParts: readonly OpaquePart[]): void {
  const canonicalSheet = worksheetPath(1);
  const kept = new Set(writtenParts.map(part => part.path.toLowerCase()));
  for (const worksheet of model.worksheets ?? []) {
    if (!worksheet.opaqueRels) {
      continue;
    }
    // Through the shared predicate — see `relationshipStillResolves`. The XLSB writer had the inverse of this test.
    const surviving = worksheet.opaqueRels.filter((relationship: OpaqueRelationship) =>
      relationshipStillResolves(canonicalSheet, relationship, kept)
    );
    worksheet.opaqueRels = surviving.length > 0 ? surviving : undefined;
  }
}

/**
 * Write the preserved parts and each one's own `.rels`.
 *
 * The relationships that point *at* these parts are not written here — they
 * belong to `[Content_Types].xml` and to the `.rels` of the modelled parts
 * that declared them, which are regenerated elsewhere and pick the preserved
 * entries up through `appendOpaqueSourceRelationships`.
 */
async function addOpaqueParts(zip: IZipWriter, model: any): Promise<void> {
  const parts: OpaquePart[] = model.opaqueParts ?? [];
  for (const part of parts) {
    await appendToZip(zip, part.data, { name: part.path });
    if (part.relationships && part.relationships.length > 0) {
      await renderToZip(
        zip,
        relationshipsPathFor(part.path),
        new RelationshipsXform(),
        part.relationships.map(rel => ({
          Id: rel.id,
          Type: rel.type,
          Target: rel.target,
          ...(rel.targetMode ? { TargetMode: rel.targetMode } : {})
        }))
      );
    }
  }
}

async function addWorkbookRels(zip: IZipWriter, model: any): Promise<void> {
  let count = 1;
  const relationships: any[] = [
    { Id: `rId${count++}`, Type: RelType.Styles, Target: OOXML_REL_TARGETS.workbookStyles },
    { Id: `rId${count++}`, Type: RelType.Theme, Target: OOXML_REL_TARGETS.workbookTheme1 }
  ];
  if (model.sharedStrings.count) {
    relationships.push({
      Id: `rId${count++}`,
      Type: RelType.SharedStrings,
      Target: OOXML_REL_TARGETS.workbookSharedStrings
    });
  }

  // Add FeaturePropertyBag relationship if checkboxes are used
  if (model.hasCheckboxes) {
    relationships.push({
      Id: `rId${count++}`,
      Type: RelType.FeaturePropertyBag,
      Target: OOXML_REL_TARGETS.workbookFeaturePropertyBag
    });
  }
  // Add metadata relationship for dynamic array formulas
  if (model.hasDynamicArrayFormulas) {
    relationships.push({
      Id: `rId${count++}`,
      Type: RelType.SheetMetadata,
      Target: OOXML_REL_TARGETS.workbookMetadata
    });
  }
  // Office 365 threaded comments need a workbook-level person
  // directory. The rel Target is `persons/person.xml` (relative to
  // the xl/ workbook home, matching how Excel writes it).
  if (model.hasPersons) {
    relationships.push({
      Id: `rId${count++}`,
      Type: RelType.Person,
      Target: "persons/person.xml"
    });
  }
  // R9-B6: Deduplicate pivot cache relationships by cacheId. When multiple pivot
  // tables share the same cache, only one workbook relationship should be created.
  // Also assigns rId to each pivot table (R9-B7: typed on PivotTable interface).
  const seenCacheIds = new Map<string, string>(); // cacheId → rId
  (model.pivotTables ?? []).forEach((pivotTable: PivotTable) => {
    const existing = seenCacheIds.get(pivotTable.cacheId);
    if (existing) {
      // Shared cache: reuse the rId from the first pivot table with this cacheId
      pivotTable.rId = existing;
    } else {
      pivotTable.rId = `rId${count++}`;
      seenCacheIds.set(pivotTable.cacheId, pivotTable.rId);
      relationships.push({
        Id: pivotTable.rId,
        Type: RelType.PivotCacheDefinition,
        Target: pivotCacheDefinitionRelTargetFromWorkbook(pivotTable.tableNumber)
      });
    }
  });
  model.worksheets.forEach((worksheet: any) => {
    worksheet.rId = `rId${count++}`;
    // fileIndex is assigned once in prepareModel() — use it directly
    relationships.push({
      Id: worksheet.rId,
      Type: RelType.Worksheet,
      Target: worksheetRelTarget(worksheet.fileIndex)
    });
  });

  // Add chartsheet relationships
  (model.chartsheets || []).forEach((cs: any) => {
    cs.rId = `rId${count++}`;
    relationships.push({
      Id: cs.rId,
      Type: RelType.Chartsheet,
      Target: `chartsheets/sheet${cs.sheetNo}.xml`
    });
  });

  // External workbook link rels are written AFTER worksheets on purpose:
  // Excel tolerates either order, but stable ordering (worksheets then
  // externalLinks) keeps the emitted workbook.xml.rels diff-friendly for
  // round-trip tests. Each external link becomes a regular Relationship
  // entry targeting `externalLinks/externalLinkN.xml` inside `xl/`; the
  // actual external file path is pointed at by the nested
  // externalLinkN.xml.rels part written by addExternalLinks().
  //
  // The list items here are the deep-copies produced by
  // `normaliseExternalLinks` — assigning `link.rId` is safe and does not
  // leak into the user's Workbook.externalLinks.
  const externalLinks = (model.externalLinks ?? []) as ExternalLinkModel[];
  for (const link of externalLinks) {
    link.rId = `rId${count++}`;
    relationships.push({
      Id: link.rId,
      Type: RelType.ExternalLink,
      Target: externalLinkRelTargetFromWorkbook(link.index)
    });
  }

  appendOpaqueSourceRelationships(relationships, model.opaqueParts, OOXML_PATHS.xlWorkbook);

  const xform = new RelationshipsXform();
  await renderToZip(zip, OOXML_PATHS.xlWorkbookRels, xform, relationships);
}

async function addFeaturePropertyBag(zip: IZipWriter, model: any): Promise<void> {
  if (!model.hasCheckboxes) {
    return;
  }
  await renderToZip(zip, OOXML_PATHS.xlFeaturePropertyBag, new FeaturePropertyBagXform(), {});
}

async function addMetadata(zip: IZipWriter, model: any): Promise<void> {
  if (!model.hasDynamicArrayFormulas) {
    return;
  }
  await renderToZip(zip, OOXML_PATHS.xlMetadata, new MetadataXform(), {
    dynamicArrayCount: model.dynamicArrayCount
  });
}

async function addSharedStrings(zip: IZipWriter, model: any): Promise<void> {
  if (model.sharedStrings && model.sharedStrings.count) {
    await renderToZip(
      zip,
      OOXML_PATHS.xlSharedStrings,
      model.sharedStrings,
      model.sharedStrings.model
    );
  }
}

async function addStyles(zip: IZipWriter, model: any): Promise<void> {
  if (model.styles) {
    await renderToZip(zip, OOXML_PATHS.xlStyles, model.styles, model.styles.model);
  }
}

/**
 * The workbook's defined names, for the parts that hold formulas: a bare function name passed as an argument is
 * written `_xleta.SUM` unless the workbook defines that name — see `toFileFormula`. Read at render time, after the
 * chart writer has added its hidden `_xlchart.` names.
 */
function formulaNames(model: any): ReadonlySet<string> {
  return formulaNameSet(model.definedNames ?? []);
}

async function addWorkbook(zip: IZipWriter, model: any): Promise<void> {
  await renderToZip(zip, OOXML_PATHS.xlWorkbook, new WorkbookXform(), model, formulaNames(model));
}

async function addWorksheets(zip: IZipWriter, model: any): Promise<void> {
  const worksheetXform = new WorkSheetXform();
  const relationshipsXform = new RelationshipsXform();

  // Lazily load the optional comment / VML / form-control xforms only when
  // some worksheet actually needs them, so comment/control-free workbooks
  // never pull these (~12 KB + VML) into the bundle.
  const needsComments = model.worksheets.some((ws: any) => ws.comments.length > 0);
  const needsVml = model.worksheets.some(
    (ws: any) =>
      ws.comments.length > 0 ||
      (ws.formControls && ws.formControls.length > 0) ||
      ws.legacyControls?.length > 0 ||
      ws.headerImage
  );
  const needsCtrlProp = model.worksheets.some(
    (ws: any) => ws.formControls && ws.formControls.length > 0
  );

  const commentsXform = needsComments
    ? new (await import("@excel/xlsx/xform/comment/comments-xform")).CommentsXform()
    : null;
  const vmlDrawingXform = needsVml
    ? new (await import("@excel/xlsx/xform/drawing/vml-drawing-xform")).VmlDrawingXform()
    : null;
  const ctrlPropXform = needsCtrlProp
    ? new (await import("@excel/xlsx/xform/drawing/ctrl-prop-xform")).CtrlPropXform()
    : null;

  for (const worksheet of model.worksheets) {
    const { fileIndex } = worksheet;

    // Worksheet XML: stream directly to the zip entry (avoids holding the
    // entire XML in memory) and use the shared completion/backpressure path.
    await renderToZip(
      zip,
      worksheetPath(fileIndex),
      worksheetXform,
      worksheet,
      formulaNames(model)
    );

    if (worksheet.rels && worksheet.rels.length) {
      await renderToZip(zip, worksheetRelsPath(fileIndex), relationshipsXform, worksheet.rels);
    }

    // Generate comments XML (separate from VML)
    if (worksheet.comments.length > 0) {
      await renderToZip(zip, commentsPath(fileIndex), commentsXform!, worksheet);
    }

    // Office 365 threaded comments sit in their own part tree
    // alongside classic VML comments. Written straight from the
    // structured model without going through an xform instance —
    // the payload is small and the shape maps 1:1 onto the output.
    if (worksheet.threadedComments && worksheet.threadedComments.length > 0) {
      const xml = renderThreadedComments(worksheet.threadedComments);
      await appendToZip(zip, xml, {
        name: `xl/threadedComments/threadedComment${fileIndex}.xml`
      });
    }

    // Generate unified VML drawing (contains both notes and form controls)
    const hasComments = worksheet.comments.length > 0;
    const hasFormControls = worksheet.formControls && worksheet.formControls.length > 0;
    const legacyControls = worksheet.legacyControls ?? [];

    if (hasComments || hasFormControls || legacyControls.length > 0) {
      await renderToZip(zip, vmlDrawingPath(fileIndex), vmlDrawingXform!, {
        comments: hasComments ? worksheet.comments : [],
        formControls: hasFormControls ? worksheet.formControls : [],
        legacyControls,
        shapeIds: worksheet.vmlShapeIds
      });
    }

    // Generate VML drawing for header/footer images (watermark in header mode)
    if (worksheet.headerImages?.length || worksheet.headerImage) {
      const headerImages = worksheet.headerImages ?? [worksheet.headerImage];

      // Write the VML file for the header image
      await renderToZip(zip, vmlDrawingHFPath(fileIndex), vmlDrawingXform!, {
        comments: [],
        formControls: [],
        headerImages: headerImages.map(hdrImage => ({
          imageRelId: hdrImage.imageRelId,
          width: hdrImage.headerWidth,
          height: hdrImage.headerHeight,
          position: hdrImage.position,
          shape: hdrImage.vmlShape
        })),
        shapeIds: worksheet.headerImageShapeIds
      });

      // Write the VML rels file referencing the image
      await renderToZip(
        zip,
        vmlDrawingHFRelsPath(fileIndex),
        relationshipsXform,
        headerImages.map(hdrImage => {
          const bookImage = hdrImage.bookImage;
          const imageFileName =
            bookImage.name &&
            bookImage.extension &&
            bookImage.name.endsWith(`.${bookImage.extension}`)
              ? bookImage.name
              : `${bookImage.name}.${bookImage.extension}`;
          return {
            Id: hdrImage.imageRelId,
            Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            Target: `../media/${imageFileName}`
          };
        })
      );
    }

    // Generate ctrlProp files for form controls
    if (hasFormControls) {
      for (const control of worksheet.formControls) {
        await renderToZip(zip, ctrlPropPath(control.ctrlPropId), ctrlPropXform!, control);
      }
    }
    // A preserved control's part, as it was read, under the number this write gave it.
    for (const control of legacyControls) {
      if (control.ctrlPropId !== undefined && control.ctrlPropXml !== undefined) {
        await appendToZip(zip, control.ctrlPropXml, { name: ctrlPropPath(control.ctrlPropId) });
      }
    }
  }
}

async function addChartsheets(zip: IZipWriter, model: any): Promise<void> {
  if (!model.chartsheets || model.chartsheets.length === 0) {
    return;
  }
  const { ChartsheetXform } = await import("@excel/xlsx/xform/sheet/chartsheet-xform");
  const { VmlDrawingXform } = await import("@excel/xlsx/xform/drawing/vml-drawing-xform");
  const chartsheetXform = new ChartsheetXform();
  const relsXform = new RelationshipsXform();
  const vmlDrawingXform = new VmlDrawingXform();
  // Track VML drawing zip paths we re-emit for chartsheets so we
  // don't accidentally write the same VML part twice when a single
  // VML file is referenced by multiple chartsheets. Writing a ZIP
  // entry twice produces a package with duplicate central-directory
  // entries — most consumers tolerate it (reading the last), but
  // validators flag it and `unzip -l` shows the duplication.
  const emittedVmlPaths = new Set<string>();

  for (const cs of model.chartsheets || []) {
    await renderToZip(zip, chartsheetPath(cs.sheetNo), chartsheetXform, cs);

    // Chartsheet rels. A chartsheet may carry rels beyond the
    // drawing reference — `legacyDrawing`, `legacyDrawingHF`,
    // `drawingHF`, `picture`, etc. — and those rels are referenced
    // by `r:id` attributes inside the raw-captured `rawChildren`
    // blocks. If we only emit the drawing rel (the previous
    // implementation), every other r:id goes dangling at save,
    // corrupting the package.
    //
    // Strategy:
    //   1. Start with the preserved `relationships` list from load
    //      (missing for newly-created chartsheets).
    //   2. Overlay / insert the current drawing rel — the drawing
    //      target may have been rewritten (e.g. chartsheet renamed
    //      or its drawing renumbered) so we replace any prior
    //      entry with the same Id.
    const baseRels: any[] = Array.isArray(cs.relationships) ? [...cs.relationships] : [];
    if (cs.drawing) {
      const drawingRel = {
        Id: cs.drawing.rId,
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
        Target: `../drawings/${cs.drawingName}.xml`
      };
      const existingIdx = baseRels.findIndex((r: any) => r?.Id === cs.drawing.rId);
      if (existingIdx >= 0) {
        baseRels[existingIdx] = drawingRel;
      } else {
        baseRels.push(drawingRel);
      }
    }
    if (baseRels.length > 0) {
      await renderToZip(zip, chartsheetRelsPath(cs.sheetNo), relsXform, baseRels);
    }

    // Re-emit any VML drawing parts this chartsheet's rels reference.
    // The worksheet loop only emits VML for worksheets that own
    // comments / form controls / header images; a chartsheet that
    // carries its own `<legacyDrawing r:id="…"/>` would preserve its
    // rel target on write but leave the VML body missing from the
    // package — a dangling relationship. Walk the chartsheet's rels,
    // resolve each VML target against the chartsheet path, and emit
    // the parsed body captured at load time.
    if (model.vmlDrawings || model.vmlDrawingHF || cs.headerImages?.length) {
      const baseDir = `xl/chartsheets/`;
      for (const rel of baseRels) {
        if (rel?.Type !== RelType.VmlDrawing || !rel.Target) {
          continue;
        }
        const vmlPath = resolveRelTarget(baseDir, rel.Target);
        if (emittedVmlPaths.has(vmlPath)) {
          continue;
        }
        // The header/footer VML: the one `legacyDrawingHF` names, whatever it is called
        // (Excel calls it `vmlDrawingN.vml`, like a comment's).
        const isHeaderVml =
          (cs.legacyDrawingHF?.rId !== undefined && rel.Id === cs.legacyDrawingHF.rId) ||
          /\/vmlDrawingHF\d+[.]vml$/.test(vmlPath);
        const vmlName = isHeaderVml
          ? /\/(vmlDrawing(?:HF)?\d+)[.]vml$/.exec(vmlPath)?.[1]
          : undefined;
        const validHeaderImages = vmlName
          ? (cs.headerImages ?? []).flatMap((image: any) => {
              const medium = model.media?.[image.imageId];
              return medium ? [{ image, medium }] : [];
            })
          : [];
        const parsedHeaderImages = vmlName ? model.vmlDrawingHF?.[vmlName] : undefined;
        const headerImages =
          parsedHeaderImages ??
          validHeaderImages.map(({ image }: any, index: number) => ({
            imageRelId: `rId${index + 1}`,
            width: image.width,
            height: image.height,
            position: image.position,
            shape: image.shape
          }));
        const vmlModel =
          model.vmlDrawings?.[vmlPath] ?? (headerImages ? { headerImages } : undefined);
        if (!vmlModel) {
          continue;
        }
        emittedVmlPaths.add(vmlPath);
        await renderToZip(zip, vmlPath, vmlDrawingXform, vmlModel);
        const headerImageRels =
          (vmlName && model.vmlDrawingRels?.[vmlName]) ??
          validHeaderImages.map(({ medium }: any, index: number) => {
            const filename = medium.name?.endsWith(`.${medium.extension}`)
              ? medium.name
              : `${medium.name}.${medium.extension}`;
            return {
              Id: `rId${index + 1}`,
              Type: RelType.Image,
              Target: `../media/${filename}`
            };
          });
        if (vmlName && headerImageRels?.length) {
          const relsPath = vmlPath.replace(/\/([^/]+)[.]vml$/, "/_rels/$1.vml.rels");
          await renderToZip(zip, relsPath, relsXform, headerImageRels);
        }
        // `prepareChartsheets` already flipped `model.hasChartsheetVml`
        // before content-types was written, so no further signalling
        // is needed here.
      }
    }
  }
}

async function addDrawings(zip: IZipWriter, model: any): Promise<void> {
  // Skip entirely (and avoid loading DrawingXform ~34 KB) when no worksheet
  // has a drawing. Chartsheets emit their drawing XML verbatim (without
  // DrawingXform), so account for them separately.
  const hasWorksheetDrawing = model.worksheets.some((ws: any) => ws.drawing);
  const hasChartsheetDrawing = (model.chartsheets ?? []).some(
    (cs: any) => cs.drawingName && (cs.chartNumber || cs.chartExNumber)
  );
  if (!hasWorksheetDrawing && !hasChartsheetDrawing) {
    return;
  }
  const relsXform = new RelationshipsXform();

  if (hasWorksheetDrawing) {
    const { DrawingXform } = await import("@excel/xlsx/xform/drawing/drawing-xform");
    const drawingXform = new DrawingXform();

    for (const worksheet of model.worksheets) {
      const { drawing } = worksheet;
      if (drawing) {
        const filteredAnchors = filterDrawingAnchors(drawing.anchors ?? []);
        const drawingForWrite = drawing.anchors
          ? { ...drawing, anchors: filteredAnchors }
          : drawing;
        drawingXform.prepare(drawingForWrite);
        await renderToZip(zip, drawingPath(drawing.name), drawingXform, drawingForWrite);

        await renderToZip(zip, drawingRelsPath(drawing.name), relsXform, drawing.rels);
      }
    }
  }

  // Chartsheet drawings — each chartsheet references a drawing
  // containing a single chart that fills the entire sheet. Unlike
  // worksheet-embedded charts (where a `<xdr:twoCellAnchor>` with
  // `<xdr:from>/<xdr:to>` cell references pins the chart to a
  // rectangle of cells, whose dimensions Excel computes from the
  // sheet's column widths and row heights), a chartsheet has no
  // cell grid — its `sheetData` is empty. A cell-based anchor on
  // a chartsheet therefore resolves to a 0×0 rectangle and Excel
  // renders an empty white canvas instead of the chart.
  //
  // Excel's own output for chartsheet drawings uses
  // `<xdr:absoluteAnchor>` with concrete EMU `pos`/`ext` values
  // (≈ 10.84″ × 6.67″ — standard A4 landscape minus default
  // margins) and leaves the inner `<xdr:graphicFrame>/<xdr:xfrm>`
  // extent at zero, which is what Excel writes. Repeating the
  // anchor's extent there is what *caused* the blank-canvas
  // rendering, rather than avoiding it. We emit the same byte
  // layout here verbatim rather than route through `DrawingXform`,
  // which is tuned for the worksheet twoCellAnchor case.
  const CHARTSHEET_EMU_CX = CHARTSHEET_DRAWING_EMU.cx;
  const CHARTSHEET_EMU_CY = CHARTSHEET_DRAWING_EMU.cy;
  for (const cs of model.chartsheets || []) {
    if (cs.drawingName && (cs.chartNumber || cs.chartExNumber)) {
      const chartRId = "rId1";
      const isChartEx = !cs.chartNumber && !!cs.chartExNumber;
      const chartName = isChartEx ? `Chart ${cs.chartExNumber}` : `Chart ${cs.chartNumber}`;
      const drawingXml = renderChartsheetDrawingXml({
        chartRId,
        chartName,
        isChartEx,
        extCx: CHARTSHEET_EMU_CX,
        extCy: CHARTSHEET_EMU_CY
      });
      const drawingRels = [
        {
          Id: chartRId,
          Type: isChartEx ? RelType.ChartEx : RelType.Chart,
          Target: isChartEx
            ? chartExRelTargetFromDrawing(cs.chartExNumber)
            : chartRelTargetFromDrawing(cs.chartNumber)
        }
      ];
      await appendToZip(zip, drawingXml, { name: drawingPath(cs.drawingName) });
      await renderToZip(zip, drawingRelsPath(cs.drawingName), relsXform, drawingRels);
    }
  }
}

async function addTables(zip: IZipWriter, model: any): Promise<void> {
  // Skip (and avoid loading TableXform ~14 KB) when no worksheet has tables.
  const hasTable = model.worksheets.some((ws: any) => ws.tables && ws.tables.length > 0);
  if (!hasTable) {
    return;
  }
  const { TableXform } = await import("@excel/xlsx/xform/table/table-xform");
  const tableXform = new TableXform();

  for (const worksheet of model.worksheets) {
    for (const table of worksheet.tables) {
      tableXform.prepare(table, {});
      await renderToZip(zip, tablePath(table.target), tableXform, table, formulaNames(model));
    }
  }
}

/**
 * Write every external workbook reference into the archive. For each
 * {@link ExternalLinkModel} in `model.externalLinks` we emit two files:
 *
 *   xl/externalLinks/externalLink{index}.xml          — sheet names + cache
 *   xl/externalLinks/_rels/externalLink{index}.xml.rels — target path
 *
 * The target-path rel carries `TargetMode="External"` with a **bare
 * relative** `Target` whenever the user supplied one. This is the single
 * line that makes Office / WPS resolve the referenced workbook relative
 * to the current file's directory (not the `%USERPROFILE%\Documents`
 * fallback) — the root of the relative-path external-link behaviour.
 */
async function addExternalLinks(zip: IZipWriter, model: any): Promise<void> {
  const externalLinks = (model.externalLinks ?? []) as ExternalLinkModel[];
  if (externalLinks.length === 0) {
    return;
  }

  const externalLinkXform = new ExternalLinkXform();
  const relsXform = new RelationshipsXform();

  for (const link of externalLinks) {
    await renderToZip(zip, externalLinkPath(link.index), externalLinkXform, link);

    // `TargetMode="External"` is what flags Office to look the file up at workbook-open time rather than embed it.
    // A link this library authored has one relationship, `rId1`. A link read from a file keeps every relationship
    // it had, because preserved XML names them; the book's own takes the model's target, which a caller may change.
    const preserved = link.preserved;
    const bookRel = {
      Id: preserved?.bookRelationshipId ?? "rId1",
      Type: RelType.ExternalLinkPath,
      Target: link.target,
      TargetMode: link.targetMode ?? "External"
    };
    await renderToZip(
      zip,
      externalLinkRelsPath(link.index),
      relsXform,
      preserved === undefined
        ? [bookRel]
        : preserved.relationships.map(rel =>
            rel.id === bookRel.Id
              ? { ...bookRel, Type: rel.type }
              : {
                  Id: rel.id,
                  Type: rel.type,
                  Target: rel.target,
                  ...(rel.targetMode === undefined ? {} : { TargetMode: rel.targetMode })
                }
          )
    );
  }
}

async function addPivotTables(zip: IZipWriter, model: any): Promise<void> {
  if (!model.pivotTables.length) {
    return;
  }
  // The cache of a pivot table built here is recomputed from its source as the source is now, so values and rows
  // the source gained after `Pivot.add` are in it — the records are read from the source at write time anyway, and
  // a record with a value the shared items lacked was refused.
  for (const pivotTable of model.pivotTables as PivotTable[]) {
    if (!pivotTable.isLoaded) {
      pivotTable.refreshCache?.();
    }
  }

  // Dynamic import: pivot serialisation (~44 KB across the three xforms) is
  // only reachable when the workbook actually contains pivot tables, so it
  // stays out of bundles whose consumers never use pivots.
  const [{ PivotCacheRecordsXform }, { PivotCacheDefinitionXform }, { PivotTableXform }] =
    await Promise.all([
      import("@excel/xlsx/xform/pivot-table/pivot-cache-records-xform"),
      import("@excel/xlsx/xform/pivot-table/pivot-cache-definition-xform"),
      import("@excel/xlsx/xform/pivot-table/pivot-table-xform")
    ]);

  const pivotCacheRecordsXform = new PivotCacheRecordsXform();
  const pivotCacheDefinitionXform = new PivotCacheDefinitionXform();
  const pivotTableXform = new PivotTableXform();
  const relsXform = new RelationshipsXform();

  // R9-B6: Track which cacheIds have already been written to avoid duplicating
  // shared caches. Maps cacheId → tableNumber used for the cache file names.
  const writtenCaches = new Map<string, number>();

  for (const pivotTable of model.pivotTables as PivotTable[]) {
    const n = pivotTable.tableNumber;
    const isLoaded = pivotTable.isLoaded;
    const cacheId = pivotTable.cacheId;

    // R9-B6: Only write cache definition/records/rels once per unique cacheId.
    const cacheAlreadyWritten = writtenCaches.has(cacheId);
    if (!cacheAlreadyWritten) {
      writtenCaches.set(cacheId, n);

      if (isLoaded) {
        if (pivotTable.cacheDefinition) {
          await renderToZip(
            zip,
            pivotCacheDefinitionPath(n),
            pivotCacheDefinitionXform,
            pivotTable.cacheDefinition
          );
        }
        if (pivotTable.cacheRecords) {
          await renderToZip(
            zip,
            pivotCacheRecordsPath(n),
            pivotCacheRecordsXform,
            pivotTable.cacheRecords
          );
        }
      } else {
        await renderToZip(zip, pivotCacheRecordsPath(n), pivotCacheRecordsXform, pivotTable);
        await renderToZip(zip, pivotCacheDefinitionPath(n), pivotCacheDefinitionXform, pivotTable);
      }

      // R9-B4: Only write cache definition rels when cache records exist.
      const hasCacheRecords = isLoaded ? !!pivotTable.cacheRecords : true;
      if (hasCacheRecords) {
        const cacheRecordsRId = (isLoaded ? pivotTable.cacheDefinition?.rId : undefined) ?? "rId1";
        await renderToZip(zip, pivotCacheDefinitionRelsPath(n), relsXform, [
          {
            Id: cacheRecordsRId,
            Type: RelType.PivotCacheRecords,
            Target: pivotCacheRecordsRelTarget(n)
          }
        ]);
      }
    }

    // Pivot table XML is always written (each pivot table has its own file).
    await renderToZip(zip, pivotTablePath(n), pivotTableXform, pivotTable);

    // Pivot table rels point to the cache definition file.
    const cacheTableNumber = writtenCaches.get(cacheId)!;
    await renderToZip(zip, pivotTableRelsPath(n), relsXform, [
      {
        Id: "rId1",
        Type: RelType.PivotCacheDefinition,
        Target: pivotCacheDefinitionRelTargetFromPivotTable(cacheTableNumber)
      }
    ]);
  }
}

function finalizeZip(zip: IZipWriter): Promise<void> {
  return new Promise((resolve, reject) => {
    zip.on("finish", () => {
      resolve();
    });
    zip.on("error", reject);
    zip.finalize();
  });
}

function prepareModel(workbook: Workbook, model: any, options: any): void {
  resolveOpaqueReachability(model);
  model.creator = model.creator ?? "Documonster";
  model.lastModifiedBy = model.lastModifiedBy ?? "Documonster";
  model.created = model.created ?? new Date();
  model.modified = model.modified ?? new Date();

  model.useSharedStrings = options.useSharedStrings !== undefined ? options.useSharedStrings : true;
  model.useStyles = options.useStyles !== undefined ? options.useStyles : true;

  model.sharedStrings = new SharedStringsXform();

  // Preserve default font from parsed styles if available
  const oldDefaultFont = model.defaultFont;
  model.styles = model.useStyles ? new StylesXform(true) : new StylesXform.Mock();
  if (oldDefaultFont && model.styles.setDefaultFont) {
    model.styles.setDefaultFont(oldDefaultFont);
  }
  // Register workbook-level named cell styles so cells referencing them by
  // name resolve to the correct cellStyleXfs index during style building.
  if (model.cellStyles && model.styles.registerNamedStyles) {
    // Keyed as Excel compares the names, without regard to case; see `namedStyleMap`.
    const namedStyles = namedStyleMap(model.cellStyles as NamedStyleEntry[]);
    model.styles.registerNamedStyles(namedStyles);
  }
  // Before any worksheet is prepared: every `addDxfStyle` must see the source table already in place.
  if (model.dxfs && model.styles.seedDxfs) {
    model.styles.seedDxfs(model.dxfs);
  }

  const workbookXform = new WorkbookXform();
  const worksheetXform = new WorkSheetXform();

  workbookXform.prepare(model);

  // Normalise external-workbook references before any formula rendering.
  // Two jobs:
  //   1. Scan every formula cell and make sure each referenced workbook
  //      has a matching ExternalLinkModel in `model.externalLinks`, with
  //      a stable 1-based index.
  //   2. Rewrite formula strings from `[filename.xlsx]Sheet!A1` to
  //      `[N]Sheet!A1`, which is the canonical on-disk form Excel
  //      expects inside `<f>` elements.
  //
  // Done once up-front (not per cell) so the index assignment is
  // deterministic and every cell sees the final externalLinks list.
  normaliseExternalLinks(workbook, model);

  const worksheetOptions: any = {
    needsArrayEvaluation: options.needsArrayEvaluation,
    unwritableFormula: options.unwritableFormula,
    sharedStrings: model.sharedStrings,
    styles: model.styles,
    date1904: model.properties?.date1904,
    drawingsCount: 0,
    media: model.media
  };
  // Seed the drawing-name allocator past any drawing name a worksheet already
  // carries verbatim from load (`drawingN`). WorkSheetXform.prepare() only
  // auto-names a drawing when its `name` is empty, using `drawing${++count}`.
  // If one sheet keeps a round-tripped "drawing1" while another needs a fresh
  // name, a count starting at 0 would hand out "drawing1" again — a duplicate
  // `/xl/drawings/drawing1.xml` Override (invalid per OPC). Starting the count
  // above the highest existing numeric drawing name avoids the collision
  // regardless of how the mix arose (e.g. Workbook.importWorksheet).
  let maxExistingDrawingNo = 0;
  for (const ws of model.worksheets as any[]) {
    const m = /^drawing(\d+)$/.exec(ws?.drawing?.name ?? "");
    if (m) {
      maxExistingDrawingNo = Math.max(maxExistingDrawingNo, Number(m[1]));
    }
  }
  for (const cs of model.chartsheets ?? []) {
    const m = /^drawing(\d+)$/.exec(cs?.drawingName ?? "");
    if (m) {
      maxExistingDrawingNo = Math.max(maxExistingDrawingNo, Number(m[1]));
    }
  }
  worksheetOptions.drawingsCount = maxExistingDrawingNo;
  worksheetOptions.drawings = model.drawings = [];
  worksheetOptions.commentRefs = model.commentRefs = [];
  worksheetOptions.formControlRefs = model.formControlRefs = [];
  // Collect the list of worksheets that carry Office 365 threaded
  // comments so the Content Types override list can include them
  // and the ZIP writer knows which per-sheet parts to emit. Sheets
  // with zero threaded comments are skipped entirely — Excel treats
  // a missing part as "no threaded comments on this sheet".
  model.threadedCommentSheetIds = [] as Array<number | string>;
  model.hasPersons = (model.persons?.length ?? 0) > 0;
  // Raw-passthrough parts captured on load. The Content-Types
  // override list and the content-types writer need these path
  // lists so the emitted bytes are registered in the package.
  model.slicerPartPaths = Object.keys(model.slicerParts ?? {}).filter(p => !p.includes("/_rels/"));
  model.slicerCachePartPaths = Object.keys(model.slicerCacheParts ?? {}).filter(
    p => !p.includes("/_rels/")
  );
  model.timelinePartPaths = Object.keys(model.timelineParts ?? {}).filter(
    p => !p.includes("/_rels/")
  );
  model.timelineCachePartPaths = Object.keys(model.timelineCacheParts ?? {}).filter(
    p => !p.includes("/_rels/")
  );
  model.hasHeaderWatermark = false;
  // VML shape ids are numbered across the workbook, so every sheet's parts draw from one space. A
  // checkbox read from a file keeps its id when it can; reserving its block first stops an earlier
  // sheet's notes from taking it.
  worksheetOptions.vmlShapeIds = createVmlShapeIdSpace();
  for (const worksheet of model.worksheets) {
    reserveVmlShapeIds(
      worksheetOptions.vmlShapeIds,
      worksheet,
      [...(worksheet.formControls ?? []), ...(worksheet.legacyControls ?? [])].map(
        (control: { shapeId: number }) => control.shapeId
      )
    );
  }
  let tableCount = 0;
  model.tables = [];
  const tableNameMap = new Map<string, string>(); // name (excelNameKey) → worksheet name
  model.worksheets.forEach((worksheet: any, index: number) => {
    // Assign fileIndex early so that worksheet-xform.prepare() can use it
    // for comment/VML relationship targets and content type names.
    // This ensures consistency with addWorksheets() which writes ZIP entries
    // using the same fileIndex.
    worksheet.fileIndex = index + 1;

    worksheet.tables.forEach((table: any) => {
      // OOXML requires table names to be unique across the entire workbook
      // (case-insensitive). Detect duplicates early to produce a clear error
      // instead of generating a corrupt file that Excel must repair.
      const nameKey = excelNameKey(table.name);
      const existingSheet = tableNameMap.get(nameKey);
      if (existingSheet !== undefined) {
        throw new TableError(
          `Duplicate table name "${table.name}": already used in worksheet "${existingSheet}". ` +
            `Table names must be unique across the entire workbook (case-insensitive).`
        );
      }
      tableNameMap.set(nameKey, worksheet.name);

      tableCount++;
      table.target = `table${tableCount}.xml`;
      table.id = tableCount;
      model.tables.push(table);
    });

    worksheetXform.prepare(worksheet, worksheetOptions);
    // Register sheets that carry threaded comments so the Content
    // Types override list and the zip emission loop find them.
    if (worksheet.threadedComments && worksheet.threadedComments.length > 0) {
      (model.threadedCommentSheetIds as Array<number | string>).push(worksheet.fileIndex);
    }
  });

  // ContentTypesXform expects this flag
  model.hasCheckboxes = model.styles.hasCheckboxes;

  // Scan all worksheets for dynamic array formulas.
  // cm=1 is assigned later by cell-xform.prepare() during worksheet rendering.
  let dynamicArrayCount = 0;
  model.worksheets.forEach((worksheet: any) => {
    (worksheet.rows ?? []).forEach((row: any) => {
      (row.cells ?? []).forEach((cell: any) => {
        if (writesDynamicArray(cell)) {
          dynamicArrayCount++;
        }
      });
    });
  });
  model.hasDynamicArrayFormulas = dynamicArrayCount > 0;
  model.dynamicArrayCount = dynamicArrayCount;

  // Propagate header watermark flag from worksheet prepare options
  if (worksheetOptions.hasHeaderWatermark) {
    model.hasHeaderWatermark = true;
  }
}

function prepareChartsheets(model: any): void {
  if (!model.chartsheets || model.chartsheets.length === 0) {
    return;
  }

  const usedDrawingNumbers = new Set<number>();
  for (const drawing of model.drawings ?? []) {
    const match = /^drawing(\d+)$/.exec(drawing.name ?? "");
    if (match) {
      usedDrawingNumbers.add(parseInt(match[1], 10));
    }
  }
  for (const cs of model.chartsheets) {
    const existingMatch = /^drawing(\d+)$/.exec(cs.drawingName ?? "");
    if (existingMatch) {
      usedDrawingNumbers.add(parseInt(existingMatch[1], 10));
    }
  }

  const nextDrawingName = (): string => {
    let n = 1;
    while (usedDrawingNumbers.has(n)) {
      n++;
    }
    usedDrawingNumbers.add(n);
    return `drawing${n}`;
  };

  for (const cs of model.chartsheets) {
    if (!cs.drawingName) {
      cs.drawingName = nextDrawingName();
    }
    if (!cs.drawing) {
      cs.drawing = { rId: "rId1" };
    }
    if (!model.drawings.some((drawing: any) => drawing.name === cs.drawingName)) {
      model.drawings.push({ name: cs.drawingName });
    }
  }

  // Signal the content-types writer that the `Default Extension="vml"`
  // declaration is required when ANY chartsheet carries a VML
  // relationship (e.g. `<legacyDrawing r:id="…"/>` referencing a
  // preserved `xl/drawings/vmlDrawing*.vml` part). Previously the
  // flag was only set inside `addChartsheets`, which runs AFTER
  // `addContentTypes` — so a chartsheet-only VML dependency silently
  // shipped without its content-type declaration, and Excel refused
  // to open the resulting package. Compute it here, during `prepare`,
  // before any part is written.
  model.hasChartsheetVml = model.chartsheets.some(
    (cs: any) =>
      cs.headerImages?.length > 0 ||
      cs.relationships?.some(
        (rel: any) =>
          rel?.Type === RelType.VmlDrawing &&
          typeof rel.Target === "string" &&
          model.vmlDrawings?.[resolveRelTarget("xl/chartsheets/", rel.Target)] !== undefined
      )
  );
}
