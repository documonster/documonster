import type {
  CacheField,
  DataField,
  PivotTableChartFormat,
  PivotTableSubtotal,
  SharedItemValue
} from "@excel/pivot-table-types";
import { VALID_SUBTOTALS, METRIC_DISPLAY_NAMES } from "@excel/pivot-table-types";
import { colCache } from "@excel/utils/col-cache";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import {
  RawXmlCollector,
  serializeAttributes
} from "@excel/xlsx/xform/pivot-table/raw-xml-collector";
import type { XmlSink } from "@xml/types";
import { StdDocAttributes } from "@xml/writer";

/** OOXML sentinel field index meaning "data values" pseudo-field (used in pivotArea references) */
const FIELD_INDEX_DATA_VALUES = 4294967294; // 0xFFFFFFFE

/**
 * Signed sentinel for the "Values" pseudo-field in colFields/rowFields.
 * OOXML represents this as x="-2" (signed int32 of 0xFFFFFFFE).
 */
const VALUES_FIELD_INDEX = -2;

/** Default pivot table style name */
const DEFAULT_PIVOT_STYLE = "PivotStyleLight16";

/** Valid OOXML axis values for pivot fields */
const VALID_PIVOT_AXES = new Set(["axisRow", "axisCol", "axisPage", "axisValues"]);

/** PivotFieldItem attribute keys — used to build the attrs object for rendering */
const PIVOT_FIELD_ITEM_KEYS = ["x", "t", "h", "sd", "f", "m", "c", "d"] as const;

/**
 * Model for generating pivot table (with live source)
 */
interface PivotTableRenderModel {
  rows: number[];
  columns: number[];
  values: number[];
  pages?: number[];
  metric: PivotTableSubtotal;
  /** Per-value metric overrides (parallel to `values` array). Falls back to `metric`. */
  valueMetrics: PivotTableSubtotal[];
  cacheFields: CacheField[];
  cacheId: string | number;
  tableNumber: number;
  applyWidthHeightFormats: "0" | "1";
  /**
   * Optional top-left anchor (e.g. `"A3"`) for the pivot's displayed block.
   * When set, the rendered `<location>` ref is computed from this anchor
   * instead of the default `A{3 + pageOffset}`. The anchor represents the
   * pivot's display origin — page filters (if any) occupy rows downward from
   * the anchor, followed by a blank separator row and then the pivot body.
   */
  ref?: string;
  name?: string;
  chartFormat?: number;
  chartFormats?: PivotTableChartFormat[];
}

/**
 * A single <item> element inside a pivotField's <items> collection.
 * Items can reference a shared item index (x attribute) or indicate a subtotal type (t attribute).
 */
interface PivotFieldItem {
  /** Shared item index (maps to cacheField sharedItems) */
  x?: number;
  /** Item type: "default" for subtotals, "sum", "count", "avg", "max", "min", "grand", etc. */
  t?: string;
  /** Hidden flag — "1" means item is filtered out */
  h?: string;
  /** Show details (hide details when "0") */
  sd?: string;
  /** Calculated item flag */
  f?: string;
  /** Missing flag */
  m?: string;
  /** Child items flag */
  c?: string;
  /** Drill across flag */
  d?: string;
}

/**
 * Parsed pivot field
 */
interface ParsedPivotField {
  axis?: "axisRow" | "axisCol" | "axisPage" | "axisValues";
  dataField?: boolean;
  items?: PivotFieldItem[];
  compact?: boolean;
  outline?: boolean;
  showAll?: boolean;
  defaultSubtotal?: boolean;
  numFmtId?: number;
  sortType?: string;
  autoSortScopeXml?: string;
  subtotalTop?: boolean;
  insertBlankRow?: boolean;
  multipleItemSelectionAllowed?: boolean;
  /** Bag of additional attributes not individually modeled (for roundtrip preservation) */
  extraAttrs?: Record<string, string>;
}

/**
 * Parsed page field (report filter)
 */
interface ParsedPageField {
  fld: number;
  item?: number;
  hier?: number;
  name?: string;
}

/**
 * Parsed pivot table model (loaded from file)
 */
interface ParsedPivotTableModel {
  // Core identifiers
  name?: string;
  cacheId: number;
  uid?: string;

  // Location info
  location?: {
    ref: string;
    firstHeaderRow?: number;
    firstDataRow?: number;
    firstDataCol?: number;
    rowPageCount?: number;
    colPageCount?: number;
  };

  // Field configurations
  pivotFields: ParsedPivotField[];
  rowFields: number[]; // Field indices for rows
  colFields: number[]; // Field indices for columns
  pageFields: ParsedPageField[]; // Page fields (report filters)
  dataFields: DataField[];

  // Style and formatting
  applyNumberFormats?: string;
  applyBorderFormats?: string;
  applyFontFormats?: string;
  applyPatternFormats?: string;
  applyAlignmentFormats?: string;
  applyWidthHeightFormats?: string;
  dataCaption?: string;
  styleName?: string;
  /** Full pivotTableStyleInfo attributes (preserved on roundtrip) */
  styleInfo?: {
    name?: string;
    showRowHeaders?: string;
    showColHeaders?: string;
    showRowStripes?: string;
    showColStripes?: string;
    showLastColumn?: string;
  };

  // Version info
  updatedVersion?: string;
  minRefreshableVersion?: string;
  createdVersion?: string;

  // Other attributes
  useAutoFormatting?: string;
  itemPrintTitles?: string;
  indent?: number;
  compact?: boolean;
  compactData?: boolean;
  multipleFieldFilters?: string;
  outline?: boolean;
  outlineData?: boolean;
  chartFormat?: number;

  // Grand totals and display options
  colGrandTotals?: string;
  rowGrandTotals?: string;
  showError?: string;
  errorCaption?: string;
  showMissing?: string;
  missingCaption?: string;
  grandTotalCaption?: string;

  // Row/col items (for grand totals etc)
  rowItems?: RowColItem[];
  colItems?: RowColItem[];

  // Track if colFields element was present in original file
  // Some pivot tables don't have colFields even when they have colItems
  hasColFields?: boolean;

  // Track if rowItems/colItems elements were present in original file
  // Prevents injecting default grand total items when the original had none
  hasRowItems?: boolean;
  hasColItems?: boolean;

  // Chart formats (for pivot charts)
  chartFormats?: ChartFormatItem[];

  // Preserved extLst XML for roundtrip (loaded models only)
  extLstXml?: string;

  // Preserved <formats> XML for roundtrip (loaded models only)
  formatsXml?: string;
  // Preserved <conditionalFormats> XML for roundtrip (loaded models only)
  conditionalFormatsXml?: string;

  // Preserved <filters> XML for roundtrip (loaded models only)
  filtersXml?: string;

  // Preserved unknown top-level child elements XML for roundtrip (loaded models only)
  unknownElementsXml?: string;

  // Flag indicating this was loaded from file
  isLoaded?: boolean;
}

/**
 * Row or column item in pivot table
 */
interface RowColItem {
  t?: string; // type (e.g., "grand" for grand total)
  r?: number; // repeated item count
  i?: number; // data field index
  x: Array<{ v: number }>; // x element values
}

/**
 * Chart format item for pivot charts
 */
interface ChartFormatItem extends PivotTableChartFormat {
  chart: number;
  format: number;
  series?: boolean;
  // Preserved pivotArea XML for round-trip
  pivotAreaXml?: string;
}

/**
 * Top-level section currently being parsed inside <pivotTableDefinition>.
 * These sections are mutually exclusive (only one can be open at a time).
 */
type PivotSection =
  | "pivotFields"
  | "rowFields"
  | "colFields"
  | "pageFields"
  | "dataFields"
  | "rowItems"
  | "colItems"
  | "chartFormats";

/**
 * Parser state for PivotTableXform
 */
interface ParserState {
  /** Which top-level section we are currently inside, or null if between sections. */
  currentSection: PivotSection | null;
  /** Whether we are inside a <pivotArea> element (nested inside chartFormats or autoSortScope). */
  inPivotArea: boolean;
  /** Whether we are inside an <autoSortScope> element (nested inside a pivotField). */
  inAutoSortScope: boolean;
}

/** Factory for default ParserState values */
function createDefaultParserState(): ParserState {
  return {
    currentSection: null,
    inPivotArea: false,
    inAutoSortScope: false
  };
}

/** Known pivotField attributes that we parse individually (hoisted to module scope) */
const KNOWN_PIVOT_FIELD_KEYS = new Set([
  "axis",
  "dataField",
  "compact",
  "outline",
  "showAll",
  "defaultSubtotal",
  "numFmtId",
  "sortType",
  "subtotalTop",
  "insertBlankRow",
  "multipleItemSelectionAllowed"
]);

class PivotTableXform extends BaseXform<ParsedPivotTableModel | null> {
  // Parser state consolidated into object for easier reset
  private state: ParserState = createDefaultParserState();

  // Current parsing context
  private currentPivotField: ParsedPivotField | null = null;
  private currentRowItem: RowColItem | null = null;
  private currentColItem: RowColItem | null = null;
  private currentChartFormat: ChartFormatItem | null = null;
  // Buffer for collecting pivotArea XML
  private pivotAreaXmlBuffer: string[] = [];
  // Buffer for collecting autoSortScope XML
  private autoSortScopeXmlBuffer: string[] = [];
  // Raw XML collectors (replacing manual in/depth/buffer triples)
  private extLstCollector = new RawXmlCollector("extLst");
  private formatsCollector = new RawXmlCollector("formats");
  private conditionalFormatsCollector = new RawXmlCollector("conditionalFormats");
  private filtersCollector = new RawXmlCollector("filters");
  private unknownCollector = new RawXmlCollector("");
  // Accumulated unknown elements XML strings (one per element)
  private unknownElementsXmlParts: string[] = [];

  constructor() {
    super();
    this.model = null;
  }

  get tag(): string {
    // http://www.datypic.com/sc/ooxml/e-ssml_pivotTableDefinition.html
    return "pivotTableDefinition";
  }

  reset(): void {
    this.model = null;
    // Reset all parser state flags
    this.state = createDefaultParserState();
    // Reset current context
    this.currentPivotField = null;
    this.currentRowItem = null;
    this.currentColItem = null;
    this.currentChartFormat = null;
    this.pivotAreaXmlBuffer = [];
    this.autoSortScopeXmlBuffer = [];
    this.extLstCollector.reset();
    this.formatsCollector.reset();
    this.conditionalFormatsCollector.reset();
    this.filtersCollector.reset();
    this.unknownCollector.reset();
    this.unknownElementsXmlParts = [];
  }

  /**
   * Render pivot table XML.
   * Supports both newly created models and loaded models.
   */
  render(xmlStream: XmlSink, model: PivotTableRenderModel | ParsedPivotTableModel): void {
    const isLoaded = "isLoaded" in model && model.isLoaded;

    if (isLoaded) {
      this.renderLoaded(xmlStream, model as ParsedPivotTableModel);
    } else {
      this.renderNew(xmlStream, model as PivotTableRenderModel);
    }
  }

  /**
   * Render newly created pivot table
   */
  private renderNew(xmlStream: XmlSink, model: PivotTableRenderModel): void {
    const {
      rows,
      columns,
      values,
      pages = [],
      cacheFields,
      cacheId,
      tableNumber,
      applyWidthHeightFormats
    } = model;

    // Multi-value with no explicit columns: the "Values" pseudo-field occupies the column axis
    const isMultiValueNoCol = columns.length === 0 && values.length > 1;

    // Page fields offset: each page field adds 1 row above the pivot table,
    // plus 1 blank separator row when any page fields are present.
    const pageCount = pages.length;
    const pageOffset = pageCount > 0 ? pageCount + 1 : 0;

    // Location ref: firstDataCol = number of row fields (row label columns),
    // endCol = row fields + data columns.
    // When the caller supplies `ref`, it anchors the pivot's displayed block
    // (page filters → blank separator → pivot body). The body — which is what
    // <location ref="..."/> actually addresses in OOXML — therefore sits
    // pageOffset rows below the anchor and shares its column.
    const firstDataCol = rows.length;
    const dataColCount = isMultiValueNoCol ? values.length : 1;
    const totalPivotCols = firstDataCol + dataColCount;

    let startRow: number;
    let startCol: number;
    if (model.ref) {
      // `ref` is pre-normalised by makePivotTable to a canonical cell address,
      // so decodeAddress cannot fail here. Guard anyway to surface any
      // bypasses via `as any` with a clear error.
      const addr = colCache.decodeAddress(model.ref);
      if (!addr.col || !addr.row) {
        throw new Error(
          `Pivot table ref "${model.ref}" must include both column and row (e.g. "A3").`
        );
      }
      startRow = addr.row + pageOffset;
      startCol = addr.col;
    } else {
      startRow = 3 + pageOffset;
      startCol = 1;
    }
    const endRow = startRow + 1; // header + 1 data row placeholder
    const endCol = startCol + totalPivotCols - 1;
    const startColLetter = colCache.n2l(startCol);
    const endColLetter = colCache.n2l(endCol);
    const locationRef = `${startColLetter}${startRow}:${endColLetter}${endRow}`;

    xmlStream.openXml(StdDocAttributes);
    xmlStream.openNode(this.tag, {
      ...PivotTableXform.PIVOT_TABLE_ATTRIBUTES,
      name: model.name ?? `PivotTable${tableNumber}`,
      cacheId,
      applyNumberFormats: "0",
      applyBorderFormats: "0",
      applyFontFormats: "0",
      applyPatternFormats: "0",
      applyAlignmentFormats: "0",
      applyWidthHeightFormats,
      dataCaption: "Values",
      updatedVersion: "8",
      minRefreshableVersion: "3",
      useAutoFormatting: "1",
      itemPrintTitles: "1",
      createdVersion: "8",
      indent: "0",
      compact: "0",
      compactData: "0",
      multipleFieldFilters: "0",
      chartFormat: model.chartFormat !== undefined ? String(model.chartFormat) : undefined
    });

    // Location
    const locAttrs: Record<string, string | number> = {
      ref: locationRef,
      firstHeaderRow: 1,
      firstDataRow: 1,
      firstDataCol
    };
    if (pageCount > 0) {
      locAttrs.rowPageCount = pageCount;
      locAttrs.colPageCount = 1;
    }
    xmlStream.leafNode("location", locAttrs);

    // Pivot fields
    renderPivotFields(xmlStream, model);

    // Row fields
    xmlStream.openNode("rowFields", { count: rows.length });
    for (const rowIndex of rows) {
      xmlStream.leafNode("field", { x: rowIndex });
    }
    xmlStream.closeNode();

    // Row items: minimal grand total row. refreshOnLoad="1" causes Excel
    // to rebuild the full row expansion on open.
    xmlStream.openNode("rowItems", { count: 1 });
    xmlStream.openNode("i", { t: "grand" });
    xmlStream.leafNode("x");
    xmlStream.closeNode(); // i
    xmlStream.closeNode(); // rowItems

    // colFields: lists the field indices on the column axis.
    // When columns is non-empty, list those field indices.
    // When columns is empty but there are multiple values, emit the synthetic
    // "Values" pseudo-field (field x="-2") so Excel knows where to position
    // the data field labels on the column axis.
    // When columns is empty and there is only one value, omit colFields entirely.
    if (columns.length > 0) {
      const fieldCount = values.length > 1 ? columns.length + 1 : columns.length;
      xmlStream.openNode("colFields", { count: fieldCount });
      for (const colIndex of columns) {
        xmlStream.leafNode("field", { x: colIndex });
      }
      if (values.length > 1) {
        xmlStream.leafNode("field", { x: VALUES_FIELD_INDEX });
      }
      xmlStream.closeNode();
    } else if (isMultiValueNoCol) {
      xmlStream.openNode("colFields", { count: 1 });
      xmlStream.leafNode("field", { x: VALUES_FIELD_INDEX });
      xmlStream.closeNode();
    }

    // colItems: for multi-value no-column pivots, one <i> per value field (referencing
    // its index in dataFields via <x v="N"/>) plus a grand total <i>.
    // For single-value or explicit-columns pivots, a single empty <i/>.
    // These are required by Excel — omitting them causes "Repaired Records" errors.
    if (isMultiValueNoCol) {
      xmlStream.openNode("colItems", { count: values.length + 1 });
      for (let idx = 0; idx < values.length; idx++) {
        xmlStream.openNode("i");
        xmlStream.leafNode("x", idx === 0 ? undefined : { v: idx });
        xmlStream.closeNode(); // i
      }
      xmlStream.openNode("i", { t: "grand" });
      xmlStream.leafNode("x");
      xmlStream.closeNode(); // i
      xmlStream.closeNode(); // colItems
    } else {
      xmlStream.openNode("colItems", { count: 1 });
      xmlStream.leafNode("i");
      xmlStream.closeNode();
    }

    // Page fields (between colItems and dataFields per OOXML spec)
    if (pageCount > 0) {
      xmlStream.openNode("pageFields", { count: pageCount });
      for (const fld of pages) {
        xmlStream.leafNode("pageField", { fld, hier: -1 });
      }
      xmlStream.closeNode();
    }

    // Data fields
    renderDataFields(xmlStream, cacheFields, values, model.valueMetrics);

    if (model.chartFormats && model.chartFormats.length > 0) {
      this.renderChartFormats(xmlStream, model.chartFormats);
    }

    // Pivot table style info
    xmlStream.leafNode("pivotTableStyleInfo", {
      name: DEFAULT_PIVOT_STYLE,
      showRowHeaders: "1",
      showColHeaders: "1",
      showRowStripes: "0",
      showColStripes: "0",
      showLastColumn: "1"
    });

    // Extensions
    xmlStream.writeRaw(PivotTableXform.EXTLST_XML);

    xmlStream.closeNode();
  }

  /**
   * Render loaded pivot table (preserving original structure)
   */
  private renderLoaded(xmlStream: XmlSink, model: ParsedPivotTableModel): void {
    const attrs = this.buildLoadedRootAttributes(model);

    xmlStream.openXml(StdDocAttributes);
    xmlStream.openNode(this.tag, attrs);

    // Location
    if (model.location) {
      const locAttrs: Record<string, string | number | undefined> = {
        ref: model.location.ref,
        firstHeaderRow: model.location.firstHeaderRow,
        firstDataRow: model.location.firstDataRow,
        firstDataCol: model.location.firstDataCol
      };
      if (model.location.rowPageCount !== undefined) {
        locAttrs.rowPageCount = model.location.rowPageCount;
      }
      if (model.location.colPageCount !== undefined) {
        locAttrs.colPageCount = model.location.colPageCount;
      }
      xmlStream.leafNode("location", locAttrs);
    }

    // Pivot fields
    if (model.pivotFields.length > 0) {
      xmlStream.openNode("pivotFields", { count: model.pivotFields.length });
      for (const pivotField of model.pivotFields) {
        this.renderPivotFieldLoaded(xmlStream, pivotField);
      }
      xmlStream.closeNode();
    }

    // Row fields
    if (model.rowFields.length > 0) {
      xmlStream.openNode("rowFields", { count: model.rowFields.length });
      for (const fieldIndex of model.rowFields) {
        xmlStream.leafNode("field", { x: fieldIndex });
      }
      xmlStream.closeNode();
    }

    // Row items - use parsed items if available; otherwise emit a minimal grand total
    if (model.rowItems && model.rowItems.length > 0) {
      xmlStream.openNode("rowItems", { count: model.rowItems.length });
      for (const item of model.rowItems) {
        this.renderRowColItem(xmlStream, item);
      }
      xmlStream.closeNode();
    } else if (model.hasRowItems) {
      xmlStream.writeRaw('<rowItems count="1"><i t="grand"><x/></i></rowItems>');
    }

    // Col fields
    // Only render colFields if it was present in the original file or if there are actual column fields
    // Some pivot tables don't have colFields element at all
    if (model.hasColFields || model.colFields.length > 0) {
      if (model.colFields.length === 0 && model.dataFields.length <= 1) {
        // Empty colFields with no multi-value need — preserve as empty element
        xmlStream.leafNode("colFields", { count: 0 });
      } else {
        const colFieldCount = model.colFields.length === 0 ? 1 : model.colFields.length;
        xmlStream.openNode("colFields", { count: colFieldCount });
        if (model.colFields.length === 0) {
          xmlStream.leafNode("field", { x: VALUES_FIELD_INDEX });
        } else {
          for (const fieldIndex of model.colFields) {
            xmlStream.leafNode("field", { x: fieldIndex });
          }
        }
        xmlStream.closeNode();
      }
    }

    // Col items - use parsed items if available
    if (model.colItems && model.colItems.length > 0) {
      xmlStream.openNode("colItems", { count: model.colItems.length });
      for (const item of model.colItems) {
        this.renderRowColItem(xmlStream, item);
      }
      xmlStream.closeNode();
    } else if (model.hasColItems) {
      xmlStream.writeRaw('<colItems count="1"><i t="grand"><x/></i></colItems>');
    }

    // Page fields (report filters)
    if (model.pageFields && model.pageFields.length > 0) {
      xmlStream.openNode("pageFields", { count: model.pageFields.length });
      for (const pf of model.pageFields) {
        const pfAttrs: Record<string, string | number> = { fld: pf.fld };
        if (pf.item !== undefined) {
          pfAttrs.item = pf.item;
        }
        if (pf.hier !== undefined) {
          pfAttrs.hier = pf.hier;
        }
        if (pf.name !== undefined) {
          pfAttrs.name = pf.name;
        }
        xmlStream.leafNode("pageField", pfAttrs);
      }
      xmlStream.closeNode();
    }

    // Data fields
    if (model.dataFields.length > 0) {
      xmlStream.openNode("dataFields", { count: model.dataFields.length });
      for (const dataField of model.dataFields) {
        const dfAttrs: Record<string, string | number> = {
          name: dataField.name,
          fld: dataField.fld
        };
        if (dataField.baseField !== undefined) {
          dfAttrs.baseField = dataField.baseField;
        }
        if (dataField.baseItem !== undefined) {
          dfAttrs.baseItem = dataField.baseItem;
        }
        if (dataField.subtotal !== undefined && dataField.subtotal !== "sum") {
          dfAttrs.subtotal = dataField.subtotal;
        }
        if (dataField.numFmtId !== undefined) {
          dfAttrs.numFmtId = dataField.numFmtId;
        }
        xmlStream.leafNode("dataField", dfAttrs);
      }
      xmlStream.closeNode();
    }

    // Formats — preserved raw XML from loaded file
    if (model.formatsXml) {
      xmlStream.writeRaw(model.formatsXml);
    }

    // Conditional formats — preserved raw XML from loaded file
    // OOXML order: formats → conditionalFormats → chartFormats
    if (model.conditionalFormatsXml) {
      xmlStream.writeRaw(model.conditionalFormatsXml);
    }

    // Chart formats (for pivot charts) - preserve original pivotArea XML
    if (model.chartFormats && model.chartFormats.length > 0) {
      this.renderChartFormats(xmlStream, model.chartFormats);
    }

    // Style info
    const si = model.styleInfo;
    xmlStream.leafNode("pivotTableStyleInfo", {
      name: si?.name ?? model.styleName ?? DEFAULT_PIVOT_STYLE,
      showRowHeaders: si?.showRowHeaders ?? "1",
      showColHeaders: si?.showColHeaders ?? "1",
      showRowStripes: si?.showRowStripes ?? "0",
      showColStripes: si?.showColStripes ?? "0",
      showLastColumn: si?.showLastColumn ?? "1"
    });

    // Filters — preserved raw XML from loaded file
    // <filters> appears between pivotTableStyleInfo and extLst per OOXML schema
    if (model.filtersXml) {
      xmlStream.writeRaw(model.filtersXml);
    }

    // Unknown top-level elements — preserved raw XML for roundtrip
    if (model.unknownElementsXml) {
      xmlStream.writeRaw(model.unknownElementsXml);
    }

    // Extensions — use preserved XML from loaded file; only inject default for new tables
    const extLstXml = model.extLstXml ?? (model.isLoaded ? "" : PivotTableXform.EXTLST_XML);
    if (extLstXml) {
      xmlStream.writeRaw(extLstXml);
    }

    xmlStream.closeNode();
  }

  /**
   * Build the root `<pivotTableDefinition>` attributes for a loaded (roundtrip) model.
   * Extracted from renderLoaded to keep the render method focused on element structure.
   */
  private buildLoadedRootAttributes(model: ParsedPivotTableModel): Record<string, string> {
    const attrs: Record<string, string> = {
      ...PivotTableXform.PIVOT_TABLE_ATTRIBUTES,
      name: model.name ?? "PivotTable1",
      cacheId: String(model.cacheId),
      applyNumberFormats: model.applyNumberFormats ?? "0",
      applyBorderFormats: model.applyBorderFormats ?? "0",
      applyFontFormats: model.applyFontFormats ?? "0",
      applyPatternFormats: model.applyPatternFormats ?? "0",
      applyAlignmentFormats: model.applyAlignmentFormats ?? "0",
      applyWidthHeightFormats: model.applyWidthHeightFormats ?? "0",
      dataCaption: model.dataCaption ?? "Values",
      updatedVersion: model.updatedVersion ?? "8",
      minRefreshableVersion: model.minRefreshableVersion ?? "3"
    };

    // Only emit these boolean-style attributes when they were present in the original.
    // Absent means the OOXML default applies; emitting "0" explicitly changes semantics.
    // Placed before createdVersion to match Excel's attribute ordering.
    if (model.useAutoFormatting !== undefined) {
      attrs.useAutoFormatting = model.useAutoFormatting;
    }
    if (model.itemPrintTitles !== undefined) {
      attrs.itemPrintTitles = model.itemPrintTitles;
    }
    if (model.multipleFieldFilters !== undefined) {
      attrs.multipleFieldFilters = model.multipleFieldFilters;
    }

    attrs.createdVersion = model.createdVersion ?? "8";
    if (model.indent !== undefined) {
      attrs.indent = String(model.indent);
    }

    // Preserve xr:uid on roundtrip
    if (model.uid) {
      attrs["xmlns:xr"] = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
      attrs["xr:uid"] = model.uid;
    }

    // Add outline attributes if present
    if (model.outline) {
      attrs.outline = "1";
    }
    if (model.outlineData) {
      attrs.outlineData = "1";
    }
    if (model.chartFormat !== undefined) {
      attrs.chartFormat = String(model.chartFormat);
    }
    // Grand totals and display option attributes — only emit when present in original
    if (model.colGrandTotals !== undefined) {
      attrs.colGrandTotals = model.colGrandTotals;
    }
    if (model.rowGrandTotals !== undefined) {
      attrs.rowGrandTotals = model.rowGrandTotals;
    }
    if (model.showError !== undefined) {
      attrs.showError = model.showError;
    }
    if (model.errorCaption !== undefined) {
      attrs.errorCaption = model.errorCaption;
    }
    if (model.showMissing !== undefined) {
      attrs.showMissing = model.showMissing;
    }
    if (model.missingCaption !== undefined) {
      attrs.missingCaption = model.missingCaption;
    }
    if (model.grandTotalCaption !== undefined) {
      attrs.grandTotalCaption = model.grandTotalCaption;
    }
    // Only write compact/compactData when false (non-default).
    // OOXML spec: absent = true (default). So if the original file had compact="0",
    // we must preserve it; omitting it would change semantics from false to true.
    if (model.compact === false) {
      attrs.compact = "0";
    }
    if (model.compactData === false) {
      attrs.compactData = "0";
    }

    return attrs;
  }

  /**
   * Render `<chartFormats>` with preserved pivotArea XML for pivot chart roundtrip.
   */
  private renderChartFormats(xmlStream: XmlSink, chartFormats: ChartFormatItem[]): void {
    xmlStream.openNode("chartFormats", { count: chartFormats.length });
    for (const cf of chartFormats) {
      xmlStream.openNode("chartFormat", {
        chart: cf.chart,
        format: cf.format,
        series: cf.series === true ? "1" : cf.series === false ? "0" : undefined
      });
      // Use preserved pivotArea XML or fallback to default
      if (cf.pivotAreaXml) {
        xmlStream.writeRaw(cf.pivotAreaXml);
      } else {
        // Fallback for newly created chart formats (shouldn't happen for loaded models)
        xmlStream.writeRaw(
          `<pivotArea type="data" outline="0" fieldPosition="0"><references count="1"><reference field="${FIELD_INDEX_DATA_VALUES}" count="1" selected="0"><x v="0"/></reference></references></pivotArea>`
        );
      }
      xmlStream.closeNode();
    }
    xmlStream.closeNode();
  }

  /**
   * Render a row or column item element
   */
  private renderRowColItem(xmlStream: XmlSink, item: RowColItem): void {
    const attrs: Record<string, string | number> = {};
    if (item.t !== undefined) {
      attrs.t = item.t;
    }
    if (item.r !== undefined) {
      attrs.r = item.r;
    }
    if (item.i !== undefined) {
      attrs.i = item.i;
    }

    if (item.x.length > 0) {
      xmlStream.openNode("i", attrs);
      for (const x of item.x) {
        if (x.v !== 0) {
          xmlStream.leafNode("x", { v: x.v });
        } else {
          xmlStream.leafNode("x");
        }
      }
      xmlStream.closeNode();
    } else {
      // Empty item (like <i/> in colItems)
      xmlStream.leafNode("i", attrs);
    }
  }

  /**
   * Render a loaded pivot field
   */
  private renderPivotFieldLoaded(xmlStream: XmlSink, field: ParsedPivotField): void {
    const attrs: Record<string, string> = {};

    // Only add attributes that were present in the original
    if (field.axis) {
      attrs.axis = field.axis;
    }
    if (field.dataField) {
      attrs.dataField = "1";
    }
    if (field.numFmtId !== undefined) {
      attrs.numFmtId = String(field.numFmtId);
    }
    if (field.sortType) {
      attrs.sortType = field.sortType;
    }
    // OOXML defaults: compact=true, outline=true, defaultSubtotal=true when absent.
    // Only write the attribute when false (non-default) to preserve round-trip fidelity.
    if (field.compact === false) {
      attrs.compact = "0";
    }
    if (field.outline === false) {
      attrs.outline = "0";
    }
    // showAll is typically always present — placed before defaultSubtotal to match Excel's ordering
    attrs.showAll = field.showAll ? "1" : "0";
    if (field.defaultSubtotal === false) {
      attrs.defaultSubtotal = "0";
    }
    if (field.subtotalTop === false) {
      attrs.subtotalTop = "0";
    }
    if (field.insertBlankRow === true) {
      attrs.insertBlankRow = "1";
    }
    if (field.multipleItemSelectionAllowed === true) {
      attrs.multipleItemSelectionAllowed = "1";
    }
    // Spread extra unknown attributes for roundtrip preservation
    if (field.extraAttrs) {
      for (const [k, v] of Object.entries(field.extraAttrs)) {
        attrs[k] = v;
      }
    }

    const hasChildren =
      (field.items !== undefined && field.items.length > 0) || field.autoSortScopeXml !== undefined;

    if (hasChildren) {
      xmlStream.openNode("pivotField", attrs);
      if (field.items !== undefined && field.items.length > 0) {
        xmlStream.openNode("items", { count: field.items.length });
        for (const item of field.items) {
          const itemAttrs: Record<string, string | number> = {};
          for (const key of PIVOT_FIELD_ITEM_KEYS) {
            if (item[key] !== undefined) {
              itemAttrs[key] = item[key];
            }
          }
          xmlStream.leafNode("item", itemAttrs);
        }
        xmlStream.closeNode(); // items
      }
      if (field.autoSortScopeXml) {
        xmlStream.writeRaw(field.autoSortScopeXml);
      }
      xmlStream.closeNode(); // pivotField
    } else {
      xmlStream.leafNode("pivotField", attrs);
    }
  }

  // TODO: Consider migrating to map-based child xform delegation (like table-xform.ts)
  // to replace this large manual switch. Currently kept as-is because the manual SAX
  // approach, while verbose, handles all OOXML edge cases correctly.
  parseOpen(node: any): boolean {
    const { name, attributes } = node;

    // Collect raw XML verbatim for roundtrip preservation (5 collectors)
    if (this.extLstCollector.active) {
      this.extLstCollector.feedOpen(name, attributes);
      return true;
    }
    if (this.formatsCollector.active) {
      this.formatsCollector.feedOpen(name, attributes);
      return true;
    }
    if (this.conditionalFormatsCollector.active) {
      this.conditionalFormatsCollector.feedOpen(name, attributes);
      return true;
    }
    if (this.filtersCollector.active) {
      this.filtersCollector.feedOpen(name, attributes);
      return true;
    }
    if (this.unknownCollector.active) {
      this.unknownCollector.feedOpen(name, attributes);
      return true;
    }

    switch (name) {
      case this.tag:
        // pivotTableDefinition root element
        this.reset();
        this.model = {
          name: attributes.name,
          cacheId: parseInt(attributes.cacheId ?? "0", 10),
          uid: attributes["xr:uid"],
          pivotFields: [],
          rowFields: [],
          colFields: [],
          pageFields: [],
          dataFields: [],
          applyNumberFormats: attributes.applyNumberFormats,
          applyBorderFormats: attributes.applyBorderFormats,
          applyFontFormats: attributes.applyFontFormats,
          applyPatternFormats: attributes.applyPatternFormats,
          applyAlignmentFormats: attributes.applyAlignmentFormats,
          applyWidthHeightFormats: attributes.applyWidthHeightFormats,
          dataCaption: attributes.dataCaption,
          updatedVersion: attributes.updatedVersion,
          minRefreshableVersion: attributes.minRefreshableVersion,
          createdVersion: attributes.createdVersion,
          useAutoFormatting: attributes.useAutoFormatting,
          itemPrintTitles: attributes.itemPrintTitles,
          indent: attributes.indent !== undefined ? parseInt(attributes.indent, 10) : undefined,
          compact: attributes.compact !== "0",
          compactData: attributes.compactData !== "0",
          multipleFieldFilters: attributes.multipleFieldFilters,
          outline: attributes.outline === "1",
          outlineData: attributes.outlineData === "1",
          chartFormat:
            attributes.chartFormat !== undefined ? parseInt(attributes.chartFormat, 10) : undefined,
          colGrandTotals: attributes.colGrandTotals,
          rowGrandTotals: attributes.rowGrandTotals,
          showError: attributes.showError,
          errorCaption: attributes.errorCaption,
          showMissing: attributes.showMissing,
          missingCaption: attributes.missingCaption,
          grandTotalCaption: attributes.grandTotalCaption,
          rowItems: [],
          colItems: [],
          chartFormats: [],
          isLoaded: true
        };
        break;

      case "location":
        if (this.model) {
          this.model.location = {
            ref: attributes.ref,
            firstHeaderRow:
              attributes.firstHeaderRow !== undefined
                ? parseInt(attributes.firstHeaderRow, 10)
                : undefined,
            firstDataRow:
              attributes.firstDataRow !== undefined
                ? parseInt(attributes.firstDataRow, 10)
                : undefined,
            firstDataCol:
              attributes.firstDataCol !== undefined
                ? parseInt(attributes.firstDataCol, 10)
                : undefined,
            rowPageCount:
              attributes.rowPageCount !== undefined
                ? parseInt(attributes.rowPageCount, 10)
                : undefined,
            colPageCount:
              attributes.colPageCount !== undefined
                ? parseInt(attributes.colPageCount, 10)
                : undefined
          };
        }
        break;

      case "pivotFields":
        this.state.currentSection = "pivotFields";
        break;

      case "pivotField":
        if (this.state.currentSection === "pivotFields") {
          // Collect unknown attributes into extraAttrs bag for roundtrip preservation
          const extraAttrs: Record<string, string> = {};
          for (const [k, v] of Object.entries(attributes)) {
            if (!KNOWN_PIVOT_FIELD_KEYS.has(k)) {
              extraAttrs[k] = String(v);
            }
          }
          this.currentPivotField = {
            axis: VALID_PIVOT_AXES.has(attributes.axis)
              ? (attributes.axis as ParsedPivotField["axis"])
              : undefined,
            dataField: attributes.dataField === "1",
            items: [],
            compact: attributes.compact !== "0",
            outline: attributes.outline !== "0",
            showAll: attributes.showAll !== "0",
            defaultSubtotal: attributes.defaultSubtotal !== "0",
            numFmtId:
              attributes.numFmtId !== undefined ? parseInt(attributes.numFmtId, 10) : undefined,
            sortType: attributes.sortType,
            subtotalTop:
              attributes.subtotalTop !== undefined ? attributes.subtotalTop === "1" : undefined,
            insertBlankRow: attributes.insertBlankRow === "1" ? true : undefined,
            multipleItemSelectionAllowed:
              attributes.multipleItemSelectionAllowed === "1" ? true : undefined,
            extraAttrs: Object.keys(extraAttrs).length > 0 ? extraAttrs : undefined
          };
        }
        break;

      case "items":
        // No state needed — item parsing is guarded by currentPivotField
        break;

      case "item":
        if (this.currentPivotField) {
          // R8-O1: Parse item attributes using a loop over PIVOT_FIELD_ITEM_KEYS
          const item: PivotFieldItem = {};
          if (attributes.x !== undefined) {
            item.x = parseInt(attributes.x, 10);
          }
          for (const key of PIVOT_FIELD_ITEM_KEYS) {
            if (key !== "x" && attributes[key] !== undefined) {
              (item as Record<string, unknown>)[key] = attributes[key];
            }
          }
          // items is always initialized as [] when currentPivotField is created (see "pivotField" case)
          this.currentPivotField.items!.push(item);
        }
        break;

      case "autoSortScope":
        // Start collecting autoSortScope XML for the current pivotField
        if (this.currentPivotField) {
          this.state.inAutoSortScope = true;
          this.autoSortScopeXmlBuffer = ["<autoSortScope>"];
        }
        break;

      case "rowFields":
        this.state.currentSection = "rowFields";
        break;

      case "colFields":
        this.state.currentSection = "colFields";
        // Track that colFields element was present in original file
        if (this.model) {
          this.model.hasColFields = true;
        }
        break;

      case "dataFields":
        this.state.currentSection = "dataFields";
        break;

      case "pageFields":
        this.state.currentSection = "pageFields";
        break;

      case "pageField":
        if (this.state.currentSection === "pageFields" && this.model) {
          this.model.pageFields.push({
            fld: parseInt(attributes.fld ?? "0", 10),
            item: attributes.item !== undefined ? parseInt(attributes.item, 10) : undefined,
            hier: attributes.hier !== undefined ? parseInt(attributes.hier, 10) : undefined,
            name: attributes.name
          });
        }
        break;

      case "rowItems":
        this.state.currentSection = "rowItems";
        if (this.model) {
          this.model.hasRowItems = true;
        }
        break;

      case "colItems":
        this.state.currentSection = "colItems";
        if (this.model) {
          this.model.hasColItems = true;
        }
        break;

      case "i":
        // Handle row/col item element
        if (this.model) {
          const rowColItem =
            this.state.currentSection === "rowItems" || this.state.currentSection === "colItems"
              ? parseRowColItem(attributes)
              : null;
          if (this.state.currentSection === "rowItems") {
            this.currentRowItem = rowColItem;
          } else if (this.state.currentSection === "colItems") {
            this.currentColItem = rowColItem;
          }
        }
        break;

      case "x":
        // Handle x element inside row/col items or pivotArea
        if (this.state.inPivotArea) {
          // Collect x element for pivotArea XML (re-encode attribute values for XML safety)
          const xAttrs = serializeAttributes(attributes);
          if (this.state.inAutoSortScope) {
            this.autoSortScopeXmlBuffer.push(xAttrs ? `<x ${xAttrs}/>` : "<x/>");
          } else {
            this.pivotAreaXmlBuffer.push(xAttrs ? `<x ${xAttrs}/>` : "<x/>");
          }
        } else if (this.currentRowItem) {
          this.currentRowItem.x.push({ v: parseInt(attributes.v ?? "0", 10) });
        } else if (this.currentColItem) {
          this.currentColItem.x.push({ v: parseInt(attributes.v ?? "0", 10) });
        }
        break;

      case "chartFormats":
        this.state.currentSection = "chartFormats";
        break;

      case "chartFormat":
        if (this.state.currentSection === "chartFormats" && this.model) {
          this.currentChartFormat = {
            chart: parseInt(attributes.chart ?? "0", 10),
            format: parseInt(attributes.format ?? "0", 10),
            series: attributes.series !== undefined ? attributes.series === "1" : undefined
          };
        }
        break;

      case "pivotArea":
        // Start collecting pivotArea XML for chartFormat or autoSortScope
        if (this.currentChartFormat) {
          this.state.inPivotArea = true;
          const attrsStr = serializeAttributes(attributes);
          this.pivotAreaXmlBuffer = [attrsStr ? `<pivotArea ${attrsStr}>` : "<pivotArea>"];
        } else if (this.state.inAutoSortScope) {
          this.state.inPivotArea = true;
          const attrsStr = serializeAttributes(attributes);
          this.autoSortScopeXmlBuffer.push(attrsStr ? `<pivotArea ${attrsStr}>` : "<pivotArea>");
        }
        break;

      case "references":
      case "reference":
        // Collect nested elements in pivotArea
        if (this.state.inPivotArea) {
          const attrsStr = serializeAttributes(attributes);
          if (this.state.inAutoSortScope) {
            this.autoSortScopeXmlBuffer.push(`<${name}${attrsStr ? " " + attrsStr : ""}>`);
          } else {
            this.pivotAreaXmlBuffer.push(`<${name}${attrsStr ? " " + attrsStr : ""}>`);
          }
        }
        break;

      case "field":
        // Handle field element (used in rowFields, colFields)
        if (this.model) {
          const fieldIndex = parseInt(attributes.x ?? "0", 10);
          if (this.state.currentSection === "rowFields") {
            this.model.rowFields.push(fieldIndex);
          } else if (this.state.currentSection === "colFields") {
            this.model.colFields.push(fieldIndex);
          }
        }
        break;

      case "dataField":
        if (this.state.currentSection === "dataFields" && this.model) {
          this.model.dataFields.push({
            name: attributes.name ?? "",
            fld: parseInt(attributes.fld ?? "0", 10),
            baseField:
              attributes.baseField !== undefined ? parseInt(attributes.baseField, 10) : undefined,
            baseItem:
              attributes.baseItem !== undefined ? parseInt(attributes.baseItem, 10) : undefined,
            subtotal: VALID_SUBTOTALS.has(attributes.subtotal)
              ? (attributes.subtotal as PivotTableSubtotal)
              : undefined,
            numFmtId:
              attributes.numFmtId !== undefined ? parseInt(attributes.numFmtId, 10) : undefined
          });
        }
        break;

      case "pivotTableStyleInfo":
        if (this.model) {
          this.model.styleName = attributes.name;
          this.model.styleInfo = {
            name: attributes.name,
            showRowHeaders: attributes.showRowHeaders,
            showColHeaders: attributes.showColHeaders,
            showRowStripes: attributes.showRowStripes,
            showColStripes: attributes.showColStripes,
            showLastColumn: attributes.showLastColumn
          };
        }
        break;

      case "extLst":
        // Start collecting extLst XML for roundtrip preservation
        if (this.model) {
          this.extLstCollector.start(attributes);
        }
        break;

      case "formats":
        // Start collecting formats XML for roundtrip preservation
        if (this.model) {
          this.formatsCollector.start(attributes);
        }
        break;

      case "conditionalFormats":
        // Start collecting conditionalFormats XML for roundtrip preservation
        if (this.model) {
          this.conditionalFormatsCollector.start(attributes);
        }
        break;

      case "filters":
        // Start collecting filters XML for roundtrip preservation
        // <filters> appears between pivotTableStyleInfo and extLst per OOXML schema
        if (this.model) {
          this.filtersCollector.start(attributes);
        }
        break;

      default:
        // Catch-all: collect any unhandled top-level child element as raw XML.
        // This preserves elements like pivotHierarchies, rowHierarchiesUsage,
        // colHierarchiesUsage, etc. that we don't individually model.
        // R8-B1: Only activate at the top level of pivotTableDefinition — NOT inside
        // known sections (pivotFields, rowFields, etc.) or pivotArea/autoSortScope,
        // otherwise the collector would steal subsequent tags from normal parsing.
        if (
          this.model &&
          this.state.currentSection === null &&
          !this.state.inPivotArea &&
          !this.state.inAutoSortScope
        ) {
          this.unknownCollector.startAs(name, attributes);
        }
        break;
    }

    return true;
  }

  parseText(text: string): void {
    // Forward text nodes to whichever raw-XML collector is active (B3 fix)
    if (this.extLstCollector.active) {
      this.extLstCollector.feedText(text);
    } else if (this.formatsCollector.active) {
      this.formatsCollector.feedText(text);
    } else if (this.conditionalFormatsCollector.active) {
      this.conditionalFormatsCollector.feedText(text);
    } else if (this.filtersCollector.active) {
      this.filtersCollector.feedText(text);
    } else if (this.unknownCollector.active) {
      this.unknownCollector.feedText(text);
    }
  }

  /** Feed a close-tag to a collector; if it completes, store the result on the model. */
  private tryCloseCollector(
    collector: RawXmlCollector,
    name: string,
    modelKey: "extLstXml" | "formatsXml" | "conditionalFormatsXml" | "filtersXml"
  ): void {
    if (collector.feedClose(name)) {
      if (this.model) {
        this.model[modelKey] = collector.result;
      }
      collector.reset();
    }
  }

  parseClose(name: string): boolean {
    // Handle raw-XML collectors — close tags
    if (this.extLstCollector.active) {
      this.tryCloseCollector(this.extLstCollector, name, "extLstXml");
      return true;
    }
    if (this.formatsCollector.active) {
      this.tryCloseCollector(this.formatsCollector, name, "formatsXml");
      return true;
    }
    if (this.conditionalFormatsCollector.active) {
      this.tryCloseCollector(this.conditionalFormatsCollector, name, "conditionalFormatsXml");
      return true;
    }
    if (this.filtersCollector.active) {
      this.tryCloseCollector(this.filtersCollector, name, "filtersXml");
      return true;
    }
    if (this.unknownCollector.active) {
      if (this.unknownCollector.feedClose(name)) {
        this.unknownElementsXmlParts.push(this.unknownCollector.result);
        this.unknownCollector.reset();
      }
      return true;
    }

    // Handle pivotArea nested elements - close tags
    if (this.state.inPivotArea) {
      if (name === "pivotArea") {
        if (this.state.inAutoSortScope) {
          this.autoSortScopeXmlBuffer.push("</pivotArea>");
        } else {
          this.pivotAreaXmlBuffer.push("</pivotArea>");
          if (this.currentChartFormat) {
            this.currentChartFormat.pivotAreaXml = this.pivotAreaXmlBuffer.join("");
          }
          this.pivotAreaXmlBuffer = [];
        }
        this.state.inPivotArea = false;
        return true;
      } else if (name === "references" || name === "reference") {
        if (this.state.inAutoSortScope) {
          this.autoSortScopeXmlBuffer.push(`</${name}>`);
        } else {
          this.pivotAreaXmlBuffer.push(`</${name}>`);
        }
        return true;
      }
      // x elements are self-closing, no need to handle close
      return true;
    }

    switch (name) {
      case this.tag:
        // End of pivotTableDefinition — store any collected unknown elements
        if (this.model && this.unknownElementsXmlParts.length > 0) {
          this.model.unknownElementsXml = this.unknownElementsXmlParts.join("");
        }
        return false;

      case "pivotFields":
      case "rowFields":
      case "colFields":
      case "dataFields":
      case "pageFields":
      case "rowItems":
      case "colItems":
      case "chartFormats":
        this.state.currentSection = null;
        break;

      case "pivotField":
        if (this.currentPivotField && this.model) {
          this.model.pivotFields.push(this.currentPivotField);
          this.currentPivotField = null;
        }
        break;

      case "items":
        // No close handling needed — item parsing guarded by currentPivotField
        break;

      case "autoSortScope":
        // Finish collecting autoSortScope XML
        if (this.state.inAutoSortScope && this.currentPivotField) {
          this.autoSortScopeXmlBuffer.push("</autoSortScope>");
          this.currentPivotField.autoSortScopeXml = this.autoSortScopeXmlBuffer.join("");
          this.autoSortScopeXmlBuffer = [];
          this.state.inAutoSortScope = false;
        }
        break;

      case "i":
        // Finish row/col item
        if (this.currentRowItem && this.model) {
          this.model.rowItems?.push(this.currentRowItem);
          this.currentRowItem = null;
        } else if (this.currentColItem && this.model) {
          this.model.colItems?.push(this.currentColItem);
          this.currentColItem = null;
        }
        break;

      case "chartFormat":
        if (this.currentChartFormat && this.model) {
          this.model.chartFormats?.push(this.currentChartFormat);
          this.currentChartFormat = null;
        }
        break;
    }

    return true;
  }

  static readonly PIVOT_TABLE_ATTRIBUTES = {
    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  };

  static readonly EXTLST_XML =
    "<extLst>" +
    '<ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
    '<x14:pivotTableDefinition hideValuesRow="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/>' +
    "</ext>" +
    '<ext uri="{747A6164-185A-40DC-8AA5-F01512510D54}" xmlns:xpdl="http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout">' +
    '<xpdl:pivotTableDefinition16 EnabledSubtotalsDefault="0" SubtotalsOnTopDefault="0"/>' +
    "</ext>" +
    "</extLst>";
}

// Helpers

/** Parse attributes of a row/col `<i>` element into a RowColItem. */
function parseRowColItem(attributes: Record<string, string>): RowColItem {
  return {
    t: attributes.t,
    r: attributes.r !== undefined ? parseInt(attributes.r, 10) : undefined,
    i: attributes.i !== undefined ? parseInt(attributes.i, 10) : undefined,
    x: []
  };
}

/**
 * Render dataField XML elements for all values in the pivot table.
 * Each value field gets its own metric from the `valueMetrics` array.
 */
function renderDataFields(
  xmlStream: XmlSink,
  cacheFields: CacheField[],
  values: number[],
  valueMetrics: PivotTableSubtotal[]
): void {
  xmlStream.openNode("dataFields", { count: values.length });
  for (let i = 0; i < values.length; i++) {
    const valueIndex = values[i]!;
    const metric = valueMetrics[i] ?? "sum";
    const metricName = METRIC_DISPLAY_NAMES[metric];

    const field = cacheFields[valueIndex];
    if (!field) {
      throw new Error(
        `Value field index ${valueIndex} is out of bounds (cacheFields has ${cacheFields.length} entries)`
      );
    }

    const attrs: Record<string, string | number> = {
      name: `${metricName} of ${field.name}`,
      fld: valueIndex,
      baseField: 0,
      baseItem: 0
    };
    // OOXML default is "sum", so omit subtotal attribute for sum
    if (metric !== "sum") {
      attrs.subtotal = metric;
    }
    xmlStream.leafNode("dataField", attrs);
  }
  xmlStream.closeNode();
}

function renderPivotFields(xmlStream: XmlSink, pivotTable: PivotTableRenderModel): void {
  // Pre-compute field type lookup for O(1) access
  const rowSet = new Set(pivotTable.rows);
  const colSet = new Set(pivotTable.columns);
  const valueSet = new Set(pivotTable.values);
  const pageSet = new Set(pivotTable.pages ?? []);

  xmlStream.openNode("pivotFields", { count: pivotTable.cacheFields.length });
  for (let fieldIndex = 0; fieldIndex < pivotTable.cacheFields.length; fieldIndex++) {
    const cacheField = pivotTable.cacheFields[fieldIndex]!;
    const isRow = rowSet.has(fieldIndex);
    const isCol = colSet.has(fieldIndex);
    const isValue = valueSet.has(fieldIndex);
    const isPage = pageSet.has(fieldIndex);
    renderPivotField(xmlStream, isRow, isCol, isValue, isPage, cacheField.sharedItems);
  }
  xmlStream.closeNode();
}

function renderPivotField(
  xmlStream: XmlSink,
  isRow: boolean,
  isCol: boolean,
  isValue: boolean,
  isPage: boolean,
  sharedItems: SharedItemValue[] | null
): void {
  // A field can be both a row/column field AND a value field
  // In this case, it needs both axis attribute AND dataField="1"

  if (isRow || isCol || isPage) {
    if (!sharedItems) {
      throw new Error("sharedItems is required for axis field (row/column/page)");
    }
    const axis = isRow ? "axisRow" : isCol ? "axisCol" : "axisPage";
    const attrs: Record<string, string> = { axis };
    if (isValue) {
      attrs.dataField = "1";
    }
    attrs.compact = "0";
    attrs.outline = "0";
    attrs.showAll = "0";

    xmlStream.openNode("pivotField", attrs);
    // items = one for each shared item + one default item
    xmlStream.openNode("items", { count: sharedItems.length + 1 });
    for (let i = 0; i < sharedItems.length; i++) {
      xmlStream.leafNode("item", { x: i });
    }
    xmlStream.leafNode("item", { t: "default" }); // Required default item for subtotals/grand totals
    xmlStream.closeNode(); // items
    xmlStream.closeNode(); // pivotField
    return;
  }

  // Value fields and non-axis fields should have defaultSubtotal="0"
  const attrs: Record<string, string> = {};
  if (isValue) {
    attrs.dataField = "1";
  }
  attrs.compact = "0";
  attrs.outline = "0";
  attrs.showAll = "0";
  attrs.defaultSubtotal = "0";
  xmlStream.leafNode("pivotField", attrs);
}

export { PivotTableXform, type ParsedPivotTableModel };
