import type { PivotTableSource } from "@excel/core/pivot-table";
import type {
  CacheField as CacheFieldType,
  ParsedCacheDefinition
} from "@excel/core/pivot-table-types";
import { rangeShortRange } from "@excel/core/range";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { renderCacheField } from "@excel/xlsx/xform/pivot-table/cache-field";
import { CacheFieldXform } from "@excel/xlsx/xform/pivot-table/cache-field-xform";
import { RawXmlCollector } from "@excel/xlsx/xform/pivot-table/raw-xml-collector";
import type { XmlSink } from "@xml/types";
import { StdDocAttributes } from "@xml/writer";

/** Attribute keys on <pivotCacheDefinition> that are individually parsed (not collected into extraRootAttrs). */
const KNOWN_CACHE_DEF_ROOT_KEYS = new Set([
  "xmlns",
  "xmlns:r",
  "r:id",
  "refreshOnLoad",
  "createdVersion",
  "refreshedVersion",
  "minRefreshableVersion",
  "recordCount",
  "backgroundQuery",
  "supportSubquery",
  "supportAdvancedDrill"
]);

/**
 * Model for generating pivot cache definition (with live source)
 */
interface CacheDefinitionModel {
  source: PivotTableSource;
  cacheFields: CacheFieldType[];
}

class PivotCacheDefinitionXform extends BaseXform<ParsedCacheDefinition | null> {
  // Parser state
  private currentCacheField: CacheFieldXform | null = null;
  private inCacheFields = false;
  private inCacheSource = false;
  // Raw XML collectors
  private extLstCollector = new RawXmlCollector("extLst");
  private unknownCollector = new RawXmlCollector("");
  private unknownElementsXmlParts: string[] = [];
  // R8-B9: Collector for non-worksheetSource children inside <cacheSource>
  private cacheSourceChildCollector = new RawXmlCollector("");
  private cacheSourceXmlParts: string[] = [];

  constructor() {
    super();
    this.model = null;
  }

  get tag(): string {
    // http://www.datypic.com/sc/ooxml/e-ssml_pivotCacheDefinition.html
    return "pivotCacheDefinition";
  }

  reset(): void {
    this.model = null;
    this.currentCacheField = null;
    this.inCacheFields = false;
    this.inCacheSource = false;
    this.extLstCollector.reset();
    this.unknownCollector.reset();
    this.unknownElementsXmlParts = [];
    this.cacheSourceChildCollector.reset();
    this.cacheSourceXmlParts = [];
  }

  /**
   * Render pivot cache definition XML.
   * Supports both newly created models (with PivotTableSource) and loaded models.
   */
  render(xmlStream: XmlSink, model: CacheDefinitionModel | ParsedCacheDefinition): void {
    // Check if this is a loaded model (has isLoaded flag or no source property)
    const isLoaded = ("isLoaded" in model && model.isLoaded) || !("source" in model);

    if (isLoaded) {
      this.renderLoaded(xmlStream, model as ParsedCacheDefinition);
    } else {
      this.renderNew(xmlStream, model as CacheDefinitionModel);
    }
  }

  /**
   * Render newly created pivot cache definition
   */
  private renderNew(xmlStream: XmlSink, model: CacheDefinitionModel): void {
    const { source, cacheFields } = model;

    // R8-O2: Use Array.isArray for type safety — getSheetValues() returns a sparse array of row arrays
    const recordCount = source.getSheetValues().slice(2).filter(Array.isArray).length;

    xmlStream.openXml(StdDocAttributes);
    xmlStream.openNode(this.tag, {
      ...PivotCacheDefinitionXform.PIVOT_CACHE_DEFINITION_ATTRIBUTES,
      "r:id": "rId1",
      refreshOnLoad: "1", // important for our implementation to work
      createdVersion: "8",
      refreshedVersion: "8",
      minRefreshableVersion: "3",
      recordCount
    });

    xmlStream.openNode("cacheSource", { type: "worksheet" });
    // When source is a named Table, reference it by name (dynamic range tracking).
    // Otherwise, reference by sheet + cell range (static).
    const worksheetSourceAttrs: Record<string, string | undefined> = source.tableName
      ? { name: source.tableName }
      : { ref: rangeShortRange(source.dimensions), sheet: source.name };
    xmlStream.leafNode("worksheetSource", worksheetSourceAttrs);
    xmlStream.closeNode();

    xmlStream.openNode("cacheFields", { count: cacheFields.length });
    // Note: keeping this pretty-printed for now to ease debugging.
    xmlStream.writeRaw(PivotCacheDefinitionXform.renderCacheFieldsXml(cacheFields));
    xmlStream.closeNode();

    xmlStream.closeNode();
  }

  /**
   * Render loaded pivot cache definition (preserving original structure)
   */
  private renderLoaded(xmlStream: XmlSink, model: ParsedCacheDefinition): void {
    const { cacheFields, sourceRef, sourceSheet, sourceTableName, recordCount } = model;

    xmlStream.openXml(StdDocAttributes);
    const rootAttrs: Record<string, string | number | undefined> = {
      ...PivotCacheDefinitionXform.PIVOT_CACHE_DEFINITION_ATTRIBUTES,
      "r:id": model.rId ?? "rId1"
    };
    // Only emit refreshOnLoad when it was present in the original file.
    // Forcing "1" causes Excel to recalculate the cache on every open.
    // Placed before createdVersion to match Excel's attribute ordering.
    if (model.refreshOnLoad) {
      rootAttrs.refreshOnLoad = model.refreshOnLoad;
    }
    rootAttrs.createdVersion = model.createdVersion ?? "8";
    rootAttrs.refreshedVersion = model.refreshedVersion ?? "8";
    rootAttrs.minRefreshableVersion = model.minRefreshableVersion ?? "3";
    // Only emit recordCount when it was present in the original file
    if (recordCount !== undefined) {
      rootAttrs.recordCount = recordCount;
    }
    // BUG-26: Preserve additional root attributes
    if (model.backgroundQuery) {
      rootAttrs.backgroundQuery = model.backgroundQuery;
    }
    if (model.supportSubquery) {
      rootAttrs.supportSubquery = model.supportSubquery;
    }
    if (model.supportAdvancedDrill) {
      rootAttrs.supportAdvancedDrill = model.supportAdvancedDrill;
    }
    // Extra unknown root attributes (roundtrip bag)
    if (model.extraRootAttrs) {
      for (const [k, v] of Object.entries(model.extraRootAttrs)) {
        rootAttrs[k] = v;
      }
    }
    xmlStream.openNode(this.tag, rootAttrs);

    xmlStream.openNode("cacheSource", { type: model.cacheSourceType ?? "worksheet" });
    // R8-B8/B14: Only emit <worksheetSource> when the source type is worksheet (or default)
    // and at least one source attribute is defined. Non-worksheet types (e.g. "consolidation",
    // "external") don't use <worksheetSource>.
    const isWorksheetSource = !model.cacheSourceType || model.cacheSourceType === "worksheet";
    if (isWorksheetSource) {
      // worksheetSource supports two reference styles:
      // 1. name: references a named Table
      // 2. ref + sheet: references a cell range on a worksheet
      const worksheetSourceAttrs: Record<string, string | undefined> = sourceTableName
        ? { name: sourceTableName }
        : { ref: sourceRef, sheet: sourceSheet };
      // BUG-28: Preserve r:id for external connections
      if (model.worksheetSourceRId) {
        worksheetSourceAttrs["r:id"] = model.worksheetSourceRId;
      }
      // R8-B14: Only emit <worksheetSource> if we have at least one meaningful attribute
      const hasAnyAttr = Object.values(worksheetSourceAttrs).some(v => v !== undefined);
      if (hasAnyAttr) {
        xmlStream.leafNode("worksheetSource", worksheetSourceAttrs);
      }
    }
    // R8-B9: Emit preserved non-worksheet cacheSource children (e.g. <consolidation>)
    if (model.cacheSourceXml) {
      xmlStream.writeRaw(model.cacheSourceXml);
    }
    xmlStream.closeNode();

    xmlStream.openNode("cacheFields", { count: cacheFields.length });
    xmlStream.writeRaw(PivotCacheDefinitionXform.renderCacheFieldsXml(cacheFields));
    xmlStream.closeNode();

    // R6-BugA: Preserved unknown child elements for roundtrip
    if (model.unknownElementsXml) {
      xmlStream.writeRaw(model.unknownElementsXml);
    }

    // BUG-29: Preserve extLst from original file
    if (model.extLstXml) {
      xmlStream.writeRaw(model.extLstXml);
    }

    xmlStream.closeNode();
  }

  parseOpen(node: any): boolean {
    const { name, attributes } = node;

    // Collect extLst XML verbatim for roundtrip preservation
    if (this.extLstCollector.active) {
      this.extLstCollector.feedOpen(name, attributes);
      return true;
    }

    // Collect unknown child element XML verbatim for roundtrip preservation
    if (this.unknownCollector.active) {
      this.unknownCollector.feedOpen(name, attributes);
      return true;
    }

    // R8-B9: Collect non-worksheetSource children inside <cacheSource>
    if (this.cacheSourceChildCollector.active) {
      this.cacheSourceChildCollector.feedOpen(name, attributes);
      return true;
    }

    // Delegate to current cacheField parser if active
    if (this.currentCacheField) {
      this.currentCacheField.parseOpen(node);
      return true;
    }

    switch (name) {
      case this.tag: {
        // pivotCacheDefinition root element
        this.reset();
        // Collect known attributes individually and put the rest in extraRootAttrs
        const extraRootAttrs: Record<string, string> = {};
        for (const [k, v] of Object.entries(attributes)) {
          if (!KNOWN_CACHE_DEF_ROOT_KEYS.has(k)) {
            extraRootAttrs[k] = String(v);
          }
        }
        this.model = {
          cacheFields: [],
          rId: attributes["r:id"],
          refreshOnLoad: attributes.refreshOnLoad,
          createdVersion: attributes.createdVersion,
          refreshedVersion: attributes.refreshedVersion,
          minRefreshableVersion: attributes.minRefreshableVersion,
          recordCount: attributes.recordCount ? parseInt(attributes.recordCount, 10) : undefined,
          backgroundQuery: attributes.backgroundQuery,
          supportSubquery: attributes.supportSubquery,
          supportAdvancedDrill: attributes.supportAdvancedDrill,
          extraRootAttrs: Object.keys(extraRootAttrs).length > 0 ? extraRootAttrs : undefined,
          isLoaded: true
        };
        break;
      }

      case "cacheSource":
        this.inCacheSource = true;
        if (this.model) {
          this.model.cacheSourceType = attributes.type;
        }
        break;

      case "worksheetSource":
        if (this.inCacheSource && this.model) {
          this.model.sourceRef = attributes.ref;
          this.model.sourceSheet = attributes.sheet;
          this.model.sourceTableName = attributes.name;
          // Preserve r:id for external connections (BUG-28)
          if (attributes["r:id"]) {
            this.model.worksheetSourceRId = attributes["r:id"];
          }
        }
        break;

      case "cacheFields":
        this.inCacheFields = true;
        break;

      case "cacheField":
        if (this.inCacheFields) {
          this.currentCacheField = new CacheFieldXform();
          this.currentCacheField.parseOpen(node);
        }
        break;

      case "extLst":
        // Start collecting extLst XML for roundtrip preservation
        if (this.model) {
          this.extLstCollector.start(attributes);
        }
        break;

      default:
        // R8-B9: Non-worksheetSource children inside <cacheSource> (e.g. <consolidation>)
        // must be collected separately so they stay inside <cacheSource> on roundtrip.
        if (this.inCacheSource && this.model) {
          this.cacheSourceChildCollector.startAs(name, attributes);
          break;
        }
        // Catch-all: collect any unhandled child element as raw XML
        // This preserves elements like calculatedItems, cacheHierarchies, kpis,
        // dimensions, measureGroups, maps, etc. that we don't individually model.
        if (this.model) {
          this.unknownCollector.startAs(name, attributes);
        }
        break;
    }

    return true;
  }

  parseText(text: string): void {
    // Forward text to active collectors (B3 fix: text nodes in raw XML)
    if (this.extLstCollector.active) {
      this.extLstCollector.feedText(text);
      return;
    }
    if (this.unknownCollector.active) {
      this.unknownCollector.feedText(text);
      return;
    }
    // R8-B9: Forward text to cacheSource child collector
    if (this.cacheSourceChildCollector.active) {
      this.cacheSourceChildCollector.feedText(text);
      return;
    }
    if (this.currentCacheField) {
      this.currentCacheField.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    // Handle extLst collection — close tags
    if (this.extLstCollector.active) {
      if (this.extLstCollector.feedClose(name)) {
        if (this.model) {
          this.model.extLstXml = this.extLstCollector.result;
        }
        this.extLstCollector.reset();
      }
      return true;
    }

    // Handle unknown element collection — close tags
    if (this.unknownCollector.active) {
      if (this.unknownCollector.feedClose(name)) {
        this.unknownElementsXmlParts.push(this.unknownCollector.result);
        this.unknownCollector.reset();
      }
      return true;
    }

    // R8-B9: Handle cacheSource child collection — close tags
    if (this.cacheSourceChildCollector.active) {
      if (this.cacheSourceChildCollector.feedClose(name)) {
        this.cacheSourceXmlParts.push(this.cacheSourceChildCollector.result);
        this.cacheSourceChildCollector.reset();
      }
      return true;
    }

    // Delegate to current cacheField parser if active
    if (this.currentCacheField) {
      if (!this.currentCacheField.parseClose(name)) {
        // cacheField parsing complete, add to model
        if (this.model && this.currentCacheField.model) {
          this.model.cacheFields.push(this.currentCacheField.model);
        }
        this.currentCacheField = null;
      }
      return true;
    }

    switch (name) {
      case this.tag:
        // End of pivotCacheDefinition — store any collected unknown elements
        if (this.model && this.unknownElementsXmlParts.length > 0) {
          this.model.unknownElementsXml = this.unknownElementsXmlParts.join("");
        }
        return false;

      case "cacheSource":
        this.inCacheSource = false;
        // R8-B9: Store collected cacheSource children on model
        if (this.model && this.cacheSourceXmlParts.length > 0) {
          this.model.cacheSourceXml = this.cacheSourceXmlParts.join("");
        }
        break;

      case "cacheFields":
        this.inCacheFields = false;
        break;
    }

    return true;
  }

  static renderCacheFieldsXml(cacheFields: CacheFieldType[]): string {
    return "\n    " + cacheFields.map(cf => renderCacheField(cf)).join("\n    ");
  }

  static readonly PIVOT_CACHE_DEFINITION_ATTRIBUTES = {
    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  };
}

export { PivotCacheDefinitionXform };
