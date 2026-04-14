import { PivotCacheDefinitionXform } from "@excel/xlsx/xform/pivot-table/pivot-cache-definition-xform";
import { describe, it, expect } from "vitest";

describe("PivotCacheDefinitionXform", () => {
  describe("parseOpen - worksheetSource", () => {
    it("should parse name attribute (table reference style)", () => {
      const xform = new PivotCacheDefinitionXform();

      // Simulate parsing pivotCacheDefinition
      xform.parseOpen({ name: "pivotCacheDefinition", attributes: { "r:id": "rId1" } });
      xform.parseOpen({ name: "cacheSource", attributes: { type: "worksheet" } });
      xform.parseOpen({ name: "worksheetSource", attributes: { name: "bookingData" } });

      expect(xform.model?.sourceTableName).toBe("bookingData");
      expect(xform.model?.sourceRef).toBeUndefined();
      expect(xform.model?.sourceSheet).toBeUndefined();
    });

    it("should parse ref and sheet attributes (cell range reference style)", () => {
      const xform = new PivotCacheDefinitionXform();

      xform.parseOpen({ name: "pivotCacheDefinition", attributes: { "r:id": "rId1" } });
      xform.parseOpen({ name: "cacheSource", attributes: { type: "worksheet" } });
      xform.parseOpen({
        name: "worksheetSource",
        attributes: { ref: "A1:C10", sheet: "DataSheet" }
      });

      expect(xform.model?.sourceRef).toBe("A1:C10");
      expect(xform.model?.sourceSheet).toBe("DataSheet");
      expect(xform.model?.sourceTableName).toBeUndefined();
    });
  });

  describe("renderLoaded - worksheetSource", () => {
    it("should render name attribute when sourceTableName is set", () => {
      const xform = new PivotCacheDefinitionXform();
      const xml = xform.toXml({
        isLoaded: true,
        sourceTableName: "bookingData",
        cacheFields: [],
        recordCount: 5
      });

      expect(xml).toContain('<worksheetSource name="bookingData"/>');
      expect(xml).not.toContain("ref=");
      expect(xml).not.toContain("sheet=");
    });

    it("should render ref and sheet attributes when sourceRef/sourceSheet are set", () => {
      const xform = new PivotCacheDefinitionXform();
      const xml = xform.toXml({
        isLoaded: true,
        sourceRef: "A1:D100",
        sourceSheet: "RawData",
        cacheFields: [],
        recordCount: 99
      });

      expect(xml).toContain('ref="A1:D100"');
      expect(xml).toContain('sheet="RawData"');
      expect(xml).not.toContain("name=");
    });

    it("should prefer sourceTableName over sourceRef/sourceSheet", () => {
      const xform = new PivotCacheDefinitionXform();
      const xml = xform.toXml({
        isLoaded: true,
        sourceTableName: "MyTable",
        sourceRef: "A1:D100",
        sourceSheet: "Sheet1",
        cacheFields: [],
        recordCount: 10
      });

      // sourceTableName takes precedence
      expect(xml).toContain('<worksheetSource name="MyTable"/>');
      expect(xml).not.toContain('ref="A1:D100"');
      expect(xml).not.toContain('sheet="Sheet1"');
    });
  });

  describe("roundtrip", () => {
    it("should preserve sourceTableName through parse → render cycle", () => {
      const xform = new PivotCacheDefinitionXform();

      // Parse original XML structure
      xform.parseOpen({ name: "pivotCacheDefinition", attributes: { "r:id": "rId1" } });
      xform.parseOpen({ name: "cacheSource", attributes: { type: "worksheet" } });
      xform.parseOpen({ name: "worksheetSource", attributes: { name: "SalesData" } });
      xform.parseClose("worksheetSource");
      xform.parseClose("cacheSource");
      xform.parseOpen({ name: "cacheFields", attributes: { count: "0" } });
      xform.parseClose("cacheFields");
      xform.parseClose("pivotCacheDefinition");

      // Render back to XML
      const xml = xform.toXml(xform.model!);

      // Verify the name attribute is preserved
      expect(xml).toContain('<worksheetSource name="SalesData"/>');
    });
  });

  // ===========================================================================
  // Round 6 Bug A/D: catch-all unknown element collector
  // ===========================================================================

  describe("R6-BugA: catch-all unknown element collector", () => {
    it("should preserve calculatedItems on roundtrip", () => {
      const xform = new PivotCacheDefinitionXform();

      xform.parseOpen({ name: "pivotCacheDefinition", attributes: { "r:id": "rId1" } });
      xform.parseOpen({ name: "cacheSource", attributes: { type: "worksheet" } });
      xform.parseOpen({ name: "worksheetSource", attributes: { ref: "A1:D10", sheet: "Sheet1" } });
      xform.parseClose("worksheetSource");
      xform.parseClose("cacheSource");
      xform.parseOpen({ name: "cacheFields", attributes: { count: "0" } });
      xform.parseClose("cacheFields");
      // Unknown element: calculatedItems
      xform.parseOpen({ name: "calculatedItems", attributes: { count: "1" } });
      xform.parseOpen({
        name: "calculatedItem",
        attributes: { field: "3", formula: "Amount * Rate" }
      });
      xform.parseClose("calculatedItem");
      xform.parseClose("calculatedItems");
      xform.parseClose("pivotCacheDefinition");

      expect(xform.model!.unknownElementsXml).toBeDefined();
      expect(xform.model!.unknownElementsXml).toContain("<calculatedItems");
      expect(xform.model!.unknownElementsXml).toContain("</calculatedItems>");
      expect(xform.model!.unknownElementsXml).toContain('formula="Amount * Rate"');

      // Render back and verify
      const xml = xform.toXml(xform.model!);
      expect(xml).toContain("<calculatedItems");
      expect(xml).toContain("</calculatedItems>");
    });

    it("should preserve multiple unknown elements", () => {
      const xform = new PivotCacheDefinitionXform();

      xform.parseOpen({ name: "pivotCacheDefinition", attributes: { "r:id": "rId1" } });
      xform.parseOpen({ name: "cacheFields", attributes: { count: "0" } });
      xform.parseClose("cacheFields");
      xform.parseOpen({ name: "cacheHierarchies", attributes: { count: "1" } });
      xform.parseOpen({ name: "cacheHierarchy", attributes: { uniqueName: "[Dim].[Hier]" } });
      xform.parseClose("cacheHierarchy");
      xform.parseClose("cacheHierarchies");
      xform.parseOpen({ name: "dimensions", attributes: { count: "1" } });
      xform.parseOpen({ name: "dimension", attributes: { name: "MyDim" } });
      xform.parseClose("dimension");
      xform.parseClose("dimensions");
      xform.parseClose("pivotCacheDefinition");

      expect(xform.model!.unknownElementsXml).toContain("<cacheHierarchies");
      expect(xform.model!.unknownElementsXml).toContain("<dimensions");
    });

    it("should not collect known elements as unknown", () => {
      const xform = new PivotCacheDefinitionXform();

      xform.parseOpen({ name: "pivotCacheDefinition", attributes: { "r:id": "rId1" } });
      xform.parseOpen({ name: "cacheSource", attributes: { type: "worksheet" } });
      xform.parseOpen({ name: "worksheetSource", attributes: { name: "T1" } });
      xform.parseClose("worksheetSource");
      xform.parseClose("cacheSource");
      xform.parseOpen({ name: "cacheFields", attributes: { count: "0" } });
      xform.parseClose("cacheFields");
      xform.parseOpen({ name: "extLst", attributes: {} });
      xform.parseClose("extLst");
      xform.parseClose("pivotCacheDefinition");

      // No unknown elements — cacheSource, cacheFields, extLst are all known
      expect(xform.model!.unknownElementsXml).toBeUndefined();
    });

    it("should render unknown elements between cacheFields and extLst", () => {
      const xform = new PivotCacheDefinitionXform();
      const xml = xform.toXml({
        isLoaded: true,
        cacheFields: [],
        recordCount: 0,
        sourceRef: "A1:D10",
        sourceSheet: "Sheet1",
        unknownElementsXml:
          '<calculatedItems count="1"><calculatedItem field="3" formula="A*B"/></calculatedItems>',
        extLstXml: "<extLst><ext/></extLst>"
      });

      const cacheFieldsClosePos = xml.indexOf("</cacheFields>");
      const unknownPos = xml.indexOf("<calculatedItems");
      const extLstPos = xml.indexOf("<extLst");

      expect(unknownPos).toBeGreaterThan(cacheFieldsClosePos);
      expect(unknownPos).toBeLessThan(extLstPos);
    });
  });

  // ===========================================================================
  // Round 7: extLstXml roundtrip
  // ===========================================================================

  describe("R7: extLstXml roundtrip", () => {
    it("should parse and preserve extLst XML on roundtrip", () => {
      const xform = new PivotCacheDefinitionXform();

      xform.parseOpen({ name: "pivotCacheDefinition", attributes: { "r:id": "rId1" } });
      xform.parseOpen({ name: "cacheFields", attributes: { count: "0" } });
      xform.parseClose("cacheFields");
      xform.parseOpen({ name: "extLst", attributes: {} });
      xform.parseOpen({
        name: "ext",
        attributes: { uri: "{725AE2AE-9491-48be-B2B4-4EB974FC3084}" }
      });
      xform.parseOpen({
        name: "x14:pivotCacheDefinition",
        attributes: { pivotCacheId: "123456" }
      });
      xform.parseClose("x14:pivotCacheDefinition");
      xform.parseClose("ext");
      xform.parseClose("extLst");
      xform.parseClose("pivotCacheDefinition");

      const model = xform.model!;
      expect(model.extLstXml).toBeDefined();
      expect(model.extLstXml).toContain("<extLst");
      expect(model.extLstXml).toContain("</extLst>");
      expect(model.extLstXml).toContain('pivotCacheId="123456"');

      // Render back and verify
      const xml = xform.toXml(model);
      expect(xml).toContain("<extLst");
      expect(xml).toContain("</extLst>");
      expect(xml).toContain('pivotCacheId="123456"');
    });

    it("should not emit extLst when absent in loaded model", () => {
      const xform = new PivotCacheDefinitionXform();
      const xml = xform.toXml({
        isLoaded: true,
        cacheFields: [],
        recordCount: 0,
        sourceRef: "A1:D10",
        sourceSheet: "Sheet1"
      });

      expect(xml).not.toContain("<extLst");
    });

    it("should render extLst after unknownElements", () => {
      const xform = new PivotCacheDefinitionXform();
      const xml = xform.toXml({
        isLoaded: true,
        cacheFields: [],
        recordCount: 0,
        sourceRef: "A1:D10",
        sourceSheet: "Sheet1",
        unknownElementsXml: '<dimensions count="1"><dimension name="MyDim"/></dimensions>',
        extLstXml: "<extLst><ext/></extLst>"
      });

      const unknownPos = xml.indexOf("<dimensions");
      const extLstPos = xml.indexOf("<extLst");

      expect(unknownPos).toBeGreaterThan(-1);
      expect(extLstPos).toBeGreaterThan(-1);
      expect(unknownPos).toBeLessThan(extLstPos);
    });
  });

  // ===========================================================================
  // Round 7: unknownElementsXml via renderLoaded
  // ===========================================================================

  describe("R7: unknownElementsXml render roundtrip", () => {
    it("should emit unknownElementsXml in rendered output", () => {
      const xform = new PivotCacheDefinitionXform();
      const xml = xform.toXml({
        isLoaded: true,
        cacheFields: [],
        recordCount: 0,
        sourceRef: "A1:D10",
        sourceSheet: "Sheet1",
        unknownElementsXml:
          '<cacheHierarchies count="1"><cacheHierarchy uniqueName="[Dim]"/></cacheHierarchies>'
      });

      expect(xml).toContain("<cacheHierarchies");
      expect(xml).toContain('uniqueName="[Dim]"');
      expect(xml).toContain("</cacheHierarchies>");
    });
  });

  // ===========================================================================
  // R8-T3: Additional attribute coverage
  // ===========================================================================

  describe("R8-T3: root attribute parsing and rendering", () => {
    it("should parse and preserve extraRootAttrs (unknown attributes)", () => {
      const xform = new PivotCacheDefinitionXform();

      xform.parseOpen({
        name: "pivotCacheDefinition",
        attributes: {
          "r:id": "rId1",
          xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          upgradeOnRefresh: "1",
          tupleCache: "1"
        }
      });
      xform.parseOpen({ name: "cacheFields", attributes: { count: "0" } });
      xform.parseClose("cacheFields");
      xform.parseClose("pivotCacheDefinition");

      const model = xform.model!;
      expect(model.extraRootAttrs).toBeDefined();
      expect(model.extraRootAttrs!.upgradeOnRefresh).toBe("1");
      expect(model.extraRootAttrs!.tupleCache).toBe("1");
    });

    it("should render extraRootAttrs in output", () => {
      const xform = new PivotCacheDefinitionXform();
      const xml = xform.toXml({
        isLoaded: true,
        cacheFields: [],
        sourceRef: "A1:B5",
        sourceSheet: "Sheet1",
        extraRootAttrs: { upgradeOnRefresh: "1" }
      });

      expect(xml).toContain('upgradeOnRefresh="1"');
    });

    it("should parse and preserve backgroundQuery attribute", () => {
      const xform = new PivotCacheDefinitionXform();

      xform.parseOpen({
        name: "pivotCacheDefinition",
        attributes: { "r:id": "rId1", backgroundQuery: "1" }
      });
      xform.parseOpen({ name: "cacheFields", attributes: { count: "0" } });
      xform.parseClose("cacheFields");
      xform.parseClose("pivotCacheDefinition");

      expect(xform.model!.backgroundQuery).toBe("1");

      const xml = xform.toXml(xform.model!);
      expect(xml).toContain('backgroundQuery="1"');
    });

    it("should parse and preserve supportSubquery attribute", () => {
      const xform = new PivotCacheDefinitionXform();

      xform.parseOpen({
        name: "pivotCacheDefinition",
        attributes: { "r:id": "rId1", supportSubquery: "1" }
      });
      xform.parseOpen({ name: "cacheFields", attributes: { count: "0" } });
      xform.parseClose("cacheFields");
      xform.parseClose("pivotCacheDefinition");

      expect(xform.model!.supportSubquery).toBe("1");

      const xml = xform.toXml(xform.model!);
      expect(xml).toContain('supportSubquery="1"');
    });

    it("should parse and preserve supportAdvancedDrill attribute", () => {
      const xform = new PivotCacheDefinitionXform();

      xform.parseOpen({
        name: "pivotCacheDefinition",
        attributes: { "r:id": "rId1", supportAdvancedDrill: "1" }
      });
      xform.parseOpen({ name: "cacheFields", attributes: { count: "0" } });
      xform.parseClose("cacheFields");
      xform.parseClose("pivotCacheDefinition");

      expect(xform.model!.supportAdvancedDrill).toBe("1");

      const xml = xform.toXml(xform.model!);
      expect(xml).toContain('supportAdvancedDrill="1"');
    });

    it("should parse and preserve worksheetSourceRId", () => {
      const xform = new PivotCacheDefinitionXform();

      xform.parseOpen({
        name: "pivotCacheDefinition",
        attributes: { "r:id": "rId1" }
      });
      xform.parseOpen({ name: "cacheSource", attributes: { type: "worksheet" } });
      xform.parseOpen({
        name: "worksheetSource",
        attributes: { ref: "A1:D10", sheet: "Data", "r:id": "rId2" }
      });
      xform.parseClose("worksheetSource");
      xform.parseClose("cacheSource");
      xform.parseOpen({ name: "cacheFields", attributes: { count: "0" } });
      xform.parseClose("cacheFields");
      xform.parseClose("pivotCacheDefinition");

      expect(xform.model!.worksheetSourceRId).toBe("rId2");

      const xml = xform.toXml(xform.model!);
      expect(xml).toContain('r:id="rId2"');
    });

    it("should parse and preserve cacheSourceType", () => {
      const xform = new PivotCacheDefinitionXform();

      xform.parseOpen({
        name: "pivotCacheDefinition",
        attributes: { "r:id": "rId1" }
      });
      xform.parseOpen({ name: "cacheSource", attributes: { type: "external" } });
      xform.parseClose("cacheSource");
      xform.parseOpen({ name: "cacheFields", attributes: { count: "0" } });
      xform.parseClose("cacheFields");
      xform.parseClose("pivotCacheDefinition");

      expect(xform.model!.cacheSourceType).toBe("external");
    });
  });

  // ===========================================================================
  // R8-B8/B9/B14: cacheSource non-worksheet source handling
  // ===========================================================================

  describe("R8: non-worksheet cacheSource handling", () => {
    it("should not emit worksheetSource for non-worksheet cacheSourceType", () => {
      const xform = new PivotCacheDefinitionXform();
      const xml = xform.toXml({
        isLoaded: true,
        cacheFields: [],
        cacheSourceType: "consolidation"
      });

      expect(xml).toContain('type="consolidation"');
      expect(xml).not.toContain("<worksheetSource");
    });

    it("should preserve non-worksheet cacheSource children on roundtrip", () => {
      const xform = new PivotCacheDefinitionXform();

      xform.parseOpen({
        name: "pivotCacheDefinition",
        attributes: { "r:id": "rId1" }
      });
      xform.parseOpen({ name: "cacheSource", attributes: { type: "consolidation" } });
      xform.parseOpen({ name: "consolidation", attributes: { autoPage: "1" } });
      xform.parseOpen({ name: "pages", attributes: { count: "1" } });
      xform.parseClose("pages");
      xform.parseClose("consolidation");
      xform.parseClose("cacheSource");
      xform.parseOpen({ name: "cacheFields", attributes: { count: "0" } });
      xform.parseClose("cacheFields");
      xform.parseClose("pivotCacheDefinition");

      const model = xform.model!;
      expect(model.cacheSourceType).toBe("consolidation");
      expect(model.cacheSourceXml).toBeDefined();
      expect(model.cacheSourceXml).toContain('<consolidation autoPage="1">');
      expect(model.cacheSourceXml).toContain("</consolidation>");

      // Render back and verify it's inside <cacheSource>
      const xml = xform.toXml(model);
      expect(xml).toContain('type="consolidation"');
      expect(xml).toContain("<consolidation");
      expect(xml).not.toContain("<worksheetSource");
    });

    it("should not emit empty worksheetSource when all source attributes are undefined", () => {
      const xform = new PivotCacheDefinitionXform();
      const xml = xform.toXml({
        isLoaded: true,
        cacheFields: [],
        cacheSourceType: "worksheet"
        // No sourceRef, sourceSheet, or sourceTableName
      });

      // Should not emit <worksheetSource/> with all undefined attributes
      expect(xml).not.toContain("<worksheetSource");
    });

    it("should emit worksheetSource when sourceRef is present", () => {
      const xform = new PivotCacheDefinitionXform();
      const xml = xform.toXml({
        isLoaded: true,
        cacheFields: [],
        sourceRef: "A1:D10",
        sourceSheet: "Data"
      });

      expect(xml).toContain('ref="A1:D10"');
      expect(xml).toContain('sheet="Data"');
    });
  });
});
