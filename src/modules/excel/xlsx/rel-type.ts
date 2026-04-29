const RelType = {
  OfficeDocument:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  Worksheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
  CalcChain: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
  SharedStrings:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
  Styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  Theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  Hyperlink: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
  Image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  CoreProperties:
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
  ExtenderProperties:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
  Comments: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
  VmlDrawing: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
  Table: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
  PivotCacheDefinition:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition",
  PivotCacheRecords:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords",
  PivotTable: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable",
  FeaturePropertyBag:
    "http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag",
  CtrlProp: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp",
  SheetMetadata:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
  /**
   * Relationship type for the externalLink part referenced from
   * xl/_rels/workbook.xml.rels. Target is an internal path like
   * `externalLinks/externalLink1.xml`.
   */
  ExternalLink: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
  /**
   * Relationship type used *inside* xl/externalLinks/_rels/externalLinkN.xml.rels
   * to point at the actual external workbook. When `TargetMode="External"` and
   * `Target` is a bare relative path (e.g. `"测试.xlsx"`), Office resolves it
   * relative to the current workbook's directory — exactly the behaviour users
   * expect from `=[测试.xlsx]Sheet1!A1`.
   */
  ExternalLinkPath:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
  Chart: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
  ChartEx: "http://schemas.microsoft.com/office/2014/relationships/chartEx",
  ChartStyle: "http://schemas.microsoft.com/office/2011/relationships/chartStyle",
  ChartColors: "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle",
  Drawing: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
  /**
   * `c:userShapes` overlay drawing targeted from a chart rels file. OOXML
   * treats this as a separate relationship type from regular worksheet
   * drawings so the spec's schema for the two parts can diverge.
   */
  ChartUserShapes:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes",
  Chartsheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
  /**
   * Office 365 threaded comments — the conversation tree for modern
   * comments, kept alongside classic VML comments so older viewers still
   * see the fallback text. From MS-XLSX ext spec (2018).
   */
  ThreadedComments: "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment",
  /**
   * `xl/persons/person.xml` — the person directory referenced by
   * threaded-comment `personId` fields. Workbook-level relationship.
   */
  Person: "http://schemas.microsoft.com/office/2017/10/relationships/person",
  /**
   * Slicer part reference from sheets. One per slicer definition.
   */
  Slicer: "http://schemas.microsoft.com/office/2007/relationships/slicer",
  /**
   * Slicer cache reference from workbook — each slicer points at a
   * cache defining its filter state.
   */
  SlicerCache: "http://schemas.microsoft.com/office/2007/relationships/slicerCache",
  /**
   * Excel 2013+ timeline control (date-based slicer).
   */
  Timeline: "http://schemas.microsoft.com/office/2011/relationships/timeline",
  /**
   * Cache behind a timeline control.
   */
  TimelineCache: "http://schemas.microsoft.com/office/2011/relationships/timelineCache"
};

export { RelType };
