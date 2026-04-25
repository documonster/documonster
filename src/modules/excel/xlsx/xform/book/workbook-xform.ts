import { colCache } from "@excel/utils/col-cache";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { DefinedNamesXform } from "@excel/xlsx/xform/book/defined-name-xform";
import { ExternalReferenceXform } from "@excel/xlsx/xform/book/external-reference-xform";
import { WorksheetXform } from "@excel/xlsx/xform/book/sheet-xform";
import { WorkbookCalcPropertiesXform } from "@excel/xlsx/xform/book/workbook-calc-properties-xform";
import { WorkbookPivotCacheXform } from "@excel/xlsx/xform/book/workbook-pivot-cache-xform";
import { WorkbookPropertiesXform } from "@excel/xlsx/xform/book/workbook-properties-xform";
import { WorkbookProtectionXform } from "@excel/xlsx/xform/book/workbook-protection-xform";
import { WorkbookViewXform } from "@excel/xlsx/xform/book/workbook-view-xform";
import { ListXform } from "@excel/xlsx/xform/list-xform";
import { StaticXform } from "@excel/xlsx/xform/static-xform";
import { StdDocAttributes } from "@xml/writer";

class WorkbookXform extends BaseXform {
  declare public parser: any;
  declare public map: { [key: string]: any };

  constructor() {
    super();

    this.map = {
      fileVersion: WorkbookXform.STATIC_XFORMS.fileVersion,
      workbookPr: new WorkbookPropertiesXform(),
      workbookProtection: new WorkbookProtectionXform(),
      bookViews: new ListXform({
        tag: "bookViews",
        count: false,
        childXform: new WorkbookViewXform()
      }),
      sheets: new ListXform({ tag: "sheets", count: false, childXform: new WorksheetXform() }),
      definedNames: new ListXform({
        tag: "definedNames",
        count: false,
        childXform: new DefinedNamesXform()
      }),
      calcPr: new WorkbookCalcPropertiesXform(),
      pivotCaches: new ListXform({
        tag: "pivotCaches",
        count: false,
        childXform: new WorkbookPivotCacheXform()
      }),
      externalReferences: new ListXform({
        tag: "externalReferences",
        count: false,
        childXform: new ExternalReferenceXform()
      })
    };
  }

  prepare(model: any): void {
    // Build the sheets list from worksheets + chartsheets.
    // Chartsheets are appended after worksheets.
    model.sheets = [...model.worksheets];
    if (model.chartsheets) {
      for (const cs of model.chartsheets) {
        model.sheets.push(cs);
      }
    }

    // collate all the print areas from all of the sheets and add them to the defined names
    const printAreas: any[] = [];
    let index = 0; // sheets is sparse array - calc index manually
    model.sheets.forEach((sheet: any) => {
      if (sheet.pageSetup && sheet.pageSetup.printArea) {
        sheet.pageSetup.printArea.split("&&").forEach((printArea: string) => {
          const printAreaComponents = printArea.split(":");
          const start = printAreaComponents[0];
          const end = printAreaComponents[1] ?? start;
          const definedName = {
            name: "_xlnm.Print_Area",
            ranges: [`'${sheet.name}'!$${start}:$${end}`],
            localSheetId: index
          };
          printAreas.push(definedName);
        });
      }

      if (
        sheet.pageSetup &&
        (sheet.pageSetup.printTitlesRow || sheet.pageSetup.printTitlesColumn)
      ) {
        const ranges: string[] = [];

        if (sheet.pageSetup.printTitlesColumn) {
          const titlesColumns = sheet.pageSetup.printTitlesColumn.split(":");
          const start = titlesColumns[0];
          const end = titlesColumns[1] ?? start;
          ranges.push(`'${sheet.name}'!$${start}:$${end}`);
        }

        if (sheet.pageSetup.printTitlesRow) {
          const titlesRows = sheet.pageSetup.printTitlesRow.split(":");
          const start = titlesRows[0];
          const end = titlesRows[1] ?? start;
          ranges.push(`'${sheet.name}'!$${start}:$${end}`);
        }

        const definedName = {
          name: "_xlnm.Print_Titles",
          ranges,
          localSheetId: index
        };

        printAreas.push(definedName);
      }
      index++;
    });
    if (printAreas.length) {
      model.definedNames = model.definedNames.concat(printAreas);
    }

    (model.media ?? []).forEach((medium: any, i: number) => {
      // assign name
      medium.name = medium.type + (i + 1);
    });
  }

  render(xmlStream: any, model: any): void {
    xmlStream.openXml(StdDocAttributes);
    xmlStream.openNode("workbook", WorkbookXform.WORKBOOK_ATTRIBUTES);

    this.map.fileVersion.render(xmlStream);
    this.map.workbookPr.render(xmlStream, model.properties);
    this.map.workbookProtection.render(xmlStream, model.protection);
    this.map.bookViews.render(xmlStream, model.views);
    this.map.sheets.render(xmlStream, model.sheets);
    this.map.definedNames.render(xmlStream, model.definedNames);
    this.map.calcPr.render(xmlStream, model.calcProperties);
    // R9-B6: Deduplicate pivot caches by cacheId before rendering.
    // Multiple pivot tables may share the same cache, but workbook.xml should
    // only list each cache once in <pivotCaches>.
    const pivotTables = model.pivotTables ?? [];
    const seenCacheIds = new Set<string>();
    const uniquePivotCaches = pivotTables.filter((pt: any) => {
      if (seenCacheIds.has(pt.cacheId)) {
        return false;
      }
      seenCacheIds.add(pt.cacheId);
      return true;
    });
    this.map.pivotCaches.render(xmlStream, uniquePivotCaches);

    // <externalReferences> order matters: one entry per externalLink in
    // declaration order, each carrying the r:id that resolves (via
    // workbook.xml.rels) to `xl/externalLinks/externalLink{N}.xml`. The
    // index inside this list is exactly the `[N]` used in formula strings.
    const externalLinks = model.externalLinks ?? [];
    if (externalLinks.length > 0) {
      const externalReferenceModels = externalLinks.map((link: any) => ({ rId: link.rId }));
      this.map.externalReferences.render(xmlStream, externalReferenceModels);
    }

    xmlStream.closeNode();
  }

  parseOpen(node: any): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case "workbook":
        return true;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        return true;
    }
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case "workbook":
        this.model = {
          sheets: this.map.sheets.model,
          properties: this.map.workbookPr.model || {},
          protection: this.map.workbookProtection.model,
          views: this.map.bookViews.model,
          calcProperties: this.map.calcPr.model || {}
        };
        if (this.map.definedNames.model) {
          this.model.definedNames = this.map.definedNames.model;
        }
        if (this.map.pivotCaches.model && this.map.pivotCaches.model.length > 0) {
          this.model.pivotCaches = this.map.pivotCaches.model;
        }
        // Attach parsed <externalReferences> as a list of { rId } objects.
        // The reader in xlsx.browser.ts will later join each entry with the
        // matching workbookRels row (to pick up the externalLink Target) and
        // the parsed externalLinkN.xml part (to pick up sheetNames / cache).
        if (this.map.externalReferences.model && this.map.externalReferences.model.length > 0) {
          this.model.externalReferences = this.map.externalReferences.model;
        }

        return false;
      default:
        // not quite sure how we get here!
        return true;
    }
  }

  reconcile(model: any): void {
    const rels = (model.workbookRels ?? []).reduce((map: any, rel: any) => {
      map[rel.Id] = rel;
      return map;
    }, {});

    // reconcile sheet ids, rIds and names
    const worksheets: any[] = [];
    const chartsheetsList: any[] = [];
    let worksheet: any;
    let index = 0;

    (model.sheets ?? []).forEach((sheet: any) => {
      const rel = rels[sheet.rId];
      if (!rel) {
        return;
      }
      const target = `xl/${rel.Target.replace(/^(\s|\/xl\/)+/, "")}`;

      // Check if this is a chartsheet
      const chartsheetMatch = /xl\/chartsheets\/sheet(\d+)\.xml/.exec(target);
      if (chartsheetMatch) {
        const csNo = parseInt(chartsheetMatch[1], 10);
        const chartsheet = model.chartsheets?.[csNo];
        if (chartsheet) {
          chartsheet.name = sheet.name;
          chartsheet.id = sheet.id;
          chartsheet.state = sheet.state;
          chartsheet.rId = sheet.rId;
          chartsheetsList.push(chartsheet);
        }
        return;
      }

      worksheet = model.worksheetHash[target];
      if (worksheet) {
        worksheet.name = sheet.name;
        worksheet.id = sheet.id;
        worksheet.state = sheet.state;
        worksheets[index++] = worksheet;
      }
    });

    // Store reconciled chartsheets on the model
    model.chartsheetsList = chartsheetsList;

    // reconcile print areas
    const definedNames: any[] = [];
    if (model.definedNames) {
      model.definedNames.forEach((definedName: any) => {
        // For print area/titles, use rawText to extract ranges since the xform
        // layer no longer pre-classifies content (two-phase design).
        const effectiveRanges: string[] =
          definedName.ranges?.length > 0
            ? definedName.ranges
            : definedName.rawText
              ? [definedName.rawText]
              : [];

        if (definedName.name === "_xlnm.Print_Area") {
          worksheet = worksheets[definedName.localSheetId];
          if (worksheet && effectiveRanges.length > 0) {
            if (!worksheet.pageSetup) {
              worksheet.pageSetup = {};
            }
            const range: any = colCache.decodeEx(effectiveRanges[0]);
            worksheet.pageSetup.printArea = worksheet.pageSetup.printArea
              ? `${worksheet.pageSetup.printArea}&&${range.dimensions}`
              : range.dimensions;
          }
        } else if (definedName.name === "_xlnm.Print_Titles") {
          worksheet = worksheets[definedName.localSheetId];
          if (worksheet && effectiveRanges.length > 0) {
            if (!worksheet.pageSetup) {
              worksheet.pageSetup = {};
            }

            const rangeString = effectiveRanges.join(",");

            const dollarRegex = /\$/g;

            const rowRangeRegex = /\$\d+:\$\d+/;
            const rowRangeMatches = rangeString.match(rowRangeRegex);

            if (rowRangeMatches && rowRangeMatches.length) {
              const range = rowRangeMatches[0];
              worksheet.pageSetup.printTitlesRow = range.replace(dollarRegex, "");
            }

            const columnRangeRegex = /\$[A-Z]+:\$[A-Z]+/;
            const columnRangeMatches = rangeString.match(columnRangeRegex);

            if (columnRangeMatches && columnRangeMatches.length) {
              const range = columnRangeMatches[0];
              worksheet.pageSetup.printTitlesColumn = range.replace(dollarRegex, "");
            }
          }
        } else {
          definedNames.push(definedName);
        }
      });
    }
    model.definedNames = definedNames;

    // used by sheets to build their image models
    model.media.forEach((media: any, i: number) => {
      media.index = i;
    });
  }

  static WORKBOOK_ATTRIBUTES = {
    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "mc:Ignorable": "x15",
    "xmlns:x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
  };

  static STATIC_XFORMS = {
    fileVersion: new StaticXform({
      tag: "fileVersion",
      $: { appName: "xl", lastEdited: 5, lowestEdited: 5, rupBuild: 9303 }
    })
  };
}

export { WorkbookXform };
