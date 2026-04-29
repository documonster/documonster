import {
  OOXML_PATHS,
  chartsheetPath,
  commentsPathFromName,
  ctrlPropPath,
  drawingPath,
  externalLinkPath,
  chartPath,
  chartUserShapesPath,
  chartExPath,
  chartStylePath,
  chartColorsPath,
  chartExStylePath,
  chartExColorsPath,
  pivotCacheDefinitionPath,
  pivotCacheRecordsPath,
  pivotTablePath,
  tablePath,
  toContentTypesPartName,
  worksheetPath
} from "@excel/utils/ooxml-paths";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { StdDocAttributes } from "@xml/writer";

// used for rendering the [Content_Types].xml file
// not used for parsing
class ContentTypesXform extends BaseXform {
  render(xmlStream: any, model: any): void {
    xmlStream.openXml(StdDocAttributes);

    xmlStream.openNode("Types", ContentTypesXform.PROPERTY_ATTRIBUTES);

    const mediaHash: { [key: string]: boolean } = {};
    (model.media ?? []).forEach((medium: any) => {
      if (medium.type === "image") {
        const imageType = medium.extension;
        if (!mediaHash[imageType]) {
          mediaHash[imageType] = true;
          xmlStream.leafNode("Default", {
            Extension: imageType,
            ContentType: `image/${imageType}`
          });
        }
      }
    });

    xmlStream.leafNode("Default", {
      Extension: "rels",
      ContentType: "application/vnd.openxmlformats-package.relationships+xml"
    });
    xmlStream.leafNode("Default", { Extension: "xml", ContentType: "application/xml" });

    xmlStream.leafNode("Override", {
      PartName: toContentTypesPartName(OOXML_PATHS.xlWorkbook),
      ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    });

    model.worksheets.forEach((worksheet: any) => {
      xmlStream.leafNode("Override", {
        PartName: toContentTypesPartName(worksheetPath(worksheet.fileIndex)),
        ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
      });
    });

    if (model.chartsheets) {
      model.chartsheets.forEach((cs: any) => {
        xmlStream.leafNode("Override", {
          PartName: toContentTypesPartName(chartsheetPath(cs.sheetNo)),
          ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"
        });
      });
    }

    if ((model.pivotTables ?? []).length) {
      // R9-B6: Deduplicate cache content types by cacheId. When multiple pivot tables
      // share the same cache, the cache definition/records files are written only once.
      const writtenCacheIds = new Set<string>();

      // Add content types for each pivot table
      (model.pivotTables ?? []).forEach((pivotTable: any) => {
        const n = pivotTable.tableNumber;
        const cacheId: string = pivotTable.cacheId;

        if (!writtenCacheIds.has(cacheId)) {
          writtenCacheIds.add(cacheId);

          xmlStream.leafNode("Override", {
            PartName: toContentTypesPartName(pivotCacheDefinitionPath(n)),
            ContentType:
              "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"
          });
          // R9-B5: Only register cacheRecords content type when the file actually exists.
          // Loaded pivot tables may lack cacheRecords (e.g. OLAP sources).
          const hasCacheRecords = pivotTable.isLoaded ? !!pivotTable.cacheRecords : true;
          if (hasCacheRecords) {
            xmlStream.leafNode("Override", {
              PartName: toContentTypesPartName(pivotCacheRecordsPath(n)),
              ContentType:
                "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"
            });
          }
        }

        // Each pivot table always has its own file
        xmlStream.leafNode("Override", {
          PartName: toContentTypesPartName(pivotTablePath(n)),
          ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"
        });
      });
    }

    xmlStream.leafNode("Override", {
      PartName: toContentTypesPartName(OOXML_PATHS.xlTheme1),
      ContentType: "application/vnd.openxmlformats-officedocument.theme+xml"
    });
    xmlStream.leafNode("Override", {
      PartName: toContentTypesPartName(OOXML_PATHS.xlStyles),
      ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
    });

    // Each externalLink part needs its own Override. Omitted entries make
    // Excel fail to load the external reference (and, in some builds,
    // trigger a "the workbook is damaged" dialog).
    if (model.externalLinks && model.externalLinks.length > 0) {
      for (const link of model.externalLinks) {
        xmlStream.leafNode("Override", {
          PartName: toContentTypesPartName(externalLinkPath(link.index)),
          ContentType:
            "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"
        });
      }
    }

    // Add FeaturePropertyBag if checkboxes are used
    if (model.hasCheckboxes) {
      xmlStream.leafNode("Override", {
        PartName: toContentTypesPartName(OOXML_PATHS.xlFeaturePropertyBag),
        ContentType: "application/vnd.ms-excel.featurepropertybag+xml"
      });
    }

    // Add metadata part for dynamic array formulas
    if (model.hasDynamicArrayFormulas) {
      xmlStream.leafNode("Override", {
        PartName: toContentTypesPartName(OOXML_PATHS.xlMetadata),
        ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
      });
    }

    const hasSharedStrings = model.sharedStrings && model.sharedStrings.count;
    if (hasSharedStrings) {
      xmlStream.leafNode("Override", {
        PartName: toContentTypesPartName(OOXML_PATHS.xlSharedStrings),
        ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
      });
    }

    if (model.tables) {
      model.tables.forEach((table: any) => {
        xmlStream.leafNode("Override", {
          PartName: toContentTypesPartName(tablePath(table.target)),
          ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
        });
      });
    }

    if (model.drawings) {
      model.drawings.forEach((drawing: any) => {
        xmlStream.leafNode("Override", {
          PartName: toContentTypesPartName(drawingPath(drawing.name)),
          ContentType: "application/vnd.openxmlformats-officedocument.drawing+xml"
        });
      });
    }

    if (model.chartEntries) {
      for (const [n, entry] of Object.entries(model.chartEntries) as Array<
        [string, { userShapesXml?: Uint8Array }]
      >) {
        xmlStream.leafNode("Override", {
          PartName: toContentTypesPartName(chartPath(n)),
          ContentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
        });
        if (entry?.userShapesXml) {
          // `c:userShapes` overlay drawings use the same DrawingML
          // content type as worksheet drawings — Excel distinguishes
          // them via the relationship type, not the content-type.
          xmlStream.leafNode("Override", {
            PartName: toContentTypesPartName(chartUserShapesPath(n)),
            ContentType: "application/vnd.openxmlformats-officedocument.drawing+xml"
          });
        }
      }
    }

    if (model.chartExEntries) {
      for (const n of Object.keys(model.chartExEntries)) {
        xmlStream.leafNode("Override", {
          PartName: toContentTypesPartName(chartExPath(n)),
          ContentType: "application/vnd.ms-office.chartEx+xml"
        });
      }
    }

    if (model.chartExStructuredEntries) {
      for (const n of Object.keys(model.chartExStructuredEntries)) {
        // Skip if already emitted for raw bytes with same number
        if (model.chartExEntries?.[n]) {
          continue;
        }
        xmlStream.leafNode("Override", {
          PartName: toContentTypesPartName(chartExPath(n)),
          ContentType: "application/vnd.ms-office.chartEx+xml"
        });
      }
    }

    if (model.chartStyles) {
      for (const n of Object.keys(model.chartStyles)) {
        xmlStream.leafNode("Override", {
          PartName: toContentTypesPartName(chartStylePath(n)),
          ContentType: "application/vnd.ms-office.chartstyle+xml"
        });
      }
    }
    if (model.chartColors) {
      for (const n of Object.keys(model.chartColors)) {
        xmlStream.leafNode("Override", {
          PartName: toContentTypesPartName(chartColorsPath(n)),
          ContentType: "application/vnd.ms-office.chartcolorstyle+xml"
        });
      }
    }
    if (model.chartExStyles) {
      for (const n of Object.keys(model.chartExStyles)) {
        xmlStream.leafNode("Override", {
          PartName: toContentTypesPartName(chartExStylePath(n)),
          ContentType: "application/vnd.ms-office.chartstyle+xml"
        });
      }
    }
    if (model.chartExColors) {
      for (const n of Object.keys(model.chartExColors)) {
        xmlStream.leafNode("Override", {
          PartName: toContentTypesPartName(chartExColorsPath(n)),
          ContentType: "application/vnd.ms-office.chartcolorstyle+xml"
        });
      }
    }

    // VML extension is needed for comments, form controls, or header watermarks
    const hasComments = model.commentRefs && model.commentRefs.length > 0;
    const hasFormControls = model.formControlRefs && model.formControlRefs.length > 0;
    const hasHeaderWatermark = model.hasHeaderWatermark === true;
    if (hasComments || hasFormControls || hasHeaderWatermark) {
      xmlStream.leafNode("Default", {
        Extension: "vml",
        ContentType: "application/vnd.openxmlformats-officedocument.vmlDrawing"
      });
    }

    if (hasComments) {
      model.commentRefs.forEach(({ commentName }: { commentName: string }) => {
        xmlStream.leafNode("Override", {
          PartName: toContentTypesPartName(commentsPathFromName(commentName)),
          ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
        });
      });
    }

    if (hasFormControls) {
      for (const ctrlPropId of model.formControlRefs) {
        xmlStream.leafNode("Override", {
          PartName: toContentTypesPartName(ctrlPropPath(ctrlPropId)),
          ContentType: "application/vnd.ms-excel.controlproperties+xml"
        });
      }
    }

    // Office 365 threaded comments: each worksheet that has them gets
    // its own override, and the workbook-level person directory gets
    // one more. Writers populate `model.threadedCommentSheetIds` with
    // the worksheet ids that need parts; absent entry means no
    // threaded comments on that sheet.
    const threadedCommentSheetIds: Array<number | string> = model.threadedCommentSheetIds ?? [];
    for (const sheetId of threadedCommentSheetIds) {
      xmlStream.leafNode("Override", {
        PartName: toContentTypesPartName(`xl/threadedComments/threadedComment${sheetId}.xml`),
        ContentType: "application/vnd.ms-excel.threadedcomments+xml"
      });
    }
    if (model.hasPersons) {
      xmlStream.leafNode("Override", {
        PartName: toContentTypesPartName("xl/persons/person.xml"),
        ContentType: "application/vnd.ms-excel.person+xml"
      });
    }

    // Raw-passthrough parts for slicers / timelines. The path lists
    // were populated in `prepareModel` from the workbook's captured
    // raw parts. Content type URIs come from MS-XLSX §2.1.32 and the
    // 2011 timeline extension addendum.
    for (const path of model.slicerPartPaths ?? []) {
      xmlStream.leafNode("Override", {
        PartName: toContentTypesPartName(path),
        ContentType: "application/vnd.ms-excel.slicer+xml"
      });
    }
    for (const path of model.slicerCachePartPaths ?? []) {
      xmlStream.leafNode("Override", {
        PartName: toContentTypesPartName(path),
        ContentType: "application/vnd.ms-excel.slicerCache+xml"
      });
    }
    for (const path of model.timelinePartPaths ?? []) {
      xmlStream.leafNode("Override", {
        PartName: toContentTypesPartName(path),
        ContentType: "application/vnd.ms-excel.timeline+xml"
      });
    }
    for (const path of model.timelineCachePartPaths ?? []) {
      xmlStream.leafNode("Override", {
        PartName: toContentTypesPartName(path),
        ContentType: "application/vnd.ms-excel.timelineCache+xml"
      });
    }

    xmlStream.leafNode("Override", {
      PartName: toContentTypesPartName(OOXML_PATHS.docPropsCore),
      ContentType: "application/vnd.openxmlformats-package.core-properties+xml"
    });
    xmlStream.leafNode("Override", {
      PartName: toContentTypesPartName(OOXML_PATHS.docPropsApp),
      ContentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml"
    });

    xmlStream.closeNode();
  }

  parseOpen(): boolean {
    return false;
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }

  static PROPERTY_ATTRIBUTES = {
    xmlns: "http://schemas.openxmlformats.org/package/2006/content-types"
  };
}

export { ContentTypesXform };
