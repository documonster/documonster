import {
  OOXML_PATHS,
  commentsPathFromName,
  ctrlPropPath,
  drawingPath,
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

    // Add passthrough content types (charts, etc.)
    if (model.passthroughContentTypes) {
      for (const { partName, contentType } of model.passthroughContentTypes) {
        xmlStream.leafNode("Override", {
          PartName: toContentTypesPartName(partName),
          ContentType: contentType
        });
      }
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
