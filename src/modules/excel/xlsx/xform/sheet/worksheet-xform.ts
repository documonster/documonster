import { colCache } from "@excel/utils/col-cache";
import { buildDrawingAnchorsAndRels, resolveMediaTarget } from "@excel/utils/drawing-utils";
import {
  chartRelTargetFromDrawing,
  chartExRelTargetFromDrawing,
  commentsRelTargetFromWorksheet,
  ctrlPropRelTargetFromWorksheet,
  drawingRelTargetFromWorksheet,
  pivotTableRelTargetFromWorksheet,
  resolveRelTarget,
  tableRelTargetFromWorksheet,
  vmlDrawingRelTargetFromWorksheet,
  vmlDrawingHFRelTargetFromWorksheet
} from "@excel/utils/ooxml-paths";
import { RelType } from "@excel/xlsx/rel-type";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { ListXform } from "@excel/xlsx/xform/list-xform";
import { AutoFilterXform } from "@excel/xlsx/xform/sheet/auto-filter-xform";
import { ConditionalFormattingsXform } from "@excel/xlsx/xform/sheet/cf/conditional-formattings-xform";
import { ColBreaksXform } from "@excel/xlsx/xform/sheet/col-breaks-xform";
import { ColXform } from "@excel/xlsx/xform/sheet/col-xform";
import { DataValidationsXform } from "@excel/xlsx/xform/sheet/data-validations-xform";
import { DimensionXform } from "@excel/xlsx/xform/sheet/dimension-xform";
import { DrawingXform } from "@excel/xlsx/xform/sheet/drawing-xform";
import { ExtLstXform } from "@excel/xlsx/xform/sheet/ext-lst-xform";
import { HeaderFooterXform } from "@excel/xlsx/xform/sheet/header-footer-xform";
import { isInternalLink, HyperlinkXform } from "@excel/xlsx/xform/sheet/hyperlink-xform";
import { IgnoredErrorsXform } from "@excel/xlsx/xform/sheet/ignored-errors-xform";
import { MergeCellXform } from "@excel/xlsx/xform/sheet/merge-cell-xform";
import { Merges } from "@excel/xlsx/xform/sheet/merges";
import { PageMarginsXform } from "@excel/xlsx/xform/sheet/page-margins-xform";
import { PageSetupXform } from "@excel/xlsx/xform/sheet/page-setup-xform";
import { PictureXform } from "@excel/xlsx/xform/sheet/picture-xform";
import { PrintOptionsXform } from "@excel/xlsx/xform/sheet/print-options-xform";
import { RowBreaksXform } from "@excel/xlsx/xform/sheet/row-breaks-xform";
import { RowXform } from "@excel/xlsx/xform/sheet/row-xform";
import { SheetFormatPropertiesXform } from "@excel/xlsx/xform/sheet/sheet-format-properties-xform";
import { SheetPropertiesXform } from "@excel/xlsx/xform/sheet/sheet-properties-xform";
import { SheetProtectionXform } from "@excel/xlsx/xform/sheet/sheet-protection-xform";
import { SheetViewXform } from "@excel/xlsx/xform/sheet/sheet-view-xform";
import { TablePartXform } from "@excel/xlsx/xform/sheet/table-part-xform";
import { StdDocAttributes } from "@xml/writer";

const mergeRule = (rule, extRule) => {
  Object.keys(extRule).forEach(key => {
    const value = rule[key];
    const extValue = extRule[key];
    if (value === undefined && extValue !== undefined) {
      rule[key] = extValue;
    }
  });
};

const mergeConditionalFormattings = (model, extModel) => {
  // conditional formattings are rendered in worksheet.conditionalFormatting and also in
  // worksheet.extLst.ext.x14:conditionalFormattings
  // some (e.g. dataBar) are even spread across both!
  if (!extModel || !extModel.length) {
    return model;
  }
  if (!model || !model.length) {
    return extModel;
  }

  // index model rules by x14Id
  const cfMap = {};
  const ruleMap = {};
  model.forEach(cf => {
    cfMap[cf.ref] = cf;
    cf.rules.forEach(rule => {
      const { x14Id } = rule;
      if (x14Id) {
        ruleMap[x14Id] = rule;
      }
    });
  });

  extModel.forEach(extCf => {
    extCf.rules.forEach(extRule => {
      const rule = ruleMap[extRule.x14Id];
      if (rule) {
        // merge with matching rule
        mergeRule(rule, extRule);
      } else if (cfMap[extCf.ref]) {
        // reuse existing cf ref
        cfMap[extCf.ref].rules.push(extRule);
      } else {
        // create new cf
        model.push({
          ref: extCf.ref,
          rules: [extRule]
        });
      }
    });
  });

  // need to cope with rules in extModel that don't exist in model
  return model;
};

class WorkSheetXform extends BaseXform {
  declare public map: { [key: string]: any };
  declare private ignoreNodes: string[];
  declare public parser: any;

  static WORKSHEET_ATTRIBUTES = {
    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "xmlns:x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "mc:Ignorable": "x14ac"
  };

  constructor(options?: any) {
    super();

    const { maxRows, maxCols, ignoreNodes } = options || {};

    this.ignoreNodes = ignoreNodes ?? [];

    this.map = {
      sheetPr: new SheetPropertiesXform(),
      dimension: new DimensionXform(),
      sheetViews: new ListXform({
        tag: "sheetViews",
        count: false,
        childXform: new SheetViewXform()
      }),
      sheetFormatPr: new SheetFormatPropertiesXform(),
      cols: new ListXform({ tag: "cols", count: false, childXform: new ColXform() }),
      sheetData: new ListXform({
        tag: "sheetData",
        count: false,
        empty: true,
        childXform: new RowXform({ maxItems: maxCols }),
        maxItems: maxRows
      }),
      autoFilter: new AutoFilterXform(),
      mergeCells: new ListXform({
        tag: "mergeCells",
        count: true,
        childXform: new MergeCellXform()
      }),
      rowBreaks: new RowBreaksXform(),
      colBreaks: new ColBreaksXform(),
      hyperlinks: new ListXform({
        tag: "hyperlinks",
        count: false,
        childXform: new HyperlinkXform()
      }),
      pageMargins: new PageMarginsXform(),
      dataValidations: new DataValidationsXform(),
      pageSetup: new PageSetupXform(),
      headerFooter: new HeaderFooterXform(),
      printOptions: new PrintOptionsXform(),
      picture: new PictureXform(),
      drawing: new DrawingXform(),
      sheetProtection: new SheetProtectionXform(),
      tableParts: new ListXform({
        tag: "tableParts",
        count: true,
        childXform: new TablePartXform()
      }),
      conditionalFormatting: new ConditionalFormattingsXform(),
      extLst: new ExtLstXform(),
      ignoredErrors: new IgnoredErrorsXform()
    };
  }

  prepare(model, options) {
    options.merges = new Merges();
    model.hyperlinks = options.hyperlinks = [];
    model.comments = options.comments = [];

    // Some Excel builds are surprisingly strict when legacy form controls exist.
    // Emitting a default sheetView (workbookViewId=0) matches typical Excel output
    // and avoids relying on optional element handling.
    if (model.formControls && model.formControls.length > 0) {
      if (!model.views || model.views.length === 0) {
        model.views = [{ workbookViewId: 0 }];
      }
    }

    options.formulae = {};
    options.siFormulae = 0;
    this.map.cols.prepare(model.cols, options);
    this.map.sheetData.prepare(model.rows, options);
    this.map.conditionalFormatting.prepare(model.conditionalFormattings, options);

    model.mergeCells = options.merges.mergeCells;

    // prepare relationships
    const rels: any[] = (model.rels = []);

    function nextRid(r) {
      return `rId${r.length + 1}`;
    }

    model.hyperlinks.forEach(hyperlink => {
      // Internal links (e.g. "#Sheet2!A1") use the location attribute only,
      // no relationship is needed in the .rels file.
      if (isInternalLink(hyperlink.target)) {
        return;
      }
      const rId = nextRid(rels);
      hyperlink.rId = rId;
      rels.push({
        Id: rId,
        Type: RelType.Hyperlink,
        Target: hyperlink.target,
        TargetMode: "External"
      });
    });

    // prepare comment relationships
    // Use fileIndex (sequential 1-based) for file naming instead of model.id
    // (the workbook-level sheet ID) because model.id can have gaps when sheets
    // have been deleted, causing a mismatch between the relationship targets
    // and the actual ZIP entry paths written by addWorksheets().
    const { fileIndex } = model;
    if (model.comments.length > 0) {
      const comment = {
        Id: nextRid(rels),
        Type: RelType.Comments,
        Target: commentsRelTargetFromWorksheet(fileIndex)
      };
      rels.push(comment);
      const vmlDrawing = {
        Id: nextRid(rels),
        Type: RelType.VmlDrawing,
        Target: vmlDrawingRelTargetFromWorksheet(fileIndex)
      };
      rels.push(vmlDrawing);

      model.comments.forEach(item => {
        item.refAddress = colCache.decodeAddress(item.ref);
      });

      options.commentRefs.push({
        commentName: `comments${fileIndex}`,
        vmlDrawing: `vmlDrawing${fileIndex}`
      });
    }

    // Office 365 threaded comments — a separate rel pointing at the
    // per-sheet `xl/threadedComments/threadedComment{N}.xml` part.
    // Excel requires the rel to exist even when the sheet already has
    // a classic comments rel (the two are independent — legacy VML
    // carries the fallback text, threaded comments carry the modern
    // conversation tree).
    if (model.threadedComments && model.threadedComments.length > 0) {
      rels.push({
        Id: nextRid(rels),
        Type: RelType.ThreadedComments,
        Target: `../threadedComments/threadedComment${fileIndex}.xml`
      });
    }

    // Handle pre-loaded drawing (from file read) that may contain charts or other non-image content.
    // Chart anchors (with chartNumber from reconcile) are preserved and get fresh rels.
    // Non-chart anchors are discarded — images are rebuilt from model.media below.
    if (model.drawing && model.drawing.anchors) {
      const drawing = model.drawing;
      drawing.rId = nextRid(rels);
      if (!drawing.name) {
        drawing.name = `drawing${++options.drawingsCount}`;
      }

      // Separate chart anchors from non-chart anchors
      const chartAnchors = drawing.anchors.filter((a: any) => a.chartNumber || a.chartExNumber);

      // Reset anchors — chart anchors will be re-added, image anchors rebuilt below
      drawing.anchors = [];
      drawing.rels = [];

      // Re-add chart anchors and build their rels
      for (const anchor of chartAnchors) {
        const chartRId = nextRid(drawing.rels);
        if (anchor.chartExNumber) {
          // Office 2016+ cx chart
          drawing.rels.push({
            Id: chartRId,
            Type: RelType.ChartEx,
            Target: chartExRelTargetFromDrawing(anchor.chartExNumber)
          });
          anchor.graphicFrame = {
            ...anchor.graphicFrame,
            rId: chartRId,
            isChartEx: true,
            index: anchor.graphicFrame?.index,
            name: anchor.graphicFrame?.name ?? `Chart ${anchor.chartExNumber}`
          };
          // ChartEx MUST be serialised inside `<mc:AlternateContent>`
          // — see TwoCellAnchorXform.render for the full rationale.
          // Tag the anchor here so re-saves of files we authored
          // ourselves (which load without the wrapper in the current
          // codebase) still emit the correct structure on the next
          // write. Files that came in already wrapped keep whatever
          // requires value the parser captured.
          if (!anchor.alternateContent) {
            anchor.alternateContent = { requires: "cx1" };
          }
        } else {
          // Traditional c: chart
          drawing.rels.push({
            Id: chartRId,
            Type: RelType.Chart,
            Target: chartRelTargetFromDrawing(anchor.chartNumber)
          });
          anchor.graphicFrame = {
            ...anchor.graphicFrame,
            rId: chartRId,
            index: anchor.graphicFrame?.index,
            name: anchor.graphicFrame?.name ?? `Chart ${anchor.chartNumber}`
          };
        }
        drawing.anchors.push(anchor);
      }

      options.drawings.push(drawing);
      rels.push({
        Id: drawing.rId,
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
        Target: drawingRelTargetFromWorksheet(drawing.name)
      });
    }

    // Handle programmatic charts (from worksheet.addChart API) that are NOT already
    // in the drawing. Charts loaded from XLSX are already in drawing.anchors with
    // chartNumber — only add truly new charts.
    if (model.charts && model.charts.length > 0) {
      // Collect chartNumbers already present in the drawing
      const existingChartNumbers = new Set<number>();
      const existingChartExNumbers = new Set<number>();
      if (model.drawing?.anchors) {
        for (const a of model.drawing.anchors) {
          if (a.chartNumber) {
            existingChartNumbers.add(a.chartNumber);
          }
          if (a.chartExNumber) {
            existingChartExNumbers.add(a.chartExNumber);
          }
        }
      }

      const newCharts = model.charts.filter((c: any) => {
        if (c.chartNumber && !existingChartNumbers.has(c.chartNumber)) {
          return true;
        }
        if (c.chartExNumber && !existingChartExNumbers.has(c.chartExNumber)) {
          return true;
        }
        return false;
      });

      if (newCharts.length > 0) {
        // Ensure a drawing exists
        if (!model.drawing) {
          model.drawing = {
            anchors: [],
            rels: [],
            rId: nextRid(rels),
            name: `drawing${++options.drawingsCount}`
          };
          options.drawings.push(model.drawing);
          rels.push({
            Id: model.drawing.rId,
            Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
            Target: drawingRelTargetFromWorksheet(model.drawing.name)
          });
        }

        const drawing = model.drawing;
        for (const chartAnchor of newCharts) {
          // Chart `ChartAnchorModel.range` stores absolute position
          // (`pos.x/y`) and extent (`ext.cx/cy`) in EMU — the native
          // OOXML unit. The drawing xform pipeline (PosXform,
          // ExtXform) expects pixels at 96 dpi (the image convention
          // — `ext.width/height`) and multiplies by `EMU_PER_PIXEL_AT_96_DPI`
          // on render. Normalise at this boundary: convert the EMU
          // values to pixels so downstream xforms re-convert to EMU
          // without doubling.
          //
          // Without this conversion an absolute-anchored chart wrote
          // `<xdr:pos x="8709660000" y="8709660000"/>` for a 914400
          // EMU position (a 9525× overshoot) and
          // `<xdr:ext cx="NaN" cy="NaN"/>` because `ext` carried
          // `{ cx, cy }` keys but ExtXform looked for `{ width,
          // height }`.
          const emuToPixel = 9525; // EMU_PER_PIXEL_AT_96_DPI
          const normalizedRange: Record<string, unknown> = { ...chartAnchor.range };
          if (chartAnchor.range.pos) {
            normalizedRange.pos = {
              x: chartAnchor.range.pos.x / emuToPixel,
              y: chartAnchor.range.pos.y / emuToPixel
            };
          }
          if (chartAnchor.range.ext && chartAnchor.range.ext.cx !== undefined) {
            normalizedRange.ext = {
              width: chartAnchor.range.ext.cx / emuToPixel,
              height: chartAnchor.range.ext.cy / emuToPixel
            };
          }
          const chartRId = nextRid(drawing.rels);
          if (chartAnchor.chartExNumber) {
            drawing.rels.push({
              Id: chartRId,
              Type: RelType.ChartEx,
              Target: chartExRelTargetFromDrawing(chartAnchor.chartExNumber)
            });
            drawing.anchors.push({
              range: normalizedRange,
              chartExNumber: chartAnchor.chartExNumber,
              alternateContent: { requires: "cx1" },
              graphicFrame: {
                rId: chartRId,
                isChartEx: true,
                name: `Chart ${chartAnchor.chartExNumber}`
              }
            });
          } else {
            drawing.rels.push({
              Id: chartRId,
              Type: RelType.Chart,
              Target: chartRelTargetFromDrawing(chartAnchor.chartNumber)
            });
            drawing.anchors.push({
              range: normalizedRange,
              chartNumber: chartAnchor.chartNumber,
              graphicFrame: {
                rId: chartRId,
                name: `Chart ${chartAnchor.chartNumber}`
              }
            });
          }
        }
      }
    }

    // Process background and image media entries
    const backgroundMedia: any[] = [];
    const imageMedia: any[] = [];
    const watermarkMedia: any[] = [];
    const headerImageMedia: any[] = [];
    model.media.forEach(medium => {
      if (medium.type === "background") {
        backgroundMedia.push(medium);
      } else if (medium.type === "image") {
        imageMedia.push(medium);
      } else if (medium.type === "watermark") {
        watermarkMedia.push(medium);
      } else if (medium.type === "headerImage") {
        headerImageMedia.push(medium);
      }
    });

    // Handle background images
    backgroundMedia.forEach(medium => {
      const rId = nextRid(rels);
      const bookImage = options.media[medium.imageId];
      rels.push({
        Id: rId,
        Type: RelType.Image,
        Target: resolveMediaTarget(bookImage)
      });
      model.background = { rId };
      model.image = options.media[medium.imageId];
    });

    // Handle embedded images — create drawing model using shared utility
    if (imageMedia.length > 0) {
      let { drawing } = model;
      if (!drawing) {
        drawing = model.drawing = {
          rId: nextRid(rels),
          name: `drawing${++options.drawingsCount}`,
          anchors: [],
          rels: []
        };
        options.drawings.push(drawing);
        rels.push({
          Id: drawing.rId,
          Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
          Target: drawingRelTargetFromWorksheet(drawing.name)
        });
      }

      const result = buildDrawingAnchorsAndRels(imageMedia, drawing.rels, {
        getBookImage: id => options.media[id as number],
        nextRId: currentRels => nextRid(currentRels)
      });
      drawing.anchors.push(...result.anchors);
      drawing.rels = result.rels;
    }

    // Handle watermark overlay images — placed as a full-sheet drawing with transparency
    if (watermarkMedia.length > 0) {
      let { drawing } = model;
      if (!drawing) {
        drawing = model.drawing = {
          rId: nextRid(rels),
          name: `drawing${++options.drawingsCount}`,
          anchors: [],
          rels: []
        };
        options.drawings.push(drawing);
        rels.push({
          Id: drawing.rId,
          Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
          Target: drawingRelTargetFromWorksheet(drawing.name)
        });
      }

      for (const medium of watermarkMedia) {
        const bookImage = options.media[medium.imageId];
        if (!bookImage) {
          continue;
        }
        const rIdImage = nextRid(drawing.rels);
        drawing.rels.push({
          Id: rIdImage,
          Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          Target: resolveMediaTarget(bookImage)
        });
        // Convert opacity (0-1) to OOXML percentage (0-100000), clamped
        const rawOpacity = medium.opacity !== undefined ? medium.opacity : 0.15;
        const clampedOpacity = Math.max(0, Math.min(1, rawOpacity));
        const alphaModFix = Math.round(clampedOpacity * 100000);

        // Compute coverage based on actual worksheet dimensions.
        // Use the model's dimensions if available, otherwise use generous defaults.
        const dims = model.dimensions;
        const maxCol = dims ? Math.max(dims.model?.right ?? 100, 100) : 100;
        const maxRow = dims ? Math.max(dims.model?.bottom ?? 200, 200) : 200;

        drawing.anchors.push({
          picture: {
            rId: rIdImage,
            alphaModFix
          },
          // Cover the full data area with extra margin
          range: {
            editAs: "absolute",
            tl: { nativeCol: 0, nativeColOff: 0, nativeRow: 0, nativeRowOff: 0 },
            br: { nativeCol: maxCol, nativeColOff: 0, nativeRow: maxRow, nativeRowOff: 0 }
          }
        });
      }
    }

    // Handle header watermark images — VML header/footer image
    if (headerImageMedia.length > 0) {
      const medium = headerImageMedia[0]; // Only one header image per sheet
      const bookImage = options.media[medium.imageId];
      if (bookImage) {
        const rIdVml = nextRid(rels);
        rels.push({
          Id: rIdVml,
          Type: RelType.VmlDrawing,
          Target: vmlDrawingHFRelTargetFromWorksheet(fileIndex)
        });
        // Store header image info on the model for the VML writer and worksheet render
        model.headerImage = {
          vmlRelId: rIdVml,
          imageId: medium.imageId,
          bookImage,
          headerWidth: medium.headerWidth,
          headerHeight: medium.headerHeight
        };

        // Flag for content-types registration
        options.hasHeaderWatermark = true;

        // Update headerFooter to include &G placeholder.
        // Respects the applyTo option: "all" (default), "odd", "even", "first".
        if (!model.headerFooter) {
          model.headerFooter = {};
        }
        const applyTo = medium.applyTo || "all";
        const insertG = (field: string): string => {
          const existing = (model.headerFooter as any)[field] || "";
          if (existing.includes("&G")) {
            return existing;
          }
          if (existing.includes("&C")) {
            return existing.replace("&C", "&C&G");
          }
          return existing + "&C&G";
        };

        if (applyTo === "all" || applyTo === "odd") {
          model.headerFooter.oddHeader = insertG("oddHeader");
        }
        if (applyTo === "all" || applyTo === "even") {
          model.headerFooter.evenHeader = insertG("evenHeader");
          model.headerFooter.differentOddEven = true;
        }
        if (applyTo === "all" || applyTo === "first") {
          model.headerFooter.firstHeader = insertG("firstHeader");
          model.headerFooter.differentFirst = true;
        }
      }
    }

    // prepare tables
    model.tables.forEach(table => {
      // relationships
      const rId = nextRid(rels);
      table.rId = rId;
      rels.push({
        Id: rId,
        Type: RelType.Table,
        Target: tableRelTargetFromWorksheet(table.target)
      });

      // dynamic styles
      table.columns.forEach(column => {
        const { style } = column;
        if (style) {
          column.dxfId = options.styles.addDxfStyle(style);
        }
      });
    });

    // prepare pivot tables
    (model.pivotTables ?? []).forEach((pivotTable: any) => {
      rels.push({
        Id: nextRid(rels),
        Type: RelType.PivotTable,
        Target: pivotTableRelTargetFromWorksheet(pivotTable.tableNumber)
      });
    });

    // prepare form controls (legacy checkboxes)
    // Form controls share the VML file with comments, but need separate ctrlProp relationships
    if (model.formControls && model.formControls.length > 0) {
      // Ensure a DrawingML drawing part exists for form controls.
      // Excel often repairs sheets that have legacy controls but no <drawing> part.
      let { drawing } = model;
      if (!drawing) {
        drawing = model.drawing = {
          rId: nextRid(rels),
          name: `drawing${++options.drawingsCount}`,
          anchors: [],
          rels: []
        };
        options.drawings.push(drawing);
        rels.push({
          Id: drawing.rId,
          Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
          Target: drawingRelTargetFromWorksheet(drawing.name)
        });
      }

      // If no comments, we need to add the VML drawing relationship for form controls
      if (model.comments.length === 0) {
        rels.push({
          Id: nextRid(rels),
          Type: RelType.VmlDrawing,
          Target: vmlDrawingRelTargetFromWorksheet(fileIndex)
        });
      }

      // Add hidden DrawingML shapes that bridge to the VML shape ids.
      // This mirrors what Excel writes when it "repairs" legacy form controls.
      const toNativePos = (p: any) => ({
        nativeCol: p.col,
        nativeColOff: p.colOff,
        nativeRow: p.row,
        nativeRowOff: p.rowOff
      });

      // Add ctrlProp relationships for each form control
      for (const control of model.formControls) {
        const globalCtrlPropId = options.formControlRefs.length + 1;
        control.ctrlPropId = globalCtrlPropId;
        const relId = nextRid(rels);
        control.ctrlPropRelId = relId;
        rels.push({
          Id: relId,
          Type: RelType.CtrlProp,
          Target: ctrlPropRelTargetFromWorksheet(globalCtrlPropId)
        });
        options.formControlRefs.push(globalCtrlPropId);

        const defaultName = `Check Box ${Math.max(1, control.shapeId - 1024)}`;
        drawing.anchors.push({
          range: {
            editAs: "absolute",
            tl: toNativePos(control.tl),
            br: toNativePos(control.br)
          },
          alternateContent: { requires: "a14" },
          shape: {
            cNvPrId: control.shapeId,
            name: (control as any).name || defaultName,
            hidden: true,
            spid: `_x0000_s${control.shapeId}`,
            text: control.text
          }
        });
      }
    }

    // prepare ext items
    this.map.extLst.prepare(model, options);
  }

  render(xmlStream, model) {
    xmlStream.openXml(StdDocAttributes);
    const worksheetAttrs: any = { ...WorkSheetXform.WORKSHEET_ATTRIBUTES };
    if (model.formControls && model.formControls.length > 0) {
      worksheetAttrs["xmlns:x14"] = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
      worksheetAttrs["xmlns:xdr"] =
        "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
      worksheetAttrs["mc:Ignorable"] = `${worksheetAttrs["mc:Ignorable"]} x14`;
    }
    xmlStream.openNode("worksheet", worksheetAttrs);

    const sheetFormatPropertiesModel: any = model.properties
      ? {
          defaultRowHeight: model.properties.defaultRowHeight,
          dyDescent: model.properties.dyDescent,
          outlineLevelCol: model.properties.outlineLevelCol,
          outlineLevelRow: model.properties.outlineLevelRow,
          customHeight: model.properties.customHeight
        }
      : undefined;
    if (model.properties && model.properties.defaultColWidth) {
      sheetFormatPropertiesModel.defaultColWidth = model.properties.defaultColWidth;
    }
    const sheetPropertiesModel = {
      outlineProperties: model.properties && model.properties.outlineProperties,
      tabColor: model.properties && model.properties.tabColor,
      pageSetup:
        model.pageSetup && model.pageSetup.fitToPage
          ? {
              fitToPage: model.pageSetup.fitToPage
            }
          : undefined
    };
    const pageMarginsModel = model.pageSetup && model.pageSetup.margins;
    const printOptionsModel = {
      showRowColHeaders: model.pageSetup && model.pageSetup.showRowColHeaders,
      showGridLines: model.pageSetup && model.pageSetup.showGridLines,
      horizontalCentered: model.pageSetup && model.pageSetup.horizontalCentered,
      verticalCentered: model.pageSetup && model.pageSetup.verticalCentered
    };
    const sheetProtectionModel = model.sheetProtection;

    this.map.sheetPr.render(xmlStream, sheetPropertiesModel);
    this.map.dimension.render(xmlStream, model.dimensions);
    this.map.sheetViews.render(xmlStream, model.views);
    this.map.sheetFormatPr.render(xmlStream, sheetFormatPropertiesModel);
    this.map.cols.render(xmlStream, model.cols);
    this.map.sheetData.render(xmlStream, model.rows);
    this.map.sheetProtection.render(xmlStream, sheetProtectionModel); // Note: must be after sheetData and before autoFilter
    this.map.autoFilter.render(xmlStream, model.autoFilter);
    this.map.mergeCells.render(xmlStream, model.mergeCells);
    this.map.conditionalFormatting.render(xmlStream, model.conditionalFormattings); // Note: must be before dataValidations
    this.map.dataValidations.render(xmlStream, model.dataValidations);

    // For some reason hyperlinks have to be after the data validations
    this.map.hyperlinks.render(xmlStream, model.hyperlinks);

    this.map.printOptions.render(xmlStream, printOptionsModel); // Note: must be before pageMargins
    this.map.pageMargins.render(xmlStream, pageMarginsModel);
    this.map.pageSetup.render(xmlStream, model.pageSetup);
    this.map.headerFooter.render(xmlStream, model.headerFooter);
    this.map.rowBreaks.render(xmlStream, model.rowBreaks);
    this.map.colBreaks.render(xmlStream, model.colBreaks);
    // `ignoredErrors` must precede `drawing` per ECMA-376 §18.3.1.99
    // CT_Worksheet (… colBreaks, customProperties, cellWatches,
    // ignoredErrors, smartTags, drawing, legacyDrawing, …). Emitting
    // it after `controls` produced out-of-order XML that strict
    // validators reject and some Excel builds repair on open
    // (dropping the ignoredErrors block entirely). Previously the
    // comment on this call claimed the opposite order — it was
    // factually wrong.
    this.map.ignoredErrors.render(xmlStream, model.ignoredErrors);
    this.map.drawing.render(xmlStream, model.drawing); // Note: must be after rowBreaks/colBreaks

    // ECMA-376 §18.3.1.99 CT_Worksheet child sequence is:
    //   … drawing → legacyDrawing → legacyDrawingHF → drawingHF → picture → oleObjects → controls →
    //   webPublishItems → tableParts → extLst
    // `legacyDrawing` and `legacyDrawingHF` therefore MUST be emitted
    // before the background `<picture>` element. The previous order here
    // had picture before legacyDrawing, which validates against strict
    // OOXML checkers as out-of-order and was the root cause of several
    // "Excel needs to repair" reports.
    if (model.rels) {
      // Add a <legacyDrawing /> node for each VML drawing relationship (comments and/or form controls).
      model.rels.forEach(rel => {
        if (rel.Type === RelType.VmlDrawing) {
          // Skip VML rels that are for header images (they use legacyDrawingHF instead)
          if (model.headerImage && rel.Id === model.headerImage.vmlRelId) {
            return;
          }
          xmlStream.leafNode("legacyDrawing", { "r:id": rel.Id });
        }
      });
    }

    // legacyDrawingHF — VML drawing for header/footer images (watermark in header mode)
    if (model.headerImage) {
      xmlStream.leafNode("legacyDrawingHF", { "r:id": model.headerImage.vmlRelId });
    }

    // Background image: must be emitted AFTER drawing/legacyDrawing per
    // the CT_Worksheet sequence noted above.
    this.map.picture.render(xmlStream, model.background);

    // Controls section for legacy form controls (checkboxes, etc.)
    // Excel expects <controls> entries that reference ctrlProp relationships.
    if (model.formControls && model.formControls.length > 0) {
      xmlStream.openNode("mc:AlternateContent");
      xmlStream.openNode("mc:Choice", { Requires: "x14" });
      xmlStream.openNode("controls");

      for (const control of model.formControls) {
        if (!control.ctrlPropRelId) {
          continue;
        }

        const defaultName = `Check Box ${Math.max(1, control.shapeId - 1024)}`;
        xmlStream.openNode("mc:AlternateContent");
        xmlStream.openNode("mc:Choice", { Requires: "x14" });
        xmlStream.openNode("control", {
          shapeId: control.shapeId,
          "r:id": control.ctrlPropRelId,
          name: (control as any).name || defaultName
        });
        xmlStream.openNode("controlPr", {
          locked: 0,
          defaultSize: 0,
          print: control.print ? 1 : 0,
          autoFill: 0,
          autoLine: 0,
          autoPict: 0
        });
        xmlStream.openNode("anchor");
        xmlStream.openNode("from");
        xmlStream.leafNode("xdr:col", undefined, control.tl.col);
        xmlStream.leafNode("xdr:colOff", undefined, control.tl.colOff);
        xmlStream.leafNode("xdr:row", undefined, control.tl.row);
        xmlStream.leafNode("xdr:rowOff", undefined, control.tl.rowOff);
        xmlStream.closeNode();
        xmlStream.openNode("to");
        xmlStream.leafNode("xdr:col", undefined, control.br.col);
        xmlStream.leafNode("xdr:colOff", undefined, control.br.colOff);
        xmlStream.leafNode("xdr:row", undefined, control.br.row);
        xmlStream.leafNode("xdr:rowOff", undefined, control.br.rowOff);
        xmlStream.closeNode(); // to
        xmlStream.closeNode(); // anchor
        xmlStream.closeNode(); // controlPr
        xmlStream.closeNode(); // control
        xmlStream.closeNode(); // mc:Choice
        xmlStream.leafNode("mc:Fallback");
        xmlStream.closeNode(); // mc:AlternateContent
      }

      xmlStream.closeNode();
      xmlStream.closeNode();
      xmlStream.leafNode("mc:Fallback");
      xmlStream.closeNode();
    }

    // `ignoredErrors` is rendered earlier (before `drawing`) per
    // ECMA-376 §18.3.1.99 schema order — see the block above
    // `this.map.drawing.render(...)`. Only `tableParts` and `extLst`
    // remain in this tail section.
    this.map.tableParts.render(xmlStream, model.tables);

    // extLst should be the last element in the worksheet.
    this.map.extLst.render(xmlStream, model);

    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    if (node.name === "worksheet") {
      Object.values(this.map).forEach((xform: any) => {
        xform.reset();
      });
      return true;
    }

    if (this.map[node.name] && !this.ignoreNodes.includes(node.name)) {
      this.parser = this.map[node.name];
      this.parser.parseOpen(node);
    }
    return true;
  }

  parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case "worksheet": {
        const properties = this.map.sheetFormatPr.model || {};
        if (this.map.sheetPr.model && this.map.sheetPr.model.tabColor) {
          properties.tabColor = this.map.sheetPr.model.tabColor;
        }
        if (this.map.sheetPr.model && this.map.sheetPr.model.outlineProperties) {
          properties.outlineProperties = this.map.sheetPr.model.outlineProperties;
        }
        const sheetProperties = {
          fitToPage:
            (this.map.sheetPr.model &&
              this.map.sheetPr.model.pageSetup &&
              this.map.sheetPr.model.pageSetup.fitToPage) ||
            false,
          margins: this.map.pageMargins.model
        };
        const pageSetup = Object.assign(
          sheetProperties,
          this.map.pageSetup.model,
          this.map.printOptions.model
        );
        const conditionalFormattings = mergeConditionalFormattings(
          this.map.conditionalFormatting.model,
          this.map.extLst.model && this.map.extLst.model["x14:conditionalFormattings"]
        );
        this.model = {
          dimensions: this.map.dimension.model,
          cols: this.map.cols.model,
          rows: this.map.sheetData.model,
          mergeCells: this.map.mergeCells.model,
          hyperlinks: this.map.hyperlinks.model,
          dataValidations: this.map.dataValidations.model,
          properties,
          views: this.map.sheetViews.model,
          pageSetup,
          headerFooter: this.map.headerFooter.model,
          background: this.map.picture.model,
          drawing: this.map.drawing.model,
          tables: this.map.tableParts.model,
          conditionalFormattings,
          rowBreaks: this.map.rowBreaks.model ?? [],
          colBreaks: this.map.colBreaks.model ?? [],
          ignoredErrors: this.map.ignoredErrors.model ?? []
        };

        if (this.map.autoFilter.model) {
          this.model.autoFilter = this.map.autoFilter.model;
        }
        if (this.map.sheetProtection.model) {
          this.model.sheetProtection = this.map.sheetProtection.model;
        }

        return false;
      }

      default:
        // not quite sure how we get here!
        return true;
    }
  }

  reconcile(model, options) {
    // options.merges = new Merges();
    // options.merges.reconcile(model.mergeCells, model.rows);
    // Build rel index first, then process comments and VML in two passes so
    // that the result is independent of the order rels appear in the file.
    const relList = model.relationships ?? [];
    const rels = relList.reduce((h, rel) => {
      h[rel.Id] = rel;
      return h;
    }, {});

    // Pass 1: resolve comments
    for (const rel of relList) {
      if (rel.Type === RelType.Comments) {
        const resolvedPath = resolveRelTarget("xl/worksheets/", rel.Target);
        const commentEntry = options.comments?.[resolvedPath];
        if (commentEntry) {
          model.comments = commentEntry.comments;
        }
      }
    }

    // Pass 2: merge VML note metadata (requires model.comments from pass 1)
    if (model.comments && model.comments.length) {
      for (const rel of relList) {
        if (rel.Type === RelType.VmlDrawing) {
          const resolvedVmlPath = resolveRelTarget("xl/worksheets/", rel.Target);
          const vmlEntry = options.vmlDrawings?.[resolvedVmlPath];
          if (vmlEntry) {
            // Build a ref-keyed map from VML comments for order-independent merge.
            // Fall back to index-based merge if VML entries lack row/col.
            const vmlComments = vmlEntry.comments;
            const vmlByRef: Record<string, any> = {};
            let hasRefInfo = false;
            for (const vc of vmlComments) {
              if (vc.row != null && vc.col != null) {
                const ref = colCache.encodeAddress(vc.row, vc.col);
                vmlByRef[ref] = vc;
                hasRefInfo = true;
              }
            }

            if (hasRefInfo) {
              // Merge by cell reference (robust against order differences)
              for (const comment of model.comments) {
                const vml = vmlByRef[comment.ref];
                if (vml) {
                  comment.note = Object.assign({}, comment.note, vml);
                }
              }
            } else {
              // Fallback: index-based merge for VML files without row/col
              model.comments.forEach((comment, index) => {
                if (index < vmlComments.length) {
                  comment.note = Object.assign({}, comment.note, vmlComments[index]);
                }
              });
            }
          }
        }
      }
    }
    options.commentsMap = (model.comments ?? []).reduce((h, comment) => {
      if (comment.ref) {
        h[comment.ref] = comment;
      }
      return h;
    }, {});
    options.hyperlinkMap = (model.hyperlinks ?? []).reduce((h, hyperlink) => {
      if (hyperlink.rId) {
        // External link: resolve target from relationship
        const rel = rels[hyperlink.rId];
        if (rel) {
          h[hyperlink.address] = rel.Target;
        }
      } else if (hyperlink.target) {
        // Internal link: target was restored from location attribute (with "#" prefix)
        h[hyperlink.address] = hyperlink.target;
      }
      return h;
    }, {});
    options.formulae = {};

    // compact the rows and cells — remove any holes from sparse parse results
    if (model.rows) {
      if (model.rows.includes(undefined as any)) {
        model.rows = model.rows.filter(Boolean);
      }
      for (let i = 0; i < model.rows.length; i++) {
        const row = model.rows[i];
        if (row.cells?.includes(undefined as any)) {
          row.cells = row.cells.filter(Boolean);
        }
      }
    } else {
      model.rows = [];
    }

    this.map.cols.reconcile(model.cols, options);
    this.map.sheetData.reconcile(model.rows, options);
    this.map.conditionalFormatting.reconcile(model.conditionalFormattings, options);

    model.media = [];
    if (model.drawing) {
      const drawingRel = rels[model.drawing.rId];
      if (drawingRel) {
        const match = drawingRel.Target.match(/\/drawings\/([a-zA-Z0-9]+)[.][a-zA-Z]{3,4}$/);
        if (match) {
          const drawingName = match[1];
          const drawing = options.drawings[drawingName];
          if (drawing) {
            // Preserve the drawing object for round-trip (charts, etc.)
            // This includes the name, anchors, and rels
            model.drawing = {
              ...drawing,
              name: drawingName,
              rels: options.drawingRels?.[drawingName] ?? drawing.rels ?? []
            };

            // Also extract images to model.media for backward compatibility.
            drawing.anchors.forEach(anchor => {
              if (anchor.medium) {
                // Detect overlay watermarks: drawings that carry alphaModFix
                const hasAlpha =
                  anchor.medium.alphaModFix !== undefined && anchor.medium.alphaModFix < 100000;
                if (hasAlpha) {
                  model.media.push({
                    type: "watermark",
                    imageId: anchor.medium.index,
                    opacity: anchor.medium.alphaModFix / 100000
                  });
                } else {
                  model.media.push({
                    type: "image",
                    imageId: anchor.medium.index,
                    range: anchor.range,
                    hyperlinks: anchor.picture.hyperlinks
                  });
                }
              }
            });
          } else {
            // Drawing data not found - clear the stale reference
            model.drawing = undefined;
          }
        } else {
          // Target path doesn't match expected drawing pattern
          model.drawing = undefined;
        }
      } else {
        // Relationship missing (corrupted/malicious file) - clear stale reference
        model.drawing = undefined;
      }
    }

    const backgroundRel = model.background && rels[model.background.rId];
    if (backgroundRel) {
      const target = backgroundRel.Target.split("/media/")[1];
      const imageId = options.mediaIndex && options.mediaIndex[target];
      if (imageId !== undefined) {
        model.media.push({
          type: "background",
          imageId
        });
      }
    } else if (model.background) {
      // Relationship missing - clear stale reference
      model.background = undefined;
    }

    model.tables = (model.tables ?? []).reduce((acc, tablePart) => {
      const rel = rels[tablePart.rId];
      if (rel) {
        const resolvedPath = resolveRelTarget("xl/worksheets/", rel.Target);
        const table = options.tables[resolvedPath];
        if (table) {
          acc.push(table);
        }
      }
      return acc;
    }, []);

    // Link pivot tables from relationships to worksheet
    // This is needed so that when writing, the worksheet knows which pivot tables it contains
    model.pivotTables = [];
    (model.relationships ?? []).forEach(rel => {
      if (rel.Type === RelType.PivotTable && options.pivotTables) {
        const resolvedPath = resolveRelTarget("xl/worksheets/", rel.Target);
        const pivotTable = options.pivotTables[resolvedPath];
        if (pivotTable) {
          model.pivotTables.push(pivotTable);
        }
      }
    });

    delete model.relationships;
    delete model.hyperlinks;
    delete model.comments;
  }
}

export { WorkSheetXform };
