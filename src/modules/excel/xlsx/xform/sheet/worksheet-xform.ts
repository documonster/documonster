import { colCache } from "@excel/utils/col-cache";
import { XmlStream } from "@excel/utils/xml-stream";
import { RelType } from "@excel/xlsx/rel-type";
import { Merges } from "@excel/xlsx/xform/sheet/merges";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { isInternalLink, HyperlinkXform } from "@excel/xlsx/xform/sheet/hyperlink-xform";
import { ListXform } from "@excel/xlsx/xform/list-xform";
import { RowXform } from "@excel/xlsx/xform/sheet/row-xform";
import { ColXform } from "@excel/xlsx/xform/sheet/col-xform";
import { DimensionXform } from "@excel/xlsx/xform/sheet/dimension-xform";
import { MergeCellXform } from "@excel/xlsx/xform/sheet/merge-cell-xform";
import { DataValidationsXform } from "@excel/xlsx/xform/sheet/data-validations-xform";
import { SheetPropertiesXform } from "@excel/xlsx/xform/sheet/sheet-properties-xform";
import { SheetFormatPropertiesXform } from "@excel/xlsx/xform/sheet/sheet-format-properties-xform";
import { SheetViewXform } from "@excel/xlsx/xform/sheet/sheet-view-xform";
import { SheetProtectionXform } from "@excel/xlsx/xform/sheet/sheet-protection-xform";
import { PageMarginsXform } from "@excel/xlsx/xform/sheet/page-margins-xform";
import { PageSetupXform } from "@excel/xlsx/xform/sheet/page-setup-xform";
import { PrintOptionsXform } from "@excel/xlsx/xform/sheet/print-options-xform";
import { AutoFilterXform } from "@excel/xlsx/xform/sheet/auto-filter-xform";
import { PictureXform } from "@excel/xlsx/xform/sheet/picture-xform";
import { DrawingXform } from "@excel/xlsx/xform/sheet/drawing-xform";
import { TablePartXform } from "@excel/xlsx/xform/sheet/table-part-xform";
import { RowBreaksXform } from "@excel/xlsx/xform/sheet/row-breaks-xform";
import { ColBreaksXform } from "@excel/xlsx/xform/sheet/col-breaks-xform";
import { HeaderFooterXform } from "@excel/xlsx/xform/sheet/header-footer-xform";
import { ConditionalFormattingsXform } from "@excel/xlsx/xform/sheet/cf/conditional-formattings-xform";
import { ExtLstXform } from "@excel/xlsx/xform/sheet/ext-lst-xform";
import {
  commentsRelTargetFromWorksheet,
  ctrlPropRelTargetFromWorksheet,
  drawingRelTargetFromWorksheet,
  pivotTableRelTargetFromWorksheet,
  tableRelTargetFromWorksheet,
  vmlDrawingRelTargetFromWorksheet
} from "@excel/utils/ooxml-paths";
import { buildDrawingAnchorsAndRels, resolveMediaTarget } from "@excel/utils/drawing-utils";

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
      extLst: new ExtLstXform()
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

    // Handle pre-loaded drawing (from file read) that may contain charts or other non-image content.
    // Reset anchors and rels so they are rebuilt cleanly from model.media (images) and
    // model.formControls (shapes) below.  Without this reset, every read-write cycle would
    // duplicate image anchors because the same images exist in both model.drawing.anchors
    // (preserved for round-trip) and model.media (the canonical image list).
    // For chart drawings, rels are preserved because the raw XML passthrough references
    // original rIds; anchors are still cleared since they are unused for chart drawings.
    if (model.drawing && model.drawing.anchors) {
      const drawing = model.drawing;
      drawing.rId = nextRid(rels);
      if (!drawing.name) {
        drawing.name = `drawing${++options.drawingsCount}`;
      }

      const hasChartRels = (drawing.rels ?? []).some(
        (rel: any) => rel.Target && rel.Target.includes("/charts/")
      );
      // Anchors are always reset: for chart drawings they are unused (raw XML passthrough),
      // for normal drawings they are rebuilt from model.media below.
      drawing.anchors = [];
      if (!hasChartRels) {
        // Non-chart drawings: clear rels so image rels are rebuilt from scratch.
        drawing.rels = [];
      }
      // Chart drawings keep their original rels intact since the raw drawing XML
      // references those rIds directly.

      options.drawings.push(drawing);
      rels.push({
        Id: drawing.rId,
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
        Target: drawingRelTargetFromWorksheet(drawing.name)
      });
    }

    // Process background and image media entries
    const backgroundMedia: any[] = [];
    const imageMedia: any[] = [];
    model.media.forEach(medium => {
      if (medium.type === "background") {
        backgroundMedia.push(medium);
      } else if (medium.type === "image") {
        imageMedia.push(medium);
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
    xmlStream.openXml(XmlStream.StdDocAttributes);
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
    this.map.drawing.render(xmlStream, model.drawing); // Note: must be after rowBreaks/colBreaks
    this.map.picture.render(xmlStream, model.background); // Note: must be after drawing

    if (model.rels) {
      // Add a <legacyDrawing /> node for each VML drawing relationship (comments and/or form controls).
      // NOTE: Excel is picky about worksheet child element order; legacyDrawing must come before controls.
      model.rels.forEach(rel => {
        if (rel.Type === RelType.VmlDrawing) {
          xmlStream.leafNode("legacyDrawing", { "r:id": rel.Id });
        }
      });
    }

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

    // Table parts must come after <controls> in worksheet element order.
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
          colBreaks: this.map.colBreaks.model ?? []
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
    const rels = (model.relationships ?? []).reduce((h, rel) => {
      h[rel.Id] = rel;
      if (rel.Type === RelType.Comments) {
        const commentEntry = options.comments?.[rel.Target];
        if (commentEntry) {
          model.comments = commentEntry.comments;
        }
      }
      if (rel.Type === RelType.VmlDrawing && model.comments && model.comments.length) {
        const vmlEntry = options.vmlDrawings?.[rel.Target];
        if (vmlEntry) {
          const vmlComment = vmlEntry.comments;
          model.comments.forEach((comment, index) => {
            comment.note = Object.assign({}, comment.note, vmlComment[index]);
          });
        }
      }
      return h;
    }, {});
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

    // compact the rows and cells
    model.rows = model.rows?.filter(Boolean) ?? [];
    model.rows.forEach(row => {
      row.cells = row.cells?.filter(Boolean) ?? [];
    });

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

            // Also extract images to model.media for backward compatibility
            drawing.anchors.forEach(anchor => {
              if (anchor.medium) {
                const image = {
                  type: "image",
                  imageId: anchor.medium.index,
                  range: anchor.range,
                  hyperlinks: anchor.picture.hyperlinks
                };
                model.media.push(image);
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
        const table = options.tables[rel.Target];
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
        const pivotTable = options.pivotTables[rel.Target];
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
