/**
 * Child order inside the parts this library writes.
 *
 * Most OOXML containers are an `xsd:sequence`, and Excel refuses — or offers to "repair" — a part whose children
 * are out of order. The worksheet and workbook checks rank the root's children; this one ranks them below the root,
 * where an out-of-order child is just as fatal and nothing looked (issue #269: `<sheetPr>` written as
 * `<pageSetUpPr/><outlinePr/>`).
 *
 * Each part is described by a tree of local names from its root. A node naming a type ranks its element's children
 * by that type's content model, which comes from `ooxml-sequences.ts`, generated from the ECMA-376 schemas — so the
 * only thing written here by hand is _where_ a type occurs, never its order. A path rather than a bare local name
 * keeps a rule to the content model it was written for: an `x14:cfRule` in an `extLst` shares a local name with
 * `cfRule` but not its children.
 *
 * A child the type does not name is skipped, not reported: an extension element may appear anywhere this check
 * cannot see, and the validator also runs on files read from elsewhere and written back. `mc:AlternateContent` is
 * unwrapped to its `mc:Choice`, whose children stand in the wrapper's place.
 */

import { CT } from "@excel/utils/ooxml-validator/check-content-types";
import type { ValidationContext } from "@excel/utils/ooxml-validator/context";
import {
  OOXML_SEQUENCES,
  type OoxmlSequenceType
} from "@excel/utils/ooxml-validator/ooxml-sequences";
import {
  checkSequenceOrder,
  describeSequence,
  sequenceRanks,
  type SequenceNode as GenericSequenceNode
} from "@xml/sequence-order";

export type SequenceNode = GenericSequenceNode<OoxmlSequenceType>;

// -----------------------------------------------------------------------------
// Shared subtrees
// -----------------------------------------------------------------------------

/** CT_Rst — a shared string, an inline string and a comment's text. */
const RST: SequenceNode = { type: "x:CT_Rst", children: { r: { type: "x:CT_RElt" } } };

const SORT_STATE: SequenceNode = { type: "x:CT_SortState" };

const AUTO_FILTER: SequenceNode = { type: "x:CT_AutoFilter", children: { sortState: SORT_STATE } };

const BORDER: SequenceNode = { type: "x:CT_Border" };

const FILL: SequenceNode = { children: { patternFill: { type: "x:CT_PatternFill" } } };

const LINE: SequenceNode = { type: "a:CT_LineProperties" };

const SHAPE_PROPERTIES: SequenceNode = {
  type: "a:CT_ShapeProperties",
  children: { xfrm: { type: "a:CT_Transform2D" }, ln: LINE }
};

const TEXT_CHARACTER: SequenceNode = {
  type: "a:CT_TextCharacterProperties",
  children: { ln: LINE }
};

const TEXT_PARAGRAPH_PROPERTIES: SequenceNode = {
  type: "a:CT_TextParagraphProperties",
  children: { defRPr: TEXT_CHARACTER }
};

/** CT_TextBody — a shape's `txBody`, and a chart's `txPr` and `rich` (see `check-chart.ts`). */
const TEXT_BODY: SequenceNode = {
  type: "a:CT_TextBody",
  children: {
    bodyPr: { type: "a:CT_TextBodyProperties" },
    lstStyle: {
      type: "a:CT_TextListStyle",
      children: Object.fromEntries(
        ["defPPr", ...Array.from({ length: 9 }, (_, i) => `lvl${i + 1}pPr`)].map(name => [
          name,
          TEXT_PARAGRAPH_PROPERTIES
        ])
      )
    },
    p: {
      type: "a:CT_TextParagraph",
      children: {
        pPr: TEXT_PARAGRAPH_PROPERTIES,
        r: { type: "a:CT_RegularTextRun", children: { rPr: TEXT_CHARACTER } },
        endParaRPr: TEXT_CHARACTER
      }
    }
  }
};

const C_NV_PR: SequenceNode = { type: "a:CT_NonVisualDrawingProps" };

const PICTURE: SequenceNode = {
  type: "xdr:CT_Picture",
  children: {
    nvPicPr: {
      type: "xdr:CT_PictureNonVisual",
      children: { cNvPr: C_NV_PR, cNvPicPr: { type: "a:CT_NonVisualPictureProperties" } }
    },
    blipFill: { type: "a:CT_BlipFillProperties" },
    spPr: SHAPE_PROPERTIES,
    style: { type: "a:CT_ShapeStyle" }
  }
};

const SHAPE: SequenceNode = {
  type: "xdr:CT_Shape",
  children: {
    nvSpPr: { type: "xdr:CT_ShapeNonVisual", children: { cNvPr: C_NV_PR } },
    spPr: SHAPE_PROPERTIES,
    style: { type: "a:CT_ShapeStyle" },
    txBody: TEXT_BODY
  }
};

const GRAPHIC_FRAME: SequenceNode = {
  type: "xdr:CT_GraphicalObjectFrame",
  children: {
    nvGraphicFramePr: { type: "xdr:CT_GraphicalObjectFrameNonVisual", children: { cNvPr: C_NV_PR } }
  }
};

const MARKER: SequenceNode = { type: "xdr:CT_Marker" };

const ANCHOR_CONTENT = { pic: PICTURE, sp: SHAPE, graphicFrame: GRAPHIC_FRAME } as const;

// -----------------------------------------------------------------------------
// Parts
// -----------------------------------------------------------------------------

const X14_CFVO: SequenceNode = { type: "x14:CT_Cfvo" };

/** What a worksheet's `x14:` extensions hold: conditional formats, data validations and sparklines. */
const X14_WORKSHEET_EXTENSIONS: Readonly<Record<string, SequenceNode>> = {
  conditionalFormattings: {
    children: {
      conditionalFormatting: {
        type: "x14:CT_ConditionalFormatting",
        children: {
          cfRule: {
            type: "x14:CT_CfRule",
            children: {
              colorScale: { type: "x14:CT_ColorScale", children: { cfvo: X14_CFVO } },
              dataBar: { type: "x14:CT_DataBar", children: { cfvo: X14_CFVO } },
              iconSet: { type: "x14:CT_IconSet", children: { cfvo: X14_CFVO } }
            }
          }
        }
      }
    }
  },
  dataValidations: { children: { dataValidation: { type: "x14:CT_DataValidation" } } },
  sparklineGroups: {
    children: {
      sparklineGroup: {
        type: "x14:CT_SparklineGroup",
        children: { sparklines: { children: { sparkline: { type: "x14:CT_Sparkline" } } } }
      }
    }
  }
};

/** The root's own children are ranked by `checkWorksheets` (`CT_Worksheet`), which reports them as its own kind. */
const WORKSHEET: SequenceNode = {
  children: {
    sheetPr: { type: "x:CT_SheetPr" },
    sheetViews: { type: "x:CT_SheetViews", children: { sheetView: { type: "x:CT_SheetView" } } },
    sheetData: {
      children: {
        row: { type: "x:CT_Row", children: { c: { type: "x:CT_Cell", children: { is: RST } } } }
      }
    },
    autoFilter: AUTO_FILTER,
    sortState: SORT_STATE,
    conditionalFormatting: {
      type: "x:CT_ConditionalFormatting",
      children: {
        cfRule: {
          type: "x:CT_CfRule",
          children: { colorScale: { type: "x:CT_ColorScale" }, dataBar: { type: "x:CT_DataBar" } }
        }
      }
    },
    dataValidations: { children: { dataValidation: { type: "x:CT_DataValidation" } } },
    headerFooter: { type: "x:CT_HeaderFooter" },
    // Microsoft's `x14:` extensions, which reach a worksheet through its `extLst`. The path keeps them apart from
    // the ECMA elements of the same local name (`cfRule`, `dataValidation`), whose content models differ.
    extLst: { children: { ext: { children: X14_WORKSHEET_EXTENSIONS } } }
  }
};

const CHARTSHEET: SequenceNode = {
  type: "x:CT_Chartsheet",
  children: { headerFooter: { type: "x:CT_HeaderFooter" } }
};

const STYLES: SequenceNode = {
  type: "x:CT_Stylesheet",
  children: {
    fills: { children: { fill: FILL } },
    borders: { children: { border: BORDER } },
    cellStyleXfs: { children: { xf: { type: "x:CT_Xf" } } },
    cellXfs: { children: { xf: { type: "x:CT_Xf" } } },
    dxfs: { children: { dxf: { type: "x:CT_Dxf", children: { fill: FILL, border: BORDER } } } }
  }
};

const SHARED_STRINGS: SequenceNode = { type: "x:CT_Sst", children: { si: RST } };

const COMMENTS: SequenceNode = {
  type: "x:CT_Comments",
  children: {
    commentList: { children: { comment: { type: "x:CT_Comment", children: { text: RST } } } }
  }
};

const TABLE: SequenceNode = {
  type: "x:CT_Table",
  children: {
    autoFilter: AUTO_FILTER,
    sortState: SORT_STATE,
    tableColumns: { children: { tableColumn: { type: "x:CT_TableColumn" } } }
  }
};

const DRAWING: SequenceNode = {
  children: {
    twoCellAnchor: {
      type: "xdr:CT_TwoCellAnchor",
      children: { from: MARKER, to: MARKER, ...ANCHOR_CONTENT }
    },
    oneCellAnchor: { type: "xdr:CT_OneCellAnchor", children: { from: MARKER, ...ANCHOR_CONTENT } },
    absoluteAnchor: { type: "xdr:CT_AbsoluteAnchor", children: ANCHOR_CONTENT }
  }
};

const EXTERNAL_LINK: SequenceNode = { children: { externalBook: { type: "x:CT_ExternalBook" } } };

const PIVOT_TABLE: SequenceNode = {
  type: "x:CT_pivotTableDefinition",
  children: { pivotFields: { children: { pivotField: { type: "x:CT_PivotField" } } } }
};

const PIVOT_CACHE_DEFINITION: SequenceNode = {
  type: "x:CT_PivotCacheDefinition",
  children: { cacheFields: { children: { cacheField: { type: "x:CT_CacheField" } } } }
};

/** The parts checked, by content type. The workbook has none: `checkWorkbook` ranks its root, and nothing below. */
export const SEQUENCE_PARTS: ReadonlyMap<string, SequenceNode> = new Map([
  [CT.worksheet, WORKSHEET],
  [CT.chartsheet, CHARTSHEET],
  [CT.styles, STYLES],
  [CT.sharedStrings, SHARED_STRINGS],
  [CT.comments, COMMENTS],
  [CT.tableDef, TABLE],
  [CT.drawing, DRAWING],
  [CT.externalLink, EXTERNAL_LINK],
  [CT.pivotTable, PIVOT_TABLE],
  [CT.pivotCacheDefinition, PIVOT_CACHE_DEFINITION]
]);

const RANKS = sequenceRanks(OOXML_SEQUENCES);

// -----------------------------------------------------------------------------
// Checker
// -----------------------------------------------------------------------------

export function checkSequences(ctx: ValidationContext): void {
  const { overrides } = ctx.readContentTypes();
  for (const [path, contentType] of overrides) {
    if (ctx.reporter.capped) {
      return;
    }
    const tree = SEQUENCE_PARTS.get(contentType);
    if (!tree) {
      continue;
    }
    // Malformed XML is reported by the well-formedness check; this one has nothing to rank.
    const dom = ctx.readDom(path);
    if (!dom) {
      continue;
    }
    checkSequenceOrder(
      dom.root,
      tree,
      RANKS,
      ({ trail, child, after, type }) =>
        ctx.reporter.error(
          "part-child-out-of-order",
          `<${trail.join("/")}> child <${child}> appears after <${after}>; ${type.slice(type.indexOf(":") + 1)} requires the order ${describeSequence(OOXML_SEQUENCES, type)}`,
          path
        ),
      () => ctx.reporter.capped
    );
  }
}
