/**
 * Child order in the WordprocessingML parts this library writes, against the ECMA-376 content models.
 *
 * Most containers in a document are an `xsd:sequence` — `w:pPr`, `w:tblPr`, `w:tcPr`, `w:sectPr`, `w:style`,
 * `w:abstractNum` — and Word refuses, or offers to repair, a part whose children are out of order. The writers emit
 * children in an order written by hand, one `if` after another, and nothing checked it: an `if` placed one line too
 * low is valid TypeScript and an invalid document.
 *
 * The order itself is not written here. It comes from `ooxml-sequences.ts`, generated from the schemas; what this
 * file says is _where_ each content model applies, as a tree of local names from a part's root. The walk is
 * `@xml/sequence-order`'s, shared with the workbook validator.
 *
 * A chart part (`c:chartSpace`) is checked by the chart rules the workbook validator uses
 * (`@xml/ooxml-chart-order`): both modules write charts, and Word's chart writer is its own.
 *
 * A part is recognised by its root element in the WordprocessingML namespace, not by its name or content type, so
 * a header is checked as a header wherever the package put it. A child a model does not name is skipped: an
 * extension element (`w14:`, `w15:`, …) may appear where this check cannot place it.
 *
 * Where a content model has no order — a run's `rPr` and the previous one in its `w:rPrChange` (`CT_RPrOriginal`), a
 * row's previous `w:trPr` (`CT_TrPrBase`), `w:txbxContent` — there is nothing to rank, and the tree only descends
 * through it where something ordered lies below.
 */

import { OOXML_SEQUENCES, type OoxmlSequenceType } from "@word/writer/ooxml-sequences";
import { SEQUENCE_RANKS } from "@word/writer/schema-order";
import { parseXml } from "@xml/dom";
import { chartContentModel } from "@xml/ooxml-chart-order";
import {
  checkSequenceOrder,
  checkSequenceOrderByName,
  describeSequence,
  type SequenceNode,
  type SequenceViolation
} from "@xml/sequence-order";

type Node = SequenceNode<OoxmlSequenceType>;

/** A child out of its content model's order in one part. */
export interface DocxOrderViolation {
  readonly part: string;
  readonly message: string;
}

// -----------------------------------------------------------------------------
// Shared subtrees. Block content nests (a cell holds paragraphs and tables), so the trees are built as mutable
// records and closed over each other below.
// -----------------------------------------------------------------------------

const R_PR: Node = { type: "w:CT_RPr" };

const SECT_PR_CHILDREN: Record<string, Node> = {
  footnotePr: { type: "w:CT_FtnProps" },
  endnotePr: { type: "w:CT_EdnProps" },
  pgBorders: { type: "w:CT_PageBorders" }
};

const SECT_PR: Node = {
  type: "w:CT_SectPr",
  children: {
    ...SECT_PR_CHILDREN,
    sectPrChange: { children: { sectPr: { type: "w:CT_SectPrBase", children: SECT_PR_CHILDREN } } }
  }
};

const P_PR_BASE: Node = {
  type: "w:CT_PPrBase",
  children: { pBdr: { type: "w:CT_PBdr" }, numPr: { type: "w:CT_NumPr" } }
};

/** `w:pPr` in a paragraph (`CT_PPr`), which adds the paragraph mark's run properties, a section and a change. */
const P_PR: Node = {
  type: "w:CT_PPr",
  children: {
    pBdr: { type: "w:CT_PBdr" },
    numPr: { type: "w:CT_NumPr" },
    rPr: {
      type: "w:CT_ParaRPr",
      children: { rPrChange: { children: { rPr: { type: "w:CT_ParaRPrOriginal" } } } }
    },
    sectPr: SECT_PR,
    pPrChange: { children: { pPr: P_PR_BASE } }
  }
};

/** `w:pPr` in a style, a level or the defaults (`CT_PPrGeneral`): the base, then a change. */
const P_PR_GENERAL: Node = {
  type: "w:CT_PPrGeneral",
  children: {
    pBdr: { type: "w:CT_PBdr" },
    numPr: { type: "w:CT_NumPr" },
    pPrChange: { children: { pPr: P_PR_BASE } }
  }
};

const A_SHAPE_PROPERTIES: Node = {
  type: "a:CT_ShapeProperties",
  children: { xfrm: { type: "a:CT_Transform2D" }, ln: { type: "a:CT_LineProperties" } }
};

/** A text box's content: block content, filled in below once that exists. */
const TXBX_CONTENT: Node = { children: {} };

/** A DrawingML shape (`wps:wsp`), whose text box holds block content. */
const WSP: Node = {
  type: "wps:CT_WordprocessingShape",
  children: {
    cNvPr: { type: "a:CT_NonVisualDrawingProps" },
    spPr: A_SHAPE_PROPERTIES,
    txbx: { children: { txbxContent: TXBX_CONTENT } }
  }
};

/**
 * A drawing group (`wpg:wgp`) or canvas (`wpc:wpc`): shapes, pictures and further groups, each descended into for
 * the text boxes they hold. Neither schema is in the tables, so nothing at this level is ranked.
 */
const SHAPE_CONTAINER: Node = { children: {} };

const GRAPHIC: Node = {
  children: {
    graphicData: {
      children: {
        wsp: WSP,
        wgp: SHAPE_CONTAINER,
        wpc: SHAPE_CONTAINER,
        pic: {
          type: "pic:CT_Picture",
          children: {
            nvPicPr: {
              type: "pic:CT_PictureNonVisual",
              children: {
                cNvPr: { type: "a:CT_NonVisualDrawingProps" },
                cNvPicPr: { type: "a:CT_NonVisualPictureProperties" }
              }
            },
            blipFill: { type: "a:CT_BlipFillProperties" },
            spPr: A_SHAPE_PROPERTIES
          }
        }
      }
    }
  }
};

const DOC_PR: Node = { type: "a:CT_NonVisualDrawingProps" };

const WRAP_POLYGON: Node = { children: { wrapPolygon: { type: "wp:CT_WrapPath" } } };

const DRAWING: Node = {
  children: {
    inline: { type: "wp:CT_Inline", children: { docPr: DOC_PR, graphic: GRAPHIC } },
    anchor: {
      type: "wp:CT_Anchor",
      children: {
        docPr: DOC_PR,
        graphic: GRAPHIC,
        wrapTight: WRAP_POLYGON,
        wrapThrough: WRAP_POLYGON
      }
    }
  }
};

/** A VML picture (`w:pict`) or embedded object (`w:object`): shapes whose `v:textbox` holds block content. */
const VML_SHAPE: Node = { children: { textbox: { children: { txbxContent: TXBX_CONTENT } } } };
const VML_SHAPES: Record<string, Node> = Object.fromEntries(
  ["shape", "rect", "roundrect", "oval", "line", "polyline", "arc", "curve", "image"].map(kind => [
    kind,
    VML_SHAPE
  ])
);
const VML_GROUP: Node = { children: VML_SHAPES };
Object.assign(VML_SHAPES, { group: VML_GROUP });
const PICT: Node = { children: VML_SHAPES };

/** Run-level content: what a paragraph, a hyperlink, a revision or a run content control holds. */
const RUN_CONTENT: Record<string, Node> = {};

const R: Node = {
  type: "w:CT_R",
  children: {
    rPr: R_PR,
    drawing: DRAWING,
    pict: PICT,
    object: PICT,
    ruby: {
      type: "w:CT_Ruby",
      children: {
        rubyPr: { type: "w:CT_RubyPr" },
        rt: { children: RUN_CONTENT },
        rubyBase: { children: RUN_CONTENT }
      }
    }
  }
};

/**
 * `w:sdtPr`, with the two Word extensions this library writes inside it: `w14:checkbox` and
 * `w15:repeatingSection`, ranked by [MS-DOCX]'s content models.
 */
const SDT_PR: Node = {
  type: "w:CT_SdtPr",
  children: {
    rPr: R_PR,
    checkbox: { type: "w14:CT_SdtCheckbox" },
    repeatingSection: { type: "w15:CT_SdtRepeatedSection" }
  }
};

const SDT_RUN: Node = {
  type: "w:CT_SdtRun",
  children: { sdtPr: SDT_PR, sdtContent: { children: RUN_CONTENT } }
};

Object.assign(RUN_CONTENT, {
  r: R,
  hyperlink: { children: RUN_CONTENT },
  ins: { children: RUN_CONTENT },
  del: { children: RUN_CONTENT },
  moveFrom: { children: RUN_CONTENT },
  moveTo: { children: RUN_CONTENT },
  smartTag: { children: RUN_CONTENT },
  customXml: { children: RUN_CONTENT },
  dir: { children: RUN_CONTENT },
  bdo: { children: RUN_CONTENT },
  fldSimple: { type: "w:CT_SimpleField", children: RUN_CONTENT },
  sdt: SDT_RUN
});

const P: Node = { type: "w:CT_P", children: { pPr: P_PR, ...RUN_CONTENT } };

/** Block-level content: what the body, a cell, a note, a comment, a header or a block content control holds. */
const BLOCK_CONTENT: Record<string, Node> = { p: P };

const TBL_BORDERS: Node = { type: "w:CT_TblBorders" };
const TBL_CELL_MAR: Node = { type: "w:CT_TblCellMar" };

/** `w:tblPr` in a style or a change record (`CT_TblPrBase`): no change record of its own. */
const TBL_PR_BASE: Node = {
  type: "w:CT_TblPrBase",
  children: { tblBorders: TBL_BORDERS, tblCellMar: TBL_CELL_MAR }
};

const TBL_PR: Node = {
  type: "w:CT_TblPr",
  children: {
    tblBorders: TBL_BORDERS,
    tblCellMar: TBL_CELL_MAR,
    tblPrChange: { children: { tblPr: TBL_PR_BASE } }
  }
};

const TR_PR: Node = { type: "w:CT_TrPr" };

const TC_PR_CHILDREN: Record<string, Node> = {
  tcBorders: { type: "w:CT_TcBorders" },
  tcMar: { type: "w:CT_TcMar" }
};

const TC_PR: Node = {
  type: "w:CT_TcPr",
  children: {
    ...TC_PR_CHILDREN,
    tcPrChange: { children: { tcPr: { type: "w:CT_TcPrInner", children: TC_PR_CHILDREN } } }
  }
};

const TC: Node = { type: "w:CT_Tc", children: { tcPr: TC_PR } };

const TR: Node = {
  type: "w:CT_Row",
  children: {
    tblPrEx: {
      type: "w:CT_TblPrEx",
      children: {
        tblBorders: TBL_BORDERS,
        tblCellMar: TBL_CELL_MAR,
        tblPrExChange: {
          children: {
            tblPrEx: {
              type: "w:CT_TblPrExBase",
              children: { tblBorders: TBL_BORDERS, tblCellMar: TBL_CELL_MAR }
            }
          }
        }
      }
    },
    trPr: TR_PR,
    tc: TC
  }
};

const TBL: Node = {
  type: "w:CT_Tbl",
  children: { tblPr: TBL_PR, tblGrid: { type: "w:CT_TblGrid" }, tr: TR }
};

const SDT_BLOCK: Node = {
  type: "w:CT_SdtBlock",
  children: { sdtPr: SDT_PR, sdtContent: { children: BLOCK_CONTENT } }
};

Object.assign(BLOCK_CONTENT, {
  tbl: TBL,
  sdt: SDT_BLOCK,
  customXml: { children: BLOCK_CONTENT },
  // `CT_RunTrackChange`: a revision at block level holds runs, as one in a paragraph does.
  ins: { children: RUN_CONTENT },
  del: { children: RUN_CONTENT },
  moveFrom: { children: RUN_CONTENT },
  moveTo: { children: RUN_CONTENT }
});
Object.assign(TC.children!, BLOCK_CONTENT);
Object.assign(TXBX_CONTENT.children!, BLOCK_CONTENT);
Object.assign(SHAPE_CONTAINER.children!, {
  wsp: WSP,
  grpSp: SHAPE_CONTAINER,
  wgp: SHAPE_CONTAINER,
  graphicFrame: { children: { graphic: GRAPHIC } }
});

/** A row inside a row content control or a `customXml` wrapper. */
const ROW_CONTENT: Record<string, Node> = { tr: TR };
Object.assign(ROW_CONTENT, {
  sdt: { children: { sdtPr: SDT_PR, sdtContent: { children: ROW_CONTENT } } },
  customXml: { children: ROW_CONTENT }
});
Object.assign(TBL.children!, ROW_CONTENT);

/** A cell inside a cell content control or a `customXml` wrapper. */
const CELL_CONTENT: Record<string, Node> = { tc: TC };
Object.assign(CELL_CONTENT, {
  sdt: { children: { sdtPr: SDT_PR, sdtContent: { children: CELL_CONTENT } } },
  customXml: { children: CELL_CONTENT }
});
Object.assign(TR.children!, CELL_CONTENT);

const BLOCKS_ONLY: Node = { children: BLOCK_CONTENT };

// -----------------------------------------------------------------------------
// Parts, by root element local name
// -----------------------------------------------------------------------------

const STYLE: Node = {
  type: "w:CT_Style",
  children: {
    pPr: P_PR_GENERAL,
    rPr: R_PR,
    tblPr: TBL_PR_BASE,
    trPr: TR_PR,
    tcPr: TC_PR,
    tblStylePr: {
      type: "w:CT_TblStylePr",
      children: { pPr: P_PR_GENERAL, rPr: R_PR, tblPr: TBL_PR_BASE, trPr: TR_PR, tcPr: TC_PR }
    }
  }
};

const LVL: Node = { type: "w:CT_Lvl", children: { pPr: P_PR_GENERAL, rPr: R_PR } };

const PARTS: Readonly<Record<string, Node>> = {
  document: {
    type: "w:CT_Document",
    children: { body: { type: "w:CT_Body", children: { ...BLOCK_CONTENT, sectPr: SECT_PR } } }
  },
  glossaryDocument: {
    children: {
      docParts: {
        children: { docPart: { children: { docPartBody: { children: BLOCK_CONTENT } } } }
      }
    }
  },
  hdr: BLOCKS_ONLY,
  ftr: BLOCKS_ONLY,
  footnotes: { children: { footnote: BLOCKS_ONLY } },
  endnotes: { children: { endnote: BLOCKS_ONLY } },
  comments: { children: { comment: BLOCKS_ONLY } },
  styles: {
    type: "w:CT_Styles",
    children: {
      docDefaults: {
        type: "w:CT_DocDefaults",
        children: {
          rPrDefault: { children: { rPr: R_PR } },
          pPrDefault: { children: { pPr: P_PR_GENERAL } }
        }
      },
      style: STYLE
    }
  },
  numbering: {
    type: "w:CT_Numbering",
    children: {
      abstractNum: { type: "w:CT_AbstractNum", children: { lvl: LVL } },
      num: {
        type: "w:CT_Num",
        children: { lvlOverride: { type: "w:CT_NumLvl", children: { lvl: LVL } } }
      }
    }
  },
  settings: {
    type: "w:CT_Settings",
    children: {
      footnotePr: { type: "w:CT_FtnDocProps" },
      endnotePr: { type: "w:CT_EdnDocProps" },
      compat: { type: "w:CT_Compat" }
    }
  },
  fonts: { children: { font: { type: "w:CT_Font" } } }
};

const WORDPROCESSINGML = new Set([
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main"
]);

const CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart";

/** The trees, for a test that every generated content model is placed somewhere. */
export const DOCX_ORDER_PARTS: Readonly<Record<string, Node>> = PARTS;

/**
 * Check the child order of every WordprocessingML part in `parts` (path → XML). A part that does not parse is
 * skipped — its well-formedness is another check's — and so is anything whose root is not WordprocessingML.
 */
export function checkDocxElementOrder(parts: ReadonlyMap<string, string>): DocxOrderViolation[] {
  const violations: DocxOrderViolation[] = [];
  for (const [part, xml] of parts) {
    if (!part.endsWith(".xml")) {
      continue;
    }
    let root;
    try {
      root = parseXml(xml).root;
    } catch {
      continue;
    }
    const local = root.name.slice(root.name.lastIndexOf(":") + 1);
    const namespace = namespaceOf(root.name, root.attributes);
    const report = ({ trail, child, after, type }: SequenceViolation<OoxmlSequenceType>) =>
      violations.push({
        part,
        message: `<${trail.join("/")}> child <${child}> appears after <${after}>; ${type.slice(type.indexOf(":") + 1)} requires the order ${describeSequence(OOXML_SEQUENCES, type)}`
      });
    // A chart part: its rules are shared with the workbook validator and apply by element name.
    if (local === "chartSpace" && namespace === CHART) {
      checkSequenceOrderByName(
        root,
        (name, parent) => chartContentModel(name, parent) as OoxmlSequenceType | undefined,
        SEQUENCE_RANKS,
        report
      );
      continue;
    }
    const tree = Object.hasOwn(PARTS, local) ? PARTS[local] : undefined;
    if (!tree || !WORDPROCESSINGML.has(namespace)) {
      continue;
    }
    checkSequenceOrder(root, tree, SEQUENCE_RANKS, report);
  }
  return violations;
}

function namespaceOf(name: string, attributes: Readonly<Record<string, string>>): string {
  const colon = name.indexOf(":");
  return attributes[colon < 0 ? "xmlns" : `xmlns:${name.slice(0, colon)}`] ?? "";
}
