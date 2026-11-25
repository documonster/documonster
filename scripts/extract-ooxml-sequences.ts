/**
 * Regenerate the OOXML content-model tables — the order a container's children must have — from the ECMA-376
 * schemas.
 *
 * Run: `node --import @oxc-node/core/register scripts/extract-ooxml-sequences.ts <dir>`
 *
 * `<dir>` holds the schemas of ECMA-376 5th edition (December 2016) in `Transitional/`, from Part 4
 * (`OfficeOpenXML-XMLSchema-Transitional.zip`, downloadable from
 * https://ecma-international.org/publications-and-standards/standards/ecma-376/). Transitional is the
 * conformance class this library writes, and the one whose content models it must follow.
 *
 * `<dir>/Extensions` holds Microsoft's extension schemas, which ECMA-376 does not contain but Office writes and
 * enforces as strictly (`x14:` conditional formats and sparklines, `w14:` and `w15:` content controls, `wps:`
 * shapes). Each is the one `<pre>` block of its "Appendix A: Full XML Schema" page on learn.microsoft.com, saved as:
 *   - `x14.xsd` — [MS-XLSX] §5.4, …/openspecs/office_standards/ms-xlsx/0a377581-c743-4ace-bfcd-22f359e70165
 *   - `xm.xsd`  — [MS-XLSX] §5.1, …/openspecs/office_standards/ms-xlsx/b0ddba06-ac73-4d3d-a913-8fbe4a56cf28
 *   - `w14.xsd` — [MS-DOCX] §5.1, …/openspecs/office_standards/ms-docx/9704b59f-bc49-4618-ac66-41beb82a0d7f
 *   - `w15.xsd` — [MS-DOCX] §5.2, …/openspecs/office_standards/ms-docx/d416013d-c112-44fa-8bef-7819b8898117
 *   - `wps.xsd` — [MS-ODRAWXML] §5.2,
 *     …/openspecs/office_standards/ms-odrawxml/f73ea06d-4749-4611-b20b-755df087eb21
 *
 * Most OOXML containers are an `xsd:sequence`, and Office refuses — or "repairs" — a part whose children are out of
 * order. The writers emitted their children in an order written by hand, and the validators checked it with tables
 * written by hand too, so a mistake made twice went unseen (issue #269: `<pageSetUpPr>` before `<outlinePr>`).
 * This script is the mechanical source both now read: it flattens each listed type's content model into ordered
 * positions ("slots"); the XLSX validator and the DOCX order check rank children by slot, and the Word writers emit
 * them by it (`emitInSchemaOrder`).
 *
 * Flattening, in the order the rules are applied:
 *   - a nested `xsd:sequence` is inlined and an `xsd:group` reference expanded;
 *   - the branches of an `xsd:choice` are merged position by position (the k-th of every branch shares one), so each
 *     branch keeps its own order — members of different branches at one index are left unordered among themselves;
 *   - a repeating sequence, choice or group of more than one position collapses to one position, because
 *     `a b a b` is a valid repeated `(a, b)` — lenient, never wrong;
 *   - an `xsd:all` is one position, since it has no order at all;
 *   - `xsd:any` contributes nothing, and `xsd:complexContent/xsd:extension` is its base's slots then its own.
 * An element the content model allows at two positions is left unranked rather than guessed at.
 *
 * References resolve by namespace, not by name: `CT_Marker` is one type in the spreadsheet drawing schema and
 * another in the chart schema. Output keys carry the conventional prefix of the schema's namespace for the same
 * reason (`xdr:CT_Marker`, `c:CT_Marker`).
 *
 * Each target is written beside the module that reads it, because the module layers forbid Word from importing
 * Excel. Add a type to a target and rerun to cover another.
 */
import { readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";
import process from "node:process";

import { parseXml } from "../src/modules/xml/dom.ts";
import { withoutAmbiguousNames } from "../src/modules/xml/sequence-order.ts";
import type { XmlElement } from "../src/modules/xml/types.ts";

/** Schema file → the prefix its target namespace is conventionally written with. */
const SCHEMAS = {
  sml: "x",
  "dml-spreadsheetDrawing": "xdr",
  "dml-main": "a",
  "dml-chart": "c",
  "dml-picture": "pic",
  "dml-wordprocessingDrawing": "wp",
  wml: "w",
  "shared-commonSimpleTypes": "s",
  "shared-relationshipReference": "r",
  "shared-math": "m",
  "shared-customXmlSchemaProperties": "sl",
  // Microsoft's extensions, from `Extensions/` (see the header).
  x14: "x14",
  xm: "xm",
  w14: "w14",
  w15: "w15",
  wps: "wps"
} as const;

/**
 * The schemas that are Microsoft's extensions rather than ECMA-376's, read from `<dir>/Extensions`, by the
 * specification each is published in.
 */
const EXTENSIONS: ReadonlyMap<SchemaFile, string> = new Map([
  ["x14", "MS-XLSX"],
  ["xm", "MS-XLSX"],
  ["w14", "MS-DOCX"],
  ["w15", "MS-DOCX"],
  ["wps", "MS-ODRAWXML"]
]);

type SchemaFile = keyof typeof SCHEMAS;

interface Target {
  /** Output path, relative to the repository root. */
  out: string;
  /** The module that reads it, for the header. */
  reader: string;
  types: Partial<Record<SchemaFile, readonly string[]>>;
}

/**
 * The chart content models. Both modules write charts — Word's own chart writer and the workbook's — so both
 * targets carry them, and the rules saying where each applies are shared (`@xml/ooxml-chart-order`).
 */
const CHART_TYPES = [
  "CT_ChartSpace",
  "CT_Chart",
  "CT_PlotArea",
  "CT_DLbls",
  "CT_DLbl",
  "CT_Scaling",
  "CT_CatAx",
  "CT_ValAx",
  "CT_DateAx",
  "CT_SerAx",
  "CT_DispUnits",
  "CT_DispUnitsLbl",
  "CT_Trendline",
  "CT_TrendlineLbl",
  "CT_ErrBars",
  "CT_View3D",
  "CT_Legend",
  "CT_LegendEntry",
  "CT_Title",
  "CT_AreaChart",
  "CT_Area3DChart",
  "CT_BarChart",
  "CT_Bar3DChart",
  "CT_LineChart",
  "CT_Line3DChart",
  "CT_PieChart",
  "CT_Pie3DChart",
  "CT_DoughnutChart",
  "CT_OfPieChart",
  "CT_ScatterChart",
  "CT_BubbleChart",
  "CT_RadarChart",
  "CT_StockChart",
  "CT_SurfaceChart",
  "CT_Surface3DChart",
  "CT_BarSer",
  "CT_LineSer",
  "CT_PieSer",
  "CT_AreaSer",
  "CT_ScatterSer",
  "CT_BubbleSer",
  "CT_RadarSer",
  "CT_SurfaceSer",
  "CT_DPt",
  "CT_Marker",
  "CT_UpDownBars",
  "CT_NumRef",
  "CT_StrRef",
  "CT_MultiLvlStrRef",
  "CT_NumData",
  "CT_StrData",
  "CT_MultiLvlStrData",
  "CT_Layout",
  "CT_ManualLayout",
  "CT_DTable",
  "CT_BandFmt",
  "CT_PictureOptions",
  "CT_Protection",
  "CT_PrintSettings"
] as const;

/** DrawingML shape and text properties, which a chart's `spPr` and `txPr` are. */
const DRAWING_TEXT_TYPES = [
  "CT_ShapeProperties",
  "CT_Transform2D",
  "CT_LineProperties",
  "CT_TextBody",
  "CT_TextBodyProperties",
  "CT_TextListStyle",
  "CT_TextParagraph",
  "CT_TextParagraphProperties",
  "CT_RegularTextRun",
  "CT_TextCharacterProperties"
] as const;

const TARGETS: readonly Target[] = [
  {
    out: "src/modules/excel/utils/ooxml-validator/ooxml-sequences.ts",
    reader:
      "the XLSX validator (`check-sequences.ts`, `check-chart.ts`, and the worksheet and workbook roots)",
    types: {
      sml: [
        // worksheet
        "CT_Worksheet",
        "CT_SheetPr",
        "CT_SheetViews",
        "CT_SheetView",
        "CT_Row",
        "CT_Cell",
        "CT_AutoFilter",
        "CT_SortState",
        "CT_ConditionalFormatting",
        "CT_CfRule",
        "CT_ColorScale",
        "CT_DataBar",
        "CT_DataValidation",
        "CT_HeaderFooter",
        // workbook
        "CT_Workbook",
        // styles
        "CT_Stylesheet",
        "CT_Xf",
        "CT_Border",
        "CT_PatternFill",
        "CT_Dxf",
        // shared strings
        "CT_Sst",
        "CT_Rst",
        "CT_RElt",
        // comments
        "CT_Comments",
        "CT_Comment",
        // tables
        "CT_Table",
        "CT_TableColumn",
        // external links
        "CT_ExternalBook",
        // pivot tables
        "CT_pivotTableDefinition",
        "CT_PivotField",
        "CT_PivotCacheDefinition",
        "CT_CacheField",
        // chartsheets
        "CT_Chartsheet"
      ],
      "dml-spreadsheetDrawing": [
        "CT_TwoCellAnchor",
        "CT_OneCellAnchor",
        "CT_AbsoluteAnchor",
        "CT_Marker",
        "CT_Picture",
        "CT_PictureNonVisual",
        "CT_Shape",
        "CT_ShapeNonVisual",
        "CT_GraphicalObjectFrame",
        "CT_GraphicalObjectFrameNonVisual"
      ],
      "dml-main": [
        "CT_NonVisualDrawingProps",
        "CT_NonVisualPictureProperties",
        "CT_BlipFillProperties",
        "CT_ShapeProperties",
        "CT_Transform2D",
        "CT_LineProperties",
        "CT_ShapeStyle",
        "CT_TextBody",
        "CT_TextBodyProperties",
        "CT_TextListStyle",
        "CT_TextParagraph",
        "CT_TextParagraphProperties",
        "CT_RegularTextRun",
        "CT_TextCharacterProperties"
      ],
      "dml-chart": CHART_TYPES,
      x14: [
        "CT_ConditionalFormatting",
        "CT_CfRule",
        "CT_DataBar",
        "CT_IconSet",
        "CT_Cfvo",
        "CT_ColorScale",
        "CT_DataValidation",
        "CT_SparklineGroup",
        "CT_Sparkline"
      ]
    }
  },
  {
    out: "src/modules/word/writer/ooxml-sequences.ts",
    reader:
      "the Word writers — through `emitInSchemaOrder` (`schema-order.ts`), and directly in `preserved-xml.ts` and " +
      "`chart-writer.ts` — and the DOCX order check (`order-check.ts`)",
    types: {
      wml: [
        // document and body content
        "CT_Document",
        "CT_Body",
        "CT_P",
        "CT_PPr",
        "CT_PPrBase",
        "CT_PPrGeneral",
        "CT_ParaRPr",
        "CT_ParaRPrOriginal",
        "CT_RPr",
        "CT_R",
        "CT_Ruby",
        "CT_RubyPr",
        "CT_PBdr",
        "CT_NumPr",
        // tables
        "CT_Tbl",
        "CT_TblPr",
        "CT_TblPrBase",
        "CT_TblPrEx",
        "CT_TblPrExBase",
        "CT_TblGrid",
        "CT_Row",
        "CT_TrPr",
        "CT_Tc",
        "CT_TcPr",
        "CT_TcPrInner",
        "CT_TblBorders",
        "CT_TcBorders",
        "CT_TblCellMar",
        "CT_TcMar",
        "CT_TblStylePr",
        // sections
        "CT_SectPr",
        "CT_SectPrBase",
        "CT_PageBorders",
        "CT_FtnProps",
        "CT_EdnProps",
        "CT_FtnDocProps",
        "CT_EdnDocProps",
        // numbering
        "CT_Numbering",
        "CT_AbstractNum",
        "CT_Lvl",
        "CT_Num",
        "CT_NumLvl",
        // styles
        "CT_Styles",
        "CT_DocDefaults",
        "CT_Style",
        // settings, fonts
        "CT_Settings",
        "CT_Compat",
        "CT_Font",
        // structured document tags, fields, hyperlinks
        "CT_SdtBlock",
        "CT_SdtRun",
        "CT_SdtPr",
        "CT_SimpleField"
      ],
      "dml-wordprocessingDrawing": ["CT_Inline", "CT_Anchor", "CT_WrapPath"],
      "dml-picture": ["CT_Picture", "CT_PictureNonVisual"],
      "dml-main": [
        "CT_NonVisualDrawingProps",
        "CT_NonVisualPictureProperties",
        "CT_BlipFillProperties",
        ...DRAWING_TEXT_TYPES
      ],
      "dml-chart": CHART_TYPES,
      w14: ["CT_SdtCheckbox"],
      w15: ["CT_SdtRepeatedSection"],
      wps: ["CT_WordprocessingShape"]
    }
  }
];

const XSD = "http://www.w3.org/2001/XMLSchema";

/** `text` as the lines of a block comment, none wider than the width Prettier formats the rest of the file to. */
function wrapComment(text: string, width = 120): string {
  const lines: string[] = [];
  let line = " *";
  for (const word of text.split(/\s+/)) {
    if (line !== " *" && line.length + 1 + word.length > width) {
      lines.push(line);
      line = " *";
    }
    line += ` ${word}`;
  }
  lines.push(line);
  return lines.join("\n");
}

interface SchemaDoc {
  file: SchemaFile;
  namespaces: Record<string, string>;
  types: Map<string, XmlElement>;
  groups: Map<string, XmlElement>;
}

/** A resolved definition: the node, and the document whose prefixes its own references use. */
interface Def {
  node: XmlElement;
  doc: SchemaDoc;
}

class Schemas {
  private readonly byNamespace = new Map<string, SchemaDoc>();
  readonly byFile = new Map<SchemaFile, SchemaDoc>();

  static async load(dir: string): Promise<Schemas> {
    const schemas = new Schemas();
    for (const file of Object.keys(SCHEMAS) as SchemaFile[]) {
      const folder = EXTENSIONS.has(file) ? "Extensions" : "Transitional";
      const root = parseXml(await readFile(join(dir, folder, `${file}.xsd`), "utf8")).root;
      if (root.attributes["xmlns:xsd"] !== XSD) {
        throw new Error(`${file}.xsd: expected the xsd prefix to name ${XSD}`);
      }
      const namespaces: Record<string, string> = {};
      for (const [name, value] of Object.entries(root.attributes)) {
        if (name === "xmlns") {
          namespaces[""] = value;
        } else if (name.startsWith("xmlns:")) {
          namespaces[name.slice(6)] = value;
        }
      }
      const doc: SchemaDoc = { file, namespaces, types: new Map(), groups: new Map() };
      for (const node of children(root)) {
        const name = node.attributes.name;
        if (name && isXsd(node, "complexType")) {
          doc.types.set(name, node);
        } else if (name && isXsd(node, "group")) {
          doc.groups.set(name, node);
        }
      }
      const target = root.attributes.targetNamespace;
      if (!target) {
        throw new Error(`${file}.xsd: no targetNamespace`);
      }
      schemas.byNamespace.set(target, doc);
      schemas.byFile.set(file, doc);
    }
    return schemas;
  }

  docOf(namespace: string): SchemaDoc | undefined {
    return this.byNamespace.get(namespace);
  }

  resolve(kind: "types" | "groups", qname: string, from: SchemaDoc, where: string): Def {
    const colon = qname.indexOf(":");
    const prefix = colon < 0 ? "" : qname.slice(0, colon);
    const ns = from.namespaces[prefix];
    const doc = ns === undefined ? undefined : this.byNamespace.get(ns);
    const node = doc?.[kind].get(qname.slice(colon + 1));
    if (!doc || !node) {
      throw new Error(
        `${where}: cannot resolve ${kind === "types" ? "type" : "group"} ${qname} (${ns ?? "unbound prefix"})`
      );
    }
    return { node, doc };
  }
}

function children(element: XmlElement): XmlElement[] {
  return element.children.filter((c): c is XmlElement => c.type === "element");
}

function localOf(name: string): string {
  return name.slice(name.indexOf(":") + 1);
}

function isXsd(element: XmlElement, local: string): boolean {
  return element.name === `xsd:${local}`;
}

function repeats(node: XmlElement): boolean {
  const max = node.attributes.maxOccurs;
  return max === "unbounded" || (max !== undefined && Number(max) > 1);
}

function collapseIfRepeating(node: XmlElement, slots: string[][]): string[][] {
  return repeats(node) && slots.length > 1 ? [[...new Set(slots.flat())]] : slots;
}

/**
 * The branches of an `xsd:choice` are exclusive, so they are laid side by side: position k of the choice holds
 * position k of every branch. That keeps each branch's own order — `CT_DLbls` is `delete` or a sequence of a dozen
 * display flags, and ranking that sequence as one position would accept the flags in any order.
 */
function mergeBranches(branches: string[][][]): string[][] {
  const merged: string[][] = [];
  for (const branch of branches) {
    branch.forEach((slot, k) => {
      merged[k] = [...new Set([...(merged[k] ?? []), ...slot])];
    });
  }
  return merged;
}

const IGNORED = new Set(["any", "attribute", "attributeGroup", "anyAttribute", "annotation"]);

/**
 * A child's name as the tables carry it: its local name when it is of the namespace of the type whose content model
 * it is in (`owner`), and prefixed when it is declared in another — `a:graphic` in `wp:CT_Anchor`, by reference, or
 * an element of a group or base type of another schema. Every schema here is `elementFormDefault="qualified"`, so a
 * local declaration is of its schema's namespace.
 */
function declaredName(
  schemas: Schemas,
  node: XmlElement,
  doc: SchemaDoc,
  owner: SchemaFile,
  where: string
): string {
  const { name, ref } = node.attributes;
  let file = doc.file;
  if (ref !== undefined) {
    const colon = ref.indexOf(":");
    const ns = doc.namespaces[colon < 0 ? "" : ref.slice(0, colon)];
    const declaring = ns === undefined ? undefined : schemas.docOf(ns);
    if (!declaring) {
      throw new Error(
        `${where}: element ${ref} is of a namespace not among the schemas (${ns ?? "unbound prefix"})`
      );
    }
    file = declaring.file;
  }
  const local = localOf(name ?? ref ?? "");
  return SCHEMAS[file] === SCHEMAS[owner] ? local : `${SCHEMAS[file]}:${local}`;
}

function flatten(schemas: Schemas, def: Def, owner: SchemaFile, where: string): string[][] {
  const { node, doc } = def;
  const inner = (c: XmlElement) => flatten(schemas, { node: c, doc }, owner, where);
  if (isXsd(node, "element")) {
    return [[declaredName(schemas, node, doc, owner, where)]];
  }
  if (isXsd(node, "sequence")) {
    return collapseIfRepeating(node, children(node).flatMap(inner));
  }
  if (isXsd(node, "choice")) {
    return collapseIfRepeating(node, mergeBranches(children(node).map(inner)));
  }
  if (isXsd(node, "all")) {
    const names = children(node).flatMap(c => inner(c).flat());
    return names.length > 0 ? [[...new Set(names)]] : [];
  }
  if (isXsd(node, "group")) {
    const group = schemas.resolve("groups", node.attributes.ref ?? "", doc, where);
    const slots = children(group.node).flatMap(c =>
      flatten(schemas, { node: c, doc: group.doc }, owner, where)
    );
    return collapseIfRepeating(node, slots);
  }
  if (isXsd(node, "complexContent")) {
    return children(node).flatMap(inner);
  }
  if (isXsd(node, "extension")) {
    const base = schemas.resolve("types", node.attributes.base ?? "", doc, where);
    return [...contentOf(schemas, base, owner, where), ...children(node).flatMap(inner)];
  }
  if (Array.from(IGNORED).some(local => isXsd(node, local))) {
    return [];
  }
  throw new Error(`${where}: unexpected <${node.name}>`);
}

function contentOf(schemas: Schemas, type: Def, owner: SchemaFile, where: string): string[][] {
  return children(type.node).flatMap(c =>
    flatten(schemas, { node: c, doc: type.doc }, owner, where)
  );
}

function slotsOf(schemas: Schemas, file: SchemaFile, typeName: string): string[][] {
  const doc = schemas.byFile.get(file)!;
  const node = doc.types.get(typeName);
  if (!node) {
    throw new Error(`${file}.xsd: no type ${typeName}`);
  }
  const model = children(node).find(c => !Array.from(IGNORED).some(local => isXsd(c, local)));
  if (!model || !(isXsd(model, "sequence") || isXsd(model, "complexContent"))) {
    throw new Error(
      `${file}.xsd ${typeName}: its content model is not a sequence, so it has no order to check — remove it`
    );
  }
  const slots = contentOf(schemas, { node, doc }, file, `${file}.xsd ${typeName}`);
  const ranked = withoutAmbiguousNames(slots);
  if (ranked.length < 2) {
    throw new Error(`${file}.xsd ${typeName}: one position, nothing to order — remove it`);
  }
  return ranked;
}

async function main(): Promise<void> {
  const dir = process.argv[2];
  if (!dir) {
    throw new Error("usage: node scripts/extract-ooxml-sequences.ts <dir with Transitional/>");
  }
  const schemas = await Schemas.load(dir);
  for (const target of TARGETS) {
    const lines: string[] = [];
    for (const [file, names] of Object.entries(target.types) as [SchemaFile, readonly string[]][]) {
      lines.push(`  // ${file}.xsd`);
      for (const name of names) {
        lines.push(
          `  "${SCHEMAS[file]}:${name}": ${JSON.stringify(slotsOf(schemas, file, name))},`
        );
      }
    }
    const extensions = Object.keys(target.types)
      .filter((file): file is SchemaFile => EXTENSIONS.has(file as SchemaFile))
      .map(file => `[${EXTENSIONS.get(file)}] \`${SCHEMAS[file]}:\``)
      .join(", ");
    const summary = [
      "Content models of ECMA-376 5th edition (December 2016) Transitional types",
      extensions ? `, and of Microsoft's extension types (${extensions}),` : ",",
      " whose children form an `xsd:sequence`, flattened into ordered slots: a child's position is the index of the",
      " slot naming it, and the names in one slot may appear in any order among themselves.",
      " Keys carry the conventional prefix of the type's namespace; a child declared in another namespace carries that",
      " one's (`a:graphic` in `wp:CT_Anchor`), and every other child is of its type's.",
      ` Read by ${target.reader}.`
    ].join("");
    const out = `/**
${wrapComment(summary)}
 *
 * Generated by \`scripts/extract-ooxml-sequences.ts\` — do not edit by hand; rerun it instead.
 */
export const OOXML_SEQUENCES = {
${lines.join("\n")}
} as const satisfies Record<string, readonly (readonly string[])[]>;

export type OoxmlSequenceType = keyof typeof OOXML_SEQUENCES;
`;
    const path = join(import.meta.dirname, "..", target.out);
    await writeFile(path, out);
    process.stdout.write(`wrote ${path}\n`);
  }
}

await main();
