import { Enums } from "@excel/core/enums";
import type { NamedStyleEntry } from "@excel/core/workbook-core";
import { ExcelUnsupportedFeatureError } from "@excel/errors";
import type { Alignment, Borders, Fill, Font, Protection, Style } from "@excel/types";
import { statePlainFacets, withInheritedFacets } from "@excel/utils/inherited-format";
import {
  FIRST_CUSTOM_NUM_FMT_ID,
  isBuiltInNumFmtId,
  keepNumFmtId,
  keptNumFmtId
} from "@excel/utils/kept-num-fmt";
import { numFmtCode } from "@excel/utils/num-fmt-code";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { WriterNamespaces } from "@excel/xlsx/xform/fragment-namespaces";
import { ListXform } from "@excel/xlsx/xform/list-xform";
import { StaticXform } from "@excel/xlsx/xform/static-xform";
import { BorderXform } from "@excel/xlsx/xform/style/border-xform";
import type { CellStyleModel } from "@excel/xlsx/xform/style/cell-style-xform";
import { CellStyleXform } from "@excel/xlsx/xform/style/cell-style-xform";
import { DxfEntryXform } from "@excel/xlsx/xform/style/dxf-xform";
import { FillXform } from "@excel/xlsx/xform/style/fill-xform";
import { FontXform } from "@excel/xlsx/xform/style/font-xform";
import { NumFmtXform } from "@excel/xlsx/xform/style/numfmt-xform";
import { StyleXform } from "@excel/xlsx/xform/style/style-xform";
import { excelNameKey } from "@utils/name-key";
import { isGeneralFormat } from "@utils/number-format";
import type { XmlSaxTag, XmlSink } from "@xml/types";
import { StdDocAttributes } from "@xml/writer";

// custom numfmt ids start here
const NUMFMT_BASE = FIRST_CUSTOM_NUM_FMT_ID;

/** The built-in "Normal" style's name as cell style names are compared — without regard to case, as Excel does. */
const NORMAL_STYLE_KEY = excelNameKey("Normal");

/**
 * The style-manager model. Each collection is dual-mode: in the write/manager
 * role it holds pre-rendered XML fragments (`_addFont` etc. push `toXml(...)`);
 * in the parse/read role (and the mock) the same arrays hold model objects.
 * The unions capture both; the few mode-specific access points cast locally.
 */
interface StylesModel {
  styles?: (string | StyleRef)[];
  numFmts?: (string | { id: number; formatCode: string })[];
  fonts?: (string | Partial<Font>)[];
  borders?: (string | Partial<Borders>)[];
  fills?: (string | Fill)[];
  dxfs?: DxfStyle[];
  /**
   * Named-style base xf pool (OOXML `cellStyleXfs`). Index 0 is always the
   * built-in "Normal" style. Dual-mode like `styles`: rendered XML strings in
   * the write/manager role, StyleRef objects in the parse/read role.
   */
  cellStyleXfs?: (string | StyleRef)[];
  /** Named-style name→xfId map (OOXML `cellStyles`). */
  cellStyles?: CellStyleModel[];
}

/** xf-level boolean attributes (pivotButton + apply* flags). */
interface XfFlags {
  pivotButton?: boolean;
  applyNumberFormat?: boolean;
  applyFont?: boolean;
  applyFill?: boolean;
  applyBorder?: boolean;
  applyAlignment?: boolean;
  applyProtection?: boolean;
}

/** Internal xf reference: a cell style expressed as indices into the collections. */
interface StyleRef extends XfFlags {
  numFmtId?: number;
  fontId?: number;
  fillId?: number;
  borderId?: number;
  xfId?: number;
  alignment?: Partial<Alignment>;
  protection?: Partial<Protection>;
  checkbox?: boolean;
  xfComplementIndex?: number;
}

/**
 * A built style model: the public {@link Style} plus the xf-level flags that the
 * read path preserves for round-tripping. `font`/`border`/`fill` are optional
 * here because they are filled in incrementally while building from a StyleRef.
 */
type StyleModelEntry = Partial<Style> & XfFlags;

/** A differential-formatting style: a cell style plus a registered numFmtId. */
type DxfStyle = Partial<Style> & { numFmtId?: number };

interface StyleIndex {
  style?: { [key: string]: number };
  // Dual-mode: in the write/manager role this maps formatCode→id
  // (`{ [code: string]: number }`); in the parse/read role it is rebuilt as the
  // inverse id→formatCode (`Record<string, string>`). The few mode-specific
  // access points narrow with a single local cast.
  numFmt?: { [key: string]: number } | Record<string, string>;
  numFmtNextId?: number;
  font?: { [key: string]: number };
  border?: { [key: string]: number };
  fill?: { [key: string]: number };
  model?: StyleModelEntry[];
}

/**
 * xf-level boolean attributes, in the order {@link packXfFlags} packs them into
 * the memo key's bitmask.
 */
const XF_FLAGS = [
  "pivotButton",
  "applyNumberFormat",
  "applyFont",
  "applyFill",
  "applyBorder",
  "applyAlignment",
  "applyProtection"
] as const satisfies readonly (keyof XfFlags)[];

/** Pack the xf-level booleans of a style model into a bitmask. */
function packXfFlags(model: Partial<Style>): number {
  let flags = 0;
  for (let i = 0; i < XF_FLAGS.length; i++) {
    if (model[XF_FLAGS[i]]) {
      flags |= 1 << i;
    }
  }
  return flags;
}

/**
 * The facets of {@link Style} that {@link styleModelKey} reads. Anything added to
 * `Style` has to be keyed or deliberately listed here, and the assertion below
 * makes that a compile error rather than a silent wrong styleId.
 */
type KeyedStyleFacet =
  | "numFmt"
  | "font"
  | "border"
  | "fill"
  | "alignment"
  | "protection"
  | "styleName";

/**
 * Fails to compile — naming the offenders — when a `Style` facet is not keyed.
 * The xf-level flags need no equivalent: {@link XF_FLAGS} is checked against
 * `XfFlags` by its `satisfies`, and both `packXfFlags` and `addStyleModel` walk
 * that one list, so a new flag reaches the key and the xf together.
 */
const _allStyleFacetsAreKeyed: Exclude<keyof Style, KeyedStyleFacet> extends never
  ? true
  : ["unkeyed Style facet(s)", Exclude<keyof Style, KeyedStyleFacet>] = true;
void _allStyleFacetsAreKeyed;

/** The key of an empty style model, per cell type — the string the fast path below would build, built once. */
const TYPE_KEYS: string[] = [];

/** Whether a style model has no keys at all, without allocating the key array `Object.keys` would. */
function isEmptyStyleModel(model: object): boolean {
  for (const _key in model) {
    return false;
  }
  return true;
}

/** The xf of a cell style that sets nothing: what a format with no other cell style is measured against. */
const PLAIN_STYLE_REF: StyleRef = {};

/** Whether two xfs' alignment (or protection) are the same; both absent is the same. */
function sameXfPart(a: object | undefined, b: object | undefined): boolean {
  return a === b || (a !== undefined && b !== undefined && JSON.stringify(a) === JSON.stringify(b));
}

/** Absent facet marker. Keeps a literal `null` in a key unambiguous — see below. */
const ABSENT = 0;

/** What each alignment attribute is when the attribute is absent; one set to this says nothing. */
const DEFAULT_ALIGNMENT: Readonly<Record<string, unknown>> = {
  horizontal: "general",
  vertical: "bottom",
  wrapText: false,
  shrinkToFit: false,
  justifyLastLine: false,
  indent: 0,
  relativeIndent: 0,
  textRotation: 0,
  readingOrder: "context"
};

/** `alignment` without the attributes that restate their default, or `undefined` when none is left. */
function withoutDefaultAlignment(
  alignment: Partial<Alignment> | undefined
): Partial<Alignment> | undefined {
  if (!alignment) {
    return undefined;
  }
  const kept = Object.fromEntries(
    Object.entries(alignment).filter(([key, value]) => DEFAULT_ALIGNMENT[key] !== value)
  ) as Partial<Alignment>;
  return isEmptyStyleModel(kept) ? undefined : kept;
}

/**
 * A content key identifying the xf {@link StylesXform.addStyleModel} would build
 * from `model` at cell type `type`. Two models sharing this key produce the same
 * styleId, so the second can be answered from a memo instead of re-registering.
 *
 * `type` is part of the key because it selects the default number format when the
 * model carries none (`Date` → `mm-dd-yy`) and adds the checkbox extLst.
 *
 * Returns `undefined` when the model cannot be keyed *soundly*, in which case the
 * caller registers it the long way and caches nothing. The invariant that matters
 * here is that the memo must never change what a registration does — neither the
 * bytes written nor the errors raised. Two cases:
 *
 *  - A cyclic facet, which `JSON.stringify` throws on.
 *  - A `null` in the output. `JSON.stringify` maps `NaN` and `±Infinity` to `null`
 *    just as a literal `null` does, and those are not interchangeable: the XML
 *    writer *refuses* a non-finite number (`<sz val="Infinity"/>` is not valid
 *    `xsd:int`), while `null` is simply absent. Keying on the JSON alone would let a
 *    hit on the `null` entry answer the `Infinity` call — silently producing a font
 *    with no size instead of reporting the bad input. Absent facets are keyed as
 *    {@link ABSENT}, so a `null` here can only have come from one of those. (A quoted
 *    string containing "null" trips this too; that costs a cache miss, which is free.)
 *
 * The key is order-sensitive — two facets carrying the same properties in a
 * different insertion order key differently — and does not model each xform's
 * normalisation, so `{ bold: false }` and `{}` key apart. Both cost a memo entry,
 * never correctness, because `_addStyle` still deduplicates by rendered XML.
 */
function styleModelKey(model: Partial<Style>, type: number): string | undefined {
  // An unstyled cell — almost every cell of a large sheet — has an empty style object, whose key is just its type.
  if (isEmptyStyleModel(model)) {
    return (TYPE_KEYS[type] ??= `${type}`);
  }
  const { numFmt, font, border, fill, alignment, protection, styleName } = model;
  // Truthiness, not `!== undefined`, to mirror `addStyleModel`'s own
  // `if (model.numFmt)`: a falsy numFmt (a JS caller's `null`, an empty string)
  // means the cell type picks the format, and must not be dereferenced. A kept
  // built-in id stands in for the code, marked so that it cannot equal one.
  const keptId = numFmt ? undefined : keptNumFmtId(model);
  const code = numFmt ? numFmtCode(numFmt) : keptId === undefined ? undefined : `\u0001${keptId}`;
  const flags = packXfFlags(model);
  // Fast path for a style that is a number format and nothing else — the shape
  // `Column.setNumFmt`/`Cell.setNumFmt` produce, and the one case where the array
  // allocation and JSON walk below cost more than they save (measured: 20k cells
  // with a distinct numFmt each, 224ms → 132ms). A digit-led key can never
  // collide with the `[`-led JSON below.
  if (!font && !border && !fill && !alignment && !protection && !flags && !styleName) {
    return code === undefined ? `${type}` : `${type}\u0000${code}`;
  }
  let key: string;
  try {
    key = JSON.stringify([
      type,
      code ?? ABSENT,
      font ?? ABSENT,
      border ?? ABSENT,
      fill ?? ABSENT,
      alignment ?? ABSENT,
      protection ?? ABSENT,
      flags,
      styleName ?? ABSENT
    ]);
  } catch {
    return undefined;
  }
  return key.includes("null") ? undefined : key;
}

// =============================================================================
// StylesXform is used to generate and parse the styles.xml file
// it manages the collections of fonts, number formats, alignments, etc
class StylesXform extends BaseXform {
  /**
   * A preserved `<dxf>` `<extLst>` is written under this writer's `<styleSheet>`, not the source's, so it
   * declares any prefix it borrowed from the source's root itself — see `fragment-namespaces.ts`.
   */
  get writerNamespaces(): WriterNamespaces {
    return STYLES_WRITER_NAMESPACES;
  }

  declare public map: Record<string, BaseXform>;
  declare public model: StylesModel;
  declare private index?: StyleIndex;
  /**
   * `addStyleModel` memo, keyed by style *content* — see {@link styleModelKey}.
   *
   * Content rather than object reference. When this was written every cell owned a
   * deep copy of its style, so a reference key could not hit at all: measured on a
   * 100-column × 200-row table sharing one number format, 204,010 calls and 0 hits,
   * every one re-rendering an `<xf>` to arrive at one of 8 styles.
   *
   * Cells now share one frozen snapshot of their row's or column's facets
   * (`core/style-sharing.ts`), so a reference key would sometimes hit — but only
   * for cells reached by the *same* owner, and never for the `numFmt` case above,
   * which is a primitive that no snapshot covers. A content key subsumes both and
   * is what the measurements below are against.
   *
   * Keying by content is -37% on that number-format table, and -68% when the shared
   * style also carries a font, fill and border.
   *
   * It is not free when it *misses*. On 20k cells with a distinct style each,
   * building keys that never hit costs +7%, and the retained keys add ~9 MB on
   * top of the ~18 MB the pre-existing `index.*` pools already hold for that
   * workbook. Both are bounded — Excel caps a workbook at 65,490 cell formats —
   * and they buy the two figures above, so this is deliberately neither capped
   * nor self-disabling: a size cap would keep the time cost while a hit-rate
   * heuristic can misfire on a sheet whose styles repeat only later, and both
   * failure modes are invisible. Revisit if a real workbook reports memory
   * pressure here.
   */
  declare private styleMemo?: Map<string, number>;
  /**
   * `dxfId` of each differential format already in `model.dxfs`, by object identity. A rule or column read
   * from a file holds the very object its `dxfId` resolved to, so identity is what lets it be handed its
   * original index back instead of appending a duplicate. See {@link seedDxfs}.
   */
  declare private dxfIds?: Map<DxfStyle, number>;
  declare private _hasCheckboxes?: boolean;
  declare public defaultFont?: Partial<Font>;
  /** {@link getDefaultStyleModel}, once built: `null` when the default format sets nothing. */
  declare private defaultStyleModel?: Partial<Style> | null;
  /** {@link listDanglingNumFmts}, found once the part is read. */
  declare private danglingNumFmts?: string[];
  /** The format {@link setDefaultStyle} gave `cellXfs[0]`, when it gave it one. */
  declare private writtenDefaultStyle?: Partial<Style>;
  /**
   * The xfs of the cell styles a cell format is based on, by `cellStyleXfs` index: the Normal style's at 0, when
   * {@link setDefaultStyle} gave it one, and each named style's — what {@link markOverrides} compares against.
   */
  declare private baseStyles?: Map<number, StyleRef>;
  /** The named styles {@link registerNamedStyles} registered, by `excelNameKey` of their name. */
  declare private namedStyles?: Map<string, Partial<Style>>;
  declare public parser?: BaseXform;
  /** name → cellStyleXfs index, populated by {@link registerNamedStyles}. */
  declare private _namedStyleIds?: Map<string, number>;
  static Mock: typeof StylesXform;

  constructor(initialise?: boolean) {
    super();

    this.map = {
      numFmts: new ListXform({ tag: "numFmts", count: true, childXform: new NumFmtXform() }),
      fonts: new ListXform({
        tag: "fonts",
        count: true,
        childXform: new FontXform(),
        $: { "x14ac:knownFonts": 1 }
      }),
      fills: new ListXform({ tag: "fills", count: true, childXform: new FillXform() }),
      borders: new ListXform({ tag: "borders", count: true, childXform: new BorderXform() }),
      cellStyleXfs: new ListXform({
        tag: "cellStyleXfs",
        count: true,
        childXform: new StyleXform()
      }),
      cellXfs: new ListXform({
        tag: "cellXfs",
        count: true,
        childXform: new StyleXform({ xfId: true })
      }),
      dxfs: new ListXform({
        tag: "dxfs",
        always: true,
        count: true,
        childXform: new DxfEntryXform()
      }),

      // for style manager
      numFmt: new NumFmtXform(),
      font: new FontXform(),
      fill: new FillXform(),
      border: new BorderXform(),
      style: new StyleXform({ xfId: true }),
      // cellStyleXfs xf elements carry no xfId attribute
      cellStyleXf: new StyleXform(),

      cellStyles: new ListXform({
        tag: "cellStyles",
        count: true,
        childXform: new CellStyleXform()
      }),
      tableStyles: StylesXform.STATIC_XFORMS.tableStyles,
      extLst: StylesXform.STATIC_XFORMS.extLst
    };

    if (initialise) {
      // StylesXform also acts as style manager and is used to build up styles-model during worksheet processing
      this.init();
    }
  }

  initIndex(): void {
    this.index = {
      style: {},
      numFmt: {},
      numFmtNextId: 164, // start custom format ids here
      font: {},
      border: {},
      fill: {}
    };
    // The memo answers with indices into the pools `index` tracks, so it cannot
    // outlive them. Rebuilding the pools without dropping it would hand out ids
    // from the discarded generation.
    this.styleMemo?.clear();
    this.dxfIds?.clear();
  }

  init(): void {
    // Prepare for Style Manager role
    this.model = {
      styles: [],
      numFmts: [],
      fonts: [],
      borders: [],
      fills: [],
      dxfs: [],
      cellStyleXfs: [],
      cellStyles: []
    };

    this.initIndex();

    // default (zero) border
    this._addBorder({});

    // add default (all zero) style
    this._addStyle({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0 });

    // add default fills
    this._addFill({ type: "pattern", pattern: "none" });
    this._addFill({ type: "pattern", pattern: "gray125" });

    // Named cell styles: cellStyleXfs[0] and the "Normal" cellStyle are always
    // present, matching Excel's implicit base style.
    this._namedStyleIds = new Map();
    this._addCellStyleXf({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 });
    this.model.cellStyles!.push({ name: "Normal", xfId: 0, builtinId: 0 });
    this._namedStyleIds.set(excelNameKey("Normal"), 0);

    this.styleMemo = new Map();
    this._hasCheckboxes = false;
    this.writtenDefaultStyle = undefined;
    this.baseStyles = undefined;
    this.namedStyles = undefined;
  }

  /**
   * Set the default font to use when no font is explicitly specified.
   * This preserves the original file's default font during round-trip.
   */
  setDefaultFont(font: Partial<Font> | undefined): void {
    this.defaultFont = font;
  }

  /**
   * Write `cellXfs[0]` — the format of every cell written without `s` — and `cellStyleXfs[0]`, the Normal style's, as
   * the formats a read found ({@link getDefaultStyleModel}, {@link getNormalStyleModel}), instead of General in the
   * default font. A cell with no style of its own is then written without `s`, which gives it that format, as it had
   * in the file read. Call after {@link registerNamedStyles}, which a default format may be based on, and before any
   * cell's style is added.
   */
  setDefaultStyle(style: Partial<Style> | undefined, normal: Partial<Style> | undefined): void {
    // Only the manager role (init()) holds the collections.
    if (!this.model.styles || !this.index?.style || !this.model.cellStyleXfs) {
      return;
    }
    this._ensureDefaultFont();
    // The Normal style first: the default format is based on it, and says which of its facets it applies.
    if (normal && !isEmptyStyleModel(normal)) {
      const ref = this._buildStyleRef(normal);
      this.model.cellStyleXfs[0] = this.map.cellStyleXf.toXml(ref);
      (this.baseStyles ??= new Map()).set(0, ref);
      this.styleMemo?.clear();
    }
    if (style && !isEmptyStyleModel(style)) {
      const ref: StyleRef = { ...this._buildStyleRef(style), xfId: 0 };
      if (style.styleName && this._namedStyleIds) {
        ref.xfId = this._namedStyleIds.get(excelNameKey(style.styleName)) ?? 0;
      }
      this.markOverrides(ref);
      const xml = this.map.style.toXml(ref);
      // Neither xf is left to be found by content: `0` is for a cell with no style of its own (see `addStyleModel`),
      // and a style that states the same facets is a style of its own, written as one — which is what it reads as.
      delete this.index.style[this.model.styles[0] as string];
      this.model.styles[0] = xml;
      this.writtenDefaultStyle = style;
      this.styleMemo?.clear();
    }
  }

  /**
   * Register workbook-level named cell styles into the cellStyleXfs/cellStyles
   * collections. Must be called (in the manager role) before cells are added so
   * that {@link addStyleModel} can resolve `styleName` → xfId. The built-in
   * "Normal" style is already seeded by {@link init}; entries named "Normal"
   * (in any case) are ignored.
   *
   * Names are compared as Excel compares them, without regard to case: an entry whose name matches a registered one
   * replaces it and keeps the registered spelling, so the package never carries two `<cellStyle>`s Excel would merge.
   */
  registerNamedStyles(namedStyles: Map<string, NamedStyleEntry>): void {
    // Only the manager role (init()) maintains the named-style collections.
    if (!this.model.cellStyles || !this.model.cellStyleXfs) {
      return;
    }
    if (!this._namedStyleIds) {
      this._namedStyleIds = new Map();
    }
    // Ensure the default font exists at index 0 before registering styles that
    // themselves omit a font (so they correctly inherit fontId 0).
    this._ensureDefaultFont();

    for (const style of namedStyles.values()) {
      const key = excelNameKey(style.name);
      if (key === NORMAL_STYLE_KEY) {
        continue;
      }
      const existingIndex = this._namedStyleIds.has(key)
        ? this.model.cellStyles!.findIndex(cs => excelNameKey(cs.name) === key)
        : -1;
      const name = existingIndex >= 0 ? this.model.cellStyles![existingIndex].name : style.name;
      const ref = this._buildStyleRef(style);
      const xfId = this._addCellStyleXf(ref);
      (this.baseStyles ??= new Map()).set(xfId, ref);
      (this.namedStyles ??= new Map()).set(key, style);
      const cellStyle: CellStyleModel = {
        name,
        xfId,
        ...(style.builtinId !== undefined ? { builtinId: style.builtinId } : {}),
        ...(style.iLevel !== undefined ? { iLevel: style.iLevel } : {}),
        ...(style.hidden ? { hidden: true } : {}),
        ...(style.customBuiltin ? { customBuiltin: true } : {})
      };
      // Replace an existing entry of the same name (e.g. a streaming writer
      // redefining a style) rather than emitting a duplicate <cellStyle>.
      if (existingIndex >= 0) {
        this.model.cellStyles![existingIndex] = cellStyle;
      } else {
        this.model.cellStyles!.push(cellStyle);
      }
      this._namedStyleIds.set(key, xfId);
    }
    // A memoised styleId embeds the xfId its `styleName` resolved to. Redefining
    // a name allocates a fresh cellStyleXf, so a surviving memo entry would hand
    // out an xf pointing at the superseded one.
    this.styleMemo?.clear();
  }

  /**
   * Say which facets `style` applies: each it holds differently from the cell style it is based on (`xfId`), and each a
   * read said it applied.
   *
   * [MS-OI29500] §18.8.45: in `cellXfs`, `apply*` says whether a later change to the cell style reaches the cell, and is
   * to be true wherever the format differs from its cell style. Excel shows a format's own facets either way; other
   * readers (LibreOffice among them) take the cell style's facet where the flag is absent. So a facet held as the cell
   * style holds it says nothing — the cell keeps following the style — and one held otherwise, at index 0 as much as
   * any other (General, the first font), says it applies.
   */
  private markOverrides(style: StyleRef): void {
    const base = this.baseStyles?.get(style.xfId ?? 0) ?? PLAIN_STYLE_REF;
    const mark = (flag: (typeof XF_FLAGS)[number], differs: boolean): void => {
      style[flag] = differs || style[flag] === true;
    };
    mark("applyNumberFormat", (style.numFmtId ?? 0) !== (base.numFmtId ?? 0));
    mark("applyFont", (style.fontId ?? 0) !== (base.fontId ?? 0));
    mark("applyFill", (style.fillId ?? 0) !== (base.fillId ?? 0));
    mark("applyBorder", (style.borderId ?? 0) !== (base.borderId ?? 0));
    mark("applyAlignment", !sameXfPart(style.alignment, base.alignment));
    mark("applyProtection", !sameXfPart(style.protection, base.protection));
  }

  /** Ensure fontId 0 (the default font) exists in the fonts collection. */
  private _ensureDefaultFont(): void {
    if (!this.model.fonts!.length) {
      this._addFont(
        this.defaultFont || {
          size: 11,
          color: { theme: 1 },
          name: "Calibri",
          family: 2,
          scheme: "minor"
        }
      );
    }
  }

  /**
   * Build a StyleRef (indices into the font/fill/border/numFmt collections plus
   * alignment/protection) from a style-like model. Only the facets present on
   * the model are registered; absent facets stay at index 0 so they inherit.
   */
  private _buildStyleRef(model: Partial<Style>): StyleRef {
    const ref: StyleRef = {};
    if (model.numFmt) {
      ref.numFmtId = this._addNumFmtStr(numFmtCode(model.numFmt));
    } else if (keptNumFmtId(model) !== undefined) {
      ref.numFmtId = keptNumFmtId(model);
    }
    if (model.font) {
      ref.fontId = this._addFont(model.font);
    }
    if (model.border) {
      ref.borderId = this._addBorder(model.border);
    }
    if (model.fill) {
      ref.fillId = this._addFill(model.fill);
    }
    if (model.alignment) {
      ref.alignment = model.alignment;
    }
    if (model.protection) {
      ref.protection = model.protection;
    }
    return ref;
  }

  /**
   * Render `<cellStyleXfs>`. In the manager role the collection holds rendered
   * XML strings; in the plain-JSON/read role it holds StyleRef objects. Falls
   * back to the implicit single "Normal" xf when no named styles are present.
   */
  private _renderCellStyleXfs(xmlStream: XmlSink, renderModel: StylesModel): void {
    const xfs = renderModel.cellStyleXfs;
    if (!xfs || !xfs.length) {
      this.map.cellStyleXfs.render(xmlStream, [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }]);
      return;
    }
    if (typeof xfs[0] === "string") {
      xmlStream.openNode("cellStyleXfs", { count: xfs.length });
      xfs.forEach(xfXml => {
        xmlStream.writeRaw(xfXml as string);
      });
      xmlStream.closeNode();
    } else {
      this.map.cellStyleXfs.render(xmlStream, xfs as StyleRef[]);
    }
  }

  /**
   * Render `<cellStyles>`. Falls back to the implicit single "Normal" entry
   * when no named styles are present, preserving legacy output.
   */
  private _renderCellStyles(xmlStream: XmlSink, renderModel: StylesModel): void {
    const cellStyles =
      renderModel.cellStyles && renderModel.cellStyles.length
        ? renderModel.cellStyles
        : [{ name: "Normal", xfId: 0, builtinId: 0 }];
    this.map.cellStyles.render(xmlStream, cellStyles);
  }

  render(xmlStream: XmlSink, model?: StylesModel): void {
    const renderModel = model || this.model;
    //
    //   <fonts count="2" x14ac:knownFonts="1">
    xmlStream.openXml(StdDocAttributes);

    xmlStream.openNode("styleSheet", StylesXform.STYLESHEET_ATTRIBUTES);

    if (this.index) {
      // model has been built by style manager role (contains xml)
      if (renderModel.numFmts && renderModel.numFmts.length) {
        xmlStream.openNode("numFmts", { count: renderModel.numFmts.length });
        renderModel.numFmts.forEach(numFmtXml => {
          xmlStream.writeRaw(numFmtXml as string);
        });
        xmlStream.closeNode();
      }

      if (!renderModel.fonts!.length) {
        // default (zero) font - use preserved font or fallback to Calibri
        this._addFont(
          this.defaultFont || {
            size: 11,
            color: { theme: 1 },
            name: "Calibri",
            family: 2,
            scheme: "minor"
          }
        );
      }
      xmlStream.openNode("fonts", { count: renderModel.fonts!.length, "x14ac:knownFonts": 1 });
      renderModel.fonts!.forEach(fontXml => {
        xmlStream.writeRaw(fontXml as string);
      });
      xmlStream.closeNode();

      xmlStream.openNode("fills", { count: renderModel.fills!.length });
      renderModel.fills!.forEach(fillXml => {
        xmlStream.writeRaw(fillXml as string);
      });
      xmlStream.closeNode();

      xmlStream.openNode("borders", { count: renderModel.borders!.length });
      renderModel.borders!.forEach(borderXml => {
        xmlStream.writeRaw(borderXml as string);
      });
      xmlStream.closeNode();

      this._renderCellStyleXfs(xmlStream, renderModel);

      xmlStream.openNode("cellXfs", { count: renderModel.styles!.length });
      renderModel.styles!.forEach(styleXml => {
        xmlStream.writeRaw(styleXml as string);
      });
      xmlStream.closeNode();
    } else {
      // model is plain JSON and needs to be xformed
      this.map.numFmts.render(xmlStream, renderModel.numFmts);
      this.map.fonts.render(xmlStream, renderModel.fonts);
      this.map.fills.render(xmlStream, renderModel.fills);
      this.map.borders.render(xmlStream, renderModel.borders);
      this._renderCellStyleXfs(xmlStream, renderModel);
      this.map.cellXfs.render(xmlStream, renderModel.styles);
    }

    this._renderCellStyles(xmlStream, renderModel);

    this.map.dxfs.render(xmlStream, renderModel.dxfs);

    StylesXform.STATIC_XFORMS.tableStyles.render(xmlStream);
    StylesXform.STATIC_XFORMS.extLst.render(xmlStream);

    xmlStream.closeNode();
  }

  parseOpen(node: XmlSaxTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case "styleSheet":
        this.initIndex();
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
      case "styleSheet": {
        this.model = {};
        const add = (propName: keyof StylesModel, xform: BaseXform): void => {
          const xformModel = xform.model as unknown[] | undefined;
          if (xformModel && xformModel.length) {
            (this.model as Record<string, unknown>)[propName] = xformModel;
          }
        };
        add("numFmts", this.map.numFmts);
        add("fonts", this.map.fonts);
        add("fills", this.map.fills);
        add("borders", this.map.borders);
        add("styles", this.map.cellXfs);
        add("dxfs", this.map.dxfs);
        add("cellStyleXfs", this.map.cellStyleXfs);
        add("cellStyles", this.map.cellStyles);

        // preserve the default (first) font from the original file
        const fontsModel = this.map.fonts.model as Partial<Font>[] | undefined;
        if (fontsModel && fontsModel.length > 0) {
          this.defaultFont = fontsModel[0];
        }
        this.defaultStyleModel = undefined;

        // index numFmts
        this.index = {
          model: [],
          numFmt: {}
        };
        if (this.model.numFmts) {
          const numFmtIndex = this.index.numFmt as Record<number, string>;
          (this.model.numFmts as { id: number; formatCode: string }[]).forEach(numFmt => {
            numFmtIndex[numFmt.id] = numFmt.formatCode;
          });
        }
        this.danglingNumFmts = this.findDanglingNumFmts();

        return false;
      }
      default:
        // not quite sure how we get here!
        return true;
    }
  }

  // add a cell's style model to the collection
  // each style property is processed and cross-referenced, etc.
  // the styleId is returned. Note: cellType is used when numFmt not defined
  addStyleModel(model: Partial<Style>, cellType?: number): number {
    if (!model) {
      return 0;
    }

    // if we have no default font, add it here now
    this._ensureDefaultFont();

    const type = cellType || Enums.ValueType.Number;

    // With a default format of its own, `cellXfs[0]` is what a cell with no style has — see `setDefaultStyle`. A date
    // or a checkbox still needs the format its value does. A style that sets some facets has the default's for the
    // rest: a cell given only a font keeps the default's number format, as it does in Excel.
    //
    // A cell of a named style has that style's facets for the ones it does not set, as a cell of the default format
    // has the default's. Written as an index of 0 they were the first font and General — which is what Excel shows a
    // format's index of 0 as, whatever its cell style holds ([MS-OI29500] §18.8.45).
    const defaults = this.writtenDefaultStyle;
    if (
      defaults &&
      isEmptyStyleModel(model) &&
      type !== Enums.ValueType.Date &&
      type !== Enums.ValueType.Checkbox
    ) {
      return 0;
    }

    // Have we already built an xf with this exact content for this cell type?
    // Everything below is a pure function of (model, type) — the default and the
    // named styles are fixed while cells are added — so an equal key means an
    // equal xf and the whole registration — up to four XML renders — is
    // redundant. Asked before the facets are filled in, so a hit allocates nothing.
    const memoKey = this.styleMemo ? styleModelKey(model, type) : undefined;
    if (memoKey !== undefined) {
      const memoised = this.styleMemo!.get(memoKey);
      if (memoised !== undefined) {
        return memoised;
      }
    }

    const named =
      model.styleName === undefined
        ? undefined
        : this.namedStyles?.get(excelNameKey(model.styleName));
    const inherited = named ?? defaults;
    if (inherited) {
      model = withInheritedFacets(model, inherited, type === Enums.ValueType.Date);
    }

    const style: StyleRef = {};

    if (model.numFmt) {
      style.numFmtId = this._addNumFmtStr(numFmtCode(model.numFmt));
    } else {
      // A built-in format the read could name only by its id is written back by it — see `kept-num-fmt.ts`.
      const keptId = keptNumFmtId(model);
      if (keptId !== undefined) {
        style.numFmtId = keptId;
      } else if (type === Enums.ValueType.Number) {
        style.numFmtId = this._addNumFmtStr("General");
      } else if (type === Enums.ValueType.Date) {
        style.numFmtId = this._addNumFmtStr("mm-dd-yy");
      }
    }

    if (model.font) {
      style.fontId = this._addFont(model.font);
    }

    if (model.border) {
      style.borderId = this._addBorder(model.border);
    }

    if (model.fill) {
      style.fillId = this._addFill(model.fill);
    }

    if (model.alignment) {
      style.alignment = model.alignment;
    }

    if (model.protection) {
      style.protection = model.protection;
    }

    // Preserve xf-level attributes (pivotButton, apply* flags)
    for (const flag of XF_FLAGS) {
      if (model[flag]) {
        style[flag] = true;
      }
    }

    // Named cell style reference: point this xf at the named style's base xf.
    // Facets the cell sets explicitly (above) override; unset facets stay at
    // index 0 and are inherited from the referenced cellStyleXf by Excel.
    if (model.styleName && this._namedStyleIds) {
      const xfId = this._namedStyleIds.get(excelNameKey(model.styleName));
      if (xfId !== undefined) {
        style.xfId = xfId;
      }
    }

    if (type === Enums.ValueType.Checkbox) {
      // Checkbox rendering relies on style extensions (extLst) and workbook-level parts.
      // Force applyAlignment="1" (without emitting an <alignment/> node) by providing
      // an empty alignment object when none is specified.
      this._hasCheckboxes = true;
      style.alignment = style.alignment || {};
      style.checkbox = true;
      style.xfComplementIndex = 0;
    }

    this.markOverrides(style);

    const styleId = this._addStyle(style);
    if (memoKey !== undefined) {
      this.styleMemo!.set(memoKey, styleId);
    }
    return styleId;
  }

  // given a styleId (i.e. s="n"), get the cell's style model
  // objects are shared where possible.
  getStyleModel(id: number): Style | null {
    // In the parse/read role the `styles` collection holds StyleRef objects
    // (not the rendered strings of the write role).
    // `?.`: a styles part with no `<cellXfs>` — malformed, but readable — has no list to index.
    const style = this.model.styles?.[id] as StyleRef | undefined;
    if (!style) {
      return null;
    }

    // have we built this model before?
    const existing = this.index!.model![id];
    if (existing) {
      return existing as Style;
    }

    // build a new model
    const model: StyleModelEntry = (this.index!.model![id] = {});

    // -------------------------------------------------------
    // number format
    //
    // The format's own, whatever its cell style's and whatever `applyNumberFormat` says: in `cellXfs` the flag decides
    // whether a later change to the cell style reaches the cell, not what it shows ([MS-OI29500] §18.8.45), and Excel
    // shows a format's own with it absent. The writer sets it wherever the two differ — see `markOverrides`.
    if (style.numFmtId) {
      this.setNumFmtById(model, style.numFmtId);
    }

    function addStyle<K extends "font" | "border" | "fill">(
      name: K,
      group: NonNullable<Style[K]>[],
      styleId?: number
    ): void {
      if (styleId || styleId === 0) {
        const part = group[styleId];
        if (part) {
          model[name] = part;
        }
      }
    }

    addStyle("font", this.model.fonts as NonNullable<Style["font"]>[], style.fontId);
    addStyle("border", this.model.borders as NonNullable<Style["border"]>[], style.borderId);
    addStyle("fill", this.model.fills as NonNullable<Style["fill"]>[], style.fillId);

    // -------------------------------------------------------
    // alignment
    if (style.alignment) {
      model.alignment = style.alignment;
    }

    // -------------------------------------------------------
    // protection
    if (style.protection) {
      model.protection = style.protection;
    }

    // -------------------------------------------------------
    // xf-level attributes (pivotButton, apply* flags)
    for (const flag of XF_FLAGS) {
      if (style[flag]) {
        model[flag] = true;
      }
    }

    // -------------------------------------------------------
    // named cell style reference (xfId → cellStyles[name])
    if (style.xfId) {
      const cellStyles = this.model.cellStyles;
      if (cellStyles) {
        const named = cellStyles.find(cs => cs.xfId === style.xfId);
        if (named) {
          // A file can carry two styles whose names differ only in case; Excel keeps the first and points the cells
          // of the other at it, and so does this (see `getNamedStyles`).
          const key = excelNameKey(named.name);
          model.styleName = cellStyles.find(cs => excelNameKey(cs.name) === key)!.name;
        }
      }
    }

    // A plain facet the default sets is stated — see `statePlainFacets`. The default's own facets are its own.
    if (id !== 0) {
      statePlainFacets(model, this.getDefaultStyleModel(), {
        font: (this.model.fonts as Partial<Font>[] | undefined)?.[0] ?? {},
        fill: (this.model.fills as Fill[] | undefined)?.[0] ?? { type: "pattern", pattern: "none" },
        border: (this.model.borders as Partial<Borders>[] | undefined)?.[0] ?? {}
      });
    }

    return model as Style;
  }

  /**
   * Give `holder` the number format `id` names: its code — the part's own, else the built-in one — or, for a built-in
   * with no code, the id itself, kept (`kept-num-fmt.ts`). A custom id the part does not define names nothing, and
   * sets nothing; {@link listDanglingNumFmts} reports it.
   */
  private setNumFmtById(holder: Partial<Style>, id: number): void {
    const code =
      (this.index?.numFmt as Record<string, string> | undefined)?.[id] ||
      NumFmtXform.getDefaultFmtCode(id);
    if (code) {
      holder.numFmt = code;
    } else if (isBuiltInNumFmtId(id)) {
      keepNumFmtId(holder, id);
    }
  }

  /**
   * The formats that name a custom number format the part does not define, one line each: `cellXfs[3] names number
   * format 170, …`. Such an id names nothing — it cannot be written back, since a writer numbers its custom formats
   * afresh and would point it at whichever format took the number — so its cells read as General, and the read
   * reports the loss (`readWithDiagnostics().lost`).
   */
  listDanglingNumFmts(): readonly string[] {
    return this.danglingNumFmts ?? [];
  }

  private findDanglingNumFmts(): string[] {
    const found: string[] = [];
    const scan = (collection: string, xfs: unknown[] | undefined): void => {
      xfs?.forEach((xf, index) => {
        const id = (xf as StyleRef | undefined)?.numFmtId;
        if (
          id !== undefined &&
          id >= FIRST_CUSTOM_NUM_FMT_ID &&
          !(this.index?.numFmt as Record<number, string> | undefined)?.[id]
        ) {
          found.push(
            `${collection}[${index}] names number format ${id}, which the styles part does not define: read as General`
          );
        }
      });
    };
    scan("cellXfs", this.model.styles);
    scan("cellStyleXfs", this.model.cellStyleXfs);
    return found;
  }

  /**
   * The workbook's default cell format: `cellXfs[0]`, which a cell with no `s` has (ECMA-376 §18.3.1.4, `s` defaults
   * to 0) — given as what it sets beyond a cell with no format at all, or `undefined` when it sets nothing.
   *
   * Excel writes its own as General in the default font, which sets nothing; changing the Normal style's number
   * format or alignment changes it. A cell without `s` used to read as unformatted whatever it said, and the writer's
   * own `cellXfs[0]` then replaced it, so such a workbook's unstyled cells — every cell it does not hold among them —
   * lost their format on a round trip.
   *
   * A facet at its default is left out, so a format that merely restates the defaults (LibreOffice writes General as
   * a custom `numFmtId="164"`, with every alignment attribute at its default) sets nothing; the xf-level `apply*`
   * flags say nothing about how a cell looks and are left out too.
   */
  getDefaultStyleModel(): Partial<Style> | undefined {
    if (this.defaultStyleModel === undefined) {
      const ref = this.model.styles?.[0] as StyleRef | undefined;
      const model = this.formatBeyondDefaults(ref) ?? {};
      // A default format based on a named style other than Normal is that style's.
      const styleName = ref?.xfId ? this.getStyleModel(0)?.styleName : undefined;
      if (styleName) {
        model.styleName = styleName;
      }
      this.defaultStyleModel = isEmptyStyleModel(model) ? null : model;
    }
    return this.defaultStyleModel ?? undefined;
  }

  /**
   * The Normal style's format, `cellStyleXfs[0]`, as {@link getDefaultStyleModel} gives the default cell format. In
   * a file Excel wrote the two are the same, since changing the Normal style is how its default format changes.
   */
  getNormalStyleModel(): Partial<Style> | undefined {
    return this.formatBeyondDefaults(this.model.cellStyleXfs?.[0] as StyleRef | undefined);
  }

  /** What the xf `ref` sets beyond a cell with no format — see {@link getDefaultStyleModel}. */
  private formatBeyondDefaults(ref: StyleRef | undefined): Partial<Style> | undefined {
    if (!ref || typeof ref !== "object") {
      return undefined;
    }
    const model: Partial<Style> = {};
    if (ref.numFmtId) {
      this.setNumFmtById(model, ref.numFmtId);
      if (typeof model.numFmt === "string" && isGeneralFormat(model.numFmt)) {
        delete model.numFmt;
      }
    }
    const font = ref.fontId
      ? (this.model.fonts as Partial<Font>[] | undefined)?.[ref.fontId]
      : undefined;
    if (font) {
      model.font = font;
    }
    const fill = ref.fillId ? (this.model.fills as Fill[] | undefined)?.[ref.fillId] : undefined;
    if (fill) {
      model.fill = fill;
    }
    const border = ref.borderId
      ? (this.model.borders as Partial<Borders>[] | undefined)?.[ref.borderId]
      : undefined;
    if (border) {
      model.border = border;
    }
    const alignment = withoutDefaultAlignment(ref.alignment);
    if (alignment) {
      model.alignment = alignment;
    }
    if (ref.protection) {
      model.protection = ref.protection;
    }
    return isEmptyStyleModel(model) ? undefined : model;
  }

  /**
   * Reconstruct the workbook's named cell styles from the parsed
   * `cellStyles` + `cellStyleXfs` collections (read role). The built-in
   * "Normal" style (builtinId 0) is excluded — it is always implicit.
   */
  getNamedStyles(): NamedStyleEntry[] {
    const cellStyles = this.model.cellStyles;
    const cellStyleXfs = this.model.cellStyleXfs as StyleRef[] | undefined;
    if (!cellStyles || !cellStyleXfs) {
      return [];
    }
    const result: NamedStyleEntry[] = [];
    // Names compare without regard to case, as Excel compares them: of two that differ only in case the first is
    // kept — what Excel does on opening such a file — and cells styled with the other are re-pointed at it.
    const seen = new Set<string>([NORMAL_STYLE_KEY]);
    for (const cs of cellStyles) {
      const key = excelNameKey(cs.name);
      if (cs.builtinId === 0 || seen.has(key)) {
        continue;
      }
      seen.add(key);
      const ref = cellStyleXfs[cs.xfId];
      const entry: NamedStyleEntry = { name: cs.name };
      if (cs.builtinId !== undefined) {
        entry.builtinId = cs.builtinId;
      }
      if (cs.iLevel !== undefined) {
        entry.iLevel = cs.iLevel;
      }
      if (cs.hidden) {
        entry.hidden = true;
      }
      if (cs.customBuiltin) {
        entry.customBuiltin = true;
      }
      if (ref) {
        if (ref.numFmtId) {
          this.setNumFmtById(entry, ref.numFmtId);
        }
        if (ref.fontId) {
          const font = (this.model.fonts as Partial<Font>[] | undefined)?.[ref.fontId];
          if (font) {
            entry.font = font;
          }
        }
        if (ref.borderId) {
          const border = (this.model.borders as Partial<Borders>[] | undefined)?.[ref.borderId];
          if (border) {
            entry.border = border;
          }
        }
        if (ref.fillId) {
          const fill = (this.model.fills as Fill[] | undefined)?.[ref.fillId];
          if (fill) {
            entry.fill = fill;
          }
        }
        if (ref.alignment) {
          entry.alignment = ref.alignment;
        }
        if (ref.protection) {
          entry.protection = ref.protection;
        }
      }
      result.push(entry);
    }
    return result;
  }

  /**
   * Start the `<dxfs>` table from the one the workbook was read with, at the same indices.
   *
   * `dxfId` is an index, and not every holder of one is modelled: a pivot table's `<formats>`, a
   * `<colorFilter>` and whatever else is carried through as preserved XML keep the number they were read
   * with. Rebuilding the table from only the styles the model can name left those pointing past its end —
   * or, worse, at a different format once something else was allocated there (#237). Seeding it first makes
   * every such reference valid without having to find them, and a modelled style that came from this table
   * is resolved back to its own index by {@link addDxfStyle}.
   *
   * Must run before any {@link addDxfStyle} call. A hole in the source table is kept as an empty format so
   * that the entries after it keep their index.
   */
  seedDxfs(dxfs: readonly (DxfStyle | undefined)[]): void {
    if (!this.model.dxfs || this.model.dxfs.length > 0) {
      return;
    }
    for (const style of dxfs) {
      this.appendDxf(style ?? {});
    }
  }

  addDxfStyle(style: DxfStyle): number {
    return this.dxfIds?.get(style) ?? this.appendDxf(style);
  }

  private appendDxf(style: DxfStyle): number {
    // **The written entry is a copy; the caller's style is never touched.** `numFmtId` is an allocation of
    // *this* write — it depends on which other formats the workbook registered first — so it belongs to the
    // table being written, not to the style. It used to be stamped onto the style itself, and a style object
    // is routinely shared: two workbooks written concurrently with one constant `{ numFmt: "0.0%" }` each
    // overwrote the other's id between preparing and rendering, and one of them shipped a `<dxf>` naming a
    // `numFmtId` its own `<numFmts>` never declared.
    const entry: DxfStyle = style.numFmt
      ? {
          ...style,
          numFmtId: this._addNumFmtStr(numFmtCode(style.numFmt))
        }
      : style;

    const dxfId = this.model.dxfs!.push(entry) - 1;
    this.dxfIds ??= new Map();
    if (!this.dxfIds.has(style)) {
      this.dxfIds.set(style, dxfId);
    }
    return dxfId;
  }

  getDxfStyle(id: number): DxfStyle | undefined {
    return this.model.dxfs![id];
  }

  // Check if workbook uses checkbox feature
  get hasCheckboxes(): boolean {
    return !!this._hasCheckboxes;
  }

  // =========================================================================
  // Private Interface
  _addStyle(style: StyleRef): number {
    const xml = this.map.style.toXml(style);
    let index = this.index!.style![xml];
    if (index === undefined) {
      index = this.index!.style![xml] = this.model.styles!.length;
      this.model.styles!.push(xml);
    }
    return index;
  }

  // =========================================================================
  // Cell Style Xfs (named-style base xf pool)
  _addCellStyleXf(style: StyleRef): number {
    // cellStyleXfs xf elements never carry an xfId attribute. Unlike cellXfs,
    // these are NOT de-duplicated by content: each named cell style owns its
    // own xf so that distinct style names with identical formatting keep a
    // 1:1 cellStyle→cellStyleXf mapping and never collapse onto one another.
    const xml = this.map.cellStyleXf.toXml(style);
    const index = this.model.cellStyleXfs!.length;
    this.model.cellStyleXfs!.push(xml);
    return index;
  }

  // =========================================================================
  // Number Formats
  _addNumFmtStr(formatCode: string): number {
    // check if default format
    let index = NumFmtXform.getDefaultFmtId(formatCode);
    if (index !== undefined) {
      return index;
    }

    // check if already in
    const numFmtIndex = this.index!.numFmt as Record<string, number>;
    index = numFmtIndex[formatCode];
    if (index !== undefined) {
      return index;
    }

    index = numFmtIndex[formatCode] = NUMFMT_BASE + this.model.numFmts!.length;
    const xml = this.map.numFmt.toXml({ id: index, formatCode });
    this.model.numFmts!.push(xml);
    return index!;
  }

  // =========================================================================
  // Fonts
  _addFont(font: Partial<Font>): number {
    const xml = this.map.font.toXml(font);
    let index = this.index!.font![xml];
    if (index === undefined) {
      index = this.index!.font![xml] = this.model.fonts!.length;
      this.model.fonts!.push(xml);
    }
    return index;
  }

  // =========================================================================
  // Borders
  _addBorder(border: Partial<Borders>): number {
    const xml = this.map.border.toXml(border);
    let index = this.index!.border![xml];
    if (index === undefined) {
      index = this.index!.border![xml] = this.model.borders!.length;
      this.model.borders!.push(xml);
    }
    return index;
  }

  // =========================================================================
  // Fills
  _addFill(fill: Fill): number {
    const xml = this.map.fill.toXml(fill);
    let index = this.index!.fill![xml];
    if (index === undefined) {
      index = this.index!.fill![xml] = this.model.fills!.length;
      this.model.fills!.push(xml);
    }
    return index;
  }

  // =========================================================================
  static STYLESHEET_ATTRIBUTES = {
    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "mc:Ignorable": "x14ac x16r2",
    "xmlns:x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xmlns:x16r2": "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main"
  };

  static STATIC_XFORMS = {
    dxfs: new StaticXform({ tag: "dxfs", $: { count: 0 } }),
    tableStyles: new StaticXform({
      tag: "tableStyles",
      $: {
        count: 0,
        defaultTableStyle: "TableStyleMedium2",
        defaultPivotStyle: "PivotStyleLight16"
      }
    }),
    extLst: new StaticXform({
      tag: "extLst",
      c: [
        {
          tag: "ext",
          $: {
            uri: "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}",
            "xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
          },
          c: [{ tag: "x14:slicerStyles", $: { defaultSlicerStyle: "SlicerStyleLight1" } }]
        },
        {
          tag: "ext",
          $: {
            uri: "{9260A510-F301-46a8-8635-F512D64BE5F5}",
            "xmlns:x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
          },
          c: [{ tag: "x15:timelineStyles", $: { defaultTimelineStyle: "TimeSlicerStyleLight1" } }]
        }
      ]
    })
  };
}

// the stylemanager mock acts like StyleManager except that it always returns 0 or {}
class StylesXformMock extends StylesXform {
  declare private _dateStyleId?: number;

  constructor() {
    super();

    this.model = {
      styles: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0 }],
      numFmts: [],
      fonts: [{ size: 11, color: { theme: 1 }, name: "Calibri", family: 2, scheme: "minor" }],
      borders: [{}],
      fills: [
        { type: "pattern", pattern: "none" },
        { type: "pattern", pattern: "gray125" }
      ],
      // `useStyles: false` drops *cell* formatting, not the differential formats a conditional format, a pivot
      // table's `<formats>` or a `<colorFilter>` refer to by index. Without a table here the first rule with a
      // style threw, and a read workbook's preserved references were written against an empty `<dxfs>`.
      dxfs: []
    };
  }

  // The mock renders its model as plain JSON (it has no `index`), so a differential format's number format is
  // registered as a `{ id, formatCode }` model rather than as pre-rendered XML.
  _addNumFmtStr(formatCode: string): number {
    const builtIn = NumFmtXform.getDefaultFmtId(formatCode);
    if (builtIn !== undefined) {
      return builtIn;
    }
    const numFmts = this.model.numFmts!;
    const existing = numFmts.find(
      (numFmt): numFmt is { id: number; formatCode: string } =>
        typeof numFmt !== "string" && numFmt.formatCode === formatCode
    );
    if (existing) {
      return existing.id;
    }
    const id = NUMFMT_BASE + numFmts.length;
    numFmts.push({ id, formatCode });
    return id;
  }

  // =========================================================================
  // Style Manager Interface

  // add a cell's style model to the collection
  // each style property is processed and cross-referenced, etc.
  // the styleId is returned. Note: cellType is used when numFmt not defined
  addStyleModel(model: Partial<Style>, cellType?: number): number {
    switch (cellType) {
      case Enums.ValueType.Checkbox:
        // Checkbox rendering relies on style extensions (extLst) and workbook-level parts.
        // The mock style manager intentionally does not build those structures.
        throw new ExcelUnsupportedFeatureError(
          "Checkbox cells",
          "require styles to be enabled (useStyles: true)"
        );
      case Enums.ValueType.Date:
        return this.dateStyleId;
      default:
        return 0;
    }
  }

  get hasCheckboxes(): boolean {
    return false;
  }

  get dateStyleId(): number {
    if (!this._dateStyleId) {
      const dateStyle = {
        numFmtId: NumFmtXform.getDefaultFmtId("mm-dd-yy")
      };
      this._dateStyleId = this.model.styles!.length;
      this.model.styles!.push(dateStyle);
    }
    return this._dateStyleId!;
  }

  // given a styleId (i.e. s="n"), get the cell's style model
  // objects are shared where possible.
  getStyleModel(/* id */): Style | null {
    return {} as Style;
  }
}

// Assign Mock after class declaration to avoid "used before declaration" error
StylesXform.Mock = StylesXformMock;

/** The styles writer declares its fixed set, whatever the source's root declared. */
const STYLES_WRITER_NAMESPACES: WriterNamespaces = { fixed: StylesXform.STYLESHEET_ATTRIBUTES };

export { StylesXform };
