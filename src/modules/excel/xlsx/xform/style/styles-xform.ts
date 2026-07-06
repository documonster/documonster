import { Enums } from "@excel/core/enums";
import type { NamedStyleEntry } from "@excel/core/workbook-core";
import { ExcelNotSupportedError } from "@excel/errors";
import type { Alignment, Borders, Fill, Font, Protection, Style } from "@excel/types";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { ListXform } from "@excel/xlsx/xform/list-xform";
import { StaticXform } from "@excel/xlsx/xform/static-xform";
import { BorderXform } from "@excel/xlsx/xform/style/border-xform";
import type { CellStyleModel } from "@excel/xlsx/xform/style/cell-style-xform";
import { CellStyleXform } from "@excel/xlsx/xform/style/cell-style-xform";
import { DxfXform } from "@excel/xlsx/xform/style/dxf-xform";
import { FillXform } from "@excel/xlsx/xform/style/fill-xform";
import { FontXform } from "@excel/xlsx/xform/style/font-xform";
import { NumFmtXform } from "@excel/xlsx/xform/style/numfmt-xform";
import { StyleXform } from "@excel/xlsx/xform/style/style-xform";
import type { SaxTag, XmlSink } from "@xml/types";
import { StdDocAttributes } from "@xml/writer";

// custom numfmt ids start here
const NUMFMT_BASE = 164;

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

// =============================================================================
// StylesXform is used to generate and parse the styles.xml file
// it manages the collections of fonts, number formats, alignments, etc
class StylesXform extends BaseXform {
  declare public map: Record<string, BaseXform>;
  declare public model: StylesModel;
  declare private index?: StyleIndex;
  declare private weakMap?: WeakMap<object, number>;
  declare private _hasCheckboxes?: boolean;
  declare public defaultFont?: Partial<Font>;
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
      dxfs: new ListXform({ tag: "dxfs", always: true, count: true, childXform: new DxfXform() }),

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
    this._namedStyleIds.set("Normal", 0);

    this.weakMap = new WeakMap();
    this._hasCheckboxes = false;
  }

  /**
   * Set the default font to use when no font is explicitly specified.
   * This preserves the original file's default font during round-trip.
   */
  setDefaultFont(font: Partial<Font> | undefined): void {
    this.defaultFont = font;
  }

  /**
   * Register workbook-level named cell styles into the cellStyleXfs/cellStyles
   * collections. Must be called (in the manager role) before cells are added so
   * that {@link addStyleModel} can resolve `styleName` → xfId. The built-in
   * "Normal" style is already seeded by {@link init}; entries named "Normal"
   * are ignored.
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

    for (const [name, style] of namedStyles) {
      if (name === "Normal") {
        continue;
      }
      const ref = this._buildStyleRef(style);
      const xfId = this._addCellStyleXf(ref);
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
      const existingIndex = this._namedStyleIds.has(name)
        ? this.model.cellStyles!.findIndex(cs => cs.name === name)
        : -1;
      if (existingIndex >= 0) {
        this.model.cellStyles![existingIndex] = cellStyle;
      } else {
        this.model.cellStyles!.push(cellStyle);
      }
      this._namedStyleIds.set(name, xfId);
    }
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
      ref.numFmtId = this._addNumFmtStr(
        typeof model.numFmt === "string" ? model.numFmt : model.numFmt.formatCode
      );
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

  parseOpen(node: SaxTag): boolean {
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

    // If we have seen this style object before, assume it has the same styleId.
    // Do not cache by object identity for checkbox cells because the styleId must
    // include checkbox-specific extLst, and the same style object may be reused
    // for non-checkbox cells.
    if (type !== Enums.ValueType.Checkbox && this.weakMap && this.weakMap.has(model)) {
      return this.weakMap.get(model)!;
    }

    const style: StyleRef = {};

    if (model.numFmt) {
      style.numFmtId = this._addNumFmtStr(
        typeof model.numFmt === "string" ? model.numFmt : model.numFmt.formatCode
      );
    } else {
      switch (type) {
        case Enums.ValueType.Number:
          style.numFmtId = this._addNumFmtStr("General");
          break;
        case Enums.ValueType.Date:
          style.numFmtId = this._addNumFmtStr("mm-dd-yy");
          break;
        default:
          break;
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
    const xfFlags = [
      "pivotButton",
      "applyNumberFormat",
      "applyFont",
      "applyFill",
      "applyBorder",
      "applyAlignment",
      "applyProtection"
    ] as const;
    for (const flag of xfFlags) {
      if (model[flag]) {
        style[flag] = true;
      }
    }

    // Named cell style reference: point this xf at the named style's base xf.
    // Facets the cell sets explicitly (above) override; unset facets stay at
    // index 0 and are inherited from the referenced cellStyleXf by Excel.
    if (model.styleName && this._namedStyleIds && this._namedStyleIds.has(model.styleName)) {
      style.xfId = this._namedStyleIds.get(model.styleName);
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

    const styleId = this._addStyle(style);
    if (type !== Enums.ValueType.Checkbox && this.weakMap) {
      this.weakMap.set(model, styleId);
    }
    return styleId;
  }

  // given a styleId (i.e. s="n"), get the cell's style model
  // objects are shared where possible.
  getStyleModel(id: number): Style | null {
    // In the parse/read role the `styles` collection holds StyleRef objects
    // (not the rendered strings of the write role).
    const style = this.model.styles![id] as StyleRef | undefined;
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
    if (style.numFmtId) {
      const numFmt =
        (this.index!.numFmt as Record<string, string>)[style.numFmtId] ||
        NumFmtXform.getDefaultFmtCode(style.numFmtId);
      if (numFmt) {
        model.numFmt = numFmt;
      }
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
    const xfFlags: (keyof XfFlags)[] = [
      "pivotButton",
      "applyNumberFormat",
      "applyFont",
      "applyFill",
      "applyBorder",
      "applyAlignment",
      "applyProtection"
    ];
    for (const flag of xfFlags) {
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
          model.styleName = named.name;
        }
      }
    }

    return model as Style;
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
    for (const cs of cellStyles) {
      if (cs.builtinId === 0 || cs.name === "Normal") {
        continue;
      }
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
          const numFmt =
            (this.index?.numFmt as Record<string, string> | undefined)?.[ref.numFmtId] ||
            NumFmtXform.getDefaultFmtCode(ref.numFmtId);
          if (numFmt) {
            entry.numFmt = numFmt;
          }
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

  addDxfStyle(style: DxfStyle): number {
    if (style.numFmt) {
      // register numFmtId to use it during dxf-xform rendering
      style.numFmtId = this._addNumFmtStr(
        typeof style.numFmt === "string" ? style.numFmt : style.numFmt.formatCode
      );
    }

    this.model.dxfs!.push(style);
    return this.model.dxfs!.length - 1;
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
      ]
    };
  }

  // =========================================================================
  // Style Manager Interface

  // override normal behaviour - consume and dispose
  parseStream(stream: { autodrain(): void }): Promise<void> {
    stream.autodrain();
    return Promise.resolve();
  }

  // add a cell's style model to the collection
  // each style property is processed and cross-referenced, etc.
  // the styleId is returned. Note: cellType is used when numFmt not defined
  addStyleModel(model: Partial<Style>, cellType?: number): number {
    switch (cellType) {
      case Enums.ValueType.Checkbox:
        // Checkbox rendering relies on style extensions (extLst) and workbook-level parts.
        // The mock style manager intentionally does not build those structures.
        throw new ExcelNotSupportedError(
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

export { StylesXform };
