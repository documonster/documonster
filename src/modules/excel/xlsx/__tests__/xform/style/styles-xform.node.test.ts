import fs from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

import { Enums } from "@excel/core/enums";
import type { Style } from "@excel/types";
import { testXformHelper, normalizeXml } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { StylesXform } from "@excel/xlsx/xform/style/styles-xform";
import { XmlWriter } from "@xml/writer";
import { describe, it, expect } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Helper function to add apply* flags to styles based on their non-zero IDs and alignments
function addApplyFlags(model: any): any {
  const result = JSON.parse(JSON.stringify(model));
  if (result.styles) {
    // Map of style property to apply flag
    const flagMappings: Array<[string, string]> = [
      ["fontId", "applyFont"],
      ["fillId", "applyFill"],
      ["borderId", "applyBorder"],
      ["numFmtId", "applyNumberFormat"],
      ["alignment", "applyAlignment"],
      ["protection", "applyProtection"]
    ];
    result.styles = result.styles.map((style: any) => {
      const newStyle = { ...style };
      for (const [prop, flag] of flagMappings) {
        if (style[prop]) {
          newStyle[flag] = true;
        }
      }
      return newStyle;
    });
  }
  return result;
}

const expectations = [
  {
    title: "Styles with fonts",
    create() {
      return new StylesXform();
    },
    preparedModel: JSON.parse(fs.readFileSync(join(__dirname, "data/styles.1.1.json")).toString()),
    xml: fs.readFileSync(join(__dirname, "data/styles.1.2.xml")).toString(),
    get parsedModel() {
      // parsedModel includes apply* flags from the XML
      const model = addApplyFlags(this.preparedModel);
      // An empty border (<border><left/><right/><top/><bottom/><diagonal/></border>)
      // parses to undefined rather than {}, because there are no edges with
      // style/color data. This is correct — cells referencing borderId 0 should
      // not get a truthy border property.
      model.borders[0] = undefined;
      // The parser now preserves the named-style collections for round-trip
      // fidelity: the implicit "Normal" cellStyle and its base cellStyleXf.
      model.cellStyleXfs = [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }];
      model.cellStyles = [{ name: "Normal", xfId: 0, builtinId: 0 }];
      return model;
    },
    tests: ["render", "renderIn", "parse"]
  }
];

describe("StylesXform", () => {
  testXformHelper(expectations);

  describe("As StyleManager", () => {
    it("Renders empty model", () => {
      const stylesXform = new StylesXform(true);
      const expectedXml = fs.readFileSync(join(__dirname, "data/styles.2.2.xml")).toString();

      const xmlStream = new XmlWriter();
      stylesXform.render(xmlStream);

      // Use normalizeXml from test-xform-helper for consistent XML comparison
      expect(normalizeXml(xmlStream.xml)).toBe(normalizeXml(expectedXml));
    });
  });

  describe("addStyleModel memoisation", () => {
    /**
     * Count the XML renders a style registration performs. Asserting on styleIds
     * alone cannot see this optimisation at all — `_addStyle` has always
     * deduplicated by rendered XML, so an id test passes with the memo deleted.
     * The renders are the work being skipped, so they are what gets asserted.
     */
    function countRenders(styles: StylesXform): Record<string, number> {
      const { map } = styles as unknown as {
        map: Record<string, { toXml(model?: unknown): string }>;
      };
      const counts: Record<string, number> = {};
      for (const name of ["style", "font", "fill", "border", "numFmt"]) {
        const xform = map[name];
        const original = xform.toXml.bind(xform);
        counts[name] = 0;
        xform.toXml = (model?: unknown) => {
          counts[name]++;
          return original(model);
        };
      }
      return counts;
    }

    it("skips every XML render for a structurally identical, reference-distinct model", () => {
      const styles = new StylesXform(true);
      const renders = countRenders(styles);
      const first = styles.addStyleModel({ numFmt: '0.00" kg"' });
      expect(renders.style).toBeGreaterThan(0);
      const after = { ...renders };

      // A separate object with the same content — what every cell in a column
      // looks like, because propagation deep-clones each facet.
      const second = styles.addStyleModel({ numFmt: '0.00" kg"' });

      expect(second).toBe(first);
      expect(renders).toEqual(after);
      expect(styles.model.styles).toHaveLength(2);
      expect(styles.model.numFmts).toHaveLength(1);
    });

    it("skips the font, fill and border renders too, not just the xf", () => {
      const styles = new StylesXform(true);
      const build = () => ({
        font: { name: "Arial", size: 12, bold: true },
        fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } },
        border: { top: { style: "thin", color: { argb: "FF000000" } } },
        alignment: { horizontal: "center", vertical: "middle" },
        protection: { locked: false }
      });
      const renders = countRenders(styles);
      const first = styles.addStyleModel(build() as Partial<Style>);
      expect(renders.font).toBeGreaterThan(0);
      expect(renders.fill).toBeGreaterThan(0);
      expect(renders.border).toBeGreaterThan(0);
      const after = { ...renders };

      const second = styles.addStyleModel(build() as Partial<Style>);

      expect(second).toBe(first);
      expect(renders).toEqual(after);
    });

    it("keys on numFmt content whether given as a string or an object", () => {
      const styles = new StylesXform(true);
      const asString = styles.addStyleModel({ numFmt: "0.000" });
      const asObject = styles.addStyleModel({ numFmt: { formatCode: "0.000" } } as Partial<Style>);
      expect(asObject).toBe(asString);
    });

    it("distinguishes models that differ only in a nested facet", () => {
      const styles = new StylesXform(true);
      const bold = styles.addStyleModel({ font: { bold: true } });
      const italic = styles.addStyleModel({ font: { italic: true } });
      expect(italic).not.toBe(bold);
    });

    it("keys on cellType, because it selects the default number format", () => {
      const styles = new StylesXform(true);
      // The same style object at two cell types: Date pulls in mm-dd-yy, Number
      // does not. A reference-keyed cache answered the second call with the
      // first call's id and silently dropped the date format.
      const shared: Partial<Style> = {};
      const asNumber = styles.addStyleModel(shared, Enums.ValueType.Number);
      const asDate = styles.addStyleModel(shared, Enums.ValueType.Date);
      expect(asDate).not.toBe(asNumber);
      expect(asDate).toBe(styles.addStyleModel({}, Enums.ValueType.Date));
    });

    it("keeps checkbox styles distinct and still flags the workbook", () => {
      const styles = new StylesXform(true);
      const plain = styles.addStyleModel({}, Enums.ValueType.String);
      const checkbox = styles.addStyleModel({}, Enums.ValueType.Checkbox);
      expect(checkbox).not.toBe(plain);
      expect(styles.hasCheckboxes).toBe(true);
      // A repeat checkbox cell is memoised but must not clear the flag.
      expect(styles.addStyleModel({}, Enums.ValueType.Checkbox)).toBe(checkbox);
      expect(styles.hasCheckboxes).toBe(true);
    });

    it("re-resolves styleName after a named style is redefined", () => {
      const styles = new StylesXform(true);
      styles.registerNamedStyles(new Map([["Accent", { name: "Accent", font: { bold: true } }]]));
      const before = styles.addStyleModel({ styleName: "Accent" });
      // Redefining the name allocates a fresh cellStyleXf; a memoised styleId
      // would keep pointing at the superseded one.
      styles.registerNamedStyles(new Map([["Accent", { name: "Accent", font: { italic: true } }]]));
      const after = styles.addStyleModel({ styleName: "Accent" });
      expect(after).not.toBe(before);
    });

    // -----------------------------------------------------------------------
    // The memo must never change the bytes. These are the inputs where keying
    // by JSON could have, and each one is a real defect the memo introduced.
    // -----------------------------------------------------------------------

    it("tolerates a falsy numFmt rather than dereferencing it", () => {
      // `addStyleModel` guards with `if (model.numFmt)`, so `null` has always
      // meant "let the cell type choose". Building a key must use the same
      // guard — reading `.formatCode` off it threw.
      for (const numFmt of [null, "", 0, false] as unknown[]) {
        const styles = new StylesXform(true);
        expect(styles.addStyleModel({ numFmt } as Partial<Style>)).toBe(0);
      }
    });

    it("does not let a cache hit skip the writer's validation", () => {
      const styles = new StylesXform(true);
      // Cast through `unknown`: these are the shapes TypeScript forbids and a
      // JavaScript caller can still hand over.
      const style = (size: unknown) => ({ font: { size } }) as unknown as Partial<Style>;

      // `size: null` is simply absent as far as the renderers go, and registers fine.
      const nulled = styles.addStyleModel(style(null));
      expect(nulled).toBe(1);

      // `size: Infinity` is a value the XML writer refuses — `<sz val="Infinity"/>`
      // is not valid `xsd:int`. JSON coerces it to the same `null`, so keying on the
      // JSON alone would have answered this from the memo and quietly produced a font
      // with no size instead of reporting the bad input.
      expect(() => styles.addStyleModel(style(Infinity))).toThrow(/no valid XML/);
      expect(() => styles.addStyleModel(style(-Infinity))).toThrow(/no valid XML/);

      // NaN is falsy, so `IntegerXform` omits it exactly as it omits `null`. Sharing
      // an id with `null` is therefore correct, and `_addStyle` establishes it.
      expect(styles.addStyleModel(style(NaN))).toBe(nulled);
    });

    it("registers a model it cannot key unambiguously, without caching it", () => {
      const styles = new StylesXform(true);
      const renders = countRenders(styles);
      // A nested `null` is indistinguishable in JSON from a value JSON coerces to
      // null, so the key is refused rather than risked. Registration still works and
      // still deduplicates through the rendered XML — it just does the work twice.
      const model = () => ({ font: { color: null } }) as unknown as Partial<Style>;
      const first = styles.addStyleModel(model());
      const before = renders.font;
      const second = styles.addStyleModel(model());
      expect(second).toBe(first);
      expect(renders.font).toBeGreaterThan(before);
    });

    it("registers a style whose facet cannot be keyed", () => {
      const styles = new StylesXform(true);
      // A cyclic facet defeats the content key; registration must still work
      // and still deduplicate through the rendered-XML index.
      const cyclic: Record<string, unknown> = { size: 11 };
      cyclic.self = cyclic;
      const first = styles.addStyleModel({ font: cyclic } as Partial<Style>);
      const second = styles.addStyleModel({ font: cyclic } as Partial<Style>);
      expect(second).toBe(first);
    });
  });
});
