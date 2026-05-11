/**
 * DOCX Reader - Math Parser
 *
 * Parses OMML (Office Math Markup Language) elements (m:* namespace) into
 * MathContent / MathBlock model objects. Each m:* element is mapped to its
 * corresponding model interface (MathFraction, MathRadical, MathDelimiter, etc).
 */

import { textContent } from "@xml/dom";
import type { XmlElement } from "@xml/types";

import { type Mutable } from "../core/internal-utils";
import type {
  MathContent,
  MathBlock,
  MathRun,
  MathFraction,
  MathRadical,
  MathDelimiter,
  MathNary,
  MathPhantom,
  MathGroupChar,
  MathBorderBox,
  MathAccent
} from "../types";
import {
  attrLocal as mathAttrVal,
  findChildLocal as findMathChild,
  findChildrenLocal as findMathChildren
} from "./parse-utils";

// =============================================================================
// Math Parser
// =============================================================================

function parseMathContent(el: XmlElement): MathContent[] {
  const result: MathContent[] = [];
  for (const child of el.children) {
    if (child.type !== "element") {
      continue;
    }
    const name = child.name.replace(/^m:/, "");
    switch (name) {
      case "r": {
        // Math run
        const tEl = findMathChild(child, "t");
        const mrPrEl = findMathChild(child, "rPr");
        const mr: Mutable<MathRun> = { type: "mathRun", text: tEl ? textContent(tEl) : "" };
        if (mrPrEl) {
          const props: { italic?: boolean; bold?: boolean; font?: string } = {};
          const sty = findMathChild(mrPrEl, "sty");
          if (sty) {
            const v = sty.attributes["m:val"] ?? sty.attributes["val"];
            if (v === "p" || v === "b") {
              props.italic = false;
            }
            if (v === "b" || v === "bi") {
              props.bold = true;
            }
          }
          if (Object.keys(props).length > 0) {
            mr.properties = props;
          }
        }
        result.push(mr);
        break;
      }
      case "f": {
        const fPrEl = findMathChild(child, "fPr");
        const num = findMathChild(child, "num");
        const den = findMathChild(child, "den");
        const frac: Mutable<MathFraction> = {
          type: "mathFraction",
          numerator: num ? parseMathContent(num) : [],
          denominator: den ? parseMathContent(den) : []
        };
        if (fPrEl) {
          const typeEl = findMathChild(fPrEl, "type");
          if (typeEl) {
            frac.fractionType = (typeEl.attributes["m:val"] ??
              typeEl.attributes["val"]) as MathFraction["fractionType"];
          }
        }
        result.push(frac);
        break;
      }
      case "sSup": {
        const base = findMathChild(child, "e");
        const sup = findMathChild(child, "sup");
        result.push({
          type: "mathSuperScript",
          base: base ? parseMathContent(base) : [],
          superScript: sup ? parseMathContent(sup) : []
        });
        break;
      }
      case "sSub": {
        const base = findMathChild(child, "e");
        const sub = findMathChild(child, "sub");
        result.push({
          type: "mathSubScript",
          base: base ? parseMathContent(base) : [],
          subScript: sub ? parseMathContent(sub) : []
        });
        break;
      }
      case "sSubSup": {
        const base = findMathChild(child, "e");
        const sub = findMathChild(child, "sub");
        const sup = findMathChild(child, "sup");
        result.push({
          type: "mathSubSuperScript",
          base: base ? parseMathContent(base) : [],
          subScript: sub ? parseMathContent(sub) : [],
          superScript: sup ? parseMathContent(sup) : []
        });
        break;
      }
      case "sPre": {
        const base = findMathChild(child, "e");
        const sub = findMathChild(child, "sub");
        const sup = findMathChild(child, "sup");
        result.push({
          type: "mathPreSubSuperScript",
          base: base ? parseMathContent(base) : [],
          preSubScript: sub ? parseMathContent(sub) : [],
          preSuperScript: sup ? parseMathContent(sup) : []
        });
        break;
      }
      case "phant": {
        const eEl = findMathChild(child, "e");
        const phantPrEl = findMathChild(child, "phantPr");
        const ph: Mutable<MathPhantom> = {
          type: "mathPhantom",
          content: eEl ? parseMathContent(eEl) : []
        };
        if (phantPrEl) {
          const boolAttr = (name: string): boolean => {
            const el = findMathChild(phantPrEl, name);
            if (!el) {
              return false;
            }
            const v = mathAttrVal(el, "val");
            return v !== "0" && v !== "false";
          };
          if (boolAttr("show")) {
            ph.show = true;
          }
          if (boolAttr("zeroWid")) {
            ph.zeroWidth = true;
          }
          if (boolAttr("zeroAsc")) {
            ph.zeroAscent = true;
          }
          if (boolAttr("zeroDesc")) {
            ph.zeroDescent = true;
          }
          if (boolAttr("transp")) {
            ph.transparent = true;
          }
        }
        result.push(ph);
        break;
      }
      case "groupChr": {
        const eEl = findMathChild(child, "e");
        const prEl = findMathChild(child, "groupChrPr");
        const g: Mutable<MathGroupChar> = {
          type: "mathGroupChar",
          base: eEl ? parseMathContent(eEl) : []
        };
        if (prEl) {
          const chrEl = findMathChild(prEl, "chr");
          if (chrEl) {
            g.char = mathAttrVal(chrEl, "val");
          }
          const posEl = findMathChild(prEl, "pos");
          if (posEl) {
            const v = mathAttrVal(posEl, "val");
            if (v === "top" || v === "bottom") {
              g.position = v;
            }
          }
          const vjcEl = findMathChild(prEl, "vertJc");
          if (vjcEl) {
            const v = mathAttrVal(vjcEl, "val");
            if (v === "top" || v === "center" || v === "bottom") {
              g.verticalAlign = v;
            }
          }
        }
        result.push(g);
        break;
      }
      case "borderBox": {
        const eEl = findMathChild(child, "e");
        const prEl = findMathChild(child, "borderBoxPr");
        const b: Mutable<MathBorderBox> = {
          type: "mathBorderBox",
          content: eEl ? parseMathContent(eEl) : []
        };
        if (prEl) {
          const boolAttr = (name: string): boolean => {
            const el = findMathChild(prEl, name);
            if (!el) {
              return false;
            }
            const v = mathAttrVal(el, "val");
            return v !== "0" && v !== "false";
          };
          if (boolAttr("hideTop")) {
            b.hideTop = true;
          }
          if (boolAttr("hideBot")) {
            b.hideBottom = true;
          }
          if (boolAttr("hideLeft")) {
            b.hideLeft = true;
          }
          if (boolAttr("hideRight")) {
            b.hideRight = true;
          }
          if (boolAttr("strikeBLTR")) {
            b.strikeBlTr = true;
          }
          if (boolAttr("strikeTLBR")) {
            b.strikeTlBr = true;
          }
          if (boolAttr("strikeH")) {
            b.strikeH = true;
          }
          if (boolAttr("strikeV")) {
            b.strikeV = true;
          }
        }
        result.push(b);
        break;
      }
      case "rad": {
        const radPrEl = findMathChild(child, "radPr");
        const deg = findMathChild(child, "deg");
        const e = findMathChild(child, "e");
        const rad: Mutable<MathRadical> = {
          type: "mathRadical",
          content: e ? parseMathContent(e) : []
        };
        if (deg) {
          rad.degree = parseMathContent(deg);
        }
        if (radPrEl) {
          const hd = findMathChild(radPrEl, "degHide");
          if (hd) {
            const v = hd.attributes["m:val"] ?? hd.attributes["val"];
            if (v === "1" || v === "on" || v === "true") {
              rad.hideDegree = true;
            }
          }
        }
        result.push(rad);
        break;
      }
      case "d": {
        const dPrEl = findMathChild(child, "dPr");
        const delim: Mutable<MathDelimiter> & { content: MathContent[][] } = {
          type: "mathDelimiter",
          content: []
        };
        if (dPrEl) {
          const bc = findMathChild(dPrEl, "begChr");
          if (bc) {
            delim.beginChar = bc.attributes["m:val"] ?? bc.attributes["val"];
          }
          const ec = findMathChild(dPrEl, "endChr");
          if (ec) {
            delim.endChar = ec.attributes["m:val"] ?? ec.attributes["val"];
          }
          const sc = findMathChild(dPrEl, "sepChr");
          if (sc) {
            delim.separatorChar = sc.attributes["m:val"] ?? sc.attributes["val"];
          }
        }
        for (const eEl of findMathChildren(child, "e")) {
          delim.content.push(parseMathContent(eEl));
        }
        result.push(delim);
        break;
      }
      case "nary": {
        const nPrEl = findMathChild(child, "naryPr");
        const sub = findMathChild(child, "sub");
        const sup = findMathChild(child, "sup");
        const e = findMathChild(child, "e");
        const nary: Mutable<MathNary> = {
          type: "mathNary",
          content: e ? parseMathContent(e) : []
        };
        if (sub) {
          nary.sub = parseMathContent(sub);
        }
        if (sup) {
          nary.sup = parseMathContent(sup);
        }
        if (nPrEl) {
          const chrEl = findMathChild(nPrEl, "chr");
          if (chrEl) {
            nary.char = chrEl.attributes["m:val"] ?? chrEl.attributes["val"];
          }
          const limLoc = findMathChild(nPrEl, "limLoc");
          if (limLoc) {
            nary.limitsLocation = (limLoc.attributes["m:val"] ??
              limLoc.attributes["val"]) as MathNary["limitsLocation"];
          }
          const sh = findMathChild(nPrEl, "supHide");
          if (sh && (sh.attributes["m:val"] ?? sh.attributes["val"]) === "1") {
            nary.supHide = true;
          }
          const sbh = findMathChild(nPrEl, "subHide");
          if (sbh && (sbh.attributes["m:val"] ?? sbh.attributes["val"]) === "1") {
            nary.subHide = true;
          }
        }
        result.push(nary);
        break;
      }
      case "func": {
        const fName = findMathChild(child, "fName");
        const e = findMathChild(child, "e");
        result.push({
          type: "mathFunction",
          name: fName ? parseMathContent(fName) : [],
          content: e ? parseMathContent(e) : []
        });
        break;
      }
      case "limLow": {
        const base = findMathChild(child, "e");
        const lim = findMathChild(child, "lim");
        result.push({
          type: "mathLimit",
          limitType: "lower",
          base: base ? parseMathContent(base) : [],
          limit: lim ? parseMathContent(lim) : []
        });
        break;
      }
      case "limUpp": {
        const base = findMathChild(child, "e");
        const lim = findMathChild(child, "lim");
        result.push({
          type: "mathLimit",
          limitType: "upper",
          base: base ? parseMathContent(base) : [],
          limit: lim ? parseMathContent(lim) : []
        });
        break;
      }
      case "m": {
        // Matrix
        const rows: MathContent[][][] = [];
        for (const mrEl of findMathChildren(child, "mr")) {
          const row: MathContent[][] = [];
          for (const eEl of findMathChildren(mrEl, "e")) {
            row.push(parseMathContent(eEl));
          }
          rows.push(row);
        }
        result.push({ type: "mathMatrix", rows });
        break;
      }
      case "acc": {
        const accPrEl = findMathChild(child, "accPr");
        const e = findMathChild(child, "e");
        const acc: Mutable<MathAccent> = {
          type: "mathAccent",
          content: e ? parseMathContent(e) : []
        };
        if (accPrEl) {
          const chr = findMathChild(accPrEl, "chr");
          if (chr) {
            acc.char = chr.attributes["m:val"] ?? chr.attributes["val"];
          }
        }
        result.push(acc);
        break;
      }
      case "bar": {
        const barPrEl = findMathChild(child, "barPr");
        const e = findMathChild(child, "e");
        let position: "top" | "bottom" = "top";
        if (barPrEl) {
          const pos = findMathChild(barPrEl, "pos");
          if (pos) {
            const v = pos.attributes["m:val"] ?? pos.attributes["val"];
            if (v === "bot") {
              position = "bottom";
            }
          }
        }
        result.push({
          type: "mathBar",
          position,
          content: e ? parseMathContent(e) : []
        });
        break;
      }
      case "box": {
        const e = findMathChild(child, "e");
        result.push({
          type: "mathBox",
          content: e ? parseMathContent(e) : []
        });
        break;
      }
      case "eqArr": {
        const rows: MathContent[][] = [];
        for (const eEl of findMathChildren(child, "e")) {
          rows.push(parseMathContent(eEl));
        }
        result.push({ type: "mathEquationArray", rows });
        break;
      }
      // Recurse into oMath elements
      case "oMath": {
        result.push(...parseMathContent(child));
        break;
      }
    }
  }
  return result;
}

function parseMathBlock(oMathParaEl: XmlElement): MathBlock {
  const content: MathContent[] = [];
  for (const child of oMathParaEl.children) {
    if (child.type === "element") {
      const n = child.name.replace(/^m:/, "");
      if (n === "oMath") {
        content.push(...parseMathContent(child));
      }
    }
  }
  return { type: "math", content };
}

export { parseMathContent, parseMathBlock };
