/**
 * DOCX Writers - Math / Equations (OMML)
 *
 * Renders Office Math Markup Language elements (m:oMath).
 */

import type { XmlSink } from "@xml/types";

import { NS_M } from "../constants";
import type {
  MathBlock,
  MathContent,
  MathRun,
  MathFraction,
  MathSuperScript,
  MathSubScript,
  MathSubSuperScript,
  MathPreSubSuperScript,
  MathPhantom,
  MathGroupChar,
  MathBorderBox,
  MathRadical,
  MathDelimiter,
  MathNary,
  MathFunction,
  MathLimit,
  MathMatrix,
  MathAccent,
  MathBar,
  MathBox,
  MathEquationArray
} from "../types";

/** Render math content array. */
function renderMathContents(xml: XmlSink, contents: readonly MathContent[]): void {
  for (const content of contents) {
    renderMathContent(xml, content);
  }
}

/** Render a single math content element. */
function renderMathContent(xml: XmlSink, content: MathContent): void {
  switch (content.type) {
    case "mathRun":
      renderMathRun(xml, content);
      break;
    case "mathFraction":
      renderMathFraction(xml, content);
      break;
    case "mathSuperScript":
      renderMathSuperScript(xml, content);
      break;
    case "mathSubScript":
      renderMathSubScript(xml, content);
      break;
    case "mathSubSuperScript":
      renderMathSubSuperScript(xml, content);
      break;
    case "mathPreSubSuperScript":
      renderMathPreSubSuperScript(xml, content);
      break;
    case "mathPhantom":
      renderMathPhantom(xml, content);
      break;
    case "mathGroupChar":
      renderMathGroupChar(xml, content);
      break;
    case "mathBorderBox":
      renderMathBorderBox(xml, content);
      break;
    case "mathRadical":
      renderMathRadical(xml, content);
      break;
    case "mathDelimiter":
      renderMathDelimiter(xml, content);
      break;
    case "mathNary":
      renderMathNary(xml, content);
      break;
    case "mathFunction":
      renderMathFunction(xml, content);
      break;
    case "mathLimit":
      renderMathLimit(xml, content);
      break;
    case "mathMatrix":
      renderMathMatrix(xml, content);
      break;
    case "mathAccent":
      renderMathAccent(xml, content);
      break;
    case "mathBar":
      renderMathBar(xml, content);
      break;
    case "mathBox":
      renderMathBox(xml, content);
      break;
    case "mathEquationArray":
      renderMathEquationArray(xml, content);
      break;
  }
}

/** Render math run (m:r). */
function renderMathRun(xml: XmlSink, run: MathRun): void {
  xml.openNode("m:r");
  if (run.properties) {
    xml.openNode("m:rPr");
    if (run.properties.italic !== undefined) {
      xml.leafNode("m:sty", {
        "m:val": run.properties.bold
          ? run.properties.italic
            ? "bi"
            : "b"
          : run.properties.italic
            ? "i"
            : "p"
      });
    }
    if (run.properties.font) {
      xml.leafNode("m:scr", { "m:val": "roman" });
    }
    xml.closeNode();
  }
  xml.openNode("m:t", { "xml:space": "preserve" });
  xml.writeText(run.text);
  xml.closeNode();
  xml.closeNode();
}

/** Render fraction (m:f). */
function renderMathFraction(xml: XmlSink, frac: MathFraction): void {
  xml.openNode("m:f");
  if (frac.fractionType && frac.fractionType !== "bar") {
    xml.openNode("m:fPr");
    xml.leafNode("m:type", { "m:val": frac.fractionType });
    xml.closeNode();
  }
  xml.openNode("m:num");
  renderMathContents(xml, frac.numerator);
  xml.closeNode();
  xml.openNode("m:den");
  renderMathContents(xml, frac.denominator);
  xml.closeNode();
  xml.closeNode();
}

/** Render superscript (m:sSup). */
function renderMathSuperScript(xml: XmlSink, sup: MathSuperScript): void {
  xml.openNode("m:sSup");
  xml.openNode("m:e");
  renderMathContents(xml, sup.base);
  xml.closeNode();
  xml.openNode("m:sup");
  renderMathContents(xml, sup.superScript);
  xml.closeNode();
  xml.closeNode();
}

/** Render subscript (m:sSub). */
function renderMathSubScript(xml: XmlSink, sub: MathSubScript): void {
  xml.openNode("m:sSub");
  xml.openNode("m:e");
  renderMathContents(xml, sub.base);
  xml.closeNode();
  xml.openNode("m:sub");
  renderMathContents(xml, sub.subScript);
  xml.closeNode();
  xml.closeNode();
}

/** Render sub-superscript (m:sSubSup). */
function renderMathSubSuperScript(xml: XmlSink, ss: MathSubSuperScript): void {
  xml.openNode("m:sSubSup");
  xml.openNode("m:e");
  renderMathContents(xml, ss.base);
  xml.closeNode();
  xml.openNode("m:sub");
  renderMathContents(xml, ss.subScript);
  xml.closeNode();
  xml.openNode("m:sup");
  renderMathContents(xml, ss.superScript);
  xml.closeNode();
  xml.closeNode();
}

/** Render pre-sub-superscript (m:sPre). */
function renderMathPreSubSuperScript(xml: XmlSink, ss: MathPreSubSuperScript): void {
  xml.openNode("m:sPre");
  xml.openNode("m:sub");
  renderMathContents(xml, ss.preSubScript);
  xml.closeNode();
  xml.openNode("m:sup");
  renderMathContents(xml, ss.preSuperScript);
  xml.closeNode();
  xml.openNode("m:e");
  renderMathContents(xml, ss.base);
  xml.closeNode();
  xml.closeNode();
}

/** Render phantom (m:phant). */
function renderMathPhantom(xml: XmlSink, p: MathPhantom): void {
  xml.openNode("m:phant");
  // phant properties
  const hasProps =
    p.show !== undefined ||
    p.zeroWidth !== undefined ||
    p.zeroAscent !== undefined ||
    p.zeroDescent !== undefined ||
    p.transparent !== undefined;
  if (hasProps) {
    xml.openNode("m:phantPr");
    if (p.show) {
      xml.leafNode("m:show", { "m:val": "1" });
    }
    if (p.zeroWidth) {
      xml.leafNode("m:zeroWid", { "m:val": "1" });
    }
    if (p.zeroAscent) {
      xml.leafNode("m:zeroAsc", { "m:val": "1" });
    }
    if (p.zeroDescent) {
      xml.leafNode("m:zeroDesc", { "m:val": "1" });
    }
    if (p.transparent) {
      xml.leafNode("m:transp", { "m:val": "1" });
    }
    xml.closeNode();
  }
  xml.openNode("m:e");
  renderMathContents(xml, p.content);
  xml.closeNode();
  xml.closeNode();
}

/** Render group character (m:groupChr). */
function renderMathGroupChar(xml: XmlSink, g: MathGroupChar): void {
  xml.openNode("m:groupChr");
  if (g.char || g.position || g.verticalAlign) {
    xml.openNode("m:groupChrPr");
    if (g.char) {
      xml.leafNode("m:chr", { "m:val": g.char });
    }
    if (g.position) {
      xml.leafNode("m:pos", { "m:val": g.position });
    }
    if (g.verticalAlign) {
      xml.leafNode("m:vertJc", { "m:val": g.verticalAlign });
    }
    xml.closeNode();
  }
  xml.openNode("m:e");
  renderMathContents(xml, g.base);
  xml.closeNode();
  xml.closeNode();
}

/** Render border box (m:borderBox). */
function renderMathBorderBox(xml: XmlSink, b: MathBorderBox): void {
  xml.openNode("m:borderBox");
  const hasProps =
    b.hideTop ||
    b.hideBottom ||
    b.hideLeft ||
    b.hideRight ||
    b.strikeBlTr ||
    b.strikeTlBr ||
    b.strikeH ||
    b.strikeV;
  if (hasProps) {
    xml.openNode("m:borderBoxPr");
    if (b.hideTop) {
      xml.leafNode("m:hideTop", { "m:val": "1" });
    }
    if (b.hideBottom) {
      xml.leafNode("m:hideBot", { "m:val": "1" });
    }
    if (b.hideLeft) {
      xml.leafNode("m:hideLeft", { "m:val": "1" });
    }
    if (b.hideRight) {
      xml.leafNode("m:hideRight", { "m:val": "1" });
    }
    if (b.strikeBlTr) {
      xml.leafNode("m:strikeBLTR", { "m:val": "1" });
    }
    if (b.strikeTlBr) {
      xml.leafNode("m:strikeTLBR", { "m:val": "1" });
    }
    if (b.strikeH) {
      xml.leafNode("m:strikeH", { "m:val": "1" });
    }
    if (b.strikeV) {
      xml.leafNode("m:strikeV", { "m:val": "1" });
    }
    xml.closeNode();
  }
  xml.openNode("m:e");
  renderMathContents(xml, b.content);
  xml.closeNode();
  xml.closeNode();
}

/** Render radical (m:rad). */
function renderMathRadical(xml: XmlSink, rad: MathRadical): void {
  xml.openNode("m:rad");
  if (rad.hideDegree) {
    xml.openNode("m:radPr");
    xml.leafNode("m:degHide", { "m:val": "1" });
    xml.closeNode();
  }
  xml.openNode("m:deg");
  if (rad.degree) {
    renderMathContents(xml, rad.degree);
  }
  xml.closeNode();
  xml.openNode("m:e");
  renderMathContents(xml, rad.content);
  xml.closeNode();
  xml.closeNode();
}

/** Render delimiter (m:d). */
function renderMathDelimiter(xml: XmlSink, d: MathDelimiter): void {
  xml.openNode("m:d");
  if (d.beginChar || d.endChar || d.separatorChar) {
    xml.openNode("m:dPr");
    if (d.beginChar) {
      xml.leafNode("m:begChr", { "m:val": d.beginChar });
    }
    if (d.endChar) {
      xml.leafNode("m:endChr", { "m:val": d.endChar });
    }
    if (d.separatorChar) {
      xml.leafNode("m:sepChr", { "m:val": d.separatorChar });
    }
    xml.closeNode();
  }
  for (const group of d.content) {
    xml.openNode("m:e");
    renderMathContents(xml, group);
    xml.closeNode();
  }
  xml.closeNode();
}

/** Render n-ary operator (m:nary) — sum, product, integral, etc. */
function renderMathNary(xml: XmlSink, nary: MathNary): void {
  xml.openNode("m:nary");

  // Properties
  const hasProps = nary.char || nary.limitsLocation || nary.supHide || nary.subHide;
  if (hasProps) {
    xml.openNode("m:naryPr");
    if (nary.char) {
      xml.leafNode("m:chr", { "m:val": nary.char });
    }
    if (nary.limitsLocation) {
      xml.leafNode("m:limLoc", { "m:val": nary.limitsLocation });
    }
    if (nary.supHide) {
      xml.leafNode("m:supHide", { "m:val": "1" });
    }
    if (nary.subHide) {
      xml.leafNode("m:subHide", { "m:val": "1" });
    }
    xml.closeNode();
  }

  xml.openNode("m:sub");
  if (nary.sub) {
    renderMathContents(xml, nary.sub);
  }
  xml.closeNode();
  xml.openNode("m:sup");
  if (nary.sup) {
    renderMathContents(xml, nary.sup);
  }
  xml.closeNode();
  xml.openNode("m:e");
  renderMathContents(xml, nary.content);
  xml.closeNode();
  xml.closeNode();
}

/** Render function (m:func). */
function renderMathFunction(xml: XmlSink, func: MathFunction): void {
  xml.openNode("m:func");
  xml.openNode("m:fName");
  renderMathContents(xml, func.name);
  xml.closeNode();
  xml.openNode("m:e");
  renderMathContents(xml, func.content);
  xml.closeNode();
  xml.closeNode();
}

/** Render limit (m:limUpp or m:limLow). */
function renderMathLimit(xml: XmlSink, lim: MathLimit): void {
  const tag = lim.limitType === "upper" ? "m:limUpp" : "m:limLow";
  xml.openNode(tag);
  xml.openNode("m:e");
  renderMathContents(xml, lim.base);
  xml.closeNode();
  xml.openNode("m:lim");
  renderMathContents(xml, lim.limit);
  xml.closeNode();
  xml.closeNode();
}

/** Render matrix (m:m). */
function renderMathMatrix(xml: XmlSink, matrix: MathMatrix): void {
  xml.openNode("m:m");
  for (const row of matrix.rows) {
    xml.openNode("m:mr");
    for (const cell of row) {
      xml.openNode("m:e");
      renderMathContents(xml, cell);
      xml.closeNode();
    }
    xml.closeNode();
  }
  xml.closeNode();
}

/** Render accent (m:acc). */
function renderMathAccent(xml: XmlSink, acc: MathAccent): void {
  xml.openNode("m:acc");
  if (acc.char) {
    xml.openNode("m:accPr");
    xml.leafNode("m:chr", { "m:val": acc.char });
    xml.closeNode();
  }
  xml.openNode("m:e");
  renderMathContents(xml, acc.content);
  xml.closeNode();
  xml.closeNode();
}

/** Render bar/overline/underline (m:bar). */
function renderMathBar(xml: XmlSink, bar: MathBar): void {
  xml.openNode("m:bar");
  xml.openNode("m:barPr");
  xml.leafNode("m:pos", { "m:val": bar.position === "top" ? "top" : "bot" });
  xml.closeNode();
  xml.openNode("m:e");
  renderMathContents(xml, bar.content);
  xml.closeNode();
  xml.closeNode();
}

/** Render box (m:box). */
function renderMathBox(xml: XmlSink, box: MathBox): void {
  xml.openNode("m:box");
  xml.openNode("m:e");
  renderMathContents(xml, box.content);
  xml.closeNode();
  xml.closeNode();
}

/** Render equation array (m:eqArr). */
function renderMathEquationArray(xml: XmlSink, eqArr: MathEquationArray): void {
  xml.openNode("m:eqArr");
  for (const row of eqArr.rows) {
    xml.openNode("m:e");
    renderMathContents(xml, row);
    xml.closeNode();
  }
  xml.closeNode();
}

/**
 * Render a math block (`m:oMathPara` containing one `m:oMath`).
 *
 * The OOXML schema only allows `m:oMathPara` inside a paragraph
 * (`EG_PContent` includes `m:oMathPara`); it must not appear directly as
 * a `EG_BlockLevelElts`. We therefore wrap it in `<w:p>` so the result is
 * valid at the body level.
 *
 * Empty math blocks (`content.length === 0`) skip rendering altogether —
 * an empty `<m:oMath/>` violates the schema (it requires at least one
 * `OMathArg` child).
 */
export function renderMathBlock(xml: XmlSink, math: MathBlock): void {
  if (math.content.length === 0) {
    // Emit an empty paragraph so the surrounding body still contains a
    // valid block, but skip the math markup entirely.
    xml.openNode("w:p");
    xml.closeNode();
    return;
  }

  xml.openNode("w:p");
  xml.openNode("m:oMathPara", { "xmlns:m": NS_M });
  xml.openNode("m:oMath");
  renderMathContents(xml, math.content);
  xml.closeNode();
  xml.closeNode();
  xml.closeNode();
}
