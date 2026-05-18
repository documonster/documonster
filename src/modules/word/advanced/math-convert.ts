/**
 * MathML ↔ OMML Conversion
 *
 * Provides bidirectional conversion between MathML strings and the internal
 * OMML (Office Math Markup Language) model used by the Word module.
 *
 * @stability experimental
 *
 * Coverage: fractions, superscript, subscript, radicals, nary operators
 * (sum/integral/product), delimiters (parentheses/brackets), and matrices.
 */

import { xmlDecode, xmlEncode } from "@xml/encode";

import type {
  MathContent,
  MathDelimiter,
  MathFraction,
  MathMatrix,
  MathNary,
  MathRadical,
  MathRun,
  MathSubScript,
  MathSubSuperScript,
  MathSuperScript
} from "../types";

// =============================================================================
// OMML → MathML
// =============================================================================

/**
 * Convert OMML model content to a MathML string.
 *
 * @param content - Array of MathContent nodes from the OMML model
 * @returns A complete MathML string wrapped in `<math>` element
 *
 * @stability experimental
 */
export function ommlToMathML(content: readonly MathContent[]): string {
  const inner = content.map(node => convertNodeToMathML(node)).join("");
  return `<math xmlns="http://www.w3.org/1998/Math/MathML">${inner}</math>`;
}

function convertNodeToMathML(node: MathContent): string {
  switch (node.type) {
    case "mathRun":
      return convertMathRunToMathML(node);
    case "mathFraction":
      return convertFractionToMathML(node);
    case "mathSuperScript":
      return convertSuperScriptToMathML(node);
    case "mathSubScript":
      return convertSubScriptToMathML(node);
    case "mathSubSuperScript":
      return convertSubSuperScriptToMathML(node);
    case "mathPreSubSuperScript":
      return convertPreSubSuperScriptToMathML(node);
    case "mathRadical":
      return convertRadicalToMathML(node);
    case "mathDelimiter":
      return convertDelimiterToMathML(node);
    case "mathNary":
      return convertNaryToMathML(node);
    case "mathFunction":
      return `<mrow><mo>${childrenToMathML(node.name)}</mo><mrow>${childrenToMathML(node.content)}</mrow></mrow>`;
    case "mathLimit":
      if (node.limitType === "lower") {
        return `<munder><mrow>${childrenToMathML(node.base)}</mrow><mrow>${childrenToMathML(node.limit)}</mrow></munder>`;
      }
      return `<mover><mrow>${childrenToMathML(node.base)}</mrow><mrow>${childrenToMathML(node.limit)}</mrow></mover>`;
    case "mathMatrix":
      return convertMatrixToMathML(node);
    case "mathAccent":
      return `<mover accent="true"><mrow>${childrenToMathML(node.content)}</mrow><mo>${xmlEncode(node.char ?? "\u0302")}</mo></mover>`;
    case "mathBar":
      if (node.position === "bottom") {
        return `<munder><mrow>${childrenToMathML(node.content)}</mrow><mo>&#x0332;</mo></munder>`;
      }
      return `<mover><mrow>${childrenToMathML(node.content)}</mrow><mo>&#x00AF;</mo></mover>`;
    case "mathBox":
      return `<mrow>${childrenToMathML(node.content)}</mrow>`;
    case "mathPhantom":
      return `<mphantom><mrow>${childrenToMathML(node.content)}</mrow></mphantom>`;
    case "mathGroupChar":
      if (node.position === "top") {
        return `<mover><mrow>${childrenToMathML(node.base)}</mrow><mo>${xmlEncode(node.char ?? "\u23DE")}</mo></mover>`;
      }
      return `<munder><mrow>${childrenToMathML(node.base)}</mrow><mo>${xmlEncode(node.char ?? "\u23DF")}</mo></munder>`;
    case "mathBorderBox":
      return `<menclose notation="box"><mrow>${childrenToMathML(node.content)}</mrow></menclose>`;
    case "mathEquationArray":
      return `<mtable>${node.rows.map(row => `<mtr><mtd><mrow>${childrenToMathML(row)}</mrow></mtd></mtr>`).join("")}</mtable>`;
    default:
      return "";
  }
}

function convertMathRunToMathML(node: MathRun): string {
  const text = xmlEncode(node.text);
  // Operators and special characters
  if (isOperator(node.text)) {
    return `<mo>${text}</mo>`;
  }
  // Numbers
  if (/^\d+(\.\d+)?$/.test(node.text)) {
    return `<mn>${text}</mn>`;
  }
  // Identifiers
  if (node.properties?.italic === false) {
    return `<mi mathvariant="normal">${text}</mi>`;
  }
  return `<mi>${text}</mi>`;
}

function convertFractionToMathML(node: MathFraction): string {
  const num = childrenToMathML(node.numerator);
  const den = childrenToMathML(node.denominator);
  if (node.fractionType === "lin") {
    return `<mrow><mrow>${num}</mrow><mo>/</mo><mrow>${den}</mrow></mrow>`;
  }
  return `<mfrac><mrow>${num}</mrow><mrow>${den}</mrow></mfrac>`;
}

function convertSuperScriptToMathML(node: MathSuperScript): string {
  return `<msup><mrow>${childrenToMathML(node.base)}</mrow><mrow>${childrenToMathML(node.superScript)}</mrow></msup>`;
}

function convertSubScriptToMathML(node: MathSubScript): string {
  return `<msub><mrow>${childrenToMathML(node.base)}</mrow><mrow>${childrenToMathML(node.subScript)}</mrow></msub>`;
}

function convertSubSuperScriptToMathML(node: MathSubSuperScript): string {
  return `<msubsup><mrow>${childrenToMathML(node.base)}</mrow><mrow>${childrenToMathML(node.subScript)}</mrow><mrow>${childrenToMathML(node.superScript)}</mrow></msubsup>`;
}

function convertPreSubSuperScriptToMathML(node: MathContent): string {
  if (node.type !== "mathPreSubSuperScript") {
    return "";
  }
  return `<mmultiscripts><mrow>${childrenToMathML(node.base)}</mrow><mprescripts/><mrow>${childrenToMathML(node.preSubScript)}</mrow><mrow>${childrenToMathML(node.preSuperScript)}</mrow></mmultiscripts>`;
}

function convertRadicalToMathML(node: MathRadical): string {
  if (node.degree && node.degree.length > 0) {
    return `<mroot><mrow>${childrenToMathML(node.content)}</mrow><mrow>${childrenToMathML(node.degree)}</mrow></mroot>`;
  }
  return `<msqrt><mrow>${childrenToMathML(node.content)}</mrow></msqrt>`;
}

function convertDelimiterToMathML(node: MathDelimiter): string {
  const open = xmlEncode(node.beginChar ?? "(");
  const close = xmlEncode(node.endChar ?? ")");
  const inner = node.content
    .map(group => `<mrow>${childrenToMathML(group)}</mrow>`)
    .join(`<mo>${xmlEncode(node.separatorChar ?? "|")}</mo>`);
  return `<mrow><mo>${open}</mo>${inner}<mo>${close}</mo></mrow>`;
}

function convertNaryToMathML(node: MathNary): string {
  const operator = xmlEncode(node.char ?? "\u2211");
  const content = childrenToMathML(node.content);

  const hasSub = node.sub && node.sub.length > 0 && !node.subHide;
  const hasSup = node.sup && node.sup.length > 0 && !node.supHide;

  if (hasSub && hasSup) {
    return `<munderover><mo>${operator}</mo><mrow>${childrenToMathML(node.sub!)}</mrow><mrow>${childrenToMathML(node.sup!)}</mrow></munderover><mrow>${content}</mrow>`;
  }
  if (hasSub) {
    return `<munder><mo>${operator}</mo><mrow>${childrenToMathML(node.sub!)}</mrow></munder><mrow>${content}</mrow>`;
  }
  if (hasSup) {
    return `<mover><mo>${operator}</mo><mrow>${childrenToMathML(node.sup!)}</mrow></mover><mrow>${content}</mrow>`;
  }
  return `<mo>${operator}</mo><mrow>${content}</mrow>`;
}

function convertMatrixToMathML(node: MathMatrix): string {
  const rows = node.rows
    .map(row => {
      const cells = row.map(cell => `<mtd><mrow>${childrenToMathML(cell)}</mrow></mtd>`).join("");
      return `<mtr>${cells}</mtr>`;
    })
    .join("");
  return `<mtable>${rows}</mtable>`;
}

function childrenToMathML(nodes: readonly MathContent[]): string {
  return nodes.map(node => convertNodeToMathML(node)).join("");
}

// =============================================================================
// MathML → OMML
// =============================================================================

/**
 * Convert a MathML string to OMML model content.
 *
 * @param mathml - A MathML string (with or without the outer `<math>` wrapper)
 * @returns Array of MathContent nodes for the OMML model
 *
 * @stability experimental
 */
export function mathMLToOmml(mathml: string): MathContent[] {
  // Simple XML parser for MathML → OMML conversion
  const nodes = parseMathMLNodes(mathml);
  return nodes;
}

// =============================================================================
// Simple MathML parser
// =============================================================================

interface MMLElement {
  tag: string;
  attrs: Record<string, string>;
  children: MMLNode[];
}

type MMLNode = MMLElement | string;

function parseMathMLNodes(mathml: string): MathContent[] {
  const tree = parseMMLTree(mathml.trim());
  // If we get a root <math> element, process its children
  if (tree.length === 1 && typeof tree[0] !== "string" && tree[0].tag === "math") {
    return convertMMLChildren(tree[0].children);
  }
  return convertMMLChildren(tree);
}

function convertMMLChildren(nodes: MMLNode[]): MathContent[] {
  const result: MathContent[] = [];
  for (const node of nodes) {
    if (typeof node === "string") {
      const trimmed = node.trim();
      if (trimmed) {
        result.push({ type: "mathRun", text: trimmed });
      }
    } else {
      const converted = convertMMLElement(node);
      if (converted) {
        if (Array.isArray(converted)) {
          result.push(...converted);
        } else {
          result.push(converted);
        }
      }
    }
  }
  return result;
}

function convertMMLElement(el: MMLElement): MathContent | MathContent[] | undefined {
  switch (el.tag) {
    case "mi":
    case "mn":
    case "mtext": {
      const text = getTextContent(el);
      return { type: "mathRun", text } as MathRun;
    }
    case "mo": {
      const text = getTextContent(el);
      return { type: "mathRun", text } as MathRun;
    }
    case "mfrac": {
      const children = getElementChildren(el);
      const numerator = children.length > 0 ? convertMMLChildren(children[0]!.children) : [];
      const denominator = children.length > 1 ? convertMMLChildren(children[1]!.children) : [];
      return {
        type: "mathFraction",
        numerator,
        denominator
      } as MathFraction;
    }
    case "msup": {
      const children = getElementChildren(el);
      const base = children.length > 0 ? convertMMLChildren(children[0]!.children) : [];
      const sup = children.length > 1 ? convertMMLChildren(children[1]!.children) : [];
      return { type: "mathSuperScript", base, superScript: sup } as MathSuperScript;
    }
    case "msub": {
      const children = getElementChildren(el);
      const base = children.length > 0 ? convertMMLChildren(children[0]!.children) : [];
      const sub = children.length > 1 ? convertMMLChildren(children[1]!.children) : [];
      return { type: "mathSubScript", base, subScript: sub } as MathSubScript;
    }
    case "msubsup": {
      const children = getElementChildren(el);
      const base = children.length > 0 ? convertMMLChildren(children[0]!.children) : [];
      const sub = children.length > 1 ? convertMMLChildren(children[1]!.children) : [];
      const sup = children.length > 2 ? convertMMLChildren(children[2]!.children) : [];
      return {
        type: "mathSubSuperScript",
        base,
        subScript: sub,
        superScript: sup
      } as MathSubSuperScript;
    }
    case "msqrt": {
      const content = convertMMLChildren(el.children);
      return { type: "mathRadical", content } as MathRadical;
    }
    case "mroot": {
      const children = getElementChildren(el);
      const content = children.length > 0 ? convertMMLChildren(children[0]!.children) : [];
      const degree = children.length > 1 ? convertMMLChildren(children[1]!.children) : [];
      return { type: "mathRadical", content, degree } as MathRadical;
    }
    case "munderover": {
      // Could be an nary operator (sum/integral) or a limit
      const children = getElementChildren(el);
      if (children.length >= 3) {
        const opEl = children[0]!;
        const opText = getTextContent(opEl);
        if (isNaryOperator(opText)) {
          const sub = convertMMLChildren(children[1]!.children);
          const sup = convertMMLChildren(children[2]!.children);
          return { type: "mathNary", char: opText, sub, sup, content: [] } as MathNary;
        }
      }
      // Fallback: treat as sub-superscript
      const base = children.length > 0 ? convertMMLChildren(children[0]!.children) : [];
      const sub = children.length > 1 ? convertMMLChildren(children[1]!.children) : [];
      const sup = children.length > 2 ? convertMMLChildren(children[2]!.children) : [];
      return { type: "mathSubSuperScript", base, subScript: sub, superScript: sup };
    }
    case "munder": {
      const children = getElementChildren(el);
      const opEl = children[0];
      if (opEl) {
        const opText = getTextContent(opEl);
        if (isNaryOperator(opText) && children.length >= 2) {
          const sub = convertMMLChildren(children[1]!.children);
          return { type: "mathNary", char: opText, sub, content: [] } as MathNary;
        }
      }
      const base = children.length > 0 ? convertMMLChildren(children[0]!.children) : [];
      const sub = children.length > 1 ? convertMMLChildren(children[1]!.children) : [];
      return { type: "mathSubScript", base, subScript: sub };
    }
    case "mover": {
      const children = getElementChildren(el);
      const opEl = children[0];
      if (opEl) {
        const opText = getTextContent(opEl);
        if (isNaryOperator(opText) && children.length >= 2) {
          const sup = convertMMLChildren(children[1]!.children);
          return { type: "mathNary", char: opText, sup, content: [] } as MathNary;
        }
      }
      const base = children.length > 0 ? convertMMLChildren(children[0]!.children) : [];
      const sup = children.length > 1 ? convertMMLChildren(children[1]!.children) : [];
      return { type: "mathSuperScript", base, superScript: sup };
    }
    case "mrow": {
      return convertMMLChildren(el.children);
    }
    case "mtable": {
      return convertMTableToOmml(el);
    }
    case "mfenced": {
      // mfenced is a deprecated MathML element but still widely used
      const open = el.attrs["open"] ?? "(";
      const close = el.attrs["close"] ?? ")";
      const separators = el.attrs["separators"] ?? ",";
      const childElements = getElementChildren(el);
      const content: MathContent[][] = childElements.map(c => convertMMLChildren(c.children));
      if (content.length === 0) {
        // If no element children, use all children
        content.push(convertMMLChildren(el.children));
      }
      return {
        type: "mathDelimiter",
        beginChar: open,
        endChar: close,
        separatorChar: separators.charAt(0) || "|",
        content
      } as MathDelimiter;
    }
    case "mmultiscripts": {
      // Pre-sub-superscript
      const children = getElementChildren(el);
      const prescriptsIdx = children.findIndex(c => c.tag === "mprescripts");
      if (prescriptsIdx >= 0 && children.length > prescriptsIdx + 2) {
        const base = prescriptsIdx > 0 ? convertMMLChildren(children[0]!.children) : [];
        const preSub = convertMMLChildren(children[prescriptsIdx + 1]!.children);
        const preSup = convertMMLChildren(children[prescriptsIdx + 2]!.children);
        return {
          type: "mathPreSubSuperScript",
          base,
          preSubScript: preSub,
          preSuperScript: preSup
        };
      }
      return convertMMLChildren(el.children);
    }
    case "menclose": {
      const content = convertMMLChildren(el.children);
      return { type: "mathBorderBox", content };
    }
    case "mphantom": {
      const content = convertMMLChildren(el.children);
      return { type: "mathPhantom", content };
    }
    case "mprescripts":
      return undefined;
    default:
      // Unknown elements: process children
      return convertMMLChildren(el.children);
  }
}

function convertMTableToOmml(el: MMLElement): MathContent {
  const rows: (readonly MathContent[])[][] = [];
  for (const child of el.children) {
    if (typeof child === "string") {
      continue;
    }
    if (child.tag === "mtr") {
      const row: (readonly MathContent[])[] = [];
      for (const td of child.children) {
        if (typeof td === "string") {
          continue;
        }
        if (td.tag === "mtd") {
          row.push(convertMMLChildren(td.children));
        }
      }
      rows.push(row);
    }
  }
  return { type: "mathMatrix", rows } as MathMatrix;
}

// =============================================================================
// MathML XML Parser (minimal, no external dependencies)
// =============================================================================

function parseMMLTree(xml: string): MMLNode[] {
  const nodes: MMLNode[] = [];
  let pos = 0;

  while (pos < xml.length) {
    if (xml[pos] === "<") {
      // Check for closing tag or self-closing
      if (xml[pos + 1] === "/") {
        // Closing tag — break out to parent
        break;
      }
      if (xml[pos + 1] === "?" || xml[pos + 1] === "!") {
        // Skip processing instructions and comments. If the closing `>` is
        // missing the document is malformed; bail out instead of looping
        // forever (indexOf returning -1 used to set pos = 0, hanging the CPU).
        const end = xml.indexOf(">", pos);
        if (end === -1) {
          // Malformed input — stop parsing. (`pos` is unused after the
          // outer loop terminates, so no further assignment is needed.)
          break;
        }
        pos = end + 1;
        continue;
      }
      // Opening tag
      const tagEnd = findTagEnd(xml, pos);
      const tagContent = xml.slice(pos + 1, tagEnd);
      const selfClosing = tagContent.endsWith("/");
      const cleanTag = selfClosing ? tagContent.slice(0, -1).trim() : tagContent.trim();

      const { tag, attrs } = parseTag(cleanTag);
      pos = tagEnd + 1;

      if (selfClosing) {
        nodes.push({ tag, attrs, children: [] });
      } else {
        // Parse children until closing tag
        const children = parseMMLTreeInner(xml, pos, tag);
        nodes.push({ tag, attrs, children: children.nodes });
        pos = children.endPos;
      }
    } else {
      // Text node
      const nextTag = xml.indexOf("<", pos);
      const text = nextTag === -1 ? xml.slice(pos) : xml.slice(pos, nextTag);
      if (text.trim()) {
        nodes.push(xmlDecode(text));
      }
      pos = nextTag === -1 ? xml.length : nextTag;
    }
  }

  return nodes;
}

function parseMMLTreeInner(
  xml: string,
  startPos: number,
  parentTag: string
): { nodes: MMLNode[]; endPos: number } {
  const nodes: MMLNode[] = [];
  let pos = startPos;

  while (pos < xml.length) {
    if (xml[pos] === "<") {
      if (xml[pos + 1] === "/") {
        // Closing tag — guard against unterminated input.
        const closeEnd = xml.indexOf(">", pos);
        if (closeEnd === -1) {
          return { nodes, endPos: xml.length };
        }
        pos = closeEnd + 1;
        return { nodes, endPos: pos };
      }
      if (xml[pos + 1] === "?" || xml[pos + 1] === "!") {
        const end = xml.indexOf(">", pos);
        if (end === -1) {
          return { nodes, endPos: xml.length };
        }
        pos = end + 1;
        continue;
      }
      // Opening tag
      const tagEnd = findTagEnd(xml, pos);
      const tagContent = xml.slice(pos + 1, tagEnd);
      const selfClosing = tagContent.endsWith("/");
      const cleanTag = selfClosing ? tagContent.slice(0, -1).trim() : tagContent.trim();

      const { tag, attrs } = parseTag(cleanTag);
      pos = tagEnd + 1;

      if (selfClosing) {
        nodes.push({ tag, attrs, children: [] });
      } else {
        const children = parseMMLTreeInner(xml, pos, tag);
        nodes.push({ tag, attrs, children: children.nodes });
        pos = children.endPos;
      }
    } else {
      const nextTag = xml.indexOf("<", pos);
      const text = nextTag === -1 ? xml.slice(pos) : xml.slice(pos, nextTag);
      if (text.trim()) {
        nodes.push(xmlDecode(text));
      }
      pos = nextTag === -1 ? xml.length : nextTag;
    }
  }

  return { nodes, endPos: pos };
}

function findTagEnd(xml: string, pos: number): number {
  let inQuote: string | null = null;
  for (let i = pos + 1; i < xml.length; i++) {
    const ch = xml[i]!;
    if (inQuote) {
      if (ch === inQuote) {
        inQuote = null;
      }
    } else if (ch === '"' || ch === "'") {
      inQuote = ch;
    } else if (ch === ">") {
      return i;
    }
  }
  return xml.length - 1;
}

function parseTag(content: string): { tag: string; attrs: Record<string, string> } {
  const spaceIdx = content.indexOf(" ");
  if (spaceIdx === -1) {
    return { tag: content, attrs: {} };
  }
  const tag = content.slice(0, spaceIdx);
  const attrStr = content.slice(spaceIdx + 1);
  const attrs: Record<string, string> = {};
  const re = /(\w[\w-]*)(?::(\w[\w-]*))?="([^"]*)"/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(attrStr)) !== null) {
    const name = match[2] ? `${match[1]}:${match[2]}` : match[1]!;
    attrs[name] = match[3]!;
  }
  return { tag, attrs };
}

// =============================================================================
// Utility
// =============================================================================

function getTextContent(el: MMLElement): string {
  let text = "";
  for (const child of el.children) {
    if (typeof child === "string") {
      text += child;
    } else {
      text += getTextContent(child);
    }
  }
  return text;
}

function getElementChildren(el: MMLElement): MMLElement[] {
  return el.children.filter((c): c is MMLElement => typeof c !== "string");
}

const OPERATORS = new Set([
  "+",
  "-",
  "=",
  "<",
  ">",
  "\u2264",
  "\u2265",
  "\u00D7",
  "\u00F7",
  "\u2212",
  "\u00B1",
  "\u2260",
  "\u2248",
  "\u221E",
  "(",
  ")",
  "[",
  "]",
  "{",
  "}",
  "|",
  ",",
  ".",
  ":",
  ";",
  "!",
  "\u2211",
  "\u220F",
  "\u222B",
  "\u222C",
  "\u222D"
]);

function isOperator(text: string): boolean {
  return OPERATORS.has(text);
}

const NARY_OPERATORS = new Set([
  "\u2211", // summation
  "\u220F", // product
  "\u2210", // coproduct
  "\u222B", // integral
  "\u222C", // double integral
  "\u222D", // triple integral
  "\u222E", // contour integral
  "\u22C0", // n-ary logical and
  "\u22C1", // n-ary logical or
  "\u22C2", // n-ary intersection
  "\u22C3" // n-ary union
]);

function isNaryOperator(text: string): boolean {
  return NARY_OPERATORS.has(text.trim());
}
