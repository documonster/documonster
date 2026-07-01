import { CompyXform } from "@excel/xlsx/__tests__/xform/compy-xform";
import { BooleanXform } from "@excel/xlsx/xform/simple/boolean-xform";
import { PassThrough } from "@stream";
import { parseSax } from "@xml/sax";

/**
 * Local test helper: deep clone (replaces the retired under-dash `cloneDeep`).
 * `preserveUndefined=false` drops `undefined` props/holes — used to normalise
 * models before structural comparison.
 */
function cloneDeep(obj: unknown, preserveUndefined = true): unknown {
  if (obj === null || typeof obj !== "object") {
    return obj;
  }
  if (obj instanceof Date) {
    return obj;
  }
  if (Array.isArray(obj)) {
    const clone: unknown[] = new Array(obj.length);
    for (let i = 0; i < obj.length; i++) {
      const value = obj[i];
      if (value !== undefined) {
        clone[i] = cloneDeep(value, preserveUndefined);
      } else if (preserveUndefined) {
        clone[i] = undefined;
      }
    }
    return clone;
  }
  const clone: Record<string, unknown> = {};
  for (const key of Object.keys(obj)) {
    const value = (obj as Record<string, unknown>)[key];
    if (value !== undefined) {
      clone[key] = cloneDeep(value, preserveUndefined);
    } else if (preserveUndefined) {
      clone[key] = undefined;
    }
  }
  return clone;
}
import { XmlWriter } from "@xml/writer";
import { expect } from "vitest";

/**
 * Decode XML entities
 */
function decodeXmlEntities(str: string): string {
  return str
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

/**
 * Simple XML node structure for comparison
 */
interface XmlNode {
  tag: string;
  attrs: Record<string, string>;
  children: (XmlNode | string)[];
}

/**
 * Parse XML string into a tree structure for semantic comparison.
 * This is a simple parser sufficient for test comparison purposes.
 */
function parseXmlTree(xml: string): XmlNode | string | null {
  const trimmed = xml.trim();
  if (!trimmed) {
    return null;
  }

  // Handle XML declaration
  const content = trimmed.replace(/<\?xml[^?]*\?>\s*/g, "");

  const stack: XmlNode[] = [];
  let root: XmlNode | null = null;
  let pos = 0;

  while (pos < content.length) {
    const nextTag = content.indexOf("<", pos);
    if (nextTag === -1) {
      // Remaining text
      const text = decodeXmlEntities(content.slice(pos).trim());
      if (text && stack.length > 0) {
        stack[stack.length - 1].children.push(text);
      }
      break;
    }

    // Text before tag
    if (nextTag > pos) {
      const text = decodeXmlEntities(content.slice(pos, nextTag).trim());
      if (text && stack.length > 0) {
        stack[stack.length - 1].children.push(text);
      }
    }

    // Find end of tag
    const tagEnd = content.indexOf(">", nextTag);
    if (tagEnd === -1) {
      break;
    }

    const tagContent = content.slice(nextTag + 1, tagEnd);
    pos = tagEnd + 1;

    // Closing tag
    if (tagContent.startsWith("/")) {
      stack.pop();
      continue;
    }

    // Self-closing tag
    const selfClosing = tagContent.endsWith("/");
    const cleanContent = selfClosing ? tagContent.slice(0, -1).trim() : tagContent;

    // Parse tag name and attributes
    const spaceIdx = cleanContent.search(/\s/);
    const tagName = spaceIdx === -1 ? cleanContent : cleanContent.slice(0, spaceIdx);
    const attrStr = spaceIdx === -1 ? "" : cleanContent.slice(spaceIdx);

    const attrs: Record<string, string> = {};
    const attrRegex = /([\w:-]+)="([^"]*)"/g;
    let match;
    while ((match = attrRegex.exec(attrStr))) {
      attrs[match[1]] = match[2];
    }

    const node: XmlNode = { tag: tagName, attrs, children: [] };

    if (stack.length > 0) {
      stack[stack.length - 1].children.push(node);
    } else {
      root = node;
    }

    if (!selfClosing) {
      stack.push(node);
    }
  }

  return root;
}

/**
 * Convert XML tree to a canonical form for comparison.
 * Sorts attributes and children by tag name for consistent comparison.
 */
function treeToCanonical(node: XmlNode | string | null): any {
  if (node === null) {
    return null;
  }
  if (typeof node === "string") {
    return node;
  }

  const result: any = { tag: node.tag };

  // Sort attributes
  const sortedAttrs = Object.keys(node.attrs).sort();
  if (sortedAttrs.length > 0) {
    result.attrs = {};
    for (const key of sortedAttrs) {
      result.attrs[key] = node.attrs[key];
    }
  }

  // Sort and canonicalize children
  if (node.children.length > 0) {
    const childNodes = node.children
      .map(c => treeToCanonical(c))
      .filter(c => c !== null && c !== "");

    // Sort child elements by tag name (text nodes stay in place conceptually)
    childNodes.sort((a, b) => {
      const aTag = typeof a === "string" ? "" : a.tag;
      const bTag = typeof b === "string" ? "" : b.tag;
      return aTag.localeCompare(bTag);
    });

    if (childNodes.length > 0) {
      result.children = childNodes;
    }
  }

  return result;
}

/**
 * Normalize XML for semantic comparison.
 */
function normalizeXml(xml: string): string {
  if (!xml || !xml.trim()) {
    return "";
  }
  const tree = parseXmlTree(xml);
  const canonical = treeToCanonical(tree);
  return JSON.stringify(canonical);
}

interface Expectation {
  title: string;
  create(): any;
  initialModel?: any;
  preparedModel?: any;
  parsedModel?: any;
  reconciledModel?: any;
  xml?: string;
  tests: string[];
  options?: any;
}

function getExpectation<K extends keyof Expectation>(
  expectation: Expectation,
  name: K
): Expectation[K] {
  if (!Object.prototype.hasOwnProperty.call(expectation, name)) {
    throw new Error(`Expectation missing required field: ${name}`);
  }
  return cloneDeep((expectation as any)[name], true);
}

// ===============================================================================================================
// provides boilerplate examples for the four transform steps: prepare, render,  parse and reconcile
//  prepare: model => preparedModel
//  render:  preparedModel => xml
//  parse:  xml => parsedModel
//  reconcile: parsedModel => reconciledModel

const its: { [key: string]: (expectation: Expectation) => () => Promise<void> } = {
  prepare(expectation: Expectation) {
    return async () => {
      const model = getExpectation(expectation, "initialModel");
      const result = getExpectation(expectation, "preparedModel");

      const xform = expectation.create();
      xform.prepare(model, expectation.options);
      expect(cloneDeep(model, false)).toEqual(result);
    };
  },

  render(expectation: Expectation) {
    return async () => {
      const model = getExpectation(expectation, "preparedModel");
      const result = getExpectation(expectation, "xml");

      const xform = expectation.create();
      const xmlStream = new XmlWriter();
      xform.render(xmlStream, model, 0);

      expect(normalizeXml(xmlStream.xml)).toBe(normalizeXml(result!));
    };
  },

  "prepare-render": function (expectation: Expectation) {
    // when implementation details get in the way of testing the prepared result
    return async () => {
      const model = getExpectation(expectation, "initialModel");
      const result = getExpectation(expectation, "xml");

      const xform = expectation.create();
      const xmlStream = new XmlWriter();

      xform.prepare(model, expectation.options);
      xform.render(xmlStream, model);

      expect(normalizeXml(xmlStream.xml)).toBe(normalizeXml(result!));
    };
  },

  renderIn(expectation: Expectation) {
    return async () => {
      const model = {
        pre: true,
        child: getExpectation(expectation, "preparedModel"),
        post: true
      };
      const result = `<compy><pre/>${getExpectation(expectation, "xml")}<post/></compy>`;

      const xform = new CompyXform({
        tag: "compy",
        children: [
          {
            name: "pre",
            xform: new BooleanXform({ tag: "pre", attr: "val" })
          },
          { name: "child", xform: expectation.create() },
          {
            name: "post",
            xform: new BooleanXform({ tag: "post", attr: "val" })
          }
        ]
      });

      const xmlStream = new XmlWriter();
      xform.render(xmlStream, model);

      expect(normalizeXml(xmlStream.xml)).toBe(normalizeXml(result));
    };
  },

  parseIn(expectation: Expectation) {
    return async () => {
      const xml = `<compy><pre/>${getExpectation(expectation, "xml")}<post/></compy>`;
      const childXform = expectation.create();
      const result: any = { pre: true };
      result[childXform.tag] = getExpectation(expectation, "parsedModel");
      result.post = true;
      const xform: any = new CompyXform({
        tag: "compy",
        children: [
          {
            name: "pre",
            xform: new BooleanXform({ tag: "pre", attr: "val" })
          },
          { name: childXform.tag, xform: childXform },
          {
            name: "post",
            xform: new BooleanXform({ tag: "post", attr: "val" })
          }
        ]
      });
      const stream = new PassThrough();
      stream.write(xml);
      stream.end();
      const model = await xform.parse(parseSax(stream));

      // eliminate the undefined
      const clone = cloneDeep(model, false);

      expect(clone).toEqual(result);
    };
  },

  parse(expectation: Expectation) {
    return async () => {
      const xml = getExpectation(expectation, "xml");
      const result = getExpectation(expectation, "parsedModel");

      const xform = expectation.create();

      const stream = new PassThrough();
      stream.write(xml);
      stream.end();
      const model = await xform.parse(parseSax(stream));

      // eliminate the undefined
      const clone = cloneDeep(model, false);

      expect(clone).toEqual(result);
    };
  },

  reconcile(expectation: Expectation) {
    return async () => {
      const model = getExpectation(expectation, "parsedModel");
      const result = getExpectation(expectation, "reconciledModel");

      const xform = expectation.create();
      xform.reconcile(model, expectation.options);

      // eliminate the undefined
      const clone = cloneDeep(model, false);

      expect(clone).toEqual(result);
    };
  }
};

export function testXformHelper(expectations: Expectation[]) {
  expectations.forEach((expectation: Expectation) => {
    const tests = getExpectation(expectation, "tests");
    describe(expectation.title, () => {
      tests.forEach((test: string) => {
        it(test, its[test](expectation));
      });
    });
  });
}

// Export normalizeXml for use in custom tests
export { normalizeXml };
