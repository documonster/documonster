/**
 * The format-independent prefix normaliser: attribute maps that are data, prefixed attributes that keep
 * their namespace, the reader's parse policy, DTD-declared namespaces, and bounded name caches.
 */
import { XmlParseError } from "@xml/errors";
import { Xml } from "@xml/index";
import {
  canonicalizeXml,
  createPrefixNormalizer,
  createPrefixVocabulary,
  needsPrefixNormalization,
  normalizeElementPrefixes
} from "@xml/namespace-prefixes";
import { describe, expect, it } from "vitest";

const PREFIXED = createPrefixVocabulary(new Map([["urn:r", "r"]]));
const DEFAULT = createPrefixVocabulary(new Map([["urn:r", ""]]));

describe("namespace prefix normalisation", () => {
  it("keeps attributes named like Object.prototype members", () => {
    expect(
      canonicalizeXml(PREFIXED, '<x:root xmlns:x="urn:r" __proto__="data" constructor="value"/>')
    ).toBe('<r:root xmlns:r="urn:r" __proto__="data" constructor="value"/>');
  });

  it("keeps a prefixed attribute in its namespace when that namespace becomes the default", () => {
    const xml = '<x:root xmlns:x="urn:r" k="plain" x:k="namespaced"/>';
    expect(canonicalizeXml(DEFAULT, xml)).toBe(
      '<root xmlns="urn:r" k="plain" xmlns:x="urn:r" x:k="namespaced"/>'
    );
    const named = createPrefixNormalizer(DEFAULT).open({
      name: "x:root",
      attributes: { "xmlns:x": "urn:r", k: "plain", "x:k": "namespaced" },
      isSelfClosing: true
    });
    expect({ ...named.attributes }).toEqual({
      xmlns: "urn:r",
      k: "plain",
      "xmlns:x": "urn:r",
      "x:k": "namespaced"
    });
  });

  it("parses under the caller's policy and throws what it cannot parse", () => {
    const deep = (n: number): string =>
      `<x:root xmlns:x="urn:r">${"<x:a>".repeat(n)}${"</x:a>".repeat(n)}</x:root>`;
    expect(() => canonicalizeXml(PREFIXED, deep(300))).toThrow(XmlParseError);
    expect(canonicalizeXml(PREFIXED, deep(300), { maxDepth: 512 })).toContain("<r:a>");
    const control = '<x:root xmlns:x="urn:r">A\u0001B</x:root>';
    expect(() => canonicalizeXml(PREFIXED, control)).toThrow(XmlParseError);
    expect(canonicalizeXml(PREFIXED, control, { invalidCharHandling: "skip" })).toBe(
      '<r:root xmlns:r="urn:r">AB</r:root>'
    );
  });

  it("sees a namespace a DTD declares", () => {
    const xml = '<!DOCTYPE x:root [<!ATTLIST x:root xmlns:x CDATA "urn:r">]><x:root/>';
    expect(needsPrefixNormalization(PREFIXED, xml)).toBe(true);
    expect(canonicalizeXml(PREFIXED, xml)).toBe('<r:root xmlns:r="urn:r"/>');
    expect(createPrefixNormalizer(PREFIXED, { lazy: true }).scan(xml)).toBe(true);
  });

  it("renames every one of many distinct names in one scope", () => {
    const normalizer = createPrefixNormalizer(PREFIXED);
    normalizer.open({ name: "x:root", attributes: { "xmlns:x": "urn:r" }, isSelfClosing: false });
    for (let i = 0; i < 5000; i++) {
      const name = normalizer.open({ name: `x:n${i}`, attributes: {}, isSelfClosing: true }).name;
      expect(name).toBe(`r:n${i}`);
      normalizer.close(`x:n${i}`);
    }
  });
});

describe("normalizeElementPrefixes", () => {
  it("renames a tree as deep as the reader's policy allows, past what recursion held", () => {
    const depth = 50000;
    const xml = '<x:r xmlns:x="urn:r">' + "<x:a>".repeat(depth) + "</x:a>".repeat(depth) + "</x:r>";
    const { root } = Xml.parse(xml, { maxDepth: 0 });
    normalizeElementPrefixes(PREFIXED, root);
    let element = root;
    let levels = 0;
    while (element.children[0]?.type === "element") {
      element = element.children[0];
      levels++;
    }
    expect([root.name, element.name, levels]).toEqual(["r:r", "r:a", depth]);
  });
});
