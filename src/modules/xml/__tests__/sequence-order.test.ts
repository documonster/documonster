import { parseXml } from "@xml/dom";
import {
  checkSequenceOrder,
  checkSequenceOrderByName,
  sequenceRanks,
  type SequenceViolation,
  withoutAmbiguousNames
} from "@xml/sequence-order";
import { describe, expect, it } from "vitest";

/** `root` holds `a` then `b`; `b` holds `x` then `y`. */
const RANKS = sequenceRanks({ root: [["a"], ["b"]], b: [["x"], ["y"]] });
const TREE = { type: "root", children: { b: { type: "b" } } } as const;
const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';

function byTree(xml: string): string[] {
  const found: string[] = [];
  checkSequenceOrder(parseXml(xml).root, TREE, RANKS, collect(found));
  return found;
}

function byName(xml: string): string[] {
  const found: string[] = [];
  checkSequenceOrderByName(
    parseXml(xml).root,
    local => (local === "root" || local === "b" ? local : undefined),
    RANKS,
    collect(found)
  );
  return found;
}

function collect(found: string[]) {
  return ({ trail, child, after }: SequenceViolation<string>) =>
    found.push(`${trail.join("/")}: ${child} after ${after}`);
}

for (const [walk, check] of [
  ["checkSequenceOrder", byTree],
  ["checkSequenceOrderByName", byName]
] as const) {
  describe(`${walk} and markup compatibility`, () => {
    it("ranks a choice's children in the wrapper's place", () => {
      expect(
        check(
          `<root ${MC}><mc:AlternateContent><mc:Choice><b/></mc:Choice></mc:AlternateContent><a/></root>`
        )
      ).toEqual(["root: a after b"]);
    });

    it("descends into a fallback and ranks its children on their own", () => {
      expect(
        check(
          `<root ${MC}><a/><mc:AlternateContent><mc:Choice><b><x/><y/></b></mc:Choice><mc:Fallback><b><y/><x/></b></mc:Fallback></mc:AlternateContent></root>`
        )
      ).toEqual(["root/AlternateContent/Fallback/b: x after y"]);
      expect(
        check(
          `<root ${MC}><mc:AlternateContent><mc:Fallback><b/><a/></mc:Fallback></mc:AlternateContent></root>`
        )
      ).toEqual(["root/AlternateContent/Fallback: a after b"]);
    });

    it("unwraps a wrapper nested in a choice, and checks the alternatives inside it", () => {
      const nested = `<mc:AlternateContent><mc:Choice><b/></mc:Choice><mc:Fallback><b><y/><x/></b></mc:Fallback></mc:AlternateContent>`;
      expect(
        check(
          `<root ${MC}><mc:AlternateContent><mc:Choice>${nested}</mc:Choice></mc:AlternateContent><a/></root>`
        )
      ).toEqual([
        "root: a after b",
        // Read through the fallback, the root still has the nested wrapper's b before its a.
        "root/AlternateContent/Fallback: a after b",
        "root/AlternateContent/Fallback/b: x after y"
      ]);
    });

    it("ranks a fallback among the siblings around its wrapper, as a consumer reading it sees them", () => {
      // The choice is in order after b; the fallback puts an a after it, which a consumer reading it would see.
      expect(
        check(
          `<root ${MC}><b/><mc:AlternateContent><mc:Choice><b/></mc:Choice><mc:Fallback><a/></mc:Fallback></mc:AlternateContent></root>`
        )
      ).toEqual(["root/AlternateContent/Fallback: a after b"]);
    });

    it("checks a wrapper nested in a fallback as that fallback's own content", () => {
      expect(
        check(
          `<root ${MC}><mc:AlternateContent><mc:Fallback><mc:AlternateContent><mc:Choice><b/></mc:Choice></mc:AlternateContent><a/></mc:Fallback></mc:AlternateContent></root>`
        )
      ).toEqual(["root/AlternateContent/Fallback: a after b"]);
    });

    it("checks a second choice as an alternative of its own", () => {
      expect(
        check(
          `<root ${MC}><mc:AlternateContent><mc:Choice><a/></mc:Choice><mc:Choice><b/><a/></mc:Choice></mc:AlternateContent></root>`
        )
      ).toEqual(["root/AlternateContent/Choice: a after b"]);
    });

    it("never ranks one alternative against the other, or against the siblings around them", () => {
      // The fallback's `a` is not after the choice's `b`, nor after the `b` that follows the wrapper.
      expect(
        check(
          `<root ${MC}><mc:AlternateContent><mc:Choice><b/></mc:Choice><mc:Fallback><a/></mc:Fallback></mc:AlternateContent><b/></root>`
        )
      ).toEqual([]);
    });

    it("ranks a wrapper nested in a later branch among the siblings around the outer wrapper", () => {
      // Read through the outer fallback and then the inner one, the root holds b, then a — out of order.
      const inner = `<mc:AlternateContent><mc:Choice><b/></mc:Choice><mc:Fallback><a/></mc:Fallback></mc:AlternateContent>`;
      expect(
        check(
          `<root ${MC}><b/><mc:AlternateContent><mc:Choice><b/></mc:Choice><mc:Fallback>${inner}</mc:Fallback></mc:AlternateContent></root>`
        )
      ).toEqual(["root/AlternateContent/Fallback/AlternateContent/Fallback: a after b"]);
    });

    it("does not unwrap an AlternateContent of another namespace", () => {
      // An extension's own AlternateContent is an element like any other, not a compatibility wrapper.
      expect(
        check(
          `<root ${MC} xmlns:e="urn:extension"><b/><e:AlternateContent><e:Choice><a/></e:Choice></e:AlternateContent></root>`
        )
      ).toEqual([]);
    });
  });
}

describe("the problem cap", () => {
  it("stops before an alternative once the check says stop", () => {
    const xml = `<root ${MC}><b/><a/><mc:AlternateContent><mc:Fallback><b/><a/></mc:Fallback></mc:AlternateContent></root>`;
    const found: string[] = [];
    checkSequenceOrder(parseXml(xml).root, TREE, RANKS, collect(found), () => found.length >= 1);
    expect(found).toEqual(["root: a after b"]);
  });
});

describe("withoutAmbiguousNames", () => {
  it("drops a name allowed at two positions, and a position left empty", () => {
    expect(withoutAmbiguousNames([["a", "m"], ["m"], ["b"]])).toEqual([["a"], ["b"]]);
  });
});

describe("matching a child to its model by namespace", () => {
  const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const STRICT = "http://purl.oclc.org/ooxml/spreadsheetml/main";
  const X_RANKS = sequenceRanks({ "x:root": [["a"], ["b"]], "x:b": [["x"], ["y"]] });
  const X_TREE = { type: "x:root", children: { b: { type: "x:b" } } } as const;
  const found = (xml: string): string[] => {
    const out: string[] = [];
    checkSequenceOrder(parseXml(xml).root, X_TREE, X_RANKS, collect(out));
    return out;
  };

  it("does not take an element of another namespace for the standard one it shares a name with", () => {
    expect(found(`<root xmlns="${MAIN}" xmlns:e="urn:extension"><b/><e:a/></root>`)).toEqual([]);
    expect(found(`<root xmlns="${MAIN}"><b/><a/></root>`)).toEqual(["root: a after b"]);
  });

  it("takes the Strict namespace for the Transitional one", () => {
    expect(found(`<root xmlns="${STRICT}"><b/><a/></root>`)).toEqual(["root: a after b"]);
  });

  it("matches by local name an element whose prefix is not declared", () => {
    expect(found(`<root><b/><a/></root>`)).toEqual(["root: a after b"]);
  });

  it("resolves a prefix declared on an mc:Choice", () => {
    expect(
      found(
        `<s:root xmlns:s="${MAIN}" ${MC}><s:b/><mc:AlternateContent><mc:Choice xmlns:t="${MAIN}" Requires="t"><t:a/></mc:Choice></mc:AlternateContent></s:root>`
      )
    ).toEqual(["root: a after b"]);
  });

  it("holds a child the table prefixes to the namespace it names, not its parent's", () => {
    const DRAW = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const ranks = sequenceRanks({ "x:root": [["a"], ["a:graphic"]] });
    const tree = { type: "x:root" } as const;
    const run = (xml: string): string[] => {
      const out: string[] = [];
      checkSequenceOrder(parseXml(xml).root, tree, ranks, collect(out));
      return out;
    };
    expect(run(`<root xmlns="${MAIN}" xmlns:d="${DRAW}"><d:graphic/><a/></root>`)).toEqual([
      "root: a after graphic"
    ]);
    // A graphic of the parent's own namespace is not the one the model declares.
    expect(run(`<root xmlns="${MAIN}"><graphic/><a/></root>`)).toEqual([]);
  });

  it("does not descend into an element of another namespace", () => {
    expect(
      found(`<root xmlns="${MAIN}" xmlns:e="urn:extension"><e:b><y/><x/></e:b></root>`)
    ).toEqual([]);
    expect(found(`<root xmlns="${MAIN}"><b><y/><x/></b></root>`)).toEqual(["root/b: x after y"]);
  });
});
