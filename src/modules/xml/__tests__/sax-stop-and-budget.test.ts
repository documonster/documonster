import { XmlArgumentError, XmlError } from "@xml/errors";
import { SaxParser } from "@xml/sax";
import { describe, it, expect } from "vitest";

const encoder = new TextEncoder();

describe("SaxParser — a stopped parse reports nothing more", () => {
  for (const xml of ["<??><r/>", "<?XML?><r/>", "<?xml?><r/>"]) {
    it(`reports no pi for ${xml} once the error handler stops`, () => {
      const parser = new SaxParser();
      const events: string[] = [];
      parser.on("error", () => {
        events.push("error");
        parser.stop();
      });
      parser.on("pi", () => events.push("pi"));
      parser.on("opentag", () => events.push("open"));
      parser.write(xml);
      parser.close();
      expect(events).toEqual(["error"]);
    });
  }

  it("keeps the position of a ]]> in text once the error handler stops", () => {
    const parser = new SaxParser();
    const seen: number[] = [];
    parser.on("error", () => {
      parser.stop();
      seen.push(parser.position);
    });
    parser.on("text", () => {});
    parser.write(encoder.encode("<r>ab]]>cd</r>"));
    seen.push(parser.position);
    expect(seen).toEqual([7, 7]);
  });
});

describe("SaxParser — attribute defaults from the internal subset", () => {
  it("charges an entity-built default once, where it is declared", () => {
    const given: string[] = [];
    const parser = new SaxParser({
      internalEntities: true,
      maxEntityExpansions: 1,
      maxEntityExpansionLength: 1
    });
    parser.on("opentag", tag => {
      if (tag.name === "a") {
        given.push(tag.attributes.v);
      }
    });
    parser.write('<!DOCTYPE r [<!ENTITY e "x"><!ATTLIST a v CDATA "&e;">]><r><a/><a/><a/></r>');
    parser.close();
    expect(given).toEqual(["x", "x", "x"]);
    // A default that takes more than the limits allow is refused where it is declared.
    expect(() =>
      new SaxParser({ internalEntities: true, maxEntityExpansionLength: 1 }).write(
        '<!DOCTYPE r [<!ENTITY e "xy"><!ATTLIST a v CDATA "&e;">]><r/>'
      )
    ).toThrow(/entity expansion/);
  });

  it("refuses a field a context's declaration does not have", () => {
    expect(
      () =>
        new SaxParser({
          resume: true,
          context: {
            root: "r",
            entities: {},
            attributes: { a: { v: { default: "x", expansions: 1 } } }
          } as never
        })
    ).toThrow(XmlArgumentError);
  });

  it("counts a prefixed attribute the DTD supplies", () => {
    const parser = new SaxParser();
    parser.countPrefixedAttributes = true;
    parser.write('<!DOCTYPE r [<!ATTLIST r x:v CDATA "1">]><r/>');
    parser.close();
    expect(parser.prefixedAttributes).toBe(1);
  });

  it("does not look into the defaults of a list skipped after a parameter entity", () => {
    const errors: string[] = [];
    const parser = new SaxParser();
    parser.on("error", e => errors.push(e.message));
    parser.write('<!DOCTYPE r [<!ENTITY % p "x">%p;<!ATTLIST r v CDATA "&nope;">]><r/>');
    parser.close();
    expect(errors).toEqual([]);
  });

  it("still looks into them in a standalone document", () => {
    const errors: string[] = [];
    const parser = new SaxParser();
    parser.on("error", e => errors.push(e.message));
    parser.write(
      '<?xml version="1.0" standalone="yes"?><!DOCTYPE r [<!ENTITY % p "x">%p;<!ATTLIST r v CDATA "&lt;">]><r/>'
    );
    parser.close();
    expect(errors).toEqual([]);
  });

  it("bounds what it keeps of the declarations without internalEntities", () => {
    const names = Array.from({ length: 200 }, (_, i) => `<!ATTLIST e${i} a${i} CDATA #IMPLIED>`);
    const parser = new SaxParser({ maxContextBytes: 1024 });
    const errors: unknown[] = [];
    parser.on("error", e => errors.push(e));
    parser.write(`<!DOCTYPE r [${names.join("")}]><r/>`);
    parser.close();
    expect(errors.length).toBeGreaterThan(0);
    expect((errors[0] as { limit?: string }).limit).toBe("prologContext");
  });
});

describe("SaxParser — references", () => {
  it("does not hold malformed references nobody reads", () => {
    const parser = new SaxParser();
    parser.on("error", () => {});
    parser.write("<r>" + "&a!".repeat(1000));
    expect((parser as unknown as { text: string }).text.length).toBeLessThan(10);
  });

  it("refuses a reference that is not a Name, whatever ENTITIES holds", () => {
    const parser = new SaxParser();
    parser.ENTITIES["1"] = "X";
    const errors: string[] = [];
    parser.on("error", e => errors.push(e.message));
    parser.write("<r>&1;</r>");
    parser.close();
    expect(errors[0]).toContain("invalid entity name");
  });

  it("reads only ENTITIES' own entries", () => {
    const parser = new SaxParser();
    parser.ENTITIES = Object.assign(Object.create({ inherited: "X" }) as Record<string, string>, {
      amp: "&"
    });
    const errors: string[] = [];
    parser.on("error", e => errors.push(e.message));
    parser.write("<r>&inherited;</r>");
    parser.close();
    expect(errors[0]).toContain("undefined entity: inherited");
  });
});

describe("SaxParser — pausing", () => {
  it("keeps a pause made in the event that declares ASCII", () => {
    const parser = new SaxParser();
    const events: string[] = [];
    parser.on("pi", () => parser.pause());
    parser.on("opentag", tag => events.push(tag.name));
    parser.write(encoder.encode('<?xml version="1.0" encoding="us-ascii"?><r/>'));
    expect(parser.paused).toBe(true);
    expect(events).toEqual([]);
    parser.resumeParsing();
    expect(events).toEqual(["r"]);
  });

  it("refuses resumeParsing from a handler", () => {
    const parser = new SaxParser();
    let thrown: unknown;
    parser.on("opentag", () => {
      parser.pause();
      try {
        parser.resumeParsing();
      } catch (e) {
        thrown = e;
      }
    });
    parser.write("<r><a/></r>");
    expect(thrown).toBeInstanceOf(XmlError);
  });
});
