import { loadIife } from "@test/browser/load-iife";
import { describe, it, expect, beforeAll } from "vitest";

/**
 * Smoke test for the shipped `documonster.xml.iife.min.js` bundle: asserts the
 * `Documonster.Xml` surface plus parse and encode running in a browser.
 */
describe("Documonster.Xml IIFE bundle", () => {
  let Xml: any;

  beforeAll(async () => {
    ({ Xml } = await loadIife<{ Xml: any }>("xml", "Xml"));
  }, 60000);

  it("exposes the Xml namespace with parse/encode", () => {
    expect(typeof Xml.parse).toBe("function");
    expect(typeof Xml.encode).toBe("function");
  });

  it("parses XML into a DOM with the expected root", () => {
    const doc = Xml.parse("<r><c>1</c></r>");
    expect(doc.root.name).toBe("r");
  });

  it("encodes special characters", () => {
    expect(Xml.encode("a<b>")).toBe("a&lt;b&gt;");
  });
});
