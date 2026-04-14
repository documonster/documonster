import {
  RawXmlCollector,
  serializeAttributes
} from "@excel/xlsx/xform/pivot-table/raw-xml-collector";
import { describe, it, expect } from "vitest";

describe("RawXmlCollector", () => {
  describe("start() / basic lifecycle", () => {
    it("should activate when start() is called", () => {
      const collector = new RawXmlCollector("extLst");
      expect(collector.active).toBe(false);
      collector.start();
      expect(collector.active).toBe(true);
    });

    it("should produce opening and closing root tags on immediate close", () => {
      const collector = new RawXmlCollector("extLst");
      collector.start();
      const done = collector.feedClose("extLst");
      expect(done).toBe(true);
      expect(collector.active).toBe(false);
      expect(collector.result).toBe("<extLst></extLst>");
    });

    it("should include root attributes from start()", () => {
      const collector = new RawXmlCollector("extLst");
      collector.start({ uri: "http://example.com", val: "42" });
      collector.feedClose("extLst");
      expect(collector.result).toBe('<extLst uri="http://example.com" val="42"></extLst>');
    });

    it("should handle start() with null attributes", () => {
      const collector = new RawXmlCollector("extLst");
      collector.start(null);
      collector.feedClose("extLst");
      expect(collector.result).toBe("<extLst></extLst>");
    });

    it("should handle start() with empty attributes object", () => {
      const collector = new RawXmlCollector("extLst");
      collector.start({});
      collector.feedClose("extLst");
      expect(collector.result).toBe("<extLst></extLst>");
    });
  });

  describe("startAs()", () => {
    it("should change the root tag and activate", () => {
      const collector = new RawXmlCollector("");
      collector.startAs("formats", { count: "3" });
      expect(collector.active).toBe(true);
      expect(collector.rootTag).toBe("formats");
      collector.feedClose("formats");
      expect(collector.result).toBe('<formats count="3"></formats>');
    });

    it("should allow reuse with different tags", () => {
      const collector = new RawXmlCollector("");

      collector.startAs("alpha");
      collector.feedClose("alpha");
      expect(collector.result).toBe("<alpha></alpha>");

      collector.reset();

      collector.startAs("beta", { id: "1" });
      collector.feedClose("beta");
      expect(collector.result).toBe('<beta id="1"></beta>');
    });
  });

  describe("reset()", () => {
    it("should deactivate and clear buffer", () => {
      const collector = new RawXmlCollector("test");
      collector.start();
      collector.feedOpen("child", { a: "1" });
      expect(collector.active).toBe(true);

      collector.reset();
      expect(collector.active).toBe(false);
      // result after reset is empty since buffer is cleared
      expect(collector.result).toBe("");
    });

    it("should allow reuse after reset", () => {
      const collector = new RawXmlCollector("root");
      collector.start();
      collector.feedOpen("a", {});
      collector.feedClose("a");
      collector.feedClose("root");
      const firstResult = collector.result;
      expect(firstResult).toContain("<root>");

      collector.reset();
      collector.start({ ver: "2" });
      collector.feedClose("root");
      expect(collector.result).toBe('<root ver="2"></root>');
    });
  });

  describe("feedOpen() / feedClose()", () => {
    it("should collect a single child element", () => {
      const collector = new RawXmlCollector("parent");
      collector.start();
      collector.feedOpen("child", { name: "foo" });
      collector.feedClose("child");
      collector.feedClose("parent");
      expect(collector.result).toBe('<parent><child name="foo" /></parent>');
    });

    it("should collect multiple child elements", () => {
      const collector = new RawXmlCollector("items");
      collector.start({ count: "2" });
      collector.feedOpen("item", { v: "a" });
      collector.feedClose("item");
      collector.feedOpen("item", { v: "b" });
      collector.feedClose("item");
      collector.feedClose("items");
      expect(collector.result).toBe('<items count="2"><item v="a" /><item v="b" /></items>');
    });

    it("should handle nested elements", () => {
      const collector = new RawXmlCollector("root");
      collector.start();
      collector.feedOpen("level1", {});
      collector.feedOpen("level2", { id: "deep" });
      collector.feedClose("level2");
      collector.feedClose("level1");
      collector.feedClose("root");
      expect(collector.result).toBe('<root><level1><level2 id="deep" /></level1></root>');
    });

    it("should not finish when closing a non-root tag at depth 0", () => {
      const collector = new RawXmlCollector("root");
      collector.start();
      collector.feedOpen("child", {});
      // close child — depth goes back to 0 but tag != rootTag, so not done
      const done = collector.feedClose("child");
      expect(done).toBe(false);
      expect(collector.active).toBe(true);
    });

    it("should return true only when root tag closes at depth 0", () => {
      const collector = new RawXmlCollector("wrap");
      collector.start();
      collector.feedOpen("inner", {});
      expect(collector.feedClose("inner")).toBe(false);
      expect(collector.feedClose("wrap")).toBe(true);
      expect(collector.active).toBe(false);
    });
  });

  describe("feedText()", () => {
    it("should collect text content inside elements", () => {
      const collector = new RawXmlCollector("note");
      collector.start();
      collector.feedOpen("body", {});
      collector.feedText("Hello World");
      collector.feedClose("body");
      collector.feedClose("note");
      expect(collector.result).toBe("<note><body>Hello World</body></note>");
    });

    it("should XML-encode special characters in text", () => {
      const collector = new RawXmlCollector("data");
      collector.start();
      collector.feedText("a < b & c > d");
      collector.feedClose("data");
      expect(collector.result).toBe("<data>a &lt; b &amp; c &gt; d</data>");
    });

    it("should prevent self-closing collapse when text appears between open and close", () => {
      const collector = new RawXmlCollector("root");
      collector.start();
      collector.feedOpen("el", {});
      collector.feedText("content");
      collector.feedClose("el");
      collector.feedClose("root");
      // Should NOT be self-closing since there's text content
      expect(collector.result).toBe("<root><el>content</el></root>");
    });
  });

  describe("self-closing collapse", () => {
    it("should collapse empty child elements to self-closing tags", () => {
      const collector = new RawXmlCollector("root");
      collector.start();
      collector.feedOpen("empty", {});
      collector.feedClose("empty");
      collector.feedClose("root");
      expect(collector.result).toBe("<root><empty /></root>");
    });

    it("should collapse empty child with attributes to self-closing", () => {
      const collector = new RawXmlCollector("root");
      collector.start();
      collector.feedOpen("leaf", { a: "1", b: "2" });
      collector.feedClose("leaf");
      collector.feedClose("root");
      expect(collector.result).toBe('<root><leaf a="1" b="2" /></root>');
    });

    it("should NOT collapse when child has nested content", () => {
      const collector = new RawXmlCollector("root");
      collector.start();
      collector.feedOpen("parent", {});
      collector.feedOpen("child", {});
      collector.feedClose("child");
      collector.feedClose("parent");
      collector.feedClose("root");
      // parent has child content, so it should NOT self-close
      expect(collector.result).toBe("<root><parent><child /></parent></root>");
    });
  });

  describe("depth tracking", () => {
    it("should track depth correctly for deeply nested structures", () => {
      const collector = new RawXmlCollector("root");
      collector.start();
      collector.feedOpen("a", {});
      collector.feedOpen("b", {});
      collector.feedOpen("c", {});
      collector.feedText("deep");
      collector.feedClose("c");
      collector.feedClose("b");
      collector.feedClose("a");
      collector.feedClose("root");
      expect(collector.active).toBe(false);
      expect(collector.result).toBe("<root><a><b><c>deep</c></b></a></root>");
    });

    it("should handle depth guard against negative depth", () => {
      // Simulates a mismatched close tag scenario — depth should not go negative
      const collector = new RawXmlCollector("root");
      collector.start();
      // Extra close without matching open — depth stays at 0
      const done = collector.feedClose("root");
      expect(done).toBe(true);
      expect(collector.active).toBe(false);
    });
  });
});

describe("serializeAttributes", () => {
  it("should return empty string for undefined", () => {
    expect(serializeAttributes(undefined)).toBe("");
  });

  it("should return empty string for null", () => {
    expect(serializeAttributes(null)).toBe("");
  });

  it("should return empty string for empty object", () => {
    expect(serializeAttributes({})).toBe("");
  });

  it("should serialize simple string attributes", () => {
    const result = serializeAttributes({ name: "test", id: "42" });
    expect(result).toBe('name="test" id="42"');
  });

  it("should serialize numeric attributes", () => {
    const result = serializeAttributes({ count: 5, min: 0 });
    expect(result).toBe('count="5" min="0"');
  });

  it("should filter out null and undefined values", () => {
    const result = serializeAttributes({ a: "keep", b: null, c: undefined, d: "also" });
    expect(result).toBe('a="keep" d="also"');
  });

  it("should XML-encode attribute values", () => {
    const result = serializeAttributes({ label: 'A & "B"' });
    expect(result).toBe('label="A &amp; &quot;B&quot;"');
  });
});
