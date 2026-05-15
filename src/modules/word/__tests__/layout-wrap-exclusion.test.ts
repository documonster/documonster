/**
 * layoutDocumentFull — square / tight / through wrap exclusion zones.
 *
 * Verifies that body paragraphs wrap their lines around a preceding
 * float with `wrap.style ∈ { "square" | "tight" | "through" }`, with
 * each affected line offset / shrunk to avoid the float's rect. The
 * cursor is NOT advanced (the float doesn't push paragraphs below it),
 * which is the distinguishing behaviour vs. `topAndBottom`.
 */

import { describe, it, expect } from "vitest";

import { layoutDocumentFull } from "../layout/layout-full";
import type { DocxDocument, FloatingImage, Paragraph } from "../types";

const longBodyText = (() => {
  // Long enough to span several lines on a US Letter content area
  // (about 6.5 inches wide).
  const word = "wrap-test-word ";
  return word.repeat(160);
})();

function para(text: string): Paragraph {
  return { type: "paragraph", children: [{ content: [{ type: "text", text }] }] };
}

const baseDoc = (overrides: Partial<DocxDocument> = {}): DocxDocument => ({
  body: [],
  styles: [],
  abstractNumberings: [],
  numberingInstances: [],
  headers: new Map(),
  footers: new Map(),
  footnotes: [],
  endnotes: [],
  comments: [],
  images: [],
  fonts: [],
  embeddedFonts: [],
  customXmlParts: [],
  customProperties: [],
  opaqueParts: [],
  ...overrides
});

describe("layoutDocumentFull — square wrap exclusion", () => {
  it("does NOT advance the cursor for a `square` float (text wraps around)", () => {
    const fi: FloatingImage = {
      type: "floatingImage",
      rId: "rId1",
      width: 914_400, // 1 inch wide
      height: 914_400, // 1 inch tall
      horizontalPosition: { relativeTo: "margin", align: "right" },
      verticalPosition: { relativeTo: "margin", align: "top" },
      wrap: { style: "square" }
    };
    const doc = baseDoc({ body: [fi, para("after the float")] });
    const layout = layoutDocumentFull(doc);
    const items = layout.pages[0].content;
    expect(items[0].type).toBe("float");
    const next = items[1];
    expect(next.type).toBe("paragraph");
    if (next.type === "paragraph") {
      // Cursor still at top — paragraph y must be ≈ 0 (or just
      // spaceBefore), NOT pushed past the float's 72pt height.
      expect(next.rect.y).toBeLessThan(20);
    }
  });

  it("shrinks lines that intersect a `square` float on the right (wrapSide: bothSides)", () => {
    // Float in upper-right corner (square wrap with bothSides).
    const fi: FloatingImage = {
      type: "floatingImage",
      rId: "rId1",
      width: 914_400, // 1 inch wide
      height: 914_400, // 1 inch tall
      horizontalPosition: { relativeTo: "margin", align: "right" },
      verticalPosition: { relativeTo: "margin", align: "top" },
      wrap: { style: "square", side: "bothSides" }
    };
    const doc = baseDoc({ body: [fi, para(longBodyText)] });
    const layout = layoutDocumentFull(doc);
    const paraItem = layout.pages[0].content[1];
    expect(paraItem.type).toBe("paragraph");
    if (paraItem.type === "paragraph") {
      // First several lines fall within the float's vertical band
      // (y in [0, 72]) so their last run's right edge must NOT pass
      // the float's left edge.
      const floatLeft = layout.pages[0].content[0].rect.x;
      const earlyLines = paraItem.lines.filter(l => l.y < 72);
      expect(earlyLines.length).toBeGreaterThan(0);
      for (const line of earlyLines) {
        const lastRun = line.runs[line.runs.length - 1];
        if (!lastRun) {
          continue;
        }
        const rightEdge = lastRun.x + lastRun.width;
        // Allow tiny rounding slack.
        expect(rightEdge).toBeLessThanOrEqual(floatLeft + 0.5);
      }
    }
  });

  it("returns lines to full width once the paragraph passes the float's bottom edge", () => {
    const fi: FloatingImage = {
      type: "floatingImage",
      rId: "rId1",
      width: 914_400,
      height: 914_400,
      horizontalPosition: { relativeTo: "margin", align: "right" },
      verticalPosition: { relativeTo: "margin", align: "top" },
      wrap: { style: "square", side: "bothSides" }
    };
    const doc = baseDoc({ body: [fi, para(longBodyText)] });
    const layout = layoutDocumentFull(doc);
    const paraItem = layout.pages[0].content[1];
    expect(paraItem.type).toBe("paragraph");
    if (paraItem.type === "paragraph") {
      // Lines below the float (y > 72pt) must be allowed to extend
      // to the full page width again — find at least one such line
      // and prove it can use the full content width.
      const lateLines = paraItem.lines.filter(l => l.y >= 72);
      expect(lateLines.length).toBeGreaterThan(0);
      const longestLate = Math.max(
        ...lateLines.map(l => {
          if (l.runs.length === 0) {
            return 0;
          }
          const lastRun = l.runs[l.runs.length - 1];
          return lastRun.x + lastRun.width;
        })
      );
      const longestEarly = Math.max(
        0,
        ...paraItem.lines
          .filter(l => l.y < 72)
          .map(l => {
            if (l.runs.length === 0) {
              return 0;
            }
            const lastRun = l.runs[l.runs.length - 1];
            return lastRun.x + lastRun.width;
          })
      );
      // Late lines are allowed to be wider than early (squeezed) lines.
      expect(longestLate).toBeGreaterThan(longestEarly);
    }
  });

  it("topAndBottom wrap still pushes the cursor (regression guard)", () => {
    const fi: FloatingImage = {
      type: "floatingImage",
      rId: "rId1",
      width: 914_400,
      height: 914_400,
      horizontalPosition: { relativeTo: "margin", align: "left" },
      verticalPosition: { relativeTo: "margin", align: "top" },
      wrap: { style: "topAndBottom" }
    };
    const doc = baseDoc({ body: [fi, para("after")] });
    const layout = layoutDocumentFull(doc);
    const items = layout.pages[0].content;
    const next = items[1];
    expect(next.type).toBe("paragraph");
    if (next.type === "paragraph") {
      // topAndBottom must still push the paragraph below the float.
      expect(next.rect.y).toBeGreaterThanOrEqual(72);
    }
  });
});
