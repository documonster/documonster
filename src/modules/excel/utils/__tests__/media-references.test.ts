import { MediaReferenceTracker } from "@excel/utils/media-references";
import { describe, expect, it } from "vitest";

const encode = (text: string) => new TextEncoder().encode(text);

describe("MediaReferenceTracker", () => {
  it("resolves relationship targets against the owning part — relative, absolute and escaped", () => {
    const tracker = new MediaReferenceTracker();
    tracker.observe(
      "xl/drawings/_rels/drawing1.xml.rels",
      `<Relationships><Relationship Id="rId1" Target="../media/image1.png"/>` +
        `<Relationship Id="rId2" Target='/xl/media/My%20Logo.JPEG'/></Relationships>`
    );
    expect(tracker.isReferenced("xl/media/image1.png")).toBe(true);
    expect(tracker.isReferenced("xl/media/my logo.jpeg")).toBe(true);
    expect(tracker.isReferenced("xl/media/image11.png")).toBe(false);
    expect(tracker.isReferenced("xl/media/image2.png")).toBe(false);
  });

  it("resolves from the package root and from a nested part", () => {
    const tracker = new MediaReferenceTracker();
    tracker.observe("_rels/.rels", `<Relationship Id="r" Target="xl/media/root.png"/>`);
    tracker.observe(
      "xl/charts/_rels/chart1.xml.rels",
      `<pr:Relationship xmlns:pr="x" Id="r" Target="../media/fill.png"/>`
    );
    expect(tracker.isReferenced("xl/media/root.png")).toBe(true);
    expect(tracker.isReferenced("xl/media/fill.png")).toBe(true);
    // Same file name, different folder: not the medium.
    tracker.observe("xl/drawings/_rels/d.xml.rels", `<Relationship Target="media/other.png"/>`);
    expect(tracker.isReferenced("xl/media/other.png")).toBe(false);
  });

  it("reads a relationship part written as bytes or in chunks", () => {
    const tracker = new MediaReferenceTracker();
    tracker.observe("xl/x/_rels/a.xml.rels", encode(`<Relationship Target="../media/a.gif"/>`));
    tracker.observeChunks("xl/x/_rels/b.xml.rels", [
      "<Relationship Tar",
      encode(`get="../media/b.gif"/>`)
    ]);
    expect(tracker.isReferenced("xl/media/a.gif")).toBe(true);
    expect(tracker.isReferenced("xl/media/b.gif")).toBe(true);
  });

  it("ignores external targets and parts that are not relationship parts", () => {
    const tracker = new MediaReferenceTracker();
    tracker.observe(
      "xl/worksheets/_rels/sheet1.xml.rels",
      `<Relationship Target="../media/image3.png" TargetMode="External"/>`
    );
    tracker.observe("xl/opaque/part.xml", encode(`<x src="xl/media/image7.png"/>`));
    tracker.observe("xl/media/image1.png", encode(`Target="image1.png"`));
    expect(tracker.wants("xl/opaque/part.xml")).toBe(false);
    expect(tracker.wants("xl/opaque/_rels/part.xml.rels")).toBe(true);
    expect(tracker.isReferenced("xl/media/image3.png")).toBe(false);
    expect(tracker.isReferenced("xl/media/image7.png")).toBe(false);
    expect(tracker.isReferenced("xl/media/image1.png")).toBe(false);
  });
});
