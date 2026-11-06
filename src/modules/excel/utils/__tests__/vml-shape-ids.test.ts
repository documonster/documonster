import {
  allocateVmlShapeIds,
  claimVmlBlock,
  createVmlShapeIdSpace,
  reserveVmlShapeIds,
  vmlIdMap
} from "@excel/utils/vml-shape-ids";
import { describe, expect, it } from "vitest";

describe("VML shape-id allocation", () => {
  it("gives each part its own block, numbered from the block's first slot", () => {
    const space = createVmlShapeIdSpace();
    expect(allocateVmlShapeIds(space, { count: 2 }).free).toEqual([1025, 1026]);
    expect(allocateVmlShapeIds(space, { count: 1 }).free).toEqual([2049]);
  });

  it("keeps preserved ids whose blocks are free, and fills the rest of the block around them", () => {
    const space = createVmlShapeIdSpace();
    const ids = allocateVmlShapeIds(space, { fixed: [1025, 1027], count: 2 });
    expect(ids.fixed).toEqual([1025, 1027]);
    expect(ids.free).toEqual([1026, 1028]);
  });

  it("renumbers preserved ids whose block another part already holds", () => {
    const space = createVmlShapeIdSpace();
    allocateVmlShapeIds(space, { fixed: [1025], count: 0 });
    expect(allocateVmlShapeIds(space, { fixed: [1025, 1026], count: 1 })).toEqual({
      fixed: [2049, 2050],
      free: [2051]
    });
  });

  it("leaves a reserved block to its owner", () => {
    const space = createVmlShapeIdSpace();
    const later = {};
    reserveVmlShapeIds(space, later, [1025]);
    expect(allocateVmlShapeIds(space, { count: 1 }).free).toEqual([2049]);
    expect(allocateVmlShapeIds(space, { owner: later, fixed: [1025], count: 0 }).fixed).toEqual([
      1025
    ]);
  });

  it("continues in a further block past 1023 shapes", () => {
    const space = createVmlShapeIdSpace();
    const { free } = allocateVmlShapeIds(space, { count: 1024 });
    expect(free.at(-2)).toBe(2047);
    expect(free.at(-1)).toBe(2049);
    expect(vmlIdMap(free)).toBe("1,2");
    expect(claimVmlBlock(space)).toBe(3);
  });

  it("renumbers ids that are duplicated or outside any block", () => {
    const space = createVmlShapeIdSpace();
    expect(allocateVmlShapeIds(space, { fixed: [1025, 1025], count: 0 }).fixed).toEqual([
      1025, 1026
    ]);
    expect(allocateVmlShapeIds(space, { fixed: [7], count: 0 }).fixed).toEqual([2049]);
  });
});
