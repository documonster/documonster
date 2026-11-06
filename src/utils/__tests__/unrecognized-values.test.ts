import {
  discardUnrecognizedValues,
  findUnrecognizedValues,
  listUnrecognizedValues,
  setAsideValue,
  unrecognizedEntry,
  unrecognizedValue
} from "@utils/unrecognized-values";
import { describe, expect, it } from "vitest";

describe("unrecognized values", () => {
  it("is in effect while the field is absent, and gives way once it is set", () => {
    const holder: { alignment?: string } = {};
    setAsideValue(holder, "alignment", "bogus");
    expect(unrecognizedValue(holder, "alignment")).toBe("bogus");
    holder.alignment = "center";
    expect(unrecognizedValue(holder, "alignment")).toBeUndefined();
    expect(listUnrecognizedValues(holder)).toEqual([]);
    delete holder.alignment;
    expect(unrecognizedValue(holder, "alignment")).toBe("bogus");
  });

  it("for a field that cannot be absent, is in effect while it still holds what the reader put there", () => {
    const tab = { type: "left" };
    setAsideValue(tab, "type", "rightish", { readAs: "left" });
    expect(unrecognizedValue(tab, "type")).toBe("rightish");
    tab.type = "center";
    expect(unrecognizedValue(tab, "type")).toBeUndefined();
  });

  it("keeps the attributes that come with it", () => {
    const run = {};
    setAsideValue(run, "underline", "evil", { with: { color: "FF0000" } });
    expect(unrecognizedEntry(run, "underline")).toEqual({
      value: "evil",
      with: { color: "FF0000" }
    });
  });

  it("travels with a spread and a structured clone, and is found with its path", () => {
    const p = { properties: {} as Record<string, unknown> };
    setAsideValue(p.properties, "alignment", "bogus");
    const copy = structuredClone({ body: [{ ...p, properties: { ...p.properties } }] });
    expect(findUnrecognizedValues(copy)).toEqual([
      { path: "body[0].properties.alignment", value: "bogus" }
    ]);
    expect(findUnrecognizedValues({ data: new Uint8Array(4) })).toEqual([]);
  });

  it("is found and discarded inside a Map", () => {
    const p = { alignment: undefined as string | undefined };
    setAsideValue(p, "alignment", "bogus");
    const root = { headers: new Map([["default", { children: [p] }]]) };
    expect(findUnrecognizedValues(root)).toEqual([
      { path: 'headers["default"].children[0].alignment', value: "bogus" }
    ]);
    expect(discardUnrecognizedValues(root)).toBe(1);
    expect(findUnrecognizedValues(root)).toEqual([]);
    expect(p).not.toHaveProperty("_unrecognized");
  });

  it("discards one a field has since taken over, without counting it", () => {
    const holder: { alignment?: string } = {};
    setAsideValue(holder, "alignment", "bogus");
    holder.alignment = "center";
    expect(discardUnrecognizedValues(holder)).toBe(0);
    delete holder.alignment;
    expect(unrecognizedValue(holder, "alignment")).toBeUndefined();
  });
});
