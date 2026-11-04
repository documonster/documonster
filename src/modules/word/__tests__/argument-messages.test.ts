/**
 * An argument error says what it was given in words that tell it apart from
 * what was wanted: "options must be an object, got a Map", not "…, got an
 * object". Booleans are printed as values, as Excel and PDF print them.
 */

import { describe, it, expect } from "vitest";

import { Document, Io, Template } from "../index";

async function message(call: () => unknown): Promise<string> {
  try {
    await call();
  } catch (e) {
    return (e as Error).message;
  }
  return "(no error)";
}

describe("argument error messages", () => {
  it("name the built-in kind of a rejected object", async () => {
    const d = Document.create();
    const bytes = await Io.toBuffer(d);
    expect(await message(() => Io.read(bytes, new Map() as never))).toBe(
      "Io.read: options must be an object, got a Map"
    );
    expect(await message(() => Io.read(bytes, new Uint8Array(1) as never))).toBe(
      "Io.read: options must be an object, got a Uint8Array"
    );
    expect(await message(() => Document.setSettings(d, new Date() as never))).toBe(
      "Document.setSettings: settings must be an object, got a Date"
    );
    expect(await message(() => Template.fill(d, new Set() as never))).toBe(
      "Template.fill: data must be an object or a Map, got a Set"
    );
    expect(await message(() => Io.read(42 as never))).toMatch(/got 42$/);
  });

  it("print a boolean as its value", async () => {
    const d = Document.create();
    expect(await message(() => Document.mailMerge(d, true as never))).toBe(
      "Document.mailMerge: data must be an object or a Map, got true"
    );
  });

  it("keep a plain object as 'an object' and an array as 'an array'", async () => {
    const d = Document.create();
    expect(await message(() => Document.setSettings(d, [] as never))).toBe(
      "Document.setSettings: settings must be an object, got an array"
    );
  });

  it("call a Date that holds no time an invalid Date, not a Date", async () => {
    const d = Document.create();
    expect(await message(() => Io.toBuffer(d, { modTime: new Date(Number.NaN) }))).toBe(
      "Io.toBuffer: options.modTime must be a valid Date, got an invalid Date"
    );
    expect(await message(() => Io.read(new Uint8Array(4), new Date() as never))).toBe(
      "Io.read: options must be an object, got a Date"
    );
  });
});
