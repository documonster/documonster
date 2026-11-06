import path from "node:path";

import { describe, expect, it } from "vitest";

// The scripts are run by Node directly and import with `.ts` extensions, which the library's tsconfig refuses.
const SCRIPT = path.resolve(import.meta.dirname, "../../../scripts/api-changes/sig-cell.ts");
const { sigCell } = (await import(SCRIPT)) as { sigCell(sig: string, other?: string): string };

describe("api-changes sigCell", () => {
  it("leaves a short signature whole, escaped for a table", () => {
    expect(sigCell("(a: string | number) => void")).toBe("`(a: string \\| number) => void`");
  });

  it("clips a long signature to its start", () => {
    const cell = sigCell(`{ ${"k: T; ".repeat(100)}}`);
    expect(cell.length).toBeLessThan(200);
    expect(cell.startsWith("`{ k: T;")).toBe(true);
    expect(cell.endsWith("…`")).toBe(true);
  });

  it("shows where two long signatures differ", () => {
    const body = "k: T; ".repeat(100);
    const a = `{ ${body}old: 1; ${body}}`;
    const b = `{ ${body}new: 2; ${body}}`;
    expect(sigCell(a, b)).toContain("old: 1;");
    expect(sigCell(b, a)).toContain("new: 2;");
    expect(sigCell(a, b).startsWith("`…")).toBe(true);
    expect(sigCell(a, b).length).toBeLessThan(200);
  });
});

describe("api-changes sigCell overloads", () => {
  it("leaves out overloads both sides share, and caps the rest", () => {
    const shared = ["(a: 1) => void", "(a: 2) => void"];
    const before = [...shared, "(a: 3) => void"].join("\n");
    const after = [...shared, "(a: 3, b?: 4) => void"].join("\n");
    expect(sigCell(before, after)).toBe("`(a: 3) => void`<br>_(2 unchanged)_");
    const many = Array.from({ length: 9 }, (_, i) => `(a: ${i}) => void`).join("\n");
    expect(sigCell(many).split("<br>")).toHaveLength(5);
    expect(sigCell(many)).toContain("_(5 more)_");
  });
});
