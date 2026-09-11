import { readFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

/**
 * The Node CRC32 variant must reach `zlib` from its first call, with nothing lazy in the way.
 *
 * Stated structurally because no functional test can see it. The variant previously loaded `zlib` with a
 * dynamic `import()` and computed on the portable lookup table until that promise settled — returning the
 * right answer, ~59x slower — so `crc32.test.ts` passed throughout and would pass again if the lazy loading
 * came back. A timing assertion is the wrong alternative: it turns a loaded CI runner into a failure and a
 * fast one into a false pass.
 */
import { crc32, crc32Finalize, crc32Update } from "@archive/compression/crc32";
import { describe, expect, it } from "vitest";

const compressionDir = join(dirname(fileURLToPath(import.meta.url)), "..", "..", "compression");

/** Code only: a scan that read prose would match a comment discussing `import()` as if it were code. */
async function readCode(file: string): Promise<string> {
  const source = await readFile(join(compressionDir, file), "utf8");
  return source.replace(/\/\*[\s\S]*?\*\//g, "").replace(/^\s*\/\/.*$/gm, "");
}

describe("crc32 native path", () => {
  it("imports zlib statically, so there is no window to fall back in", async () => {
    const code = await readCode("crc32.ts");

    expect(code).toMatch(/^import \{[^}]*\bcrc32 as nativeCrc32\b[^}]*\} from "zlib";$/m);
    // A dynamic import reintroduces the window; `require` cannot resolve in an ES module at all.
    expect(code).not.toMatch(/\bimport\s*\(/);
    expect(code).not.toMatch(/\brequire\s*\(/);
  });

  it("keeps zlib out of the browser variant entirely", async () => {
    const code = await readCode("crc32.browser.ts");

    expect(code).not.toMatch(/\bzlib\b/);
    expect(code).not.toMatch(/\bimport\s*\(/);
    expect(code).not.toMatch(/\brequire\s*\(/);
  });

  it("agrees with the lookup table on the very first call, unwarmed", () => {
    // First statement to touch the module: before the fix this ran on the table, and the point is that the
    // answer is the same either way — only the speed differed, which is why this needs the scan above.
    const data = new TextEncoder().encode("Hello, World!");
    expect(crc32(data)).toBe(0xec4ac3d0);

    let chunked = 0xffffffff;
    chunked = crc32Update(chunked, data.subarray(0, 5));
    chunked = crc32Update(chunked, data.subarray(5));
    expect(crc32Finalize(chunked)).toBe(0xec4ac3d0);
  });
});
