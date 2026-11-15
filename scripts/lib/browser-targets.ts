/**
 * The browser floor, as the bundler's lowering target.
 *
 * `package.json`'s `browserslist` is the one statement of the floor; the README's table restates it. The
 * IIFE bundles are built from source by rolldown, whose transform defaults to `esnext` — no lowering — so
 * without a target the bundles shipped whatever syntax the transform produced. TypeScript's static class
 * fields came out as class static blocks, which Safari parses only from 16.4, and every bundle that held a
 * class with one failed to load on Safari 15.4–16.3. `dist/esm` was never affected: `tsc` lowers to the
 * `target` in `tsconfig.json`.
 *
 * Only the `browser >= version` form is used there, so only that form is read; anything else is refused
 * rather than silently dropped from the target.
 */

import fs from "node:fs";
import path from "node:path";

/** `browserslist` as rolldown's `transform.target`: `["chrome98", "safari15.4", …]`. */
export function browserTargets(root = process.cwd()): string[] {
  const pkg = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8")) as {
    browserslist?: unknown;
  };
  if (!Array.isArray(pkg.browserslist) || pkg.browserslist.length === 0) {
    throw new Error("package.json: browserslist must be a non-empty array");
  }
  return pkg.browserslist.map(query => {
    const match = typeof query === "string" ? /^(\w+) >= ([\d.]+)$/.exec(query) : null;
    if (!match) {
      throw new Error(
        `package.json: browserslist entry ${JSON.stringify(query)} is not "<browser> >= <version>"`
      );
    }
    return `${match[1]}${match[2]}`;
  });
}
