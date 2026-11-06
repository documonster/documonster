/**
 * Fetch a published package from npm into a directory, for the scripts that compare the current tree with a release
 * (`scripts/api-changes.ts`, `scripts/verify-excelts-guide.ts`).
 */
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

export function run(cmd: string, args: string[], cwd: string): string {
  return execFileSync(cmd, args, {
    cwd,
    encoding: "utf8",
    stdio: ["ignore", "pipe", "inherit"]
  }).trim();
}

/** `npm pack` a published package into `dir` and unpack it; returns the package directory. */
export function fetchPackage(spec: string, dir: string): string {
  fs.mkdirSync(dir, { recursive: true });
  const tarball = run("npm", ["pack", spec, "--silent", "--pack-destination", dir], dir)
    .split("\n")
    .pop()!;
  run("tar", ["-xzf", path.join(dir, tarball), "-C", dir], dir);
  return path.join(dir, "package");
}
