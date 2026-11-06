/**
 * Which lines of a TypeScript snippet call something `@deprecated`.
 *
 * Deprecation is a suggestion, which `tsc` never reports, and the TypeScript in this repository
 * (7, native) has no JS language-service API — so it is read from the type-aware lint rule
 * `typescript/no-deprecated` (tsgolint).
 *
 * The probe is written to its own directory under `tmp/` with a `tsconfig.json` that extends the
 * project's (paths, strictness) and includes the probe alone, so probes written by tests running
 * in parallel are never part of each other's program. The report is read as JSON and a lint that
 * did not run throws: matching a regex over the human-readable output read any change in that
 * output — or a failed run — as "nothing deprecated".
 */
import { spawnSync } from "node:child_process";
import { mkdirSync, rmSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";

const ROOT = path.resolve(__dirname, "../..");
// The package's own entry, run with this Node: `.bin/oxlint` is a shell shim on Windows.
const OXLINT = path.join(
  path.dirname(createRequire(path.join(ROOT, "package.json")).resolve("oxlint/package.json")),
  "bin/oxlint"
);

export function deprecatedLines(lines: readonly string[]): number[] {
  const dir = path.join(
    ROOT,
    "tmp",
    `deprecation-probe-${process.pid}-${Math.random().toString(36).slice(2)}`
  );
  mkdirSync(dir, { recursive: true });
  try {
    writeFileSync(
      path.join(dir, "tsconfig.json"),
      // `exclude` is reset because the root's excludes `tmp`, which would drop the probe silently.
      JSON.stringify({
        extends: path.join(ROOT, "tsconfig.json"),
        include: ["probe.ts"],
        exclude: []
      })
    );
    const probe = path.join(dir, "probe.ts");
    writeFileSync(probe, lines.join("\n") + "\n");
    const result = spawnSync(
      process.execPath,
      [
        OXLINT,
        "--no-ignore",
        "--type-aware",
        "--format",
        "json",
        "-A",
        "all",
        "-D",
        "typescript/no-deprecated",
        probe
      ],
      { cwd: ROOT, encoding: "utf8" }
    );
    // 0: nothing reported; 1: something reported, on stdout. Anything else is a failure to lint,
    // which must not read as "nothing deprecated".
    if (result.status !== 0 && result.status !== 1) {
      throw new Error(
        `oxlint exited with ${result.status ?? result.signal}: ${result.stderr || result.stdout}`
      );
    }
    const report = JSON.parse(result.stdout) as {
      diagnostics: { code: string; labels: { span: { line: number } }[] }[];
    };
    return report.diagnostics
      .filter(d => d.code === "typescript(no-deprecated)")
      .map(d => d.labels[0].span.line)
      .sort((a, b) => a - b);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}
