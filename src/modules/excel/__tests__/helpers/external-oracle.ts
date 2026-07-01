import { execFile } from "node:child_process";
import { mkdir, mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, join } from "node:path";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

/** Cache resolved executables so we don't probe the filesystem repeatedly. */
const resolveCache = new Map<string, string | undefined>();

export interface ExternalOracleResult {
  available: boolean;
  skipped?: string;
  executable?: string;
  exitCode?: number | null;
  stdout?: string;
  stderr?: string;
  outputs: Array<{ name: string; data: Uint8Array }>;
}

export interface ExternalOracleOptions {
  /**
   * Environment variable that gates the oracle. When set to `"1"` the
   * oracle runs unconditionally (and fails the test if the executable
   * cannot be located or the conversion errors).
   *
   * Pass `null` to use {@link autoMode} semantics instead — the oracle
   * runs whenever the executable is auto-discovered, otherwise it
   * gracefully reports `available: false` so callers can `expect(result.skipped).toBeTruthy()`.
   */
  envFlag: string | null;
  executableEnv: string;
  candidates: string[];
  args: string[];
  input: Uint8Array;
  inputName: string;
  outputGlob?: RegExp;
  timeoutMs?: number;
  /** Set to false for proprietary CLIs that do not support a cheap --version probe. */
  versionArgs?: string[] | false;
  /**
   * When true, run the oracle whenever the executable is discoverable,
   * skipping otherwise. Useful for "default-on if installed, off in
   * minimal CI" semantics that do not require an env flag opt-in.
   *
   * Ignored when {@link envFlag} is set to a non-null string and that
   * env var explicitly equals `"1"` (the explicit opt-in always wins).
   */
  autoMode?: boolean;
}

export async function runExternalOracle(
  options: ExternalOracleOptions
): Promise<ExternalOracleResult> {
  const explicitOptIn = options.envFlag !== null && process.env[options.envFlag] === "1";
  const auto = options.autoMode === true;
  if (!explicitOptIn && !auto) {
    return {
      available: false,
      skipped: options.envFlag
        ? `Set ${options.envFlag}=1 to enable.`
        : "autoMode disabled and no envFlag opt-in",
      outputs: []
    };
  }
  const executable = await resolveExecutable(
    options.executableEnv,
    options.candidates,
    options.versionArgs
  );
  if (!executable) {
    return {
      available: false,
      skipped: `${options.executableEnv} executable not found.`,
      outputs: []
    };
  }
  const dir = await mkdtemp(join(tmpdir(), "documonster-oracle-"));
  const outDir = join(dir, "out");
  try {
    await mkdir(outDir);
    const inputPath = join(dir, options.inputName);
    await writeFile(inputPath, options.input);
    const args = options.args.map(arg =>
      arg.replace(/\{input\}/g, inputPath).replace(/\{outDir\}/g, outDir)
    );
    const { stdout, stderr } = await execFileAsync(executable, args, {
      timeout: options.timeoutMs ?? 120_000
    });
    const outputs = await collectOutputs(outDir, options.outputGlob);
    return { available: true, executable, exitCode: 0, stdout, stderr, outputs };
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

export interface OfficeOpenValidationOptions {
  /** Pass `null` to rely on {@link autoMode} only. */
  envFlag: string | null;
  executableEnv: string;
  candidates: string[];
  args?: string[];
  input: Uint8Array;
  inputName: string;
  timeoutMs?: number;
  repairLogPatterns?: RegExp[];
  versionArgs?: string[] | false;
  /** See {@link ExternalOracleOptions.autoMode}. */
  autoMode?: boolean;
}

export async function runOfficeOpenValidation(
  options: OfficeOpenValidationOptions
): Promise<ExternalOracleResult> {
  const result = await runExternalOracle({
    envFlag: options.envFlag,
    executableEnv: options.executableEnv,
    candidates: options.candidates,
    args: options.args ?? ["--headless", "--convert-to", "xlsx", "--outdir", "{outDir}", "{input}"],
    input: options.input,
    inputName: options.inputName,
    outputGlob: /[.]xlsx$/i,
    timeoutMs: options.timeoutMs,
    versionArgs: options.versionArgs,
    autoMode: options.autoMode
  });
  if (!result.available) {
    return result;
  }
  const output = `${result.stdout ?? ""}\n${result.stderr ?? ""}`;
  const repairLogPatterns = options.repairLogPatterns ?? [
    /repair/i,
    /repaired/i,
    /corrupt/i,
    /error/i
  ];
  const repairHit = repairLogPatterns.find(pattern => pattern.test(output));
  if (repairHit) {
    const message = `Office open validation reported a possible repair/error (${repairHit}).`;
    return {
      ...result,
      available: true,
      exitCode: 1,
      stderr: [result.stderr, message].filter(Boolean).join("\n"),
      outputs: result.outputs
    };
  }
  return result;
}

async function resolveExecutable(
  envName: string,
  candidates: string[],
  versionArgs: string[] | false = ["--version"]
): Promise<string | undefined> {
  const cacheKey = `${envName}:${candidates.join(",")}:${String(versionArgs)}`;
  if (resolveCache.has(cacheKey)) {
    return resolveCache.get(cacheKey);
  }
  const values = [process.env[envName], ...candidates].filter((value): value is string => !!value);
  for (const value of values) {
    if (versionArgs === false) {
      resolveCache.set(cacheKey, value);
      return value;
    }
    try {
      await execFileAsync(value, versionArgs, { timeout: 10_000 });
      resolveCache.set(cacheKey, value);
      return value;
    } catch {
      // Try next optional oracle executable.
    }
  }
  resolveCache.set(cacheKey, undefined);
  return undefined;
}

/**
 * Convenience wrapper that runs LibreOffice's `--convert-to xlsx`
 * round-trip in auto mode: if `LIBREOFFICE_BIN` (or `soffice` /
 * `libreoffice` on PATH) is discoverable the validation runs and
 * `expect(exitCode).toBe(0)` passes; otherwise the result reports
 * `available: false` and callers should `expect(skipped).toBeTruthy()`.
 *
 * Used by the synthetic chart corpus tests so the open-validation gate
 * runs by default for everyone with LibreOffice installed, without
 * forcing an explicit env-var opt-in or breaking minimal CI environments.
 */
export async function runLibreOfficeOpenValidationAuto(
  input: Uint8Array,
  inputName: string
): Promise<ExternalOracleResult> {
  return runOfficeOpenValidation({
    envFlag: "DOCUMONSTER_LIBREOFFICE_OPEN_VALIDATION",
    executableEnv: "LIBREOFFICE_BIN",
    candidates: [
      "soffice",
      "libreoffice",
      // macOS app bundle install (default Homebrew Cask + manual install location).
      "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    ],
    input,
    inputName,
    autoMode: true
  });
}

async function collectOutputs(
  dir: string,
  include: RegExp | undefined
): Promise<Array<{ name: string; data: Uint8Array }>> {
  const result: Array<{ name: string; data: Uint8Array }> = [];
  for (const name of await readdir(dir)) {
    if (include && !include.test(name)) {
      continue;
    }
    const filePath = join(dir, name);
    // Read directly — skip on failure (avoids TOCTOU race between stat and read).
    try {
      const data = new Uint8Array(await readFile(filePath));
      result.push({ name: basename(name), data });
    } catch {
      // Not a readable file (directory, permission error, removed between readdir and read).
    }
  }
  return result.sort((a, b) => a.name.localeCompare(b.name));
}
