import { execFile } from "node:child_process";
import { mkdir, mkdtemp, readFile, readdir, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, join } from "node:path";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

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
  envFlag: string;
  executableEnv: string;
  candidates: string[];
  args: string[];
  input: Uint8Array;
  inputName: string;
  outputGlob?: RegExp;
  timeoutMs?: number;
  /** Set to false for proprietary CLIs that do not support a cheap --version probe. */
  versionArgs?: string[] | false;
}

export async function runExternalOracle(
  options: ExternalOracleOptions
): Promise<ExternalOracleResult> {
  if (process.env[options.envFlag] !== "1") {
    return { available: false, skipped: `Set ${options.envFlag}=1 to enable.`, outputs: [] };
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
  const dir = await mkdtemp(join(tmpdir(), "excelts-oracle-"));
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
  envFlag: string;
  executableEnv: string;
  candidates: string[];
  args?: string[];
  input: Uint8Array;
  inputName: string;
  timeoutMs?: number;
  repairLogPatterns?: RegExp[];
  versionArgs?: string[] | false;
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
    versionArgs: options.versionArgs
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
  const values = [process.env[envName], ...candidates].filter((value): value is string => !!value);
  for (const value of values) {
    if (versionArgs === false) {
      return value;
    }
    try {
      await execFileAsync(value, versionArgs, { timeout: 10_000 });
      return value;
    } catch {
      // Try next optional oracle executable.
    }
  }
  return undefined;
}

async function collectOutputs(
  dir: string,
  include: RegExp | undefined
): Promise<Array<{ name: string; data: Uint8Array }>> {
  const result: Array<{ name: string; data: Uint8Array }> = [];
  for (const name of await readdir(dir)) {
    const filePath = join(dir, name);
    if (!(await stat(filePath)).isFile()) {
      continue;
    }
    if (include && !include.test(name)) {
      continue;
    }
    result.push({ name: basename(name), data: new Uint8Array(await readFile(filePath)) });
  }
  return result.sort((a, b) => a.name.localeCompare(b.name));
}
