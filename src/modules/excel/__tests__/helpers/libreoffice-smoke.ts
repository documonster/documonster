import { execFile } from "node:child_process";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, join } from "node:path";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

export interface LibreOfficeSmokeResult {
  available: boolean;
  skipped?: string;
  output?: Uint8Array;
  stdout?: string;
  stderr?: string;
}

export async function libreOfficeAvailable(): Promise<string | undefined> {
  const candidates = [process.env.LIBREOFFICE_BIN, "soffice", "libreoffice"].filter(
    (value): value is string => !!value
  );
  for (const candidate of candidates) {
    try {
      await execFileAsync(candidate, ["--version"], { timeout: 10_000 });
      return candidate;
    } catch {
      // Try next candidate; LibreOffice smoke tests are optional.
    }
  }
  return undefined;
}

export async function smokeRoundTripWithLibreOffice(
  input: Uint8Array,
  filename = "workbook.xlsx"
): Promise<LibreOfficeSmokeResult> {
  if (process.env.DOCUMONSTER_LIBREOFFICE_SMOKE !== "1") {
    return { available: false, skipped: "Set DOCUMONSTER_LIBREOFFICE_SMOKE=1 to enable." };
  }
  const binary = await libreOfficeAvailable();
  if (!binary) {
    return { available: false, skipped: "LibreOffice executable not found." };
  }
  const dir = await mkdtemp(join(tmpdir(), "documonster-lo-smoke-"));
  const outDir = join(dir, "out");
  try {
    await mkdir(outDir);
    const inputPath = join(dir, filename);
    await writeFile(inputPath, input);
    const { stdout, stderr } = await execFileAsync(
      binary,
      ["--headless", "--convert-to", "xlsx", "--outdir", outDir, inputPath],
      { timeout: 60_000 }
    );
    const outputPath = join(outDir, basename(filename));
    const output = await readFile(outputPath);
    return { available: true, output: new Uint8Array(output), stdout, stderr };
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}
