import { readFile, readdir, stat } from "node:fs/promises";
import { join, relative } from "node:path";

export interface EnterpriseCorpusEntry {
  path: string;
  tags?: string[];
  expectCharts?: boolean;
  expectChartEx?: boolean;
  expectPivotTables?: boolean;
  /** Require an Office-compatible open/convert validation for this workbook when the gate is enabled. */
  openValidation?: boolean;
  /** Optional reason or source for traceability, e.g. "Excel 365 generated". */
  source?: string;
}

export interface EnterpriseCorpusManifest {
  entries: EnterpriseCorpusEntry[];
}

export async function loadEnterpriseCorpusManifest(
  manifestPath: string
): Promise<EnterpriseCorpusManifest> {
  const parsed = JSON.parse(await readFile(manifestPath, "utf8")) as EnterpriseCorpusManifest;
  if (!parsed || !Array.isArray(parsed.entries)) {
    throw new Error("Enterprise corpus manifest must contain an entries array");
  }
  return parsed;
}

export async function discoverEnterpriseCorpus(root: string): Promise<EnterpriseCorpusEntry[]> {
  const entries: EnterpriseCorpusEntry[] = [];
  await walk(root, async file => {
    if (/\.xlsx$/i.test(file)) {
      entries.push({ path: relative(root, file) });
    }
  });
  return entries.sort((a, b) => a.path.localeCompare(b.path));
}

async function walk(dir: string, onFile: (file: string) => Promise<void>): Promise<void> {
  for (const name of await readdir(dir)) {
    const full = join(dir, name);
    const info = await stat(full);
    if (info.isDirectory()) {
      await walk(full, onFile);
    } else if (info.isFile()) {
      await onFile(full);
    }
  }
}
