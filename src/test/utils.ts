import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

function toPath(fileUrlOrPath: string): string {
  if (fileUrlOrPath.startsWith("file:")) {
    return fileURLToPath(fileUrlOrPath);
  }
  return fileUrlOrPath;
}

function findRepoRoot(startDir: string): string | undefined {
  let currentDir = startDir;

  while (true) {
    const candidate = path.join(currentDir, "package.json");
    if (fs.existsSync(candidate)) {
      return currentDir;
    }

    const parentDir = path.dirname(currentDir);
    if (parentDir === currentDir) {
      return undefined;
    }
    currentDir = parentDir;
  }
}

let cachedOutDir: string | undefined;

export function testOutDir(): string {
  if (cachedOutDir) {
    return cachedOutDir;
  }

  const repoRoot =
    findRepoRoot(process.cwd()) ??
    findRepoRoot(path.dirname(fileURLToPath(import.meta.url))) ??
    process.cwd();
  const outDir = path.join(repoRoot, "out");

  fs.mkdirSync(outDir, { recursive: true });
  cachedOutDir = outDir;
  return outDir;
}

function stripKnownTestSuffix(baseName: string): string {
  // Keeps legacy behavior for names like *.integration.test.ts
  return baseName.replace(/\.(?:test|spec)(?:\.[cm]?[jt]s)?$/i, "");
}

function uniqueTestFilePathFromFilename(filename: string, extension: string): string {
  const base = stripKnownTestSuffix(path.basename(filename));
  return path.join(testOutDir(), `${base}${extension}`);
}

export function getUniqueTestFilePath(testFileUrlOrPath: string, extension = ".xlsx"): string {
  return uniqueTestFilePathFromFilename(toPath(testFileUrlOrPath), extension);
}

export function testFilePath(name: string, extension = ".xlsx"): string {
  return path.join(testOutDir(), `${name}${extension}`);
}

export function makeTestDataPath(importMetaUrl: string, dataDirRelative: string) {
  const baseDir = path.dirname(fileURLToPath(importMetaUrl));
  const dataDir = path.resolve(baseDir, dataDirRelative);

  return (filename: string): string => path.join(dataDir, filename);
}
