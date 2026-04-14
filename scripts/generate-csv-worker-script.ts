#!/usr/bin/env node
// Generates src/modules/csv/worker/worker-script.generated.ts by bundling
// src/modules/csv/worker/worker.entry.ts into a single classic-worker script.

import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const projectRoot = process.cwd();
const configPath = path.join(projectRoot, "rolldown.csv-worker.config.ts");

const outDir = path.join(projectRoot, "tmp", "csv-worker-bundle");
const outJs = path.join(outDir, "csv-worker.iife.js");

const generatedPath = path.join(
  projectRoot,
  "src",
  "modules",
  "csv",
  "worker",
  "worker-script.generated.ts"
);

function readUtf8(filePath) {
  return fs.readFileSync(filePath, "utf8");
}

function writeUtf8(filePath, content) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, content, "utf8");
}

function normalizeEol(s) {
  return s.replace(/\r\n/g, "\n");
}

try {
  // Ensure outDir exists and is clean-ish
  fs.mkdirSync(outDir, { recursive: true });

  // Bundle via rolldown (already a dev dependency).
  // Run rolldown's CLI entry directly via Node to avoid Windows .cmd shim
  // issues with execFileSync and to prevent shell injection.
  const rolldownCli = path.join(projectRoot, "node_modules", "rolldown", "bin", "cli.mjs");
  execFileSync(process.execPath, [rolldownCli, "-c", configPath], { stdio: "inherit" });

  const js = normalizeEol(readUtf8(outJs));

  const generated =
    `/**\n` +
    ` * GENERATED FILE - DO NOT EDIT.\n` +
    ` *\n` +
    ` * Regenerate with: npm run generate:csv-worker\n` +
    ` */\n` +
    `\n` +
    `export const CSV_WORKER_SCRIPT = ${JSON.stringify(js)};\n`;

  writeUtf8(generatedPath, generated);

  // Format the generated file
  execFileSync(
    process.execPath,
    [path.join(projectRoot, "node_modules", "oxfmt", "bin", "oxfmt"), generatedPath],
    { stdio: "inherit" }
  );
} catch (err) {
  console.error("Failed to generate CSV worker script.");
  console.error(err);
  process.exitCode = 1;
}
