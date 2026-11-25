/**
 * Validates every Word document and workbook the examples write with Microsoft's Open XML SDK (`OpenXmlValidator`),
 * and fails on any error it reports that is not a known disagreement with Office's own files.
 *
 * **Why this one, beside `verify:office`.** Opening a file in Office is the final word on whether it opens, but Word
 * is lenient: it repairs a child out of schema order, or an element it does not know, in silence — so a Word document
 * that opens says little about whether it is right. The SDK's validator is Office's schema, made executable: it refuses
 * what Word quietly repairs, and what Office documents it does not allow although ECMA-376 does ([MS-OI29500]: a table
 * style's row height, both names of one table edge). Run against the examples it found a sparkline colour without its
 * alpha, a table column format Excel ignored, a form control's checked state as a number, 3-D pie and ruby properties
 * the schema does not have or requires, and more — none of which any other check here had caught.
 *
 * Calibrated before it was trusted: an out-of-order `<sheetPr>` (issue #269), an out-of-order `<w:tcPr>`, an unknown
 * element and malformed XML each fail it, and a file Excel itself saved passes but for the disagreements below.
 *
 * **Known disagreements.** A few of its reports are Office's own conventions, which a file Excel saves carries too.
 * Each is listed in {@link KNOWN}, matched by the validator's error id and where it is, with the evidence; anything else
 * fails.
 *
 * ## Usage
 *
 * ```
 * pnpm verify:examples                       # writes the files this checks
 * pnpm verify:openxml                        # skips loudly when .NET is absent
 * pnpm verify:openxml --require              # absence is a failure
 * pnpm verify:openxml --dir some/where       # another directory instead of the examples' output
 * ```
 *
 * It needs the .NET SDK (8 or later); the validator is built once, under `tmp/openxml-validator`, from
 * `scripts/openxml-validator`, which pins the SDK version so results do not move under it.
 *
 * CI runs it with `--require` in the `Examples` job, over the files that job has just written.
 */
import { execFileSync, spawn } from "node:child_process";
import { cp, readFile, readdir } from "node:fs/promises";
import { join, relative, resolve } from "node:path";
import process from "node:process";

/** Where `pnpm verify:examples` writes Word documents and workbooks. */
const EXAMPLE_OUTPUT = ["tmp/excel-examples", "tmp/word-examples", "tmp/pdf-examples"];

const SOURCE = "scripts/openxml-validator";
const BUILD = "tmp/openxml-validator";

interface ValidationError {
  readonly part?: string;
  readonly path?: string;
  readonly id: string;
  readonly description: string;
  /** The element the report is about, as written. */
  readonly node?: string;
}

interface Result {
  readonly file: string;
  readonly errors?: readonly ValidationError[];
  readonly failure?: string;
}

/**
 * `xml` — an element as the validator reports it, `<p:name attr="…" xmlns:p="…" />` — when it holds exactly the
 * attributes named, each value passing its test, and nothing else: no other attribute, no child, no text.
 */
function elementIs(
  xml: string | undefined,
  name: string,
  attributes: Readonly<Record<string, (value: string) => boolean>>
): boolean {
  const match = new RegExp(`^<${name}((?:\\s+[\\w:]+="[^"]*")*)\\s*/>$`).exec(xml ?? "");
  if (match === null) {
    return false;
  }
  const found = new Map<string, string>();
  for (const [, attr, value] of match[1]!.matchAll(/\s+([\w:]+)="([^"]*)"/g)) {
    if (!attr!.startsWith("xmlns")) {
      found.set(attr!, value!);
    }
  }
  return (
    found.size === Object.keys(attributes).length &&
    Object.entries(attributes).every(([attr, valid]) => found.has(attr) && valid(found.get(attr)!))
  );
}

/** A bin count as Excel writes it: `val`, a positive whole number, and nothing else. */
const isExcelBinCount = (xml: string | undefined): boolean =>
  elementIs(xml, "cx:binCount", { val: value => /^[1-9]\d*$/.test(value) });

/** A bin size as Excel writes it: `val`, a positive number, and nothing else. */
const isExcelBinSize = (xml: string | undefined): boolean =>
  elementIs(xml, "cx:binSize", {
    val: value => /^\d*\.?\d+(?:[eE][+-]?\d+)?$/.test(value) && Number(value) > 0
  });

/**
 * Reports that are Office's convention rather than a fault: a file Excel 16 saved from the same workbook carries each
 * of them. Each is matched by the validator's error id, the part, the error's XPath with its indexes removed, the text
 * of the report, **and the element itself** — written as Excel writes it, with a valid value. The report's text cannot
 * tell `<cx:binCount val="8"/>` from `val="abc"` or no `val` at all; the element can.
 */
const KNOWN: readonly {
  id: string;
  part: RegExp;
  path: RegExp;
  description: RegExp;
  node: (xml: string | undefined) => boolean;
  why: string;
}[] = [
  {
    id: "Sem_UniqueAttributeValue",
    part: /^\/xl\/drawings\/drawing\d+\.xml$/,
    path: /\/mc:AlternateContent\/mc:Fallback\/xdr:sp\/xdr:nvSpPr\/xdr:cNvPr$/,
    description:
      /^Attribute 'id' should have unique value\. Its current value '0' duplicates with others\.$/,
    node: xml =>
      elementIs(xml, "xdr:cNvPr", { id: value => value === "0", name: value => value === "" }),
    why: 'a chartEx anchor\'s fallback shape is `id="0" name=""`, as Excel writes every one of them'
  },
  {
    id: "Sch_UndeclaredAttribute",
    part: /^\/xl\/charts\/chartEx\d+\.xml$/,
    path: /\/cx:binning\/cx:bin(?:Count|Size)$/,
    description: /^The 'val' attribute is not declared\.$/,
    node: xml => isExcelBinCount(xml) || isExcelBinSize(xml),
    why: 'Excel writes a histogram\'s bin count and size as `<cx:binCount val="8"/>`; the SDK expects text'
  },
  {
    id: "Sch_ElementValueDataTypeDetailed",
    part: /^\/xl\/charts\/chartEx\d+\.xml$/,
    path: /\/cx:binning\/cx:bin(?:Count|Size)$/,
    description: /has invalid value ''\. The text value cannot be empty\.$/,
    node: xml => isExcelBinCount(xml) || isExcelBinSize(xml),
    why: "the same element: with the value in `val`, its text is empty"
  }
];

/** The compound-file signature an encrypted OOXML package is wrapped in; the SDK cannot open one. */
const COMPOUND_FILE_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

function dotnet(): string | undefined {
  const candidate = process.env["DOTNET"] ?? "dotnet";
  try {
    execFileSync(candidate, ["--version"], { stdio: "ignore" });
    return candidate;
  } catch {
    return undefined;
  }
}

async function collect(dir: string, out: string[]): Promise<void> {
  let entries: Awaited<ReturnType<typeof readdir>>;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }
  for (const entry of entries) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) {
      await collect(path, out);
    } else if (/\.(docx|docm|xlsx|xlsm)$/.test(entry.name) && !entry.name.startsWith("~$")) {
      out.push(path);
    }
  }
}

function isKnown(error: ValidationError): boolean {
  const path = (error.path ?? "").replace(/\[\d+\]/g, "");
  return KNOWN.some(
    known =>
      known.id === error.id &&
      known.part.test(error.part ?? "") &&
      known.path.test(path) &&
      known.description.test(error.description) &&
      known.node(error.node)
  );
}

function validate(tool: string, files: readonly string[]): Promise<Result[]> {
  return new Promise((done, fail) => {
    const child = spawn(tool, [join(BUILD, "out", "OpenXmlValidator.dll")], {
      stdio: ["pipe", "pipe", "inherit"]
    });
    let output = "";
    child.stdout.on("data", (chunk: Buffer) => (output += chunk.toString()));
    child.on("error", fail);
    child.on("close", code => {
      if (code !== 0) {
        fail(new Error(`the validator exited ${code}`));
        return;
      }
      done(
        output
          .split("\n")
          .filter(line => line.trim() !== "")
          .map(line => JSON.parse(line) as Result)
      );
    });
    child.stdin.end(files.map(file => resolve(file)).join("\n") + "\n");
  });
}

async function main(): Promise<void> {
  const argv = process.argv.slice(2);
  const required = argv.includes("--require");
  const dirs = argv.flatMap((arg, i) => (arg === "--dir" ? [argv[i + 1]!] : []));

  const tool = dotnet();
  if (tool === undefined) {
    const message =
      "the .NET SDK was not found (`dotnet`, or set DOTNET). Install .NET 8 or later to run Microsoft's Open XML validator.";
    if (required) {
      console.error(`✗ verify:openxml — ${message}`);
      process.exit(1);
    }
    // Loud, and it says what was not checked: a quiet skip is how a gate dies green.
    console.log(`⚠ verify:openxml — skipped, nothing was verified.\n  ${message}`);
    return;
  }

  const files: string[] = [];
  for (const dir of dirs.length > 0 ? dirs : EXAMPLE_OUTPUT) {
    await collect(dir, files);
  }
  files.sort();
  if (files.length === 0) {
    console.error(
      `✗ verify:openxml — no documents to check${dirs.length > 0 ? ` under ${dirs.join(", ")}` : ". Run `pnpm verify:examples` first."}`
    );
    process.exit(1);
  }

  const encrypted: string[] = [];
  for (const file of files) {
    const head = (await readFile(file)).subarray(0, COMPOUND_FILE_SIGNATURE.length);
    if (COMPOUND_FILE_SIGNATURE.every((byte, i) => head[i] === byte)) {
      encrypted.push(file);
    }
  }
  const checked = files.filter(file => !encrypted.includes(file));
  if (checked.length === 0) {
    console.error(
      `✗ verify:openxml — all ${files.length} file(s) are encrypted; nothing could be checked.`
    );
    process.exit(1);
  }

  await cp(SOURCE, BUILD, { recursive: true });
  execFileSync(
    tool,
    ["build", BUILD, "-c", "Release", "-o", join(BUILD, "out"), "--nologo", "-v", "q"],
    {
      stdio: ["ignore", "ignore", "inherit"]
    }
  );
  const results = await validate(tool, checked);

  let failures = 0;
  let known = 0;
  // One result per file given, and no other: a crash part-way, or a line the validator did not write, would otherwise
  // pass on less than the run was asked to check.
  const answered = new Set(results.map(result => result.file));
  for (const file of checked) {
    if (!answered.has(resolve(file))) {
      failures += 1;
      console.log(`  ✗ ${file}\n      the validator returned no result for it`);
    }
  }
  if (results.length !== checked.length) {
    failures += 1;
    console.log(`  ✗ ${results.length} result(s) for ${checked.length} file(s)`);
  }
  for (const result of results) {
    if (result.failure !== undefined) {
      failures += 1;
      console.log(
        `  ✗ ${relative(process.cwd(), result.file)}\n      could not be read: ${result.failure}`
      );
      continue;
    }
    const errors = (result.errors ?? []).filter(error => {
      const isKnownError = isKnown(error);
      known += isKnownError ? 1 : 0;
      return !isKnownError;
    });
    if (errors.length > 0) {
      failures += 1;
      console.log(`  ✗ ${relative(process.cwd(), result.file)}`);
      for (const error of errors.slice(0, 5)) {
        console.log(
          `      ${error.part ?? ""} ${error.path ?? ""}\n        ${error.id}: ${error.description}`
        );
      }
      if (errors.length > 5) {
        console.log(`      … and ${errors.length - 5} more`);
      }
    }
  }
  if (encrypted.length > 0) {
    console.log(
      `\n${encrypted.length} encrypted file(s) not checked (the SDK opens no encrypted package):\n` +
        encrypted.map(file => `  - ${file}`).join("\n")
    );
  }
  if (failures > 0) {
    console.error(
      `\n✗ verify:openxml — ${failures} of ${checked.length} file(s) fail Office's schema (Open XML SDK).`
    );
    process.exit(1);
  }
  console.log(
    `\n✓ verify:openxml — all ${checked.length} file(s) pass Office's schema` +
      (known > 0 ? `, but for ${known} report(s) of Office's own conventions (see KNOWN).` : ".")
  );
}

await main();
