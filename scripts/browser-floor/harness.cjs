// The browser floor, checked in the browsers themselves.
//
// Plain CommonJS on purpose: this runs inside the Playwright image that carries the floor's engine, and those
// images ship Node 16. `scripts/verify-browser-floor.ts` starts it; see there for why it exists.
//
// Usage: node harness.cjs <dist-dir> <workers-dir>  (with `playwright` resolvable, and ENGINE / EXPECT_VERSION /
// EXPECT_BUNDLES set). <workers-dir> holds the worker scripts, one file each, as `verify-browser-floor.ts` wrote them.
"use strict";

const fs = require("fs");
const http = require("http");
const path = require("path");

const ENGINE = process.env.ENGINE;
const EXPECT_VERSION = process.env.EXPECT_VERSION;
/**
 * The bundles the package publishes, `[{ global, file }]` from `scripts/lib/iife-bundles.ts` — passed in because that
 * table is TypeScript and this runs where nothing compiles it. It is what the check is held to: what the build
 * happened to leave in `dist/iife` cannot be, since a bundle the build failed to write would then be expected by
 * no one.
 */
const EXPECT_BUNDLES = process.env.EXPECT_BUNDLES;
const DIST = path.resolve(process.argv[2] || "dist");
const WORKERS = process.argv[3] && path.resolve(process.argv[3]);
/** The port the server listens on: any free one, so nothing else on the machine can be in the way. */
let port = 0;

function walk(dir, out) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) walk(full, out);
    else if (entry.name.endsWith(".js")) out.push(full);
  }
  return out;
}

const url = file =>
  `http://localhost:${port}/${path.relative(DIST, file).split(path.sep).join("/")}`;

/** The worker scripts in WORKERS, which the package hands to `new Worker` as strings: no file parse sees them. */
function workerScripts() {
  if (!WORKERS || !fs.existsSync(WORKERS)) {
    throw new Error(
      "no worker scripts directory: verify-browser-floor.ts writes one and passes it"
    );
  }
  const names = fs.readdirSync(WORKERS).filter(name => name.endsWith(".js"));
  if (names.length === 0) {
    throw new Error(`${WORKERS}: holds no worker scripts`);
  }
  return names.map(name => ({ name, text: fs.readFileSync(path.join(WORKERS, name), "utf8") }));
}

/** The bundle file a published bundle is written to, as `rolldown.config.ts` names it. */
const bundleFile = file => `documonster.${file}.iife.min.js`;

/**
 * `EXPECT_BUNDLES`, and how `dist/iife` differs from it: a listed bundle the build did not write, or a script there
 * that nothing lists. Either way the build is not what is published, and no browser is needed to say so.
 */
function bundleManifest() {
  if (!EXPECT_BUNDLES) {
    throw new Error(
      "EXPECT_BUNDLES is not set: the bundles to expect come from scripts/lib/iife-bundles.ts"
    );
  }
  const bundles = JSON.parse(EXPECT_BUNDLES);
  const dir = path.join(DIST, "iife");
  const present = new Set(fs.readdirSync(dir).filter(name => name.endsWith(".js")));
  const listed = new Set(bundles.map(bundle => bundleFile(bundle.file)));
  const problems = [
    ...[...listed]
      .filter(name => !present.has(name))
      .map(name => `iife/${name} — listed, but not built`),
    ...[...present]
      .filter(name => !listed.has(name))
      .map(name => `iife/${name} — built, but not listed`)
  ];
  return { bundles, problems };
}

async function main() {
  const manifest = bundleManifest();
  const scripts = workerScripts();
  if (manifest.problems.length > 0) {
    console.error(`✗ ${ENGINE}: dist/iife is not the published set of bundles`);
    for (const problem of manifest.problems) console.error(`  ${problem}`);
    process.exitCode = 1;
    return;
  }
  const { [ENGINE]: engine } = require("playwright");
  const server = http.createServer((req, res) => {
    if (req.url === "/") {
      res.setHeader("content-type", "text/html; charset=utf-8");
      res.end("<!doctype html><meta charset=utf-8><title>floor</title>");
      return;
    }
    const file = path.join(DIST, decodeURIComponent(req.url.split("?")[0]));
    if (!file.startsWith(DIST) || !fs.existsSync(file)) {
      res.statusCode = 404;
      res.end();
      return;
    }
    res.setHeader("content-type", "text/javascript; charset=utf-8");
    res.end(fs.readFileSync(file));
  });
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, resolve);
  });
  port = server.address().port;
  // Everything after the server listens is inside the `try`: a listening server keeps the process alive, so a
  // failure that skipped closing it — the browser not launching, say — would hang the CI job until its timeout
  // instead of failing it.
  let browser;
  const failures = [];
  try {
    browser = await engine.launch();
    const version = browser.version();
    console.log(`${ENGINE} ${version}`);
    if (EXPECT_VERSION && !version.startsWith(EXPECT_VERSION)) {
      failures.push(`expected ${ENGINE} ${EXPECT_VERSION}, the image provides ${version}`);
    }

    // 1. Every shipped file parses. A module is imported from a blob: its relative imports cannot resolve there,
    //    which is a TypeError, while a file the engine cannot read is a SyntaxError — reported before linking.
    const page = await browser.newPage();
    await page.goto(`http://localhost:${port}/`);
    const files = [...walk(path.join(DIST, "esm"), []), ...walk(path.join(DIST, "iife"), [])];
    for (const file of files) {
      const isModule = file.startsWith(path.join(DIST, "esm"));
      const result = await page.evaluate(
        async ({ href, isModule }) => {
          const text = await (await fetch(href)).text();
          if (!isModule) {
            try {
              new Function(text);
              return "ok";
            } catch (e) {
              return `${e.name}: ${e.message}`;
            }
          }
          const blob = URL.createObjectURL(
            new Blob([text], { type: "text/javascript;charset=utf-8" })
          );
          try {
            await import(blob);
            return "ok";
          } catch (e) {
            return e instanceof SyntaxError ? `${e.name}: ${e.message}` : "ok";
          } finally {
            URL.revokeObjectURL(blob);
          }
        },
        { href: url(file), isModule }
      );
      if (result !== "ok") failures.push(`${path.relative(DIST, file)} — ${result}`);
    }

    // 2. The worker scripts parse as the classic scripts a Worker runs.
    for (const script of scripts) {
      const result = await page.evaluate(text => {
        try {
          new Function(text);
          return "ok";
        } catch (e) {
          return `${e.name}: ${e.message}`;
        }
      }, script.text);
      if (result !== "ok") failures.push(`worker script ${script.name} — ${result}`);
    }
    console.log(`  parsed ${files.length} files and ${scripts.length} worker scripts`);

    // 3. Every published bundle loads as a <script>, and every module it carries works: each global has a check of
    //    its own, and a global with none is a failure rather than a module taken on trust.
    const app = await browser.newPage();
    const pageErrors = [];
    app.on("pageerror", error => pageErrors.push(String(error)));
    await app.goto(`http://localhost:${port}/`);
    for (const { file } of manifest.bundles) {
      await app
        .addScriptTag({ url: url(path.join(DIST, "iife", bundleFile(file))) })
        .catch(error => pageErrors.push(`${bundleFile(file)}: ${error}`));
    }
    for (const error of pageErrors) failures.push(`loading the IIFE bundles — ${error}`);
    const smoke = await app.evaluate(async () => {
      const D = globalThis.Documonster || {};
      const checks = {
        Excel: async () => {
          const { Workbook, Cell } = D.Excel;
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "S1");
          Cell.setValue(ws, "A1", { formula: "LET(x,2,x*3)" });
          Cell.setValue(ws, "A2", { formula: "LAMBDA(a,b,a+b)(1,2)" });
          Cell.setValue(ws, "A3", "Safari");
          const back = await Workbook.read(await Workbook.toBuffer(wb));
          const sheet = Workbook.getWorksheet(back, "S1");
          return [
            Cell.getFormula(sheet, "A1"),
            Cell.getFormula(sheet, "A2"),
            Cell.getValue(sheet, "A3")
          ];
        },
        Word: async () => {
          const { Document } = D.Word;
          const doc = Document.create();
          Document.addParagraph(doc, "Hello floor");
          const back = await Document.read(await Document.toBuffer(doc));
          return JSON.stringify(Document.getContent(back, 0)).includes("Hello floor");
        },
        Pdf: async () => {
          const bytes = await D.Pdf.Pdf.fromTable([
            ["a", "b"],
            ["1", "2"]
          ]);
          return String.fromCharCode(...bytes.slice(0, 5));
        },
        Csv: () => JSON.stringify(D.Csv.Csv.parse("a,b\n1,2")),
        Markdown: () => {
          const { Markdown } = D.Markdown;
          const table = Markdown.parse("| a | b |\n|---|---|\n| 1 | 2 |");
          return Markdown.format(table.headers, table.rows);
        },
        Xml: () => {
          const { Xml } = D.Xml;
          const root = Xml.parse('<r a="1"><c>t</c></r>').root;
          return [Xml.getAttr(root, "a"), Xml.getText(Xml.findChild(root, "c"))];
        },
        Formula: () => D.Formula.Formula.format(D.Formula.Formula.parse("SUM(A1,2)*3")),
        // Compressible, so DEFLATE runs — natively, or by the fallback where CompressionStream is missing.
        Archive: async () => {
          const text = "floor ".repeat(5000);
          const zip = await D.Archive.Archive.packZip(new Map([["a.txt", text]]));
          const entry = (await D.Archive.Archive.unpackZip(zip)).get("a.txt");
          return [entry.info.compressionMethod, new TextDecoder().decode(entry.data) === text];
        },
        Stream: async () => {
          const source = new ReadableStream({
            start(controller) {
              controller.enqueue(new Uint8Array([1, 2]));
              controller.enqueue(new Uint8Array([3]));
              controller.close();
            }
          });
          return Array.from(await D.Stream.Streams.bytes(source));
        },
        // Markup and pixels: the SVG serialiser and the rasteriser are separate backends.
        Draw: () => {
          const { Draw } = D.Draw;
          const list = {
            width: 4,
            height: 4,
            children: [
              {
                kind: "rect",
                x: 0,
                y: 0,
                width: 4,
                height: 4,
                paint: { fill: Draw.parseColor("#ff0000") }
              }
            ]
          };
          return [
            Draw.toSvg(list).includes('fill="#ff0000"'),
            Array.from(Draw.toRgba(list).data.slice(0, 4))
          ];
        },
        Mermaid: () => D.Mermaid.Mermaid.toSvg("graph TD; A-->B").startsWith("<svg"),
        Utils: () => {
          const { Bytes } = D.Utils;
          return [
            Bytes.toBase64(Bytes.fromText("floor")),
            Bytes.toText(Bytes.fromHex("666c6f6f72"))
          ];
        }
      };
      const out = { loaded: Object.keys(D).sort(), results: {} };
      for (const name of Object.keys(checks)) {
        try {
          out.results[name] = await checks[name]();
        } catch (e) {
          out.results[name] = `threw ${e && e.name}: ${e && e.message}`;
        }
      }
      return out;
    });
    const expected = {
      Excel: [
        "_xlfn.LET(_xlpm.x,2,_xlpm.x*3)",
        "_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)(1,2)",
        "Safari"
      ],
      Word: true,
      Pdf: "%PDF-",
      Csv: '[["a","b"],["1","2"]]',
      Markdown: "| a   | b   |\n|:----|:----|\n| 1   | 2   |\n",
      Xml: ["1", "t"],
      Formula: "SUM(A1,2)*3",
      Archive: [8, true],
      Stream: [1, 2, 3],
      Draw: [true, [255, 0, 0, 255]],
      Mermaid: true,
      Utils: ["Zmxvb3I=", "floor"]
    };
    const globals = manifest.bundles.map(bundle => bundle.global).sort();
    if (JSON.stringify(smoke.loaded) !== JSON.stringify(globals)) {
      failures.push(
        `the bundles define Documonster.{${smoke.loaded.join(", ")}}, not Documonster.{${globals.join(", ")}}`
      );
    }
    for (const name of globals) {
      if (!(name in expected)) {
        failures.push(
          `Documonster.${name}: no check in harness.cjs, so nothing shows the module runs`
        );
      } else if (JSON.stringify(smoke.results[name]) !== JSON.stringify(expected[name])) {
        failures.push(
          `Documonster.${name}: expected ${JSON.stringify(expected[name])}, got ${JSON.stringify(smoke.results[name])}`
        );
      }
    }
    console.log(
      `  loaded ${smoke.loaded.length} IIFE bundles and ran each module: ${globals.join(", ")}`
    );
  } finally {
    try {
      if (browser) await browser.close();
    } finally {
      server.close();
    }
  }
  if (failures.length > 0) {
    console.error(`✗ ${ENGINE}: ${failures.length} failure(s)`);
    for (const failure of failures) console.error(`  ${failure}`);
    process.exitCode = 1;
  } else {
    console.log(`✓ ${ENGINE}`);
  }
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
