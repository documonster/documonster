#!/usr/bin/env node
/**
 * Compatibility report scaffold.
 *
 * This script is the deterministic scaffold for the compatibility
 * matrix published in `docs/COMPATIBILITY.md`. Running it without any
 * environment variables emits the **static matrix** — the same one
 * that lives in `src/modules/excel/README.md` — as a standalone
 * Markdown document, which keeps `docs/COMPATIBILITY.md` in source
 * control and easy to diff.
 *
 * Setting `DOCUMONSTER_ENTERPRISE_CORPUS_DIR=/path` additionally appends a
 * per-fixture section summarising load / audit / LibreOffice
 * open-validation results for the private corpus. That path runs the
 * same workbook round-trip the oracle integration test uses; see
 * `src/modules/excel/__tests__/helpers/ooxml-package-audit.ts` and
 * `src/modules/excel/__tests__/helpers/external-oracle.ts` for the
 * underlying checks.
 *
 * The corpus section is a thin wrapper around
 * `chart-oracle.integration.test.ts` — deliberately so, so there is no
 * second implementation of the round-trip logic.
 *
 * Run:
 *
 *   node scripts/compatibility-report.ts
 *
 *   DOCUMONSTER_ENTERPRISE_CORPUS_DIR=/path/to/corpus \
 *     DOCUMONSTER_LIBREOFFICE_OPEN_VALIDATION=1 \
 *     LIBREOFFICE_BIN=/usr/bin/soffice \
 *     node scripts/compatibility-report.ts
 */
import { existsSync } from "node:fs";
import { mkdir, readdir, stat, writeFile } from "node:fs/promises";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(__dirname, "..");
const DEFAULT_OUT = resolve(REPO_ROOT, "docs/COMPATIBILITY.md");

interface TypeRow {
  type: string;
  create: "yes" | "no" | "partial";
  read: "yes" | "no" | "partial";
  edit: "yes" | "no" | "partial";
  roundTrip: "yes" | "no" | "via-preset";
  rawPreserve: "yes" | "no";
  svg: "hash" | "generic" | "none";
  png: "hash" | "generic" | "none";
  pdf: "type-specific" | "generic" | "vector" | "raster-only" | "none";
  libreoffice: "direct" | "via-preset" | "none";
}

interface FeatureRow {
  feature: string;
  status: "yes" | "partial" | "no";
  note: string;
}

interface LibraryRow {
  library: string;
  language: string;
  classicCreate: "full" | "partial" | "none";
  chartEx: "full" | "partial" | "none";
  combo: "yes" | "partial" | "no";
  pivotChart: "yes" | "partial" | "no";
  chartsheet: "yes" | "partial" | "no" | "none";
  editLoaded: "full" | "partial" | "none";
  preview: "svg+png+pdf" | "partial" | "none";
  rawXmlPreserve: "byte" | "partial" | "rebuild";
  ecosystem: "node-browser" | "node" | "python" | "java" | "go" | "dotnet" | "any";
}

const CLASSIC_ROWS: TypeRow[] = [
  {
    type: "bar",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "type-specific",
    libreoffice: "direct"
  },
  {
    type: "bar3D",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "type-specific",
    libreoffice: "via-preset"
  },
  {
    type: "line",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "generic",
    libreoffice: "direct"
  },
  {
    type: "line3D",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "via-preset",
    rawPreserve: "yes",
    svg: "generic",
    png: "generic",
    pdf: "generic",
    libreoffice: "via-preset"
  },
  {
    type: "pie",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "type-specific",
    libreoffice: "via-preset"
  },
  {
    type: "pie3D",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "via-preset",
    rawPreserve: "yes",
    svg: "generic",
    png: "generic",
    pdf: "generic",
    libreoffice: "via-preset"
  },
  {
    type: "doughnut",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "generic",
    libreoffice: "via-preset"
  },
  {
    type: "area",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "type-specific",
    libreoffice: "via-preset"
  },
  {
    type: "area3D",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "via-preset",
    rawPreserve: "yes",
    svg: "generic",
    png: "generic",
    pdf: "generic",
    libreoffice: "via-preset"
  },
  {
    type: "scatter",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "generic",
    libreoffice: "via-preset"
  },
  {
    type: "bubble",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "type-specific",
    libreoffice: "via-preset"
  },
  {
    type: "radar",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "type-specific",
    libreoffice: "via-preset"
  },
  {
    type: "stock",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "via-preset",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "generic",
    libreoffice: "via-preset"
  },
  {
    type: "surface",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "generic",
    libreoffice: "via-preset"
  },
  {
    type: "surface3D",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "via-preset",
    rawPreserve: "yes",
    svg: "generic",
    png: "generic",
    pdf: "generic",
    libreoffice: "via-preset"
  },
  {
    type: "ofPie",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "via-preset",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "generic",
    libreoffice: "via-preset"
  }
];

const CHARTEX_ROWS: TypeRow[] = [
  {
    type: "sunburst",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "vector",
    libreoffice: "via-preset"
  },
  {
    type: "treemap",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "vector",
    libreoffice: "direct"
  },
  {
    type: "waterfall",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "vector",
    libreoffice: "via-preset"
  },
  {
    type: "funnel",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "vector",
    libreoffice: "direct"
  },
  {
    type: "histogram",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "vector",
    libreoffice: "via-preset"
  },
  {
    type: "pareto",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "vector",
    libreoffice: "via-preset"
  },
  {
    type: "boxWhisker",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "vector",
    libreoffice: "via-preset"
  },
  {
    type: "regionMap",
    create: "yes",
    read: "yes",
    edit: "yes",
    roundTrip: "yes",
    rawPreserve: "yes",
    svg: "hash",
    png: "generic",
    pdf: "vector",
    libreoffice: "via-preset"
  }
];

/**
 * Cross-cutting chart features that span every chart type — so they
 * don't belong on the per-type grid. Each feature has its own note
 * pointing at the API entry point and the test that exercises it.
 */
const FEATURE_ROWS: FeatureRow[] = [
  {
    feature: "Combo charts",
    status: "yes",
    note: "`worksheet.addComboChart({ groups: [...] })`; primary / secondary axes; see `AddComboChartOptions`"
  },
  {
    feature: "Secondary axes",
    status: "yes",
    note: "`secondaryXAxis` / `secondaryYAxis` options on `addChart`; axis picker at scene / render level"
  },
  {
    feature: "Pivot chart metadata",
    status: "yes",
    note: "`worksheet.addPivotChart(pivotTable, options, range)` — pivotSource, field buttons, `c14:pivotOptions`, `c16:pivotOptions16`; rendering is still Excel's responsibility"
  },
  {
    feature: "Chartsheet (single-chart sheet)",
    status: "yes",
    note: "`workbook.addChartsheet({ chart })` / `addPivotChartsheet`; tabSelected, state, pageSetup preserved"
  },
  {
    feature: "Data table (`c:dTable`)",
    status: "yes",
    note: "Round-trips through XML, rendered below the plot area in SVG/PNG/PDF as of the dataTable-rendering milestone (showOutline/showHorzBorder/showVertBorder/showKeys honoured)"
  },
  {
    feature: "bar3D true 3D projection",
    status: "yes",
    note: "`view3D.rotX` / `rotY` drive an axonometric projection in the preview renderer; three shaded faces (top + front + right) per bar. Other 3D variants still preview as 2D — see the 3D note"
  },
  {
    feature: "User-shape overlays (`c:userShapes`)",
    status: "yes",
    note: "`Chart.userShapesXml` / `setUserShapesXml(xml)` / `removeUserShapes()` — byte-preserving round-trip + programmatic replacement. The DrawingML shape model itself stays opaque"
  },
  {
    feature: "ChartEx helper APIs",
    status: "yes",
    note: "`chartExOptionsFromTable(ws, table, options)` / `chartExOptionsFromRows(ws, rows, options)` + `worksheet.addChartExFromTable` / `addChartExFromRows` — mirror classic helpers for sunburst/treemap/waterfall/funnel/histogram/pareto/boxWhisker. `regionMap` intentionally excluded (needs geographic labels)"
  },
  {
    feature: "Unknown-element surfacing",
    status: "yes",
    note: '`Chart.unknownElements` returns `c15:` / `cx14:` vendor tags observed at parse time; `templateMode: "strict"` fails the write when a rebuild would drop them'
  },
  {
    feature: "Strict template mode",
    status: "yes",
    note: '`workbook.xlsx.writeBuffer({ templateMode: "strict" })` / `{ strictTemplateMode: true }` — refuses any rebuild that would silently drop vendor XML'
  },
  {
    feature: "Raw-XML patching",
    status: "yes",
    note: "`chart.mutate(fn, { preferRawPatch: true })` applies surgical byte patches for narrow edits (title, series value ref, grouping flags) so unknown XML stays intact"
  },
  {
    feature: "Byte-preserving round-trip",
    status: "yes",
    note: "Clean loads write back byte-identical chart XML; edits that can't raw-patch rebuild the chart structurally with every known element preserved"
  },
  {
    feature: "Built-in chart styles 1–48",
    status: "yes",
    note: "`chart.setStyle(n)` / `setBuiltInStyle(n)` — matches xlsxwriter `chart.set_style(n)` semantics"
  },
  {
    feature: "Modern chartStyle/chartColors sidecars",
    status: "yes",
    note: "`chartStyle: ChartStyleModel` / `chartColors: ChartColorsModel` on `addChart` write `styleN.xml` / `colorsN.xml` + rels + content-types"
  },
  {
    feature: "Trendlines",
    status: "yes",
    note: "6 types: exp / linear / log / movingAvg / poly / power; `forward` / `backward` / `dispRSqr` / `dispEq` / `intercept` / `trendlineLab`"
  },
  {
    feature: "Error bars",
    status: "yes",
    note: "`errBarType`: both/minus/plus; `errValType`: cust/fixedVal/percentage/stdDev/stdErr; per-series plus/minus custom refs"
  },
  {
    feature: "Data labels (per-point override)",
    status: "yes",
    note: "All `DataLabelPosition` values, `DataLabelEntry[]` for per-point overrides, `DataLabelsRange` for formula-driven labels, pie leader lines with collision avoidance"
  },
  {
    feature: "Gradient / pattern / picture fills",
    status: "yes",
    note: "Linear + path gradients; 48 preset patterns; `ChartBlipFill.image` auto-registers media + rels"
  }
];

/**
 * How Documonster stacks up against the leading alternatives. The matrix
 * intentionally uses lossy categories ("full / partial / none") because
 * each library's chart surface is big enough that detailed rows would
 * take an entire page — the accompanying migration guides
 * (`docs/FROM_*.md`) carry the detail.
 */
const LIBRARY_ROWS: LibraryRow[] = [
  {
    library: "documonster",
    language: "TypeScript / JavaScript",
    classicCreate: "full",
    chartEx: "full",
    combo: "yes",
    pivotChart: "yes",
    chartsheet: "yes",
    editLoaded: "full",
    preview: "svg+png+pdf",
    rawXmlPreserve: "byte",
    ecosystem: "node-browser"
  },
  {
    library: "ExcelJS",
    language: "JavaScript",
    classicCreate: "none",
    chartEx: "none",
    combo: "no",
    pivotChart: "no",
    chartsheet: "no",
    editLoaded: "none",
    preview: "none",
    rawXmlPreserve: "partial",
    ecosystem: "node-browser"
  },
  {
    library: "SheetJS (xlsx)",
    language: "JavaScript",
    classicCreate: "none",
    chartEx: "none",
    combo: "no",
    pivotChart: "no",
    chartsheet: "none" as const,
    editLoaded: "none",
    preview: "none",
    rawXmlPreserve: "partial",
    ecosystem: "node-browser"
  },
  {
    library: "xlsxwriter",
    language: "Python (write-only)",
    classicCreate: "full",
    chartEx: "none",
    combo: "yes",
    pivotChart: "no",
    chartsheet: "yes",
    editLoaded: "none",
    preview: "none",
    rawXmlPreserve: "rebuild",
    ecosystem: "python"
  },
  {
    library: "openpyxl",
    language: "Python",
    classicCreate: "full",
    chartEx: "none",
    combo: "partial",
    pivotChart: "no",
    chartsheet: "yes",
    editLoaded: "partial",
    preview: "none",
    rawXmlPreserve: "rebuild",
    ecosystem: "python"
  },
  {
    library: "excelize",
    language: "Go",
    classicCreate: "full",
    chartEx: "none",
    combo: "partial",
    pivotChart: "no",
    chartsheet: "partial",
    editLoaded: "partial",
    preview: "none",
    rawXmlPreserve: "partial",
    ecosystem: "go"
  },
  {
    library: "Apache POI",
    language: "Java",
    classicCreate: "full",
    chartEx: "partial",
    combo: "yes",
    pivotChart: "partial",
    chartsheet: "yes",
    editLoaded: "full",
    preview: "none",
    rawXmlPreserve: "partial",
    ecosystem: "java"
  },
  {
    library: "EPPlus",
    language: "C# / .NET",
    classicCreate: "full",
    chartEx: "none",
    combo: "yes",
    pivotChart: "partial",
    chartsheet: "yes",
    editLoaded: "full",
    preview: "none",
    rawXmlPreserve: "partial",
    ecosystem: "dotnet"
  },
  {
    library: "ClosedXML",
    language: "C# / .NET",
    classicCreate: "partial",
    chartEx: "none",
    combo: "partial",
    pivotChart: "no",
    chartsheet: "no",
    editLoaded: "partial",
    preview: "none",
    rawXmlPreserve: "partial",
    ecosystem: "dotnet"
  },
  {
    library: "Aspose.Cells",
    language: "multi (paid)",
    classicCreate: "full",
    chartEx: "full",
    combo: "yes",
    pivotChart: "yes",
    chartsheet: "yes",
    editLoaded: "full",
    preview: "svg+png+pdf",
    rawXmlPreserve: "byte",
    ecosystem: "any"
  }
];

const SYMBOL: Record<string, string> = {
  yes: "✅",
  no: "❌",
  partial: "⚠️",
  "via-preset": "⬛",
  hash: "✅",
  generic: "⬛",
  "type-specific": "✅",
  "raster-only": "🟨",
  vector: "✅",
  direct: "✅",
  none: "➖",
  full: "✅",
  byte: "✅",
  rebuild: "❌",
  "svg+png+pdf": "✅"
};

function renderTypeTable(title: string, rows: TypeRow[]): string {
  const header =
    "| Type | Create | Read | Edit | Round-trip | Raw preserve | SVG | PNG | PDF | LibreOffice |";
  const separator =
    "| ---- | :----: | :--: | :--: | :--------: | :----------: | :-: | :-: | :-: | :---------: |";
  const body = rows
    .map(
      r =>
        `| ${r.type} | ${SYMBOL[r.create]} | ${SYMBOL[r.read]} | ${SYMBOL[r.edit]} | ${SYMBOL[r.roundTrip]} | ${SYMBOL[r.rawPreserve]} | ${SYMBOL[r.svg]} | ${SYMBOL[r.png]} | ${SYMBOL[r.pdf]} | ${SYMBOL[r.libreoffice]} |`
    )
    .join("\n");
  return `## ${title}\n\n${header}\n${separator}\n${body}\n`;
}

function renderFeatureTable(title: string, rows: FeatureRow[]): string {
  const header = "| Feature | Status | Notes |";
  const separator = "| ------- | :----: | ----- |";
  const body = rows.map(r => `| ${r.feature} | ${SYMBOL[r.status]} | ${r.note} |`).join("\n");
  return `## ${title}\n\n${header}\n${separator}\n${body}\n`;
}

function renderLibraryTable(title: string, rows: LibraryRow[]): string {
  const header =
    "| Library | Language | Classic create | ChartEx | Combo | Pivot chart | Chartsheet | Edit loaded | Preview (SVG/PNG/PDF) | Raw-XML preserve |";
  const separator =
    "| ------- | -------- | :------------: | :-----: | :---: | :---------: | :--------: | :---------: | :-------------------: | :--------------: |";
  const body = rows
    .map(r => {
      const libCell = r.library === "documonster" ? `**${r.library}**` : r.library;
      return `| ${libCell} | ${r.language} | ${SYMBOL[r.classicCreate]} | ${SYMBOL[r.chartEx]} | ${SYMBOL[r.combo]} | ${SYMBOL[r.pivotChart]} | ${SYMBOL[r.chartsheet]} | ${SYMBOL[r.editLoaded]} | ${SYMBOL[r.preview]} | ${SYMBOL[r.rawXmlPreserve]} |`;
    })
    .join("\n");
  return `## ${title}\n\n${header}\n${separator}\n${body}\n`;
}

async function discoverCorpus(root: string): Promise<string[]> {
  const files: string[] = [];
  const walk = async (dir: string): Promise<void> => {
    for (const name of await readdir(dir)) {
      const full = join(dir, name);
      const info = await stat(full);
      if (info.isDirectory()) {
        await walk(full);
      } else if (info.isFile() && /\.xlsx$/i.test(name)) {
        files.push(relative(root, full));
      }
    }
  };
  await walk(root);
  return files.sort();
}

async function renderCorpusSection(root: string): Promise<string> {
  const files = await discoverCorpus(root);
  const lines = [`## Enterprise corpus (${files.length} fixtures)\n`];
  lines.push(`Discovered under: \`${root}\`\n`);
  lines.push("| File | Size |");
  lines.push("| ---- | ---: |");
  for (const rel of files) {
    const full = join(root, rel);
    const info = await stat(full);
    lines.push(`| \`${rel}\` | ${(info.size / 1024).toFixed(1)} KB |`);
  }
  lines.push("");
  lines.push(
    "_Full audit / LibreOffice open-validation runs live in" +
      " `src/modules/excel/__tests__/chart-oracle.integration.test.ts`." +
      " Set `DOCUMONSTER_ENTERPRISE_CORPUS_DIR` + `DOCUMONSTER_LIBREOFFICE_OPEN_VALIDATION=1`" +
      " when running the test to produce pass/fail decisions;" +
      " this scaffold only lists presence so `docs/COMPATIBILITY.md` stays small._"
  );
  lines.push("");
  return lines.join("\n");
}

async function main(): Promise<void> {
  const outPath = process.env.DOCUMONSTER_COMPATIBILITY_REPORT_OUT
    ? resolve(process.env.DOCUMONSTER_COMPATIBILITY_REPORT_OUT)
    : DEFAULT_OUT;
  const lines: string[] = [];
  lines.push("# Documonster chart compatibility report");
  lines.push("");
  lines.push(`Generated by \`scripts/compatibility-report.ts\` on ${new Date().toISOString()}.`);
  lines.push("");
  lines.push(
    "This document mirrors the per-type capability grid in" +
      " [`src/modules/excel/README.md`](../src/modules/excel/README.md)" +
      " and extends it with a cross-cutting feature inventory and a" +
      " side-by-side comparison against the leading open-source and" +
      " commercial chart libraries. It also appends the current" +
      " enterprise-corpus inventory when the" +
      " `DOCUMONSTER_ENTERPRISE_CORPUS_DIR` environment variable is set." +
      " Run `node scripts/compatibility-report.ts` to regenerate."
  );
  lines.push("");
  lines.push("## Legend");
  lines.push("");
  lines.push("Per-type grid symbols:");
  lines.push("");
  lines.push("- ✅ direct type-specific test / full support");
  lines.push("- ⬛ exercised via generic / preset-scan loop (no value-level assert)");
  lines.push("- 🟨 raster-only PDF path (no longer used — all ChartEx goes vector)");
  lines.push("- ⚠️ partial / caveats apply (see notes)");
  lines.push("- ➖ not implemented / not applicable");
  lines.push("- ❌ not supported");
  lines.push("");
  lines.push(
    "Cross-library comparison uses the same symbols, where ✅ means" +
      " a first-class feature with structured API, ⚠️ means it works" +
      " but with caveats (partial fields, raw-XML only, metadata-only)," +
      " and ❌/➖ mean no structured support. See the migration guides" +
      " in `docs/FROM_*.md` for the concrete feature deltas."
  );
  lines.push("");
  lines.push("## Rendering scope");
  lines.push("");
  lines.push(
    "Documonster ships a **zero-dependency deterministic preview renderer**" +
      " for SVG / PNG / PDF, not an Excel-pixel-perfect compositor." +
      " The preview is driven by the same `ChartScene` intermediate" +
      " representation on all three backends, so what you see in SVG" +
      " is what the PDF surface emits (modulo rasterisation) and what" +
      " the Node PNG rasteriser paints. It is suitable for thumbnails," +
      " email attachments, server-side report generation and quick" +
      " sanity checks; **it is not a replacement for Excel or LibreOffice**" +
      " when pixel-identical output matters. For production-grade" +
      " rendering, round-trip the `.xlsx` through headless LibreOffice" +
      " (`soffice --convert-to pdf`) — the metadata preservation" +
      " guarantees in this library make that a byte-safe handoff."
  );
  lines.push("");
  lines.push(renderTypeTable("Classic charts", CLASSIC_ROWS));
  lines.push(renderTypeTable("ChartEx types", CHARTEX_ROWS));
  lines.push(renderFeatureTable("Cross-cutting features", FEATURE_ROWS));
  lines.push(renderLibraryTable("Cross-library comparison", LIBRARY_ROWS));
  const corpusRoot = process.env.DOCUMONSTER_ENTERPRISE_CORPUS_DIR;
  if (corpusRoot && existsSync(corpusRoot)) {
    lines.push(await renderCorpusSection(resolve(corpusRoot)));
  } else {
    lines.push("## Enterprise corpus");
    lines.push("");
    lines.push(
      "_No corpus directory supplied._ Set `DOCUMONSTER_ENTERPRISE_CORPUS_DIR`" +
        " to a directory containing real `.xlsx` files authored by Excel," +
        " WPS Office, LibreOffice or Aspose.Cells and rerun this script to" +
        " emit a per-fixture inventory. An example manifest is at" +
        " [`docs/enterprise-corpus-manifest.example.json`](./enterprise-corpus-manifest.example.json)."
    );
    lines.push("");
  }
  await mkdir(dirname(outPath), { recursive: true });
  await writeFile(outPath, lines.join("\n"), "utf8");
  // eslint-disable-next-line no-console
  console.log(`Wrote ${outPath}`);
}

main().catch(err => {
  // eslint-disable-next-line no-console
  console.error(err);
  process.exit(1);
});
