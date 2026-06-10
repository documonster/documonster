/**
 * Word Example 40 — Templates: advanced workflows
 *
 * Goes beyond `16-templates.ts` to cover:
 *   - compileTemplate / patchTemplate — pre-parse a .docx once, reuse for
 *     hundreds of patch operations (much faster than patchDocument in a loop).
 *   - fillTemplateFromBuffer — read .docx + fillTemplate + write in one call.
 *   - bindChartData — fill chart series from runtime data (ChartTemplateData).
 *   - CompositeDataSource — combine JSON + XML + CSV sources, with array
 *     merge semantics.
 *
 * Output: tmp/word-examples/40-templates-advanced/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  Document,
  toBuffer,
  compileTemplate,
  patchTemplate,
  fillTemplateFromBuffer,
  fillTemplateFromSource,
  bindChartData,
  JsonDataSource,
  XmlDataSource,
  CsvDataSource,
  CompositeDataSource,
  chart,
  cmToEmu
} from "../index";
import type { ChartBinding } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/40-templates-advanced"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// 1. Build a template .docx with simple {{...}} placeholders
// ---------------------------------------------------------------------------
function buildTemplateBytes(): Promise<Uint8Array> {
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Invoice {{invoiceNo}}", 1);
  Document.addParagraph(d, "Customer: {{customer}}");
  Document.addParagraph(d, "Total: {{total}}");
  return toBuffer(Document.build(d));
}
const templateBytes = await buildTemplateBytes();
fs.writeFileSync(path.join(outDir, "00-template.docx"), templateBytes);

// ---------------------------------------------------------------------------
// 2. compileTemplate + patchTemplate — bulk fill in a tight loop
// ---------------------------------------------------------------------------
const compiled = await compileTemplate(templateBytes);
const t0 = performance.now();
for (let i = 1; i <= 5; i++) {
  const patched = await patchTemplate(compiled, [
    { placeholder: "{{invoiceNo}}", content: { type: "text", text: `INV-${1000 + i}` } },
    { placeholder: "{{customer}}", content: { type: "text", text: `Customer #${i}` } },
    { placeholder: "{{total}}", content: { type: "text", text: `$${i * 100}` } }
  ]);
  fs.writeFileSync(path.join(outDir, `01-bulk-${i}.docx`), patched);
}
console.log(
  `  patchTemplate × 5: ${(performance.now() - t0).toFixed(0)} ms (parsing happens once)`
);

// ---------------------------------------------------------------------------
// 3. fillTemplateFromBuffer — single-shot fill using {{var}} placeholders
// ---------------------------------------------------------------------------
const filled = await fillTemplateFromBuffer(templateBytes, {
  invoiceNo: "INV-9999",
  customer: "From-Buffer Co.",
  total: "$1,234.56"
});
fs.writeFileSync(path.join(outDir, "02-from-buffer.docx"), filled);
console.log(`  → 02-from-buffer.docx (${filled.length} bytes)`);

// ---------------------------------------------------------------------------
// 4. CompositeDataSource — JSON + XML + CSV combined, array merge
// ---------------------------------------------------------------------------
const json = new JsonDataSource({
  invoiceNo: "JSON-1",
  customer: "Composite Co.",
  notes: ["from json"]
});
const xml = new XmlDataSource(`<?xml version="1.0"?>
<root>
  <total>$5,000</total>
  <notes>from xml</notes>
</root>`);
const csv = new CsvDataSource("name,price\nWidget,10\nGadget,25", { rowsKey: "lineItems" });

// keep each source's arrays distinct where keys collide (mergeArrays: false)
const composite = new CompositeDataSource([json, xml, csv], { mergeArrays: false });
console.log(`  composite.getData() keys: ${Object.keys(composite.getData()).join(", ")}`);

const compositeFilled = fillTemplateFromSource(compiled._doc, composite, { strict: false });
const compositeBytes = await toBuffer(compositeFilled);
fs.writeFileSync(path.join(outDir, "03-composite.docx"), compositeBytes);
console.log(`  → 03-composite.docx (${compositeBytes.length} bytes)`);

// ---------------------------------------------------------------------------
// 5. bindChartData — replace a chart's series at runtime
// ---------------------------------------------------------------------------
const chartTemplate = (() => {
  const dd = Document.create();
  Document.useDefaultStyles(dd);
  Document.addHeading(dd, "Quarterly figures", 1);
  Document.addContent(
    dd,
    chart({
      type: "column",
      title: "TEMPLATE TITLE",
      width: cmToEmu(15),
      height: cmToEmu(8),
      legend: "b",
      series: [
        {
          name: "Series 1",
          categories: ["Q1", "Q2", "Q3", "Q4"],
          values: [0, 0, 0, 0],
          color: "4472C4"
        }
      ]
    })
  );
  return Document.build(dd);
})();

const bindings: ChartBinding[] = [
  {
    chartRef: 0, // first chart in the body
    title: "Actual Q4 results",
    categories: ["Q1", "Q2", "Q3", "Q4"],
    series: [
      { name: "North", values: [120, 140, 160, 190], color: "4472C4" },
      { name: "South", values: [90, 110, 130, 150], color: "ED7D31" }
    ]
  }
];
const boundDoc = bindChartData(chartTemplate, bindings);
fs.writeFileSync(path.join(outDir, "04-chart-bound.docx"), await toBuffer(boundDoc));
console.log(`  → 04-chart-bound.docx`);

// ---------------------------------------------------------------------------
// 6. Edge case: bindChartData with mismatched chart count is a no-op
// ---------------------------------------------------------------------------
const noChartDoc = (() => {
  const dd = Document.create();
  Document.useDefaultStyles(dd);
  Document.addParagraph(dd, "no charts here");
  return Document.build(dd);
})();
const stillNoChart = bindChartData(noChartDoc, bindings);
console.log(`  bindChartData on no-chart doc: body length unchanged=${stillNoChart.body.length}`);
