import { existsSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { join, resolve } from "node:path";

import { extractAll } from "@archive/unzip/extract";
import { Workbook } from "@excel/workbook";
import { describe, expect, it } from "vitest";

import {
  discoverEnterpriseCorpus,
  loadEnterpriseCorpusManifest
} from "./helpers/enterprise-corpus";
import { runExternalOracle, runOfficeOpenValidation } from "./helpers/external-oracle";
import { auditOoxmlPackage } from "./helpers/ooxml-package-audit";

describe("optional chart oracle and enterprise corpus harness", () => {
  it("optionally exports chart workbooks through LibreOffice as a visual oracle hook", async () => {
    const workbook = new Workbook();
    const ws = workbook.addWorksheet("Data");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    ws.addChart(
      { type: "bar", series: [{ categories: "Data!$A$1:$A$2", values: "Data!$B$1:$B$2" }] },
      "D1:J10"
    );
    const input = new Uint8Array(await workbook.xlsx.writeBuffer());
    const zip = await extractAll(input);
    const audit = auditOoxmlPackage(zip);

    expect(audit.errors, audit.errors.join("\n")).toEqual([]);

    const result = await runExternalOracle({
      envFlag: "EXCELTS_LIBREOFFICE_VISUAL_ORACLE",
      executableEnv: "LIBREOFFICE_BIN",
      candidates: ["soffice", "libreoffice"],
      args: ["--headless", "--convert-to", "pdf", "--outdir", "{outDir}", "{input}"],
      input,
      inputName: "chart-oracle.xlsx",
      outputGlob: /\.pdf$/i
    });

    if (!result.available) {
      expect(result.skipped).toBeTruthy();
      return;
    }
    expect(result.outputs.length).toBeGreaterThan(0);
    expect(result.outputs[0].data.length).toBeGreaterThan(100);
  });

  it("optionally open-validates generated chart workbooks with Office-compatible binaries", async () => {
    const workbook = new Workbook();
    const ws = workbook.addWorksheet("Data");
    ws.addRows([
      ["A", 10, 1],
      ["B", 20, 2],
      ["C", 30, 3]
    ]);
    ws.addComboChart(
      {
        groups: [
          {
            type: "bar",
            series: [{ name: "Revenue", categories: "Data!$A$1:$A$3", values: "Data!$B$1:$B$3" }]
          },
          {
            type: "line",
            useSecondaryAxis: true,
            series: [{ name: "Growth", categories: "Data!$A$1:$A$3", values: "Data!$C$1:$C$3" }]
          }
        ]
      },
      "E1:L12"
    );
    ws.addChartEx(
      { type: "treemap", categories: "Data!$A$1:$A$3", series: [{ values: "Data!$B$1:$B$3" }] },
      "E14:L25"
    );
    workbook.addChartsheet("Chart Sheet", {
      chart: {
        type: "funnel",
        categories: "Data!$A$1:$A$3",
        series: [{ values: "Data!$B$1:$B$3" }]
      }
    });

    const input = new Uint8Array(await workbook.xlsx.writeBuffer());
    const audit = auditOoxmlPackage(await extractAll(input));
    expect(audit.errors, audit.errors.join("\n")).toEqual([]);

    const libreOffice = await runOfficeOpenValidation({
      envFlag: "EXCELTS_LIBREOFFICE_OPEN_VALIDATION",
      executableEnv: "LIBREOFFICE_BIN",
      candidates: ["soffice", "libreoffice"],
      input,
      inputName: "chart-open-validation.xlsx"
    });
    if (libreOffice.available) {
      expect(libreOffice.exitCode).toBe(0);
      expect(libreOffice.outputs.length).toBeGreaterThan(0);
    } else {
      expect(libreOffice.skipped).toBeTruthy();
    }

    const office = await runOfficeOpenValidation({
      envFlag: "EXCELTS_OFFICE_OPEN_VALIDATION",
      executableEnv: "EXCEL_OFFICE_BIN",
      candidates: [],
      args: process.env.EXCELTS_OFFICE_OPEN_ARGS
        ? process.env.EXCELTS_OFFICE_OPEN_ARGS.split(" ").filter(Boolean)
        : undefined,
      versionArgs: process.env.EXCELTS_OFFICE_VERSION_ARGS
        ? process.env.EXCELTS_OFFICE_VERSION_ARGS.split(" ").filter(Boolean)
        : false,
      input,
      inputName: "chart-open-validation.xlsx"
    });
    if (office.available) {
      expect(office.exitCode).toBe(0);
    } else {
      expect(office.skipped).toBeTruthy();
    }
  });

  it("optionally runs configured enterprise chart corpus round-trips", async () => {
    const rootValue = process.env.EXCELTS_ENTERPRISE_CORPUS_DIR;
    if (!rootValue) {
      expect("Set EXCELTS_ENTERPRISE_CORPUS_DIR to enable.").toBeTruthy();
      return;
    }
    const root = resolve(rootValue);
    const manifestPath = process.env.EXCELTS_ENTERPRISE_CORPUS_MANIFEST
      ? resolve(process.env.EXCELTS_ENTERPRISE_CORPUS_MANIFEST)
      : join(root, "manifest.json");
    const entries = existsSync(manifestPath)
      ? (await loadEnterpriseCorpusManifest(manifestPath)).entries
      : await discoverEnterpriseCorpus(root);

    expect(entries.length).toBeGreaterThan(0);
    for (const entry of entries) {
      const input = await readFile(join(root, entry.path));
      const wb = new Workbook();
      await wb.xlsx.load(input);
      const output = await wb.xlsx.writeBuffer();
      const zip = await extractAll(new Uint8Array(output));
      const audit = auditOoxmlPackage(zip);

      expect(audit.errors, `${entry.path}\n${audit.errors.join("\n")}`).toEqual([]);
      expect(zip.get("xl/workbook.xml")).toBeDefined();
      if (entry.expectCharts) {
        expect([...zip.keys()].some(path => /^xl\/charts\/chart\d+\.xml$/.test(path))).toBe(true);
      }
      if (entry.expectChartEx) {
        expect([...zip.keys()].some(path => /^xl\/charts\/chartEx\d+\.xml$/.test(path))).toBe(true);
      }
      if (entry.expectPivotTables) {
        expect(
          [...zip.keys()].some(path => /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(path))
        ).toBe(true);
      }

      if (entry.openValidation || process.env.EXCELTS_CORPUS_LIBREOFFICE_OPEN_VALIDATION === "1") {
        const opened = await runOfficeOpenValidation({
          envFlag: "EXCELTS_CORPUS_LIBREOFFICE_OPEN_VALIDATION",
          executableEnv: "LIBREOFFICE_BIN",
          candidates: ["soffice", "libreoffice"],
          input: new Uint8Array(output),
          inputName: entry.path.replace(/[^A-Za-z0-9_.-]/g, "_")
        });
        if (opened.available) {
          expect(opened.exitCode).toBe(0);
          expect(opened.outputs.length).toBeGreaterThan(0);
        } else {
          expect(opened.skipped).toBeTruthy();
        }
      }
    }
  });
});
