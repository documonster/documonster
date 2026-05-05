/**
 * Content-types check: missing entries, wrong content types, default
 * overrides.
 */

import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { describe, expect, it } from "vitest";

import { baseParts, buildPackage, contentTypesWith } from "./fixtures";

describe("ooxml-validator / content-types", () => {
  it("flags missing Default for .rels", async () => {
    const parts = baseParts();
    parts["[Content_Types].xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`;
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "content-types-missing-default")).toBe(true);
  });

  it("flags Override pointing to a non-existent part", async () => {
    const parts = baseParts();
    parts["[Content_Types].xml"] = contentTypesWith([
      {
        partName: "/xl/workbook.xml",
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
      },
      {
        partName: "/xl/styles.xml",
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
      },
      {
        partName: "/xl/sharedStrings.xml",
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
      },
      {
        partName: "/xl/theme/theme1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.theme+xml"
      },
      {
        partName: "/xl/worksheets/sheet1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
      },
      { partName: "/xl/nonExistent.xml", contentType: "application/xml" }
    ]);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(
      report.problems.some(
        p => p.kind === "content-types-missing" && p.message.includes("/xl/nonExistent.xml")
      )
    ).toBe(true);
  });

  it("flags part without any content-type mapping", async () => {
    const parts = baseParts();
    // Drop the theme override AND the xml default → theme1.xml has no resolvable content-type.
    parts["[Content_Types].xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`;
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(
      report.problems.some(
        p => p.kind === "content-types-missing-for-part" && p.message.includes("theme1.xml")
      )
    ).toBe(true);
  });

  it("flags wrong content type for a well-known part", async () => {
    const parts = baseParts();
    // Declare workbook with styles content-type.
    parts["[Content_Types].xml"] = contentTypesWith([
      {
        partName: "/xl/workbook.xml",
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
      },
      {
        partName: "/xl/styles.xml",
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
      },
      {
        partName: "/xl/sharedStrings.xml",
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
      },
      {
        partName: "/xl/theme/theme1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.theme+xml"
      },
      {
        partName: "/xl/worksheets/sheet1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
      }
    ]);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(
      report.problems.some(
        p => p.kind === "content-types-wrong-for-part" && p.message.includes("xl/workbook.xml")
      )
    ).toBe(true);
  });

  it("flags duplicate Override PartName", async () => {
    const parts = baseParts();
    parts["[Content_Types].xml"] = contentTypesWith([
      {
        partName: "/xl/workbook.xml",
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
      },
      {
        partName: "/xl/workbook.xml",
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
      },
      {
        partName: "/xl/styles.xml",
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
      },
      {
        partName: "/xl/sharedStrings.xml",
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
      },
      {
        partName: "/xl/theme/theme1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.theme+xml"
      },
      {
        partName: "/xl/worksheets/sheet1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
      }
    ]);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "content-types-duplicate-override")).toBe(true);
  });
});
