/**
 * Generic Workbook Round-trip Test
 *
 * This test verifies that reading and writing an Excel file preserves all critical data.
 * It uses atomic-level exclusions - only specific known differences are allowed.
 *
 * Each exclusion is documented with its reason.
 */

import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

import { Workbook } from "@excel/index";
import { describe, it, expect, beforeAll } from "vitest";

import { ZipParser } from "../../archive/unzip/zip-parser";
import { expectValidXlsx } from "./helpers/expect-valid-xlsx";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Test file path
const SAMPLE_FILE_PATH = path.join(__dirname, "data/workbook-roundtrip.xlsx");

/**
 * Atomic exclusion rules for XML comparison.
 * Each rule documents exactly what difference is allowed and why.
 */
interface AtomicExclusion {
  /** Description of what this exclusion allows */
  reason: string;
  /** Function to normalize the difference */
  normalize: (content: string) => string;
}

/**
 * All allowed atomic differences with their reasons.
 * IMPORTANT: Do not add exclusions without clear justification.
 */
const ATOMIC_EXCLUSIONS: AtomicExclusion[] = [
  // ============================================
  // rId Renumbering (relationship IDs)
  // ============================================
  {
    reason: "rId numbers may be reordered, but targets must match",
    normalize: s => s.replace(/Id="rId\d+"/g, 'Id="rId#"')
  },
  {
    reason: "r:id references may be renumbered",
    normalize: s => s.replace(/r:id="rId\d+"/g, 'r:id="rId#"')
  },

  // ============================================
  // Self-closing tag whitespace (must run early)
  // ============================================
  {
    reason: "Self-closing tag space: <tag /> vs <tag/>",
    normalize: s => s.replace(/ \/>/g, "/>")
  },

  // ============================================
  // XML Declaration & Namespace ordering
  // ============================================
  {
    reason: "XML namespace declarations may be reordered or removed",
    normalize: s => s.replace(/xmlns:[a-z0-9]+="[^"]+"\s*/g, "")
  },
  {
    reason: "mc:Ignorable attribute may be removed",
    normalize: s => s.replace(/mc:Ignorable="[^"]*"\s*/g, "")
  },

  // ============================================
  // UUID/GUID attributes (session-specific)
  // ============================================
  {
    reason: "xr:uid is a session-specific UUID",
    normalize: s => s.replace(/xr:uid="\{[^}]+\}"/g, "")
  },
  {
    reason: "xr2:uid is a session-specific UUID",
    normalize: s => s.replace(/xr2:uid="\{[^}]+\}"/g, "")
  },

  // ============================================
  // Workbook version metadata (non-functional)
  // ============================================
  {
    reason: "fileVersion attributes are application metadata",
    normalize: s =>
      s
        .replace(/<fileVersion[^>]*\/>/g, "<fileVersion/>")
        .replace(/<fileVersion[^>]*>/g, "<fileVersion>")
  },
  {
    reason: "calcPr calcId is Excel version specific",
    normalize: s => s.replace(/calcId="\d+"/g, 'calcId="#"')
  },
  {
    reason: "workbookPr attributes vary by Excel version",
    normalize: s =>
      s
        .replace(/defaultThemeVersion="\d+"/g, "")
        .replace(/filterPrivacy="1"/g, "")
        .replace(/<workbookPr\s*\/>/g, "<workbookPr/>")
        .replace(/<workbookPr\s+>/g, "<workbookPr>")
  },

  // ============================================
  // Extension elements (extLst)
  // ============================================
  {
    reason: "mc:AlternateContent contains local file paths",
    normalize: s => s.replace(/<mc:AlternateContent[\s\S]*?<\/mc:AlternateContent>/g, "")
  },
  {
    reason: "xr:revisionPtr is revision tracking metadata",
    normalize: s => s.replace(/<xr:revisionPtr[^>]*\/>/g, "")
  },
  {
    reason: "extLst extensions may not be preserved (known limitation)",
    normalize: s =>
      s.replace(/<extLst>[\s\S]*?<\/extLst>/g, "").replace(/<ext [^>]*>[\s\S]*?<\/ext>/g, "")
  },

  // ============================================
  // Style-related differences (known limitation)
  // ============================================
  {
    reason: "Font count may differ due to font normalization",
    normalize: s => s.replace(/<fonts count="\d+"/g, '<fonts count="#"')
  },
  {
    reason: "cellXfs count may differ due to style rebuild",
    normalize: s => s.replace(/<cellXfs count="\d+"/g, '<cellXfs count="#"')
  },
  {
    reason: "Style index (s attribute) may be renumbered",
    normalize: s => s.replace(/\ss="\d+"/g, ' s="#"')
  },

  // ============================================
  // Pivot Cache metadata (non-critical)
  // ============================================
  {
    reason: "refreshedBy is user metadata",
    normalize: s => s.replace(/refreshedBy="[^"]*"/g, "")
  },
  {
    reason: "refreshedDate is timestamp metadata",
    normalize: s => s.replace(/refreshedDate="[^"]*"/g, "")
  },
  {
    reason: "refreshOnLoad may be added",
    normalize: s => s.replace(/refreshOnLoad="1"/g, "")
  },
  {
    reason: "createdVersion is version metadata",
    normalize: s => s.replace(/createdVersion="\d+"/g, 'createdVersion="#"')
  },
  {
    reason: "refreshedVersion is version metadata",
    normalize: s => s.replace(/refreshedVersion="\d+"/g, 'refreshedVersion="#"')
  },
  {
    reason: "minRefreshableVersion is version metadata",
    normalize: s => s.replace(/minRefreshableVersion="\d+"/g, 'minRefreshableVersion="#"')
  },

  // ============================================
  // Shared strings references (index may change)
  // ============================================
  {
    reason: "Shared string index may be renumbered",
    normalize: s => {
      // All cells with t="s" reference sharedStrings by index
      // The index may be renumbered, but the actual string content is preserved
      return s.replace(/t="s"([^>]*)><v>\d+<\/v>/g, 't="s"$1><v>#</v>');
    }
  },

  // ============================================
  // workbookView self-closing vs non-self-closing
  // ============================================
  {
    reason: "workbookView may have different closing style",
    normalize: s => s.replace(/ \//g, "/")
  },

  // ============================================
  // Relationship file ordering
  // ============================================
  {
    reason: "Relationship element order may differ",
    normalize: s => {
      // Extract all Relationship elements and sort by Target
      const matches = s.match(/<Relationship[^>]+\/>/g) || [];
      if (matches.length === 0) {
        return s;
      }
      const sorted = [...matches].sort((a, b) => {
        const targetA = a.match(/Target="([^"]+)"/)?.[1] || "";
        const targetB = b.match(/Target="([^"]+)"/)?.[1] || "";
        return targetA.localeCompare(targetB);
      });
      let result = s;
      for (const m of matches) {
        result = result.replace(m, "{{REL}}");
      }
      for (const m of sorted) {
        result = result.replace("{{REL}}", m);
      }
      return result;
    }
  }
];

/**
 * Normalize XML attributes within a tag to be sorted alphabetically
 */
function normalizeAttributeOrder(xml: string): string {
  // Match opening tags with attributes
  return xml.replace(/<([a-zA-Z0-9:]+)(\s+[^>]+)(\/?>)/g, (match, tagName, attrs, closing) => {
    // Extract attributes
    const attrRegex = /([a-zA-Z0-9:]+)="([^"]*)"/g;
    const attributes: Array<{ name: string; value: string }> = [];
    let attrMatch;
    while ((attrMatch = attrRegex.exec(attrs)) !== null) {
      attributes.push({ name: attrMatch[1], value: attrMatch[2] });
    }
    // Sort attributes by name
    attributes.sort((a, b) => a.name.localeCompare(b.name));
    // Rebuild the tag
    const sortedAttrs = attributes.map(a => `${a.name}="${a.value}"`).join(" ");
    return `<${tagName} ${sortedAttrs}${closing}`;
  });
}

/**
 * Apply all atomic exclusions to normalize content for comparison
 */
function normalizeForComparison(content: string): string {
  let result = content;
  for (const exclusion of ATOMIC_EXCLUSIONS) {
    result = exclusion.normalize(result);
  }
  // Sort attributes within each tag to handle attribute reordering
  result = normalizeAttributeOrder(result);
  // Normalize whitespace between tags
  result = result.replace(/>\s+</g, "><");
  // Normalize multiple spaces to single space within tags
  result = result.replace(/\s+/g, " ");
  // Remove trailing space before > or />
  result = result.replace(/ >/g, ">");
  result = result.replace(/ \/>/g, "/>");
  // Trim
  result = result.trim();
  return result;
}

/**
 * Extract relationship targets (ignoring rId numbering)
 */
function extractRelationshipTargets(xml: string): Map<string, string> {
  const targets = new Map<string, string>();
  const regex = /Type="([^"]+)"[^>]*Target="([^"]+)"/g;
  let match;
  while ((match = regex.exec(xml)) !== null) {
    const type = match[1].split("/").pop() || match[1];
    const target = match[2];
    // Use type+target as key to handle multiple relationships of same type
    targets.set(`${type}:${target}`, `${type} -> ${target}`);
  }
  return targets;
}

/**
 * Compare two sets and return differences
 */
function compareMaps(
  inputMap: Map<string, string>,
  outputMap: Map<string, string>
): { missing: string[]; extra: string[] } {
  const missing: string[] = [];
  const extra: string[] = [];

  for (const [key, value] of inputMap) {
    if (!outputMap.has(key)) {
      missing.push(value);
    }
  }
  for (const [key, value] of outputMap) {
    if (!inputMap.has(key)) {
      extra.push(value);
    }
  }

  return { missing, extra };
}

describe("Workbook Round-trip Preservation", () => {
  let inputBuffer: Buffer;
  let outputBuffer: Buffer;
  let inputZip: Record<string, Uint8Array>;
  let outputZip: Record<string, Uint8Array>;

  beforeAll(async () => {
    // Read the sample file
    inputBuffer = fs.readFileSync(SAMPLE_FILE_PATH);

    // Load and save the workbook
    const workbook = Workbook.create();
    await Workbook.read(workbook, inputBuffer);
    outputBuffer = (await Workbook.toBuffer(workbook)) as Buffer;

    // OOXML conformance gate on every round-trip output.
    await expectValidXlsx(new Uint8Array(outputBuffer));

    // Parse both zip files
    inputZip = new ZipParser(inputBuffer).extractAllSync();
    outputZip = new ZipParser(outputBuffer).extractAllSync();
  });

  describe("File Structure", () => {
    it("should preserve all files from input (no missing files)", () => {
      const inputFiles = Object.keys(inputZip).sort();
      const outputFiles = new Set(Object.keys(outputZip));

      const missing = inputFiles.filter(f => !outputFiles.has(f));
      expect(missing, `Missing files: ${missing.join(", ")}`).toEqual([]);
    });

    it("should not add unexpected files", () => {
      const inputFiles = new Set(Object.keys(inputZip));
      const outputFiles = Object.keys(outputZip).sort();

      const extra = outputFiles.filter(f => !inputFiles.has(f));
      expect(extra, `Extra files: ${extra.join(", ")}`).toEqual([]);
    });
  });

  describe("Content Types ([Content_Types].xml)", () => {
    it("should preserve all content type mappings", () => {
      const inputContent = new TextDecoder().decode(inputZip["[Content_Types].xml"]);
      const outputContent = new TextDecoder().decode(outputZip["[Content_Types].xml"]);

      // Extract PartName -> ContentType mappings
      const extractTypes = (xml: string) => {
        const types = new Map<string, string>();
        const overrideRegex = /PartName="([^"]+)"[^>]*ContentType="([^"]+)"/g;
        let match;
        while ((match = overrideRegex.exec(xml)) !== null) {
          types.set(match[1], match[2]);
        }
        const defaultRegex = /Extension="([^"]+)"[^>]*ContentType="([^"]+)"/g;
        while ((match = defaultRegex.exec(xml)) !== null) {
          types.set(`*.${match[1]}`, match[2]);
        }
        return types;
      };

      const inputTypes = extractTypes(inputContent);
      const outputTypes = extractTypes(outputContent);

      const { missing, extra } = compareMaps(inputTypes, outputTypes);
      expect(missing, `Missing content types: ${missing.join(", ")}`).toEqual([]);
      expect(extra, `Extra content types: ${extra.join(", ")}`).toEqual([]);
    });
  });

  describe("Relationships Preservation", () => {
    const relsFiles = [
      "_rels/.rels",
      "xl/_rels/workbook.xml.rels",
      "xl/worksheets/_rels/sheet2.xml.rels",
      "xl/drawings/_rels/drawing1.xml.rels",
      "xl/pivotTables/_rels/pivotTable1.xml.rels",
      "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"
    ];

    for (const relsFile of relsFiles) {
      it(`should preserve ${relsFile} relationship targets`, () => {
        const inputContent = inputZip[relsFile];
        const outputContent = outputZip[relsFile];

        if (!inputContent) {
          // File doesn't exist in input, skip
          return;
        }

        expect(outputContent, `Missing file: ${relsFile}`).toBeDefined();

        const inputTargets = extractRelationshipTargets(new TextDecoder().decode(inputContent));
        const outputTargets = extractRelationshipTargets(new TextDecoder().decode(outputContent));

        const { missing, extra } = compareMaps(inputTargets, outputTargets);
        expect(missing, `Missing relationships in ${relsFile}: ${missing.join(", ")}`).toEqual([]);
        expect(extra, `Extra relationships in ${relsFile}: ${extra.join(", ")}`).toEqual([]);
      });
    }
  });

  describe("Chart Files (must be byte-identical)", () => {
    // Style, colors and theme are passed through without modification
    const passthroughFiles = [
      "xl/charts/colors1.xml",
      "xl/charts/style1.xml",
      "xl/theme/theme1.xml"
    ];

    for (const chartFile of passthroughFiles) {
      it(`should preserve ${chartFile} exactly`, () => {
        const inputContent = inputZip[chartFile];
        const outputContent = outputZip[chartFile];

        if (!inputContent) {
          return;
        }

        expect(outputContent, `Missing file: ${chartFile}`).toBeDefined();

        const inputStr = new TextDecoder().decode(inputContent);
        const outputStr = new TextDecoder().decode(outputContent);

        expect(outputStr, `${chartFile} should be byte-identical`).toBe(inputStr);
      });
    }
  });

  describe("Chart XML (regenerated, newline-normalized check)", () => {
    it("should preserve xl/charts/chart1.xml after newline normalization", () => {
      const inputContent = inputZip["xl/charts/chart1.xml"];
      const outputContent = outputZip["xl/charts/chart1.xml"];

      if (!inputContent) {
        return;
      }

      expect(outputContent, "Missing file: xl/charts/chart1.xml").toBeDefined();

      // Normalize line endings for comparison (writer uses LF, source may use CRLF)
      const inputStr = new TextDecoder().decode(inputContent).replace(/\r\n/g, "\n");
      const outputStr = new TextDecoder().decode(outputContent).replace(/\r\n/g, "\n");

      expect(outputStr, "xl/charts/chart1.xml should match after newline normalization").toBe(
        inputStr
      );
    });
  });

  describe("Drawing Files (regenerated, structural check)", () => {
    it("should preserve xl/drawings/drawing1.xml structure", () => {
      const inputContent = inputZip["xl/drawings/drawing1.xml"];
      const outputContent = outputZip["xl/drawings/drawing1.xml"];

      if (!inputContent) {
        return;
      }

      expect(outputContent, "Missing file: xl/drawings/drawing1.xml").toBeDefined();

      // Normalize line endings for comparison (writer uses LF, source may use CRLF)
      const inputStr = new TextDecoder().decode(inputContent).replace(/\r\n/g, "\n");
      const outputStr = new TextDecoder().decode(outputContent).replace(/\r\n/g, "\n");

      expect(outputStr, "xl/drawings/drawing1.xml should match after newline normalization").toBe(
        inputStr
      );
    });
  });

  describe("Pivot Table Structure", () => {
    it("should preserve pivotTable1.xml critical elements", () => {
      const inputContent = new TextDecoder().decode(inputZip["xl/pivotTables/pivotTable1.xml"]);
      const outputContent = new TextDecoder().decode(outputZip["xl/pivotTables/pivotTable1.xml"]);

      // Critical: name and cacheId
      expect(outputContent).toContain('name="PivotTable1"');
      expect(outputContent).toContain('cacheId="3"');

      // Critical: location reference
      expect(outputContent).toContain('ref="A1:B6"');

      // Critical: field counts must match
      const inputPivotFieldsCount = inputContent.match(/<pivotFields count="(\d+)"/)?.[1];
      const outputPivotFieldsCount = outputContent.match(/<pivotFields count="(\d+)"/)?.[1];
      expect(outputPivotFieldsCount).toBe(inputPivotFieldsCount);

      // Critical: rowFields must match
      const inputRowFields = inputContent.match(/<rowFields count="(\d+)"/)?.[1];
      const outputRowFields = outputContent.match(/<rowFields count="(\d+)"/)?.[1];
      expect(outputRowFields).toBe(inputRowFields);

      // Critical: dataFields must match
      const inputDataFields = inputContent.match(/<dataFields count="(\d+)"/)?.[1];
      const outputDataFields = outputContent.match(/<dataFields count="(\d+)"/)?.[1];
      expect(outputDataFields).toBe(inputDataFields);
      expect(outputContent).toContain('name="Sum of Revenue"');

      const inputHasColFields = inputContent.includes("<colFields");
      const outputHasColFields = outputContent.includes("<colFields");
      expect(outputHasColFields, "colFields presence must match input").toBe(inputHasColFields);

      // Critical: chartFormats for pivot chart
      const inputChartFormats = inputContent.match(/<chartFormats count="(\d+)"/)?.[1];
      const outputChartFormats = outputContent.match(/<chartFormats count="(\d+)"/)?.[1];
      expect(outputChartFormats).toBe(inputChartFormats);
    });

    it("should preserve pivotCacheDefinition1.xml data source", () => {
      const inputContent = new TextDecoder().decode(
        inputZip["xl/pivotCache/pivotCacheDefinition1.xml"]
      );
      const outputContent = new TextDecoder().decode(
        outputZip["xl/pivotCache/pivotCacheDefinition1.xml"]
      );

      // Critical: data source reference
      expect(outputContent).toContain('ref="A1:F51"');
      expect(outputContent).toContain('sheet="Data"');

      // Critical: cache fields count
      const inputCacheFields = inputContent.match(/<cacheFields count="(\d+)"/)?.[1];
      const outputCacheFields = outputContent.match(/<cacheFields count="(\d+)"/)?.[1];
      expect(outputCacheFields).toBe(inputCacheFields);

      // Critical: record count
      const inputRecordCount = inputContent.match(/recordCount="(\d+)"/)?.[1];
      const outputRecordCount = outputContent.match(/recordCount="(\d+)"/)?.[1];
      expect(outputRecordCount).toBe(inputRecordCount);

      // Critical: shared items for Region field
      expect(outputContent).toContain('v="Northeast"');
      expect(outputContent).toContain('v="Midwest"');
      expect(outputContent).toContain('v="South"');
      expect(outputContent).toContain('v="West"');
    });

    it("should preserve pivotCacheRecords1.xml record data", () => {
      const inputContent = new TextDecoder().decode(
        inputZip["xl/pivotCache/pivotCacheRecords1.xml"]
      );
      const outputContent = new TextDecoder().decode(
        outputZip["xl/pivotCache/pivotCacheRecords1.xml"]
      );

      // Critical: record count
      const inputCount = inputContent.match(/count="(\d+)"/)?.[1];
      const outputCount = outputContent.match(/count="(\d+)"/)?.[1];
      expect(outputCount).toBe(inputCount);

      // Critical: sample data values must be present
      expect(outputContent).toContain('v="C001"');
      expect(outputContent).toContain('v="Customer 1"');
    });
  });

  describe("Workbook Structure", () => {
    it("should preserve sheet names and order", () => {
      const inputContent = new TextDecoder().decode(inputZip["xl/workbook.xml"]);
      const outputContent = new TextDecoder().decode(outputZip["xl/workbook.xml"]);

      // Extract sheet names in order
      const extractSheets = (xml: string) => {
        const sheets: string[] = [];
        const regex = /name="([^"]+)"/g;
        let match;
        // Only match within <sheets> section
        const sheetsSection = xml.match(/<sheets>([\s\S]*?)<\/sheets>/)?.[1] || "";
        while ((match = regex.exec(sheetsSection)) !== null) {
          sheets.push(match[1]);
        }
        return sheets;
      };

      const inputSheets = extractSheets(inputContent);
      const outputSheets = extractSheets(outputContent);

      expect(outputSheets).toEqual(inputSheets);
    });

    it("should preserve pivotCaches", () => {
      const outputContent = new TextDecoder().decode(outputZip["xl/workbook.xml"]);

      expect(outputContent).toContain("<pivotCaches>");
      expect(outputContent).toContain('cacheId="3"');
    });
  });

  describe("Worksheet Data", () => {
    it("should preserve sheet1.xml (Data) row count", () => {
      const inputContent = new TextDecoder().decode(inputZip["xl/worksheets/sheet1.xml"]);
      const outputContent = new TextDecoder().decode(outputZip["xl/worksheets/sheet1.xml"]);

      // Count <row> elements
      const inputRows = (inputContent.match(/<row /g) || []).length;
      const outputRows = (outputContent.match(/<row /g) || []).length;

      expect(outputRows).toBe(inputRows);
    });

    it("should preserve sheet1.xml numeric cell values", () => {
      const inputContent = new TextDecoder().decode(inputZip["xl/worksheets/sheet1.xml"]);
      const outputContent = new TextDecoder().decode(outputZip["xl/worksheets/sheet1.xml"]);

      // Extract only numeric values (not shared string references)
      // Numeric cells don't have t="s" attribute
      const extractNumericValues = (xml: string) => {
        const values: string[] = [];
        // Match cells without t="s" that have values
        const regex = /<c r="[^"]+"(?![^>]*t="s")[^>]*><v>([^<]+)<\/v><\/c>/g;
        let match;
        while ((match = regex.exec(xml)) !== null) {
          values.push(match[1]);
        }
        return values;
      };

      const inputValues = extractNumericValues(inputContent);
      const outputValues = extractNumericValues(outputContent);

      expect(outputValues).toEqual(inputValues);
    });

    it("should preserve sheet2.xml drawing reference", () => {
      const outputContent = new TextDecoder().decode(outputZip["xl/worksheets/sheet2.xml"]);

      // Must have drawing reference (rId may differ)
      expect(outputContent).toMatch(/<drawing r:id="rId\d+"\/>/);
    });

    it("should preserve sheet2.xml pivot table cell structure", () => {
      const inputContent = new TextDecoder().decode(inputZip["xl/worksheets/sheet2.xml"]);
      const outputContent = new TextDecoder().decode(outputZip["xl/worksheets/sheet2.xml"]);

      // Count cells in each row
      const countCells = (xml: string) => {
        const rows = xml.match(/<row[^>]*>[\s\S]*?<\/row>/g) || [];
        return rows.map(row => {
          const cells = row.match(/<c /g) || [];
          return cells.length;
        });
      };

      const inputCellCounts = countCells(inputContent);
      const outputCellCounts = countCells(outputContent);

      expect(outputCellCounts).toEqual(inputCellCounts);
    });
  });

  describe("Normalized Content Comparison", () => {
    // Files that should match after applying atomic exclusions
    const criticalFiles = [
      "xl/workbook.xml",
      "xl/worksheets/sheet1.xml",
      "xl/worksheets/sheet2.xml",
      "xl/pivotTables/pivotTable1.xml",
      "xl/pivotCache/pivotCacheDefinition1.xml",
      "xl/pivotCache/pivotCacheRecords1.xml"
    ];

    for (const file of criticalFiles) {
      it(`should match ${file} after applying atomic exclusions`, () => {
        const inputContent = inputZip[file];
        const outputContent = outputZip[file];

        if (!inputContent) {
          return;
        }
        expect(outputContent, `Missing file: ${file}`).toBeDefined();

        const inputNormalized = normalizeForComparison(new TextDecoder().decode(inputContent));
        const outputNormalized = normalizeForComparison(new TextDecoder().decode(outputContent));

        // If they don't match, show the first difference
        if (inputNormalized !== outputNormalized) {
          // Find first difference position
          let diffPos = 0;
          while (diffPos < inputNormalized.length && diffPos < outputNormalized.length) {
            if (inputNormalized[diffPos] !== outputNormalized[diffPos]) {
              break;
            }
            diffPos++;
          }

          const contextStart = Math.max(0, diffPos - 50);
          const contextEnd = Math.min(
            Math.max(inputNormalized.length, outputNormalized.length),
            diffPos + 100
          );

          const inputContext = inputNormalized.substring(contextStart, contextEnd);
          const outputContext = outputNormalized.substring(contextStart, contextEnd);

          expect.fail(
            `${file} differs at position ${diffPos}:\n` +
              `Input:  ...${inputContext}...\n` +
              `Output: ...${outputContext}...`
          );
        }
      });
    }
  });

  describe("Style Preservation", () => {
    it("should preserve pivotButton attribute in cellXfs", () => {
      const outputContent = new TextDecoder().decode(outputZip["xl/styles.xml"]);

      // pivotButton="1" must be preserved for pivot table cells
      expect(outputContent).toContain('pivotButton="1"');
    });

    it("should preserve applyNumberFormat attribute in cellXfs", () => {
      const outputContent = new TextDecoder().decode(outputZip["xl/styles.xml"]);

      // applyNumberFormat="1" must be preserved
      expect(outputContent).toContain('applyNumberFormat="1"');
    });

    it("should preserve font name (Aptos Narrow)", () => {
      const inputContent = new TextDecoder().decode(inputZip["xl/styles.xml"]);
      const outputContent = new TextDecoder().decode(outputZip["xl/styles.xml"]);

      // Extract font name from input
      const inputFontName = inputContent.match(/<name val="([^"]+)"/)?.[1];
      expect(inputFontName).toBe("Aptos Narrow");

      // Output must have the same font name
      expect(outputContent).toContain('val="Aptos Narrow"');
    });

    it("should preserve font count and structure", () => {
      const inputContent = new TextDecoder().decode(inputZip["xl/styles.xml"]);
      const outputContent = new TextDecoder().decode(outputZip["xl/styles.xml"]);

      // Extract font count
      const inputFontCount = inputContent.match(/<fonts count="(\d+)"/)?.[1];
      const outputFontCount = outputContent.match(/<fonts count="(\d+)"/)?.[1];

      expect(outputFontCount).toBe(inputFontCount);
    });
  });

  describe("docProps (metadata files - allowed to differ)", () => {
    // These files contain application metadata and are expected to change
    it("should have docProps/core.xml", () => {
      expect(outputZip["docProps/core.xml"]).toBeDefined();
    });

    it("should have docProps/app.xml", () => {
      expect(outputZip["docProps/app.xml"]).toBeDefined();
    });
  });
});
