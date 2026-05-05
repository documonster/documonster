/**
 * Synthetic xlsx package builder for validator unit tests.
 *
 * Each checker is tested with **hand-crafted minimal packages** rather
 * than by running the full serialiser — this keeps tests deterministic
 * and makes it trivial to reproduce specific structural violations.
 *
 * The helpers below build just enough OOXML scaffolding to exercise
 * one check at a time; callers override specific parts via the options
 * object.
 */

import { ZipArchive } from "@archive";

const TEXT = new TextEncoder();

export interface PackageParts {
  /** Override / add entries. Key is zip path, value is file content. */
  [path: string]: string | Uint8Array;
}

/**
 * Build an xlsx zip from a parts map. Caller is responsible for
 * providing each part's contents; this helper just packages them.
 */
export function buildPackage(parts: PackageParts): Uint8Array {
  const zip = new ZipArchive({
    level: 0,
    timestamps: "dos",
    modTime: new Date(1980, 0, 1, 0, 0, 0)
  });
  for (const [path, data] of Object.entries(parts)) {
    zip.add(path, typeof data === "string" ? TEXT.encode(data) : data);
  }
  return zip.bytesSync();
}

// -----------------------------------------------------------------------------
// Minimal valid package
// -----------------------------------------------------------------------------

/**
 * A package layout that is valid OPC and passes every default check.
 * Use `{ ...baseParts(), "xl/workbook.xml": customXml }` to mutate a
 * single part for targeted tests.
 */
export function baseParts(): PackageParts {
  return {
    "[Content_Types].xml": baseContentTypes(),
    "_rels/.rels": baseRootRels(),
    "xl/_rels/workbook.xml.rels": baseWorkbookRels(),
    "xl/workbook.xml": baseWorkbook(),
    "xl/styles.xml": baseStyles(),
    "xl/sharedStrings.xml": baseSharedStrings(),
    "xl/theme/theme1.xml": baseTheme(),
    "xl/worksheets/sheet1.xml": baseSheet()
  };
}

export function baseContentTypes(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`;
}

export function baseRootRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
}

export function baseWorkbookRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`;
}

export function baseWorkbook(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;
}

export function baseStyles(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
}

export function baseSharedStrings(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>`;
}

export function baseTheme(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
  <a:themeElements/>
</a:theme>`;
}

export function baseSheet(body?: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>${body ?? ""}</sheetData>
</worksheet>`;
}

/**
 * Replace the content-types Override list with the given overrides.
 * Easier than hand-editing the XML in tests that vary content types.
 */
export function contentTypesWith(
  overrides: Array<{ partName: string; contentType: string }>
): string {
  const lines = overrides
    .map(o => `  <Override PartName="${o.partName}" ContentType="${o.contentType}"/>`)
    .join("\n");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
${lines}
</Types>`;
}

/**
 * Wrap a body in a `<Relationships>` root. Each entry becomes a
 * `<Relationship>` element with the standard 4 attributes.
 */
export function relsWith(
  entries: Array<Partial<{ id: string; type: string; target: string; targetMode: string }>>
): string {
  const lines = entries
    .map(e => {
      const attrs: string[] = [];
      if (e.id !== undefined) {
        attrs.push(`Id="${e.id}"`);
      }
      if (e.type !== undefined) {
        attrs.push(`Type="${e.type}"`);
      }
      if (e.target !== undefined) {
        attrs.push(`Target="${e.target}"`);
      }
      if (e.targetMode !== undefined) {
        attrs.push(`TargetMode="${e.targetMode}"`);
      }
      return `  <Relationship ${attrs.join(" ")}/>`;
    })
    .join("\n");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${lines}
</Relationships>`;
}
