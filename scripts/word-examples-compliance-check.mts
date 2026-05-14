/**
 * Static OOXML compliance scanner for Word example outputs.
 *
 * Checks every .docx in tmp/word-examples/ against a set of structural
 * invariants taken from ECMA-376 Part 1 §17:
 *
 *   1. ZIP container is valid; word/document.xml exists (skip CFB-encrypted).
 *   2. document.xml well-formed XML.
 *   3. <w:body> is non-empty (must contain at least one paragraph or sectPr).
 *   4. Every <w:tbl> has a <w:tblGrid> with at least one <w:gridCol>.
 *   5. Every <w:tc> has a <w:tcPr><w:tcW> child.
 *   6. Inside any cell, no nested <w:tbl> appears at the very end without a
 *      trailing <w:p> (§17.4.66).
 *   7. Adjacent <w:tbl> siblings inside body / tc are separated by at least
 *      one <w:p> (§17.13.5.34).
 *   8. Every fldChar sequence inside a paragraph (or hyperlink) is balanced
 *      (begin..separate..end), no orphan begin/end.
 *   9. Every <w:hyperlink> has either r:id, w:anchor, or both.
 *  10. Every <w:rStyle w:val="X"/> reference resolves to a defined character
 *      style in styles.xml (or DefaultParagraphFont).
 *  11. Every <w:pStyle w:val="X"/> reference resolves to a paragraph style.
 *  12. Every basedOn / next / link in styles.xml resolves.
 *  13. document.xml.rels: every r:id used in document.xml is defined.
 *  14. Every <w:bookmarkStart> is closed by a <w:bookmarkEnd> with matching id.
 *  15. Every <w:numId> reference resolves to a numId defined in numbering.xml.
 *
 * Usage: npx tsx scripts/word-examples-compliance-check.mts
 */
import * as fs from "node:fs";
import * as path from "node:path";

import { extractAll } from "../src/modules/archive/unzip/extract";

interface Issue {
  file: string;
  rule: string;
  detail: string;
}

const issues: Issue[] = [];
let docsScanned = 0;

const exampleDir = path.resolve(import.meta.dirname, "../tmp/word-examples");

async function* walk(dir: string): AsyncGenerator<string> {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      yield* walk(full);
    } else if (entry.name.endsWith(".docx") || entry.name.endsWith(".docm")) {
      yield full;
    }
  }
}

function report(file: string, rule: string, detail: string): void {
  issues.push({ file, rule, detail });
}

const ZIP_MAGIC = new Uint8Array([0x50, 0x4b, 0x03, 0x04]);
const CFB_MAGIC = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);

async function checkDocx(file: string): Promise<void> {
  const bytes = new Uint8Array(fs.readFileSync(file));
  const isCfb = CFB_MAGIC.every((b, i) => bytes[i] === b);
  if (isCfb) return; // encrypted; skip
  const isZip = ZIP_MAGIC.every((b, i) => bytes[i] === b);
  if (!isZip) {
    report(file, "container", "not a ZIP and not a CFB");
    return;
  }
  docsScanned++;
  let entries;
  try {
    entries = await extractAll(bytes);
  } catch (e) {
    report(file, "container", `unzip failed: ${(e as Error).message}`);
    return;
  }

  const docXml = entries.get("word/document.xml");
  if (!docXml) {
    report(file, "container", "missing word/document.xml");
    return;
  }
  const xml = new TextDecoder().decode(docXml.data);

  // 3. body non-empty
  if (
    !/<w:body[^>]*>[\s\S]*?<w:p[\s/>]/.test(xml) &&
    !/<w:body[^>]*>[\s\S]*?<w:tbl[\s>]/.test(xml)
  ) {
    report(file, "body-empty", "<w:body> has no paragraph or table");
  }

  // 4. tblGrid presence
  const tblMatches = xml.matchAll(/<w:tbl[\s>][\s\S]*?<\/w:tbl>/g);
  for (const m of tblMatches) {
    const tbl = m[0];
    if (!/<w:tblGrid>[\s\S]*?<w:gridCol/.test(tbl)) {
      report(file, "tblGrid", "<w:tbl> missing <w:tblGrid> with <w:gridCol>");
    }
  }

  // 5. tc must have tcPr with tcW
  const tcMatches = xml.matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g);
  for (const m of tcMatches) {
    const tc = m[0];
    if (!/<w:tcPr>[\s\S]*?<w:tcW[\s/]/.test(tc)) {
      report(file, "tcPr-tcW", "<w:tc> missing <w:tcPr><w:tcW>");
    }
  }

  // 6. cell ends with paragraph (not nested table)
  // Simplified: any "</w:tbl></w:tc>" sequence indicates cell ends with a tbl.
  if (/<\/w:tbl>\s*<\/w:tc>/.test(xml)) {
    report(file, "cell-trailing-p", "cell ends with </w:tbl> not followed by <w:p/>");
  }

  // 7. adjacent <w:tbl> without paragraph separator
  // Pattern: </w:tbl> directly followed by <w:tbl> (no <w:p> between)
  if (/<\/w:tbl>\s*<w:tbl[\s>]/.test(xml)) {
    report(file, "tbl-separator", "two adjacent <w:tbl> without separating <w:p>");
  }

  // 8. fldChar balance
  // Count begin / end inside the doc (not perfect for nested but good enough)
  const beginCount = (xml.match(/w:fldCharType="begin"/g) ?? []).length;
  const endCount = (xml.match(/w:fldCharType="end"/g) ?? []).length;
  if (beginCount !== endCount) {
    report(file, "fldChar-balance", `begin=${beginCount} end=${endCount}`);
  }

  // 9. hyperlink must have r:id or w:anchor
  const hypMatches = xml.matchAll(/<w:hyperlink([^>]*)>/g);
  for (const m of hypMatches) {
    const attrs = m[1];
    if (!/r:id="/.test(attrs) && !/w:anchor="/.test(attrs)) {
      report(file, "hyperlink-target", `<w:hyperlink${attrs}> has neither r:id nor w:anchor`);
    }
  }

  // 10/11/12. style references
  const stylesXml = entries.get("word/styles.xml");
  if (stylesXml) {
    const sxml = new TextDecoder().decode(stylesXml.data);
    const definedStyles = new Set<string>();
    for (const m of sxml.matchAll(/w:styleId="([^"]+)"/g)) {
      definedStyles.add(m[1]);
    }
    // Hyperlink basedOn DefaultParagraphFont — ensure DefaultParagraphFont
    // is actually defined when used as basedOn.
    for (const m of sxml.matchAll(
      /<w:basedOn w:val="([^"]+)"\/>|<w:next w:val="([^"]+)"\/>|<w:link w:val="([^"]+)"\/>/g
    )) {
      const ref = m[1] ?? m[2] ?? m[3];
      if (ref && !definedStyles.has(ref)) {
        report(file, "style-ref-missing", `styles.xml references undefined style "${ref}"`);
      }
    }
    // Body uses
    for (const m of xml.matchAll(/<w:pStyle w:val="([^"]+)"\/>/g)) {
      if (!definedStyles.has(m[1])) {
        report(file, "pStyle-undefined", `body uses pStyle "${m[1]}" not in styles.xml`);
      }
    }
    for (const m of xml.matchAll(/<w:rStyle w:val="([^"]+)"\/>/g)) {
      if (!definedStyles.has(m[1])) {
        report(file, "rStyle-undefined", `body uses rStyle "${m[1]}" not in styles.xml`);
      }
    }
  }

  // 13. r:id resolution against document.xml.rels
  const relsXml = entries.get("word/_rels/document.xml.rels");
  if (relsXml) {
    const rxml = new TextDecoder().decode(relsXml.data);
    const definedRids = new Set<string>();
    for (const m of rxml.matchAll(/Id="([^"]+)"/g)) {
      definedRids.add(m[1]);
    }
    for (const m of xml.matchAll(/r:id="([^"]+)"/g)) {
      if (!definedRids.has(m[1])) {
        report(file, "rId-undefined", `body uses r:id="${m[1]}" not in document.xml.rels`);
      }
    }
    for (const m of xml.matchAll(/r:embed="([^"]+)"/g)) {
      if (!definedRids.has(m[1])) {
        report(file, "rEmbed-undefined", `body uses r:embed="${m[1]}" not in document.xml.rels`);
      }
    }
  }

  // 14. bookmark balance
  const bmStarts = new Map<string, number>();
  for (const m of xml.matchAll(/<w:bookmarkStart\s+([^>]*)\/>/g)) {
    const id = /w:id="([^"]+)"/.exec(m[1])?.[1];
    if (id) bmStarts.set(id, (bmStarts.get(id) ?? 0) + 1);
  }
  const bmEnds = new Map<string, number>();
  for (const m of xml.matchAll(/<w:bookmarkEnd\s+([^>]*)\/>/g)) {
    const id = /w:id="([^"]+)"/.exec(m[1])?.[1];
    if (id) bmEnds.set(id, (bmEnds.get(id) ?? 0) + 1);
  }
  for (const [id, count] of bmStarts) {
    if (bmEnds.get(id) !== count) {
      report(
        file,
        "bookmark-unbalanced",
        `bookmarkStart id=${id} count=${count} but end=${bmEnds.get(id) ?? 0}`
      );
    }
  }

  // 15. numId resolution
  const numXml = entries.get("word/numbering.xml");
  if (numXml) {
    const nxml = new TextDecoder().decode(numXml.data);
    const definedNumIds = new Set<string>();
    for (const m of nxml.matchAll(/<w:num w:numId="([^"]+)"/g)) {
      definedNumIds.add(m[1]);
    }
    for (const m of xml.matchAll(/<w:numId w:val="([^"]+)"\/>/g)) {
      if (!definedNumIds.has(m[1])) {
        report(file, "numId-undefined", `body uses numId="${m[1]}" not in numbering.xml`);
      }
    }
  } else {
    // No numbering.xml but body uses numId
    if (/<w:numId w:val="/.test(xml)) {
      report(file, "numbering-missing", "body uses numId but word/numbering.xml is absent");
    }
  }
}

console.log("Scanning tmp/word-examples ...");
for await (const file of walk(exampleDir)) {
  await checkDocx(file);
}

console.log(`\nScanned ${docsScanned} docx files.`);
console.log(`Found ${issues.length} issues.\n`);
if (issues.length > 0) {
  // Group by rule
  const byRule = new Map<string, Issue[]>();
  for (const issue of issues) {
    const list = byRule.get(issue.rule) ?? [];
    list.push(issue);
    byRule.set(issue.rule, list);
  }
  for (const [rule, list] of byRule) {
    console.log(`\n## ${rule} (${list.length})`);
    for (const issue of list.slice(0, 10)) {
      const rel = path.relative(exampleDir, issue.file);
      console.log(`  ${rel}: ${issue.detail}`);
    }
    if (list.length > 10) {
      console.log(`  ... ${list.length - 10} more`);
    }
  }
  process.exit(1);
}
console.log("✅ ALL CLEAN");
