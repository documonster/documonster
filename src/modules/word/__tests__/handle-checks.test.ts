/**
 * Every public function that takes a handle — a document, a glossary, a style
 * map — checks it at the boundary with one helper (`checkHandle`) and refuses
 * anything else with `DocxArgumentError` naming the parameter, before any
 * other argument. Before, most failed inside with an untyped `TypeError`
 * ("Cannot read properties of undefined (reading 'some')").
 *
 * The list is every namespace function whose first parameter is a handle, as
 * `tmp/api-review/r2/word/scan-handles.py` finds them.
 */

import { describe, it, expect } from "vitest";

import { WordCrypto } from "../crypto";
import { DocxArgumentError } from "../errors";
import { WordExcel } from "../excel";
import { WordHtml } from "../html";
import * as Word from "../index";
import { WordMarkdown } from "../markdown";

const NAMESPACES: Record<string, Record<string, unknown>> = {
  ...(Word as unknown as Record<string, Record<string, unknown>>),
  WordHtml,
  WordMarkdown,
  WordExcel,
  WordCrypto
};

const HANDLE_FUNCTIONS: readonly (readonly [name: string, param: string])[] = [
  ["Document.discardUnrecognizedValues", "doc"],
  ["Document.clone", "doc"],
  ["Document.addContent", "doc"],
  ["Document.insertContent", "doc"],
  ["Document.removeContent", "doc"],
  ["Document.getContentCount", "doc"],
  ["Document.getContent", "doc"],
  ["Document.addParagraph", "doc"],
  ["Document.addParagraphElement", "doc"],
  ["Document.addHeading", "doc"],
  ["Document.addPageBreak", "doc"],
  ["Document.addTable", "doc"],
  ["Document.addTableElement", "doc"],
  ["Document.addImage", "doc"],
  ["Document.addFloatingImage", "doc"],
  ["Document.addFont", "doc"],
  ["Document.setWatermark", "doc"],
  ["Document.addFootnote", "doc"],
  ["Document.addEndnote", "doc"],
  ["Document.addComment", "doc"],
  ["Document.addTableOfContents", "doc"],
  ["Document.addMath", "doc"],
  ["Document.addTextBox", "doc"],
  ["Document.addListDefinition", "doc"],
  ["Document.addBulletList", "doc"],
  ["Document.addNumberedList", "doc"],
  ["Document.setSectionProperties", "doc"],
  ["Document.setDocDefaults", "doc"],
  ["Document.addStyle", "doc"],
  ["Document.useDefaultStyles", "doc"],
  ["Document.setHeader", "doc"],
  ["Document.setFooter", "doc"],
  ["Document.setSettings", "doc"],
  ["Document.setCoreProperties", "doc"],
  ["Document.setAppProperties", "doc"],
  ["Document.setBackground", "doc"],
  ["Document.addCustomProperty", "doc"],
  ["Document.addSectionBreak", "doc"],
  ["Document.allocateBookmarkId", "doc"],
  ["Document.build", "doc"],
  ["Document.replaceText", "doc"],
  ["Document.mailMerge", "doc"],
  ["Document.acceptAllRevisions", "doc"],
  ["Document.rejectAllRevisions", "doc"],
  ["Document.acceptRevision", "doc"],
  ["Document.rejectRevision", "doc"],
  ["Document.setCompatibilityMode", "doc"],
  ["Query.each", "doc"],
  ["Query.getParagraphStyle", "doc"],
  ["Query.getRunStyle", "doc"],
  ["Query.getTableStyle", "doc"],
  ["Query.getParagraphCount", "doc"],
  ["Query.getTableCount", "doc"],
  ["Query.getWordCount", "doc"],
  ["Query.getText", "doc"],
  ["Query.getBookmark", "doc"],
  ["Query.getComment", "doc"],
  ["Query.listHeadings", "doc"],
  ["Query.listImages", "doc"],
  ["Query.listTables", "doc"],
  ["Query.listHyperlinks", "doc"],
  ["Query.listSections", "doc"],
  ["Query.searchText", "doc"],
  ["Query.listFields", "doc"],
  ["Query.listRevisions", "doc"],
  ["Query.createStyleIndex", "doc"],
  ["Query.getNumberingLevel", "doc"],
  ["Query.getCompatibilityMode", "doc"],
  ["Query.listFormFields", "doc"],
  ["Query.searchByFormat", "doc"],
  ["Query.getCountByFormat", "doc"],
  ["Query.listUsedFormats", "doc"],
  ["Query.listParagraphs", "doc"],
  ["Query.listRuns", "doc"],
  ["Io.toBuffer", "doc"],
  ["Io.toBase64", "doc"],
  ["Io.toFlatOpc", "doc"],
  // `Io.writeFile` is Node-only and is checked in `io-node-surface.node.test.ts`.
  ["Transform.split", "doc"],
  ["Transform.updateFields", "doc"],
  ["Transform.updateTableOfContents", "doc"],
  ["Transform.map", "doc"],
  ["Transform.fillFormFields", "doc"],
  ["Transform.resolveDataBindings", "doc"],
  ["Template.bindChartData", "doc"],
  ["Template.listTags", "doc"],
  ["Template.fillFromSource", "doc"],
  ["Convert.toOdt", "doc"],
  ["Convert.toSemantic", "doc"],
  ["Font.addEmbedded", "doc"],
  ["Layout.getPageMap", "doc"],
  ["Layout.compute", "doc"],
  ["Layout.toSvg", "doc"],
  ["Layout.pageToSvg", "doc"],
  ["Security.protect", "doc"],
  ["Security.unprotect", "doc"],
  ["Security.isProtected", "doc"],
  ["Security.getState", "doc"],
  ["Security.verifyPassword", "doc"],
  ["Ole.extract", "doc"],
  ["Ole.has", "doc"],
  ["Ole.getData", "doc"],
  ["Ole.add", "doc"],
  ["Vba.has", "doc"],
  ["Vba.getInfo", "doc"],
  ["Vba.getData", "doc"],
  ["Vba.add", "doc"],
  ["Vba.remove", "doc"],
  ["Vba.listParts", "doc"],
  ["Glossary.findBlock", "glossary"],
  ["Glossary.listBlocks", "glossary"],
  ["Glossary.listAutoTextEntries", "glossary"],
  ["Glossary.listQuickParts", "glossary"],
  ["Styles.match", "map"],
  ["Diff.compare", "oldDoc"],
  ["Validation.validate", "doc"],
  ["WordHtml.fromDocx", "doc"],
  ["WordMarkdown.fromDocx", "doc"],
  ["WordExcel.extractTables", "doc"],
  // Takes a document or a list of part names; anything else is `source`.
  ["WordCrypto.hasSignatures", "source"],
  ["WordCrypto.listSignatures", "doc"]
];

function lookup(name: string): (...args: unknown[]) => unknown {
  const [ns, member] = name.split(".") as [string, string];
  const fn = NAMESPACES[ns]?.[member];
  expect(typeof fn, name).toBe("function");
  return fn as (...args: unknown[]) => unknown;
}

async function refusal(run: () => unknown): Promise<unknown> {
  try {
    await run();
  } catch (error) {
    return error;
  }
  return undefined;
}

describe("handles are checked at the boundary", () => {
  it.each(HANDLE_FUNCTIONS)("%s refuses a value that is not its handle", async (name, param) => {
    const fn = lookup(name);
    for (const bad of [42, null, {}, "doc"]) {
      const error = await refusal(() => fn(bad));
      expect(error, `${name}(${JSON.stringify(bad)})`).toBeInstanceOf(DocxArgumentError);
      expect((error as DocxArgumentError).field, name).toBe(param);
      expect((error as Error).message, name).toMatch(new RegExp(`^${name.replace(".", "\\.")}: `));
    }
  });
});
