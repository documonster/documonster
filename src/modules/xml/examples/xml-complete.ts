/**
 * Example: XML Module — Complete Guide
 *
 * Covers:
 * - Xml.Writer: building XML documents in memory
 * - Xml.StreamWriter: streaming XML to a writable target
 * - XML declaration, namespaces, attributes
 * - Text content, CDATA sections, comments
 * - Save/commit/rollback (transactional writes)
 * - DOM parsing with Xml.parse
 * - DOM navigation: Xml.findChild, Xml.findChildren, Xml.attr, Xml.textContent, Xml.walk
 * - Query engine: Xml.query, Xml.queryAll (XPath-like syntax)
 * - SAX streaming parser: Xml.SaxParser, Xml.parseSax
 * - Xml.parseToObject: convert XML to plain JS objects
 * - Encoding utilities: Xml.encode, Xml.decode, Xml.encodeAttr
 * - Error handling: XmlError, XmlParseError, XmlWriteError
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Xml, isXmlError, isXmlParseError } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/xml-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// =============================================================================
// 1. Xml.Writer — basic document
// =============================================================================

const w1 = new Xml.Writer();
w1.openXml(Xml.StdDocAttributes);
w1.openNode("catalog");
w1.leafNode("book", { id: "1", lang: "en" }, "The Great Gatsby");
w1.leafNode("book", { id: "2", lang: "fr" }, "Le Petit Prince");
w1.closeNode();

const xml1 = w1.xml;
fs.writeFileSync(path.join(outDir, "basic.xml"), xml1);
console.log("=== 1. Basic Xml.Writer ===");
console.log(xml1);

// =============================================================================
// 2. Nested elements with attributes
// =============================================================================

const w2 = new Xml.Writer();
w2.openXml(Xml.StdDocAttributes);
w2.openNode("employees");

w2.openNode("employee", { id: "101", department: "Engineering" });
w2.leafNode("name", undefined, "Alice Johnson");
w2.leafNode("title", undefined, "Senior Developer");
w2.leafNode("salary", { currency: "USD" }, "120000");
w2.closeNode(); // employee

w2.openNode("employee", { id: "102", department: "Design" });
w2.leafNode("name", undefined, "Bob Smith");
w2.leafNode("title", undefined, "UI Designer");
w2.leafNode("salary", { currency: "EUR" }, "95000");
w2.closeNode(); // employee

w2.closeAll();

fs.writeFileSync(path.join(outDir, "nested.xml"), w2.xml);
console.log("\n=== 2. Nested Elements ===");
console.log(w2.xml.substring(0, 200) + "...");

// =============================================================================
// 3. CDATA, comments, raw XML
// =============================================================================

const w3 = new Xml.Writer();
w3.openXml(Xml.StdDocAttributes);
w3.openNode("page");

w3.writeComment("This is a comment");

w3.openNode("script");
w3.writeCData('if (x < 10 && y > 20) { alert("hello"); }');
w3.closeNode();

w3.openNode("html");
w3.writeRaw("<b>Bold</b> and <i>italic</i>");
w3.closeNode();

w3.closeAll();

fs.writeFileSync(path.join(outDir, "special.xml"), w3.xml);
console.log("\n=== 3. CDATA, Comments, Raw XML ===");
console.log(w3.xml);

// =============================================================================
// 4. Save/commit/rollback (transactional writes)
// =============================================================================

const w4 = new Xml.Writer();
w4.openXml(Xml.StdDocAttributes);
w4.openNode("root");
w4.leafNode("kept", undefined, "This stays");

w4.save(); // Save checkpoint

w4.leafNode("tentative", undefined, "This might be removed");
w4.leafNode("also-tentative", undefined, "This too");

w4.rollback(); // Undo back to checkpoint

w4.leafNode("final", undefined, "This is added after rollback");
w4.closeAll();

console.log("\n=== 4. Save/Commit/Rollback ===");
console.log(w4.xml);

// =============================================================================
// 5. Xml.Writer cursor and depth
// =============================================================================

const w5 = new Xml.Writer();
w5.openXml();
w5.openNode("a");
w5.openNode("b");
w5.openNode("c");

console.log("\n=== 5. Cursor and Depth ===");
console.log("Depth:", w5.depth); // 3
console.log("Current element:", w5.currentElement); // "c"

w5.closeAll();
console.log("After closeAll depth:", w5.depth); // 0

// =============================================================================
// 6. Xml.StreamWriter — streaming to file
// =============================================================================

const filePath6 = path.join(outDir, "streamed.xml");
const chunks6: string[] = [];
const fileHandle = fs.openSync(filePath6, "w");
const writeTarget = {
  write(str: string): void {
    fs.writeSync(fileHandle, str);
    chunks6.push(str);
  }
};

const sw6 = new Xml.StreamWriter(writeTarget);
sw6.openXml(Xml.StdDocAttributes);
sw6.openNode("data");
for (let i = 0; i < 5; i++) {
  sw6.leafNode("item", { index: i }, `Value ${i}`);
}
sw6.closeAll();
fs.closeSync(fileHandle);

console.log("\n=== 6. Xml.StreamWriter ===");
console.log("Written to:", filePath6);
console.log(chunks6.join(""));

// =============================================================================
// 7. DOM parsing — Xml.parse
// =============================================================================

const xmlDoc = `<?xml version="1.0" encoding="UTF-8"?>
<library>
  <book id="1" genre="fiction">
    <title>1984</title>
    <author>George Orwell</author>
    <year>1949</year>
  </book>
  <book id="2" genre="science">
    <title>A Brief History of Time</title>
    <author>Stephen Hawking</author>
    <year>1988</year>
  </book>
  <book id="3" genre="fiction">
    <title>Brave New World</title>
    <author>Aldous Huxley</author>
    <year>1932</year>
  </book>
</library>`;

const doc = Xml.parse(xmlDoc);
console.log("\n=== 7. DOM Parsing ===");
console.log("Root element:", doc.root.name);
console.log("Children:", doc.root.children.length);

// =============================================================================
// 8. DOM navigation — Xml.findChild, Xml.findChildren, Xml.attr, Xml.textContent
// =============================================================================

console.log("\n=== 8. DOM Navigation ===");

// Find first book
const firstBook = Xml.findChild(doc.root, "book");
if (firstBook) {
  console.log("First book id:", Xml.attr(firstBook, "id"));
  console.log("First book genre:", Xml.attr(firstBook, "genre"));

  const titleEl = Xml.findChild(firstBook, "title");
  if (titleEl) {
    console.log("Title:", Xml.textContent(titleEl));
  }
}

// Find all books
const allBooks = Xml.findChildren(doc.root, "book");
console.log("Total books:", allBooks.length);
for (const book of allBooks) {
  const title = Xml.findChild(book, "title");
  const author = Xml.findChild(book, "author");
  console.log(`  "${Xml.textContent(title!)}" by ${Xml.textContent(author!)}`);
}

// =============================================================================
// 9. Walk — visit all elements
// =============================================================================

console.log("\n=== 9. Walk ===");

const elementNames: string[] = [];
Xml.walk(doc.root, el => {
  elementNames.push(el.name);
});
console.log("All elements:", elementNames);

// =============================================================================
// 10. Query engine — XPath-like queries
// =============================================================================

console.log("\n=== 10. Query Engine ===");

// Direct path
const firstTitle = Xml.query(doc.root, "book/title");
console.log("First title:", firstTitle ? Xml.textContent(firstTitle) : "not found");

// All titles
const allTitles = Xml.queryAll(doc.root, "book/title");
console.log(
  "All titles:",
  allTitles.map(t => Xml.textContent(t))
);

// Attribute filter
const fictionBooks = Xml.queryAll(doc.root, "book[@genre='fiction']");
console.log(
  "Fiction books:",
  fictionBooks.map(b => Xml.textContent(Xml.findChild(b, "title")!))
);

// Indexed access
const secondBook = Xml.query(doc.root, "book[1]");
if (secondBook) {
  console.log("Second book:", Xml.textContent(Xml.findChild(secondBook, "title")!));
}

// Wildcard
const allChildren = Xml.queryAll(doc.root, "book/*");
console.log("All book children:", allChildren.length, "elements");

// =============================================================================
// 11. Xml.toPlainObject — convert DOM to JS objects
// =============================================================================

console.log("\n=== 11. Xml.toPlainObject ===");

if (firstBook) {
  const obj = Xml.toPlainObject(firstBook);
  console.log("Book as object:", JSON.stringify(obj, null, 2));
}

// =============================================================================
// 12. Xml.parseToObject — direct XML to JS object
// =============================================================================

console.log("\n=== 12. Xml.parseToObject ===");

const simpleXml = `<config>
  <host>localhost</host>
  <port>8080</port>
  <debug>true</debug>
</config>`;

const configObj = Xml.parseToObject(simpleXml);
console.log("Config:", JSON.stringify(configObj, null, 2));

// =============================================================================
// 13. SAX streaming parser
// =============================================================================

console.log("\n=== 13. SAX Parser ===");

const sax = new Xml.SaxParser();
const saxEvents: string[] = [];

sax.on("opentag", tag => {
  saxEvents.push(`OPEN: <${tag.name}> attrs=${JSON.stringify(tag.attributes)}`);
});
sax.on("text", text => {
  const trimmed = text.trim();
  if (trimmed) {
    saxEvents.push(`TEXT: "${trimmed}"`);
  }
});
sax.on("closetag", tag => {
  saxEvents.push(`CLOSE: </${tag.name}>`);
});

sax.write('<root><item id="1">Hello</item><item id="2">World</item></root>');
sax.close();

for (const event of saxEvents.slice(0, 8)) {
  console.log("  " + event);
}
console.log(`  ... (${saxEvents.length} total events)`);

// =============================================================================
// 14. SAX parser — error handling
// =============================================================================

console.log("\n=== 14. SAX Error Handling ===");

const saxErr = new Xml.SaxParser();
let parseError: Error | null = null;
saxErr.on("error", err => {
  parseError = err;
});

saxErr.write("<root><unclosed>");
saxErr.close();

if (parseError) {
  console.log("Parse error caught:", (parseError as Error).message);
}

// =============================================================================
// 15. Encoding utilities
// =============================================================================

console.log("\n=== 15. Encoding Utilities ===");

// Xml.encode: escape special chars in text content
const encoded = Xml.encode('Hello <world> & "quotes"');
console.log("Encoded:", encoded);
// "Hello &lt;world&gt; &amp; &quot;quotes&quot;"

// Xml.decode: unescape back
const decoded = Xml.decode("Hello &lt;world&gt; &amp; &quot;quotes&quot;");
console.log("Decoded:", decoded);
// "Hello <world> & "quotes""

// Xml.encodeAttr: for attribute values
const attrEncoded = Xml.encodeAttr('value with "quotes" & <tags>');
console.log("Attr encoded:", attrEncoded);

// Xml.encodeCData: handle ]]> in CDATA
const cdataEncoded = Xml.encodeCData("Data with ]]> inside");
console.log("CDATA encoded:", cdataEncoded);

// =============================================================================
// 16. XML name validation
// =============================================================================

console.log("\n=== 16. Name Validation ===");

try {
  Xml.validateXmlName("valid-name");
  console.log("'valid-name' is valid");
} catch {
  // won't reach here
}

try {
  Xml.validateXmlName("123invalid");
  console.log("'123invalid' is valid");
} catch (err) {
  if (isXmlError(err)) {
    console.log("'123invalid' is invalid:", err.message);
  }
}

// =============================================================================
// 17. Namespace support
// =============================================================================

const w17 = new Xml.Writer();
w17.openXml(Xml.StdDocAttributes);
w17.openNode("worksheet", {
  xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
});
w17.openNode("sheetData");
w17.openNode("row", { r: 1 });
w17.leafNode("c", { r: "A1", t: "s" }, "0");
w17.closeAll();

console.log("\n=== 17. Namespaces ===");
console.log(w17.xml);

// =============================================================================
// 18. Parse with options — tolerant/strict modes
// =============================================================================

console.log("\n=== 18. Parse Options ===");

// Strict parsing (default)
try {
  const strictDoc = Xml.parse("<root><child>text</child></root>");
  console.log("Strict parse OK, root:", strictDoc.root.name);
} catch (err) {
  if (isXmlParseError(err)) {
    console.log("Parse error at line", err.line, "col", err.column);
  }
}

// With invalid char handling
const xmlWithBadChars = "<root>text with \x01 control char</root>";
try {
  const tolerantDoc = Xml.parse(xmlWithBadChars, { invalidCharHandling: "replace" });
  console.log("Tolerant parse OK:", Xml.textContent(tolerantDoc.root).length, "chars");
} catch (err) {
  console.log("Even tolerant mode failed:", (err as Error).message);
}

// =============================================================================
// 19. Building complex XML — OOXML-like structure
// =============================================================================

const w19 = new Xml.Writer();
w19.openXml(Xml.StdDocAttributes);
w19.openNode("workbook", {
  xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
});

w19.openNode("sheets");
w19.leafNode("sheet", { name: "Sheet1", sheetId: 1, "r:id": "rId1" });
w19.leafNode("sheet", { name: "Sheet2", sheetId: 2, "r:id": "rId2" });
w19.closeNode(); // sheets

w19.openNode("definedNames");
w19.leafNode("definedName", { name: "_xlnm.Print_Area", localSheetId: 0 }, "Sheet1!$A:$H");
w19.closeNode(); // definedNames

w19.closeAll();

const ooxmlPath = path.join(outDir, "workbook.xml");
fs.writeFileSync(ooxmlPath, w19.xml);
console.log("\n=== 19. OOXML-like Structure ===");
console.log(w19.xml.substring(0, 300) + "...");

// =============================================================================
// 20. Round-trip: write → parse → Xml.query → extract
// =============================================================================

console.log("\n=== 20. Round-Trip ===");

const roundTrip = Xml.parse(w19.xml);
const sheets = Xml.queryAll(roundTrip.root, "sheets/sheet");
console.log("Sheets found:", sheets.length);
for (const sheet of sheets) {
  console.log(`  Sheet: name="${Xml.attr(sheet, "name")}" id=${Xml.attr(sheet, "sheetId")}`);
}

const printArea = Xml.query(roundTrip.root, "definedNames/definedName[@name='_xlnm.Print_Area']");
if (printArea) {
  console.log("Print area:", Xml.textContent(printArea));
}

console.log("\n=== XML Examples Complete ===");
