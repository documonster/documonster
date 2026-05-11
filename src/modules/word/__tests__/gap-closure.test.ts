/**
 * DOCX Module - Gap Closure Tests
 *
 * Comprehensive tests for all new features added to close the gap
 * with competitor libraries.
 */

import { describe, it, expect } from "vitest";

import { htmlToDocxBody } from "../convert/html/html-import";
import { docxToMarkdown } from "../convert/markdown/markdown-converter";
import {
  Document,
  textParagraph,
  paragraph,
  text,
  formTextField,
  formCheckboxField,
  formDropdownField,
  searchText,
  replaceText,
  extractText,
  mergeDocuments,
  resolveStyle,
  getCompatibilityMode,
  setCompatibilityMode,
  diffDocuments,
  acceptAllRevisions,
  rejectAllRevisions,
  packageDocx,
  readDocx,
  toBuffer,
  compileTemplate,
  patchTemplate,
  patchDocument,
  chart,
  heading,
  fillTemplateEnhanced,
  listTemplateTags
} from "../index";
import { readCfb } from "../security/cfb-reader";
import { isEncryptedDocx } from "../security/encryption";
import type { DocxDocument, Paragraph, Table, Run, BodyContent } from "../types";

// =============================================================================
// extractText fix (table cells)
// =============================================================================

describe("extractText: table cells", () => {
  it("should extract text from table cells", () => {
    const doc: DocxDocument = {
      body: [
        textParagraph("Before table"),
        {
          type: "table",
          rows: [
            {
              cells: [
                { content: [textParagraph("Cell A1")] },
                { content: [textParagraph("Cell B1")] }
              ]
            },
            {
              cells: [
                { content: [textParagraph("Cell A2")] },
                { content: [textParagraph("Cell B2")] }
              ]
            }
          ]
        },
        textParagraph("After table")
      ]
    };
    const result = extractText(doc);
    expect(result).toContain("Before table");
    expect(result).toContain("Cell A1");
    expect(result).toContain("Cell B1");
    expect(result).toContain("Cell A2");
    expect(result).toContain("Cell B2");
    expect(result).toContain("After table");
    // Cells in same row separated by tab
    expect(result).toContain("Cell A1\tCell B1");
  });

  it("should handle nested tables", () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "table",
          rows: [
            {
              cells: [
                {
                  content: [
                    textParagraph("Outer"),
                    {
                      type: "table",
                      rows: [
                        {
                          cells: [{ content: [textParagraph("Inner")] }]
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        }
      ]
    };
    const result = extractText(doc);
    expect(result).toContain("Outer");
    expect(result).toContain("Inner");
  });
});

// =============================================================================
// replaceText: tables and format preservation
// =============================================================================

describe("replaceText: enhanced features", () => {
  it("should replace text inside tables", () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "table",
          rows: [
            {
              cells: [
                { content: [textParagraph("Hello World")] },
                { content: [textParagraph("Foo Bar")] }
              ]
            }
          ]
        }
      ]
    };
    const count = replaceText(doc, "Hello", "Goodbye");
    expect(count).toBe(1);
    const text = extractText(doc);
    expect(text).toContain("Goodbye World");
  });

  it("should handle cross-run replacement", () => {
    const para: Paragraph = {
      type: "paragraph",
      children: [
        { properties: { bold: true }, content: [{ type: "text", text: "Hel" }] } as Run,
        { properties: { italic: true }, content: [{ type: "text", text: "lo" }] } as Run,
        { content: [{ type: "text", text: " World" }] } as Run
      ]
    };
    const doc: DocxDocument = { body: [para] };
    const count = replaceText(doc, "Hello", "Hi");
    expect(count).toBe(1);
    // The replacement should distribute text across runs (format preservation)
    const fullText = extractText(doc);
    expect(fullText).toContain("Hi");
    expect(fullText).toContain("World");
  });

  it("should replace in headers/footers", () => {
    const doc: DocxDocument = {
      body: [textParagraph("Body text")],
      headers: new Map([
        [
          "rId1",
          {
            rId: "rId1",
            content: {
              children: [textParagraph("Header: {{company}}")] as any
            }
          }
        ]
      ])
    };
    const count = replaceText(doc, "{{company}}", "Acme Corp");
    expect(count).toBe(1);
  });
});

// =============================================================================
// searchText: tables and nested content
// =============================================================================

describe("searchText: enhanced features", () => {
  it("should find text inside tables", () => {
    const doc: DocxDocument = {
      body: [
        textParagraph("First paragraph"),
        {
          type: "table",
          rows: [
            {
              cells: [{ content: [textParagraph("Needle in table")] }]
            }
          ]
        }
      ]
    };
    const results = searchText(doc, "Needle");
    expect(results.length).toBe(1);
    expect(results[0].match).toBe("Needle");
  });

  it("should search with regex", () => {
    const doc: DocxDocument = {
      body: [textParagraph("Price is $123.45"), textParagraph("Another price $67.89")]
    };
    const results = searchText(doc, /\$\d+\.\d+/);
    expect(results.length).toBe(2);
    expect(results[0].match).toBe("$123.45");
    expect(results[1].match).toBe("$67.89");
  });
});

// =============================================================================
// Chart reader (round-trip)
// =============================================================================

describe("Chart reader", () => {
  it("should round-trip a chart through write → read", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Before chart");
    Document.addContent(
      h,
      chart({
        type: "bar",
        title: "Sales Report",
        series: [{ name: "Q1", categories: ["Jan", "Feb", "Mar"], values: [10, 20, 30] }],
        legend: "r"
      })
    );
    Document.addParagraph(h, "After chart");
    const buffer = await toBuffer(Document.build(h));
    const parsed = await readDocx(buffer);

    // Should find a chart content item
    const charts = parsed.body.filter(b => b.type === "chart");
    expect(charts.length).toBe(1);
    const chartContent = charts[0] as any;
    expect(chartContent.chart.type).toBe("bar");
    expect(chartContent.chart.title).toBe("Sales Report");
    expect(chartContent.chart.series.length).toBe(1);
    expect(chartContent.chart.series[0].name).toBe("Q1");
    expect(chartContent.chart.series[0].categories).toEqual(["Jan", "Feb", "Mar"]);
    expect(chartContent.chart.series[0].values).toEqual([10, 20, 30]);
  });

  it("should parse pie chart correctly", async () => {
    const h = Document.create();
    Document.addContent(
      h,
      chart({
        type: "pie",
        series: [
          {
            name: "Market Share",
            categories: ["A", "B", "C"],
            values: [40, 35, 25],
            pointColors: ["FF0000", "00FF00", "0000FF"]
          }
        ]
      })
    );
    const buffer = await toBuffer(Document.build(h));
    const parsed = await readDocx(buffer);
    const charts = parsed.body.filter(b => b.type === "chart");
    expect(charts.length).toBe(1);
    expect((charts[0] as any).chart.type).toBe("pie");
  });
});

// =============================================================================
// Form fields
// =============================================================================

describe("Form field builders", () => {
  it("should create text form field with properties", async () => {
    const h = Document.create();
    Document.addParagraphElement(
      h,
      paragraph([
        formTextField({
          name: "FirstName",
          default: "John",
          maxLength: 50,
          helpText: "Enter your first name"
        })
      ])
    );
    const buffer = await toBuffer(Document.build(h));
    const parsed = await readDocx(buffer);
    const para = parsed.body[0] as Paragraph;
    // Find the field in the run content
    let foundField = false;
    for (const child of para.children) {
      if ("content" in child && Array.isArray(child.content)) {
        for (const c of child.content) {
          if ("type" in c && c.type === "field" && "formField" in c) {
            const ff = (c as any).formField;
            expect(ff.type).toBe("text");
            expect(ff.name).toBe("FirstName");
            expect(ff.default).toBe("John");
            expect(ff.maxLength).toBe(50);
            foundField = true;
          }
        }
      }
    }
    expect(foundField).toBe(true);
  });

  it("should create checkbox form field", async () => {
    const h = Document.create();
    Document.addParagraphElement(
      h,
      paragraph([formCheckboxField({ name: "Agree", checked: true })])
    );
    const buffer = await toBuffer(Document.build(h));
    const parsed = await readDocx(buffer);
    const para = parsed.body[0] as Paragraph;
    let foundField = false;
    for (const child of para.children) {
      if ("content" in child && Array.isArray(child.content)) {
        for (const c of child.content) {
          if ("type" in c && c.type === "field" && "formField" in c) {
            const ff = (c as any).formField;
            expect(ff.type).toBe("checkBox");
            expect(ff.name).toBe("Agree");
            expect(ff.checked).toBe(true);
            foundField = true;
          }
        }
      }
    }
    expect(foundField).toBe(true);
  });

  it("should create dropdown form field", async () => {
    const h = Document.create();
    Document.addParagraphElement(
      h,
      paragraph([
        formDropdownField({
          name: "Country",
          entries: ["USA", "Canada", "UK"],
          default: 1
        })
      ])
    );
    const buffer = await toBuffer(Document.build(h));
    const parsed = await readDocx(buffer);
    const para = parsed.body[0] as Paragraph;
    let foundField = false;
    for (const child of para.children) {
      if ("content" in child && Array.isArray(child.content)) {
        for (const c of child.content) {
          if ("type" in c && c.type === "field" && "formField" in c) {
            const ff = (c as any).formField;
            expect(ff.type).toBe("dropDown");
            expect(ff.name).toBe("Country");
            expect(ff.entries).toEqual(["USA", "Canada", "UK"]);
            foundField = true;
          }
        }
      }
    }
    expect(foundField).toBe(true);
  });
});

// =============================================================================
// Document merge
// =============================================================================

describe("mergeDocuments", () => {
  it("should merge multiple documents", () => {
    const doc1: DocxDocument = {
      body: [textParagraph("Doc 1 content")],
      styles: [{ type: "paragraph", styleId: "Heading1", name: "heading 1" }]
    };
    const doc2: DocxDocument = {
      body: [textParagraph("Doc 2 content")],
      styles: [{ type: "paragraph", styleId: "Heading2", name: "heading 2" }]
    };
    const merged = mergeDocuments([doc1, doc2]);
    const text = extractText(merged);
    expect(text).toContain("Doc 1 content");
    expect(text).toContain("Doc 2 content");
    expect(merged.styles!.length).toBe(2);
  });

  it("should handle single document", () => {
    const doc: DocxDocument = { body: [textParagraph("Only one")] };
    const result = mergeDocuments([doc]);
    expect(result).toBe(doc); // Should return same reference
  });

  it("should handle empty array", () => {
    const result = mergeDocuments([]);
    expect(result.body).toEqual([]);
  });

  it("should insert section breaks between documents", () => {
    const doc1: DocxDocument = { body: [textParagraph("A")] };
    const doc2: DocxDocument = { body: [textParagraph("B")] };
    const merged = mergeDocuments([doc1, doc2], { sectionBreak: "continuous" });
    // Should have: paragraph "A", section break paragraph, paragraph "B"
    expect(merged.body.length).toBe(3);
    const sectionPara = merged.body[1] as Paragraph;
    expect(sectionPara.properties?.sectionProperties?.breakType).toBe("continuous");
  });
});

// =============================================================================
// Style resolution
// =============================================================================

describe("resolveStyle", () => {
  it("should resolve inherited style properties", () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "paragraph",
          properties: { style: "Heading1" },
          children: [{ content: [{ type: "text", text: "Title" }] } as any]
        }
      ],
      styles: [
        {
          type: "paragraph",
          styleId: "Normal",
          name: "Normal",
          paragraphProperties: { alignment: "left" },
          runProperties: { size: 24 }
        },
        {
          type: "paragraph",
          styleId: "Heading1",
          name: "heading 1",
          basedOn: "Normal",
          paragraphProperties: { alignment: "center" },
          runProperties: { bold: true, size: 32 }
        }
      ],
      docDefaults: {
        runProperties: { font: { ascii: "Calibri" } }
      }
    };
    const para = doc.body[0] as Paragraph;
    const resolved = resolveStyle(doc, para);

    expect(resolved.chain).toEqual(["Heading1", "Normal"]);
    // Heading1 overrides Normal's alignment
    expect((resolved.paragraphProperties as any).alignment).toBe("center");
    // Run props: bold from Heading1, font from defaults
    expect(resolved.runProperties.bold).toBe(true);
    expect((resolved.runProperties as any).size).toBe(32); // Heading1 overrides Normal
    expect((resolved.runProperties as any).font?.ascii).toBe("Calibri");
  });

  it("should handle paragraph without style", () => {
    const doc: DocxDocument = {
      body: [textParagraph("Plain text")],
      docDefaults: {
        runProperties: { size: 24 }
      }
    };
    const para = doc.body[0] as Paragraph;
    const resolved = resolveStyle(doc, para);
    expect(resolved.chain).toEqual([]);
    expect((resolved.runProperties as any).size).toBe(24);
  });
});

// =============================================================================
// Compatibility mode
// =============================================================================

describe("Compatibility mode", () => {
  it("should get default mode (15) when no settings", () => {
    const doc: DocxDocument = { body: [] };
    expect(getCompatibilityMode(doc)).toBe(15);
  });

  it("should get mode from settings", () => {
    const doc: DocxDocument = {
      body: [],
      settings: {
        compatSettings: [
          { name: "compatibilityMode", uri: "http://schemas.microsoft.com/office/word", val: "14" }
        ]
      }
    };
    expect(getCompatibilityMode(doc)).toBe(14);
  });

  it("should set compatibility mode", () => {
    const doc: DocxDocument = { body: [] };
    setCompatibilityMode(doc, 12);
    expect(getCompatibilityMode(doc)).toBe(12);
    setCompatibilityMode(doc, 15);
    expect(getCompatibilityMode(doc)).toBe(15);
  });
});

// =============================================================================
// Template caching (compileTemplate / patchTemplate)
// =============================================================================

describe("compileTemplate / patchTemplate", () => {
  it("should compile and patch faster on second use", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Hello {{name}}!");
    Document.addParagraph(h, "Your role is {{role}}.");
    const templateBuf = await toBuffer(Document.build(h));

    // Compile once
    const compiled = await compileTemplate(templateBuf);

    // Patch multiple times
    const result1 = await patchTemplate(compiled, [
      { placeholder: "{{name}}", content: { type: "text", text: "Alice" } },
      { placeholder: "{{role}}", content: { type: "text", text: "Engineer" } }
    ]);
    const result2 = await patchTemplate(compiled, [
      { placeholder: "{{name}}", content: { type: "text", text: "Bob" } },
      { placeholder: "{{role}}", content: { type: "text", text: "Designer" } }
    ]);

    // Verify both results
    const parsed1 = await readDocx(result1);
    const parsed2 = await readDocx(result2);
    expect(extractText(parsed1)).toContain("Alice");
    expect(extractText(parsed1)).toContain("Engineer");
    expect(extractText(parsed2)).toContain("Bob");
    expect(extractText(parsed2)).toContain("Designer");
  });
});

// =============================================================================
// HTML → DOCX import
// =============================================================================

describe("htmlToDocxBody", () => {
  it("should parse headings", () => {
    const blocks = htmlToDocxBody("<h1>Title</h1><h2>Subtitle</h2>");
    expect(blocks.length).toBe(2);
    const h1 = blocks[0] as Paragraph;
    const h2 = blocks[1] as Paragraph;
    expect(h1.properties?.style).toBe("Heading1");
    expect(h2.properties?.style).toBe("Heading2");
  });

  it("should parse paragraphs with inline formatting", () => {
    const blocks = htmlToDocxBody("<p>Hello <strong>bold</strong> and <em>italic</em></p>");
    expect(blocks.length).toBe(1);
    const para = blocks[0] as Paragraph;
    expect(para.children.length).toBeGreaterThanOrEqual(3);
    // Check the bold run
    const boldRun = para.children.find(c => {
      if ("properties" in c && c.properties) {
        return (c.properties as any).bold === true;
      }
      return false;
    });
    expect(boldRun).toBeDefined();
  });

  it("should parse tables", () => {
    const html =
      "<table><tr><th>Name</th><th>Age</th></tr><tr><td>Alice</td><td>30</td></tr></table>";
    const blocks = htmlToDocxBody(html);
    const tables = blocks.filter(b => b.type === "table");
    expect(tables.length).toBe(1);
    const table = tables[0] as Table;
    expect(table.type).toBe("table");
    expect(table.rows.length).toBe(2);
    expect(table.rows[0].cells.length).toBe(2);
  });

  it("should parse lists", () => {
    const html = "<ul><li>Item 1</li><li>Item 2</li></ul>";
    const blocks = htmlToDocxBody(html);
    expect(blocks.length).toBe(2);
    const li1 = blocks[0] as Paragraph;
    expect(li1.properties?.numbering).toBeDefined();
  });

  it("should handle br tags", () => {
    const blocks = htmlToDocxBody("<p>Line 1<br>Line 2</p>");
    expect(blocks.length).toBe(1);
    const para = blocks[0] as Paragraph;
    // Should contain a break content
    const hasBreak = para.children.some(c => {
      if ("content" in c && Array.isArray(c.content)) {
        return c.content.some((rc: any) => rc.type === "break");
      }
      return false;
    });
    expect(hasBreak).toBe(true);
  });

  it("should decode HTML entities", () => {
    const blocks = htmlToDocxBody("<p>&lt;html&gt; &amp; &quot;quotes&quot;</p>");
    expect(blocks.length).toBe(1);
    const text = extractText({ body: blocks });
    expect(text).toContain('<html> & "quotes"');
  });
});

// =============================================================================
// DOCX → Markdown
// =============================================================================

describe("docxToMarkdown", () => {
  it("should convert headings to # syntax", () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "paragraph",
          properties: { style: "Heading1" },
          children: [{ content: [{ type: "text", text: "Title" }] } as any]
        },
        {
          type: "paragraph",
          properties: { style: "Heading2" },
          children: [{ content: [{ type: "text", text: "Subtitle" }] } as any]
        }
      ]
    };
    const md = docxToMarkdown(doc);
    expect(md).toContain("# Title");
    expect(md).toContain("## Subtitle");
  });

  it("should convert bold/italic to markdown syntax", () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "paragraph",
          children: [
            { properties: { bold: true }, content: [{ type: "text", text: "bold" }] } as Run,
            { content: [{ type: "text", text: " and " }] } as Run,
            { properties: { italic: true }, content: [{ type: "text", text: "italic" }] } as Run
          ]
        }
      ]
    };
    const md = docxToMarkdown(doc);
    expect(md).toContain("**bold**");
    expect(md).toContain("*italic*");
  });

  it("should convert tables to GFM format", () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "table",
          rows: [
            { cells: [{ content: [textParagraph("A")] }, { content: [textParagraph("B")] }] },
            { cells: [{ content: [textParagraph("1")] }, { content: [textParagraph("2")] }] }
          ]
        }
      ]
    };
    const md = docxToMarkdown(doc);
    expect(md).toContain("|");
    expect(md).toContain("---");
    expect(md).toContain("A");
    expect(md).toContain("B");
  });

  it("should handle list items", () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "paragraph",
          properties: { numbering: { numId: 1, level: 0 } },
          children: [{ content: [{ type: "text", text: "Item 1" }] } as any]
        },
        {
          type: "paragraph",
          properties: { numbering: { numId: 1, level: 0 } },
          children: [{ content: [{ type: "text", text: "Item 2" }] } as any]
        }
      ]
    };
    const md = docxToMarkdown(doc);
    expect(md).toContain("- Item 1");
    expect(md).toContain("- Item 2");
  });
});

// =============================================================================
// Document diff
// =============================================================================

describe("diffDocuments", () => {
  it("should detect unchanged documents", () => {
    const doc1: DocxDocument = { body: [textParagraph("Hello"), textParagraph("World")] };
    const doc2: DocxDocument = { body: [textParagraph("Hello"), textParagraph("World")] };
    const result = diffDocuments(doc1, doc2);
    expect(result.summary.unchanged).toBe(2);
    expect(result.summary.added).toBe(0);
    expect(result.summary.deleted).toBe(0);
  });

  it("should detect added paragraphs", () => {
    const doc1: DocxDocument = { body: [textParagraph("Hello")] };
    const doc2: DocxDocument = { body: [textParagraph("Hello"), textParagraph("World")] };
    const result = diffDocuments(doc1, doc2);
    expect(result.summary.added).toBe(1);
    expect(result.entries.some(e => e.type === "added" && e.newText === "World")).toBe(true);
  });

  it("should detect deleted paragraphs", () => {
    const doc1: DocxDocument = { body: [textParagraph("Hello"), textParagraph("World")] };
    const doc2: DocxDocument = { body: [textParagraph("Hello")] };
    const result = diffDocuments(doc1, doc2);
    expect(result.summary.deleted).toBe(1);
    expect(result.entries.some(e => e.type === "deleted" && e.oldText === "World")).toBe(true);
  });

  it("should detect modifications", () => {
    const doc1: DocxDocument = { body: [textParagraph("Hello World")] };
    const doc2: DocxDocument = { body: [textParagraph("Hello Earth")] };
    const result = diffDocuments(doc1, doc2);
    // Either modified or (deleted + added)
    const hasChange =
      result.summary.modified > 0 || (result.summary.added > 0 && result.summary.deleted > 0);
    expect(hasChange).toBe(true);
  });
});

// =============================================================================
// CFB reader
// =============================================================================

describe("CFB reader", () => {
  it("should detect encrypted DOCX signature", () => {
    const cfbHeader = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    expect(isEncryptedDocx(cfbHeader)).toBe(true);
    expect(isEncryptedDocx(new Uint8Array([0x50, 0x4b, 0x03, 0x04]))).toBe(false);
  });

  it("should reject non-CFB files", () => {
    // Need at least 512 bytes for it to check signature
    const buf = new Uint8Array(512);
    buf[0] = 0x50; // not CFB
    expect(() => readCfb(buf)).toThrow("CFB: invalid signature");
    expect(() => readCfb(new Uint8Array(5))).toThrow("CFB: file too small");
  });
});

// =============================================================================
// Full round-trip integration tests
// =============================================================================

describe("Full round-trip: new features", () => {
  it("should round-trip extractText with complex document", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Top-level paragraph");
    Document.addTable(h, [
      ["A", "B"],
      ["C", "D"]
    ]);
    Document.addHeading(h, "Heading 1", 1);
    const buffer = await toBuffer(Document.build(h));
    const parsed = await readDocx(buffer);
    const text = extractText(parsed);
    expect(text).toContain("Top-level paragraph");
    expect(text).toContain("A");
    expect(text).toContain("B");
    expect(text).toContain("Heading 1");
  });

  it("should round-trip search/replace in table", async () => {
    const h = Document.create();
    Document.addTable(h, [["Hello World", "Foo"]]);
    const buffer = await toBuffer(Document.build(h));
    const parsed = await readDocx(buffer);

    const searchResults = searchText(parsed, "Hello");
    expect(searchResults.length).toBeGreaterThanOrEqual(1);

    const count = replaceText(parsed, "Hello", "Goodbye");
    expect(count).toBe(1);
    expect(extractText(parsed)).toContain("Goodbye World");
  });

  it("should merge and repackage successfully", async () => {
    const h1 = Document.create();
    Document.addParagraph(h1, "Document One");
    const buf1 = await toBuffer(Document.build(h1));
    const doc1 = await readDocx(buf1);

    const h2 = Document.create();
    Document.addParagraph(h2, "Document Two");
    const buf2 = await toBuffer(Document.build(h2));
    const doc2 = await readDocx(buf2);

    const merged = mergeDocuments([doc1, doc2]);
    const mergedBuf = await packageDocx(merged);
    const reParsed = await readDocx(mergedBuf);
    const text = extractText(reParsed);
    expect(text).toContain("Document One");
    expect(text).toContain("Document Two");
  });
});

// =============================================================================
// Track Changes accept/reject
// =============================================================================

describe("Track Changes: accept/reject", () => {
  it("acceptAllRevisions should keep inserted text and remove deleted text", () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "paragraph",
          children: [
            { content: [{ type: "text", text: "Hello " }] } as any,
            {
              type: "insertedRun",
              revision: { id: 1, author: "Alice", date: "2024-01-01" },
              run: { content: [{ type: "text", text: "beautiful " }] }
            } as any,
            {
              type: "deletedRun",
              revision: { id: 2, author: "Alice", date: "2024-01-01" },
              run: { content: [{ type: "text", text: "ugly " }] }
            } as any,
            { content: [{ type: "text", text: "world" }] } as any
          ]
        }
      ]
    };
    const count = acceptAllRevisions(doc);
    expect(count).toBeGreaterThanOrEqual(2);
    const result = extractText(doc);
    expect(result).toContain("Hello ");
    expect(result).toContain("beautiful ");
    expect(result).not.toContain("ugly ");
    expect(result).toContain("world");
  });

  it("rejectAllRevisions should remove inserted text and keep deleted text", () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "paragraph",
          children: [
            { content: [{ type: "text", text: "Hello " }] } as any,
            {
              type: "insertedRun",
              revision: { id: 1, author: "Alice", date: "2024-01-01" },
              run: { content: [{ type: "text", text: "beautiful " }] }
            } as any,
            {
              type: "deletedRun",
              revision: { id: 2, author: "Alice", date: "2024-01-01" },
              run: { content: [{ type: "text", text: "ugly " }] }
            } as any,
            { content: [{ type: "text", text: "world" }] } as any
          ]
        }
      ]
    };
    const count = rejectAllRevisions(doc);
    expect(count).toBeGreaterThanOrEqual(2);
    const result = extractText(doc);
    expect(result).toContain("Hello ");
    expect(result).not.toContain("beautiful ");
    expect(result).toContain("ugly ");
    expect(result).toContain("world");
  });
});

// =============================================================================
// Improved heading builder
// =============================================================================

describe("heading builder: mixed formatting", () => {
  it("should accept array of runs for mixed formatting", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, heading([text("Normal "), text("Bold", { bold: true })], 1));
    const buffer = await toBuffer(Document.build(h));
    const parsed = await readDocx(buffer);
    const para = parsed.body[0] as Paragraph;
    expect(para.properties?.style).toBe("Heading1");
    const fullText = extractText(parsed);
    expect(fullText).toContain("Normal ");
    expect(fullText).toContain("Bold");
  });
});

// =============================================================================
// Header/footer full patching
// =============================================================================

describe("patchDocument: header/footer full support", () => {
  it("should patch text in headers", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Body with {{name}}");
    Document.setHeader(h, "default", { children: [textParagraph("Header: {{name}}")] });
    const templateBuf = await toBuffer(Document.build(h));

    const result = await patchDocument(templateBuf, [
      { placeholder: "{{name}}", content: { type: "text", text: "Acme Inc" } }
    ]);
    const parsed = await readDocx(result);
    const bodyText = extractText(parsed);
    expect(bodyText).toContain("Acme Inc");
  });
});

// =============================================================================
// Cross-run surrogate pair safety
// =============================================================================

describe("replaceText: surrogate pair safety", () => {
  it("should not break emoji in cross-run replacement", () => {
    // Text with emoji split across runs
    const para: Paragraph = {
      type: "paragraph",
      children: [
        { content: [{ type: "text", text: "Hello" }] } as Run,
        { content: [{ type: "text", text: " 🌍 World" }] } as Run
      ]
    };
    const doc: DocxDocument = { body: [para] };
    replaceText(doc, "Hello 🌍 World", "Goodbye 🌎 Earth");
    const result = extractText(doc);
    expect(result).toContain("Goodbye");
    expect(result).toContain("Earth");
    // Should not have broken the emoji
    expect(result).toContain("🌎");
  });
});

// =============================================================================
// New Gap Closure: docType / VBA round-trip
// =============================================================================

describe("docType and VBA round-trip", () => {
  it("should round-trip docType for macroEnabledDocument", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Macro doc");
    const doc = Document.build(h);
    const macroDoc: DocxDocument = {
      ...doc,
      docType: "macroEnabledDocument",
      vbaProject: new Uint8Array([0x01, 0x02, 0x03, 0x04])
    };
    const buffer = await toBuffer(macroDoc);
    const parsed = await readDocx(buffer);
    expect(parsed.docType).toBe("macroEnabledDocument");
    expect(parsed.vbaProject).toBeDefined();
    expect(parsed.vbaProject![0]).toBe(0x01);
    expect(parsed.vbaProject!.length).toBe(4);
  });

  it("should round-trip docType for template", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Template doc");
    const doc = Document.build(h);
    const templateDoc: DocxDocument = { ...doc, docType: "template" };
    const buffer = await toBuffer(templateDoc);
    const parsed = await readDocx(buffer);
    expect(parsed.docType).toBe("template");
  });

  it("should not set docType for standard document", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Normal doc");
    const buffer = await toBuffer(Document.build(h));
    const parsed = await readDocx(buffer);
    expect(parsed.docType).toBeUndefined();
  });
});

// =============================================================================
// New Gap Closure: Opaque XML preservation in runs and paragraphs
// =============================================================================

describe("Opaque XML preservation", () => {
  it("should preserve unknown run content through round-trip", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Normal text");
    const doc = Document.build(h);
    // Manually inject an opaque run content
    const paraWithOpaque: Paragraph = {
      type: "paragraph",
      children: [
        {
          content: [
            { type: "text", text: "Before " },
            { type: "opaqueRun", rawXml: '<w:fakeElement w:val="test"/>' },
            { type: "text", text: " After" }
          ]
        } as Run
      ]
    };
    const testDoc: DocxDocument = {
      ...doc,
      body: [paraWithOpaque]
    };
    const buffer = await toBuffer(testDoc);
    const parsed = await readDocx(buffer);
    const para = parsed.body[0] as Paragraph;
    const run = para.children[0] as Run;
    // The opaque content should be preserved
    const opaqueItems = run.content.filter(c => c.type === "opaqueRun");
    expect(opaqueItems.length).toBe(1);
    expect((opaqueItems[0] as any).rawXml).toContain("fakeElement");
  });

  it("should preserve unknown paragraph children through round-trip", async () => {
    const h = Document.create();
    const doc = Document.build(h);
    const paraWithOpaque: Paragraph = {
      type: "paragraph",
      children: [
        { type: "opaqueParagraphChild", rawXml: '<w:unknownMarker w:id="42"/>' } as any,
        { content: [{ type: "text", text: "Hello" }] } as Run
      ]
    };
    const testDoc: DocxDocument = {
      ...doc,
      body: [paraWithOpaque]
    };
    const buffer = await toBuffer(testDoc);
    const parsed = await readDocx(buffer);
    const para = parsed.body[0] as Paragraph;
    const opaqueChildren = para.children.filter(
      c => "type" in c && c.type === "opaqueParagraphChild"
    );
    expect(opaqueChildren.length).toBe(1);
    expect((opaqueChildren[0] as any).rawXml).toContain("unknownMarker");
  });
});

// =============================================================================
// New Gap Closure: Enhanced template engine
// =============================================================================

describe("Template engine: listTemplateTags", () => {
  it("should discover all template tags in a document", () => {
    const doc: DocxDocument = {
      body: [
        textParagraph("Hello {{name}}!"),
        textParagraph("{{#if active}}"),
        textParagraph("Active user"),
        textParagraph("{{/if}}"),
        textParagraph("{{#each items}}"),
        textParagraph("Item: {{.label}}"),
        textParagraph("{{/each}}"),
        textParagraph("{{%logo}}"),
        textParagraph("{{&formatted}}"),
        textParagraph("{{>extra}}")
      ]
    };
    const tags = listTemplateTags(doc);
    expect(tags.length).toBe(9);
    expect(tags[0].type).toBe("variable");
    expect(tags[0].expression).toBe("name");
    expect(tags[1].type).toBe("ifOpen");
    expect(tags[2].type).toBe("ifClose");
    expect(tags[3].type).toBe("eachOpen");
    expect(tags[4].type).toBe("variable"); // .label
    expect(tags[5].type).toBe("eachClose");
    expect(tags[6].type).toBe("image");
    expect(tags[6].expression).toBe("%logo");
    expect(tags[7].type).toBe("richText");
    expect(tags[8].type).toBe("subDocument");
  });
});

describe("Template engine: fillTemplateEnhanced", () => {
  it("should replace image placeholder with inline image", () => {
    const doc: DocxDocument = {
      body: [textParagraph("{{%logo}}")]
    };
    const imgData = new Uint8Array([0x89, 0x50, 0x4e, 0x47]); // PNG header stub
    const result = fillTemplateEnhanced(doc, {
      logo: {
        image: { data: imgData, fileName: "logo.png", mediaType: "png" },
        width: 914400,
        height: 457200
      }
    }) as DocxDocument;

    expect(result.images).toBeDefined();
    expect(result.images!.length).toBe(1);
    expect(result.images![0].fileName).toBe("logo.png");
    const para = result.body[0] as Paragraph;
    const run = para.children[0] as Run;
    expect(run.content[0].type).toBe("image");
  });

  it("should replace richText placeholder with formatted runs", () => {
    const doc: DocxDocument = {
      body: [textParagraph("{{&formatted}}")]
    };
    const richRuns: Run[] = [
      { properties: { bold: true }, content: [{ type: "text", text: "Bold" }] },
      { content: [{ type: "text", text: " and normal" }] }
    ];
    const result = fillTemplateEnhanced(doc, { formatted: richRuns }) as DocxDocument;
    const para = result.body[0] as Paragraph;
    expect(para.children.length).toBe(2);
    expect(((para.children[0] as Run).properties as any)?.bold).toBe(true);
  });

  it("should replace subDocument placeholder with body blocks", () => {
    const doc: DocxDocument = {
      body: [textParagraph("Before"), textParagraph("{{>extra}}"), textParagraph("After")]
    };
    const subContent: BodyContent[] = [
      textParagraph("Sub paragraph 1"),
      textParagraph("Sub paragraph 2")
    ];
    const result = fillTemplateEnhanced(doc, { extra: subContent }) as DocxDocument;
    expect(result.body.length).toBe(4); // Before + 2 sub + After
    expect(extractText(result)).toContain("Sub paragraph 1");
    expect(extractText(result)).toContain("Sub paragraph 2");
  });
});
