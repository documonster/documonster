# ExcelTS Roadmap

## Excel Module

### 🔥 P1: High Priority

#### Chart Support

- [ ] Create basic charts (bar, line, pie)
- [ ] Chart customization (title, legend, colors)

#### Stream/Performance

- [ ] Streaming shared strings (currently built in memory, then flushed)

#### Formula

- [ ] Dynamic array formula support (`FILTER`, `SORT`, `UNIQUE`, `XLOOKUP`, `SEQUENCE`) — requires `cm` attribute on `<c>`, `xl/metadata.xml` parsing/writing, `<xda:dynamicArrayProperties>` handling

### 🟡 P2: Medium Priority

#### Image Handling

- [ ] Basic shape support (rectangles, arrows)
- [ ] SVG image support

#### Data Validation

- [ ] Dynamic list formulae support (INDIRECT/OFFSET in list validation)

#### Comment/Note

- [ ] Configurable note size (width/height)

#### Column Key

- [ ] Nested property path support (e.g. `"address.city"`)

#### Hyperlink

- [ ] Hyperlink auto-styling (blue underline for cells with hyperlinks)

### 🟢 P3: Low Priority

#### Formula

- [ ] Formula parsing/validation engine

#### Protection

- [ ] Runtime enforcement of cell `locked` property (currently serialization-only)

---

## Word Module ✅ Completed

### Core Reading ✅

- [x] Parse DOCX documents (Open XML WordprocessingML)
- [x] Extract text content with formatting (bold, italic, underline, colors, highlights, shading)
- [x] Extract paragraphs, headings, and sections
- [x] Extract tables with cell styles, borders, merges (vMerge/gridSpan), nested tables
- [x] Extract images (inline + floating) and embedded objects (opaque preservation)

### Core Writing ✅

- [x] Create DOCX documents from structured data (DocumentBuilder fluent API)
- [x] Paragraph styling (fonts, colors, alignment, spacing, indentation, borders, shading)
- [x] Table creation with styles, borders, merges (horizontal + vertical)
- [x] Image embedding (JPEG, PNG, GIF, BMP, TIFF, SVG with raster fallback)
- [x] Header/footer support (default, first, even)

### Advanced Features ✅

- [x] Lists (numbered, bulleted, multi-level, picture bullets, mail-merge-compatible)
- [x] Table of contents with field switches (headingStyleRange, captionLabel, stylesWithLevels)
- [x] Page setup (size, orientation, margins, columns, borders including art borders)
- [x] Section breaks and page breaks
- [x] Footnotes, endnotes, and comments (with commentsExtended/done/parentId)
- [x] Hyperlinks (external + internal bookmarks, with history/tgtFrame)
- [x] Document metadata (core/app/custom properties)
- [x] Track changes (insertedRun/deletedRun/movedRun + pPrChange/rPrChange/sectPrChange/tblPrChange/trPrChange/tcPrChange/cellIns/cellDel/cellMerge)
- [x] Mail merge (MERGEFIELD + removeUnmatched option)
- [x] Template patching (patchDocument API)
- [x] Drawing shapes (wsp:wsp) with preset geometries
- [x] Math equations (OMML: fraction, sub/sup, radical, matrix, nary, phantom, groupChar, borderBox, etc.)
- [x] Ruby (phonetic annotations for CJK text)
- [x] East Asian support (kinsoku, topLinePunct, autoSpaceDN)
- [x] Theme support (color scheme + font scheme + format scheme)
- [x] Font embedding (ODTTF with obfuscation)
- [x] Charts (opaque preservation + from-scratch builder for bar/line/pie/area/scatter)
- [x] Watermarks (text + image)
- [x] SDT/Content controls with data binding + Custom XML parts
- [x] Document encryption (agile encryption primitives)
- [x] Digital signatures (metadata extraction)
- [x] AltChunk (embedded HTML/RTF content)

### Conversion ✅

- [x] DOCX to HTML conversion (semantic HTML5 with CSS)
- [x] DOCX to PDF conversion (flow-based layout)
- [ ] DOCX to Markdown conversion
- [ ] Excel to DOCX table embedding
