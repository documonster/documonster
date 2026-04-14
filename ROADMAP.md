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

## DOCX Module (Future)

### 🔥 P1: High Priority

#### Core Reading

- [ ] Parse DOCX documents (Open XML WordprocessingML)
- [ ] Extract text content with formatting (bold, italic, underline, colors)
- [ ] Extract paragraphs, headings, and sections
- [ ] Extract tables with cell styles and merges
- [ ] Extract images and embedded objects

#### Core Writing

- [ ] Create DOCX documents from structured data
- [ ] Paragraph styling (fonts, colors, alignment, spacing)
- [ ] Table creation with styles and merges
- [ ] Image embedding (JPEG, PNG)
- [ ] Header/footer support

### 🟡 P2: Medium Priority

#### Advanced Features

- [ ] Lists (numbered, bulleted, multi-level)
- [ ] Table of contents generation
- [ ] Page setup (size, orientation, margins)
- [ ] Section breaks and page breaks
- [ ] Footnotes and endnotes
- [ ] Hyperlinks
- [ ] Document metadata

#### Conversion

- [ ] DOCX to PDF conversion
- [ ] DOCX to Markdown conversion
- [ ] Excel to DOCX table embedding
