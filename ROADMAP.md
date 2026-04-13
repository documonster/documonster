# ExcelTS Roadmap

## Excel Module

### 🔥 P1: High Priority

#### Chart Support

- [ ] Preserve existing charts when reading/writing
- [ ] Create basic charts (bar, line, pie)
- [ ] Chart customization (title, legend, colors)

#### Stream/Performance

- [ ] Streaming shared strings
- [ ] Memory optimization for large files

#### Style/Formatting

- [ ] Fix font style bleeding to other cells
- [ ] Fix default font not applied correctly
- [ ] Fix styles lost when reading/writing files

#### Image Handling

- [ ] Header/footer image support
- [ ] Basic shape support (rectangles, arrows)
- [ ] Fix image position for non-integer coordinates
- [ ] SVG image support

### 🟡 P2: Medium Priority

#### Formula

- [ ] Fix CSP (Content Security Policy) violations
- [ ] Fix formula parsing errors
- [ ] Improve shared formula handling

#### Print/Page Setup

- [ ] Fix header/footer options
- [ ] Print area improvements

#### Data Validation

- [ ] Dynamic list formulae support
- [ ] Fix template data validation lost
- [ ] Fix protection locked not preventing editing

#### Table Features

- [ ] Append rows to existing table
- [ ] Fix table corruption on save
- [ ] Load tables correctly from disk

#### Protection & Security

- [ ] Fix cell protection not working
- [ ] Column key nested property path support

#### Comment/Note

- [ ] Configurable note size (width/height)
- [ ] Rich text in notes

### 🟢 P3: Low Priority

#### Merge Cell

- [ ] Fix date timezone issue in merged cells
- [ ] Fix addRow() not adding merged cells correctly

#### Row/Column Operations

- [ ] Outline/grouping improvements
- [ ] Hidden row/column improvements

#### Sheet Operations

- [ ] Sheet tab color
- [ ] Sheet copy improvements

#### Hyperlink

- [ ] Hyperlink styling support
- [ ] Internal hyperlinks to named ranges

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
