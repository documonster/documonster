# Word Module — Examples

Self-contained, runnable examples that cover every public Word-module API
plus its conversion bridges (HTML, Markdown, Excel, PDF, Flat OPC, ODT)
plus the `documonster/word/crypto` subpath for decryption. Each file is
hermetic — no external assets are required; all images / fonts are inline
byte literals.

## Run

```bash
# Run a single example
pnpm example --filter 01-basics

# Run every Word example (writes everything into tmp/word-examples/)
pnpm example --filter word/examples
```

`pnpm example` runs the whole tree; `--filter` narrows it by path substring. It
resolves imports the way the test suite does, which a bare `tsx` invocation does
not — see `scripts/run-examples.ts`.

Output is written to `tmp/word-examples/`.

## Index

### 1–14 Authoring building blocks

| File                                | Topic                                                          |
| ----------------------------------- | -------------------------------------------------------------- |
| `01-basics.ts`                      | Document creation, headings, paragraphs, breaks, Unicode/RTL   |
| `02-formatting.ts`                  | All run-level formatting (bold, color, font, effects, …)       |
| `03-tables.ts`                      | Tables: merging, borders, nested, vertical text, floating      |
| `04-layout.ts`                      | Pages, margins, columns, sections, line numbers, RTL           |
| `05-lists.ts`                       | Bullet / numbered / multi-level / custom abstract numbering    |
| `06-headers-footers.ts`             | Default / first / even-odd headers, page-of-N footer, table FT |
| `07-images.ts`                      | Inline / floating / behind / rotated / nested images           |
| `08-hyperlinks-bookmarks.ts`        | URL / mailto / anchor links, bookmarks, REF / PAGEREF          |
| `09-styles-numbering.ts`            | Custom paragraph / character / table styles, asian numerals    |
| `10-toc-fields.ts`                  | Table of contents, every field helper                          |
| `11-footnotes-endnotes-comments.ts` | Notes (rich content) and review comments + 100-footnote stress |
| `12-math.ts`                        | OMML — fractions, integrals, matrices, eq arrays               |
| `13-shapes.ts`                      | DrawingML shapes, gradients, effects, 3D bevel                 |
| `14-charts.ts`                      | Built-in chart types (column, line, pie, scatter, combo)       |

### 15–24 Editing, querying & document features

| File                          | Topic                                                                                          |
| ----------------------------- | ---------------------------------------------------------------------------------------------- |
| `15-track-changes.ts`         | Insertion / deletion / move + accept/reject helpers                                            |
| `16-templates.ts`             | `Template.fill` + JSON / XML / CSV data sources, custom delim, `Template.isChart`              |
| `17-forms-sdt.ts`             | Legacy form fields + SDT (content controls) + data binding                                     |
| `18-read-modify.ts`           | `Document.read`, `Query.getText`, `Query.searchText`, `Document.replaceText` (in 9 containers) |
| `19-merge-split-diff.ts`      | `Transform.merge` / `Transform.split` / `Diff.compare`                                         |
| `20-protection-encryption.ts` | `Security.protect` + `WordCrypto.encrypt`                                                      |
| `21-validation.ts`            | `Validation.validate` (clean / malformed / strict / compat)                                    |
| `22-streaming.ts`             | `Stream.DocumentWriter` for huge documents                                                     |
| `23-fonts.ts`                 | `Font.embed` (subsetting & ODTTF obfuscation) + `Font.obfuscate`                               |
| `24-watermark-bg.ts`          | Text & image watermark, page background, custom theme                                          |

### 25–30 Conversions & edge cases

| File                        | Topic                                                                                   |
| --------------------------- | --------------------------------------------------------------------------------------- |
| `25-html-conversion.ts`     | `WordHtml.toDocxBody` / `WordHtml.toDocx` / `WordHtml.fromDocx`                         |
| `26-markdown-conversion.ts` | `Markdown.toDocx` / `Markdown.fromDocx`                                                 |
| `27-excel-conversion.ts`    | `WordExcel.toDocx` / `WordExcel.extractTables` / chart bridge (`toChartSvg`, …)         |
| `28-pdf-conversion.ts`      | `Pdf.fromDocx` (custom page size, custom chart renderer)                                |
| `29-flat-opc-odt.ts`        | `Document.toFlatOpc` / `Convert.flatOpcToParts` / `Convert.fromOdt` / `Convert.toOdt`   |
| `30-edge-cases.ts`          | Encryption rejection, empty / large / nested-table docs, corrupted ZIP, 1 000 bookmarks |

### 31–47 Advanced internals & low-level surface

| File                             | Topic                                                                                                                        |
| -------------------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `31-walker-mapper.ts`            | `Query.each` / `Query.eachBlock` / `Transform.map` / `Query.list*`                                                           |
| `32-style-resolve-themes.ts`     | `Query.getParagraphStyle` / `Query.getRunStyle` / `Query.getTableStyle` / `Theme.resolveColor` / `Styles.parse`              |
| `33-incremental-edit.ts`         | `Package.edit` / `Package.listParts` / `Package.readPart`                                                                    |
| `34-decrypt-roundtrip.ts`        | `WordCrypto.isEncrypted` / `WordCrypto.decrypt` / `Document.read` with password                                              |
| `35-mathml-conversion.ts`        | `Convert.ommlToMathMl` / `Convert.mathMlToOmml` + `ruby` + `mathPhantom`                                                     |
| `36-text-shaping-hyphenation.ts` | `Font.shapeText` / `Font.detectScript` / `Font.detectDirection` / `Font.hyphenateText`                                       |
| `37-svg-render.ts`               | `Layout.getPageMap` / `Layout.compute` / `Layout.pageToSvg` / `Layout.computedPageToSvg` / `Layout.toSvg`                    |
| `38-glossary-buildingblocks.ts`  | `Glossary.createBlock` / `Glossary.createDocument` / `Glossary.listQuickParts`                                               |
| `39-ole-vba.ts`                  | `Ole.extract` / `Ole.createEmbedding` / `Vba.add` (.docm round-trip)                                                         |
| `40-templates-advanced.ts`       | `Package.compileTemplate` / `Package.patchTemplate` / `Template.bindChartData` / `Template.createCompositeDataSource`        |
| `41-fields-engine.ts`            | `Transform.updateFields` / `Transform.updateTableOfContents` / `tcField` / `indexEntryField` / `indexField` / `noteRefField` |
| `42-conversion-ir.ts`            | `Convert.toSemantic` / `Convert.createConversionContext`                                                                     |
| `43-compat-modes.ts`             | `getCompatibilityMode` / `setCompatibilityMode` for Word 2003-2013+                                                          |
| `44-queries-deep.ts`             | `Query.getBookmark` / `Query.getComment` / `Query.searchByFormat` / `Query.listUsedFormats`                                  |
| `45-digital-signatures.ts`       | `WordCrypto.listSignatures` / `parseSignature` / `isWellFormedSignature` / `hasSignatures`                                   |
| `46-security-policy.ts`          | `Security.DEFAULT_POLICY` / `Security.resolvePolicy` / `RenderContext.create` / `RenderContext.createIds`                    |
| `47-low-level-crypto.ts`         | `WordCrypto.verifyPassword` / `readCfb` / `writeCfb` / `parseEncryptionInfo` / `deriveKey` / `decryptPackage`                |

## Conventions

- Every file is a standalone ES module. They import from `../index` for the
  core API, and from `../html` / `../markdown` / `../excel` / `../crypto` /
  `../../pdf` for the conversion bridges and crypto sub-path.
- Output files use the file name as a prefix (e.g. `07-images.docx`).
- Examples that produce many outputs use a sub-directory
  (e.g. `19-merge-split-diff/`).
- All output paths are relative to the project root: `tmp/word-examples/`.
  This directory is gitignored.
