# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [9.5.4](https://github.com/cjnoname/excelts/compare/v9.5.3...v9.5.4) (2026-05-08)


### Bug Fixes

* **stream:** Update read() behavior to align with Node.js 26+ and maintain backward compatibility ([85314fa](https://github.com/cjnoname/excelts/commit/85314fa624b94f82faeaf66de07b7238bf4e93d6))

## [9.5.3](https://github.com/cjnoname/excelts/compare/v9.5.2...v9.5.3) (2026-05-08)


### Bug Fixes

* **pdf:** Draw borders after overflow erase to prevent white rects from covering border lines ([2b8c997](https://github.com/cjnoname/excelts/commit/2b8c9974429d1f69127d76873c2e59a074470d1a))
* **stream:** Align read() with Node.js 26 behavior (return single chunk) ([5e90746](https://github.com/cjnoname/excelts/commit/5e9074695653307cf4e1f1ce072cf68bfab2aabb))

## [9.5.2](https://github.com/cjnoname/excelts/compare/v9.5.1...v9.5.2) (2026-05-08)


### Bug Fixes

* **pdf:** Resolve rich text bold inheritance and border rendering issues ([#154](https://github.com/cjnoname/excelts/issues/154)) ([50be87e](https://github.com/cjnoname/excelts/commit/50be87e950869a00e6b62cc0b55ece34a416b053))

## [9.5.1](https://github.com/cjnoname/excelts/compare/v9.5.0...v9.5.1) (2026-05-08)


### Bug Fixes

* **pdf:** Correct rich text font inheritance, indexed colors, and overflow borders ([8090686](https://github.com/cjnoname/excelts/commit/8090686e4f345e13bd4182472796faa4718fe536))


### Code Refactoring

* **word:** Decompose document.ts into focused modules ([772cc2b](https://github.com/cjnoname/excelts/commit/772cc2be03b41ac56c776ac91bd909943e406d4f))

## [9.5.0](https://github.com/cjnoname/excelts/compare/v9.4.2...v9.5.0) (2026-05-08)


### Features

* **chart:** Add ChartEx cache population for improved preview rendering ([cc3b389](https://github.com/cjnoname/excelts/commit/cc3b3899ea5a632d78a2f2d0ed1d6f995da5975f))
* **excel:** Add chart, chartEx, sparkline, and chartsheet support ([d13f95b](https://github.com/cjnoname/excelts/commit/d13f95b41146e21a5616612a1cdc239daadc87a5))
* **excel:** Chart compatibility matrix, docs & ChartEx/3D/dataTable/userShapes gaps ([3c31a9e](https://github.com/cjnoname/excelts/commit/3c31a9edb5d409234620d47167854972271fc45f))
* **excel:** Chart integration test suite + removeChart anchor cleanup ([465fa6e](https://github.com/cjnoname/excelts/commit/465fa6ea48ffb7192e764f3b5e2721fda832b958))
* **excel:** Classic-chart schema validator + dLbls child order + pivot anchor ([5f6a37c](https://github.com/cjnoname/excelts/commit/5f6a37c306ba0c2080620b6632619a69b08fb083))
* **excel:** Lazy chart support + test infrastructure improvements ([6b5a24d](https://github.com/cjnoname/excelts/commit/6b5a24d55a7ab883ee8784960750e52ce1aff861))
* **excel:** OOXML validator overhaul + chart compliance fixes + full chart catalogue ([3638020](https://github.com/cjnoname/excelts/commit/363802034aefa3f0fe99cc87a58d5e820e01a61e))
* **word:** Add comprehensive Word (DOCX) processing module ([b3d09e9](https://github.com/cjnoname/excelts/commit/b3d09e9b810dbb825c4c3144b8e54bb3176c8a69))


### Bug Fixes

* **chart:** Correct documentation, rendering, and cross-workbook copy issues ([3039fec](https://github.com/cjnoname/excelts/commit/3039fecede48115fe261dda03788b878a1f64d1f))
* **excel:** Eleven chart bugs surfaced by a review + TZ-independent date serial ([0f240e0](https://github.com/cjnoname/excelts/commit/0f240e0b1960dc86029c9010207d5b06b5234c4d))
* **excel:** External links cause Excel to reject file as corrupt ([eeaa273](https://github.com/cjnoname/excelts/commit/eeaa273bf0d87b7817faeab1da57af260bebae4f))
* **excel:** Replace pseudo-glyph PNG text rendering with real TTF rasterizer ([3adefdb](https://github.com/cjnoname/excelts/commit/3adefdb85bf36ce012b6517cdd0932e4d47c2c2d))
* **package:** Remove Word references from description, keywords, and exports ([d31d27f](https://github.com/cjnoname/excelts/commit/d31d27f1b130212e98dbb64c9bc9ac42e391ab49))
* **pdf:** Overhaul chart rendering in PDF export ([72711e7](https://github.com/cjnoname/excelts/commit/72711e7b7c5b87e912a81568490f74d749db8406))
* **pdf:** Rich text overflow, per-run wrap measurement, and overflow erase ([82faf0f](https://github.com/cjnoname/excelts/commit/82faf0f68ef8553dbd6853268de813dc7bd915ef))
* **security:** Eliminate remaining CodeQL regex and sanitization alerts ([7b2d5e7](https://github.com/cjnoname/excelts/commit/7b2d5e76d412cf4cf4efa5f1805520fee78492f3))
* **security:** Resolve all CodeQL polynomial regex and sanitization alerts ([af2d0c5](https://github.com/cjnoname/excelts/commit/af2d0c5378a38e92b5209e261918e9f4a812a44a))
* **svg:** Replace regex parsing with manual attribute parser to avoid backtracking issues ([ce37fd1](https://github.com/cjnoname/excelts/commit/ce37fd1fdf45e8d3bb9bd125d46d535266cda1f3))
* **test:** Replace platform-dependent PNG hash with determinism check ([709d834](https://github.com/cjnoname/excelts/commit/709d83456e9062bef33812bceebc5b46a790540a))

## [9.4.2](https://github.com/cjnoname/excelts/compare/v9.4.1...v9.4.2) (2026-04-25)


### Bug Fixes

* **pdf:** Unicode text rendering and long text clipping in PDF export ([#150](https://github.com/cjnoname/excelts/issues/150)) ([6b6aea9](https://github.com/cjnoname/excelts/commit/6b6aea94a6d51f2f41155a34f3627329d620e317))

## [9.4.1](https://github.com/cjnoname/excelts/compare/v9.4.0...v9.4.1) (2026-04-21)


### Features

* **issue-templates:** Add "Formula" and "Utils" options to module dropdowns ([71f4c32](https://github.com/cjnoname/excelts/commit/71f4c32cd3d671006612a0b3d733cea7c447eee5))


### Bug Fixes

* **excel:** Support comments subdirectory layout and author round-trip ([#148](https://github.com/cjnoname/excelts/issues/148)) ([480d108](https://github.com/cjnoname/excelts/commit/480d108c9296bafa45148b80a8085ac853478f23))


### Miscellaneous Chores

* Release 9.4.1 ([0cf74c7](https://github.com/cjnoname/excelts/commit/0cf74c7574b2e8e503e326b00814d5ab79357de0))

## [9.4.0](https://github.com/cjnoname/excelts/compare/v9.3.1...v9.4.0) (2026-04-20)


### Features

* **excel:** Add Cell.displayText getter, export format helpers, fix date formatting bugs ([a9f4fb5](https://github.com/cjnoname/excelts/commit/a9f4fb5cd97c79accba6c2a5e65d0132a606ec93))
* **formula:** Comprehensive Excel compliance overhaul + 4 architectural features ([2fcf19d](https://github.com/cjnoname/excelts/commit/2fcf19dcc78a2a7e9313e4ab86a6790a6863fda4))

## [9.3.1](https://github.com/cjnoname/excelts/compare/v9.3.0...v9.3.1) (2026-04-18)


### Code Refactoring

* Remove dead code and tighten let/const usage ([29e0166](https://github.com/cjnoname/excelts/commit/29e016654e5cd61e9266449996c5491bef3fe082))

## [9.3.0](https://github.com/cjnoname/excelts/compare/v9.2.1...v9.3.0) (2026-04-18)


### Features

* **excel:** Add COUP family + BINOM.DIST.RANGE ([a7e3e54](https://github.com/cjnoname/excelts/commit/a7e3e54b19fcf6927cc85f74c04aa56ecfad7892))
* **excel:** Add Excel 365 dynamic array formula support (FILTER/SORT/UNIQUE/XLOOKUP/SEQUENCE) ([17e6c22](https://github.com/cjnoname/excelts/commit/17e6c22e9f3f394610f2705c06f666e025edcaed))
* **excel:** Add external workbook link support ([Book]Sheet!Ref) ([8d0d046](https://github.com/cjnoname/excelts/commit/8d0d0467658e32a609ca72e062fc8b216b0c6cae))
* **excel:** Add formula calculation engine with tokenizer, parser, evaluator, dependency graph, and 220+ Excel functions ([6b6c9a8](https://github.com/cjnoname/excelts/commit/6b6c9a84394becf658d4db3e2e1de17676197aeb))
* **excel:** Add ignoredErrors support, absoluteAnchor images, and reconcilePicture null guard ([da34761](https://github.com/cjnoname/excelts/commit/da3476107efe52dd95bf91db3df8bb1b77ed17e3))
* **excel:** Add matrix + series math functions (MMULT, MDETERM, MINVERSE, MUNIT, SERIESSUM) ([ee6a094](https://github.com/cjnoname/excelts/commit/ee6a09499d4f1f269369e78875ae767d7a702b88))
* **excel:** Add REGEX, VALUETOTEXT, ARRAYTOTEXT, PERCENTRANK, PROB functions ([03913e5](https://github.com/cjnoname/excelts/commit/03913e52d41c9e9d4d38e2634486c34329bb9c01))
* **excel:** Add SQRTPI, ENCODEURL, ACCRINTM, TBILL*, PRICEMAT, YIELDMAT ([96813b8](https://github.com/cjnoname/excelts/commit/96813b81a14ad2c413a8ffad171e38d949c015cf))
* **excel:** Add workbook structure protection and public defaultFont API ([7f15a58](https://github.com/cjnoname/excelts/commit/7f15a583aa4dcb2abda25b0152a561ca9ecc212b))
* **excel:** Cast-free hyperlink + formula+hyperlink cell input ([86340a0](https://github.com/cjnoname/excelts/commit/86340a0586846218447b869b457c3923fe194feb))
* **excel:** Harden formula engine — +5000 tests, 30+ bug fixes, 34 new functions, 3D ref + dynamic-array dependency fixes ([06b0cf4](https://github.com/cjnoname/excelts/commit/06b0cf462e936f4646313467cdf5af089d2a7c0f))


### Bug Fixes

* **excel:** Fix 21 formula engine bugs, add 13 functions, 2 language features ([487079a](https://github.com/cjnoname/excelts/commit/487079a2136f8f84e4e4b0a6e7edfbfe1e727293))
* **excel:** Fix data validation sort, empty border truthy, and table name uniqueness ([59f24cf](https://github.com/cjnoname/excelts/commit/59f24cfa4dd840e40c97fb672b029a798daa224c))
* **excel:** Preserve formula+hyperlink and rich-text on round-trip ([#142](https://github.com/cjnoname/excelts/issues/142)) ([fa350f7](https://github.com/cjnoname/excelts/commit/fa350f7ffc18deacad5088311e033c846b4a1606))
* **excel:** SUBTOTAL/AGGREGATE full Excel semantics, totals-row SUBTOTAL codes, multi-area mask merging ([128c54c](https://github.com/cjnoname/excelts/commit/128c54cf4f4ae84d422ec1b2677b27ce30b62ad4))
* **excel:** Support range operator ':' with dynamic refs, add 50+ functions, fix SUBTOTAL/AGGREGATE semantics ([f912857](https://github.com/cjnoname/excelts/commit/f9128570091d0bd6ab566a528f923c27657e5763))
* Update TypeScript linting rules — add no-misused-spread and no-useless-default-assignment ([480b7c0](https://github.com/cjnoname/excelts/commit/480b7c083ad84e9c7518d9cabd38186f0c6ba5fa))


### Code Refactoring

* **excel:** Replace old formula engine with new compile→evaluate→materialize pipeline ([b747df7](https://github.com/cjnoname/excelts/commit/b747df7f2a78174fdfcf7d83e3e273e46371a6a8))
* **formula:** Promote to top-level module with standalone entry + strict defined-name classification ([b34fe81](https://github.com/cjnoname/excelts/commit/b34fe818f3d7cd4d8a8514cbcba9e800cf20eae7))

## [9.2.1](https://github.com/cjnoname/excelts/compare/v9.2.0...v9.2.1) (2026-04-14)


### Bug Fixes

* Data bar CF ext section in streaming writer, add WorksheetWriter.addRows() ([ea47ced](https://github.com/cjnoname/excelts/commit/ea47ced479679861e03f70f3a11790058e20bc80))
* Resolve CodeQL [#80](https://github.com/cjnoname/excelts/issues/80) [#82](https://github.com/cjnoname/excelts/issues/82) — isolate PDF key derivation from shared crypto ([ef2347c](https://github.com/cjnoname/excelts/commit/ef2347cb6db1656797b1cb1f869b8186c80a9d81))
* Resolve CodeQL security alerts [#79](https://github.com/cjnoname/excelts/issues/79), [#80](https://github.com/cjnoname/excelts/issues/80), [#81](https://github.com/cjnoname/excelts/issues/81) ([b35777c](https://github.com/cjnoname/excelts/commit/b35777c33213f6bf42289c713071300dd5a891fa))
* Resolve CodeQL security alerts [#80](https://github.com/cjnoname/excelts/issues/80) [#82](https://github.com/cjnoname/excelts/issues/82) — exclude PDF-mandated hashes ([b99bd23](https://github.com/cjnoname/excelts/commit/b99bd23deba0d1325036f1ad68234ded79b2de24))
* Use crypto.hash() instead of createHash() to resolve CodeQL alerts ([0125c19](https://github.com/cjnoname/excelts/commit/0125c192e86e49a840875430366d2a72f3677914))

## [9.2.0](https://github.com/cjnoname/excelts/compare/v9.1.0...v9.2.0) (2026-04-14)

### Features

- Add annotation/form creation, SVG path, digital signatures; extract shared crypto ([12f1a7b](https://github.com/cjnoname/excelts/commit/12f1a7bc5a50d2e0bd5fa355c9e6d5e4498ba560))
- Add generateTestCertificate() API; remove raw DER bytes from example ([e09f2c7](https://github.com/cjnoname/excelts/commit/e09f2c7100bb3d1738df6997d52b266cff9b8e95))
- Add high-level sign() API to PdfDocumentBuilder and PdfEditor ([5b61561](https://github.com/cjnoname/excelts/commit/5b615616590220b72042e74b0b665065554c00d4))
- Add PDF builder and editor APIs with unified image/text utilities ([e7e4709](https://github.com/cjnoname/excelts/commit/e7e47091b09414fa076eb009328f914ed44ae1ab))
- Complete PDF builder, editor, and reader roadmap ([7572e0b](https://github.com/cjnoname/excelts/commit/7572e0b3964ecea5eee8f042e24615c4251df608))

### Bug Fixes

- Make signature widget visible so Adobe Acrobat shows signature status ([d4cfcef](https://github.com/cjnoname/excelts/commit/d4cfcef5c911131a5795108d4251abe2f604dd78))
- Move signature widget to invisible rect so visible text is not clickable ([4d2bc96](https://github.com/cjnoname/excelts/commit/4d2bc9690cf073b7c1f89ea4ea2ba20c9bab268a))
- Radio button Annots, signature digest algorithm, editor radio overlay ([1d2d785](https://github.com/cjnoname/excelts/commit/1d2d785979be1a5ba58b68cc48edb2ec32e985b9))
- Resolve PDF border overlap and text-border collision ([#136](https://github.com/cjnoname/excelts/issues/136)) ([6ce6c2c](https://github.com/cjnoname/excelts/commit/6ce6c2cb9012a940f5017fa33cb1c3cff314cbde))
- Signing+forms AcroForm resources, restore drawText, normalize browser hash ([71d2281](https://github.com/cjnoname/excelts/commit/71d2281d13e226105d31c001cd037a95a52a2d9a))
- SVG T command bug, editor state leaks, signPdf pattern matching, orphaned catalogs ([a248e34](https://github.com/cjnoname/excelts/commit/a248e343b0cacad3a37cc914adf9b4ac4d08e147))
- Update cspell words list and add oxlint dependency in package.json ([1a8691f](https://github.com/cjnoname/excelts/commit/1a8691f4e9282997b1e77b37c8664fa900dc88a3))
- Use process.execPath to run oxfmt in generate script for Windows compat ([fe83656](https://github.com/cjnoname/excelts/commit/fe83656e3a65cde20b3b084a8026c4f836613601))

### Code Refactoring

- Replace prettier with oxfmt, migrate typescript-eslint to tsgolint ([8e24fdb](https://github.com/cjnoname/excelts/commit/8e24fdb901863e929065266c83575529f457d5e6))

## [9.1.0](https://github.com/cjnoname/excelts/compare/v9.0.0...v9.1.0) (2026-04-09)

### Features

- Add watermark support for PDF and Excel ([a0d72ec](https://github.com/cjnoname/excelts/commit/a0d72ec81246259b2182270ed234aa78763be53e)), closes [#134](https://github.com/cjnoname/excelts/issues/134)

### Bug Fixes

- Correct rotated text alignment and add slanted borders in PDF rendering ([10b2d8c](https://github.com/cjnoname/excelts/commit/10b2d8c2ffc5ce271a37e9a962cfef2d4686a068))

## [9.0.0](https://github.com/cjnoname/excelts/compare/v8.1.2...v9.0.0) (2026-04-09)

### ⚠ BREAKING CHANGES

- pdf(), readPdf(), and excelToPdf() now return Promise instead of synchronous results. All call sites must use await.

### Features

- Make PDF APIs async to avoid blocking the event loop ([2f521cd](https://github.com/cjnoname/excelts/commit/2f521cd3c7706c191f01e9483ebff659c963c73d))

### Bug Fixes

- Improve Excel-to-PDF style fidelity and fix cell style mutation leaks ([0448604](https://github.com/cjnoname/excelts/commit/044860403f6d5454f85b92fe701e2f7a13fbab73))

## [8.1.2](https://github.com/cjnoname/excelts/compare/v8.1.1...v8.1.2) (2026-04-07)

### Bug Fixes

- Add PDF subpath export and complete tree-shake coverage for all modules ([4faedf6](https://github.com/cjnoname/excelts/commit/4faedf6949471e9412113e33605553a5e4a93189))

## [8.1.1](https://github.com/cjnoname/excelts/compare/v8.1.0...v8.1.1) (2026-04-05)

### Bug Fixes

- **pdf:** Prevent double-unescaping in XML entity decoding (CWE-116) ([4423033](https://github.com/cjnoname/excelts/commit/442303343894cc7686fb43b0b7bed88772ac8d62))

## [8.1.0](https://github.com/cjnoname/excelts/compare/v8.0.0...v8.1.0) (2026-04-05)

### Features

- **pdf:** Add zero-dependency PDF reader with text, image, and metadata extraction ([23663b3](https://github.com/cjnoname/excelts/commit/23663b3337392c5fc098cc4b15ee6243f6879ff2))
- **pdf:** Upgrade writer to PDF 2.0/AES-256, add annotation and form field extraction ([c725147](https://github.com/cjnoname/excelts/commit/c725147efd803ea28d34fe0aaf8210cfbd11ee0f))

### Performance Improvements

- **archive:** Batch small async push() calls in ZipDeflateFile (4.4x browser speedup) ([54a7a78](https://github.com/cjnoname/excelts/commit/54a7a78779e2eb9772dbdf32647e63d81132fc0d)), closes [#127](https://github.com/cjnoname/excelts/issues/127)

## [8.0.0](https://github.com/cjnoname/excelts/compare/v7.6.0...v8.0.0) (2026-04-05)

### ⚠ BREAKING CHANGES

- All Md-prefixed identifiers renamed to Markdown.

### Features

- **md:** Add Markdown table module with GFM parsing, formatting, and Workbook integration ([dfd55b7](https://github.com/cjnoname/excelts/commit/dfd55b77594fd3c1c4e7332aee3b4bbc1090433a))

### Bug Fixes

- **archive:** Use Dynamic Huffman and native CompressionStream for browser compression ([b18414d](https://github.com/cjnoname/excelts/commit/b18414d38d5f2cfaed5daa80aa50899157624aa8)), closes [#125](https://github.com/cjnoname/excelts/issues/125)
- Avoid file system race condition in XML stream writer example ([#77](https://github.com/cjnoname/excelts/issues/77)) ([7a83866](https://github.com/cjnoname/excelts/commit/7a8386694da3e347fce8ec200cce5db860ecae15))

### Code Refactoring

- Rename Md module to Markdown globally ([d0072ae](https://github.com/cjnoname/excelts/commit/d0072aee6f2e0cc2bfc6f8679f0a1125acd0eed6))
- Rename remaining md variables to markdown in examples, tests, and docs ([767bc12](https://github.com/cjnoname/excelts/commit/767bc1292212b48fc0e01664696b54e63904b0e9))

## [7.6.0](https://github.com/cjnoname/excelts/compare/v7.5.0...v7.6.0) (2026-04-04)

### Features

- **xml:** Add invalidCharHandling option for tolerant XML parsing ([f4c5ecc](https://github.com/cjnoname/excelts/commit/f4c5ecc5eb82becbe58da0909f3fc0d9556da5e9))

## [7.5.0](https://github.com/cjnoname/excelts/compare/v7.4.0...v7.5.0) (2026-04-02)

### Features

- **xml:** Add ignoreAttributes option and remove fast-xml-parser references ([3ae7f67](https://github.com/cjnoname/excelts/commit/3ae7f675f0c7749ca49f404471528983eaf96a6a))

## [7.4.0](https://github.com/cjnoname/excelts/compare/v7.3.0...v7.4.0) (2026-04-02)

### Features

- **xml:** Add isArray callback for per-tag array wrapping in toPlainObject/parseXmlToObject ([edd69c3](https://github.com/cjnoname/excelts/commit/edd69c3a5d92810483733692626c64ffdd4070b3))

## [7.3.0](https://github.com/cjnoname/excelts/compare/v7.2.0...v7.3.0) (2026-04-02)

### Features

- **xml:** Add toPlainObject and parseXmlToObject for XML-to-plain-object conversion ([b7275dc](https://github.com/cjnoname/excelts/commit/b7275dc0215b4c2d507adbc1d4efb208213c4cbf))

## [7.2.0](https://github.com/cjnoname/excelts/compare/v7.1.1...v7.2.0) (2026-04-01)

### Features

- Add ./xml subpath export and treeshake verification ([6dfee7a](https://github.com/cjnoname/excelts/commit/6dfee7aaa84b076d275be50db8a4c356fc46c634))

## [7.1.1](https://github.com/cjnoname/excelts/compare/v7.1.0...v7.1.1) (2026-04-01)

### Bug Fixes

- Consolidate module links, remove duplicate sections from READMEs ([4a85b2c](https://github.com/cjnoname/excelts/commit/4a85b2c8a2f1183cac7882a90955cab4d660e859))

## [7.1.0](https://github.com/cjnoname/excelts/compare/v7.0.1...v7.1.0) (2026-04-01)

### Features

- **xml:** Add standalone [@xml](https://github.com/xml) module with streaming write pipeline and backpressure support ([a0a0630](https://github.com/cjnoname/excelts/commit/a0a063031f00cb4c8837b7cb47a37ca8a28c7b8c))

### Bug Fixes

- Enhance backpressure handling in StreamingZipWriterAdapter for sync and async write() return types ([156b5fa](https://github.com/cjnoname/excelts/commit/156b5fa95d6bcae158b2f6020183350a61c2de8d))
- Extract \_checkBackpressure to resolve CodeQL missing-await alerts [#74](https://github.com/cjnoname/excelts/issues/74) [#75](https://github.com/cjnoname/excelts/issues/75) [#76](https://github.com/cjnoname/excelts/issues/76) ([7081e1f](https://github.com/cjnoname/excelts/commit/7081e1f0ec46c3e6a933d0feb9b418e6f8a61010))
- Handle async write() return in StreamingZipWriterAdapter backpressure check ([dc8c3fd](https://github.com/cjnoname/excelts/commit/dc8c3fd317134bf099650e306103152ff3b90a71))
- Wait for file stream close before resolving on Windows ([84503eb](https://github.com/cjnoname/excelts/commit/84503eba664b538eed74729de013c6345a2ed335))
- **xml:** Complete namespace support, harden parser, simplify module, and update validator ([9a59fe8](https://github.com/cjnoname/excelts/commit/9a59fe88e35700d73612f05b1b4feb825e687d73))
- **xml:** Correct HAN CELL prefix detection, xmlEncode FFFE/FFFF, and tag mutation ([56c8665](https://github.com/cjnoname/excelts/commit/56c8665b40f3172b62156cb012c97216ed761b36))
- **xml:** Reject second root element, fatal UTF-8 decode, lightweight decodeCol ([95b676e](https://github.com/cjnoname/excelts/commit/95b676e07dbcb8e74b4a928051bc73a1961b9929))

### Performance Improvements

- Cache isDateFmt, use encodeInto in StringBuf, optimize row spans parsing ([3c63a07](https://github.com/cjnoname/excelts/commit/3c63a0723f4fe7c5eb4e030eeae91b723531f1ec))
- **xml:** Eliminate XML parse overhead with direct SAX callbacks ([74ebdaf](https://github.com/cjnoname/excelts/commit/74ebdaf0c3a6ad154f15161b2c09388567c008d5))
- **xml:** Optimize hot paths for large-data throughput ([97bd074](https://github.com/cjnoname/excelts/commit/97bd07497c7eba6afd5e805043459ce83355801b))
- **xml:** Optimize text decoding, SAX consumers, and write batching ([aaa41a9](https://github.com/cjnoname/excelts/commit/aaa41a956ab47f690c126ace6885f4885a7df85a))

## [7.0.1](https://github.com/cjnoname/excelts/compare/v7.0.0...v7.0.1) (2026-03-29)

### Bug Fixes

- **ci:** Match release-please merge commit message pattern ([4fedfd8](https://github.com/cjnoname/excelts/commit/4fedfd8d1d261c755fbd70d0084bc51709d6ec7e))
- **stream:** Increase async transform error test timeout to prevent flaky failure ([a91642e](https://github.com/cjnoname/excelts/commit/a91642ec8fc3dfdfb6bc65977a71e0dd9882afff))

## [7.0.0](https://github.com/cjnoname/excelts/compare/v6.2.0...v7.0.0) (2026-03-28)

### ⚠ BREAKING CHANGES

- **excel:** Deprecated type alias `ZipOptions` removed — use `WorkbookZipOptions` instead.

### Features

- **pdf:** Add image embedding support for standalone pdf() and excelToPdf() ([9d0f2b2](https://github.com/cjnoname/excelts/commit/9d0f2b2c0c472684b5a40d8817931b7e8d6c11a4))

### Bug Fixes

- **ci:** Resolve merge conflict, keep quoted if expression ([f909a62](https://github.com/cjnoname/excelts/commit/f909a6210703fb7573add9f5a45e14b6e4864134))

### Code Refactoring

- **excel:** Remove deprecated Image and ZipOptions type aliases ([ab499d5](https://github.com/cjnoname/excelts/commit/ab499d5566bff2ee5a530571c849de3a3a9d0168))
- **pdf:** Decouple PDF engine from Excel, add standalone pdf() API ([a426eb4](https://github.com/cjnoname/excelts/commit/a426eb46ce71c594878c6fd4abeee1d5b7cba7ab))

## [6.3.0] (Unreleased)

### Features

- **pdf:** Add standalone `pdf()` API for generating PDFs from plain data without Excel
- **pdf:** Add `PdfImage` support for embedding JPEG/PNG images in standalone `pdf()` API
- **pdf:** Add image embedding examples (`pdf-images.ts`) for both `excelToPdf()` and `pdf()` paths

### Bug Fixes

- **pdf:** Fix header-only empty data sheets producing blank pages
- **pdf:** Fix sparse column headers rendering at wrong positions
- **pdf:** Fix image-only standalone sheets being dropped (empty bounds early return)
- **pdf:** Fix `excelToPdf()` dropping images anchored outside cell data bounds
- **pdf:** Fix `excelToPdf()` not extending bounds to image `br` (bottom-right) anchor

### Code Refactoring

- **pdf:** Decouple PDF engine from Excel module via `excel-bridge.ts`
- **pdf:** Remove `numFmt` from standalone `PdfCell` (Excel-specific concept)

### BREAKING CHANGES

- **pdf:** `PdfExporter` class has been removed. Use `excelToPdf(workbook, options?)` instead of `new PdfExporter(workbook).export(options)`.
- **pdf:** `exportPdf()` is no longer a public export. Use `excelToPdf()` for Excel workbooks or `pdf()` for standalone PDF generation.
- **pdf:** `ExcelPdfExporter` class has been removed. Use `excelToPdf()` instead.
- **excel:** Deprecated type alias `Image` has been removed. Use `ImageData` instead.
- **excel:** Deprecated type alias `ZipOptions` has been removed. Use `WorkbookZipOptions` instead.

## [6.2.0](https://github.com/cjnoname/excelts/compare/v6.1.3...v6.2.0) (2026-03-28)

### Features

- **excel:** Add autoFitColumns() and autoFitRows() for auto-sizing ([c209bdf](https://github.com/cjnoname/excelts/commit/c209bdfff0a7a6e947368fe9e2adaabc2687eec8)), closes [#114](https://github.com/cjnoname/excelts/issues/114)
- **pdf:** Add zero-dependency Excel-to-PDF export module ([fdc568b](https://github.com/cjnoname/excelts/commit/fdc568be0a2d14d5d3df5a7d00d2124c5273c421))

### Bug Fixes

- **security:** Harden input parsing against DoS, prototype pollution, and crash vectors ([2eca761](https://github.com/cjnoname/excelts/commit/2eca761765f3bc5a42f56e6b79852967d50394a5))
- **treeshake:** Improve tree-shaking for rspack/webpack and add verification script ([012493c](https://github.com/cjnoname/excelts/commit/012493c1ebb95550ff30c02e5bea70b38a24ecb9))
- **worksheet:** Remove unnecessary null check for worksheet name length ([b804e9c](https://github.com/cjnoname/excelts/commit/b804e9ce2dfb59d2c4b28a1c9b66564fa13993d8))

### Code Refactoring

- **excel:** Extract shared worksheet utils and fix TS6 build compatibility ([2ac2885](https://github.com/cjnoname/excelts/commit/2ac2885a1452efef3233adcaa479fe5b1e395ff2))

## [6.1.3](https://github.com/cjnoname/excelts/compare/v6.1.2...v6.1.3) (2026-03-24)

### Bug Fixes

- **excel:** Change WorkbookWriter default compression level from 1 to 6 ([d30313a](https://github.com/cjnoname/excelts/commit/d30313a05ab401072df094dbd02881bb96330c1e))

## [6.1.2](https://github.com/cjnoname/excelts/compare/v6.1.1...v6.1.2) (2026-03-24)

### Bug Fixes

- **archive:** Drastically improve compression ratio for streaming writes ([78f9f22](https://github.com/cjnoname/excelts/commit/78f9f22cbe75be144b7fae53a66c5c706ae35762))
- **build:** Prevent race condition on dist/types during parallel builds ([d0db0dc](https://github.com/cjnoname/excelts/commit/d0db0dc36dd950df8e7281ca0200c5387837aadb))

## [6.1.1](https://github.com/cjnoname/excelts/compare/v6.1.0...v6.1.1) (2026-03-23)

### Bug Fixes

- **excel:** Case-insensitive worksheet name lookup and correct internal hyperlink OOXML output ([2e5f0dc](https://github.com/cjnoname/excelts/commit/2e5f0dc1641e7aee3af7ae916432d2bb202cd58a))
- **excel:** Make worksheet name lookup case-insensitive to match Excel semantics ([f735884](https://github.com/cjnoname/excelts/commit/f73588411a102b913d9fdc971124b835b17a69ea))

## [6.1.0](https://github.com/cjnoname/excelts/compare/v6.0.0...v6.1.0) (2026-03-16)

### Features

- **stream:** Add WorksheetWriter.addImage support ([#108](https://github.com/cjnoname/excelts/issues/108)) ([a91d9e1](https://github.com/cjnoname/excelts/commit/a91d9e11b8304037658b9bbfc169bde497fd2521))

### Bug Fixes

- **ci:** Remove stale release-as pinning to unblock version bumps ([1da5585](https://github.com/cjnoname/excelts/commit/1da558575de7a1e79de16f48e07e6bf1b6bf6961))
- **test:** Pin TAR modTime in byte-for-byte consistency test ([c163e49](https://github.com/cjnoname/excelts/commit/c163e49568ef7ef58c705dc4ce35b88eb59e86e8))

## [6.0.0](https://github.com/cjnoname/excelts/compare/v6.0.0-beta.10...v6.0.0) (2026-03-16)

### Bug Fixes

- Promote to 6.0.0 stable release ([083e0e0](https://github.com/cjnoname/excelts/commit/083e0e0be35f49ff866223b569e30d8ca288f115))

## [6.0.0-beta.10](https://github.com/cjnoname/excelts/compare/v6.0.0-beta.9...v6.0.0-beta.10) (2026-03-16)

### Bug Fixes

- Widen RowValues object type from Record&lt;string, unknown&gt; to Record&lt;string, any&gt; ([9d29be6](https://github.com/cjnoname/excelts/commit/9d29be6dceecc92f76176316e01212a82d568399))

## [6.0.0-beta.9](https://github.com/cjnoname/excelts/compare/v6.0.0-beta.8...v6.0.0-beta.9) (2026-03-16)

### Bug Fixes

- **ci:** Avoid npm preversion hook in canary version bump ([3bcf30f](https://github.com/cjnoname/excelts/commit/3bcf30f7eb1eed7c8e8e63331d689c2101aa2113))
- Consume data descriptor after known-size pump in streaming parser ([f7be681](https://github.com/cjnoname/excelts/commit/f7be68135e55fc62d82240166fe18f53673d86dd))

## [6.0.0-beta.8](https://github.com/cjnoname/excelts/compare/v6.0.0-beta.7...v6.0.0-beta.8) (2026-03-15)

### Bug Fixes

- Correct dishonest type tightenings and remove unsafe toJSON generic ([d974843](https://github.com/cjnoname/excelts/commit/d9748433666b2b0870098fe2900901aa8cde7245))
- Restore package.json and manifest to current released version ([5111e60](https://github.com/cjnoname/excelts/commit/5111e60d4fd1d7cee301b61e11190c67c66ece00))

## [6.0.0-beta.7](https://github.com/cjnoname/excelts/compare/v6.0.0-beta.6...v6.0.0-beta.7) (2026-03-15)

### Features

- Integrate sheet-utils into native Worksheet/Workbook API ([34148b1](https://github.com/cjnoname/excelts/commit/34148b1d85d21d2a1d08f428669c6fe3842c2c1a))

### Bug Fixes

- Chai assertion syntax, anchor copy-paste bug, duplicate test, and 8 weak assertions ([4919a36](https://github.com/cjnoname/excelts/commit/4919a3613317709cfafbcacdce184aafec80008c))
- Unify model field naming, strengthen types, and rewrite importSheet as deep copy ([cb381c7](https://github.com/cjnoname/excelts/commit/cb381c7acc54d341102f90b1e12b97638a705e69))
- WorkbookReader emits wrong worksheet name when workbook.xml is parsed after worksheets ([206e424](https://github.com/cjnoname/excelts/commit/206e4246201e74998038b07c7ede327ad6596956))
- WorksheetWriter.findCell used wrong property name (address.column → address.col) ([72ee159](https://github.com/cjnoname/excelts/commit/72ee159541536be85c7f893de3614b1b79f35cb1))

### Code Refactoring

- Convert all .then() chains to async/await in integration tests ([6ba5dcc](https://github.com/cjnoname/excelts/commit/6ba5dcc11f87f6a4be52a21ba92a0a4781ec8dcc))
- Modernize excel module types and patterns ([07e3b89](https://github.com/cjnoname/excelts/commit/07e3b890e107b056bd8f521619b0ceed6b084b38))

## [6.0.0-beta.6](https://github.com/cjnoname/excelts/compare/v6.0.0-beta.5...v6.0.0-beta.6) (2026-03-13)

### Bug Fixes

- Revert bad release and fix release-please tag format ([79bb0be](https://github.com/cjnoname/excelts/commit/79bb0be3b2792d4e87efbd43f623c466113d724f))

## [6.0.0-beta.4](https://github.com/cjnoname/excelts/compare/v6.0.0-beta.3...v6.0.0-beta.4) (2026-03-13)

### Features

- **archive:** export binary/encoding utilities for standalone usage ([9d31a75](https://github.com/cjnoname/excelts/commit/9d31a75f9901438f526abf5e62b7166081ea1016))

### Bug Fixes

- resolve all 19 CodeQL security alerts ([8e7a79a](https://github.com/cjnoname/excelts/commit/8e7a79a96e220775ecb2e4b7497fc9d761b03be2))
- resolve last 4 CodeQL alerts with inline barriers ([fa727e4](https://github.com/cjnoname/excelts/commit/fa727e4aa68d4346d42bb420a7f45a2fb35445af))
- resolve remaining 13 CodeQL security alerts ([91e6628](https://github.com/cjnoname/excelts/commit/91e6628b6f2286abcc3144cd56768705842cb2b7))
- resolve Windows ENOENT when running rolldown via execFileSync ([d804282](https://github.com/cjnoname/excelts/commit/d804282d60e95521a84716366e5219218b690465))
- support encrypted entries in streaming ZIP parse mode ([32a6c33](https://github.com/cjnoname/excelts/commit/32a6c330b53cf175a0f7ea8ee66e0548e153b89b))

## [6.0.0-beta.3](https://github.com/cjnoname/excelts/compare/v6.0.0-beta.2...v6.0.0-beta.3) (2026-03-12)

### Bug Fixes

- change Worksheet.columns return type from Column[] | null to Column[] ([ab3f3fe](https://github.com/cjnoname/excelts/commit/ab3f3fef022d8f6d081fdc064bf3a1a38a0ef121))

## [6.0.0-beta.2](https://github.com/cjnoname/excelts/compare/v6.0.0-beta.1...v6.0.0-beta.2) (2026-03-12)

### Bug Fixes

- decode OOXML _xHHHH_ escapes in table column name attributes ([#94](https://github.com/cjnoname/excelts/issues/94)) ([bbfe148](https://github.com/cjnoname/excelts/commit/bbfe1484799d21ed477cdaad3d7d23e4a1404e50))
- stabilize flaky ZipCrypto checkPassword test ([b54eb15](https://github.com/cjnoname/excelts/commit/b54eb1544ce9f2e6e8f31c4006306a139e1f0c1d))

## [6.0.0-beta.1](https://github.com/cjnoname/excelts/compare/v5.1.18...v6.0.0-beta.1) (2026-03-12)

### ⚠ BREAKING CHANGES

- Module structure and entry points have been reorganized. The archive, CSV, and stream submodules are now first-class exports. See MIGRATION.md for details.

### Features

- excelts v6 — cross-platform streaming, archive, and CSV ([28d4f5a](https://github.com/cjnoname/excelts/commit/28d4f5ab129f57977d3d9fe6b0bfa90e6dcce560))
- expose isEncrypted on UnzipEntry in streaming mode ([bd03cf5](https://github.com/cjnoname/excelts/commit/bd03cf56b48e121629e69b8ee8aeff83f2dfe1ae))

### Bug Fixes

- decode OOXML _xHHHH_ escapes with lowercase hex digits ([#94](https://github.com/cjnoname/excelts/issues/94)) ([9c3163f](https://github.com/cjnoname/excelts/commit/9c3163fef636f6b600f133fdb1a9f94aa35617cc))
- merge main fixes, resolve lint errors, and improve code quality ([8a33545](https://github.com/cjnoname/excelts/commit/8a335459b5acde8e4a8a25b3fcdfb2f382f3228e))
- rename zip export to archive in package.json ([f5b3efb](https://github.com/cjnoname/excelts/commit/f5b3efbbe9cc57f2999fe9aff5eab8db6f1d5359))
- resolve CI failures on Node.js 22 and Windows ([83bd891](https://github.com/cjnoname/excelts/commit/83bd891d4e85d77adfd9bed9f6b2efd222956b38))
- **stream:** align browser stream behavior with Node.js parity ([34f69c4](https://github.com/cjnoname/excelts/commit/34f69c4c4f9a4284efa504fbf491a79d1a45e462))
- **stream:** constant-memory streaming for ZIP and Excel writers ([#88](https://github.com/cjnoname/excelts/issues/88)) ([532d7bb](https://github.com/cjnoname/excelts/commit/532d7bb7261893b2d13c54ced27e8db7c85c8a37))

## [5.1.18](https://github.com/cjnoname/excelts/compare/v5.1.17...v5.1.18) (2026-03-10)

### Bug Fixes

- sanitize table names to comply with OOXML defined name rules ([#91](https://github.com/cjnoname/excelts/issues/91)) ([b6f9b0e](https://github.com/cjnoname/excelts/commit/b6f9b0e7dd46872b066e90c6ece705931f082ffe))

## [5.1.17](https://github.com/cjnoname/excelts/compare/v5.1.16...v5.1.17) (2026-03-10)

### Bug Fixes

- prevent unbounded memory growth in StreamBuf when data listeners are attached ([090b2e4](https://github.com/cjnoname/excelts/commit/090b2e42b659f251f2409cbf7691939f0fb7e9a0))

## [5.1.16](https://github.com/cjnoname/excelts/compare/v5.1.15...v5.1.16) (2026-03-09)

### Bug Fixes

- handle empty defined name ranges and missing colon in print area/titles ([74ce4e6](https://github.com/cjnoname/excelts/commit/74ce4e6a7bc711c334580bc9cb603800c9b07888))

## [5.1.15](https://github.com/cjnoname/excelts/compare/v5.1.14...v5.1.15) (2026-03-08)

### Bug Fixes

- write ht="1" for height=0 rows to reliably trigger Excel auto-height ([69728a6](https://github.com/cjnoname/excelts/commit/69728a6ad6da6cb7bcf7c128dc2a7a75998f9193))

## [5.1.14](https://github.com/cjnoname/excelts/compare/v5.1.13...v5.1.14) (2026-03-07)

### Bug Fixes

- row height=0 ignored due to falsy-zero checks, add customHeight support ([9c91fdc](https://github.com/cjnoname/excelts/commit/9c91fdc8e3d2ac4a8dc44ea654a8c8cd2767e0e8)), closes [#82](https://github.com/cjnoname/excelts/issues/82)

## [5.1.13](https://github.com/cjnoname/excelts/compare/v5.1.12...v5.1.13) (2026-03-07)

### Bug Fixes

- add post-publish verification step to CI workflows ([c93c9c4](https://github.com/cjnoname/excelts/commit/c93c9c4b10600821fd6689533e89fbd06b005f3e))

## [5.1.12](https://github.com/cjnoname/excelts/compare/v5.1.11...v5.1.12) (2026-03-06)

### Bug Fixes

- isDateFmt now correctly recognizes date formats with text fallback sections ([#79](https://github.com/cjnoname/excelts/issues/79)) ([2d5f238](https://github.com/cjnoname/excelts/commit/2d5f2389a8d7d145b34a92dc6c7ef7231be7beab))
- mergeCells now preserves perimeter borders like Excel ([d9d28d6](https://github.com/cjnoname/excelts/commit/d9d28d66500b05665fab2399277ba2536b2a7d65))

## [5.1.11](https://github.com/cjnoname/excelts/compare/v5.1.10...v5.1.11) (2026-03-06)

### Bug Fixes

- resolve TypeError when loading workbook with table column child elements ([f4bcbe6](https://github.com/cjnoname/excelts/commit/f4bcbe63921060984762fb48f2aa07f5446962a7)), closes [#76](https://github.com/cjnoname/excelts/issues/76)

## [5.1.10](https://github.com/cjnoname/excelts/compare/v5.1.9...v5.1.10) (2026-03-05)

### Bug Fixes

- empty style object shadowing in \_mergeStyle and shared style references in row/cell operations ([7df419d](https://github.com/cjnoname/excelts/commit/7df419daeac85c743c4ae50c885025c7b69bcee6))

## [5.1.9](https://github.com/cjnoname/excelts/compare/v5.1.8...v5.1.9) (2026-03-02)

### Bug Fixes

- improve normalizeWritable function to handle Web WritableStream correctly ([f5cf6f5](https://github.com/cjnoname/excelts/commit/f5cf6f5fd5167004d90d7f12ae9bfb0ddb053c08))

## [5.1.8](https://github.com/cjnoname/excelts/compare/v5.1.7...v5.1.8) (2026-02-28)

### Bug Fixes

- improve public API return types and enum types ([d261785](https://github.com/cjnoname/excelts/commit/d261785d4ade30ad5db953a17286e309a0753193))

## [5.1.7](https://github.com/cjnoname/excelts/compare/v5.1.6...v5.1.7) (2026-02-28)

### Bug Fixes

- add main, module, and types fields to package.json for legacy moduleResolution compatibility ([3f67511](https://github.com/cjnoname/excelts/commit/3f67511325b183a0752debaebca223b6cbe99d67)), closes [#69](https://github.com/cjnoname/excelts/issues/69)
- hide internal underscore-prefixed members from public type declarations ([f94d157](https://github.com/cjnoname/excelts/commit/f94d1579a2c964e4b4d42269e560d3642f1e06cb)), closes [#68](https://github.com/cjnoname/excelts/issues/68)

## [5.1.6](https://github.com/cjnoname/excelts/compare/v5.1.5...v5.1.6) (2026-02-28)

### Bug Fixes

- revert unnecessary optional chaining for date1904 property access ([042f6c3](https://github.com/cjnoname/excelts/commit/042f6c35981208e284c4aeb60f6f2631cb32362b))

## [5.1.5](https://github.com/cjnoname/excelts/compare/v5.1.4...v5.1.5) (2026-02-27)

### Bug Fixes

- use optional chaining for date1904 property access in XLSX class ([4e74f80](https://github.com/cjnoname/excelts/commit/4e74f805b1817a0a917c5fd73d7faf34f791810d))

## [5.1.4](https://github.com/cjnoname/excelts/compare/v5.1.3...v5.1.4) (2026-02-27)

### Bug Fixes

- update worksheet fileIndex handling for consistency in ZIP entry paths ([5cda867](https://github.com/cjnoname/excelts/commit/5cda86708efb9007e86435157687a705019fb046))

## [5.1.3](https://github.com/cjnoname/excelts/compare/v5.1.2...v5.1.3) (2026-02-11)

### Bug Fixes

- clone images when duplicating rows ([#57](https://github.com/cjnoname/excelts/issues/57)) ([bd7d949](https://github.com/cjnoname/excelts/commit/bd7d949694641f2968a94424f14123827c5634c3))
- prevent image duplication on read-write round-trips ([#58](https://github.com/cjnoname/excelts/issues/58)) ([3da3461](https://github.com/cjnoname/excelts/commit/3da3461c68bd751e2b43527e458e4ee847b5a9b4))

## [5.1.2](https://github.com/cjnoname/excelts/compare/v5.1.1...v5.1.2) (2026-02-10)

### Bug Fixes

- preserve merged cell styles when splicing rows/columns ([#55](https://github.com/cjnoname/excelts/issues/55)) ([668bec7](https://github.com/cjnoname/excelts/commit/668bec7956818fd6d30b57993d6c74b9b3f213a6))

## [5.1.1](https://github.com/cjnoname/excelts/compare/v5.1.0...v5.1.1) (2026-02-09)

### Bug Fixes

- preserve merge information when splicing rows/columns and duplicating rows ([#53](https://github.com/cjnoname/excelts/issues/53)) ([62bbc16](https://github.com/cjnoname/excelts/commit/62bbc160cfb76ed6686b4c81620feb4c3fc5c143))

## [5.1.0](https://github.com/cjnoname/excelts/compare/v5.0.6...v5.1.0) (2026-02-08)

### Features

- complete pivot table implementation with roundtrip support and codebase refactoring ([2801053](https://github.com/cjnoname/excelts/commit/2801053450c369bfaeb6a14701cb08a01ed156a7))

## [5.0.6](https://github.com/cjnoname/excelts/compare/v5.0.5...v5.0.6) (2026-02-06)

### Bug Fixes

- update image anchor positions when rows or columns are spliced ([#50](https://github.com/cjnoname/excelts/issues/50)) ([c164bec](https://github.com/cjnoname/excelts/commit/c164becdf233e1b96e9cff7ece2e8e2e9dc45990))

## [5.0.5](https://github.com/cjnoname/excelts/compare/v5.0.4...v5.0.5) (2026-02-06)

### Bug Fixes

- handle styleId=0 correctly in reconcile functions ([01a532b](https://github.com/cjnoname/excelts/commit/01a532bd0185c6b1381e5dc2c7caf9e875fe8e40))
- handle styleId=0 correctly in reconcile functions ([50e097a](https://github.com/cjnoname/excelts/commit/50e097a4edf977331ecb30ee7b86ae000016b755))
- **streaming:** add null guard in WorksheetWriter.eachRow() for sparse rows ([b0f3079](https://github.com/cjnoname/excelts/commit/b0f30795045794de1231ef53bc761232c50201f2))

## [5.0.4](https://github.com/cjnoname/excelts/compare/v5.0.3...v5.0.4) (2026-01-25)

### Bug Fixes

- **pivot-table:** preserve worksheetSource name attribute for table references ([#45](https://github.com/cjnoname/excelts/issues/45)) ([ef1722b](https://github.com/cjnoname/excelts/commit/ef1722b69c3b84a22e4083024958797f9c1b5b6a))

## [5.0.3](https://github.com/cjnoname/excelts/compare/v5.0.2...v5.0.3) (2026-01-25)

### Bug Fixes

- **excel:** add default cfvo and color for dataBar conditional formatting ([d7abd28](https://github.com/cjnoname/excelts/commit/d7abd28db0cc8acd987324c6a3057d39d6a2ce17))
- remove redundant no-op string replacement in workbook roundtrip test ([09ee8dc](https://github.com/cjnoname/excelts/commit/09ee8dc4ea0336be72f0e1c4236548c88e305ddf))

## [5.0.2](https://github.com/cjnoname/excelts/compare/v5.0.1...v5.0.2) (2026-01-23)

### Bug Fixes

- **browser:** fix drawing parsing failure in loadFromFiles path ([98c7ee0](https://github.com/cjnoname/excelts/commit/98c7ee0a91caf82a5f43bce8cdd97970152ef2ec))

## [5.0.1](https://github.com/cjnoname/excelts/compare/v5.0.0...v5.0.1) (2026-01-23)

### Bug Fixes

- **xlsx:** preserve metadata attributes during round-trip ([#41](https://github.com/cjnoname/excelts/issues/41)) ([5f5d54d](https://github.com/cjnoname/excelts/commit/5f5d54d1825ce5e64bd69ded42a2511c8070f8bf))

## [5.0.0](https://github.com/cjnoname/excelts/compare/v4.2.3...v5.0.0) (2026-01-22)

### ⚠ BREAKING CHANGES

- **excel:** None - passthrough is opt-in via workbook options

### Features

- **excel:** add chart and drawing passthrough preservation ([a4ea35e](https://github.com/cjnoname/excelts/commit/a4ea35e2ea4297d7e7a06f9dbc83ffd37fd90de4))

### Bug Fixes

- **stream:** remove unreachable streamError rethrow ([b80904d](https://github.com/cjnoname/excelts/commit/b80904dee784b6dc0c1ed92f846194bccec0dcc9))

### Performance Improvements

- move HAN CELL namespace handling from SAX parser to BaseXform ([cc11b20](https://github.com/cjnoname/excelts/commit/cc11b206bec2e0eabf5bd2164743bca0cc43fc97))

## [4.2.3](https://github.com/cjnoname/excelts/compare/v4.2.2...v4.2.3) (2026-01-15)

### Bug Fixes

- support HAN CELL xlsx files with namespace prefixes ([88820eb](https://github.com/cjnoname/excelts/commit/88820eb94192c2b9a10c7794cf698aaa66254387))

## [4.2.2](https://github.com/cjnoname/excelts/compare/v4.2.1...v4.2.2) (2026-01-12)

### Bug Fixes

- **archive:** stabilize streaming unzip and browser parsing ([a503090](https://github.com/cjnoname/excelts/commit/a50309085bceb6986a07d098c57749b4c1476f5a))
- **excel:** make legacy form controls OOXML-valid ([fe7a444](https://github.com/cjnoname/excelts/commit/fe7a444a3586977089ee6b8ad9b24f13d8830152))

### Performance Improvements

- **archive:** speed up streaming unzip hot path ([f808a37](https://github.com/cjnoname/excelts/commit/f808a37255750a26e2db10a95de98461f52b8241))

## [4.2.1](https://github.com/cjnoname/excelts/compare/v4.2.0...v4.2.1) (2026-01-10)

### Bug Fixes

- **archive:** keep ZIP parse streaming for large entries ([c88c61c](https://github.com/cjnoname/excelts/commit/c88c61cc3b3e22b693147303be1e500cd4402a6a))
- **build:** copy LICENSE and THIRD_PARTY_NOTICES to dist/iife ([0919d4d](https://github.com/cjnoname/excelts/commit/0919d4d6313f4b54dd3dfb20d450be287b71830a))
- **excel:** improve legacy form checkbox anchors and controls ([7805a16](https://github.com/cjnoname/excelts/commit/7805a16a85f81eaccf542c6ad093beb5e7d1e73d))

## [4.2.0](https://github.com/cjnoname/excelts/compare/v4.1.0...v4.2.0) (2026-01-10)

### Features

- **excel:** add legacy Form Control Checkbox support ([e7d8c4e](https://github.com/cjnoname/excelts/commit/e7d8c4e4b650aba90d83bb9a2a7d6934945e8a7e))

## [4.1.0](https://github.com/cjnoname/excelts/compare/v4.0.4...v4.1.0) (2026-01-10)

### Features

- **excel:** add Office Online-compatible in-cell checkboxes ([8ac37ef](https://github.com/cjnoname/excelts/commit/8ac37efb46a5f33e85462ef53bf8c6a6cc38025d))

### Bug Fixes

- **excel:** hydrate loaded table rows for mutations ([4f97ebb](https://github.com/cjnoname/excelts/commit/4f97ebb00671c157fac89cdb789fef8727b6deaa))

## [4.0.4](https://github.com/cjnoname/excelts/compare/v4.0.3...v4.0.4) (2026-01-06)

### Bug Fixes

- make generated types NodeNext-safe ([b618378](https://github.com/cjnoname/excelts/commit/b618378a19871d2175a452cc658fbd8859d50704))

## [4.0.3](https://github.com/cjnoname/excelts/compare/v4.0.2...v4.0.3) (2026-01-04)

### Bug Fixes

- **excel:** keep table formulas readable ([3972145](https://github.com/cjnoname/excelts/commit/3972145fe1eec92a3d0895583e6dc91eb9aea9fe)), closes [#29](https://github.com/cjnoname/excelts/issues/29)

## [4.0.2](https://github.com/cjnoname/excelts/compare/v4.0.1...v4.0.2) (2026-01-04)

### Bug Fixes

- **excel:** make table structured refs work ([302e682](https://github.com/cjnoname/excelts/commit/302e6827bb0a286bddaeeaf7abf563ada77cda08)), closes [#26](https://github.com/cjnoname/excelts/issues/26)

## [4.0.1](https://github.com/cjnoname/excelts/compare/v4.0.0...v4.0.1) (2026-01-03)

### Bug Fixes

- **types:** avoid .d.ts specifiers in declarations ([0e5d37f](https://github.com/cjnoname/excelts/commit/0e5d37f63b650ad15f02d13c9548899023152c95))

## [4.0.0](https://github.com/cjnoname/excelts/compare/v3.1.0...v4.0.0) (2026-01-02)

### ⚠ BREAKING CHANGES

- The main package entrypoints no longer re-export the internal stream utility surface. If you were importing stream helpers from the root package, migrate to standard Web Streams (ReadableStream/WritableStream) or pin to an older version.
- **exports:** Use top-level WorkbookWriter/WorkbookReader/WorksheetWriter/WorksheetReader exports instead.

### Features

- **csv:** add valueMapperOptions for decimalSeparator ([b93d66e](https://github.com/cjnoname/excelts/commit/b93d66e5488e4c9833c913c901bb78f9bfb8a1cf))
- **exports:** unify node and browser entrypoints ([c8bc979](https://github.com/cjnoname/excelts/commit/c8bc979725b97b33eac7c8433fe9a60c593483f3))
- remove stream utility re-exports ([ea16582](https://github.com/cjnoname/excelts/commit/ea16582e8434d845ced099ffca80a63c970d3da2))
- **streaming:** browser streaming support ([381817c](https://github.com/cjnoname/excelts/commit/381817ce46b7e367542d251a4da06e98fa747810))
- **streaming:** support Web Streams across environments ([204ba36](https://github.com/cjnoname/excelts/commit/204ba365f4100e05d5fe668ab71f3550a789f94a))
- **xlsx:** allow deterministic zip entry timestamps ([d17da6a](https://github.com/cjnoname/excelts/commit/d17da6a8eea6db0b2c0fff208fcc11c13d71723f))
- **xlsx:** store data validations as ranges ([09c2a40](https://github.com/cjnoname/excelts/commit/09c2a4062c2a4daf1b703dfca195fff0f8dc1987))

### Bug Fixes

- **build:** rewrite tsconfig path aliases in dist outputs ([6791d4e](https://github.com/cjnoname/excelts/commit/6791d4ea91f0296f70a9914d9817fc96cc3e6f53))
- post-merge csv parsing + pivot test import ([8f31be3](https://github.com/cjnoname/excelts/commit/8f31be3b6afd806363d419f3f88365229023a11c))
- **security:** address CodeQL findings ([77dafd9](https://github.com/cjnoname/excelts/commit/77dafd9c012bf6ca45f6ec245d320c7edc1ab7ed))
- **stream:** avoid extra args in browser transform ([6ddacdd](https://github.com/cjnoname/excelts/commit/6ddacddf719aaf784100179dabac26fdd29432bb))
- **stream:** handle browser transform/flush arity safely ([d80b29d](https://github.com/cjnoname/excelts/commit/d80b29de30d667125b51684cb3d2d0aa23bfe634))
- **stream:** make transform arity dispatch CodeQL-friendly ([4ffdd8b](https://github.com/cjnoname/excelts/commit/4ffdd8b66847e809a31ed0d43f69051793083bb2))
- **stream:** use direct call with known signature to satisfy CodeQL ([ea53170](https://github.com/cjnoname/excelts/commit/ea531702c0450c6116c26538da17e48f16c2f6a3))
- **stream:** use proper type assertion for userFlush call ([51d781c](https://github.com/cjnoname/excelts/commit/51d781ca286c7747789bb066071112dd5012bb80))
- **test:** align csv mapper typing ([ab0d509](https://github.com/cjnoname/excelts/commit/ab0d509c1c1d8b8867a6a4ec1be9f74686f03888))

### Performance Improvements

- **csv:** optimize streaming parser hot path ([50bf1ef](https://github.com/cjnoname/excelts/commit/50bf1ef27b59eb8994934262c5deb2217b3346af))

## [3.1.0](https://github.com/cjnoname/excelts/compare/v3.0.1...v3.1.0) (2025-12-30)

### Features

- **csv:** support decimalSeparator option ([418eccf](https://github.com/cjnoname/excelts/commit/418eccf56d8ce94127b89a913f6194743d7157d1)), closes [#20](https://github.com/cjnoname/excelts/issues/20)
- **row:** add getValues and valuesToString helpers ([9dca08f](https://github.com/cjnoname/excelts/commit/9dca08f1ad30144121719ad65b4eb622eef66226)), closes [#19](https://github.com/cjnoname/excelts/issues/19)

## [3.0.1](https://github.com/cjnoname/excelts/compare/v3.0.0...v3.0.1) (2025-12-28)

### Bug Fixes

- PivotTable not forming correctly when rows and values fields are equal ([#15](https://github.com/cjnoname/excelts/issues/15)) ([d3eb98d](https://github.com/cjnoname/excelts/commit/d3eb98d2e04a54d05ec3895ec6f3ee49d90a520a))

## [3.0.0](https://github.com/cjnoname/excelts/compare/v2.0.1...v3.0.0) (2025-12-28)

### ⚠ BREAKING CHANGES

- dyDescent is no longer output by default for new worksheets

### Bug Fixes

- correct Table headerRowCount parsing per ECMA-376 ([6cc6016](https://github.com/cjnoname/excelts/commit/6cc60169b2cd6934b28bbff844195a184990af08))
- improve XML output to match Excel's minimal format ([379d895](https://github.com/cjnoname/excelts/commit/379d895a52b33cfa1c3815953d59d4170e3ca7ec))
- make dyDescent optional per ECMA-376 minimum output principle ([76f9c2b](https://github.com/cjnoname/excelts/commit/76f9c2b7b87d659f254836b8c92c2bfc071be3d6))
- resolve PivotTable XML generation bugs (Issue [#5](https://github.com/cjnoname/excelts/issues/5)) ([d564470](https://github.com/cjnoname/excelts/commit/d564470c3989405a5d4783c669727723cfe020e2))

## [2.0.1](https://github.com/cjnoname/excelts/compare/v2.0.0...v2.0.1) (2025-12-28)

### Bug Fixes

- correct PivotTable XML generation for rowItems, colItems, and recordCount ([2e956a6](https://github.com/cjnoname/excelts/commit/2e956a6e39be34db854f1d7d56cfca2646b98dc6))

## [2.0.0](https://github.com/cjnoname/excelts/compare/v1.6.3...v2.0.0) (2025-12-26)

### ⚠ BREAKING CHANGES

- **deps:** All external runtime dependencies removed
- **datetime:** dayjs is no longer used internally
- Minimum Node.js version is now 20.0.0. Node.js 18 is no longer supported.
- TypeScript configuration now uses bundler moduleResolution

### Features

- **browser:** add pure JavaScript DEFLATE fallback for older browsers ([2a9c29c](https://github.com/cjnoname/excelts/commit/2a9c29cc7020d9834883827142330d136706f07b))
- **browser:** native browser support with zero config ([ea3620c](https://github.com/cjnoname/excelts/commit/ea3620cd363d7fa0c2d8c62293e7b222c2687066))
- **csv:** implement native CSV parser with browser support ([9e9ff9c](https://github.com/cjnoname/excelts/commit/9e9ff9c9e1d9548327a9c6d668f09fb0782d4dda))
- **worksheet:** add column page breaks support ([ad90492](https://github.com/cjnoname/excelts/commit/ad90492a29b6b21f618f0533e20c1e505804e6c6))

### Bug Fixes

- **security:** address CodeQL security warnings ([e89b618](https://github.com/cjnoname/excelts/commit/e89b618872e488e9b5c677fae3389610574817db))

### Performance Improvements

- **datetime:** replace dayjs with high-performance native datetime utilities ([f804811](https://github.com/cjnoname/excelts/commit/f8048114d5a1dbd043017d688976d14172acb867))
- **sax:** optimize XML SAX parser with lookup tables and fast paths ([4dc99eb](https://github.com/cjnoname/excelts/commit/4dc99ebd548f669b38f96e76d7db9ed3078210d5))

### Miscellaneous Chores

- **deps:** remove all runtime dependencies ([15e7b50](https://github.com/cjnoname/excelts/commit/15e7b501344042bb8240eb5404df0bc21b59202e))
- drop Node.js 18 support, require Node.js 20+ ([9568b93](https://github.com/cjnoname/excelts/commit/9568b9354d8fc84a18c03822d2f49e35acd57f3c))

### Code Refactoring

- switch TypeScript moduleResolution from nodenext to bundler ([73c5d94](https://github.com/cjnoname/excelts/commit/73c5d941ae2cd18c99752e3e22415cbb23353cd5))

## [1.6.3](https://github.com/cjnoname/excelts/compare/v1.6.2...v1.6.3) (2025-12-24)

### Bug Fixes

- **docs:** add Vite polyfill configuration for browser usage ([0b06ae9](https://github.com/cjnoname/excelts/commit/0b06ae93fa98dfeb04b17a14de1553df8c8ce526))

## [1.6.2](https://github.com/cjnoname/excelts/compare/v1.6.1...v1.6.2) (2025-12-20)

### Miscellaneous Chores

- release 1.6.2 ([d075b45](https://github.com/cjnoname/excelts/commit/d075b45009aee8e699f02d9ba4f3926415250946))

## [1.6.1](https://github.com/cjnoname/excelts/compare/v1.5.0...v1.6.1) (2025-12-18)

This release includes all changes from 1.6.0 (which was not published to npm).

### Features

- add release-please for automated releases ([735d7ef](https://github.com/cjnoname/excelts/commit/735d7efc114a7aa1c1ebbbbae9894ed2a971dc66))
- **column:** support CellValue types for column headers (fixes [#2740](https://github.com/cjnoname/excelts/issues/2740)) ([18a6eb6](https://github.com/cjnoname/excelts/commit/18a6eb617607e14cf968ebe7f9d72f71c387f7ef))
- **pivot-table:** enhance pivot table support with multiple improvements ([ad9f123](https://github.com/cjnoname/excelts/commit/ad9f123cfe7739438f3bfaf5b96fc70966d68de8))
- **pivot-table:** implement pivot table read and preserve functionality (Issue [#261](https://github.com/cjnoname/excelts/issues/261)) ([9883e5c](https://github.com/cjnoname/excelts/commit/9883e5c6484fe3a15d6d386b22e64fb0cb418839))

### Bug Fixes

- **ci:** add npm publish job to release-please workflow ([a84e54e](https://github.com/cjnoname/excelts/commit/a84e54e2e238e349fe0218af41036d987a8aa089))
- **ci:** add outputs to release-please for better integration ([cddf12a](https://github.com/cjnoname/excelts/commit/cddf12ada88a9e172388c24a61699edc409a0619))
- **pivot-table:** correctly link pivot cache data using pivotCaches from workbook.xml (Issue [#1678](https://github.com/cjnoname/excelts/issues/1678)) ([3bfc50e](https://github.com/cjnoname/excelts/commit/3bfc50eda13f0454cdd3f5a6d01cc7b988153ccb))
- simplify release-please to only manage versions, keep tag-based npm publish ([f1236e6](https://github.com/cjnoname/excelts/commit/f1236e6f36e783cf8012ae29f7dd6c79746f9c64))
- **xlsx:** correct worksheet file naming and pivot table linking (fixes [#2315](https://github.com/cjnoname/excelts/issues/2315)) ([84144cc](https://github.com/cjnoname/excelts/commit/84144cc99a8143810f7bd08d65305ca0b8e352e1))

## [Unreleased]

### Added

- **csv:** unified `workbook.readCsv(input, options)` entry for reading from string/ArrayBuffer/Uint8Array/File/Blob/stream
- **csv:** `workbook.writeCsv(options)` and `workbook.writeCsvBuffer(options)` for writing
- **csv:** `workbook.readCsvFile()` / `workbook.writeCsvFile()` for Node.js file I/O
- **csv:** `createCsvParserStream()` and `createCsvFormatterStream()` factory functions
- **csv:** `detectDelimiter()` helper export
- **csv:** standalone `parseCsvAsync()`, `parseCsvRows()`, `parseCsvWithProgress()` for non-workbook usage
- **archive:** TAR archive support (`TarArchive`, `TarReader`, `tar()`, `untar()`)
- **archive:** ZIP editor (`ZipEditor`, `editZip()`, `ZipEditPlan`)
- **archive:** HTTP range reading (`RemoteZipReader`, `HttpRangeReader`)
- **archive:** ZIP/TAR encryption support (ZipCrypto, AES-256)
- **archive:** Gzip/Zlib compression (`gzip()`, `gunzip()`, `zlib()`, `unzlib()`)
- **archive:** ZIP64 large file support
- **archive:** progress/abort support for all archive operations
- **stream:** new subpath export `@cj-tech-master/excelts/stream`
- **stream:** cross-platform stream error classes (`StreamError`, `StreamStateError`, `StreamTypeError`)
- **stream:** type guards (`isReadableStream`, `isWritableStream`, `isAsyncIterable`, `isTransformStream`)
- **excel:** structured error hierarchy (16 typed error classes extending `ExcelError`)
- **package:** subpath exports for `./zip`, `./csv`, `./stream`

### Removed

- **csv:** `workbook.csv` accessor (use `workbook.readCsv()`, `workbook.writeCsv()`, etc. directly)
- **csv:** legacy type aliases `CsvReadOptions`, `CsvWriteOptions`, `CsvStreamReadOptions`, `CsvStreamWriteOptions` (use `CsvOptions`)
- **stream:** `BufferChunk` (renamed to `ByteChunk`)
- **stream:** `normalizeWritable` / `Writeable` (replaced by `toWritable`)
- **stream:** `EventEmitter` no longer re-exported from stream module (moved to `@utils/event-emitter`)
- **stream:** `once` function (replaced by `onceEvent`)
- **stream:** binary utilities (`textEncoder`, `stringToUint8Array`, etc.) no longer re-exported from stream module (moved to `@utils/binary`)
- **stream:** `ReadWriteBufferOptions` type
- **archive:** `UnzipEntry.isDirectory` (replaced by `UnzipEntry.type: "file" | "directory" | "symlink"`)
- **archive:** archive APIs removed from browser main entry (use `@cj-tech-master/excelts/zip` subpath instead)

### Breaking Changes

- **csv:** `workbook.csv` accessor removed; use `workbook.readCsv()` / `workbook.writeCsv()` / `workbook.writeCsvBuffer()` / `workbook.readCsvFile()` / `workbook.writeCsvFile()` directly
- **csv:** when no delimiter is provided, parsing now auto-detects the delimiter (previously defaulted to ","); pass `delimiter: ","` to keep the old behavior
- **csv:** removed type aliases `CsvReadOptions`, `CsvWriteOptions`, `CsvStreamReadOptions`, `CsvStreamWriteOptions`; use `CsvOptions` instead
- **csv:** parse option `transform` renamed to `rowTransform`; format option `rowDelimiter` renamed to `lineEnding`
- **stream:** `BufferChunk` renamed to `ByteChunk`
- **stream:** `normalizeWritable` / `Writeable` replaced by `toWritable`
- **stream:** `BufferedStream.toUint8Array()` now consumes the buffer (resets to empty after call)
- **archive:** `UnzipEntry.isDirectory` removed; use `entry.type === "directory"` instead
- **archive:** archive APIs removed from browser main entry; use `@cj-tech-master/excelts/zip` subpath
- **excel:** `Image` type renamed to `ImageData` (deprecated alias preserved)
- **excel:** `ZipOptions` renamed to `WorkbookZipOptions` (deprecated alias preserved)
- **eventemitter:** `emit("error")` now throws if no listener (matches Node.js behavior)

## [1.5.0] - 2025-12-13

### Added

- `ZipParser` class for cross-platform ZIP parsing (browser + Node.js)
- `extractAll`, `extractFile`, `listFiles`, `forEachEntry` now work in browser environments
- Native `DecompressionStream` support for browser decompression
- Comprehensive tests for new zip-parser module

### Changed

- Refactored `extract.ts` to use `ZipParser` instead of Node.js streams
- Updated tests to use `TextDecoder` instead of `Buffer.toString()`

### Removed

- Unused `global.d.ts` type declarations

### Breaking Changes

- `extractAll`, `extractFile`, `forEachEntry` now return `Uint8Array` instead of `Buffer`

## [1.4.5] - 2025-12-10

### Added

- Proper typing for `Row` and `Cell` classes with JSDoc comments
- Type safety improvements across `Row`, `Cell`, `Anchor`, `Column`, `Range`, `Image`, `Table` and stream classes

### Changed

- Relaxed return types for row methods (`getRow`, `findRow`, `eachRow`) to improve flexibility

## [1.4.4] - 2025-12-08

### Changed

- Replaced fflate with native zlib for ZIP compression (performance improvement)

### Fixed

- Ignore dynamicFilter nodes in filterColumn parsing (#2972)
- Prevent memory overflow when loading files with many definedNames (#2925)
- Prevent string formula results from being converted to date (#2970)
- Handle missing `r` attribute in row and cell elements (#2961)

## [1.4.3] - 2025-12-05

### Fixed

- Date and duration format handling

## [1.4.2] - 2025-12-04

### Changed

- Relaxed performance test thresholds for CI and Windows compatibility

## [1.4.1] - 2025-12-03

### Changed

- Optimized parsing of large data validation ranges (performance improvement)

## [1.4.0] - 2025-12-02

### Changed

- Code cleanup and optimizations

## [1.3.0] - 2025-11-28

### Changed

- Updated all dependencies to latest versions

### Added

- Cell format parser
- Improved browser compatibility

## [1.1.0] - 2025-11-15

### Added

- Major improvements and bug fixes

## [1.0.0] - 2025-10-30

### 🎉 First Stable Release

This is the first stable 1.0 release of ExcelTS! The library is now production-ready with comprehensive features, excellent TypeScript support, and thorough testing.

### Added

- Full TypeScript rewrite with strict typing
- Named exports for better tree-shaking
- Browser testing support with Playwright
- Husky v9 for Git hooks
- lint-staged for pre-commit checks
- Prettier configuration for consistent code style
- .npmignore for optimized package publishing
- Comprehensive browser and Node.js version requirements documentation

### Changed

- Public API and packaging updates
- All default exports converted to named exports
- Updated all dependencies to latest versions
- Migrated testing framework from Mocha to Vitest
- Switched bundler from Webpack to Rolldown
- Build system using tsgo (TypeScript native compiler)
- Target ES2020 for better compatibility
- Node.js requirement: >= 18.0.0 (previously >= 12.0.0)
- Browser requirements: Chrome 85+, Firefox 79+, Safari 14+, Edge 85+, Opera 71+

### Improved

- Enhanced type safety with proper access modifiers
- Performance optimizations in build process
- Reduced package size by excluding source files from npm publish
- Optimized IIFE builds with conditional sourcemaps
- Better error handling and logging (development-only console warnings)
