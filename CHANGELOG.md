# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [5.1.0](https://github.com/cjnoname/excelts/compare/v5.0.4...v5.1.0) (2026-01-25)


### Features

* **archive:** add noSort option and getChildCount method ([742dade](https://github.com/cjnoname/excelts/commit/742dadede78221715b82fa97bd9ba5810bba6802))

## [5.0.4](https://github.com/cjnoname/excelts/compare/v5.0.3...v5.0.4) (2026-01-25)


### Bug Fixes

* **pivot-table:** preserve worksheetSource name attribute for table references ([#45](https://github.com/cjnoname/excelts/issues/45)) ([ef1722b](https://github.com/cjnoname/excelts/commit/ef1722b69c3b84a22e4083024958797f9c1b5b6a))

## [5.0.3](https://github.com/cjnoname/excelts/compare/v5.0.2...v5.0.3) (2026-01-25)


### Bug Fixes

* **excel:** add default cfvo and color for dataBar conditional formatting ([d7abd28](https://github.com/cjnoname/excelts/commit/d7abd28db0cc8acd987324c6a3057d39d6a2ce17))
* remove redundant no-op string replacement in workbook roundtrip test ([09ee8dc](https://github.com/cjnoname/excelts/commit/09ee8dc4ea0336be72f0e1c4236548c88e305ddf))

## [5.0.2](https://github.com/cjnoname/excelts/compare/v5.0.1...v5.0.2) (2026-01-23)


### Bug Fixes

* **browser:** fix drawing parsing failure in loadFromFiles path ([98c7ee0](https://github.com/cjnoname/excelts/commit/98c7ee0a91caf82a5f43bce8cdd97970152ef2ec))

## [5.0.1](https://github.com/cjnoname/excelts/compare/v5.0.0...v5.0.1) (2026-01-23)


### Bug Fixes

* **xlsx:** preserve metadata attributes during round-trip ([#41](https://github.com/cjnoname/excelts/issues/41)) ([5f5d54d](https://github.com/cjnoname/excelts/commit/5f5d54d1825ce5e64bd69ded42a2511c8070f8bf))

## [5.0.0](https://github.com/cjnoname/excelts/compare/v4.2.3...v5.0.0) (2026-01-22)


### ⚠ BREAKING CHANGES

* **excel:** None - passthrough is opt-in via workbook options

### Features

* **excel:** add chart and drawing passthrough preservation ([a4ea35e](https://github.com/cjnoname/excelts/commit/a4ea35e2ea4297d7e7a06f9dbc83ffd37fd90de4))


### Bug Fixes

* **stream:** remove unreachable streamError rethrow ([b80904d](https://github.com/cjnoname/excelts/commit/b80904dee784b6dc0c1ed92f846194bccec0dcc9))


### Performance Improvements

* move HAN CELL namespace handling from SAX parser to BaseXform ([cc11b20](https://github.com/cjnoname/excelts/commit/cc11b206bec2e0eabf5bd2164743bca0cc43fc97))

## [4.2.3](https://github.com/cjnoname/excelts/compare/v4.2.2...v4.2.3) (2026-01-15)


### Bug Fixes

* support HAN CELL xlsx files with namespace prefixes ([88820eb](https://github.com/cjnoname/excelts/commit/88820eb94192c2b9a10c7794cf698aaa66254387))

## [4.2.2](https://github.com/cjnoname/excelts/compare/v4.2.1...v4.2.2) (2026-01-12)


### Bug Fixes

* **archive:** stabilize streaming unzip and browser parsing ([a503090](https://github.com/cjnoname/excelts/commit/a50309085bceb6986a07d098c57749b4c1476f5a))
* **excel:** make legacy form controls OOXML-valid ([fe7a444](https://github.com/cjnoname/excelts/commit/fe7a444a3586977089ee6b8ad9b24f13d8830152))


### Performance Improvements

* **archive:** speed up streaming unzip hot path ([f808a37](https://github.com/cjnoname/excelts/commit/f808a37255750a26e2db10a95de98461f52b8241))

## [4.2.1](https://github.com/cjnoname/excelts/compare/v4.2.0...v4.2.1) (2026-01-10)


### Bug Fixes

* **archive:** keep ZIP parse streaming for large entries ([c88c61c](https://github.com/cjnoname/excelts/commit/c88c61cc3b3e22b693147303be1e500cd4402a6a))
* **build:** copy LICENSE and THIRD_PARTY_NOTICES to dist/iife ([0919d4d](https://github.com/cjnoname/excelts/commit/0919d4d6313f4b54dd3dfb20d450be287b71830a))
* **excel:** improve legacy form checkbox anchors and controls ([7805a16](https://github.com/cjnoname/excelts/commit/7805a16a85f81eaccf542c6ad093beb5e7d1e73d))

## [4.2.0](https://github.com/cjnoname/excelts/compare/v4.1.0...v4.2.0) (2026-01-10)


### Features

* **excel:** add legacy Form Control Checkbox support ([e7d8c4e](https://github.com/cjnoname/excelts/commit/e7d8c4e4b650aba90d83bb9a2a7d6934945e8a7e))

## [4.1.0](https://github.com/cjnoname/excelts/compare/v4.0.4...v4.1.0) (2026-01-10)


### Features

* **excel:** add Office Online-compatible in-cell checkboxes ([8ac37ef](https://github.com/cjnoname/excelts/commit/8ac37efb46a5f33e85462ef53bf8c6a6cc38025d))


### Bug Fixes

* **excel:** hydrate loaded table rows for mutations ([4f97ebb](https://github.com/cjnoname/excelts/commit/4f97ebb00671c157fac89cdb789fef8727b6deaa))

## [4.0.4](https://github.com/cjnoname/excelts/compare/v4.0.3...v4.0.4) (2026-01-06)


### Bug Fixes

* make generated types NodeNext-safe ([b618378](https://github.com/cjnoname/excelts/commit/b618378a19871d2175a452cc658fbd8859d50704))

## [4.0.3](https://github.com/cjnoname/excelts/compare/v4.0.2...v4.0.3) (2026-01-04)


### Bug Fixes

* **excel:** keep table formulas readable ([3972145](https://github.com/cjnoname/excelts/commit/3972145fe1eec92a3d0895583e6dc91eb9aea9fe)), closes [#29](https://github.com/cjnoname/excelts/issues/29)

## [4.0.2](https://github.com/cjnoname/excelts/compare/v4.0.1...v4.0.2) (2026-01-04)


### Bug Fixes

* **excel:** make table structured refs work ([302e682](https://github.com/cjnoname/excelts/commit/302e6827bb0a286bddaeeaf7abf563ada77cda08)), closes [#26](https://github.com/cjnoname/excelts/issues/26)

## [4.0.1](https://github.com/cjnoname/excelts/compare/v4.0.0...v4.0.1) (2026-01-03)


### Bug Fixes

* **types:** avoid .d.ts specifiers in declarations ([0e5d37f](https://github.com/cjnoname/excelts/commit/0e5d37f63b650ad15f02d13c9548899023152c95))

## [4.0.0](https://github.com/cjnoname/excelts/compare/v3.1.0...v4.0.0) (2026-01-02)


### ⚠ BREAKING CHANGES

* The main package entrypoints no longer re-export the internal stream utility surface. If you were importing stream helpers from the root package, migrate to standard Web Streams (ReadableStream/WritableStream) or pin to an older version.
* **exports:** Browser build no longer exports the exceljs-compatible stream.xlsx namespace. Use top-level WorkbookWriter/WorkbookReader/WorksheetWriter/WorksheetReader exports instead.

### Features

* **csv:** add valueMapperOptions for decimalSeparator ([b93d66e](https://github.com/cjnoname/excelts/commit/b93d66e5488e4c9833c913c901bb78f9bfb8a1cf))
* **exports:** unify node and browser entrypoints ([c8bc979](https://github.com/cjnoname/excelts/commit/c8bc979725b97b33eac7c8433fe9a60c593483f3))
* remove stream utility re-exports ([ea16582](https://github.com/cjnoname/excelts/commit/ea16582e8434d845ced099ffca80a63c970d3da2))
* **streaming:** browser streaming support ([381817c](https://github.com/cjnoname/excelts/commit/381817ce46b7e367542d251a4da06e98fa747810))
* **streaming:** support Web Streams across environments ([204ba36](https://github.com/cjnoname/excelts/commit/204ba365f4100e05d5fe668ab71f3550a789f94a))
* **xlsx:** allow deterministic zip entry timestamps ([d17da6a](https://github.com/cjnoname/excelts/commit/d17da6a8eea6db0b2c0fff208fcc11c13d71723f))
* **xlsx:** store data validations as ranges ([09c2a40](https://github.com/cjnoname/excelts/commit/09c2a4062c2a4daf1b703dfca195fff0f8dc1987))


### Bug Fixes

* **build:** rewrite tsconfig path aliases in dist outputs ([6791d4e](https://github.com/cjnoname/excelts/commit/6791d4ea91f0296f70a9914d9817fc96cc3e6f53))
* post-merge csv parsing + pivot test import ([8f31be3](https://github.com/cjnoname/excelts/commit/8f31be3b6afd806363d419f3f88365229023a11c))
* **security:** address CodeQL findings ([77dafd9](https://github.com/cjnoname/excelts/commit/77dafd9c012bf6ca45f6ec245d320c7edc1ab7ed))
* **stream:** avoid extra args in browser transform ([6ddacdd](https://github.com/cjnoname/excelts/commit/6ddacddf719aaf784100179dabac26fdd29432bb))
* **stream:** handle browser transform/flush arity safely ([d80b29d](https://github.com/cjnoname/excelts/commit/d80b29de30d667125b51684cb3d2d0aa23bfe634))
* **stream:** make transform arity dispatch CodeQL-friendly ([4ffdd8b](https://github.com/cjnoname/excelts/commit/4ffdd8b66847e809a31ed0d43f69051793083bb2))
* **stream:** use direct call with known signature to satisfy CodeQL ([ea53170](https://github.com/cjnoname/excelts/commit/ea531702c0450c6116c26538da17e48f16c2f6a3))
* **stream:** use proper type assertion for userFlush call ([51d781c](https://github.com/cjnoname/excelts/commit/51d781ca286c7747789bb066071112dd5012bb80))
* **test:** align csv mapper typing ([ab0d509](https://github.com/cjnoname/excelts/commit/ab0d509c1c1d8b8867a6a4ec1be9f74686f03888))


### Performance Improvements

* **csv:** optimize streaming parser hot path ([50bf1ef](https://github.com/cjnoname/excelts/commit/50bf1ef27b59eb8994934262c5deb2217b3346af))

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
