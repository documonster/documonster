# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

> **Project rename — `excelts` → `documonster`.**
> This package was previously published as `@cj-tech-master/excelts`. Starting
> with **v11.0.0** it is published as **`documonster`** (no org scope). To
> migrate, replace every `@cj-tech-master/excelts` import specifier with
> `documonster` (subpaths are unchanged, e.g.
> `@cj-tech-master/excelts/excel` → `documonster/excel`). Release history prior
> to the rename (≤ v10.2.0) lived under the `excelts` project and is not
> reproduced here.

## [11.0.0](https://github.com/documonster/documonster) (2026-06-18)

### ⚠ BREAKING CHANGES

- **Renamed package** from `@cj-tech-master/excelts` to `documonster`. Update
  all imports: `@cj-tech-master/excelts/<module>` → `documonster/<module>`.
- **excel:** Unified `Workbook` IO naming — `loadXlsx`→`read`,
  `readXlsxFile`→`readFile`, `writeXlsx`→`writeFile`, `toXlsxBuffer`→`toBuffer`,
  `readXlsxStream`/`writeXlsxStream`→`readStream`/`writeStream`.
- **excel:** Renamed error class `XmlParseError` → `XlsxParseError`.
- **archive:** Public API moved under the `Archive` namespace
  (`import { Archive } from "documonster/archive"`).
- **word:** `toBuffer`/`toBase64`/`packageDocx` now take an options object
  instead of a bare `compressionLevel` number.

### Improvements

- Project-wide architecture consistency pass: all error subclasses thread
  `{ cause }`, ~234 bare `throw new Error` replaced with module error classes,
  same-module imports unified to path aliases across all nine modules, and
  `TemplateError` consolidated into the word errors module.
