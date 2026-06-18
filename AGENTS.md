# AGENTS.md

## Project Overview

**documonster** — zero-dependency TypeScript toolkit. Nine modules: Excel, Word, Formula, PDF, CSV, Markdown, XML, Archive, Stream.

- Zero runtime dependencies — never add packages to `dependencies`
- Cross-platform: Node.js 22+ and modern browsers
- ESM-first with CommonJS compatibility

## Hard Rules

1. **No runtime dependencies.** All functionality must be self-contained.
2. **Prefer native APIs.** If a browser or Node.js built-in can do the job, use it instead of writing a custom implementation. Only roll your own when the native API is missing, insufficient, or unavailable on a target platform.
3. **No circular imports.** Enforced by `import/no-cycle`.
4. **Named exports only.** No default exports.
5. **Respect module dependency direction.** See layer diagram below. Never introduce upward dependencies.
6. **Run `pnpm check` then `pnpm format` before committing.**

## Bug Fixing & Code Changes

- **Fix root causes, not symptoms.** Trace every bug to its origin. Never patch over a problem — fix the underlying logic.
- **Read before writing.** Before modifying any file, read the surrounding code to understand context, patterns, and invariants. Do not assume — verify.
- **Match existing patterns.** Follow the conventions already present in the file and module. When unsure, search for similar code in the codebase first.
- **No speculative code.** If you are uncertain about an API, type, or behavior, look it up in the source. Do not guess.
- **Fix it properly.** If the correct fix requires changing multiple files, refactoring a helper, or adjusting an interface — do it. Do not take shortcuts to minimize the diff. The goal is the best solution, not the smallest patch.
- **Do not be afraid of large changes.** If the best solution means rewriting a function, restructuring a module, or breaking an existing API — do it. Correctness and quality come first. Tests exist to catch regressions; use them.
- **Do not touch unrelated files.** Only modify files directly relevant to the task. Never make drive-by changes to code you were not asked to work on.
- **Verify your fix.** After making changes, run the relevant tests or `pnpm check` to confirm the fix works. Never claim a problem is resolved without evidence.
- **No over-engineering.** Solve the actual problem, not a hypothetical general case. If unsure whether a design is over-engineered, summarize the tradeoffs and ask before proceeding.

## Commands

```bash
pnpm i                  # Install (use pnpm, not npm/yarn)
pnpm check                # Type check + lint + format check — run before commit (do not run lint separately)
pnpm format               # Prettier format — run before commit
pnpm test                 # All tests
pnpm build                # Production build

# Single test file
pnpm exec vitest run src/modules/excel/__tests__/cell.test.ts
# Pattern match
pnpm exec vitest run -t "should handle empty cells"
```

## Project Structure

```
src/
├── modules/
│   ├── excel/          # Workbook, Worksheet, Cell; stream/ xlsx/
│   ├── word/           # DocxDocument, DocumentBuilder, readDocx, packageDocx
│   ├── formula/        # Tokenizer, parser, evaluator, 433 functions, spill engine
│   ├── pdf/            # core/ builder/ font/ render/ reader/ + excel-bridge.ts + word-bridge.ts
│   ├── csv/            # Parsing/formatting + streaming
│   ├── markdown/       # GFM table parsing/formatting
│   ├── xml/            # SAX/DOM parser, query engine, writer
│   ├── archive/        # ZIP/TAR compression
│   └── stream/         # Cross-platform streaming primitives
├── utils/              # Shared: errors, datetime, fs, binary, crypto
└── test/               # Test utilities and fixtures
```

## Module Dependency Layers

```
Layer 5:  pdf      → excel (only excel-bridge.ts), word (only word-bridge.ts), archive, utils
Layer 4:  excel, word → formula, archive, xml, csv, markdown, stream, utils
Layer 3:  formula  → utils    (independent calc engine; no excel imports)
Layer 2:  csv, archive → stream, utils
Layer 1:  xml, markdown, stream → utils
Layer 0:  utils    (no module dependencies)
```

- Modules may only import from **lower** layers — never sideways or upward.
- **Sole exceptions**:
  - `pdf/excel-bridge.ts` may import from `@excel/`. No other file in `pdf/` may.
  - `pdf/word-bridge.ts` may import from `@word/`. No other file in `pdf/` may.
  - `word/bridge/excel-bridge.ts` may import from `@excel/`. No other file in `word/` may.
  - `formula/` defines structural interfaces (`WorkbookLike`, `WorksheetLike`, `CellLike`) that `excel/` implements; `formula/` never imports concrete types from `@excel/*`.
- `utils/` must never import from any module.

## Path Aliases

`@excel/*`, `@word/*`, `@formula/*`, `@pdf/*`, `@csv/*`, `@markdown/*`, `@xml/*`, `@archive/*`, `@stream/*` → `./src/modules/<name>/*`
`@utils/*` → `./src/utils/*` | `@test/*` → `./src/test/*`

Use aliases for **all** module imports — both cross-module (`@archive/...` from excel) and same-module (`@excel/cell` from within excel). This matches the IDE auto-import setting (`importModuleSpecifier: "non-relative"`) and keeps imports stable when files move. The only exception is `src/utils/` (Layer 0), whose internal files use relative paths (`./errors`, `./glob`).

## Code Style

- **Type-only imports**: `import type { Foo } from "..."`
- **Error handling**: Extend `BaseError` from `@utils/errors`, use `{ cause }` for chaining.
- **Files**: kebab-case. **Browser variants**: `*.browser.ts`.
- **Formatting**: Handled entirely by Prettier — just run `pnpm format`.
- **Tests**: Vitest, in `__tests__/*.test.ts`. Timeout: 30s.

## Example Output

All runnable examples write output to `tmp/` under the project root. This directory is gitignored.
