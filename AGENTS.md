# AGENTS.md

## Project Overview

**documonster** — zero-dependency TypeScript toolkit. Modules: Excel, Word, Formula, PDF, CSV, Markdown, XML, Archive, Stream, Draw, Mermaid.

- Zero runtime dependencies — never add packages to `dependencies`
- Cross-platform: Node.js LTS lines (22, 24) and Current (26), plus modern browsers
- ESM-only; CommonJS consumers load it through `require(esm)` (Node >= 22.13)

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
- **Moving a member between namespaces? Run `pnpm build`.** It runs `verify:treeshake`, which `pnpm check` does not — a member re-homed into another namespace object can pull that namespace's whole module into every bundle that imports the old one.
- **No over-engineering.** Solve the actual problem, not a hypothetical general case. If unsure whether a design is over-engineered, summarize the tradeoffs and ask before proceeding.

## Commands

```bash
pnpm i                  # Install (use pnpm, not npm/yarn)
pnpm check                # Type check + lint + format check — run before commit (do not run lint separately)
pnpm format               # Prettier format — run before commit
pnpm test                 # All tests
pnpm build                # Production build

# Workspace satellites (packages/*) — need `pnpm build:esm` first, see below
pnpm verify:packages      # satellites must use the public API only
pnpm type:packages
pnpm test:packages
pnpm build:packages

pnpm verify:doc-examples  # documented imports & members must exist — see "Documentation" below
pnpm verify:doc-links     # local Markdown targets & headings must exist
pnpm verify:guide-tables  # the calls in the migration guides' tables must compile
pnpm verify:guide-examples  # every ts code block of the migration guides compiles whole (`ts browser`, `ts before`)
pnpm verify:examples      # every example must actually run — see "Examples" below
pnpm verify:compat        # the last release's examples still type-check and run
pnpm api:changes --migration  # before a release: regenerate docs/api-changes-since-v<last>.md and
                              # the generated parts of MIGRATION.md (needs build:esm + build:packages)
pnpm verify:excelts-guide     # the excelts guide accounts for every name excelts 10.2.0 published; --write
                              # regenerates its name index and docs/api-changes-from-excelts-10.2.0.md
pnpm verify:exceljs-guide     # the exceljs guide accounts for every name exceljs 4.4.0 published
pnpm verify:exceljs-parity    # the guides' ports write the same workbooks as exceljs 4.4.0 (and
pnpm verify:excelts-parity    # excelts 10.2.0) did — both installed from npm

# Single test file
pnpm exec vitest run src/modules/excel/core/__tests__/cell.test.ts
# Pattern match
pnpm exec vitest run -t "should handle empty cells"

# Run examples (same runner as the gate; --filter narrows by path substring)
pnpm example --filter pdf-page-setup
```

## Project Structure

```
src/
├── modules/
│   ├── excel/          # core/ (Workbook, Worksheet, Cell, …) surface/ stream/ xlsx/
│   ├── word/           # DocxDocument, DocumentBuilder, readDocx, packageDocx
│   ├── formula/        # Tokenizer, parser, evaluator, 467 functions, spill engine
│   ├── pdf/            # core/ builder/ font/ render/ reader/ + excel-bridge.ts + word-bridge.ts + word-chart-bridge.ts + word-layout-to-pdf.ts
│   ├── csv/            # Parsing/formatting + streaming
│   ├── markdown/       # GFM table parsing/formatting
│   ├── xml/            # SAX/DOM parser, query engine, writer
│   ├── archive/        # ZIP/TAR compression; core/ shared primitives
│   ├── draw/           # Shared drawing engine: display-list IR, one walker, SVG surface
│   ├── mermaid/        # Mermaid text → DrawList (21 diagram types)
│   └── stream/         # Cross-platform streaming primitives; core/ shared primitives
├── utils/              # Shared: errors, datetime, fs, binary, crypto
└── test/               # Test utilities and fixtures

packages/               # Workspace satellites — MAY have runtime dependencies
└── mcp/                # @documonster/mcp — Model Context Protocol server (node-only)
```

## Workspace Layout

This is a pnpm workspace. The repository **root is the `documonster` package**
itself; `packages/*` holds satellites.

**Why satellites exist.** The core must have zero runtime dependencies (rule 1),
but some things genuinely need them — the MCP server needs the MCP SDK. A
satellite package keeps that dependency out of `documonster` entirely while the
code still lives in this repo, so a core API change and its MCP counterpart land
in one PR instead of drifting across two repositories.

**Satellite rules — machine-enforced by `scripts/verify-package-imports.ts`
(`pnpm verify:packages`, part of `pnpm check`):**

1. A satellite imports `documonster` **only through its published `exports` map**
   (`documonster/excel`, `documonster/csv`, …). Internal aliases (`@excel/*`,
   `@utils/*`) and relative reaches into `../../src` are build failures. This
   makes every satellite an honest consumer of the public API — and a real check
   on whether that API is sufficient.
2. A satellite is node-only and ESM-only. It does not participate in the core's
   IIFE build matrix.
3. Satellites are excluded from the root `tsconfig.json` and the root
   `vitest.config.ts`; each carries its own.

**A satellite's own surface follows its host protocol, not the core's vocabulary.**
`@documonster/mcp` publishes a server (`createServer`, `resolveConfig`), wire-named tools
(`sheet_read`, snake_case and stable) and the error codes a model retries on
(`McpToolError.code`), because those are what an MCP client and the SDK expect. The
"Public API Conventions" below are the core's; they apply to a satellite only where it
re-exposes a core capability, which it does by calling the public entry, never by
re-publishing it under a name of its own. One core rule does apply to it unchanged: a type
its public signatures name is importable from its entry — `pnpm verify:public-types` checks
`@documonster/mcp` as it checks the core entries.

**Consequence for the dev loop:** because satellites resolve the core through
`exports`, `dist/esm` + `dist/types` must exist before they will type-check.
Run `pnpm build:esm` once (not the full `pnpm build`), then iterate.

```bash
pnpm build:esm          # prerequisite: produces dist/esm + dist/types
pnpm verify:packages    # no internal imports (no build needed)
pnpm type:packages
pnpm test:packages
pnpm build:packages
```

**Watch out:** root `pnpm check` builds ESM/types first because type-aware lint
resolves `packages/*` through the public exports map; root tsc itself still
excludes satellites, whose dedicated type check runs under `pnpm test` and CI.
Root `pnpm test` (and its `test:all` alias) runs
everything in the required order (core node tests →
`build:esm` → package tests → browser tests; the browser step wipes `dist/`, so
package tests must precede it). `pnpm check` includes `verify:packages` because
that is a pure source scan; `type:packages` is left out because it needs a build,
and the pre-commit hook therefore runs it only when `dist/types` already exists.

**Releases — one version for the whole repository.** `@documonster/mcp` always
carries the same version as `documonster`. release-please tracks a single
package (`.`) and writes the new version into `packages/mcp/package.json` through
`extra-files`, so the two stay identical by construction — there is no second
manifest entry, no `mcp-v*` tag and no separate MCP changelog. MCP changes appear
in the root `CHANGELOG.md` like any other change, and a commit that only touches
`packages/mcp` still bumps the shared version.

That coupling is deliberate. `@documonster/mcp` depends on
`documonster: workspace:*`, which `pnpm publish` rewrites to an **exact** version
— so an MCP release that skipped a core release would leave MCP users pinned to
an old core forever. Releasing both together keeps that pin fresh. It also means
the core must reach npm first, which is why `publish-mcp` needs `publish` in
`.github/workflows/release.yml`.

Canaries follow the same rule. Both manifests receive the deterministic version
`<current>-canary.sha.<commit>`, then core publishes before MCP. The version must
not contain a timestamp: a retry after only one package published must reproduce
the same version so the idempotent checks can skip it and finish the pair.

**A scoped package needs `--access public` on the command line.** pnpm does not
apply `publishConfig.access`, and npm defaults a scoped package to restricted, so
`pnpm publish` prints `✅ Published` while the package stays invisible to
everyone but the owner — `documonster` is unscoped and never showed this. The
publish steps therefore pass `--access public`, and every verify step reads the
registry **unauthenticated**, because an authenticated read succeeds for a
restricted package and would hide the problem.

Do not try to give the satellite its own version line with the `linked-versions`
plugin: it skips any package whose component resolves to empty, and the core's
component is empty by design (`include-component-in-tag: false`, which is what
keeps the `v0.9.0` tag format instead of `documonster-v0.9.0`).

## Module Dependency Layers

```
Layer 5:  pdf      → excel (only excel-bridge.ts + word-chart-bridge.ts), word (only word-bridge.ts + word-chart-bridge.ts + word-layout-to-pdf.ts), draw, archive, xml, utils
Layer 4:  excel, word → formula, draw, archive, xml, csv, markdown, stream, utils; mermaid → draw, utils
Layer 3:  formula  → utils    (independent calc engine; no excel imports); draw → utils, archive (only api/png.ts)
Layer 2:  csv, archive → stream, utils
Layer 1:  xml, markdown, stream → utils
Layer 0:  utils    (no module dependencies; owns the shared TTF parser and font discovery)
```

### The `draw` module

`draw` is the shared drawing engine. It owns a structured display list
(`DrawList`), the single walker that consumes it (`renderDrawList`), the abstract
backend interface (`DrawSurface`), an SVG serialiser and a rasteriser. A producer
that emits a `DrawList` therefore gets markup and pixels from `documonster/draw`
alone, and a PDF page from `documonster/pdf` — the claim below is testable from
outside the repository, not only inside it.

**Why it exists.** The renderers used to pass an _SVG string_ between themselves:
the chart engine serialised SVG, then the Node PNG fallback re-parsed it with a
regex scanner and the PDF importer parsed it again with a second, differently
capable parser. Each backend therefore had its own reading of every attribute,
and the same picture came out differently depending on which one you asked —
dashes disappeared, opacity was dropped, rotations mirrored, 8-digit colours were
read with the wrong byte order. A display list removes the round trip.

**The rule: one walker, many surfaces.** `renderDrawList` owns transform
composition, stroke/dash scaling and text rotation, and hands surfaces
**absolute, already-transformed** coordinates in a Y-down space. A surface never
implements a transform stack, so a new backend is a few dozen lines instead of a
parallel renderer. Adding a _producer_ (a chart type, a diagram engine) gets every
existing backend for free; adding a _backend_ serves every producer.

A surface lives with the engine when it needs nothing but the IR, and next to its
target when it draws onto something the engine cannot know about:

| Surface    | Location                     | Output                                                   |
| ---------- | ---------------------------- | -------------------------------------------------------- |
| SVG        | `draw/svg.ts`                | Markup in the subset the rest of the library understands |
| Raster     | `draw/raster/surface.ts`     | RGBA pixels via `BasicRasterCanvas`                      |
| PDF vector | `pdf/render/draw-surface.ts` | Page operators; owns the Y flip                          |

The rasteriser used to live in `excel/chart/render/`, which was an accident of where
charts were written rather than a statement about what it does: it paints a display
list and knows nothing about a workbook. It sits in `draw/` now, alongside the glyph
rasteriser and the stroke font it needs, and `documonster/draw` exports it.

**Fonts are shared with `pdf`, at Layer 0.** Drawing text needs a font file read, and
that was written twice: `pdf/font/ttf-parser.ts` to embed and subset a face, and
`draw/raster/glyph-rasterizer.ts` to rasterise glyphs, each with its own big-endian
readers, table directory, `cmap` formats 4 and 12, `hmtx`, `loca` and `.ttc` header —
around 330 duplicated lines that were free to disagree.

The duplication was not the real cost. **System font _discovery_ could only be built
once, and it was built on the PDF copy**, so `draw` could not reach it: `pdf` is
Layer 5. The rasteriser therefore had a hardcoded list of five Latin filenames per
platform, took the first that existed, and used that one face for everything — and
since none of Arial, Helvetica, DejaVu or Liberation contains a Han glyph, a PNG of a
Chinese label came out **blank** while the PDF of the same display list embedded
`Songti SC` and rendered correctly. Nothing reported it: a missing outline advanced
the pen and painted nothing.

So the split is now by job, not by module:

| Job                                                | Lives in                          |
| -------------------------------------------------- | --------------------------------- |
| TTF/TTC tables — directory, `cmap`, `hmtx`, `loca` | `@utils/font-ttf`                 |
| Walking a composite glyph's component records      | `@utils/font-ttf`                 |
| Finding an installed face that covers a text       | `@utils/font-discovery`           |
| `glyf` outlines → contours, and a `RasterFont`     | `draw/raster/glyph-outline.ts`    |
| Contours → an alpha bitmap                         | `draw/raster/glyph-rasterizer.ts` |
| Choosing and chaining faces for the rasteriser     | `draw/raster/system-raster-font`  |
| PDF descriptors, subsetting, Type3                 | `pdf/font/*`                      |
| Word subsetting, ODTTF obfuscation                 | `word/font/font-embed.ts`         |

Layer 0 is forced, not chosen: `draw` may only import `utils`, so anything shared
with it has to be there — and `utils` already held `font-data`, `font-metrics` and
`text-measure`. Outline _decoding_ deliberately stayed in `draw`, because PDF embeds
`glyf` bytes verbatim and never needs a point; lowering it would put code at Layer 0
that exactly one consumer can use.

**Word was the third copy, and it had already drifted.** `word/font/font-embed.ts`
carried its own table directory, `cmap` formats 4 and 12, `head`, `maxp` and `loca` —
and accepted only Windows (platform 3) `cmap` subtables, where the shared reader also
accepts the Unicode platform. A font publishing only a `(0,3)` table, which several
macOS system faces do, was therefore declared unsubsettable and embedded whole, with
nothing reported. That is the second bug this split produced, after the blank CJK PNG,
and the reason the rule is one reader rather than one reader per consumer.

**A composite glyph's component records are walked in one place too.** A record's
length depends on its own flags, and there were four copies of that arithmetic: the
rasteriser needed the transform, the PDF subsetter needed both to collect dependencies
and to renumber glyph IDs, and the Word subsetter needed the dependencies. Three of
them could disagree with the fourth about where a record ends, and none of them bounded
the read to the glyph — they stepped with a bare cursor and relied on a `DataView`
throwing once it left the buffer. `glyphComponents` yields `flags`, the component's
glyph ID, **the offset of that ID** so a subsetter can write a new one back, the
placement offsets and the 2×2 transform; each consumer reads the fields it needs.

Two things did **not** move, and both are policy rather than mechanism. Whether a code
point _requires_ a real face is passed in (`RequiresFacePredicate`): PDF draws an arrow
or a checkbox with a Type3 glyph so those must not disqualify a Chinese font, while the
rasteriser has no such fallback and needs everything. And `PdfFontError` stayed in
`pdf` — it is published, `isPdfError()` answers for it, and the exporter branches on
it — so `pdf/font/ttf-parser.ts` remains as a boundary that translates Layer 0's
`FontParseError` into it. Both files are thin and exist to keep a contract, not to
forward calls.

**The rasteriser resolves each character against a chain, not a font.** PDF can fall
back to a Type3 drawing for a glyph its chosen face lacks; a rasteriser has nothing to
fall back to but another face, and no single font covers `Mixed 混合 ABC`. The Latin
face is ordered _before_ the discovered one on purpose — a CJK face carries Latin
glyphs too, and putting it first set the ASCII of a mixed label in Songti's serif hand
while the SVG of the same drawing used Arial. A character nothing can draw is recorded
in `RgbaImage.uncoveredCodePoints` rather than silently skipped, and
`RasterizeOptions.fonts` / `useSystemFonts` let a caller supply faces outright — which
is the only route to a non-ASCII glyph in a browser, and the way to make a render
reproducible across machines. Supplied fonts belong to the **canvas**, not to the
process: routing them through a module-level registry left one render's fonts in place
for the next render that passed none, so there is no such registry in the public
surface. `fonts` takes bytes, a parsed `RasterFont` (parse once, reuse across renders —
bytes are re-parsed every call) or `{ data, collectionIndex }`, which is the only way to
name a face inside a `.ttc`. Bytes are recognised structurally rather than with
`instanceof Uint8Array`, because a typed array from another realm fails that test and
was being dropped in silence — in a browser, where supplying bytes is the only option.
Normalising a supplied font touches no filesystem, so it lives in `glyph-outline.ts` and
is shared; only _discovery_ is a platform variant, which is what keeps the browser
sibling at 42 lines instead of a second copy of it.

**A formatting control is not a character.** A variation selector, a joiner, a bidi
isolate or a zero-width space has no glyph in any font, so it is never looked up, never
charged an advance and never reported as uncovered. Doing otherwise told callers to
install a font for `U+FE0F`, and — because a line's advances are normalised to its
measured width — spread the phantom half-em over every _visible_ glyph in the label.
Whether it appeared at all depended on whether the host's face happened to carry an
empty glyph for it, so the report differed between machines.

The rule is expressed as two predicates, and the difference between them is the point.
`isZeroWidthCodePoint` is "takes no width" and is what `@utils/text-measure` consults;
`isNonPrintingControl` is "draws nothing" and is what the rasteriser consults. They
differ by **combining marks**, which take no width but do draw. **Tab is in neither**: it
draws nothing yet occupies space, so the measurer charges it and the rasteriser advances
the pen for it without reporting a gap — the one control character that is width without
ink.

Both halves are needed. The rasteriser skipping a control is not enough on its own,
because a line's glyph advances are normalised to its _measured_ width: while the
measurer still charged a zero-width space half an em, inserting one moved every visible
glyph in the label even though it was no longer drawn.

**Style reaches the face, not just the width.** `family`, `bold` and `italic` used to
affect only measurement, so a label asking for `Courier New` was drawn in Arial and a
bold one was drawn regular, then stretched to the measured width — the right size in the
wrong shape, while SVG and PDF honoured the same request. Discovery ranks candidates
`family` → `italic` → `weight` (`FaceStyle`, `styleDistance`), because a wrong slant is
seen immediately and a weight one step off is not. The Latin slot is resolved for the
requested style too, not only the CJK one: it is consulted first, so leaving Arial in
front of a requested `Courier New` claimed every ASCII character and the request had no
visible effect. Its fallback list is Latin families by name — searching for "covers A, a
and 0" with no family preference answers with whichever face heads the built-in order,
which is a CJK one, so `bold` produced a Chinese face for Latin text.

Measurement still comes from the static advance tables rather than the chosen face, and
that is deliberate: layout has to be identical on every machine, and it cannot be if it
depends on which fonts are installed.

**Discovery keeps asking until it stops making progress.** It may return a face
covering only _part_ of what was requested, which is deliberate — most of the text beats
none of it — so taking the first answer dropped the rest: `αא` loaded the Greek face and
lost the Hebrew with a Hebrew font installed. Each round must cover at least one
previously-missing code point, faces are deduplicated, the chain is capped, and whatever
is still missing is recorded so a search that cannot succeed is not repeated on every
later label.

Two caches sit behind that chain, keyed differently on purpose because the costs are
different. Faces are found **per script** (`detectCjkLanguage`), so a diagram's eighteen
Chinese labels run one search rather than eighteen — measured at 17 ms against 293 ms,
since every search re-runs candidate enumeration and the coverage check. Faces are then
deduplicated **by the face that came back**, because two keys legitimately resolve to one
font: `detectCjkLanguage` answers `undefined` for text shaped identically in Simplified
and Traditional (`开始处理`) and `zh-Hans` for text that is not (`数据校验`), so one
diagram alternates between two keys and built `Songti SC` twice — two 43,000-entry
`cmap`s, and a glyph cache that missed on every label because it is keyed on outline
identity. Deduplicating on the answer rather than on the question fixes that class of
problem instead of guessing which keys ought to have been equal.

**The rasteriser yields pixels; `Draw.toPng` adds the file.** The encoder lives with
DEFLATE and CRC-32 in `archive/png.ts`, published as `encodePng` from
`documonster/archive` — a PNG is a DEFLATE stream plus CRC-32-checked chunks, so that is
where it belongs rather than in whichever module first wanted a picture. `draw/api/png.ts`
is the one file in `draw/` allowed to import `archive/` (registered in
`scripts/verify-layers.ts`), which is why `draw` now sits above `archive` in the diagram;
`archive` imports only `stream` and `utils`, so the edge cannot form a cycle. It used to be
two calls across two entries, and a Mermaid diagram took three — the encoder was placed by
where its code lived, and a user looking for a PNG looked in `Draw` and `Mermaid`.
`scripts/treeshake-verify.ts` pins that, under rolldown and rspack, only `Draw.toPng` /
`Mermaid.toPng` bring DEFLATE into a bundle: `Draw.toSvg` and `Draw.toRgba` stay archive-free.
esbuild keeps every member of an imported namespace, so there the edge costs every `Draw`
consumer the encoder; that scenario excludes esbuild, as the other member-level ones do. Moving DEFLATE down to
`utils` instead was rejected — it would put a compression engine, with its Node and browser
variants, at Layer 0.

That placement was arrived at the second time. The encoder sat in `excel/utils/png.ts`
while both its callers were in `excel/`, with a note to move it down rather than copy it
a third time. What triggered the move was not a second internal caller but a public
entry: a consumer of `documonster/draw` or `documonster/mermaid` could obtain pixels and
had no way to encode them, while a chart consumer could — and publishing the encoder from
`documonster/excel` to close that gap would have been the wrong answer to the right
question. `excel/chart/render/draw-raster-png.ts` remains as the internal eight-line
`DrawList` → PNG convenience the chart and sparkline renderers use.

`createPdfDrawSurface` takes a `PdfDrawPage` — a structural interface naming the six
marks it puts down — rather than a `PdfPageBuilder`. The concrete builder also carries
annotations, form fields and a content stream, and naming it in a public signature
would publish all of them.

Because the Y flip lives in exactly one adapter, so does the fact that a
reflection reverses a rotation's sense — a bug each per-backend renderer
previously had to discover for itself.

**Clipping** is axis-aligned rectangles only (`DrawNode.group.clip`), authored in
the group's own space so it moves with `transform`. Every backend expresses that
exactly — SVG `clipPath`, PDF `q … W n … Q`, a scissor test in `setPixel` — and an
arbitrary clip path in a scanline rasteriser is a different order of problem. A
surface may omit `pushClip`/`popClip`, in which case the walker draws unclipped
rather than dropping the content.

**Primitives** are rect, ellipse, **sector**, line, polyline, path (move/line/cubic)
and text. A sector is first-class rather than lowered to a path because the three
backends do it three genuinely different ways and lowering upstream would make two
of them worse: SVG has arcs natively, PDF needs cubics, and the rasteriser tests
radius-and-angle per pixel for an exact edge. The old pipeline smuggled the
parameters past the SVG string in a `data-sector` attribute so the rasteriser
could recover them; modelling them removes that channel.

**Text measurement belongs here too.** A producer sizes its boxes around the text they
contain, so it has to measure before it can build a display list at all. `draw/text.ts`
expresses `@utils/text-measure` (Layer 0 — the glyph-advance tables, moved down from
`@excel/utils` where nothing below Layer 4 could reach them) in terms of `DrawTextStyle`,
and adds `wrapText`. Before that, measurement was reachable only through
`@excel/chart/shared/chart-utils`, which is Layer 4: anything sitting beside `draw` could
draw text but not measure it, and therefore could not lay anything out. That was the one
thing stopping a non-chart producer from being written against this engine.

Note the unit: `measureTextWidthPx` returns CSS pixels for a point size, while a display
list draws text at `style.size` units. `measureText` applies `POINTS_PER_PIXEL`; skipping
it over-reports every label by 4/3, which is how legends came out too wide and centred
titles sat left of centre.

**Deliberately absent**: element-level opacity, gradients and patterns. Opacity is
expressed as an alpha on each paint instead. That is _not_ identical to the SVG
`opacity` attribute, and it is worth being precise about the difference rather than
claiming an equivalence: element opacity composites the finished shape once, so a
stroke that overlaps its own fill is flattened first and then faded, whereas
per-paint alpha fades the fill and then blends the stroke over the result. The two
agree everywhere except the band where a translucent stroke covers its own fill.
For the region map's circles — a 1.5-unit white ring at `0.92` over a near-white
panel — that band is under a pixel wide and the channels differ by at most
15/255; the fill is bit-identical. Closing that gap properly means offscreen
compositing in every backend (a transparency group and a Form XObject in PDF, a
second buffer in the rasteriser), which is a large amount of machinery for a
sub-pixel seam, so the alpha is where it stays. Note that
`describeSvgGeometry` folds element opacity into both paints, so it cannot see this
difference — do not read a matching geometry hash as proof that compositing agrees.
Nothing produces gradients or patterns —
the chart engine already degrades gradient fills to a representative colour before
it reaches any renderer — so adding them would grow the IR and every surface
without changing a single output. Filters are the one exception, and they are
named for the single backend that can express them (`DrawList.svgDefs` +
`group.svgFilterId`): a DrawingML shadow has no counterpart in a content stream or
a scanline rasteriser, so the field says SVG in its name rather than pretending to
be portable. That mirrors what the per-backend renderers already did.

**Metadata that only one backend can carry does not enter the IR.** The region map
reports which path drew it — real TopoJSON, the built-in centroid preview, or the
hex-tile fallback — and that is the only way a caller learns whether their topology
actually matched their categories. It is a fact about a decision, not about geometry
or paint, so the SVG renderer attaches it as a `data-region-map-mode` attribute on a
wrapping `<g>` when it serialises, and the PDF path simply draws the nodes. Nothing
was added to `DrawNode` for it. Prefer that shape for any future SVG-only
annotation: decide it in the producer, attach it at the serialisation boundary.

### The `mermaid` module

`mermaid` is the first producer written _against_ the drawing engine rather than
alongside it, and it exists as much to test the claim as to draw diagrams: it
implements no backend, and gets SVG, pixels and a PDF page anyway.

Twenty-one diagram types, but far fewer layouts, because most of them are the same picture
in a different language. A state diagram, a class diagram, an ER diagram, a requirement
diagram and a C4 diagram are all a directed graph of boxes, so each is converted to the
flowchart form and laid out by the same ranking, ordering, straightening and routing; only
the boxes and the marks at the ends of the edges differ, which is the part that genuinely
differs. A box that sizes itself — a class compartment stack, a C4 element with its
technology line — hands the result to the layout through `measureNode`.

The rest are _not_ graphs and are not forced through that layout. A Gantt bar's position
comes from the calendar, a journey's dot from a score, a mind map's node from the height of
its own subtree, a commit from its lane and its ordinal. Each of those is a layout and a
renderer in one pass, which is honest about how little there is to lay out. What they share
is the theme, the primitives and — for the Gantt chart — the date arithmetic in
`parse/dates`, which is kept separate precisely so it can be tested as arithmetic rather
than through a picture: a bar in the wrong place is not an ugly chart, it is a wrong one.

Three passes, and keeping them apart is what makes the module tractable. The
parser turns text into a syntax tree that says nothing about geometry. The layout
turns that tree into coordinates — a layered ("Sugiyama") arrangement: rank,
order to reduce crossings, position, route. The renderer turns coordinates into a
`DrawList`. A caller can stop after any of them; `parseMermaid` and
`layoutFlowchart` are both public because inspecting the tree or re-using the
positions are real things to want.

**Nothing was added to the IR for it.** Every outline is a rect, an ellipse, a
closed polyline or a path. The case worth naming is the arrowhead: a marker is an
SVG concept with no counterpart in a content stream or a scanline rasteriser, so
the producer lowers it to a small filled triangle aimed along the last segment —
which is the same picture in all three backends and needed no new primitive. The
text baseline is the other one: the engine positions text by its baseline because
that is what every backend can honour, so vertical centring is arithmetic the
producer does (`BASELINE_SHIFT`), not a portable attribute it can ask for.

Two layout details are worth knowing before changing them, because both were
bugs first:

- **A rank is re-centred after its overlaps are resolved.** Pushing left to right
  until every pair clears the gap is one-directional and drags the whole rank
  right; the trunk of a diagram then visibly bends away from a fork, because the
  fork's children have been shoved sideways and their parent follows them.
  Translating the finished rank back by its mean displacement cannot re-introduce
  an overlap, since it moves every node equally.
- **A long edge is threaded, not routed round.** An edge spanning several ranks gets a
  chain of narrow dummy nodes, one per rank it crosses, so it takes part in the ordering
  like any other edge and is drawn straight through the lane that was left for it. Without
  them it has no say in how the ranks between its ends are arranged and has to be routed
  afterwards around whatever ended up in the way — which is why a diagram with several
  skip edges used to grow a bundle of lanes down its margin. Dummies are scaffolding: they
  steer the layout and are dropped before it is returned, and they are deliberately kept
  out of the straightening pass, or a chain threading past a trunk drags it sideways.
- **The ordering measures its own output.** Median sweeps, then a transpose pass that
  swaps adjacent pairs while that removes crossings, then keep whichever iteration
  actually counted fewest. A heuristic that does not measure what it produced has no
  reason to improve it, and the last sweep is regularly worse than one already passed
  through.
- **Every edge gets its own place on the borders it joins.** Bending at the midpoint of
  the flow axis makes an incoming edge arrive at the _centre_ of its target's leading
  edge whatever it came from, so a node with three parents collected three arrowheads on
  one pixel — and in a class diagram the marks that tell inheritance from composition
  landed on top of each other. Ports are ordered by where the other end sits, so the
  lines arrive in the order they left and do not cross on the way in.
- **An edge that runs backwards or skips a rank is routed around the outside.**
  Drawn straight it crosses whatever sits between its ends, and the piece that
  shows is the segment between two _other_ nodes — a cycle's back edge grew an
  arrowhead on a pair that had not asked for one.

Cross-backend agreement is enforced by
`src/modules/pdf/__tests__/draw-backend-parity.test.ts`: the same display list is
rendered to markup, pixels and PDF operators, and the geometry is asserted to
match modulo each backend's coordinate convention.

- `pdf` imports `xml` because XMP metadata is XML: editing it with regular expressions wrote unbound prefixes, duplicated properties and edits inside nested descriptions, so the editor parses and rewrites the packet with the real parser.
- Office file encryption (MS-OFFCRYPTO Agile Encryption) is shared by `excel` and `word`, so it lives below both, in `archive` — CFB is a container format, as ZIP is: `archive/cfb/` (the compound file; `signature.ts` is the tiny test readers use before loading anything) and `archive/office-crypto/` (key derivation, verifier, package encryption, data-integrity HMAC). Both are internal; each module checks its own arguments, imports them with `import()` only when a password is involved, and reports `CfbError` / `OfficeCryptoError` as its own errors (`Docx*`, `Excel*`).
- Modules may only import from **lower** layers — never sideways or upward.
- **Sole exceptions**:
  - `pdf/excel-bridge.ts` may import from `@excel/`. No other file in `pdf/` may import `@excel/` except `pdf/word-chart-bridge.ts` (Word charts rendered by the Excel chart engine).
  - `pdf/word-bridge.ts`, `pdf/word-chart-bridge.ts`, and `pdf/word-layout-to-pdf.ts` may import from `@word/` (the Word→PDF bridge family). No other file in `pdf/` may.
  - `word/bridge/excel-bridge.ts` may import from `@excel/`. No other file in `word/` may.
  - `draw/api/png.ts` may import from `@archive/` (the PNG encoder, for `Draw.toPng`). No other file in `draw/` may.
  - `formula/` consumes immutable snapshots and emits writeback plans; `excel/` owns the host adapter. `formula/` never imports concrete types from `@excel/*`.
- `utils/` must never import from any module.

These rules are **machine-enforced** by `scripts/verify-layers.ts` (run via `pnpm verify:layers`, included in `pnpm check`). It scans every production `.ts` import and fails on any forbidden cross-module import. A new bridge file that legitimately needs a cross-module import must be registered in that script's `EXCEPTIONS` map and documented above.

## Path Aliases

`@excel/*`, `@word/*`, `@formula/*`, `@pdf/*`, `@csv/*`, `@markdown/*`, `@xml/*`, `@archive/*`, `@stream/*`, `@draw/*` → `./src/modules/<name>/*`
`@utils/*` → `./src/utils/*` | `@test/*` → `./src/test/*`

Use aliases for **all** module imports — both cross-module (`@archive/...` from excel) and same-module (`@excel/cell` from within excel). This matches the IDE auto-import setting (`importModuleSpecifier: "non-relative"`) and keeps imports stable when files move. The only exception is `src/utils/` (Layer 0), whose internal files use relative paths (`./errors`, `./glob`).

## Platform Variants (`*.browser.ts`)

A module gets a `*.browser.ts` sibling when its Node form cannot work in a browser, or
must not ship there. **Write the Node name at the import site and nothing else** —
`import { readFileIfExists } from "@utils/fs"`. Selecting the variant is not your job and
not the compiler's; it is resolution.

`scripts/link-platform-variants.ts` runs after `fix-esm-imports.ts` and rewrites every
emitted import of a variant-having module to `#platform/<path>`. One wildcard entry in the
manifest expresses the choice:

```json
"imports": {
  "#platform/*": {
    "browser": { "types": "./dist/types/*.browser.d.ts", "default": "./dist/esm/*.browser.js" },
    "default": { "types": "./dist/types/*.d.ts",         "default": "./dist/esm/*.js" }
  }
}
```

So **adding a variant needs no manifest edit, no build-target edit and no source edit
beyond the file itself.** The wildcard is deliberate: an explicit list of 27 entries is a
list that drifts from the tree it describes.

**This replaced a second build of the whole tree.** `tsconfig.browser.json` used to compile
`src` again into `dist/browser`, and `fix-browser-imports.ts` walked that copy rewriting
relative specifiers to prefer the `.browser` sibling — its own log line was
`modified 112 files; rewrote 158 specifiers`. Comparing the trees file by file, 705 of 784
`.js` were byte-identical to `dist/esm` and 751 of 784 `.d.ts` to `dist/types`, and nothing
existed in `dist/browser` that was not already in `dist/esm`: `tsc` compiles the
`.browser.ts` files into both. It was 12.6 MB of published artifact to change 158 strings.
The files that genuinely differed were mostly not platform code but _importers_ whose
specifier had been rewritten — and an importer of `#platform/utils/fs` is byte-identical in
both worlds, which is why the difference collapses to the 27 variant files.

Three invariants keep it honest, and each fails loudly rather than degrade:

- `pnpm verify:browser-stubs` scans `dist/esm` and `dist/types` for a _relative_ import of
  a module that has a variant. Any hit means the browser condition cannot fire for it, so a
  browser consumer silently receives the Node module. Without the linker it reports all 158.
  It also holds a short policy list of modules that must keep a variant, because a deleted
  variant leaves nothing for a structural check to call a violation.
- `pnpm build:verify:browser` type-checks the browser declarations with
  `--customConditions browser --lib ESNext,DOM` and no Node types available. Drop the
  condition and it fails on `Buffer` and `node:stream`, which is how you know the check is
  doing work rather than passing vacuously.
- `pnpm verify:public-types` reads the browser entries the way `#platform/*` resolves them
  — a module's `.browser.ts` sibling wins — so a shared surface file needs no hand-copied
  `.browser` twin to type-check the browser build.

**One capability was deliberately given up:** `dist/esm` is no longer loadable in a browser
straight from a URL, because `#platform/*` is a bare specifier and a browser needs an import
map to resolve one. Bundlers resolve it (verified against esbuild, rolldown and rspack), and
the no-bundler case is what `dist/iife` is for. Deep paths were never part of the contract —
`exports` has only ever published the 19 subpaths.

**A second consumer is affected, and it is worth naming rather than discovering:** Jest
cannot resolve `imports` (`#`) specifiers under its experimental ESM mode. Measured on
jest 30.4.2 / Node 26.8.1, `await import("documonster/excel")` with
`--experimental-vm-modules` fails with `does not provide an export named …`, and a minimal
package with a single non-wildcard `#plat` entry fails identically — so this is the
`imports` field in general, not the wildcard and not this mapping. 14 of the 19 subpaths
reach a platform variant and are therefore affected; the other five (`markdown`, `mermaid`,
`formula`, `xml`, `word/markdown`) are not.

This is moot for the route a Jest user actually takes, because that route no longer goes
through Jest's ESM loader at all: the package is ESM-only, so a Jest consumer transpiles it
with `babel-jest` (see "CommonJS"), and Babel lowers `import "#platform/…"` to
`require("#platform/…")`, which Node's CommonJS resolver handles. The `imports` field is only
unresolvable in Jest's _own_ ESM implementation.

The alternative was to express the internal seam as a public `./internal/*` subpath in
`exports`, which Jest does resolve natively. It was rejected: it would publish 27 internal
modules as importable paths to work around one runner's experimental mode, and this
package's position is that the `dist/` layout is not a contract.

Note that TypeScript does **not** apply the `browser` condition by default under either
`bundler` or `nodenext` resolution: a consumer who wants the browser surface's types sets
`customConditions: ["browser"]`. That was equally true of the two-tree layout, so nothing
changed for them.

`src/utils/browser.ts` (`preferBrowserFilesPlugin`) is a separate mechanism for a separate
job: it swaps variants while bundling **from source**, which is what `rolldown.config.ts`
and `vitest.browser.config.ts` do. Source has no `#platform/*` specifiers to resolve.

## CommonJS

**The package is ESM-only.** Every `exports` subpath is `{ browser, types, default }` and
`default` names a file in `dist/esm`; there is no `require` condition and no transpiled tree.
A CommonJS consumer reaches it through `require(esm)`, and their call site does not change:

```js
const { Workbook } = require("documonster/excel"); // works verbatim
```

`dist/cjs` used to be a full second transpile — 10.7 MB unpacked, 2.4 MB of the packed
tarball, a third of everything an `npm install` downloaded — and it existed for a runtime
capability Node has since absorbed.

### Why `>=22.13.0` and not `>=22.12.0`

`require(esm)` was unflagged in 22.12, so 22.12 is where it starts _working_. But measured on
a real 22.12.0 binary, every `require()` of an ES module also prints:

```
ExperimentalWarning: CommonJS module … is loading ES Module … using require().
Support for loading ES Module in require() is an experimental feature and might change at any time
```

That is stderr noise on every process start, in every consumer's CI log. Verified across
22.12.0 / 22.13.0 / 22.22.1 / 24.16.0 / 26.8.1: the warning is present on 22.12.0 and gone
from 22.13.0 onward, while `require()` of all 19 subpaths succeeds on every one of them. So
the floor is the first version where the feature is both available and quiet.

### What this costs a consumer, and what it does not

Nothing changes for Node ESM, for a bundler, or for a `<script>` tag. Two populations are
affected:

- **Node 22.0 – 22.12.** 22.0–22.11 need `--experimental-require-module`; 22.12 works but
  warns. Upgrade.
- **TypeScript consumers on `moduleResolution: node16`.** That mode predates `require(esm)`
  and reports `TS1479` before Node is involved; `nodenext` accepts it. Measured on tsc 5.9:
  `nodenext`, `bundler` and the legacy `commonjs` + `node10` pair all compile, `node16` does
  not, and the emitted JavaScript runs in every case. It is a one-word change in their
  tsconfig, and `nodenext` is what TypeScript recommends for Node projects anyway.
- **Jest users.** Jest cannot `require()` an ES module at all — measured on jest 30.4.2 with
  transforms disabled, on Node 26.8.1, and against a local `.mjs` as a control, so this is
  Jest and not this package. Its own error text points at "Node v24.9+ where Jest supports
  require(esm) natively"; that did not hold. The fix is two files, verified working against
  the published shape:

  ```js
  // babel.config.cjs
  module.exports = { presets: [["@babel/preset-env", { targets: { node: "current" } }]] };
  ```

  ```json
  // jest.config.json
  {
    "transform": { "\\.[jt]sx?$": "babel-jest" },
    "transformIgnorePatterns": ["/node_modules/(?!documonster)"]
  }
  ```

That second cost is real but it is not unusual, and that is the calibration that decided
this. Requiring six of the eight most-depended-on packages tested — `strip-ansi` (527 M
weekly), `chalk` (507 M), `p-limit` (333 M), `uuid` (295 M), `nanoid` (243 M), `node-fetch`
(194 M) — fails in Jest with the identical message; only `axios` and `zod` still ship
CommonJS. Any Jest project whose dependency tree reaches one of those already carries that
`transformIgnorePatterns` line, so adding a name to it is a one-line change rather than a
new configuration. And all six switched at a _major_ version; this package is pre-1.0 with
`bump-minor-pre-major`, so now is the cheapest moment it will ever have.

### The standing rule this creates

`require(esm)` throws `ERR_REQUIRE_ASYNC_MODULE` on a module graph containing top-level
await. With no transpiled tree to fall back on, **top-level await would make the package
unrequirable from CommonJS.** So: do not introduce one. A zero-dependency document toolkit
has no need for it.

The whole matrix was measured against a real `npm install` of the packed tarball, not a
symlinked `node_modules`: all 19 subpaths `require()`d, destructured and executed on
22.12.0 / 22.13.0 / 22.22.1 / 24.16.0 / 26.8.1 and on Bun 1.4, plus ESM `import` on each.
`pnpm verify:install` is that check, and the `Consumer floor` CI job runs it on 22.13.0.

**Two floors, and conflating them breaks CI.** `engines.node` (>=22.13) is what a _consumer_
needs; the _toolchain_ needs >=22.18, because every gate is a `node scripts/*.ts` and Node's
type stripping is unflagged only from 22.18 — 22.17.1 still throws
`ERR_UNKNOWN_FILE_EXTENSION`. Pinning the consumer floor in the `test` matrix therefore failed
all nine jobs for a reason unrelated to the package. So the `test` matrix runs `22.x`, and the
floor is verified where it can be: against the installed artifact, by the one script here
written in plain CommonJS (`scripts/verify-installed-package.cjs`) so that it runs there at all.

**Support policy: Node LTS lines and Current, while each is still alive.** The floor is the
oldest non-EOL LTS line, so it moves on Node's schedule rather than ours — as of Node 22 (Jod,
maintenance until 2027-04-30) and 24 (Krypton, active until 2028-04-30), with 26 as Current and
becoming LTS on 2026-10-28. That is what the `test` matrix `22.x / 24.x / 26.x` covers. Odd
lines are absent because 23 and 25 are **EOL**, not because they are odd: Current rotates every
six months and is odd-numbered half the time, so when 27 ships (2027-04-22) it becomes Current
and joins the matrix while 26 will already be LTS. `engines.node` still carries a **floor, not
an allowlist** — an upper bound or a line-by-line range would reject a future release and trip
`engine-strict` installs for no benefit — so an unsupported line is untested rather than
blocked. When 22 reaches EOL, raise the floor to the next living LTS and drop `22.x` from the
matrix in the same commit; do not leave a comment claiming support for a line nothing runs.

**Bun is supported on `latest` only, and that is why there is no `engines.bun`.** Bun has no
LTS line and no maintenance branches, so there is nothing to express a range against: a floor
like `>=1` would claim every 1.x works, which is a stronger promise than anything here tests.
"Latest" is not a semver range, so the policy cannot live in `engines` at all — it lives in the
one place it can be enforced, the `bun-version: [latest]` CI matrix.

One methodological trap is worth recording, because it produced a wrong "verified" reading
first time round: the `ExperimentalWarning` goes through `process.emitWarning`, which defers
to the next tick, so a harness ending in `process.exit()` never sees it and 22.12 looks
clean. `verify:cjs` waits for the child to exit for that reason.

That is a gate, not a hope. `pnpm verify:cjs` and
`src/test/__tests__/cjs-entrypoints.node.test.ts` `require()` all 19 entry points, and the
test first pins the mechanism against a fixture — one case proves `require(esm)` works on a
plain module, another proves it throws on top-level await — so the suite cannot pass because
Node stopped caring. Verified by appending an `await` to `markdown/index.ts`:
`✗ verify:cjs — 1 of 19 entry point(s) failed to require()`.

Deleting the CommonJS transpile also removed a whole class of bug with it. `tsc`'s CommonJS
emit preserved statement order where ESM hoists its bindings, so an import written at the end
of a file landed below its consumer and threw
`ReferenceError: Cannot access 'grapheme_1' before initialization` at load. That shipped once,
breaking 12 of the 19 CJS entries while the entire test suite stayed green — because the tests
run against source, through ESM. There is now one emit, so source and artifact cannot disagree
that way.

## Documentation

Examples are the first thing a consumer copies, and until recently they were the only artefact
here that nothing in the toolchain read — so they drifted, silently and with the authority of
sitting next to the implementation. `scripts/verify-doc-examples.ts` (`pnpm verify:doc-examples`,
included in `pnpm check`) closes that: it resolves every documented import **and** every
`Namespace.member` reference against the built public surface by handing them to `tsc`.

What it found on its first two runs is the reason it exists: seventeen `@example` blocks in
`pdf/` and `word/` still naming the flat exports those modules replaced with namespace surfaces
(`readPdf` → `Pdf.read`, `toBuffer` → `Io.toBuffer`, …); eleven more of the same in the READMEs,
which carry five times as many imports; a `Document.addBodyContent` that has never existed; and
a comment instructing readers to call four members on `StyleMap`, which resolves — to a _type_
of that name, while the namespace is published as `Styles`.

Rules when writing documentation:

- **Name the public form**, not the internal function it forwards to — and not a deprecated
  forward either. `Pdf.extract`, not `readPdf` (nor `Pdf.read`, which now forwards to it);
  `Worksheet.Handle`, not `WorksheetData`.
- **Put runnable code in a fence.** Markdown is checked inside fenced blocks only; inline code
  in a sentence is treated as a reference to a symbol, which is what lets a document _discuss_
  a broken import. TSDoc comments are read in full, prose included, because their prose
  instructs.
- **Use braced imports** (`import { A, B } from "documonster/x"`). A namespace or default import
  is reported rather than skipped — a gate with a silent hole is worse than no gate.
- **Keep local links real.** `scripts/verify-doc-links.ts` checks relative files, directories,
  images and Markdown heading fragments. External URLs are deliberately not fetched. Heading
  anchors follow GitHub's rule, including the two details that surprise people: a dropped
  character between spaces leaves _two_ hyphens (`Excel — XLSX` → `excel--xlsx`), and a
  leading emoji leaves a leading hyphen (`⚠ BREAKING` → `-breaking`).
- **Do not restate counts.** A number only a human keeps in sync is eventually wrong: the
  function count said 433 in sixteen places while the registry held 448.
  `src/modules/formula/__tests__/function-count.node.test.ts` pins the twelve that remain against
  `listFunctionNames()`. It carries the `.node` suffix because it reads Markdown off disk: the
  browser config excludes `*.node.test.ts` by glob, which is a rule rather than a hand-kept list.
- A member a comment names _in order to say it is absent_ goes in that script's
  `DELIBERATELY_ABSENT` map, keyed by file, and is documented there.

## Public API Conventions

Every public entry follows one vocabulary, so a name learnt in one module predicts the
name in the next. A new export that breaks a rule here needs a reason written beside it.

- **One way to do a thing.** A new name must add a capability or replace a deprecated one;
  it is not added for symmetry. Legacy return conventions (e.g. `null` for absence) are
  kept rather than duplicated. So a class gets no namespace factory that only forwards to
  its constructor or a static opener, a type gets no second name (`Handle`, or a flat
  re-export of a namespace member), and a lookup gets no twin that differs only in
  returning `undefined` where the existing one returns `null`. The rules below describe
  the vocabulary for _new_ capability; they are not a reason to rename what already works.

- **Namespaces of plain functions.** Each subpath publishes its operations as namespaces
  (`Workbook.*`, `Csv.*`, `Mermaid.*`). Flat exports are for types, error classes and
  constants only. The exception is `documonster/stream`'s Node-compatible surface
  (`Readable`, `pipeline`, `isDestroyed`, …), whose names must match `node:stream`;
  everything else in that entry lives in `Streams`. A subpath that bridges two modules
  prefixes its namespace with the module it extends (`ExcelCsv`, `WordHtml`), so it can
  be imported beside the standalone one (`Csv`, `Markdown`) without renaming.
- **A namespace is a subject, not a promise about effects.** It may hold more than one
  collection (`Image`: the workbook's media and a sheet's placements of it), and a plain
  function in it is not necessarily pure — `ChartClassic.createOptionsFromRows` writes its
  rows to the sheet it is given, `fillCaches` edits the model. Say so in the member's TSDoc
  and in the namespace header when the subject holds more than one collection. Likewise a
  number keeps the unit of the format it belongs to (Word `size` in half-points, Draw
  `size` in user units, PDF `fontSize` in points); a shared field name across modules does
  not imply a shared unit, so each field documents its unit and conversions go through
  `Units`.
- **Low-level members go in a labelled sub-namespace**, not beside the everyday ones:
  `Draw.Backend`, `Draw.Geometry`, `Streams.LowLevel`, `WordCrypto.Agile`,
  `WordCrypto.Cfb`, `Pdf.Signature`. They are supported API; the label says who they are
  for.
- **Handles.** A namespace whose functions take a handle that has no other public name
  exports it as `type Handle` (`Table.Handle`). Where the handle already has a public name
  — a data-model interface such as `DocxDocument`, `DocxStyleMap` or `ChartHandle` — that name
  is the only one; no `Handle` alias is added beside it.
- **One primary object per document type, and one namespace for it.** `create`, `read`
  and `toBuffer` live together (`Workbook.*`, `Document.*`) and return or take the same thing,
  every namespace operates on it, `toBuffer` writes it — `Workbook.Handle` for Excel,
  `DocxDocument` for Word. Word's is a transparent content tree rather than an opaque
  record because a Word document _is_ a tree callers assemble; the shape of the API is
  the same. A plain serialisable snapshot is `<Noun>Model` (`WorkbookModel`).
- **Prefixes.** A format prefix (`Xlsx`, `Docx`) is for format-level things — read/write
  options, parse errors (`XlsxReadOptions`, `DocxParseError`) — not for the document
  object. An entry whose vocabulary is generic (`Row`, `Table`, `Transform…`) prefixes
  _every_ flat type with its module (`Pdf…`, `Csv…`), not only the ones that collide
  today: PDF's `Table`, `Row`, `Image`, `Sheet` would otherwise meet Excel's
  namespaces, and CSV's `Row` already did. A bridge entry prefixes its flat types with
  its namespace, as it prefixes the namespace itself (`ExcelCsvReadOptions`,
  `ExcelMarkdownWriteOptions`), so they cannot be mistaken for the standalone module's
  (`CsvParseOptions`). A format-level type that serves every format of the document is
  named for the document (`WorkbookEncryptOptions` covers XLSX and XLSB); `Xlsx…` is for
  what only XLSX has.
  The one exception is Word's document model — the content tree declared in
  `word/types.ts` (`Paragraph`, `Run`, `Table`, `RunProperties`, …), which callers build
  documents from and which mirrors OOXML element for element; everything else Word
  publishes (option, result, layout and conversion types) carries `Docx`. That includes
  Word's bridge entries, whose types take `Docx` plus the bridge's subject
  (`DocxHtmlRenderOptions`, `DocxMarkdownImportOptions`, `DocxChartSvgOptions`) rather than
  the namespace (`WordHtml…`): `Word…` is the prefix Word's flat types moved away from
  (`WordChartExOptions`, published in 0.21.1, is now a deprecated name for
  `DocxChartExOptions`), and the subject already tells them from the standalone module's
  (`MarkdownParseOptions`).
- **Every public error class carries a prefix**: its module's (`ExcelTableError`,
  `DocxTemplateError`, `ArchiveFileTooLargeError`, `PdfFontError`) or, for a failure
  only one format can produce, that format's (`ZipEocdNotFoundError`, `XlsxParseError`).
  The one exception is `AbortError`, whose name is the DOM's and is what
  `isAbortError` and every `AbortSignal` consumer test for. Each module's argument
  error is `<Module>ArgumentError` (`StreamArgumentError`, not `StreamTypeError`). A failure
  every module can have has the same name in each: `<Module>FileError` for a refused file
  operation, `<Module>UnsupportedFeatureError` for something the module recognises but does
  not do, `<Module>LimitExceededError` for a configured read limit (`ExcelLimitExceededError`,
  `DocxLimitExceededError`, `PdfLimitExceededError`). An operation's own computational cap
  that is not such a limit (the size of `Diff.compare`'s LCS table) is the module's base
  error with the call in its message. An index or other argument the caller got wrong is
  the `ArgumentError`, never a `*StructureError` (which describes a malformed file) or a
  bare `RangeError`; "it always threw that class" is not a reason to keep one.
- **A namespace member does not repeat its namespace's noun**, and a getter/setter pair
  shares one: `Chart.isEx` (the namespace is the chart) but `Chartsheet.isChartEx` (the
  sheet is not the chart); `Security.DEFAULT_POLICY`, not `DEFAULT_SECURITY_POLICY`;
  `Table.hasRowStripes` / `Table.setRowStripes`, not `setShowRowStripes`.
- **Constants**: an enum-like name → code table is PascalCase (`ValueType`, `PaperSize`,
  `Pdf.PageSizes`, `Archive.TarType`); any other constant is SCREAMING_CASE
  (`TAR_BLOCK_SIZE`, `Xml.DEFAULT_DECLARATION`, `Styles.DEFAULT`).
- **A namespace holds its module's own things.** Generic helpers the platform already
  provides (byte/base64/UTF-8 conversion, `signal.throwIfAborted()`) are not
  re-published from whichever module first used them; that is how `Archive` came to carry
  `base64ToUint8Array` and `concatUint8Arrays`, now deprecated in favour of the native
  APIs. An error guard sits flat beside its class (`isAbortError` beside `AbortError`).
  A helper that only _looks_ like a platform one stays when it does something the platform
  does not, and says so beside it: `Archive.throwIfAborted` always throws this package's
  `AbortError` class (the signal's reason becomes its `cause`), where
  `signal.throwIfAborted()` rethrows whatever reason was given, so `isAbortError` and
  `instanceof AbortError` cannot rely on the native one.
- **Classes only for stateful incremental objects** — the PDF `Builder`/`Editor`, the
  streaming readers and writers, archive builders and readers, XML writers. They are
  public values, built with their constructor or static opener (`new Pdf.Builder()`,
  `Pdf.Editor.read(bytes)`, `new Xml.Writer()`, `Archive.RemoteZipReader.open(url)`); a
  namespace factory exists only where it adds something the class cannot — the `Archive`
  format matrix (`Archive.createZip`, `Archive.openTarGz`, …) picks the class from the
  format, and replaced the deprecated `zip` / `unzip`; `Streams.createReadable` and its
  siblings return one cross-platform typed interface over two platform classes. A factory
  that only forwards to `new X(options)` is deprecated in favour of the constructor
  (`Workbook.createStreamWriter`, `Csv.createParserStream`, `Streams.createPullStream`).
  Their methods follow the same verbs (`toBuffer`, not `save`/`build`/`bytes`).
  Two refinements apply to class methods only. A builder whose `add*` / `remove*` methods
  return `this` for chaining (`Pdf.Builder`, `Stream.DocumentWriter`) is exempt from the
  collection return rules below, because what it returns is the builder, not an item. And a
  transactional object marks, keeps and undoes with `snapshot()` / `commit()` /
  `rollback()` (`Streams.ChunkedBuilder`, `Xml.Writer`) — a transaction's `commit` keeps
  the edits since the snapshot, which is a different object from a streaming writer's
  `commit` below and never appears on the same class.
  Two vocabularies are borrowed from the platform on purpose. An object edited by entry
  path (`Archive.ZipEditor`, `ArchiveFile`, and `ZipEditPlan`, which records the same edits)
  speaks `Map`'s `set` / `delete` (and `has` where it can answer) — `set` adds or replaces,
  which no collection verb above says, and `remove` beside `set` would mix two vocabularies.
  And something that holds a body names its readers as `Blob` and `Response` do — `bytes()`,
  `text()`, `stream()` — as `Streams.bytes` / `Streams.text` do. `stream()` here yields an
  async iterable of chunks, not the web `ReadableStream` `Blob#stream()` returns nor the Node
  `Readable` that `toStream` names; the name is for the role, and its TSDoc says the type.
- **A document module's streaming API is its `Stream` namespace**, and the classes in it
  are named for the primary object they write or read: `Stream.WorkbookWriter` in Excel,
  `Stream.DocumentWriter` in Word. `documonster/stream` calls its helper namespace
  `Streams` because `Stream` is a `node:stream` name that entry must keep free.
- **A streaming writer finishes the same way in every module.** `commit()` completes the
  file and ends the output the writer was given (a sink, a stream, a file); `toBuffer()`
  is for a writer that owns its output, commits if it has not, and returns the whole file —
  given an output of its own it refuses, rather than answering with something that is not
  the file. Word's `finalize()`, which returned the bytes in one mode and an empty array in
  the other, is the reason this is written down.
- **Who owns a returned object is part of the name's contract, and the same name means the
  same ownership everywhere.** The two document models answer this differently, because
  they are different kinds of object.

  Excel's model is opaque and addressed by position, so a reader returns a copy: editing it
  changes nothing, and the setter is how to change the document (`Cell.getFont` answers as
  `Row.getFont` and `Column.getFont` do, whether or not the cell exists). The live
  exceptions are named and few: `getStyle` (for in-place editing), `Chart.getModel` (edited
  through `Chart.mutate`), and **handles** — an item a collection's `remove` takes by identity
  (`Image.list`, `Pivot.list`, `ConditionalFormatting.list`, `Sparkline.list`) is the item
  itself, because a copy could not be handed back. A collection keyed by something else
  (`DataValidation` by reference) hands out copies. A value given to a setter is copied in as
  well, so the caller's object never becomes the cell's. Values are copied at the public
  boundary (`cellReadValue`), not in the stored value classes: the library's own reads — the
  writers, the formula engine — take the stored value (`cellGetValue`) and pay nothing.

  Word's model is a transparent tree whose nodes are typed `readonly` all the way down, so
  a reader (`Document.getContent`, `Query.listParagraphs`, `Query.getComment`, …) returns
  the tree's own nodes: the type system already refuses an edit through them, copying a
  subtree per lookup would cost what the tree is meant to save, and identity is what lets a
  node found by `Query` be handed back to `Document` or a `Transform`. The document changes
  only through `Document.*` (in place) or a `Transform` (a new document that may share
  unchanged nodes). Bytes held in a node — `ImageDef.data`, what `Vba.getData` returns —
  are the node's own too: a `Uint8Array` has no `readonly` form, and writing into one is
  outside the contract exactly as editing any other node is.

- **`clone` returns an independent copy and attaches it to nothing** (`Document.clone`,
  `Anchor.clone`). Copying something into a container is `copyTo(item, container, …)`,
  which is why `Chart.clone` — which placed a second chart on the sheet and returned its
  part number — is deprecated.
- **A file operation the platform refuses is still the module's error.** A file-path read or
  write wraps the file system's failure in the module's file error, with the system error
  as `cause` (`ExcelFileError`, `DocxFileError`), so `isExcelError` / `isDocxError` see
  every failure of the call. `Fs.*` in `documonster/archive` are file-system helpers
  themselves and pass the platform's error on.
- **An entry does not publish what cannot work on its platform.** A Node-only member is
  absent from the browser surface (a `*.browser.ts` variant of the namespace), as
  `Workbook.readFile` and `Document.readFile` are — a type error at the call beats a runtime one.
  The one sanctioned exception is a member isomorphic code must be able to name on both
  platforms, and then the entry carries a probe for it: `Compression.Workers` stays on the
  Node entry because `Compression.hasWorkerSupport()` answers `false` there, and every
  member it guards throws a typed, prefixed error (`ArchiveWorkerUnsupportedError`), never
  the bare base class. A Node-only _option_ that is harmless elsewhere
  (`preferSystemFonts`) stays and says in its documentation that it has no effect in a
  browser.
- **Positional arguments are what the operation acts on; settings go in options.** A row
  index or a count that names the target (`Table.addRow(t, values, rowNumber)`,
  `Table.removeColumns(t, colIndex, count)`) is positional; a flag or a preference is never.
  Two kinds of parameter list stay positional, and say so beside them: one copied from a
  format the caller translates from, in its order (`Draw.Geometry.arcToCubics` is SVG's `A`
  command, `Matrix.rotate(deg, cx, cy)` its `rotate()`, `createRectNode`'s `rx` a `<rect>`'s),
  and a value that is one entry of an options map handed through as it is
  (`Csv.applyDynamicTyping(value, columnConfig)` takes one column's `dynamicTyping`).
- **Accessors.** Readers are `getX`, writers `setX`, booleans `isX`/`hasX`. A bare noun
  (`Table.name(t)`) is not an accessor name. A reader never creates, materialises or
  mutates anything — reading a row's height must not add the row. The one sanctioned
  exception is `getStyle` on `Cell`/`Row`/`Column`: it returns the **live, mutable** style
  record so callers can edit it in place, and therefore materialises the cell, row or
  column it names. That is its purpose, not a leak. Side-effect-free reads are the facet
  getters (`Cell.getFont`, `Row.getHeight`, …), `findStyle` on `Cell`/`Row`/`Column` (a
  copy of the whole record, or `undefined`) and `Cell.getView`; do not add a second
  live-record reader. The deprecated get-or-create forwards (`Worksheet.getRow`,
  `Worksheet.getRows`, `Row.getCell`) keep creating what they name, as they always did. So
  do the streaming handle getters (`Stream.getRowCell`, `WorksheetWriter#getRow` /
  `getCell`): handing out the handle every `Stream.setCell*` writes through is what they
  are for. A streamed row is being written, not read, so there is little to look up in it;
  the writer still offers `findRow` / `findCell` for a lookup that creates nothing.
- **Collections.** `add(…)` attaches and returns what it attached, `get(…, key)` returns
  one item or `undefined`, `list(…)` returns all, `remove(…)` removes one and returns
  `boolean`, `clear(…)` removes all, `each(…, cb)` iterates. `find(…)` is `get`'s form
  for a lookup by something other than the collection's own key (`DataValidation.find` by
  address); it too creates nothing and returns `undefined`. Excel's deprecated chart `add*`
  return the part number and keep doing so until they are removed in 1.0; `Chart.place` and
  `Chart.copyTo` return the chart's handle, which carries the part number. Likewise an
  existing `remove` that returned the removed item keeps doing so (`Chart.removeSeries`).

  Those rules are for a collection edited in place. Three other kinds of operation share
  the verbs, and each has its own return because each has a different thing to hand back:
  a **transform** returns the new document, since the input is left as it was and nothing
  else changed (`Vba.add`, `Vba.remove`, `Font.addEmbedded`, `Security.protect`); a
  **registration** returns the identifier the registered thing is later referred to by
  (`Image.add` → the id `Image.place` takes, `Document.addFootnote` → the note id,
  `Document.addListDefinition` → the `numId`); and a **builder command** returns `this`
  (see the class rules above). An append into Word's tree (`Document.addContent`,
  `Document.addParagraph`) returns nothing: the node is the caller's, or is reachable with
  `Document.getContent`. Choose the kind before the verb, and name an operation that fits
  none of them for what it does.

- **Absence is `undefined`**, in every public return. The archive readers, `Fs.safeStats`
  and `Watermark.get` returned `null` until they were changed together; a deprecated name
  keeps the `null` it returned until it is removed. Two kinds of `null` are not absences and
  stay: Node's own semantics in `documonster/stream` (`Readable#read()`, `errored`), and a
  three-way result whose `null` means "not applicable" (`RemoteZipReader#checkPassword` and `#checkEntryPassword`: the entry is not encrypted).
  `null` as data — an empty cell, a JSON value, a callback's "drop this row" — is not a
  return convention at all.
- **Conversions and I/O.** `toX` / `fromX` for format conversions (`toSvg`, `fromDocx`);
  `read` / `readFile` / `readStream` in, `toBuffer` / `toStream` / `writeFile` /
  `writeStream` out. `format` serialises to a `string`; `toBuffer` to bytes and is async,
  with `toBufferSync` beside it when the work is synchronous. `extract` pulls content out
  of a document without producing the primary object (`Pdf.extract`,
  `WordExcel.extractTables`). The unsuffixed name is the primary form and a sibling is
  suffixed by how it differs: work that is I/O or compression is async with a `Sync`
  sibling (`packZip` / `packZipSync`, `gzip` / `gzipSync`), while parsing or formatting
  text already in memory is synchronous, with an `Async` sibling for input that arrives
  asynchronously (`Csv.parse` / `Csv.parseAsync`). `Async` marks the asynchronous sibling
  of a synchronous form whether that form returns or yields (`Csv.generateRows` /
  `Csv.generateRowsAsync`); `Stream` marks a form that consumes its source incrementally,
  holding one piece at a time (`Workbook.readStream`, `Archive.parseTarStream`,
  `Csv.parseStream`, `Xml.parseSaxStream`). `toBuffer` resolves to a `Uint8Array` on every
  platform, and to a plain one at runtime — not a `Buffer` typed as one, whose `slice` is a view
  where the declared type's copies. One verb means one
  thing: in `Archive`, `create*` returns a builder, `open*` a reader, `pack*` bytes,
  `unpack*` a path map and `parse*` an entry list — the last for TAR and tar.gz only, since a
  ZIP's entry list is what its reader already gives (`openZip(…).list()` / `entries()`) and a
  name is not added for symmetry. `toSvg` returns a complete `<svg>` document; a surface that
  holds only a fragment serialises it with `format()` (`Draw.Backend.SvgSurface#format`). A
  function that returns a changed copy and leaves its input alone is `with*`
  (`Draw.withAlpha`, `DrawingMl.withSpPrFill`), never `set*`, which promises an edit in place.
- **Bridges pick their verbs by what they are.** A bridge that is a file format of the
  module's primary object uses the I/O verbs (`ExcelCsv.read` / `format` / `toBuffer` /
  `writeFile`, `ExcelMarkdown.read`), because a CSV or Markdown table _is_ a way to store a
  workbook. A bridge that converts between two document models uses `toX` / `fromX`
  (`WordHtml.toDocx`, `WordMarkdown.fromDocx`, `WordExcel.toDocx`, `Pdf.fromDocx`), because
  the conversion is lossy and the result is a different kind of document.
- **Arguments.** The handle comes first; optional settings go in one trailing options
  object — no positional booleans or chains of optional positionals. An overload that
  accepts both a new options object and a deprecated positional form dispatches with
  `isRecordArgument` (`@utils/object`), never `typeof x === "object"`, `instanceof` or
  `Array.isArray`, so `null`, typed arrays and `Date` keep their old positional meaning;
  pin it with a parity test whose expected values come from the previous release.
- **Arguments are declared with `@utils/schema` and checked at the boundary.** A public
  function checks its arguments before it changes anything, with
  `check(schema, value, { where, field, error })` — or `checkAt(...)` with the same four
  values where a check runs per node or per call. `where` is the public call as the caller
  writes it (`Csv.parse`, `new Xml.SaxParser`), `field` the parameter name in the public
  signature, and `error` the module's own `<Module>ArgumentError`; a module whose error
  records more than field and message builds it from the factory's third argument,
  `{ where, label, expected, got }`. The schema derives the path (`options.columns[2].width`,
  `options.headers["a"]`) and the
  message `${where}: ${path} must be ${expected}, got ${describeValue(value)}`; neither is
  written by hand. A rule that relates two values uses `fail()`, so it reads the same way.
  Each public options interface keeps its TSDoc and is pinned with
  `SchemaMatches<typeof schema, Interface>`, so the two cannot drift without `pnpm type`
  failing. Declare readonly keys with `{ readonly: true }` and flatten an intersection with
  `Flatten<A & B>` — never cast through a `Writable<>` helper. An options argument an
  overload must tell from a positional value uses `record: "argument"`; a string vocabulary
  uses `s.member`; a tagged union uses `s.variant`. A guard that lives in a module is passed
  to `s.custom` and never imported into `@utils/schema`, which is Layer 0.

  A check that runs per node, per cell or per surface call may stay a hand-written predicate
  when the schema form measures slower (draw and pdf do, at about 1.6–2.3 times; Excel's
  per-row setters likewise); it still words its refusal through the schema's `fail`, so the
  message is the same, and says so beside it with the measurement. A check shared by hundreds
  of entry points (a worksheet-handle check, the Word writer's model validation) cannot know
  its caller; it states the problem without a prefix (`ws must be a worksheet, got 42`) and
  `field` plus the stack locate it.

  Nothing unchecked reaches a part: a value written into a package still goes through the
  checked serialisers (`xmlValue` and the drawing-geometry checks). And it is a gate rather
  than a habit: `src/test/__tests__/public-argument-errors.node.test.ts` calls every
  namespace member of every entry, and constructs every class, with wrong arguments, and
  fails on an error that is not the library's own, that belongs to another module than the
  entry's, or that is an argument error naming no `field`. A member it exempts is named
  there with the reason (a deprecated forward, a refusal the platform owns, a private
  constructor), and an exemption that stops being needed fails too. It does not reach
  instance methods or a wrong argument behind a valid handle, so a module's own tests still
  cover those. An instance method is as much the boundary as a namespace function — so is a
  method of an object a public call hands out (`Pdf.PageBuilder#getContentStream` returns a
  `PdfContentStream`) — and a refusal there is the module's error too, a destroyed or
  finished object's included (`StreamStateError`, never a bare `Error`).

- **A read keeps what it cannot type.** A file's value outside a field's vocabulary is
  neither put in the typed field nor dropped: the field reads as absent, the text is kept
  beside it (`@utils/unrecognized-values`) and written back until something sets the
  field, and the module's read report lists it. A vocabulary is the schema's, held to the
  ECMA-376 schemas by a test (`src/test/fixtures/ooxml-enumerations.ts`) — a narrow
  hand-written list otherwise surfaces only when a file from elsewhere carries the value.
- **Names compare the way Office compares them.** Case-insensitive name matching uses
  `excelNameKey` (tables, columns, defined names) or `sheetNameKey` (sheet names, sheet
  references, `LET`/`LAMBDA` parameters) from `@utils/name-key`; never `toUpperCase()` or
  `toLowerCase()`.
- **Acronyms are words**: `toSvg`, `toPng`, `addJson`, `MathMl`, `XmlParseOptions`.
  US spelling (`color`, `centered`).
- **Excel's document model is addressed by position** — `(ws, "A1")`, `(ws, row)`,
  `(ws, "B")` — and hands out no row, column or cell handles; handles belong to the
  streaming API (`Stream.*`), where a row is written once and committed. Row-level work
  lives on `Row`, column-level on `Column`; `Worksheet` keeps what concerns the sheet.
  A cell is **one** argument, `addr: string | CellPosition` (`"A1"` or `{ row, col }`,
  1-based); a position is never split across positional arguments, and an overload is never
  chosen by `arguments.length` — a missing argument must be a type error, not a column read
  as a value. A batch of cells is a range (`Range.setValues`, `Range.setStyle`), not a loop
  of `Cell.*` calls.

  A streaming writer takes the document model's vocabulary. A namespace function whose
  operation makes sense for a writer accepts `Stream.WorkbookWriter` /
  `Stream.WorksheetWriter` beside the document handle — `Workbook.*` (sheets, protection,
  cell styles, defined names), `Worksheet.merge` / `protect`, `Column.*` settings,
  `Cell.*`, `Row.*`, `Range.*`, `Image.*`, `Watermark.*`, `ConditionalFormatting.*`,
  `DataValidation.*` — and dispatches to the writer's own implementation, so the writer's
  state rules apply: a committed workbook or sheet refuses an edit with
  `ExcelStreamStateError`; a row the writer has written refuses an edit **and** a read
  (`surface/stream-targets.ts`: `checkStreamRowsOpen` / `checkStreamRowsHeld`), because a
  silently incomplete answer is worse than none; a multi-row write checks the span before
  touching a cell; column settings are refused once written with the first committed row.
  Detection is structural, never `instanceof`, and is not moved into `checkWorkbook` /
  `checkWorksheet`: the writers pass themselves to core functions on purpose. What needs
  rows already written or the whole document — `Row.insert` / `splice` / `each`,
  `Column.getValues`, `Worksheet.unmerge`, `Workbook.removeWorksheet`, `Workbook.toBuffer`
  (a writer finishes with its own `commit()` / `toBuffer()`) — refuses a writer with
  `ExcelUnsupportedFeatureError` naming the reason, through `refuseStreamWorksheet` or
  `workbookOnly`. That refusal is for JavaScript and `any`; a typed caller gets a compile error
  first, because those signatures name the document model only and a writer is not
  structurally one. `excel/__tests__/type/stream-writer-refusals.typecheck.ts` holds every
  refused member to that, so a new refusal belongs there too. A writer overload that changes the return type keeps the document model's
  signature both first (what `any` picks) and last (what `ReturnType` reads).

  Two ways to address a streamed cell are kept, because they take different things: `Cell` /
  `Row` / `Range` address **by position**; `Stream.setCell*` / `Stream.setRow*` /
  `Stream.getRowCell` take **the handle `addRow` returns**, the row just appended, with no
  address to build. Committing is a handle operation only (`Stream.commitRow`). The
  exceljs-named class methods are deprecated forwards; what stays on the classes has no
  positional counterpart — `addRow`, `getRow` / `getCell` (get-or-create handles),
  `findRow` / `findCell`, `getColumn`, `eachRow`, `commit`, `toBuffer`, and the
  column-key registry. A capability the writer gains is added by teaching the namespace
  function to accept it, not as a new class method.

- **Names are honest about what they do.** A verifier that checks integrity but not trust
  reports `intact` and `claimedSigner`, not `valid` and `signer`; a capability check
  that answers "vector or raster" is `canDrawVectorPdf`, not `canDrawPdf`.
- **A display-list node factory is `create*`** (`Draw.Geometry.createRectNode`), not a
  bare noun (`rectNode`).
- **Who changes what is visible in the namespace.** In Word, `Document.*` edits the
  document it is given, in place; every other function that returns a `DocxDocument`
  leaves its input unchanged and may share unchanged nodes with it. The two never reach each
  other: a function that returns a document derived from another calls
  `markDerived(result, ...inputs)` (`@word/core/ownership`), and a `Document.*` edit takes
  its containers through `ownContainer` and calls `ownTree` before changing nodes inside the
  tree, so what two documents share is copied once, by whichever is edited first. Never mutate
  a container or node reached from a caller's document without one of them. `Document.clone`
  is for a snapshot, not for isolation. A new in-place edit goes in `Document`, never in
  `Query`.
- **A read returns what it read.** `Workbook.read(bytes)` returns a new workbook, as
  `Document.read(bytes)` returns a new document; nothing reads _into_ a handle the caller
  made. An edit applies when it is made; there is no separate commit step. An edit
  that validates checks before it changes anything (as every `Table.*` edit does), so
  one that throws leaves the object as it was.
  A `getModel` says in its documentation whether it hands out a copy or the live
  model; where it is a copy (`Table.getModel`), the matching `setModel` applies it in
  full, and where it is live (`Cell.getModel`, `Chart.getModel`) edits go through the
  namespace (`Chart.mutate`).
- **A handle's published type is what a caller may touch.** A new internal field
  carries `@internal` and is stripped from the declarations (`stripInternal`). One
  that was already published is not removed: it is marked `@deprecated` and stripped
  at 1.0, like any other name. A field left in the type is a promise.

**Renaming a public name** keeps the old one as a forward, with the old behaviour, until
1.0 — every deprecated name is removed together there, not one minor at a time, so a
consumer migrates once. After 1.0 the package follows semver strictly. "The old behaviour"
is what the old name did with arguments it accepted; where a forward shares its replacement's
implementation, an argument the old name never accepted may now be the module's argument
error instead of a `TypeError` from inside. What an accepted argument returned must not
change — `Archive.getEncryptionMethodName` still answers an unknown method with `undefined`
while `Archive.Crypto.getEncryptionMethodName` refuses it.

```ts
/** @deprecated Use {@link getName}. Removed in 1.0. */
export const name = getName;
```

Use a `const`, not `export { getName as name }`: tooling does not reliably carry
`@deprecated` through an export specifier. A deprecated name appears in no README,
example or TSDoc `@example`. A capability is never removed by a rename: the old name
forwards to a non-deprecated member that can do everything it did.

Two shapes cannot be a `const`, and say so beside them. An `enum` keeps its old name with
`export { FormulaTokenType as TokenType }`, because a `const` plus `type` pair loses the
member types (`NodeType.Number` stops compiling as a type). And a class method whose
positional form becomes an options object keeps the old form as a `@deprecated` overload
of the same method, dispatched with `isRecordArgument`
(`BasicRasterCanvas#drawText(x, y, text, options)`).

## Code Style

- **Type-only imports**: `import type { Foo } from "..."`
- **Error handling**: Extend `BaseError` from `@utils/errors`, use `{ cause }` for chaining.
- **Files**: kebab-case. **Browser variants**: `*.browser.ts` — see below.
- **Formatting**: Handled entirely by Prettier — just run `pnpm format`.
- **Tests**: Vitest, in `__tests__/*.test.ts`. Timeout: 30s.
  - **Co-locate tests next to the code they cover.** A test lives in the `__tests__/` directory of the module subfolder it exercises — e.g. `core/__tests__/`, `surface/__tests__/`, `stream/__tests__/`, `chart/__tests__/`, `bridge/__tests__/`, `utils/__tests__/`, `xlsx/__tests__/`. The `xlsx/__tests__/` tree mirrors the `xlsx/xform/` source layout. Do not pile module tests into a single top-level `__tests__/`.
  - **Shared fixtures stay centralized.** Cross-cutting test assets — `data/` (binary `.xlsx`/`.png`/`.csv` fixtures), `helpers/` (e.g. `expect-valid-xlsx`, `zip-text`, `external-oracle`), and `shared/` (reusable sheet builders) — live in `src/modules/excel/__tests__/` and are imported via the `@excel/__tests__/...` alias from any co-located test. A test that is private to one subfolder may keep a private helper beside it (e.g. `chart/__tests__/chart-builder.helpers.ts`).
  - **Browser tests** stay under a `__tests__/browser/` directory (matched by `vitest.browser.config.ts`); keep their `__screenshots__/` baselines alongside them.
    - **The browser config has no `setupFiles`, and it must stay that way.** A setup file runs once per test file, in that file's own iframe, so anything it does is paid ~410 times. The one that existed injected the 1.5 MB Excel IIFE so two smoke tests could read `Documonster.Excel` off the global; deleting it and having those two call `loadIife` themselves cut the local browser run from 72s to 39s, and most of the CI run's 98s import phase with it. It also removed the largest thing each iframe left behind: Chromium does not reclaim a finished test file's memory ([530892387](https://issues.chromium.org/issues/530892387)), and the renderer died mid-run with `Browser connection was closed while running tests` — abandoning the ~120 files that had not run yet, under a summary that says `290 passed`. A test that needs a shipped bundle loads it in its own `beforeAll` through `@test/browser/load-iife`, which caches per iframe.
  - **A Node-only test is named `*.node.test.ts`, and that is machine-enforced.** `vitest.browser.config.ts` re-runs whole directories of Excel/Word/formula/xml/markdown/draw tests in a real Chromium, so a test that reads `process`, `Buffer`, `__dirname`/`__filename` or imports `node:*` has to opt out by name — the config excludes `*.node.test.ts` by glob. `.oxlintrc.json` carries an override whose `files`/`excludeFiles` mirror that config's `include`/`exclude` and turn on `no-restricted-globals` and `no-restricted-imports` there, so `pnpm lint` catches it in seconds. Before that, the only thing catching it was the `Browser` CI job — a `dist/` build plus a real browser, over two minutes, after the push — and three files landed at once breaking it.
    - Prefer keeping a file platform-free over renaming it: a `globalThis` probe (`(globalThis as { process?: … }).process`) or a platform-free helper from `@utils/binary` (`decodeBytesToString` instead of `Buffer.from(x).toString("hex")`) keeps the browser coverage. Rename, or split the Node-only cases into a `.node.test.ts` sibling, only when the test genuinely needs the platform — a top-level `node:` import cannot be skipped at runtime, so it must be split.
    - Two known edges. `typeof process !== "undefined"` is reported like a bare read, so write the `globalThis` probe instead — it is the form the rest of the tree uses. And the globs match `*.test.ts` only, so a `*.helpers.ts` beside a browser-run test is not linted; the `Browser` job is still the backstop for that one.

## Functions, Arrow Functions & Classes

Choose the form by purpose, not by preference. Do **not** make everything an arrow function — each form exists for a reason.

- **Top-level named functions → `function` declarations.** Use `function foo() {}` for module-level functions. They are hoisted (free ordering, no top-of-file dependency dance), carry a real name in stack traces, and support recursion cleanly.
- **Callbacks & inline functions → arrow functions.** Use arrows for `map`/`filter`/`forEach`, promise chains, event handlers, and anywhere lexical `this` is wanted. Keep bodies expression-form when possible (`x => x * 2`, not `x => { return x * 2; }`).
- **Overloads & generators → must be `function`.** Multiple call signatures (TS overloads) and `function*` generators cannot be expressed as arrows.
- **Class members → method syntax, never arrow fields.** Write `load() {}`, not `load = () => {}`. Methods live on the prototype and are shared across instances; arrow fields allocate a fresh function per instance (measured ~5× memory on hot value types like `Cell`/`Token`/XML nodes) and add per-`new` construction cost. Only use an arrow field when a method is detached and passed as a callback that genuinely needs bound `this`.
- **Avoid named function expressions** (`const foo = function bar() {}`); prefer a `function` declaration or an arrow.

### Prefer plain functions over classes

- **Don't reach for `class` by default.** If a unit of behavior is just data + a few transforms, prefer plain functions operating on plain objects/interfaces over a class. Modules with named exports already give you encapsulation and namespacing.
- **Use a `class` only when you genuinely need** instance identity with mutable state, inheritance/polymorphism, lifecycle (`implements`/`extends`), or a public API where `new`/methods read more naturally than free functions.
- **Avoid classes that are just namespaces** — a class with only static members (or a single method) should be plain exported functions instead.

## Examples

All runnable examples write output to `tmp/` under the project root. This directory is gitignored.

`scripts/run-examples.ts` (`pnpm verify:examples`, and the `Examples` CI job) **discovers and
runs every one of them**. Until it did, examples were the only artefact here that nothing
executed: the runner carried a hand-written list covering a sixth of the tree, and no CI job
invoked it. The first full run found six failures — five examples that could not resolve their
imports and one that hit a genuine `RangeError: Maximum call stack size exceeded` in the PDF
exporter (`push(...parts)` turning an unbounded array into an argument list). A seventh surfaced
on the second run: an example that could not run twice.

It runs at two levels, because the two ways an example breaks are different. The pre-commit hook
runs `--changed`: the examples the commit edits, plus every example beside a `utils/` helper it
edits, which is seconds. The `Examples` CI job runs **all** of them, catching an example broken
by a change to the library it calls; that is over a minute, too slow for a hook. It is
deliberately absent from `pnpm test`, which CI runs across every supported Node version and
three operating systems: the failures this catches are neither version- nor platform-specific,
so repeating them once per matrix cell would buy nothing.

Both the discovery rules and `--changed` are covered by
`src/test/__tests__/run-examples.test.ts`, against fixture trees built to break each one — the
runner takes `--root` for that, like the other gates. The hook used to carry its own shell
pattern for the same job, and it had drifted three ways: renames were skipped, nested example
directories were invisible, and a path containing a space was split in two.

Rules when writing an example:

- **Use the public API.** An example that has to import from `@excel/core/...`, `@pdf/reader/...`
  or `__tests__/` is either the wrong file or evidence of a missing public member — treat it as
  the latter until proven otherwise. That is how `Row.addPageBreak`, `Column.setNumFmt` and
  `Cell.getView` came to exist.
- **Be re-runnable.** The gate runs an example against the `tmp/` output of its own previous
  run. Anything that defaults to `overwrite: "error"` needs the opt-in.
- **Write only to `tmp/`.** Accept an optional output path as `process.argv[2]` and default it
  under `tmp/`; never write into the source tree.
- **Run it with `pnpm example`, not `npx tsx`.** `tsx` hands the resolved path to Node's ESM
  resolver, which treats a dot-suffix as a complete filename — so a specifier ending `.node` or
  `.browser` never gets `.ts` appended, and this repository's platform-variant convention breaks.
  `pnpm example` runs `node --import @oxc-node/core/register`, which resolves both tsconfig
  `paths` and an extensionless TypeScript target. Node's own `--experimental-strip-types` is not
  enough either: it is strip-only and this tree uses parameter properties. Fifty-three
  `Run: npx tsx …` comments said otherwise; five of them were instructions that could not work.
- **A benchmark is not an example.** Something that measures internal machinery against the
  platform's belongs in `benchmark/`, where reaching into internals is legitimate. Two stream
  benchmarks sat in `examples/` and were the sole reason anything wanted `@utils/event-emitter`
  published.
- **Fixtures live in `examples/data/`,** not in `__tests__/data/`. The test tree is excluded from
  every build and absent from the published package, so an example reading from it cannot be
  copied — and its output would be coupled to an asset free to change for unrelated reasons.
  A file needed by both is duplicated on purpose.
