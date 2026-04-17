/**
 * Formula Calculation Implementation
 *
 * The pipeline that implements the snapshot → compile → evaluate →
 * materialize → apply architecture.
 *
 * ## Pipeline Steps
 *
 * 1. **Snapshot** — `buildWorkbookSnapshot()` creates an immutable snapshot
 *    of the entire workbook.
 * 2. **Normalize** — `collectFormulaInstances()` extracts all formula cells
 *    into uniform `FormulaInstance` objects.
 * 3. **Parse** — Each formula's source text is tokenized and parsed into an AST.
 * 4. **Compile** — The binder transforms each AST into a `BoundExpr` tree,
 *    resolving names, structured references, and sheet references.
 * 5. **Dependency Analysis** — Static dependencies are extracted from bound
 *    expressions and a topological evaluation order is computed.
 * 6. **Evaluate** — Formulas are evaluated in dependency order using the
 *    evaluator which operates on `BoundExpr` and produces `RuntimeValue`.
 * 7. **Materialize** — Evaluation results are converted into a `WritebackPlan`.
 * 8. **Apply** — The plan is applied to the live workbook.
 */

import { parseRefRange } from "../compile/address-utils";
import { bind, type BindingContext } from "../compile/binder";
import type { BoundExpr } from "../compile/bound-ast";
import {
  extractStaticDeps,
  analyzeExpr,
  detectDynamicArrayFunction,
  type NameDepResolver,
  type CompiledFormula
} from "../compile/compiled-formula";
import {
  buildDependencyGraphFromDeps,
  topologicalSort,
  mergeDynamicDeps
} from "../compile/dependency-analysis";
import { buildWritebackPlan } from "../materialize/build-writeback-plan";
import { getPersistentSpillMap, getGhostSnapshots } from "../materialize/spill-engine";
import type { WorkbookLike } from "../materialize/types";
import {
  EvalSession,
  evaluateFormula,
  evaluateFormulaRaw,
  type EvalContext
} from "../runtime/evaluator";
import type { RuntimeValue } from "../runtime/values";
import { RVKind, rvNumber, BLANK, ERRORS } from "../runtime/values";
import type { AstNode } from "../syntax/ast";
import { parse } from "../syntax/parser";
import { tokenize } from "../syntax/tokenizer";
import { applyWritebackPlan } from "./apply-writeback-plan";
import { collectFormulaInstances } from "./formula-instance";
import type { FormulaInstance } from "./formula-instance";
import { buildWorkbookSnapshot } from "./workbook-adapter";
import { formulaCellKey, resolveDefinedName, type WorkbookSnapshot } from "./workbook-snapshot";

// ============================================================================
// Persistent Caches (keyed by workbook — survive across invocations)
// ============================================================================

/**
 * Persistent AST cache: formula text → parsed AstNode.
 * Since AST is a pure function of formula text, this is safe to cache
 * across calculation cycles. Keyed by workbook to allow GC.
 */
const persistentAstCache = new WeakMap<WeakKey, Map<string, AstNode>>();

function getPersistentAstCache(workbook: WorkbookLike): Map<string, AstNode> {
  let cache = persistentAstCache.get(workbook);
  if (!cache) {
    cache = new Map();
    persistentAstCache.set(workbook, cache);
  }
  return cache;
}

// ============================================================================
// Main: Formula Calculation Implementation
// ============================================================================

/**
 * Recalculate all formula cells using the new pipeline.
 *
 * This implements the full snapshot → compile → evaluate → materialize → apply
 * architecture. The workbook is mutated only at the final apply step.
 */
export function calculateFormulasImpl(workbook: WorkbookLike): void {
  // ── Step 1: Snapshot ──
  const snapshot = buildWorkbookSnapshot(workbook);

  // ── Step 2: Normalize ──
  const instances = collectFormulaInstances(snapshot);
  if (instances.length === 0) {
    // Clean up stale spills even when there are no formulas
    cleanupStaleSpillsIfNeeded(workbook, snapshot);
    return;
  }

  // ── Step 3: Parse ──
  // Use persistent AST cache — formula text → AST is a pure function,
  // so parsed ASTs can be safely reused across calculation cycles.
  const astCache = getPersistentAstCache(workbook);
  for (const inst of instances) {
    parseFormulaText(inst.sourceText, astCache);
  }

  // ── Step 4: Compile (Bind) ──
  const compiledMap = new Map<string, CompiledFormula>();
  const results = new Map<string, RuntimeValue>();
  for (const inst of instances) {
    const compiled = compileFormula(inst, astCache, snapshot);
    const key = formulaCellKey(inst.sheetName, inst.row, inst.col);
    if ("reason" in compiled) {
      // Parse or bind failure — produce an explicit error result.
      // #CALC! for engine-level failures, #NAME? for parse errors that
      // may indicate an unsupported construct.
      results.set(key, compiled.reason === "parse" ? ERRORS.NAME : ERRORS.CALC);
    } else {
      compiledMap.set(key, compiled);
    }
  }

  // ── Step 5: Dependency Analysis + Topological Sort ──
  // Build a producer map so that formulas depending on cells that are:
  //   (a) CSE target range slaves, or
  //   (b) previous-cycle spill ghost cells
  // correctly get ordered after their producing formula.
  const producerMap = new Map<string, string>();

  // CSE: distribute targetRef across all slave cells → master key
  for (const [masterKey, cf] of compiledMap) {
    const inst = cf.instance;
    if (inst.kind === "cse" && inst.targetRef) {
      const rng = parseRefRange(inst.targetRef);
      if (rng) {
        for (let r = rng.top; r <= rng.bottom; r++) {
          for (let c = rng.left; c <= rng.right; c++) {
            const slaveKey = formulaCellKey(inst.sheetName, r, c);
            if (slaveKey !== masterKey) {
              producerMap.set(slaveKey, masterKey);
            }
          }
        }
      }
    }
  }

  // Dynamic array spill: use previous-cycle spill regions as static hint.
  // The actual spill may differ this cycle, but this covers the common case
  // of a formula that reads a stable spill target.
  const persistentSpills = getPersistentSpillMap(workbook);
  for (const [, region] of persistentSpills) {
    const ws = snapshot.worksheetsById.get(region.worksheetId);
    if (!ws) {
      continue;
    }
    const masterKey = formulaCellKey(ws.name, region.sourceRow, region.sourceCol);
    if (!compiledMap.has(masterKey)) {
      continue; // source formula no longer exists
    }
    for (let r = 0; r < region.rows; r++) {
      for (let c = 0; c < region.cols; c++) {
        if (r === 0 && c === 0) {
          continue; // skip source
        }
        const ghostKey = formulaCellKey(ws.name, region.sourceRow + r, region.sourceCol + c);
        if (!producerMap.has(ghostKey)) {
          producerMap.set(ghostKey, masterKey);
        }
      }
    }
  }

  // Build the dependency graph with producer remapping.
  let graph = buildDependencyGraphFromDeps(compiledMap, producerMap);
  let evalOrder = topologicalSort(graph);

  // ── Step 6: Evaluate ──
  const session = new EvalSession();
  const indirectCache = new Map<string, BoundExpr>();
  const ctx: EvalContext = {
    snapshot,
    compiledFormulas: compiledMap,
    astCache: indirectCache,
    currentSheet: snapshot.worksheets[0]?.name ?? ""
  };

  // Evaluate in topological order
  evaluateInOrder(evalOrder, compiledMap, results, ctx, session);

  // ── Step 6b: Merge dynamic dependencies and re-evaluate if needed ──
  // After the first pass, formulas with INDIRECT/OFFSET have recorded their
  // actual runtime cell accesses in session.dynamicDeps. Merge these edges
  // into the graph and re-evaluate any formulas whose dependencies changed.
  if (session.dynamicDeps.size > 0) {
    const mergeResult = mergeDynamicDeps(graph, session.dynamicDeps);
    if (mergeResult.changed) {
      const prevCircularKeys = graph.circularKeys;
      graph = mergeResult.graph;
      evalOrder = topologicalSort(graph);

      // Collect formulas that gained new deps AND their transitive dependents.
      // Without clearing dependents, downstream cells could see stale values.
      const toClear = new Set<string>();
      const queue: string[] = [];
      for (const [formulaKey] of session.dynamicDeps) {
        if (!toClear.has(formulaKey)) {
          toClear.add(formulaKey);
          queue.push(formulaKey);
        }
      }
      // If the merge introduced new circular refs, include all new members too
      for (const key of graph.circularKeys) {
        if (!prevCircularKeys.has(key) && !toClear.has(key)) {
          toClear.add(key);
          queue.push(key);
        }
      }
      // BFS through reverse edges to find all transitive dependents
      let head = 0;
      while (head < queue.length) {
        const key = queue[head++];
        const deps = graph.dependedBy.get(key);
        if (deps) {
          for (const depKey of deps) {
            if (!toClear.has(depKey)) {
              toClear.add(depKey);
              queue.push(depKey);
            }
          }
        }
      }
      // Clear all affected formulas
      for (const key of toClear) {
        session.resultCache.delete(key);
        results.delete(key);
      }
      // Formula-based name results may depend on cell values that just changed
      session.nameCache.clear();

      // Re-evaluate the full order (evaluateInOrder skips already-computed cells)
      evaluateInOrder(evalOrder, compiledMap, results, ctx, session);
    }
  }

  // ── Iterative Calculation for Circular References ──
  const iterateEnabled = snapshot.calcProperties.iterate === true;
  if (iterateEnabled && graph.circularKeys.size > 0) {
    const maxIter = snapshot.calcProperties.iterateCount ?? 100;
    const delta = snapshot.calcProperties.iterateDelta ?? 0.001;

    const circularKeys: string[] = [];
    for (const key of evalOrder) {
      if (graph.circularKeys.has(key)) {
        circularKeys.push(key);
      }
    }

    for (let iter = 0; iter < maxIter; iter++) {
      let maxChange = 0;

      // Clear ALL cached values — non-circular cells may depend on circular
      // cells and must be re-evaluated with updated values each iteration.
      session.resultCache.clear();
      session.nameCache.clear();

      // Seed fallback from previous results
      for (const key of circularKeys) {
        const compiled = compiledMap.get(key);
        if (!compiled) {
          continue;
        }
        const prev = results.get(key);
        if (prev !== undefined && prev.kind !== RVKind.Error) {
          session.circularFallback.set(key, prev);
        } else {
          session.circularFallback.set(key, rvNumber(0));
        }
      }

      for (const key of circularKeys) {
        const compiled = compiledMap.get(key);
        if (!compiled) {
          continue;
        }

        try {
          const oldResult = results.get(key);
          const newResult = evaluateFormula(compiled, ctx, session);
          results.set(key, newResult);
          session.circularFallback.set(key, newResult);

          if (oldResult && oldResult.kind === RVKind.Number && newResult.kind === RVKind.Number) {
            maxChange = Math.max(maxChange, Math.abs(newResult.value - oldResult.value));
          }
        } catch {
          // Iterative evaluation threw — set error and continue convergence
          results.set(key, ERRORS.CALC);
        }
      }

      if (maxChange <= delta) {
        break;
      }
    }

    session.circularFallback.clear();

    // Re-evaluate non-circular cells that depend (transitively) on circular cells,
    // since they were initially computed with stale fallback values.
    const circularKeySet = new Set(circularKeys);
    const affectedDownstream = new Set<string>();
    const queue: string[] = [...circularKeys];
    let head = 0;
    while (head < queue.length) {
      const key = queue[head++];
      const deps = graph.dependedBy.get(key);
      if (deps) {
        for (const depKey of deps) {
          if (!circularKeySet.has(depKey) && !affectedDownstream.has(depKey)) {
            affectedDownstream.add(depKey);
            queue.push(depKey);
          }
        }
      }
    }
    if (affectedDownstream.size > 0) {
      for (const key of affectedDownstream) {
        session.resultCache.delete(key);
        results.delete(key);
      }
      // Clear name cache — formula-based names may depend on circular cell values
      session.nameCache.clear();
      evaluateInOrder(evalOrder, compiledMap, results, ctx, session);
    }
  }

  // ── Step 7: Materialize (Build Writeback Plan) ──
  const previousSpills = getPersistentSpillMap(workbook);
  const previousGhosts = getGhostSnapshots(workbook);

  const plan = buildWritebackPlan(
    snapshot,
    [...compiledMap.values()],
    results,
    previousSpills,
    previousGhosts
  );

  // ── Step 8: Apply ──
  applyWritebackPlan(workbook, plan);
}

// ============================================================================
// Helper: Evaluate Formulas in Order
// ============================================================================

/**
 * Evaluate compiled formulas in the given topological order.
 *
 * For each formula in `evalOrder`:
 * - If already in `results` (from a previous pass or compile failure), skip.
 * - Volatile formulas always re-evaluate (bypass scalar cache).
 * - CSE / dynamic array formulas use `evaluateFormulaRaw`.
 * - Normal scalar formulas use `evaluateFormula` with cache.
 */
function evaluateInOrder(
  evalOrder: readonly string[],
  compiledMap: ReadonlyMap<string, CompiledFormula>,
  results: Map<string, RuntimeValue>,
  ctx: EvalContext,
  session: EvalSession
): void {
  for (const key of evalOrder) {
    // Skip if already computed (e.g. compile failure → #CALC!, or previous pass)
    if (results.has(key)) {
      continue;
    }

    const compiled = compiledMap.get(key);
    if (!compiled) {
      continue;
    }

    const inst = compiled.instance;
    const isCSE = inst.kind === "cse" && inst.targetRef;
    const isDynamic = inst.isDynamicArray || compiled.isDynamicArrayFunction;

    // Volatile formulas always re-evaluate — bypass cache
    if (compiled.isVolatile) {
      session.resultCache.delete(key);
    }

    if (isCSE || isDynamic) {
      try {
        const raw = evaluateFormulaRaw(compiled, ctx, session);
        results.set(key, raw);
        if (isCSE && inst.targetRef) {
          populateCSECache(inst, raw, session);
        }
      } catch {
        results.set(key, ERRORS.CALC);
      }
    } else {
      if (!compiled.isVolatile) {
        const cachedResult = session.resultCache.get(key);
        if (cachedResult !== undefined) {
          results.set(key, cachedResult.scalar);
          continue;
        }
      }
      try {
        const scalar = evaluateFormula(compiled, ctx, session);
        results.set(key, scalar);
      } catch {
        results.set(key, ERRORS.CALC);
      }
    }
  }
}

// ============================================================================
// Helper: Parse Formula Text
// ============================================================================

function parseFormulaText(formula: string, astCache: Map<string, AstNode>): AstNode | null {
  const cached = astCache.get(formula);
  if (cached) {
    return cached;
  }
  try {
    const tokens = tokenize(formula);
    const ast = parse(tokens);
    astCache.set(formula, ast);
    return ast;
  } catch {
    return null;
  }
}

// ============================================================================
// Helper: Compile Formula (with failure diagnostics)
// ============================================================================

/**
 * Failure reason codes for formula compilation.
 * These provide fine-grained diagnostics beyond the generic #CALC! error.
 */
type CompileFailure =
  | { reason: "parse"; formula: string }
  | { reason: "bind"; formula: string; sheet: string };

function compileFormula(
  inst: FormulaInstance,
  astCache: Map<string, AstNode>,
  snapshot: WorkbookSnapshot
): CompiledFormula | CompileFailure {
  const ast = astCache.get(inst.sourceText);
  if (!ast) {
    return { reason: "parse", formula: inst.sourceText };
  }

  const bindCtx: BindingContext = {
    snapshot,
    currentSheet: inst.sheetName
  };

  try {
    const bound = bind(ast, bindCtx);

    // Build a name resolver that parses formula-based defined names
    // so their deps and dynamic-ref flags propagate to the outer formula.
    const nameDepCache = new Map<string, ReturnType<NameDepResolver>>();
    const nameResolver: NameDepResolver = upperName => {
      // Include sheet in cache key for scope-aware resolution
      const cacheKey = `${upperName}\0${inst.sheetName}`;
      if (nameDepCache.has(cacheKey)) {
        return nameDepCache.get(cacheKey);
      }
      // Prevent infinite recursion
      nameDepCache.set(cacheKey, undefined);

      // Use scope-aware resolution (sheet-local → workbook-global)
      const dn = resolveDefinedName(snapshot.definedNames, upperName, inst.sheetName);
      if (!dn || dn.ranges.length !== 1) {
        return undefined;
      }
      const rangeStr = dn.ranges[0];
      // Only process formula-based names (not cell/range refs)
      if (parseRefRange(rangeStr)) {
        return undefined;
      }
      try {
        const nameTokens = tokenize(rangeStr);
        const nameAst = parse(nameTokens);
        const nameBound = bind(nameAst, { snapshot, currentSheet: inst.sheetName });
        const nameDeps = extractStaticDeps(nameBound, snapshot, nameResolver);
        const nameAnalysis = analyzeExpr(nameBound, nameResolver);
        const result = { deps: nameDeps, hasDynamicRefs: nameAnalysis.hasDynamicRefs };
        nameDepCache.set(cacheKey, result);
        return result;
      } catch {
        return undefined;
      }
    };

    const staticDeps = extractStaticDeps(bound, snapshot, nameResolver);
    const analysis = analyzeExpr(bound, nameResolver);

    return {
      instance: inst,
      ast,
      bound,
      staticDeps,
      isVolatile: analysis.isVolatile,
      hasDynamicRefs: analysis.hasDynamicRefs,
      containsLambda: analysis.containsLambda,
      isDynamicArrayFunction: detectDynamicArrayFunction(ast, bound)
    };
  } catch {
    return { reason: "bind", formula: inst.sourceText, sheet: inst.sheetName };
  }
}

// ============================================================================
// Helper: Cleanup Stale Spills
// ============================================================================

function cleanupStaleSpillsIfNeeded(workbook: WorkbookLike, snapshot: WorkbookSnapshot): void {
  const previousSpills = getPersistentSpillMap(workbook);
  if (previousSpills.size === 0) {
    return;
  }

  // No formula cells → all spills are stale
  const plan = buildWritebackPlan(
    snapshot,
    [],
    new Map(),
    previousSpills,
    getGhostSnapshots(workbook)
  );

  applyWritebackPlan(workbook, plan);
}

// ============================================================================
// Helper: Populate CSE Session Cache
// ============================================================================

/**
 * For CSE array formulas, immediately populate the session cache for all
 * cells in the target range. This ensures that when other formulas reference
 * cells within a CSE range, they see the correct distributed values instead
 * of re-evaluating independently.
 */
function populateCSECache(inst: FormulaInstance, result: RuntimeValue, session: EvalSession): void {
  const ref = inst.targetRef;
  if (!ref) {
    return;
  }
  const range = parseRefRange(ref);
  if (!range) {
    return;
  }
  const { top, left, bottom, right } = range;

  const numRows = bottom - top + 1;
  const numCols = right - left + 1;

  if (result.kind === RVKind.Array) {
    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < numCols; c++) {
        const cellKey = formulaCellKey(inst.sheetName, top + r, left + c);
        const val = result.rows[r]?.[c];
        const sv = val ?? BLANK;
        session.resultCache.set(cellKey, { scalar: sv, raw: sv });
      }
    }
  } else {
    // Scalar — fill entire range
    const scalar =
      result.kind === RVKind.Error ||
      result.kind === RVKind.Number ||
      result.kind === RVKind.String ||
      result.kind === RVKind.Boolean
        ? result
        : BLANK;
    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < numCols; c++) {
        const cellKey = formulaCellKey(inst.sheetName, top + r, left + c);
        session.resultCache.set(cellKey, { scalar, raw: scalar });
      }
    }
  }
}
