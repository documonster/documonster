/**
 * Formula cells read from a file as ordinary formulas, by cell model.
 *
 * Such a formula means what Excel computes for it, implicit intersection included, so it is evaluated that way (see
 * `ordinaryEvaluationFormula` in `array-evaluation.ts`) and written back as it came. A formula set through the API is
 * not in here: it is written as a dynamic array where that keeps what Excel 365 computes for it typed, without
 * spilling (see `needsArrayEvaluation`), and is otherwise computed as the ordinary formula it is written as.
 * Kept beside the model rather than on it, so the published model shape does not change; a cell given a new value has
 * a new model and is no longer in here.
 */
const storedOrdinaryFormulas = new WeakSet<object>();

export function markStoredOrdinaryFormula(model: object): void {
  storedOrdinaryFormulas.add(model);
}

export function isStoredOrdinaryFormula(model: object): boolean {
  return storedOrdinaryFormulas.has(model);
}
