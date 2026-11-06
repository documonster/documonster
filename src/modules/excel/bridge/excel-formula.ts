/**
 * `ExcelFormula` namespace surface — workbook recalculation.
 *
 * `import { ExcelFormula } from "documonster/excel/formula"` →
 *   `ExcelFormula.calculate(wb)`.
 *
 * Named `ExcelFormula` rather than `Formula` so it can be imported next to
 * `Formula` from `documonster/formula` without aliasing.
 */
export { calculateFormulas as calculate } from "@excel/core/formula-adapter";
