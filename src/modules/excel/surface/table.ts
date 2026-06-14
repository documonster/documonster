/**
 * `Table` namespace surface.
 * `import { Table } from "documonster/excel"` → `Table.add(ws, model)`.
 */
export {
  addTable as add,
  getTable as get,
  getTables as list,
  removeTable as remove
} from "@excel/worksheet";
