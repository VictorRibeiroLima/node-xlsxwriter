export = Cell;
/**
 * @typedef {(number|string|Link|Formula)} CellValue
 * @typedef {("number"|"string"|"link"|"date"|"formula")} CellType
 */
/**
 *
 * @class Cell
 * @classdesc Represents a cell in the grid
 * @property {number} col - The column index of the cell
 * @property {number} row - The row index of the cell
 * @property {CellValue} value - The value of the cell
 * @property {CellType} [celType] - The type of the cell.
 * @property {Format} [format] - The format of the cell
 */
declare class Cell {
    /**
     * @param {Object} opts - Options for the cell
     * @param {number} opts.col - The column index of the cell
     * @param {number} opts.row - The row index of the cell
     * @param {CellValue} opts.value - The value of the cell
     * @param {CellType} [opts.cellType] - The type of the cell
     * @param {Format} [opts.format] - The format of the cell
     */
    constructor(opts: {
        col: number;
        row: number;
        value: CellValue;
        cellType?: CellType;
        format?: Format;
    });
    /**
     * The column index of the cell
     * @type {number}
     */
    col: number;
    /**
     * The row index of the cell
     * @type {number}
     */
    row: number;
    /**
     * The value of the cell
     * @type {CellValue}
     */
    value: CellValue;
    /**
     * The type of the cell
     * @type {CellType|undefined}
     */
    cellType: CellType | undefined;
    /**
     * The format of the cell
     * @type {Format|undefined}
     */
    format: Format | undefined;
}
declare namespace Cell {
    export { CellValue, CellType };
}
import Format = require("./format");
type CellValue = (number | string | Link | Formula);
type CellType = ("number" | "string" | "link" | "date" | "formula");
import Link = require("./link");
import Formula = require("./formula");
//# sourceMappingURL=cell.d.ts.map