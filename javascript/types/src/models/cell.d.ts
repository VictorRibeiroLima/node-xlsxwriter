export = Cell;
/**
 * @typedef {(number|string|Link)} CellValue
 * @typedef {("number"|"string"|"link"|"date")} CellType
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
     * @param {number} col - The column index of the cell
     * @param {number} row - The row index of the cell
     * @param {string|number|Link} value - The value of the cell
     * @param {CellType} [cellType] - The type of the cell
     * @param {Format} [format] - The format of the cell
     */
    constructor(col: number, row: number, value: string | number | Link, cellType?: CellType, format?: Format);
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
     * @type {CellType}
     */
    cellType: CellType;
    /**
     * The format of the cell
     * @type {Format}
     */
    format: Format;
}
declare namespace Cell {
    export { CellValue, CellType };
}
type CellValue = (number | string | Link);
type CellType = ("number" | "string" | "link" | "date");
import Format = require("./format");
import Link = require("./link");
//# sourceMappingURL=cell.d.ts.map