export = MergedCell;
/**
 * @typedef {(number|string|Link|Formula)} CellValue
 * @typedef {("number"|"string"|"link"|"date"|"formula")} CellType
 */
/**
 *
 * @class MergedCell
 * @classdesc Represents a merged cell in the grid
 * @property {number} firstRow - The first row index of the merged cell
 * @property {number} firstCol - The first column index of the merged cell
 * @property {number} lastRow - The last row index of the merged cell
 * @property {number} lastCol - The last column index of the merged cell
 * @property {CellValue} value - The value of the cell
 * @property {true} merged - Whether the cell is part of a merged cell
 * @property {Format} format - The format of the cell
 * @property {CellType} [celType] - The type of the cell.
 */
declare class MergedCell {
    /**
     * @param {Object} opts - Options for the cell
     * @param {number} opts.firstRow - The first row index of the merged cell
     * @param {number} opts.firstCol - The first column index of the merged cell
     * @param {number} opts.lastRow - The last row index of the merged cell
     * @param {number} opts.lastCol - The last column index of the merged cell
     * @param {CellValue} opts.value - The value of the cell
     * @param {Format} opts.format - The format of the cell
     * @param {CellType} [opts.cellType] - The type of the cell
  
     */
    constructor(opts: {
        firstRow: number;
        firstCol: number;
        lastRow: number;
        lastCol: number;
        value: CellValue;
        format: Format;
        cellType?: CellType;
    });
    /**
     * The first row index of the merged cell
     * @type {number}
     */
    firstRow: number;
    /**
     * The first column index of the merged cell
     * @type {number}
     */
    firstCol: number;
    /**
     * The last row index of the merged cell
     * @type {number}
     */
    lastRow: number;
    /**
     * The last column index of the merged cell
     * @type {number}
     */
    lastCol: number;
    /**
     * The value of the merged cell
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
     * @type {Format}
     */
    format: Format;
}
declare namespace MergedCell {
    export { CellValue, CellType };
}
import Format = require("./format");
type CellValue = (number | string | Link | Formula);
type CellType = ("number" | "string" | "link" | "date" | "formula");
import Link = require("./link");
import Formula = require("./formula");
//# sourceMappingURL=merged_cell.d.ts.map