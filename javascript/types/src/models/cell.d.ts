export = Cell;
/**
 * @typedef {(number|string)} CellValue
 * @typedef {("number"|"string"|"link")} CellType
 */
/**
 *
 * @class Cell
 * @classdesc Represents a cell in the grid
 * @property {number} col - The column index of the cell
 * @property {number} row - The row index of the cell
 * @property {CellValue} value - The value of the cell
 * @property {CellType} celType - The type of the cell.
 */
declare class Cell {
    /**
     * @param {number} col - The column index of the cell
     * @param {number} row - The row index of the cell
     * @param {string|number} value - The value of the cell
     * @param {CellType} cellType - The type of the cell
     */
    constructor(col: number, row: number, value: string | number, cellType: CellType);
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
}
declare namespace Cell {
    export { CellValue, CellType };
}
type CellValue = (number | string);
type CellType = ("number" | "string" | "link");
//# sourceMappingURL=cell.d.ts.map