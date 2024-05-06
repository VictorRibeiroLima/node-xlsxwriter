export = Sheet;
/**
 *
 * @class Sheet
 * @classdesc A sheet is a collection of cells.
 * @property {string} name - The name of the sheet
 * @property {Cell[]} cells - The cells in the sheet
 */
declare class Sheet {
    /**
     * @param {string} name - The name of the sheet
     */
    constructor(name: string);
    /**
     * The name of the sheet
     * @type {string}
     */
    name: string;
    /**
     * The cells in the sheet
     * @type {Cell[]}
     */
    cells: Cell[];
    /**
     * Writes a cell to the sheet
     *
     * @param {number} col - The column index of the cell
     * @param {number} row - The row index of the cell
     * @param {string|number} value - The value of the cell
     * @param {("number"|"string"|"link")} cellType - The type of the cell
     * @returns {void}
     * @throws {Error} - col > 65_535 or col < 0
     * @throws {Error} - row >  4_294_967_295 or row < 0
     */
    writeCell(col: number, row: number, value: string | number, cellType: ("number" | "string" | "link")): void;
    /**
     * writes a string value to a cell
     * @param {number} col - The column index of the cell
     * @param {number} row - The row index of the cell
     * @param {string} value - The value to write to the cell
     * @returns {void}
     * @throws {Error} - col > 65_535 or col < 0
     * @throws {Error} - row >  4_294_967_295 or row < 0
     * @throws {Error} - value is not a string
     */
    writeString(col: number, row: number, value: string): void;
    /**
     * writes a number value to a cell
     * @param {number} col - The column index of the cell
     * @param {number} row - The row index of the cell
     * @param {number} value - The value to write to the cell
     * @returns {void}
     * @throws {Error} - col > 65_535 or col < 0
     * @throws {Error} - row >  4_294_967_295 or row < 0
     * @throws {Error} - value is not a number
     */
    writeNumber(col: number, row: number, value: number): void;
    /**
     * writes a link value to a cell
     * @param {number} col - The column index of the cell
     * @param {number} row - The row index of the cell
     * @param {string} value - The value to write to the cell
     * @returns {void}
     * @throws {Error} - col > 65_535 or col < 0
     * @throws {Error} - row >  4_294_967_295 or row < 0
     * @throws {Error} - value is not a string
     */
    writeLink(col: number, row: number, value: string): void;
}
import Cell = require("./cell");
//# sourceMappingURL=sheet.d.ts.map