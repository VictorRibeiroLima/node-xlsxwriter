export = Sheet;
/**
 *
 * @class Sheet
 * @classdesc A sheet is a collection of cells.
 * @property {string} name - The name of the sheet
 * @property {Cell[]} cells - The cells in the sheet
 * @property {ConditionalFormatSheetValue[]} conditionalFormats - The conditional format values of the sheet
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
     * The conditional format values of the sheet
     * @type {ConditionalFormatSheetValue[]}
     */
    conditionalFormats: ConditionalFormatSheetValue[];
    /**
     * Adds a conditional format to the sheet
     * @param {Object} opts - The options for the conditional format
     * @param {number} opts.firstRow - The first row of the range
     * @param {number} opts.lastRow - The last row of the range
     * @param {number} opts.firstColumn - The first column of the range
     * @param {number} opts.lastColumn - The last column of the range
     * @param {ConditionalFormat} opts.format - The format of the range
     * @returns {void}
     */
    addConditionalFormat(opts: {
        firstRow: number;
        lastRow: number;
        firstColumn: number;
        lastColumn: number;
        format: ConditionalFormat;
    }): void;
    /**
     * Writes a cell to the sheet
     *
     * @param {number} col - The column index of the cell
     * @param {number} row - The row index of the cell
     * @param {string|number|Link|Date|Formula|any} value - The value of the cell.
     * @param {("number"|"string"|"link"|"date"|"formula")} [cellType] - The type of the cell(if not provider .toString() will be used)
     * @param {Format} [format] - The format of the cell
     * @returns {void}
     * @throws {Error} - col > 65_535 or col < 0
     * @throws {Error} - row > 1_048_577 or row < 0
     */
    writeCell(col: number, row: number, value: string | number | Link | Date | Formula | any, cellType?: ("number" | "string" | "link" | "date" | "formula"), format?: Format): void;
    /**
     * writes a string value to a cell
     * @param {number} col - The column index of the cell
     * @param {number} row - The row index of the cell
     * @param {string} value - The value to write to the cell
     * @param {Format} [format] - The format of the cell
     * @returns {void}
     * @throws {Error} - col > 65_535 or col < 0
     * @throws {Error} - row > 1_048_577 or row < 0
     */
    writeString(col: number, row: number, value: string, format?: Format): void;
    /**
     * writes a number value to a cell
     * @param {number} col - The column index of the cell
     * @param {number} row - The row index of the cell
     * @param {number} value - The value to write to the cell
     * @param {Format} [format] - The format of the cell
     * @returns {void}
     * @throws {Error} - col > 65_535 or col < 0
     * @throws {Error} - row > 1_048_577 or row < 0
     */
    writeNumber(col: number, row: number, value: number, format?: Format): void;
    /**
     * writes a link value to a cell
     * @param {number} col - The column index of the cell
     * @param {number} row - The row index of the cell
     * @param {Link} value - The value to write to the cell
     * @param {Format} [format] - The format of the cell
     * @returns {void}
     * @throws {Error} - col > 65_535 or col < 0
     * @throws {Error} - row > 1_048_577 or row < 0
     */
    writeLink(col: number, row: number, value: Link, format?: Format): void;
    /**
     * writes a date value to a cell
     * @param {number} col - The column index of the cell
     * @param {number} row - The row index of the cell
     * @param {Date} value - The value to write to the cell
     * @param {Format} [format] - The format of the cell
     * @returns {void}
     * @throws {Error} - col > 65_535 or col < 0
     * @throws {Error} - row > 1_048_577 or row < 0
     */
    writeDate(col: number, row: number, value: Date, format?: Format): void;
    /**
     * writes a formula value to a cell
     * @param {number} col - The column index of the cell
     * @param {number} row - The row index of the cell
     * @param {Formula} value - The value to write to the cell
     * @param {Format} [format] - The format of the cell
     * @returns {void}
     */
    writeFormula(col: number, row: number, value: Formula, format?: Format): void;
    /**
     *
     * @typedef {Object} FormatOptions
     * @property {Format} [headerFormat] - The format of the header cells
     * @property {Format} [cellFormat] - The format of the data cells
     * @property {Object.<string, Format>} [columnFormats] - The format of the cells in the columns
     * Writes the sheet based on the provided array of objects
     * For performance reasons the headers will be generated based on the first object
     * and those headers will be used for the rest of the objects,so if the objects have different keys
     * the keys that are not in the first object will not be written to the sheet
     * @param {Object[]} objects - The objects to write to the sheet
     * @param {FormatOptions} [opts] - The format of the header cells
     * @returns {void}
     */
    writeFromJson(objects: any[], { headerFormat, cellFormat, columnFormats }?: {
        /**
         * - The format of the header cells
         */
        headerFormat?: Format;
        /**
         * - The format of the data cells
         */
        cellFormat?: Format;
        /**
         * - The format of the cells in the columns
         * Writes the sheet based on the provided array of objects
         * For performance reasons the headers will be generated based on the first object
         * and those headers will be used for the rest of the objects,so if the objects have different keys
         * the keys that are not in the first object will not be written to the sheet
         */
        columnFormats?: {
            [x: string]: Format;
        };
    }): void;
}
import Cell = require("./cell");
/**
 * @class ConditionalFormatSheetValue
 * @classdesc Represents the values of a conditional format sheet.
 * @property {number} firstRow - The first row of the range
 * @property {number} lastRow - The last row of the range
 * @property {number} firstColumn - The first column of the range
 * @property {number} lastColumn - The last column of the range
 * @property {ConditionalFormat} format - The format of the range
 *
 */
declare class ConditionalFormatSheetValue {
    /**
     * @param {number} firstRow - The first row of the range
     * @param {number} lastRow - The last row of the range
     * @param {number} firstColumn - The first column of the range
     * @param {number} lastColumn - The last column of the range
     * @param {ConditionalFormat} format - The format of the range
     */
    constructor(firstRow: number, lastRow: number, firstColumn: number, lastColumn: number, format: ConditionalFormat);
    /**
     * The first row of the range
     * @type {number}
     */
    firstRow: number;
    /**
     * The last row of the range
     * @type {number}
     */
    lastRow: number;
    /**
     * The first column of the range
     * @type {number}
     */
    firstColumn: number;
    /**
     * The last column of the range
     * @type {number}
     */
    lastColumn: number;
    /**
     * The format of the range
     * @type {ConditionalFormat}
     */
    format: ConditionalFormat;
}
import { ConditionalFormat } from "./conditional_format";
import Link = require("./link");
import Formula = require("./formula");
import Format = require("./format");
//# sourceMappingURL=sheet.d.ts.map