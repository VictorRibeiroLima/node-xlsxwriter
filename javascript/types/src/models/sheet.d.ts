export type SizeConfig = {
    /**
     * - The value of the size
     */
    value: number;
    /**
     * - The unit of the size
     */
    unit?: "auto" | "px";
};
export type RowCellConfig = {
    /**
     * - The index of the row/column, 0-based
     */
    index: number;
    /**
     * - The format of the cell (will be overwritten by the cell format)
     */
    format?: Format;
    /**
     * - The height/width of the row/column
     */
    size?: SizeConfig;
    /**
     * - Whether the row is hidden
     */
    hidden?: boolean;
};
/**
 *
 * @class Sheet
 * @classdesc A sheet is a collection of cells.
 * @property {string} name - The name of the sheet
 * @property {Cell[]} cells - The cells in the sheet
 * @property {ConditionalFormatSheetValue[]} conditionalFormats - The conditional format values of the sheet
 * @property {ArrayFormulaSheetValue[]} arrayFormulas - The array formulas of the sheet
 * @property {TableSheetValue[]} tables - The tables of the sheet
 * @property {RowCellConfig[]} rowConfigs - The rows of the sheet
 * @property {RowCellConfig[]} columnConfigs - The columns of the sheet
 */
export class Sheet {
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
     * The array formulas of the sheet
     * @type {ArrayFormulaSheetValue[]}
     */
    arrayFormulas: ArrayFormulaSheetValue[];
    /**
     * The rows of the sheet
     * @type {RowCellConfig[]}
     * @default []
     * */
    rowConfigs: RowCellConfig[];
    /**
     * The columns of the sheet
     * @type {RowCellConfig[]}
     * @default []
     * */
    columnConfigs: RowCellConfig[];
    /**
     * The tables of the sheet
     * @type {TableSheetValue[]}
     * @default []
     * */
    tables: TableSheetValue[];
    /**
     * Adds a row configuration to the sheet.
     * Rows are the first ones to be processed,so if any value overlaps with the columns it will be overwritten
     * @param {RowCellConfig} config - The configuration of the row
     * @returns {void}
     */
    addRowConfig(config: RowCellConfig): void;
    /**
     * Adds a column configuration to the sheet
     * @param {RowCellConfig} config - The configuration of the column
     * @returns {void}
     */
    addColumnConfig(config: RowCellConfig): void;
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
     * @param {Object} opts - The options for the array formula
     * @param {number} opts.firstRow - The first row of the range
     * @param {number} opts.lastRow - The last row of the range
     * @param {number} opts.firstColumn - The first column of the range
     * @param {number} opts.lastColumn - The last column of the range
     * @param {Formula} opts.formula - The formula of the range
     * @param {Format} [opts.format] - The format of the range
     */
    addArrayFormula(opts: {
        firstRow: number;
        lastRow: number;
        firstColumn: number;
        lastColumn: number;
        formula: Formula;
        format?: Format;
    }): void;
    /**
     * Adds a table to the sheet
     * @param {Object} opts - The options for the table
     * @param {number} opts.firstRow - The first row of the range
     * @param {number} opts.lastRow - The last row of the range
     * @param {number} opts.firstColumn - The first column of the range
     * @param {number} opts.lastColumn - The last column of the range
     * @param {Table} opts.table - The table of the range
     * @throws {Error} - Invalid table range
     * @returns {void}
     */
    addTable(opts: {
        firstRow: number;
        lastRow: number;
        firstColumn: number;
        lastColumn: number;
        table: Table;
    }): void;
    /**
     * Writes a cell to the sheet
     *
     * @param {number} row - the cell row
     * @param {number} col - the cell col
     * @param {string|number|Link|Date|Formula|any} value - The value of the cell.
     * @param {("number"|"string"|"link"|"date"|"formula")} [cellType] - The type of the cell(if not provider .toString() will be used)
     * @param {Format} [format] - The format of the cell
     * @returns {void}
     * @throws {Error} - col > 65_535 or col < 0
     * @throws {Error} - row > 1_048_577 or row < 0
     */
    writeCell(row: number, col: number, value: string | number | Link | Date | Formula | any, cellType?: ("number" | "string" | "link" | "date" | "formula"), format?: Format): void;
    /**
     * writes a string value to a cell
     * @param {number} row - the cell row
     * @param {number} col - the cell col
     * @param {string} value - The value to write to the cell
     * @param {Format} [format] - The format of the cell
     * @returns {void}
     * @throws {Error} - col > 65_535 or col < 0
     * @throws {Error} - row > 1_048_577 or row < 0
     */
    writeString(row: number, col: number, value: string, format?: Format): void;
    /**
     * writes a number value to a cell
     * @param {number} row - the cell row
     * @param {number} col - the cell col
     * @param {number} value - The value to write to the cell
     * @param {Format} [format] - The format of the cell
     * @returns {void}
     * @throws {Error} - col > 65_535 or col < 0
     * @throws {Error} - row > 1_048_577 or row < 0
     */
    writeNumber(row: number, col: number, value: number, format?: Format): void;
    /**
     * writes a link value to a cell
     * @param {number} row - the cell row
     * @param {number} col - the cell col
     * @param {Link} value - The value to write to the cell
     * @param {Format} [format] - The format of the cell
     * @returns {void}
     * @throws {Error} - col > 65_535 or col < 0
     * @throws {Error} - row > 1_048_577 or row < 0
     */
    writeLink(row: number, col: number, value: Link, format?: Format): void;
    /**
     * writes a date value to a cell
     * @param {number} row - the cell row
     * @param {number} col - the cell col
     * @param {Date} value - The value to write to the cell
     * @param {Format} [format] - The format of the cell
     * @returns {void}
     * @throws {Error} - col > 65_535 or col < 0
     * @throws {Error} - row > 1_048_577 or row < 0
     */
    writeDate(row: number, col: number, value: Date, format?: Format): void;
    /**
     * writes a formula value to a cell
     * @param {number} row - the cell row
     * @param {number} col - the cell col
     * @param {Formula} value - The value to write to the cell
     * @param {Format} [format] - The format of the cell
     * @returns {void}
     */
    writeFormula(row: number, col: number, value: Formula, format?: Format): void;
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
/**
 * @class ArrayFormulaSheetValue
 * @classdesc Represents the values of an array formula sheet.
 * @property {number} firstRow - The first row of the range
 * @property {number} lastRow - The last row of the range
 * @property {number} firstColumn - The first column of the range
 * @property {number} lastColumn - The last column of the range
 * @property {Formula} formula - The formula of the range
 * @property {Format} [format] - The format of the range
 */
export class ArrayFormulaSheetValue {
    /**
     * @param {Object} opts - The options for the array formula
     * @param {number} opts.firstRow - The first row of the range
     * @param {number} opts.lastRow - The last row of the range
     * @param {number} opts.firstColumn - The first column of the range
     * @param {number} opts.lastColumn - The last column of the range
     * @param {Formula} opts.formula - The formula of the range
     * @param {Format} [opts.format] - The format of the range
     */
    constructor(opts: {
        firstRow: number;
        lastRow: number;
        firstColumn: number;
        lastColumn: number;
        formula: Formula;
        format?: Format;
    });
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
     * The formula of the range
     * @type {Formula}
     */
    formula: Formula;
    /**
     * The format of the range
     * @type {Format|undefined}
     */
    format: Format | undefined;
}
import Format = require("./format");
import Cell = require("./cell");
/**
 * @typedef {Object} SizeConfig
 * @property {number} value - The value of the size
 * @property {"auto"|"px"} [unit] - The unit of the size
 */
/**
 * @typedef {Object} RowCellConfig
 * @property {number} index - The index of the row/column, 0-based
 * @property {Format} [format] - The format of the cell (will be overwritten by the cell format)
 * @property {SizeConfig} [size] - The height/width of the row/column
 * @property {boolean} [hidden] - Whether the row is hidden
 */
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
/**
 * @class TableSheetValue
 * @classdesc Represents the values of a table sheet.
 * @property {number} firstRow - The first row of the range
 * @property {number} lastRow - The last row of the range
 * @property {number} firstColumn - The first column of the range
 * @property {number} lastColumn - The last column of the range
 * @property {Table} table - The table of the range
 */
declare class TableSheetValue {
    /**
     * @param {number} firstRow - The first row of the range
     * @param {number} lastRow - The last row of the range
     * @param {number} firstColumn - The first column of the range
     * @param {number} lastColumn - The last column of the range
     * @param {Table} table - The table of the range
     */
    constructor(firstRow: number, lastRow: number, firstColumn: number, lastColumn: number, table: Table);
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
     * The table of the range
     * @type {Table}
     */
    table: Table;
}
import { ConditionalFormat } from "./conditional_format";
import Formula = require("./formula");
import { Table } from "./table";
import Link = require("./link");
export {};
//# sourceMappingURL=sheet.d.ts.map