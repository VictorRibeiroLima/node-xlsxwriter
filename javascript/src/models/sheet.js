// @ts-check

const Cell = require('./cell');
const Format = require('./format');
const Link = require('./link');
const Formula = require('./formula');
const { ConditionalFormat } = require('./conditional_format');

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
class ConditionalFormatSheetValue {
  /**
   * @param {number} firstRow - The first row of the range
   * @param {number} lastRow - The last row of the range
   * @param {number} firstColumn - The first column of the range
   * @param {number} lastColumn - The last column of the range
   * @param {ConditionalFormat} format - The format of the range
   */
  constructor(firstRow, lastRow, firstColumn, lastColumn, format) {
    /**
     * The first row of the range
     * @type {number}
     */
    this.firstRow = firstRow;
    /**
     * The last row of the range
     * @type {number}
     */
    this.lastRow = lastRow;
    /**
     * The first column of the range
     * @type {number}
     */
    this.firstColumn = firstColumn;
    /**
     * The last column of the range
     * @type {number}
     */
    this.lastColumn = lastColumn;
    /**
     * The format of the range
     * @type {ConditionalFormat}
     */
    this.format = format;
  }
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
class ArrayFormulaSheetValue {
  /**
   * @param {Object} opts - The options for the array formula
   * @param {number} opts.firstRow - The first row of the range
   * @param {number} opts.lastRow - The last row of the range
   * @param {number} opts.firstColumn - The first column of the range
   * @param {number} opts.lastColumn - The last column of the range
   * @param {Formula} opts.formula - The formula of the range
   * @param {Format} [opts.format] - The format of the range
   */
  constructor(opts) {
    const { firstRow, lastRow, firstColumn, lastColumn, formula } = opts;
    /**
     * The first row of the range
     * @type {number}
     */
    this.firstRow = firstRow;
    /**
     * The last row of the range
     * @type {number}
     */
    this.lastRow = lastRow;
    /**
     * The first column of the range
     * @type {number}
     */
    this.firstColumn = firstColumn;
    /**
     * The last column of the range
     * @type {number}
     */
    this.lastColumn = lastColumn;
    /**
     * The formula of the range
     * @type {Formula}
     */
    this.formula = formula;

    /**
     * The format of the range
     * @type {Format|undefined}
     */
    this.format = opts.format ?? undefined;
  }
}

/**
 *
 * @class Sheet
 * @classdesc A sheet is a collection of cells.
 * @property {string} name - The name of the sheet
 * @property {Cell[]} cells - The cells in the sheet
 * @property {ConditionalFormatSheetValue[]} conditionalFormats - The conditional format values of the sheet
 * @property {ArrayFormulaSheetValue[]} arrayFormulas - The array formulas of the sheet
 * @property {RowCellConfig[]} rowConfigs - The rows of the sheet
 * @property {RowCellConfig[]} columnConfigs - The columns of the sheet
 */
class Sheet {
  /**
   * @param {string} name - The name of the sheet
   */
  constructor(name) {
    /**
     * The name of the sheet
     * @type {string}
     */
    this.name = name;
    /**
     * The cells in the sheet
     * @type {Cell[]}
     */
    this.cells = [];

    /**
     * The conditional format values of the sheet
     * @type {ConditionalFormatSheetValue[]}
     */
    this.conditionalFormats = [];

    /**
     * The array formulas of the sheet
     * @type {ArrayFormulaSheetValue[]}
     */
    this.arrayFormulas = [];

    /**
     * The rows of the sheet
     * @type {RowCellConfig[]}
     * @default []
     * */
    this.rowConfigs = [];

    /**
     * The columns of the sheet
     * @type {RowCellConfig[]}
     * @default []
     * */
    this.columnConfigs = [];
  }

  /**
   * Adds a row configuration to the sheet.
   * Rows are the first ones to be processed,so if any value overlaps with the columns it will be overwritten
   * @param {RowCellConfig} config - The configuration of the row
   * @returns {void}
   */
  addRowConfig(config) {
    this.rowConfigs.push(config);
  }

  /**
   * Adds a column configuration to the sheet
   * @param {RowCellConfig} config - The configuration of the column
   * @returns {void}
   */
  addColumnConfig(config) {
    this.columnConfigs.push(config);
  }

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
  addConditionalFormat(opts) {
    const { firstRow, lastRow, firstColumn, lastColumn, format } = opts;
    const conditionalSheetValue = new ConditionalFormatSheetValue(
      firstRow,
      lastRow,
      firstColumn,
      lastColumn,
      format,
    );
    this.conditionalFormats.push(conditionalSheetValue);
  }

  /**
   * @param {Object} opts - The options for the array formula
   * @param {number} opts.firstRow - The first row of the range
   * @param {number} opts.lastRow - The last row of the range
   * @param {number} opts.firstColumn - The first column of the range
   * @param {number} opts.lastColumn - The last column of the range
   * @param {Formula} opts.formula - The formula of the range
   * @param {Format} [opts.format] - The format of the range
   */
  addArrayFormula(opts) {
    const arrayFormula = new ArrayFormulaSheetValue(opts);
    this.arrayFormulas.push(arrayFormula);
  }

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
  writeCell(row, col, value, cellType, format) {
    if (col > 65_535 || col < 0) {
      throw new Error('Invalid column index');
    }
    if (row > 1_048_577 || row < 0) {
      throw new Error('Invalid row index');
    }
    const cell = new Cell({
      col,
      row,
      value,
      cellType,
      format,
    });
    this.cells.push(cell);
  }

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
  writeString(row, col, value, format) {
    this.writeCell(row, col, value, 'string', format);
  }

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
  writeNumber(row, col, value, format) {
    if (!isNaN(value)) {
      this.writeCell(row, col, value, 'number', format);
      return;
    }
    this.writeCell(row, col, value, 'string', format);
  }

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
  writeLink(row, col, value, format) {
    this.writeCell(row, col, value, 'link', format);
  }

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
  writeDate(row, col, value, format) {
    this.writeCell(row, col, value, 'date', format);
  }

  /**
   * writes a formula value to a cell
   * @param {number} row - the cell row
   * @param {number} col - the cell col
   * @param {Formula} value - The value to write to the cell
   * @param {Format} [format] - The format of the cell
   * @returns {void}
   */
  writeFormula(row, col, value, format) {
    this.writeCell(row, col, value, 'formula', format);
  }

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
  writeFromJson(objects, { headerFormat, cellFormat, columnFormats } = {}) {
    if (objects.length === 0) {
      return;
    }

    const keys = Object.keys(objects[0]);
    // write headers
    for (let col = 0; col < keys.length; col++) {
      const format = headerFormat || columnFormats?.[keys[col]];
      this.writeString(0, col, keys[col], format);
    }

    // write data
    for (let col = 0; col < objects.length; col++) {
      for (let row = 0; row < keys.length; row++) {
        const format = columnFormats?.[keys[row]] || cellFormat;
        const type = typeof objects[col][keys[row]];
        const value = objects[col][keys[row]];
        switch (type) {
          case 'string':
            this.writeString(col + 1, row, value, format);
            break;
          case 'number':
            this.writeNumber(col + 1, row, value, format);
            break;
          case 'object':
            if (value instanceof Link) {
              this.writeLink(col + 1, row, value, format);
            } else if (value instanceof Date) {
              this.writeCell(col + 1, row, value, 'date', format);
            } else {
              this.writeCell(col + 1, row, value, null, format);
            }
            break;
          default:
            this.writeCell(col + 1, row, value, null, format);
        }
      }
    }
  }
}

module.exports = { Sheet, ArrayFormulaSheetValue };
