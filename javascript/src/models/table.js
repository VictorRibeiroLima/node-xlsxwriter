//@ts-check

const Formula = require('./formula');
const Format = require('./format');

/**
 * @typedef {(
 * "none"|
 * "light1"|
 * "light2"|
 * "light3"|
 * "light4"|
 * "light5"|
 * "light6"|
 * "light7"|
 * "light8"|
 * "light9"|
 * "light10"|
 * "light11"|
 * "light12"|
 * "light13"|
 * "light14"|
 * "light15"|
 * "light16"|
 * "light17"|
 * "light18"|
 * "light19"|
 * "light20"|
 * "light21"|
 * "medium1"|
 * "medium2"|
 * "medium3"|
 * "medium4"|
 * "medium5"|
 * "medium6"|
 * "medium7"|
 * "medium8"|
 * "medium9"|
 * "medium10"|
 * "medium11"|
 * "medium12"|
 * "medium13"|
 * "medium14"|
 * "medium15"|
 * "medium16"|
 * "medium17"|
 * "medium18"|
 * "medium19"|
 * "medium20"|
 * "medium21"|
 * "medium22"|
 * "medium23"|
 * "medium24"|
 * "medium25"|
 * "medium26"|
 * "medium27"|
 * "medium28"|
 * "dark1"|
 * "dark2"|
 * "dark3"|
 * "dark4"|
 * "dark5"|
 * "dark6"|
 * "dark7"|
 * "dark8"|
 * "dark9"|
 * "dark10"|
 * "dark11")} TableStyle
 */

/**
 * @typedef {(
 * "none"|
 * "average"|
 * "count"|
 * "countNumbers"|
 * "max"|
 * "min"|
 * "sum"|
 * "stdDev"|
 * "var"|
 * "custom"
 * )} TableFunctionType
 */

/**
 * @class
 * @classdesc defines functions for worksheet table total rows
 * @property {TableFunctionType} type - function type
 * @property {Formula} [formula] - custom formula
 */
class TableFunction {
  /**
   * @param {Object} obj
   * @param {TableFunctionType} obj.type
   * @param {Formula} [obj.formula]
   * @throws {Error} if type is not valid
   * @throws {Error} if type is custom and formula is not provided
   */
  constructor(obj) {
    /**
     * @type {TableFunctionType}
     */
    this.type = obj.type;
    if (this.type === 'custom' && !obj.formula) {
      throw new Error('custom function requires formula');
    }

    /**
     * @type {Formula|undefined|null}
     * @default undefined
     */
    this.formula = obj.formula;
  }
}

/**
 * @class
 * @classdesc defines a table column
 * @property {Format} [format] - column format
 * @property {Formula} [formula] - column formula
 * @property {string} [header] - column header
 * @property {Format} [headerFormat] - column header format
 * @property {TableFunction} [totalFunction] - column total function
 * @property {string} [totalLabel] - column total label
 */
class TableColumn {
  /**
   * @param {Object} obj
   * @param {Format} [obj.format]
   * @param {Formula} [obj.formula]
   * @param {string} [obj.header]
   * @param {Format} [obj.headerFormat]
   * @param {Object} [obj.totalFunction]
   * @param {TableFunctionType} obj.totalFunction.type
   * @param {Formula} [obj.totalFunction.formula]
   * @param {string} [obj.totalLabel]
   * @throws {Error} if totalFunction is not valid
   */
  constructor(obj) {
    /**
     * @type {Format|undefined|null}
     * @default undefined
     */
    this.format = obj.format;

    /**
     * @type {Formula|undefined|null}
     * @default undefined
     */
    this.formula = obj.formula;

    /**
     * @type {string|undefined|null}
     * @default undefined
     */
    this.header = obj.header;

    /**
     * @type {Format|undefined|null}
     * @default undefined
     */
    this.headerFormat = obj.headerFormat;

    /**
     * @type {TableFunction|undefined|null}
     * @default undefined
     */
    this.totalFunction = obj.totalFunction
      ? new TableFunction(obj.totalFunction)
      : undefined;

    /**
     * @type {string|undefined|null}
     * @default undefined
     */
    this.totalLabel = obj.totalLabel;
  }

  /**
   * @param {Format} format
   * @returns {void}
   * Set the format for a table column
   */
  setFormat(format) {
    this.format = format;
  }

  /**
   * @param {Object} obj
   * @param {TableFunctionType} obj.type
   * @param {Formula} [obj.formula]
   * @throws {Error} if type is not valid
   * @throws {Error} if type is custom and formula is not provided
   * @returns {void}
   * Set the total function for the total row of a table column.
   * Set the SUBTOTAL() function for the “totals” row of a table column.
   * Note, overwriting the total row cells with worksheet.write() calls will cause Excel to warn that the table is corrupt when loading the file.
   */
  setTotalFunction(obj) {
    this.totalFunction = new TableFunction(obj);
  }

  /**
   * @param {string} header
   * @returns {void}
   * Set the header for a table column
   */
  setHeader(header) {
    this.header = header;
  }

  /**
   * @param {Format} format
   * @returns {void}
   * Set the header format for a table column
   */
  setHeaderFormat(format) {
    this.headerFormat = format;
  }

  /**
   * @param {string} totalLabel
   * @returns {void}
   * Set a label for the total row of a table column.
   * It is possible to set a label for the totals row of a column instead of a subtotal function.
   * Note, overwriting the total row cells with worksheet.write() calls will cause Excel to warn that the table is corrupt when loading the file.
   */
  setTotalLabel(totalLabel) {
    this.totalLabel = totalLabel;
  }

  /**
   * @param {Formula} formula
   * @returns {void}
   * Set the formula for a table column
   */
  setFormula(formula) {
    this.formula = formula;
  }
}

/**
 * @class
 * @classdesc Represents a worksheet Table
 * @property {string} [name] - table name - By default Excel (Table1 .. TableN)
 * @property {boolean} [autoFilter=true] - enable auto filtering
 * @property {boolean} [bandedColumns=false] - enable banded columns
 * @property {boolean} [bandedRows=false] - enable banded rows
 * @property {TableColumn[]} columns - table columns
 * @property {boolean} [firstColumnHighlighted=false] - enable first column highlighting
 * @property {boolean} [lastColumnHighlighted=false] - enable last column highlighting
 * @property {boolean} [headerRow=true] - enable header row
 * @property {TableStyle} [style] - table style
 * @property {boolean} [totalRow=false] - enable total row
 */
class Table {
  /**
   * @param {Object} obj
   * @param {string} [obj.name]
   * @param {boolean} [obj.autoFilter=true]
   * @param {boolean} [obj.bandedColumns=false]
   * @param {boolean} [obj.bandedRows=false]
   * @param {TableColumn[]} [obj.columns]
   * @param {boolean} [obj.firstColumnHighlighted=false]
   * @param {boolean} [obj.lastColumnHighlighted=false]
   * @param {boolean} [obj.headerRow=true]
   * @param {TableStyle} [obj.style]
   * @param {boolean} [obj.totalRow=false]
   */
  constructor(obj) {
    /**
     * @type {string|undefined|null}
     * @default undefined
     */
    this.name = obj.name;

    /**
     * @type {boolean}
     * @default true
     */
    this.autoFilter =
      obj.autoFilter === undefined || obj.autoFilter === null
        ? true
        : obj.autoFilter;

    /**
     * @type {boolean}
     * @default false
     */
    this.bandedColumns = obj.bandedColumns || false;

    /**
     * @type {boolean}
     * @default false
     */
    this.bandedRows = obj.bandedRows || false;

    /**
     * @type {TableColumn[]}
     */
    this.columns = obj.columns || [];

    /**
     * @type {boolean}
     * @default false
     */
    this.firstColumnHighlighted = obj.firstColumnHighlighted || false;

    /**
     * @type {boolean}
     * @default false
     */
    this.lastColumnHighlighted = obj.lastColumnHighlighted || false;

    /**
     * @type {boolean}
     * @default true
     */
    this.headerRow =
      obj.headerRow === undefined || obj.headerRow === null
        ? true
        : obj.headerRow;

    /**
     * @type {TableStyle|undefined|null}
     * @default undefined
     */
    this.style = obj.style;

    /**
     * @type {boolean}
     * @default false
     */
    this.totalRow = obj.totalRow || false;
  }
}

module.exports = {
  Table,
  TableColumn,
  TableFunction,
};
