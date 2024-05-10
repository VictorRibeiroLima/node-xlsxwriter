// @ts-check
const Format = require('./format');
const Link = require('./link');
const Formula = require('./formula');

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
class Cell {
  /**
   * @param {Object} opts - Options for the cell
   * @param {number} opts.col - The column index of the cell
   * @param {number} opts.row - The row index of the cell
   * @param {CellValue} opts.value - The value of the cell
   * @param {CellType} [opts.cellType] - The type of the cell
   * @param {Format} [opts.format] - The format of the cell
   */
  constructor(opts) {
    /**
     * The column index of the cell
     * @type {number}
     */
    this.col = opts.col;
    /**
     * The row index of the cell
     * @type {number}
     */
    this.row = opts.row;
    /**
     * The value of the cell
     * @type {CellValue}
     */
    this.value = opts.value;

    /**
     * The type of the cell
     * @type {CellType|undefined}
     */
    this.cellType = opts.cellType ?? undefined;

    /**
     * The format of the cell
     * @type {Format|undefined}
     */
    this.format = opts.format ?? undefined;
  }
}

module.exports = Cell;
