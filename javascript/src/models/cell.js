// @ts-check
const Format = require('./format');
const Link = require('./link');

/**
 * @typedef {(number|string|Link)} CellValue
 * @typedef {("number"|"string"|"link")} CellType
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
   * @param {number} col - The column index of the cell
   * @param {number} row - The row index of the cell
   * @param {string|number|Link} value - The value of the cell
   * @param {CellType} [cellType] - The type of the cell
   * @param {Format} [format] - The format of the cell
   */
  constructor(col, row, value, cellType, format) {
    /**
     * The column index of the cell
     * @type {number}
     */
    this.col = col;
    /**
     * The row index of the cell
     * @type {number}
     */
    this.row = row;
    /**
     * The value of the cell
     * @type {CellValue}
     */
    this.value = value;
    if (cellType) {
      /**
       * The type of the cell
       * @type {CellType}
       */
      this.cellType = cellType;
    }
    if (format) {
      /**
       * The format of the cell
       * @type {Format}
       */
      this.format = format;
    }
  }
}

module.exports = Cell;
