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
class MergedCell {
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
  constructor(opts) {
    /**
     * The first row index of the merged cell
     * @type {number}
     */
    this.firstRow = opts.firstRow;
    /**
     * The first column index of the merged cell
     * @type {number}
     */
    this.firstCol = opts.firstCol;
    /**
     * The last row index of the merged cell
     * @type {number}
     */
    this.lastRow = opts.lastRow;
    /**
     * The last column index of the merged cell
     * @type {number}
     */
    this.lastCol = opts.lastCol;
    /**
     * The value of the merged cell
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
     * @type {Format}
     */
    this.format = opts.format;

   /**
     * Whether the cell is part of a merged cell
     * @type {true}
     */
    Object.defineProperty(this, 'merged', {
      value: true,
      writable: false,
      enumerable: true,
      configurable: false,
    });
  }
}

module.exports = MergedCell;
