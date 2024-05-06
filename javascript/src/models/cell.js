// @ts-check

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
class Cell {
  /**
   * @param {number} col - The column index of the cell
   * @param {number} row - The row index of the cell
   * @param {string|number} value - The value of the cell
   * @param {CellType} cellType - The type of the cell
   */
  constructor(col, row, value, cellType) {
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
    /**
     * The type of the cell
     * @type {CellType}
     */
    this.cellType = cellType;
  }
}

module.exports = Cell;
