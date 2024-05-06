// @ts-check

const Cell = require('./cell');

/**
 *
 * @class Sheet
 * @classdesc A sheet is a collection of cells.
 * @property {string} name - The name of the sheet
 * @property {Cell[]} cells - The cells in the sheet
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
  }

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
  writeCell(col, row, value, cellType) {
    if (col > 65_535 || col < 0) {
      throw new Error('Invalid column index');
    }
    if (row > 4_294_967_295 || row < 0) {
      throw new Error('Invalid row index');
    }
    const cell = new Cell(col, row, value, cellType);
    this.cells.push(cell);
  }

  /**
   * writes a string value to a cell
   * @param {number} col - The column index of the cell
   * @param {number} row - The row index of the cell
   * @param {string} value - The value to write to the cell
   * @returns {void}
   * @throws {Error} - col > 65_535 or col < 0
   * @throws {Error} - row >  4_294_967_295 or row < 0
   * @throws {Error} - value is can't be converted to a string( null and undefined are allowed)
   */
  writeString(col, row, value) {
    if (value === null || value === undefined) {
      this.writeCell(col, row, value, 'string');
    }
    const type = typeof value;
    if (type !== 'string' && type !== 'number') {
      throw new Error('Value must be capable of being converted to a string');
    }
    this.writeCell(col, row, value, 'string');
  }

  /**
   * writes a number value to a cell
   * @param {number} col - The column index of the cell
   * @param {number} row - The row index of the cell
   * @param {number} value - The value to write to the cell
   * @returns {void}
   * @throws {Error} - col > 65_535 or col < 0
   * @throws {Error} - row >  4_294_967_295 or row < 0
   * @throws {Error} - value is not a number (null and undefined are allowed)
   */
  writeNumber(col, row, value) {
    if (value === null || value === undefined) {
      this.writeCell(col, row, value, 'number');
    }
    if (typeof value !== 'number') {
      throw new Error('Value must be a number');
    }
  }

  /**
   * writes a link value to a cell
   * @param {number} col - The column index of the cell
   * @param {number} row - The row index of the cell
   * @param {string} value - The value to write to the cell
   * @returns {void}
   * @throws {Error} - col > 65_535 or col < 0
   * @throws {Error} - row >  4_294_967_295 or row < 0
   * @throws {Error} - value is not a string (null and undefined are allowed)
   */
  writeLink(col, row, value) {
    if (value === null || value === undefined) {
      this.writeCell(col, row, value, 'link');
    }
    if (typeof value !== 'string') {
      throw new Error('Value must be a string');
    }
    this.writeCell(col, row, value, 'link');
  }
}

module.exports = Sheet;
