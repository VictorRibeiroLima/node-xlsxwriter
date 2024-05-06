// @ts-check

const Cell = require('./cell');
const Format = require('./format');
const Link = require('./link');

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
   * @param {string|number|Link} value - The value of the cell
   * @param {("number"|"string"|"link")} [cellType] - The type of the cell
   * @param {Format} [format] - The format of the cell
   * @returns {void}
   * @throws {Error} - col > 65_535 or col < 0
   * @throws {Error} - row > 1_048_577 or row < 0
   */
  writeCell(col, row, value, cellType, format) {
    if (col > 65_535 || col < 0) {
      throw new Error('Invalid column index');
    }
    if (row > 1_048_577 || row < 0) {
      throw new Error('Invalid row index');
    }
    const cell = new Cell(col, row, value, cellType, format);
    this.cells.push(cell);
  }

  /**
   * writes a string value to a cell
   * @param {number} col - The column index of the cell
   * @param {number} row - The row index of the cell
   * @param {string} value - The value to write to the cell
   * @param {Format} [format] - The format of the cell
   * @returns {void}
   * @throws {Error} - col > 65_535 or col < 0
   * @throws {Error} - row > 1_048_577 or row < 0
   * @throws {Error} - value is can't be converted to a string( null and undefined are allowed)
   */
  writeString(col, row, value, format) {
    this.writeCell(col, row, value, 'string', format);
  }

  /**
   * writes a number value to a cell
   * @param {number} col - The column index of the cell
   * @param {number} row - The row index of the cell
   * @param {number} value - The value to write to the cell
   * @param {Format} [format] - The format of the cell
   * @returns {void}
   * @throws {Error} - col > 65_535 or col < 0
   * @throws {Error} - row > 1_048_577 or row < 0
   * @throws {Error} - value is not a number (null and undefined are allowed)
   */
  writeNumber(col, row, value, format) {
    this.writeCell(col, row, value, 'number', format);
  }

  /**
   * writes a link value to a cell
   * @param {number} col - The column index of the cell
   * @param {number} row - The row index of the cell
   * @param {string} value - The value to write to the cell
   * @param {Format} [format] - The format of the cell
   * @returns {void}
   * @throws {Error} - col > 65_535 or col < 0
   * @throws {Error} - row > 1_048_577 or row < 0
   * @throws {Error} - value is not a string (null and undefined are allowed)
   */
  writeLink(col, row, value, format) {
    this.writeCell(col, row, value, 'link', format);
  }
}

module.exports = Sheet;
