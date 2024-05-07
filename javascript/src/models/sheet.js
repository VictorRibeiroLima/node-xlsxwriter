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
   * @param {string|number|Link|any} value - The value of the cell.
   * @param {("number"|"string"|"link")} [cellType] - The type of the cell(if not provider .toString() will be used)
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
    if (!isNaN(value)) {
      this.writeCell(col, row, value, 'number', format);
      return;
    }
    this.writeCell(col, row, value, 'string', format);
  }

  /**
   * writes a link value to a cell
   * @param {number} col - The column index of the cell
   * @param {number} row - The row index of the cell
   * @param {Link} value - The value to write to the cell
   * @param {Format} [format] - The format of the cell
   * @returns {void}
   * @throws {Error} - col > 65_535 or col < 0
   * @throws {Error} - row > 1_048_577 or row < 0
   * @throws {Error} - value is not a string (null and undefined are allowed)
   */
  writeLink(col, row, value, format) {
    this.writeCell(col, row, value, 'link', format);
  }

  /**
   * Writes the sheet based on the provided array of objects
   * For performance reasons the headers will be generated based on the first object
   * and those headers will be used for the rest of the objects,so if the objects have different keys
   * the keys that are not in the first object will not be written to the sheet
   * @param {Object[]} objects - The objects to write to the sheet
   * @param {Format} [headerFormat] - The format of the header cells
   * @param {Format} [cellFormat] - The format of the data cells
   * @returns {void}
   */
  writeFromJson(objects, headerFormat, cellFormat) {
    if (objects.length === 0) {
      return;
    }
    const keys = Object.keys(objects[0]);
    // write headers
    for (let i = 0; i < keys.length; i++) {
      this.writeString(i, 0, keys[i], headerFormat);
    }

    // write data
    for (let i = 0; i < objects.length; i++) {
      for (let j = 0; j < keys.length; j++) {
        const type = typeof objects[i][keys[j]];
        const value = objects[i][keys[j]];
        switch (type) {
          case 'string':
            this.writeString(j, i + 1, value, cellFormat);
            break;
          case 'number':
            this.writeNumber(j, i + 1, value, cellFormat);
            break;
          case 'object':
            if (value instanceof Link) {
              this.writeLink(j, i + 1, value, cellFormat);
            } else {
              this.writeCell(j, i + 1, value, null, cellFormat);
            }
            break;
          default:
            this.writeCell(j, i + 1, value, null, cellFormat);
        }
      }
    }
  }
}

module.exports = Sheet;
