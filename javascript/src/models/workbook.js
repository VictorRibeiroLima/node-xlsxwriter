const funcs = require('../../../native/node-xlsxwriter.node');
const saveToBuffer = funcs.saveToBuffer;
const saveToBufferSync = funcs.saveToBufferSync;
const saveToFileSync = funcs.saveToFileSync;
const saveToFile = funcs.saveToFile;
const saveToBase64 = funcs.saveToBase64;
const saveToBase64Sync = funcs.saveToBase64Sync;
// @ts-check

const { Sheet } = require('./sheet');
/**
 *
 * @class Workbook
 * @classdesc Represents a workbook
 * @property {Sheet[]} sheets - The sheets in the workbook
 */
class Workbook {
  constructor() {
    /**
     * The sheets in the workbook
     * @type {Sheet[]}
     */
    this.sheets = [];
  }

  /**
   * Adds a sheet to the workbook
   *
   * @return {Sheet} The new sheet
   */
  addSheet() {
    const length = this.sheets.length;
    const name = `Sheet${length + 1}`;
    const sheet = new Sheet(name);
    this.sheets.push(sheet);
    return sheet;
  }

  /**
   * Pushes a sheet to the workbook
   * @param {Sheet} sheet - The sheet to be pushed
   * @returns {void}
   */
  pushSheet(sheet) {
    this.sheets.push(sheet);
  }

  /**
   * Gets a sheet from the workbook
   * @param {number} index - The index of the sheet
   * @returns {Sheet|undefined} The sheet
   * @throws {Error} The index is out of range
   */
  worksheetFromIndex(index) {
    return this.sheets[index].worksheet;
  }

  /**
   * Gets a sheet from the workbook
   * @param {string} name - The name of the sheet
   * @returns {Sheet|undefined} The sheet
   */
  worksheetFromName(name) {
    return this.sheets.find((sheet) => sheet.name === name).worksheet;
  }

  /**
   * Creates a workbook to be written as a buffer.(using a child process for the asynchronous operation)
   * @returns {Promise<Buffer>}
   * @throws {Error}
   */
  async saveToBuffer() {
    return saveToBuffer(this);
  }

  /**
   * Creates a workbook to be written as a buffer.
   * @returns {Buffer}
   * @throws {Error}
   */
  saveToBufferSync() {
    return saveToBufferSync(this);
  }

  /**
   * Writes a workbook to a file.
   * @param {string} path - The path of the file.
   * @returns {void}
   * @throws {Error}
   */
  saveToFileSync(path) {
    return saveToFileSync(this, path);
  }

  /**
   * Writes a workbook to a file.(using a child process for the asynchronous operation)
   * @param {string} path - The path of the file.
   * @returns {Promise<void>}
   * @throws {Error}
   */
  async saveToFile(path) {
    return saveToFile(this, path);
  }

  /**
   * Writes a workbook to a base64 string.
   * @returns {Promise<string>}
   * @throws {Error}
   */
  async saveToBase64() {
    return saveToBase64(this);
  }

  /**
   * Writes a workbook to a base64 string.
   * @returns {string}
   * @throws {Error}
   */
  saveToBase64Sync() {
    return saveToBase64Sync(this);
  }
}

module.exports = Workbook;
