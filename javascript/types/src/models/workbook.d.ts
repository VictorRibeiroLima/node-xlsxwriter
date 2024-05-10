export = Workbook;
/**
 *
 * @class Workbook
 * @classdesc Represents a workbook
 * @property {Sheet[]} sheets - The sheets in the workbook
 */
declare class Workbook {
    /**
     * The sheets in the workbook
     * @type {Sheet[]}
     */
    sheets: Sheet[];
    /**
     * Adds a sheet to the workbook
     *
     * @return {Sheet} The new sheet
     */
    addSheet(): Sheet;
    /**
     * Pushes a sheet to the workbook
     * @param {Sheet} sheet - The sheet to be pushed
     * @returns {void}
     */
    pushSheet(sheet: Sheet): void;
    /**
     * Gets a sheet from the workbook
     * @param {number} index - The index of the sheet
     * @returns {Sheet|undefined} The sheet
     * @throws {Error} The index is out of range
     */
    worksheetFromIndex(index: number): Sheet | undefined;
    /**
     * Gets a sheet from the workbook
     * @param {string} name - The name of the sheet
     * @returns {Sheet|undefined} The sheet
     */
    worksheetFromName(name: string): Sheet | undefined;
    /**
     * Creates a workbook to be written as a buffer.(using a child process for the asynchronous operation)
     * @returns {Promise<Buffer>}
     * @throws {Error}
     */
    saveToBuffer(): Promise<Buffer>;
    /**
     * Creates a workbook to be written as a buffer.
     * @returns {Buffer}
     * @throws {Error}
     */
    saveToBufferSync(): Buffer;
    /**
     * Writes a workbook to a file.
     * @param {string} path - The path of the file.
     * @returns {void}
     * @throws {Error}
     */
    saveToFileSync(path: string): void;
    /**
     * Writes a workbook to a file.(using a child process for the asynchronous operation)
     * @param {string} path - The path of the file.
     * @returns {Promise<void>}
     * @throws {Error}
     */
    saveToFile(path: string): Promise<void>;
    /**
     * Writes a workbook to a base64 string.
     * @returns {Promise<string>}
     * @throws {Error}
     */
    saveToBase64(): Promise<string>;
    /**
     * Writes a workbook to a base64 string.
     * @returns {string}
     * @throws {Error}
     */
    saveToBase64Sync(): string;
}
import { Sheet } from "./sheet";
//# sourceMappingURL=workbook.d.ts.map