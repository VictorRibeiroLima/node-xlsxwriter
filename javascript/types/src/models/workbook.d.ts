export = Workbook;
declare class Workbook {
    sheets: Sheet[];
    addSheet(): Sheet;
    pushSheet(sheet: Sheet): void;
    worksheetFromIndex(index: number): Sheet | undefined;
    worksheetFromName(name: string): Sheet | undefined;
    saveToBuffer(): Promise<Buffer>;
    saveToBufferSync(): Buffer;
    saveToFileSync(path: string): void;
}
import Sheet = require("./sheet");
//# sourceMappingURL=workbook.d.ts.map