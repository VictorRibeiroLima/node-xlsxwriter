export = Sheet;
declare class Sheet {
    constructor(name: string);
    name: string;
    cells: Cell[];
    writeCell(col: number, row: number, value: string | number, cellType: ("number" | "string" | "link")): void;
    writeString(col: number, row: number, value: string): void;
    writeNumber(col: number, row: number, value: number): void;
    writeLink(col: number, row: number, value: string): void;
}
import Cell = require("./cell");
//# sourceMappingURL=sheet.d.ts.map