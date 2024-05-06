export = Cell;
declare class Cell {
    constructor(col: number, row: number, value: string | number, cellType: CellType);
    col: number;
    row: number;
    value: CellValue;
    cellType: CellType;
}
declare namespace Cell {
    export { CellValue, CellType };
}
type CellValue = (number | string);
type CellType = ("number" | "string" | "link");
//# sourceMappingURL=cell.d.ts.map