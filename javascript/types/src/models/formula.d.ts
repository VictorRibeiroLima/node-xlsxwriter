export = Formula;
/**
 * @class Formula
 * @classdesc Represents a worksheet formula.
 * @property {string} formula - The formula string.
 * @property {string} [result] - The result of the formula.
 * @property {boolean} [useFutureFunctions=false] - Enable the use of newer Excel future functions in the formula
 * @property {boolean} [useTableFunctions=false] - Enable backward compatible formulas in table.
 * @property {boolean} [dynamic=false] - Enable the use of dynamic arrays in the formula
 */
declare class Formula {
    /**
     * @param {Object} opts - Options for the formula.
     * @param {string} opts.formula - The formula string.
     * @param {string} [opts.result] - The result of the formula.
     * @param {boolean} [opts.useFutureFunctions=false] - Enable the use of newer Excel future functions in the formula
     * @param {boolean} [opts.useTableFunctions=false] - Enable backward compatible formulas in table.
     * @param {boolean} [opts.dynamic=false] - Enable the use of dynamic arrays in the formula
     */
    constructor(opts: {
        formula: string;
        result?: string;
        useFutureFunctions?: boolean;
        useTableFunctions?: boolean;
        dynamic?: boolean;
    });
    /**
     * The formula string.
     * @type {string}
     */
    formula: string;
    /**
     * The result of the formula.
     * @type {?string|undefined}
     */
    result: (string | undefined) | null;
    /**
     * Enable the use of newer Excel future functions in the formula
     * @type {boolean}
     */
    useFutureFunctions: boolean;
    /**
     * Enable backward compatible formulas in table.
     * @type {boolean}
     */
    useTableFunctions: boolean;
    /**
     * Enable the use of dynamic arrays in the formula
     * @type {boolean}
     */
    dynamic: boolean;
    /**
     * Set the result of the formula.
     * @param {string} result - The result of the formula.
     */
    setResult(result: string): void;
    /**
     * Set the formula string.
     * @param {string} formula - The formula string.
     */
    setFormula(formula: string): void;
    /**
     * Set the useFutureFunctions flag.
     * @param {boolean} useFutureFunctions - Enable the use of newer Excel future functions in the formula
     */
    setUseFutureFunctions(useFutureFunctions: boolean): void;
    /**
     * Set the useTableFunctions flag.
     * @param {boolean} useTableFunctions - Enable backward compatible formulas in table.
     */
    setUseTableFunctions(useTableFunctions: boolean): void;
}
//# sourceMappingURL=formula.d.ts.map