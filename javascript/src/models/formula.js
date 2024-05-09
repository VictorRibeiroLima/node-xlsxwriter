// @ts-check

/**
 * @class Formula
 * @classdesc Represents a worksheet formula.
 * @property {string} formula - The formula string.
 * @property {string} [result] - The result of the formula.
 * @property {boolean} [useFutureFunctions=false] - Enable the use of newer Excel future functions in the formula
 * @property {boolean} [useTableFunctions=false] - Enable backward compatible formulas in table.
 * @property {boolean} [dynamic=false] - Enable the use of dynamic arrays in the formula
 */
class Formula {
  /**
   * @param {Object} opts - Options for the formula.
   * @param {string} opts.formula - The formula string.
   * @param {string} [opts.result] - The result of the formula.
   * @param {boolean} [opts.useFutureFunctions=false] - Enable the use of newer Excel future functions in the formula
   * @param {boolean} [opts.useTableFunctions=false] - Enable backward compatible formulas in table.
   * @param {boolean} [opts.dynamic=false] - Enable the use of dynamic arrays in the formula
   */
  constructor(opts) {
    /**
     * The formula string.
     * @type {string}
     */
    this.formula = opts.formula;
    /**
     * The result of the formula.
     * @type {?string|undefined}
     */
    this.result = opts.result;
    /**
     * Enable the use of newer Excel future functions in the formula
     * @type {boolean}
     */
    this.useFutureFunctions = opts.useFutureFunctions || false;
    /**
     * Enable backward compatible formulas in table.
     * @type {boolean}
     */
    this.useTableFunctions = opts.useTableFunctions || false;

    /**
     * Enable the use of dynamic arrays in the formula
     * @type {boolean}
     */
    this.dynamic = opts.dynamic || false;
  }

  /**
   * Set the result of the formula.
   * @param {string} result - The result of the formula.
   */
  setResult(result) {
    this.result = result;
  }

  /**
   * Set the formula string.
   * @param {string} formula - The formula string.
   */
  setFormula(formula) {
    this.formula = formula;
  }

  /**
   * Set the useFutureFunctions flag.
   * @param {boolean} useFutureFunctions - Enable the use of newer Excel future functions in the formula
   */
  setUseFutureFunctions(useFutureFunctions) {
    this.useFutureFunctions = useFutureFunctions;
  }

  /**
   * Set the useTableFunctions flag.
   * @param {boolean} useTableFunctions - Enable backward compatible formulas in table.
   */
  setUseTableFunctions(useTableFunctions) {
    this.useTableFunctions = useTableFunctions;
  }
}

module.exports = Formula;
