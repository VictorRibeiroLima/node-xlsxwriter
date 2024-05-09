// @ts-check

const Color = require('./color');
const Format = require('./format');

//Enums
/**
 * @typedef {(
 * "automatic" |
 * "lowest" |
 * "number" |
 * "percent" |
 * "formula" |
 * "percentile" |
 * "highest"
 * )} ConditionalFormatEnumType
 *
 * @typedef {(
 * "aboveAverage" |
 * "belowAverage" |
 * "equalOrAboveAverage" |
 * "equalOrBelowAverage" |
 * "oneStandardDeviationAbove" |
 * "oneStandardDeviationBelow" |
 * "twoStandardDeviationsAbove" |
 * "twoStandardDeviationsBelow" |
 * "threeStandardDeviationsAbove" |
 * "threeStandardDeviationsBelow"
 * )} ConditionalFormatAverageRule
 *
 * @typedef {(
 * "automatic" |
 * "midpoint" |
 * "none"
 * )} ConditionalFormatDataBarAxisPosition
 *
 * @typedef {(
 * "context" |
 * "leftToRight" |
 * "rightToLeft"
 * )} ConditionalFormatDataBarDirection
 *
 * @typedef {(
 * "yesterday" |
 * "today" |
 * "tomorrow" |
 * "last7Days" |
 * "lastWeek" |
 * "thisWeek" |
 * "nextWeek" |
 * "lastMonth" |
 * "thisMonth" |
 * "nextMonth"
 * )} ConditionalFormatDateRule
 *
 * @typedef {(
 * "treeArrows" |
 * "threeArrowsGray" |
 * "threeFlags" |
 * "threeTrafficLights" |
 * "threeTrafficLightsWithRim" |
 * "threeSigns" |
 * "threeSymbolsCircled" |
 * "threeSymbols" |
 * "threeStars" |
 * "threeTriangles" |
 * "fourArrows" |
 * "fourArrowsGray" |
 * "fourRedToBlack" |
 * "fourHistograms" |
 * "fourTrafficLights" |
 * "fiveArrows" |
 * "fiveArrowsGray" |
 * "fiveHistograms" |
 * "fiveQuadrants" |
 * "fiveBoxes"
 * )} ConditionalFormatIconType
 */

//Values
/**
 *   @typedef {(string|number|Date)} ConditionalFormatValue
 */

//Rules
/**
 * @typedef {(
 * "equalTo" |
 * "notEqualTo" |
 * "greaterThan" |
 * "greaterThanOrEqualTo" |
 * "lessThan" |
 * "lessThanOrEqualTo" |
 * "between" |
 * "notBetween"
 * )} ConditionalFormatCellRuleType
 * @Class ConditionalFormatCellRule
 * @classdesc Represents a rule for a conditional format cell.
 * @property {ConditionalFormatCellRuleType} type - The type of the rule.
 * @property {ConditionalFormatValueValue} value - The value of the rule.
 */
class ConditionalFormatCellRule {
  /**
   * @param {ConditionalFormatCellRuleType} type
   * @param {ConditionalFormatValue} value
   */
  constructor(type, value) {
    /**
     * @type {ConditionalFormatCellRuleType}
     */
    this.type = type;
    /**
     * @type {ConditionalFormatValue}
     */
    this.value = value;
  }
}

/**
 * @typedef {(
 * "contains" |
 * "doesNotContain" |
 * "beginsWith" |
 * "endsWith"
 * )} ConditionalFormatTextRuleType
 *
 * @class ConditionalFormatTextRule
 * @classdesc Represents a rule for a conditional format text.
 * @property {ConditionalFormatTextRule} type - The type of the rule.
 * @property {string} value - The value of the rule.
 */

class ConditionalFormatTextRule {
  /**
   * @param {ConditionalFormatTextRuleType} type
   * @param {string} value
   */
  constructor(type, value) {
    /**
     * @type {ConditionalFormatTextRuleType}
     */
    this.type = type;
    /**
     * @type {string}
     */
    this.value = value;
  }
}

/**
 * @typedef {(
 * "top" |
 * "bottom" |
 * "topPercent" |
 * "bottomPercent"
 * )} ConditionalFormatTopRuleType
 *
 * @class ConditionalFormatTopRule
 * @classdesc Represents a rule for a conditional format top.
 * @property {ConditionalFormatTopRuleType} type - The type of the rule.
 * @property {number} value - The value of the rule.
 */
class ConditionalFormatTopRule {
  /**
   * @param {ConditionalFormatTopRuleType} type
   * @param {number} value
   */
  constructor(type, value) {
    /**
     * @type {ConditionalFormatTopRuleType}
     */
    this.type = type;
    /**
     * @type {number}
     */
    this.value = value;
  }
}

//Classes

/**
 * @typedef {(
 * "twoColorScale" |
 * "threeColorScale" |
 * "average" |
 * "blank" |
 * "cell" |
 * "dataBar" |
 * "date" |
 * "duplicate" |
 * "error" |
 * "formula" |
 * "iconSet" |
 * "text" |
 * "top"
 * )} ConditionalFormatClassType
 *
 * @class ConditionalFormat
 * @description Represents a generic conditional format.
 * @abstract
 * @property {number} id - The id of the conditional format.
 * @property {ConditionalFormatType} type - The type of the conditional format.
 * @property {string} [multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
 * @property {boolean} [stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
 */
class ConditionalFormat {
  /**
   * @param {ConditionalFormatClassType} type
   * @param {string} [multiRange]
   * @param {boolean} [stopIfTrue]
   */
  constructor(type, multiRange, stopIfTrue) {
    /**
     * @type {number}
     * @default Math.floor(Math.random() * 1000000)
     */
    this.id = Math.floor(Math.random() * 1000000);

    /**
     * @type {ConditionalFormatClassType}
     */
    this.type = type;

    /**
     * @type {string|undefined}
     * @default undefined
     */
    this.multiRange = multiRange;

    /**
     * @type {boolean|undefined}
     * @default undefined
     */
    this.stopIfTrue = stopIfTrue;
  }

  /**
   * @param {string} multiRange
   * @returns {void}
   */
  setMultiRange(multiRange) {
    this.multiRange = multiRange;
  }

  /**
   * @param {boolean} stopIfTrue
   * @returns {void}
   */
  setStopIfTrue(stopIfTrue) {
    this.stopIfTrue = stopIfTrue;
  }
}

/**
 * @typedef {Object} ConditionalFormatColorScaleRule
 * @property {ConditionalFormatEnumType} type - The type of the rule.
 * @property {ConditionalFormatValue} value - The value of the rule.
 * /

/**
 * @class ConditionalFormatTwoColorScale
 * @classdesc Represents a 2 Color Scale conditional format.
 * Used to represent a Cell style conditional format in Excel. A 2 Color Scale Cell conditional format shows a per cell color gradient from the minimum value to the maximum value.
 * @extends ConditionalFormat
 * @property {Color} [minColor] - The color for the minimum value.(If not set, Excel will use the default color for the minimum value)
 * @property {Color} [maxColor] - The color for the maximum value.(If not set, Excel will use the default color for the maximum value)
 * @property {ConditionalFormatTwoColorScaleRule} [minRule] - The rule for the minimum value.
 * @property {ConditionalFormatTwoColorScaleRule} [maxRule] - The rule for the maximum value.
 * @property {string} [multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
 * @property {boolean} [stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
 */
class ConditionalFormatTwoColorScale extends ConditionalFormat {
  /**
   * @param {Object} [options] - The options object
   * @param {Color} [options.minColor] - The color for the minimum value.
   * @param {Color} [options.maxColor] - The color for the maximum value.
   * @param {ConditionalFormatColorScaleRule} [options.minRule] - The rule for the minimum value.
   * @param {ConditionalFormatColorScaleRule} [options.maxRule] - The rule for the maximum value.
   * @param {string} [options.multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
   * @param {boolean} [options.stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
   */
  constructor(options = {}) {
    super('twoColorScale', options.multiRange, options.stopIfTrue);
    /**
     * @type {Color|undefined}
     * @default undefined
     */
    this.minColor = options.minColor;

    /**
     * @type {Color|undefined}
     * @default undefined
     */
    this.maxColor = options.maxColor;

    /**
     * @type {ConditionalFormatColorScaleRule|undefined}
     * @default undefined
     */
    this.minRule = options.minRule;

    /**
     * @type {ConditionalFormatColorScaleRule|undefined}
     * @default undefined
     */
    this.maxRule = options.maxRule;
  }

  /**
   * @param {Color} color
   */
  setMinColor(color) {
    this.minColor = color;
  }

  /**
   * @param {Color} color
   */
  setMaxColor(color) {
    this.maxColor = color;
  }

  /**
   * @param {ConditionalFormatColorScaleRule} rule
   */
  setMinRule(rule) {
    this.minRule = rule;
  }

  /**
   * @param {ConditionalFormatColorScaleRule} rule
   */
  setMaxRule(rule) {
    this.maxRule = rule;
  }

  /**
   * @param {string} multiRange
   */
  setMultiRange(multiRange) {
    this.multiRange = multiRange;
  }

  /**
   * @param {boolean} stopIfTrue
   */
  setStopIfTrue(stopIfTrue) {
    this.stopIfTrue = stopIfTrue;
  }
}

/**
 * @class ConditionalFormatTreeColorScale
 * @classdesc Represents a 3 Color Scale conditional format.
 * Used to represent a Cell style conditional format in Excel. A 3 Color Scale Cell conditional format shows a per cell color gradient from the minimum value to the maximum value.
 * @extends ConditionalFormatTwoColorScale
 * @property {Color} [midColor] - The color for the mid value.(If not set, Excel will use the default color for the mid value)
 * @property {ConditionalFormatTwoColorScaleRule} [midRule] - The rule for the mid value.
 */
class ConditionalFormatThreeColorScale extends ConditionalFormatTwoColorScale {
  /**
   * @param {Object} [options] - The options object
   * @param {Color} [options.minColor] - The color for the minimum value.
   * @param {Color} [options.midColor] - The color for the mid value.
   * @param {Color} [options.maxColor] - The color for the maximum value.
   * @param {ConditionalFormatColorScaleRule} [options.minRule] - The rule for the minimum value.
   * @param {ConditionalFormatColorScaleRule} [options.midRule] - The rule for the maximum value.
   * @param {ConditionalFormatColorScaleRule} [options.maxRule] - The rule for the maximum value.
   * @param {string} [options.multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
   * @param {boolean} [options.stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
   */
  constructor(options = {}) {
    super(options);
    /**
     * @type {ConditionalFormatClassType}
     */
    this.type = 'threeColorScale';

    /**
     * @type {Color|undefined}
     * @default undefined
     */
    this.midColor = options.midColor;

    /**
     * @type {ConditionalFormatColorScaleRule|undefined}
     * @default undefined
     */
    this.midRule = options.midRule;
  }

  /**
   * @param {Color} color
   */
  setMidColor(color) {
    this.midColor = color;
  }

  /**
   * @param {ConditionalFormatColorScaleRule} rule
   */
  setMidRule(rule) {
    this.midRule = rule;
  }
}

/**
 * @class ConditionalFormatAverage
 * @classdesc Represents an Average/Standard Deviation style conditional format
 * Is used to represent a Average or Standard Deviation style conditional format in Excel
 * @property {ConditionalFormatAverageRule} rule - The rule for the average value.(default: 'aboveAverage')
 * @property {Format} [format] - The format for the average value.
 * @property {string} [multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
 * @property {boolean} [stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
 */
class ConditionalFormatAverage extends ConditionalFormat {
  /**
   * @param {Object} [options] - The options object
   * @param {ConditionalFormatAverageRule} [options.rule] - The rule for the average value.(default: 'aboveAverage')
   * @param {Format} [options.format] - The format for the average value.
   * @param {string} [options.multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
   * @param {boolean} [options.stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
   */
  constructor(options = {}) {
    super('average', options.multiRange, options.stopIfTrue);
    /**
     * @type {ConditionalFormatAverageRule}
     * @default 'aboveAverage'
     */
    this.rule = options.rule || 'aboveAverage';

    /**
     * @type {Format|undefined}
     * @default undefined
     */
    this.format = options.format;
  }

  /**
   * @param {ConditionalFormatAverageRule} rule
   */
  setRule(rule) {
    this.rule = rule;
  }

  /**
   * @param {Format} format
   */
  setFormat(format) {
    this.format = format;
  }
}

/**
 * @class ConditionalFormatBlank
 * @classdesc Represents a a Blank/Non-blank conditional format.
 * @extends ConditionalFormat
 * @property {boolean} invert - Inverts the conditional format.
 * @property {Format} [format] - The format for the average value.
 * @property {string} [multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
 * @property {boolean} [stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
 */
class ConditionalFormatBlank extends ConditionalFormat {
  /**
   * @param {Object} [options] - The options object
   * @param {boolean} [options.invert] - Inverts the conditional format.
   * @param {Format} [options.format] - The format for the average value.
   * @param {string} [options.multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
   * @param {boolean} [options.stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
   */
  constructor(options = {}) {
    super('blank', options.multiRange, options.stopIfTrue);

    /**
     * @type {boolean}
     * @default false
     */
    this.invert = options.invert || false;

    /**
     * @type {Format|undefined}
     * @default undefined
     */

    this.format = options.format;
  }

  /**
   * @param {boolean} invert
   */
  setInvert(invert) {
    this.invert = invert;
  }

  /**
   * @param {Format} format
   */
  setFormat(format) {
    this.format = format;
  }
}

module.exports = {
  ConditionalFormat,
  ConditionalFormatCellRule,
  ConditionalFormatTextRule,
  ConditionalFormatTopRule,
  ConditionalFormatTwoColorScale,
  ConditionalFormatThreeColorScale,
  ConditionalFormatAverage,
  ConditionalFormatBlank,
};
