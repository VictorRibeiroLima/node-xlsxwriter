// @ts-check

const Color = require('./color');

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
 * "customIcon" |
 * "dataBar" |
 * "date" |
 * "duplicate" |
 * "error" |
 * "formula" |
 * "iconSet" |
 * "text" |
 * "top"|
 * "value"
 * )} ConditionalFormatClassType
 *
 * @class ConditionalFormat
 * @description Represents a generic conditional format.
 * @abstract
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
}

/**
 * @typedef {Object} ConditionalFormatTwoColorScaleRule
 * @property {ConditionalFormatEnumType} type - The type of the rule.
 * @property {ConditionalFormatValue} value - The value of the rule.
 *
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
   * @param {ConditionalFormatTwoColorScaleRule} [options.minRule] - The rule for the minimum value.
   * @param {ConditionalFormatTwoColorScaleRule} [options.maxRule] - The rule for the maximum value.
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
     * @type {ConditionalFormatTwoColorScaleRule|undefined}
     * @default undefined
     */
    this.minRule = options.minRule;

    /**
     * @type {ConditionalFormatTwoColorScaleRule|undefined}
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
   * @param {ConditionalFormatTwoColorScaleRule} rule
   */
  setMinRule(rule) {
    this.minRule = rule;
  }

  /**
   * @param {ConditionalFormatTwoColorScaleRule} rule
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

module.exports = {
  ConditionalFormat,
  ConditionalFormatCellRule,
  ConditionalFormatTextRule,
  ConditionalFormatTopRule,
  ConditionalFormatTwoColorScale,
};
