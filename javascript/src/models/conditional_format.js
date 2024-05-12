// @ts-check

const Color = require('./color');
const Format = require('./format');
const Formula = require('./formula');

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
 * "threeArrows" |
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
 * @typedef {Object} ConditionalFormatCellRule
 * @property {ConditionalFormatCellRuleType} type - The type of the rule.
 * @property {ConditionalFormatValue} value - The value of the rule.
 * @property {ConditionalFormatValue} [optionalValue] - The optional value of the rule (between, notBetween).
 */

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
 * @typedef {Object} ConditionalFormatTypeRule
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
   * @param {ConditionalFormatTypeRule} [options.minRule] - The rule for the minimum value.
   * @param {ConditionalFormatTypeRule} [options.maxRule] - The rule for the maximum value.
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
     * @type {ConditionalFormatTypeRule|undefined}
     * @default undefined
     */
    this.minRule = options.minRule;

    /**
     * @type {ConditionalFormatTypeRule|undefined}
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
   * @param {ConditionalFormatTypeRule} rule
   */
  setMinRule(rule) {
    this.minRule = rule;
  }

  /**
   * @param {ConditionalFormatTypeRule} rule
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
   * @param {ConditionalFormatTypeRule} [options.minRule] - The rule for the minimum value.
   * @param {ConditionalFormatTypeRule} [options.midRule] - The rule for the maximum value.
   * @param {ConditionalFormatTypeRule} [options.maxRule] - The rule for the maximum value.
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
     * @type {ConditionalFormatTypeRule|undefined}
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
   * @param {ConditionalFormatTypeRule} rule
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

/**
 * @class ConditionalFormatCell
 * @classdesc Represents a cell style conditional format.
 * @extends ConditionalFormat
 * @property {ConditionalFormatCellRule} rule - The rule for the cell.
 * @property {Format} [format] - The format for the cell.
 * @property {string} [multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
 * @property {boolean} [stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
 */
class ConditionalFormatCell extends ConditionalFormat {
  /**
   * @param {Object} [options] - The options object
   * @param {ConditionalFormatCellRule} [options.rule] - The rule for the cell.
   * @param {Format} [options.format] - The format for the cell.
   * @param {string} [options.multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
   * @param {boolean} [options.stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
   */
  constructor(options = {}) {
    super('cell', options.multiRange, options.stopIfTrue);
    /**
     * @type {ConditionalFormatCellRule}
     */
    this.rule = options.rule;

    /**
     * @type {Format|undefined}
     * @default undefined
     */
    this.format = options.format;
  }

  /**
   * @param {ConditionalFormatCellRule} rule
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
 * @class ConditionalFormatDataBar
 * @classdesc Represents a Data Bar style conditional format.
 * @extends ConditionalFormat
 * @property {Color} [axisColor] - The color of the axis.
 * @property {ConditionalFormatDataBarAxisPosition} [axisPosition] - The position of the axis.
 * @property {boolean} [barOnly] - Show only the bar.
 * @property {Color} [borderColor] - The color of the border.
 * @property {boolean} [borderOff] - Turn off the border.
 * @property {ConditionalFormatDataBarDirection} [direction] - The direction of the data bar.
 * @property {Color} [fillColor] - The color of the fill.
 * @property {ConditionalFormatTypeRule} [maxRule] - The rule for the maximum value.
 * @property {ConditionalFormatTypeRule} [minRule] - The rule for the minimum value.
 * @property {Color} [negativeBorderColor] - The color of the negative border.
 * @property {Color} [negativeFillColor] - The color of the negative fill.
 * @property {boolean} [solidFill] - Show a solid fill.
 * @property {string} [multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
 * @property {boolean} [stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
 */
class ConditionalFormatDataBar extends ConditionalFormat {
  /**
   * @param {Object} [options] - The options object
   * @param {Color} [options.axisColor] - The color of the axis.
   * @param {ConditionalFormatDataBarAxisPosition} [options.axisPosition] - The position of the axis.
   * @param {boolean} [options.barOnly] - Show only the bar.
   * @param {Color} [options.borderColor] - The color of the border.
   * @param {boolean} [options.borderOff] - Turn off the border.
   * @param {ConditionalFormatDataBarDirection} [options.direction] - The direction of the data bar.
   * @param {Color} [options.fillColor] - The color of the fill.
   * @param {ConditionalFormatTypeRule} [options.maxRule] - The rule for the maximum value.
   * @param {ConditionalFormatTypeRule} [options.minRule] - The rule for the minimum value.
   * @param {Color} [options.negativeBorderColor] - The color of the negative border.
   * @param {Color} [options.negativeFillColor] - The color of the negative fill.
   * @param {boolean} [options.solidFill] - Show a solid fill.
   * @param {string} [options.multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
   * @param {boolean} [options.stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
   */
  constructor(options = {}) {
    super('dataBar', options.multiRange, options.stopIfTrue);
    /**
     * @type {Color|undefined}
     * @default undefined
     */
    this.axisColor = options.axisColor;

    /**
     * @type {ConditionalFormatDataBarAxisPosition|undefined}
     * @default undefined
     */
    this.axisPosition = options.axisPosition;

    /**
     * @type {boolean|undefined}
     * @default undefined
     */
    this.barOnly = options.barOnly;

    /**
     * @type {Color|undefined}
     * @default undefined
     */
    this.borderColor = options.borderColor;

    /**
     * @type {boolean|undefined}
     * @default undefined
     */
    this.borderOff = options.borderOff;

    /**
     * @type {ConditionalFormatDataBarDirection|undefined}
     * @default undefined
     */
    this.direction = options.direction;

    /**
     * @type {Color|undefined}
     * @default undefined
     */
    this.fillColor = options.fillColor;

    /**
     * @type {ConditionalFormatTypeRule|undefined}
     * @default undefined
     */
    this.maxRule = options.maxRule;

    /**
     * @type {ConditionalFormatTypeRule|undefined}
     * @default undefined
     */
    this.minRule = options.minRule;

    /**
     * @type {Color|undefined}
     * @default undefined
     */
    this.negativeBorderColor = options.negativeBorderColor;

    /**
     * @type {Color|undefined}
     * @default undefined
     */
    this.negativeFillColor = options.negativeFillColor;

    /**
     * @type {boolean|undefined}
     * @default undefined
     */
    this.solidFill = options.solidFill;
  }

  /**
   * @param {Color} color
   */
  setAxisColor(color) {
    this.axisColor = color;
  }

  /**
   * @param {ConditionalFormatDataBarAxisPosition} position
   */
  setAxisPosition(position) {
    this.axisPosition = position;
  }

  /**
   * @param {boolean} barOnly
   */
  setBarOnly(barOnly) {
    this.barOnly = barOnly;
  }

  /**
   * @param {Color} color
   */
  setBorderColor(color) {
    this.borderColor = color;
  }

  /**
   * @param {boolean} borderOff
   */
  setBorderOff(borderOff) {
    this.borderOff = borderOff;
  }

  /**
   * @param {ConditionalFormatDataBarDirection} direction
   */
  setDirection(direction) {
    this.direction = direction;
  }

  /**
   * @param {Color} color
   */
  setFillColor(color) {
    this.fillColor = color;
  }

  /**
   * @param {ConditionalFormatTypeRule} rule
   */
  setMaxRule(rule) {
    this.maxRule = rule;
  }

  /**
   * @param {ConditionalFormatTypeRule} rule
   */
  setMinRule(rule) {
    this.minRule = rule;
  }

  /**
   * @param {Color} color
   */
  setNegativeBorderColor(color) {
    this.negativeBorderColor = color;
  }

  /**
   * @param {Color} color
   */
  setNegativeFillColor(color) {
    this.negativeFillColor = color;
  }

  /**
   * @param {boolean} solidFill
   */
  setSolidFill(solidFill) {
    this.solidFill = solidFill;
  }
}

/**
 * @class ConditionalFormatDate
 * @classdesc Represents a Date style conditional format.
 * @extends ConditionalFormat
 * @property {Format} [format] - The format for the date.
 * @property {ConditionalFormatDateRule} [rule] - The rule for the date.
 * @property {string} [multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
 * @property {boolean} [stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
 */
class ConditionalFormatDate extends ConditionalFormat {
  /**
   * @param {Object} [options] - The options object
   * @param {Format} [options.format] - The format for the date.
   * @param {ConditionalFormatDateRule} [options.rule] - The rule for the date.
   * @param {string} [options.multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
   * @param {boolean} [options.stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
   */
  constructor(options = {}) {
    super('date', options.multiRange, options.stopIfTrue);
    /**
     * @type {Format|undefined}
     * @default undefined
     */
    this.format = options.format;

    /**
     * @type {ConditionalFormatDateRule|undefined}
     */
    this.rule = options.rule;
  }

  /**
   * @param {Format} format
   */
  setFormat(format) {
    this.format = format;
  }

  /**
   * @param {ConditionalFormatDateRule} rule
   */
  setRule(rule) {
    this.rule = rule;
  }
}

/**
 * @class ConditionalFormatDuplicate
 * @classdesc Represents a Duplicate/Unique conditional format.
 * @extends ConditionalFormat
 * @property {boolean} invert - Inverts the conditional format.
 * @property {Format} [format] - The format for the average value.
 * @property {string} [multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
 * @property {boolean} [stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
 */
class ConditionalFormatDuplicate extends ConditionalFormat {
  /**
   * @param {Object} [options] - The options object
   * @param {boolean} [options.invert] - Inverts the conditional format.
   * @param {Format} [options.format] - The format for the average value.
   * @param {string} [options.multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
   * @param {boolean} [options.stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
   */
  constructor(options = {}) {
    super('duplicate', options.multiRange, options.stopIfTrue);
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

/**
 * @class ConditionalFormatError
 * @classdesc Represents an Error style conditional format.
 * @extends ConditionalFormat
 * @property {boolean} invert - Inverts the conditional format.
 * @property {Format} [format] - The format for the average value.
 * @property {string} [multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
 * @property {boolean} [stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
 */
class ConditionalFormatError extends ConditionalFormat {
  /**
   * @param {Object} [options] - The options object
   * @param {boolean} [options.invert] - Inverts the conditional format.
   * @param {Format} [options.format] - The format for the average value.
   * @param {string} [options.multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
   * @param {boolean} [options.stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
   */
  constructor(options = {}) {
    super('error', options.multiRange, options.stopIfTrue);

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
   * Sets the format
   * @param {Format} format
   */
  setFormat(format) {
    this.format = format;
  }

  /**
   * Sets the invert flag
   * @param {boolean} invert
   */
  setInvert(invert) {
    this.invert = invert;
  }
}

/**
 * @class ConditionalFormatFormula
 * @classdesc Represents a Formula style conditional format.
 * @extends ConditionalFormat
 * @property {Formula} [formula] - The formula(non-dynamic) for the conditional format.
 * @property {Format} [format] - The format for the average value.
 * @property {string} [multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
 * @property {boolean} [stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
 */
class ConditionalFormatFormula extends ConditionalFormat {
  /**
   * @param {Object} [options] - The options object
   * @param {Formula} [options.formula] - The formula(non-dynamic) for the conditional format.
   * @param {Format} [options.format] - The format for the average value.
   * @param {string} [options.multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
   * @param {boolean} [options.stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
   */
  constructor(options = {}) {
    super('formula', options.multiRange, options.stopIfTrue);
    /**
     * @type {Formula|undefined}
     * @default undefined
     */
    this.formula = options.formula;

    /**
     * @type {Format|undefined}
     * @default undefined
     */
    this.format = options.format;
  }

  /**
   * @param {Formula} formula
   */
  setFormula(formula) {
    this.formula = formula;
  }

  /**
   * @param {Format} format
   */
  setFormat(format) {
    this.format = format;
  }
}

/**
 * @class ConditionalFormatCustomIcon
 * @classdesc Represents an icon in an Icon Set style conditional format
 * @property {boolean} [greaterThan = false] - Set the rule to be “greater than” instead of the Excel default of “greater than or equal to”.
 * @property {ConditionalFormatTypeRule} iconRule - The rule for the icon.
 * @property {boolean} [noIcon = false] - Turn off the icon in the cell.
 * @property {ConditionalFormatIconType} [iconType] - The type of icon.
 * @property {number} [iconTypeIndex] - Index to the icon within the type.
 */
class ConditionalFormatCustomIcon {
  /**
   * @param {Object} options - The options object
   * @param {ConditionalFormatTypeRule} options.iconRule - The rule for the icon.
   * @param {boolean} [options.greaterThan] - Set the rule to be “greater than” instead of the Excel default of “greater than or equal to”.
   * @param {boolean} [options.noIcon] - Turn off the icon in the cell.
   * @param {Object} [options.iconType] - The type of icon.
   * @param {number} options.iconType.index
   * @param {ConditionalFormatIconType} options.iconType.type - The rule for the icon.
   */
  constructor(options) {
    /**
     * @type {boolean}
     * @default false
     */
    this.greaterThan = options.greaterThan || false;

    /**
     * @type {boolean}
     * @default false
     */
    this.noIcon = options.noIcon || false;

    /**
     * @type {ConditionalFormatIconType|undefined}
     * @default undefined
     */
    this.iconType = options.iconType?.type;

    /**
     * @type {number|undefined}
     * @default undefined
     */
    this.iconTypeIndex = options.iconType?.index;

    /**
     * @type {ConditionalFormatTypeRule}
     */
    this.iconRule = options.iconRule;
  }

  /**
   * @param {boolean} greaterThan
   */
  setGreaterThan(greaterThan) {
    this.greaterThan = greaterThan;
  }

  /**
   * @param {boolean} noIcon
   */
  setNoIcon(noIcon) {
    this.noIcon = noIcon;
  }

  /**
   * @param {ConditionalFormatIconType} iconType
   * @param {number} index
   */
  setIconType(iconType, index) {
    this.iconType = iconType;
    this.iconTypeIndex = index;
  }
}

/**
 * @class ConditionalFormatIconSet
 * @classdesc Represents a Icon Set style conditional format.
 * @extends ConditionalFormat
 * @property {boolean} [reverse=false] - Reverse the order of icons from lowest to highest.
 * @property {boolean} [showIconsOnly=false] - Show only the icons and not the data in the cells.
 * @property {ConditionalFormatCustomIcon[]} icons - The icons for the set.
 * @property {ConditionalFormatIconType} iconType - The type of icon set.
 * @property {string} [multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
 * @property {boolean} [stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
 */
class ConditionalFormatIconSet extends ConditionalFormat {
  /**
   * @param {Object} options - The options object
   * @param {ConditionalFormatCustomIcon[]} options.icons - The icons for the set.
   * @param {ConditionalFormatIconType} options.iconType - The type of icon set.
   * @param {boolean} [options.reverse] - Reverse the order of icons from lowest to highest.
   * @param {boolean} [options.showIconsOnly] - Show only the icons and not the data in the cells.
   * @param {string} [options.multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
   * @param {boolean} [options.stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
   * @throws {Error} When the number of icons does not match the expected number for the icon type.
   */
  constructor(options) {
    super('iconSet', options.multiRange, options.stopIfTrue);

    let expectedIcons = 0;
    if (options.iconType.indexOf('three') > -1) {
      expectedIcons = 3;
    } else if (options.iconType.indexOf('four') > -1) {
      expectedIcons = 4;
    } else if (options.iconType.indexOf('five') > -1) {
      expectedIcons = 5;
    }

    if (options.icons.length !== expectedIcons) {
      throw new Error(
        `Expected ${expectedIcons} icons, but got ${options.icons.length}`,
      );
    }

    /**
     * @type {boolean}
     * @default false
     */
    this.reverse = options.reverse || false;

    /**
     * @type {boolean}
     * @default false
     */
    this.showIconsOnly = options.showIconsOnly || false;

    /**
     * @type {ConditionalFormatCustomIcon[]}
     */
    this.icons = options.icons;

    /**
     * @type {ConditionalFormatIconType}
     */
    this.iconType = options.iconType;
  }

  /**
   * @param {boolean} reverse
   */
  setReverse(reverse) {
    this.reverse = reverse;
  }

  /**
   * @param {boolean} showIconsOnly
   */
  setShowIconsOnly(showIconsOnly) {
    this.showIconsOnly = showIconsOnly;
  }

  /**
   * @param {ConditionalFormatIconType} iconType
   * @param {ConditionalFormatCustomIcon[]} icons
   * @throws {Error} When the number of icons does not match the expected number for the icon type.
   */
  setIcons(iconType, icons) {
    let expectedIcons = 0;
    if (iconType.indexOf('three') > -1) {
      expectedIcons = 3;
    } else if (iconType.indexOf('four') > -1) {
      expectedIcons = 4;
    } else if (iconType.indexOf('five') > -1) {
      expectedIcons = 5;
    }

    if (icons.length !== expectedIcons) {
      throw new Error(
        `Expected ${expectedIcons} icons, but got ${icons.length}`,
      );
    }
    this.iconType = iconType;
    this.icons = icons;
  }
}

module.exports = {
  ConditionalFormat,
  ConditionalFormatTextRule,
  ConditionalFormatTopRule,
  ConditionalFormatTwoColorScale,
  ConditionalFormatThreeColorScale,
  ConditionalFormatAverage,
  ConditionalFormatBlank,
  ConditionalFormatCell,
  ConditionalFormatDataBar,
  ConditionalFormatDate,
  ConditionalFormatDuplicate,
  ConditionalFormatError,
  ConditionalFormatFormula,
  ConditionalFormatCustomIcon,
  ConditionalFormatIconSet,
};
