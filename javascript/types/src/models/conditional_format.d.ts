export type ConditionalFormatEnumType = ("automatic" | "lowest" | "number" | "percent" | "formula" | "percentile" | "highest");
export type ConditionalFormatAverageRule = ("aboveAverage" | "belowAverage" | "equalOrAboveAverage" | "equalOrBelowAverage" | "oneStandardDeviationAbove" | "oneStandardDeviationBelow" | "twoStandardDeviationsAbove" | "twoStandardDeviationsBelow" | "threeStandardDeviationsAbove" | "threeStandardDeviationsBelow");
export type ConditionalFormatDataBarAxisPosition = ("automatic" | "midpoint" | "none");
export type ConditionalFormatDataBarDirection = ("context" | "leftToRight" | "rightToLeft");
export type ConditionalFormatDateRule = ("yesterday" | "today" | "tomorrow" | "last7Days" | "lastWeek" | "thisWeek" | "nextWeek" | "lastMonth" | "thisMonth" | "nextMonth");
export type ConditionalFormatIconType = ("treeArrows" | "threeArrowsGray" | "threeFlags" | "threeTrafficLights" | "threeTrafficLightsWithRim" | "threeSigns" | "threeSymbolsCircled" | "threeSymbols" | "threeStars" | "threeTriangles" | "fourArrows" | "fourArrowsGray" | "fourRedToBlack" | "fourHistograms" | "fourTrafficLights" | "fiveArrows" | "fiveArrowsGray" | "fiveHistograms" | "fiveQuadrants" | "fiveBoxes");
export type ConditionalFormatValue = (string | number | Date);
export type ConditionalFormatCellRuleType = ("equalTo" | "notEqualTo" | "greaterThan" | "greaterThanOrEqualTo" | "lessThan" | "lessThanOrEqualTo" | "between" | "notBetween");
export type ConditionalFormatTextRuleType = ("contains" | "doesNotContain" | "beginsWith" | "endsWith");
export type ConditionalFormatTopRuleType = ("top" | "bottom" | "topPercent" | "bottomPercent");
export type ConditionalFormatClassType = ("twoColorScale" | "threeColorScale" | "average" | "blank" | "cell" | "dataBar" | "date" | "duplicate" | "error" | "formula" | "iconSet" | "text" | "top");
export type ConditionalFormatTwoColorScaleRule = {
    /**
     * - The type of the rule.
     */
    type: ConditionalFormatEnumType;
    /**
     * - The value of the rule.
     */
    value: ConditionalFormatValue;
};
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
export class ConditionalFormat {
    /**
     * @param {ConditionalFormatClassType} type
     * @param {string} [multiRange]
     * @param {boolean} [stopIfTrue]
     */
    constructor(type: ConditionalFormatClassType, multiRange?: string, stopIfTrue?: boolean);
    /**
     * @type {number}
     * @default Math.floor(Math.random() * 1000000)
     */
    id: number;
    /**
     * @type {ConditionalFormatClassType}
     */
    type: ConditionalFormatClassType;
    /**
     * @type {string|undefined}
     * @default undefined
     */
    multiRange: string | undefined;
    /**
     * @type {boolean|undefined}
     * @default undefined
     */
    stopIfTrue: boolean | undefined;
}
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
/**
 *   @typedef {(string|number|Date)} ConditionalFormatValue
 */
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
export class ConditionalFormatCellRule {
    /**
     * @param {ConditionalFormatCellRuleType} type
     * @param {ConditionalFormatValue} value
     */
    constructor(type: ConditionalFormatCellRuleType, value: ConditionalFormatValue);
    /**
     * @type {ConditionalFormatCellRuleType}
     */
    type: ConditionalFormatCellRuleType;
    /**
     * @type {ConditionalFormatValue}
     */
    value: ConditionalFormatValue;
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
export class ConditionalFormatTextRule {
    /**
     * @param {ConditionalFormatTextRuleType} type
     * @param {string} value
     */
    constructor(type: ConditionalFormatTextRuleType, value: string);
    /**
     * @type {ConditionalFormatTextRuleType}
     */
    type: ConditionalFormatTextRuleType;
    /**
     * @type {string}
     */
    value: string;
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
export class ConditionalFormatTopRule {
    /**
     * @param {ConditionalFormatTopRuleType} type
     * @param {number} value
     */
    constructor(type: ConditionalFormatTopRuleType, value: number);
    /**
     * @type {ConditionalFormatTopRuleType}
     */
    type: ConditionalFormatTopRuleType;
    /**
     * @type {number}
     */
    value: number;
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
export class ConditionalFormatTwoColorScale extends ConditionalFormat {
    /**
     * @param {Object} [options] - The options object
     * @param {Color} [options.minColor] - The color for the minimum value.
     * @param {Color} [options.maxColor] - The color for the maximum value.
     * @param {ConditionalFormatTwoColorScaleRule} [options.minRule] - The rule for the minimum value.
     * @param {ConditionalFormatTwoColorScaleRule} [options.maxRule] - The rule for the maximum value.
     * @param {string} [options.multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
     * @param {boolean} [options.stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
     */
    constructor(options?: {
        minColor?: Color;
        maxColor?: Color;
        minRule?: ConditionalFormatTwoColorScaleRule;
        maxRule?: ConditionalFormatTwoColorScaleRule;
        multiRange?: string;
        stopIfTrue?: boolean;
    });
    /**
     * @type {Color|undefined}
     * @default undefined
     */
    minColor: Color | undefined;
    /**
     * @type {Color|undefined}
     * @default undefined
     */
    maxColor: Color | undefined;
    /**
     * @type {ConditionalFormatTwoColorScaleRule|undefined}
     * @default undefined
     */
    minRule: ConditionalFormatTwoColorScaleRule | undefined;
    /**
     * @type {ConditionalFormatTwoColorScaleRule|undefined}
     * @default undefined
     */
    maxRule: ConditionalFormatTwoColorScaleRule | undefined;
    /**
     * @param {Color} color
     */
    setMinColor(color: Color): void;
    /**
     * @param {Color} color
     */
    setMaxColor(color: Color): void;
    /**
     * @param {ConditionalFormatTwoColorScaleRule} rule
     */
    setMinRule(rule: ConditionalFormatTwoColorScaleRule): void;
    /**
     * @param {ConditionalFormatTwoColorScaleRule} rule
     */
    setMaxRule(rule: ConditionalFormatTwoColorScaleRule): void;
    /**
     * @param {string} multiRange
     */
    setMultiRange(multiRange: string): void;
    /**
     * @param {boolean} stopIfTrue
     */
    setStopIfTrue(stopIfTrue: boolean): void;
}
import Color = require("./color");
//# sourceMappingURL=conditional_format.d.ts.map