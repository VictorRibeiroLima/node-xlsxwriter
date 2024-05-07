/**
 * - The border style
 */
export type BorderType = ("none" | "thin" | "medium" | "thick" | "double" | "hair" | "mediumDashed" | "dashDot" | "mediumDashDot" | "dashDotDot" | "mediumDashDotDot" | "slantDashDot");
/**
 * - The diagonal border style
 */
export type DiagonalBorderType = ("none" | "borderUp" | "borderDown" | "borderUpDown");
/**
 *@typedef {(
 *    "none"|
 *    "thin"|
 *    "medium"|
 *    "thick"|
 *    "double"|
 *    "hair"|
 *    "mediumDashed"|
 *    "dashDot"|
 *    "mediumDashDot"|
 *    "dashDotDot"|
 *    "mediumDashDotDot"|
 *    "slantDashDot"
 * )} BorderType - The border style
 *
 * @typedef {(
 *    "none"|
 *    "borderUp"|
 *    "borderDown"|
 *    "borderUpDown"
 * )} DiagonalBorderType - The diagonal border style
 *
 */
/**
 * @class Border
 * @classdesc Represents a border
 * @property {BorderType} style - The style of the border
 * @property {Color} color - The color of the border
 */
export class Border {
    /**
     * @param {BorderType} [style] - The style of the border
     * @param {Color} [color] - The color of the border
     */
    constructor(style?: BorderType, color?: Color);
    /**
     * The style of the border
     * @type {BorderType}
     */
    style: BorderType;
    /**
     * The color of the border
     * @type {Color}
     */
    color: Color;
}
/**
 * @class DiagonalBorder
 * @classdesc Represents a diagonal border
 * @property {DiagonalBorderType} dStyle - The style of the border
 * @property {BorderType} style - The style of the border
 * @property {Color} color - The color of the border
 */
export class DiagonalBorder extends Border {
    /**
     * @param {BorderType} [style] - The style of the border
     * @param {Color} [color] - The color of the border
     * @param {DiagonalBorderType} [dStyle] - The style of the border
     */
    constructor(style?: BorderType, color?: Color, dStyle?: DiagonalBorderType);
    /**
     * The style of the border
     * @type {DiagonalBorderType}
     */
    dStyle: DiagonalBorderType;
}
import Color = require("./color");
//# sourceMappingURL=border.d.ts.map