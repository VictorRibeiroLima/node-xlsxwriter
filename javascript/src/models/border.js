// @ts-check
const Color = require('./color');

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
class Border {
  /**
   * @param {BorderType} [style] - The style of the border
   * @param {Color} [color] - The color of the border
   */
  constructor(style, color) {
    if (!color) {
      color = new Color(0, 0, 0);
    }
    if (!style) {
      style = 'none';
    }
    /**
     * The style of the border
     * @type {BorderType}
     */
    this.style = style;
    /**
     * The color of the border
     * @type {Color}
     */
    this.color = color;
  }
}

/**
 * @class DiagonalBorder
 * @classdesc Represents a diagonal border
 * @property {DiagonalBorderType} dStyle - The style of the border
 * @property {BorderType} style - The style of the border
 * @property {Color} color - The color of the border
 */
class DiagonalBorder extends Border {
  /**
   * @param {BorderType} [style] - The style of the border
   * @param {Color} [color] - The color of the border
   * @param {DiagonalBorderType} [dStyle] - The style of the border
   */
  constructor(style, color, dStyle) {
    super(style, color);
    /**
     * The style of the border
     * @type {DiagonalBorderType}
     */
    this.dStyle = dStyle;
  }
}

module.exports = { Border, DiagonalBorder };
