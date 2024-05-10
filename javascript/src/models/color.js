// @ts-check

/**
 * @class Color
 * @classdesc Represents a color
 * @property {number} red - The red value of the color
 * @property {number} green - The green value of the color
 * @property {number} blue - The blue value of the color
 */
class Color {
  /**
   * @param {Object} [opts = {}] - Options for the color
   * @param {number} [opts.red=0]
   * @param {number} [opts.green=0]
   * @param {number} [opts.blue=0]
   */
  constructor(opts = {}) {
    /**
     * The red value of the color
     * @type {number}
     */
    this.red = opts.red ?? 0;
    /**
     * The green value of the color
     * @type {number}
     */
    this.green = opts.green ?? 0;
    /**
     * The blue value of the color
     * @type {number}
     */
    this.blue = opts.blue ?? 0;
  }
}

module.exports = Color;
