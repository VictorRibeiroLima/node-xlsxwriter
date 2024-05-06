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
   * @param {number} red
   * @param {number} green
   * @param {number} blue
   */
  constructor(red, green, blue) {
    /**
     * The red value of the color
     * @type {number}
     */
    this.red = red;
    /**
     * The green value of the color
     * @type {number}
     */
    this.green = green;
    /**
     * The blue value of the color
     * @type {number}
     */
    this.blue = blue;
  }
}

module.exports = Color;
