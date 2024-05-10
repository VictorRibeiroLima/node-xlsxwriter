// @ts-check
const Color = require('./models/color');

const colorUtils = {
  /**
   * Converts a hexadecimal color string to an RGB color object.
   *
   * @param {string} hex - The hexadecimal color string.
   * @returns {Color} the color
   * @throws {Error} If the input is not a valid hexadecimal color string.
   */
  hexToColor: function (hex) {
    if (!/^#?[0-9A-F]{6}$/i.test(hex)) {
      throw new Error('Invalid hexadecimal color string');
    }

    // Remove the leading '#' if it exists
    hex = hex.indexOf('#') === 0 ? hex.slice(1) : hex;

    const red = parseInt(hex.slice(0, 2), 16);
    const green = parseInt(hex.slice(2, 4), 16);
    const blue = parseInt(hex.slice(4, 6), 16);

    return new Color({
      red,
      green,
      blue,
    });
  },
};

module.exports = {
  colorUtils,
};
