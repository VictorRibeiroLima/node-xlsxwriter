export = Color;
/**
 * @class Color
 * @classdesc Represents a color
 * @property {number} red - The red value of the color
 * @property {number} green - The green value of the color
 * @property {number} blue - The blue value of the color
 */
declare class Color {
    /**
     * @param {Object} [opts = {}] - Options for the color
     * @param {number} [opts.red=0]
     * @param {number} [opts.green=0]
     * @param {number} [opts.blue=0]
     */
    constructor(opts?: {
        red?: number;
        green?: number;
        blue?: number;
    });
    /**
     * The red value of the color
     * @type {number}
     */
    red: number;
    /**
     * The green value of the color
     * @type {number}
     */
    green: number;
    /**
     * The blue value of the color
     * @type {number}
     */
    blue: number;
}
//# sourceMappingURL=color.d.ts.map