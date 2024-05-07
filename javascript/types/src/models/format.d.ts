export = Format;
/**
 * @typedef {(
 *   "general"|
 *   "left"|
 *   "center"|
 *   "right"|
 *   "fill"|
 *   "justify"|
 *   "centerAcross"|
 *   "distributed"|
 *   "top"|
 *   "bottom"|
 *   "verticalCenter"|
 *   "verticalJustify"|
 *   "verticalDistributed"
 * )} FormatAlign - The alignment of the cell
 *
 * @typedef {(
 *   "none" |
 *   "solid" |
 *   "mediumGray" |
 *   "darkGray" |
 *   "lightGray" |
 *   "darkHorizontal" |
 *   "darkVertical" |
 *   "darkDown" |
 *   "darkUp" |
 *   "darkGrid" |
 *   "darkTrellis" |
 *   "lightHorizontal" |
 *   "lightVertical" |
 *   "lightDown" |
 *   "lightUp" |
 *   "lightGrid" |
 *   "lightTrellis" |
 *   "gray125" |
 *   "gray0625"
 * )} FormatPattern - The pattern of the cell
 *
 * @typedef {(
 *   "none"|
 *   "single"|
 *   "double"|
 *   "singleAccounting"|
 *   "doubleAccounting"
 * )} FormatUnderline - The underline style
 */
/**
 * @class Format
 * @classdesc Represents a format
 * @property {number} id - the id of the format (do not set this directly)
 * @property {FormatAlign} [align] - The alignment of the cell
 * @property {Color} [backgroundColor] - The background color of the cell
 * @property {boolean} [bold] - If the font is bold
 * @property {Border} [leftBorder] - The left border of the cell
 * @property {Border} [rightBorder] - The right border of the cell
 * @property {Border} [topBorder] - The top border of the cell
 * @property {Border} [bottomBorder] - The bottom border of the cell
 * @property {DiagonalBorder} [diagonalBorder] - The diagonal border of the cell
 * @property {number} [charset] - The charset of the font
 * @property {Color} [fontColor] - The color of the font
 * @property {number} [fontFamily] - The family of the font
 * @property {string} [fontName] - The name of the font
 * @property {string} [fontScheme] - The font scheme
 * @property {number} [fontSize] - The font size
 * @property {boolean} [strikeThrough] - If the font is strike through
 * @property {Color} [foregroundColor] - The foreground color
 * @property {boolean} [hidden] - If the format is hidden
 * @property {boolean} [hyperlink] - If the format is hyperlinked
 * @property {number} [indent] - The indent level
 * @property {boolean} [italic] - If the font is italic
 * @property {boolean} [locked] - If the format is locked
 * @property {string} [numFmt] - The number format
 * @property {number} [numFmtId] - The number format id
 * @property {FormatPattern} [pattern] - The pattern
 * @property {FormatUnderline} [underline] - The underline style
 */
declare class Format {
    /**
     * the id of the format
     * @private
     * @type {number}
     */
    private id;
    /**
     * Sets alignment of the cell
     * @param {FormatAlign} align - The alignment of the cell
     * @returns {void}
     */
    setAlignment(align: FormatAlign): void;
    align: FormatAlign;
    /**
     * Sets the font color
     * @param {Color} color - The color of the font
     * @returns {void}
     */
    setBackgroundColor(color: Color): void;
    backgroundColor: Color;
    /**
     * Sets if the font is bold
     * @param {boolean} bold - If the font is bold
     * @returns {void}
     */
    setBold(bold: boolean): void;
    bold: boolean;
    /**
     * Sets the border of the cell
     * @param {Border} border - The border of the cell
     * @returns {void}
     */
    setBorder(border: Border): void;
    leftBorder: Border;
    rightBorder: Border;
    topBorder: Border;
    bottomBorder: Border;
    /**
     * Sets the bottom border of the cell
     * @param {Border} border - The bottom border of the cell
     * @returns {void}
     */
    setBorderBottom(border: Border): void;
    /**
     * Sets the left border of the cell
     * @param {Border} border - The left border of the cell
     * @returns {void}
     */
    setBorderLeft(border: Border): void;
    /**
     * Sets the right border of the cell
     * @param {Border} border - The right border of the cell
     * @returns {void}
     */
    setBorderRight(border: Border): void;
    /**
     * Sets the top border of the cell
     * @param {Border} border - The top border of the cell
     * @returns {void}
     */
    setBorderTop(border: Border): void;
    /**
     * Sets the diagonal border of the cell
     * @param {DiagonalBorder} border - The diagonal border of the cell
     * @returns {void}
     */
    setBorderDiagonal(border: DiagonalBorder): void;
    diagonalBorder: DiagonalBorder;
    /**
     * Sets the font charset
     * @param {number} charset - The charset of the font
     * @returns {void}
     */
    setFontCharset(charset: number): void;
    charset: number;
    /**
     * Sets the font color
     * @param {Color} color - The color of the font
     * @returns {void}
     */
    setFontColor(color: Color): void;
    fontColor: Color;
    /**
     * Sets the font family
     * @param {number} fontFamily - The family of the font
     * @returns {void}
     */
    setFontFamily(fontFamily: number): void;
    fontFamily: number;
    /**
     * Sets the font name
     * @param {string} fontName - The name of the font
     * @returns {void}
     */
    setFontName(fontName: string): void;
    fontName: string;
    /**
     * Set the Format font scheme property.
     * @param {string} fontScheme - The font scheme
     * @returns {void}
     */
    setFontScheme(fontScheme: string): void;
    fontScheme: string;
    /**
     * Set the font size property
     * @param {number} fontSize - The font size
     * @returns {void}
     */
    setFontSize(fontSize: number): void;
    fontSize: number;
    /**
     * Set the font strike through property
     * @returns {void}
     */
    setFontStrikeThrough(): void;
    strikeThrough: boolean;
    /**
     * Set the foreground color property
     * @param {Color} color - The color
     * @returns {void}
     */
    setForegroundColor(color: Color): void;
    foregroundColor: Color;
    /**
     * Set if the format is hidden
     * @param {boolean} hidden - If the format is hidden
     * @returns {void}
     */
    setHidden(hidden: boolean): void;
    hidden: boolean;
    /**
     * Set if the format in hyperlinked
     * @param {boolean} hyperlink - If the format is hyperlinked
     * @returns {void}
     */
    setHyperlink(hyperlink: boolean): void;
    hyperlink: boolean;
    /**
     * Set if the format is indented
     * @param {number} indent - The indent level
     * @returns {void}
     */
    setIndent(indent: number): void;
    indent: number;
    /**
     * Set the italic property
     * @param {boolean} italic - If the format is italic
     * @returns {void}
     */
    setItalic(italic: boolean): void;
    italic: boolean;
    /**
     * Set the lock property
     * @param {boolean} locked - If the format is locked
     * @returns {void}
     */
    setLocked(locked: boolean): void;
    locked: boolean;
    /**
     * Set the number format property
     * @param {string} numFmt - The number format
     * @returns {void}
     */
    setNumFmt(numFmt: string): void;
    numFmt: string;
    /**
     * Set the number format id property
     * @param {number} numFmtId - The number format id
     * @returns {void}
     */
    setNumFmtId(numFmtId: number): void;
    numFmtId: number;
    /**
     * Set the format pattern
     * @param {FormatPattern} pattern - The pattern
     * @returns {void}
     */
    setPattern(pattern: FormatPattern): void;
    pattern: FormatPattern;
    /**
     * Set the underline properties for a format.
     * @param {FormatUnderline} underline - The underline style
     * @returns {void}
     */
    setUnderline(underline: FormatUnderline): void;
    underline: FormatUnderline;
}
declare namespace Format {
    export { FormatAlign, FormatPattern, FormatUnderline };
}
import Color = require("./color");
import { Border } from "./border";
import { DiagonalBorder } from "./border";
/**
 * - The alignment of the cell
 */
type FormatAlign = ("general" | "left" | "center" | "right" | "fill" | "justify" | "centerAcross" | "distributed" | "top" | "bottom" | "verticalCenter" | "verticalJustify" | "verticalDistributed");
/**
 * - The pattern of the cell
 */
type FormatPattern = ("none" | "solid" | "mediumGray" | "darkGray" | "lightGray" | "darkHorizontal" | "darkVertical" | "darkDown" | "darkUp" | "darkGrid" | "darkTrellis" | "lightHorizontal" | "lightVertical" | "lightDown" | "lightUp" | "lightGrid" | "lightTrellis" | "gray125" | "gray0625");
/**
 * - The underline style
 */
type FormatUnderline = ("none" | "single" | "double" | "singleAccounting" | "doubleAccounting");
//# sourceMappingURL=format.d.ts.map