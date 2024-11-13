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
 * @property {FormatAlign|undefined} [align=undefined] - The alignment of the cell
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
 *
 */
declare class Format {
    /**
     *
     * @param {Object} [options] - The options object
     * @param {FormatAlign} [options.align] - The alignment of the cell
     * @param {Color} [options.backgroundColor] - The background color of the cell
     * @param {boolean} [options.bold] - If the font is bold
     * @param {Border} [options.leftBorder] - The left border of the cell
     * @param {Border} [options.rightBorder] - The right border of the cell
     * @param {Border} [options.topBorder] - The top border of the cell
     * @param {Border} [options.bottomBorder] - The bottom border of the cell
     * @param {DiagonalBorder} [options.diagonalBorder] - The diagonal border of the cell
     * @param {number} [options.charset] - The charset of the font
     * @param {Color} [options.fontColor] - The color of the font
     * @param {number} [options.fontFamily] - The family of the font
     * @param {string} [options.fontName] - The name of the font
     * @param {string} [options.fontScheme] - The font scheme
     * @param {number} [options.fontSize] - The font size
     * @param {boolean} [options.strikeThrough] - If the font is strike through
     * @param {Color} [options.foregroundColor] - The foreground color
     * @param {boolean} [options.hidden] - If the format is hidden
     * @param {boolean} [options.hyperlink] - If the format is hyperlinked
     * @param {number} [options.indent] - The indent level
     * @param {boolean} [options.italic] - If the font is italic
     * @param {boolean} [options.locked] - If the format is locked
     * @param {string} [options.numFmt] - The number format
     * @param {number} [options.numFmtId] - The number format id
     * @param {FormatPattern} [options.pattern] - The pattern
     * @param {FormatUnderline} [options.underline] - The underline style
     */
    constructor({ align, backgroundColor, bold, leftBorder, rightBorder, topBorder, bottomBorder, diagonalBorder, charset, fontColor, fontFamily, fontName, fontScheme, fontSize, strikeThrough, foregroundColor, hidden, hyperlink, indent, italic, locked, numFmt, numFmtId, pattern, underline, }?: {
        align?: FormatAlign;
        backgroundColor?: Color;
        bold?: boolean;
        leftBorder?: Border;
        rightBorder?: Border;
        topBorder?: Border;
        bottomBorder?: Border;
        diagonalBorder?: DiagonalBorder;
        charset?: number;
        fontColor?: Color;
        fontFamily?: number;
        fontName?: string;
        fontScheme?: string;
        fontSize?: number;
        strikeThrough?: boolean;
        foregroundColor?: Color;
        hidden?: boolean;
        hyperlink?: boolean;
        indent?: number;
        italic?: boolean;
        locked?: boolean;
        numFmt?: string;
        numFmtId?: number;
        pattern?: FormatPattern;
        underline?: FormatUnderline;
    });
    id: number;
    /**
     * The alignment of the cell
     * @type {?FormatAlign}
     * @default undefined
     */
    align: FormatAlign | null;
    /**
     * The background color of the cell
     * @type {?Color}
     * @default undefined
     */
    backgroundColor: Color | null;
    /**
     * If the font is bold
     * @type {?boolean} [bold]
     * @default undefined
     */
    bold: boolean | null;
    /**
     * The left border of the cell
     * @type {?Border}
     * @default undefined
     */
    leftBorder: Border | null;
    /**
     * The right border of the cell
     * @type {?Border}
     * @default undefined
     */
    rightBorder: Border | null;
    /**
     * The top border of the cell
     * @type {?Border}
     * @default undefined
     */
    topBorder: Border | null;
    /**
     * The bottom border of the cell
     * @type {?Border}
     * @default undefined
     */
    bottomBorder: Border | null;
    /**
     * The diagonal border of the cell
     * @type {?Border}
     * @default undefined
     */
    diagonalBorder: Border | null;
    /**
     * The charset of the cell
     * @type {?number}
     * @default undefined
     */
    charset: number | null;
    /**
     * The font color of the cell
     * @type {?Color}
     * @default undefined
     */
    fontColor: Color | null;
    /**
     * The font family of the cell
     * @type {?number}
     * @default undefined
     */
    fontFamily: number | null;
    /**
     * The font name of the cell
     * @type {?string}
     * @default undefined
     */
    fontName: string | null;
    /**
     * The font scheme of the cell
     * @type {?string}
     * @default undefined
     */
    fontScheme: string | null;
    /**
     * The font size of the cell
     * @type {?number}
     * @default undefined
     */
    fontSize: number | null;
    /**
     * If the font is strikethrough
     * @type {?boolean}
     * @default undefined
     */
    strikeThrough: boolean | null;
    /**
     * The foreground color of the cell
     * @type {?Color}
     * @default undefined
     */
    foregroundColor: Color | null;
    /**
     * If the cell is hidden
     * @type {?boolean}
     * @default undefined
     */
    hidden: boolean | null;
    /**
     * The hyperlink of the cell
     * @type {?boolean}
     * @default undefined
     */
    hyperlink: boolean | null;
    /**
     * The indent of the cell
     * @type {?number}
     * @default undefined
     */
    indent: number | null;
    /**
     * If the font is italic
     * @type {?boolean}
     * @default undefined
     */
    italic: boolean | null;
    /**
     * If the cell is locked
     * @type {?boolean}
     * @default undefined
     */
    locked: boolean | null;
    /**
     * The number format of the cell
     * @type {?string}
     * @default undefined
     */
    numFmt: string | null;
    /**
     * The number format ID of the cell
     * @type {?number}
     * @default undefined
     */
    numFmtId: number | null;
    /**
     * The pattern of the cell
     * @type {?FormatPattern}
     * @default undefined
     */
    pattern: FormatPattern | null;
    /**
     * If the font is underlined
     * @type {?string}
     * @default undefined
     */
    underline: string | null;
    /**
     * Sets alignment of the cell
     * @param {FormatAlign} align - The alignment of the cell
     * @returns {void}
     */
    setAlignment(align: FormatAlign): void;
    /**
     * Sets the font color
     * @param {Color} color - The color of the font
     * @returns {void}
     */
    setBackgroundColor(color: Color): void;
    /**
     * Sets if the font is bold
     * @param {boolean} bold - If the font is bold
     * @returns {void}
     */
    setBold(bold: boolean): void;
    /**
     * Sets the border of the cell (left, right, top, bottom)
     * @param {Border} border - The border of the cell
     * @returns {void}
     */
    setBorder(border: Border): void;
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
    /**
     * Sets the font charset
     * @param {number} charset - The charset of the font
     * @returns {void}
     */
    setFontCharset(charset: number): void;
    /**
     * Sets the font color
     * @param {Color} color - The color of the font
     * @returns {void}
     */
    setFontColor(color: Color): void;
    /**
     * Sets the font family
     * @param {number} fontFamily - The family of the font
     * @returns {void}
     */
    setFontFamily(fontFamily: number): void;
    /**
     * Sets the font name
     * @param {string} fontName - The name of the font
     * @returns {void}
     */
    setFontName(fontName: string): void;
    /**
     * Set the Format font scheme property.
     * @param {string} fontScheme - The font scheme
     * @returns {void}
     */
    setFontScheme(fontScheme: string): void;
    /**
     * Set the font size property
     * @param {number} fontSize - The font size
     * @returns {void}
     */
    setFontSize(fontSize: number): void;
    /**
     * Set the font strike through property
     * @returns {void}
     */
    setFontStrikeThrough(): void;
    /**
     * Set the foreground color property
     * @param {Color} color - The color
     * @returns {void}
     */
    setForegroundColor(color: Color): void;
    /**
     * Set if the format is hidden
     * @param {boolean} hidden - If the format is hidden
     * @returns {void}
     */
    setHidden(hidden: boolean): void;
    /**
     * Set if the format in hyperlinked
     * @param {boolean} hyperlink - If the format is hyperlinked
     * @returns {void}
     */
    setHyperlink(hyperlink: boolean): void;
    /**
     * Set if the format is indented
     * @param {number} indent - The indent level
     * @returns {void}
     */
    setIndent(indent: number): void;
    /**
     * Set the italic property
     * @param {boolean} italic - If the format is italic
     * @returns {void}
     */
    setItalic(italic: boolean): void;
    /**
     * Set the lock property
     * @param {boolean} locked - If the format is locked
     * @returns {void}
     */
    setLocked(locked: boolean): void;
    /**
     * Set the number format property
     * @param {string} numFmt - The number format
     * @returns {void}
     */
    setNumFmt(numFmt: string): void;
    /**
     * Set the number format id property
     * @param {number} numFmtId - The number format id
     * @returns {void}
     */
    setNumFmtId(numFmtId: number): void;
    /**
     * Set the format pattern
     * @param {FormatPattern} pattern - The pattern
     * @returns {void}
     */
    setPattern(pattern: FormatPattern): void;
    /**
     * Set the underline properties for a format.
     * @param {FormatUnderline} underline - The underline style
     * @returns {void}
     */
    setUnderline(underline: FormatUnderline): void;
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