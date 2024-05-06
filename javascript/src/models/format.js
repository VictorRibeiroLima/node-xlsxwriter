// @ts-check

const { Border, DiagonalBorder } = require('./border');
const Color = require('./color');

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
 * @property {boolean} default - If the format is the default format (do not set this directly)
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
class Format {
  constructor() {
    /**
     * the id of the format
     * @private
     * @type {number}
     */
    this.id = Math.floor(Math.random() * 1_000_000);
    /**
     * if the format is the default format
     * @private
     * @type {boolean}
     */
    this.default = true;
  }

  /**
   * Sets alignment of the cell
   * @param {FormatAlign} align - The alignment of the cell
   * @returns {void}
   */
  setAlignment(align) {
    this.default = false;
    this.align = align;
  }

  /**
   * Sets the font color
   * @param {Color} color - The color of the font
   * @returns {void}
   */
  setBackgroundColor(color) {
    this.default = false;
    this.backgroundColor = color;
  }

  /**
   * Sets if the font is bold
   * @param {boolean} bold - If the font is bold
   * @returns {void}
   */
  setBold(bold) {
    this.default = false;
    this.bold = bold;
  }

  /**
   * Sets the border of the cell
   * @param {Border} border - The border of the cell
   * @returns {void}
   */
  setBorder(border) {
    this.default = false;
    this.leftBorder = border;
    this.rightBorder = border;
    this.topBorder = border;
    this.bottomBorder = border;
  }

  /**
   * Sets the bottom border of the cell
   * @param {Border} border - The bottom border of the cell
   * @returns {void}
   */
  setBorderBottom(border) {
    this.default = false;
    this.bottomBorder = border;
  }

  /**
   * Sets the left border of the cell
   * @param {Border} border - The left border of the cell
   * @returns {void}
   */
  setBorderLeft(border) {
    this.default = false;
    this.leftBorder = border;
  }

  /**
   * Sets the right border of the cell
   * @param {Border} border - The right border of the cell
   * @returns {void}
   */
  setBorderRight(border) {
    this.default = false;
    this.rightBorder = border;
  }

  /**
   * Sets the top border of the cell
   * @param {Border} border - The top border of the cell
   * @returns {void}
   */
  setBorderTop(border) {
    this.default = false;
    this.topBorder = border;
  }

  /**
   * Sets the diagonal border of the cell
   * @param {DiagonalBorder} border - The diagonal border of the cell
   * @returns {void}
   */
  setBorderDiagonal(border) {
    this.default = false;
    this.diagonalBorder = border;
  }

  /**
   * Sets the font charset
   * @param {number} charset - The charset of the font
   * @returns {void}
   */
  setFontCharset(charset) {
    this.default = false;
    this.charset = charset;
  }

  /**
   * Sets the font color
   * @param {Color} color - The color of the font
   * @returns {void}
   */
  setFontColor(color) {
    this.default = false;
    this.fontColor = color;
  }

  /**
   * Sets the font family
   * @param {number} fontFamily - The family of the font
   * @returns {void}
   */
  setFontFamily(fontFamily) {
    this.default = false;
    this.fontFamily = fontFamily;
  }

  /**
   * Sets the font name
   * @param {string} fontName - The name of the font
   * @returns {void}
   */
  setFontName(fontName) {
    this.default = false;
    this.fontName = fontName;
  }

  /**
   * Set the Format font scheme property.
   * @param {string} fontScheme - The font scheme
   * @returns {void}
   */
  setFontScheme(fontScheme) {
    this.default = false;
    this.fontScheme = fontScheme;
  }

  /**
   * Set the font size property
   * @param {number} fontSize - The font size
   * @returns {void}
   */
  setFontSize(fontSize) {
    this.default = false;
    this.fontSize = fontSize;
  }

  /**
   * Set the font strike through property
   * @returns {void}
   */
  setFontStrikeThrough() {
    this.default = false;
    this.strikeThrough = true;
  }

  /**
   * Set the foreground color property
   * @param {Color} color - The color
   * @returns {void}
   */
  setForegroundColor(color) {
    this.default = false;
    this.foregroundColor = color;
  }

  /**
   * Set if the format is hidden
   * @param {boolean} hidden - If the format is hidden
   * @returns {void}
   */
  setHidden(hidden) {
    this.default = false;
    this.hidden = hidden;
  }

  /**
   * Set if the format in hyperlinked
   * @param {boolean} hyperlink - If the format is hyperlinked
   * @returns {void}
   */
  setHyperlink(hyperlink) {
    this.default = false;
    this.hyperlink = hyperlink;
  }

  /**
   * Set if the format is indented
   * @param {number} indent - The indent level
   * @returns {void}
   */
  setIndent(indent) {
    this.default = false;
    this.indent = indent;
  }

  /**
   * Set the italic property
   * @param {boolean} italic - If the format is italic
   * @returns {void}
   */
  setItalic(italic) {
    this.default = false;
    this.italic = italic;
  }

  /**
   * Set the lock property
   * @param {boolean} locked - If the format is locked
   * @returns {void}
   */
  setLocked(locked) {
    this.default = false;
    this.locked = locked;
  }

  /**
   * Set the number format property
   * @param {string} numFmt - The number format
   * @returns {void}
   */
  setNumFmt(numFmt) {
    this.default = false;
    this.numFmt = numFmt;
  }

  /**
   * Set the number format id property
   * @param {number} numFmtId - The number format id
   * @returns {void}
   */
  setNumFmtId(numFmtId) {
    this.default = false;
    this.numFmtId = numFmtId;
  }

  /**
   * Set the format pattern
   * @param {FormatPattern} pattern - The pattern
   * @returns {void}
   */
  setPattern(pattern) {
    this.default = false;
    this.pattern = pattern;
  }

  /**
   * Set the underline properties for a format.
   * @param {FormatUnderline} underline - The underline style
   * @returns {void}
   */
  setUnderline(underline) {
    this.default = false;
    this.underline = underline;
  }
}

module.exports = Format;
