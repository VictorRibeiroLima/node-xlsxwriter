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
 *
 */
class Format {
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
  constructor({
    align,
    backgroundColor,
    bold,
    leftBorder,
    rightBorder,
    topBorder,
    bottomBorder,
    diagonalBorder,
    charset,
    fontColor,
    fontFamily,
    fontName,
    fontScheme,
    fontSize,
    strikeThrough,
    foregroundColor,
    hidden,
    hyperlink,
    indent,
    italic,
    locked,
    numFmt,
    numFmtId,
    pattern,
    underline,
  } = {}) {
    this.id = Math.floor(Math.random() * 1_000_000);
    this.align = align;
    this.backgroundColor = backgroundColor;
    this.bold = bold;
    this.leftBorder = leftBorder;
    this.rightBorder = rightBorder;
    this.topBorder = topBorder;
    this.bottomBorder = bottomBorder;
    this.diagonalBorder = diagonalBorder;
    this.charset = charset;
    this.fontColor = fontColor;
    this.fontFamily = fontFamily;
    this.fontName = fontName;
    this.fontScheme = fontScheme;
    this.fontSize = fontSize;
    this.strikeThrough = strikeThrough;
    this.foregroundColor = foregroundColor;
    this.hidden = hidden;
    this.hyperlink = hyperlink;
    this.indent = indent;
    this.italic = italic;
    this.locked = locked;
    this.numFmt = numFmt;
    this.numFmtId = numFmtId;
    this.pattern = pattern;
    this.underline = underline;
  }

  /**
   * Sets alignment of the cell
   * @param {FormatAlign} align - The alignment of the cell
   * @returns {void}
   */
  setAlignment(align) {
    this.align = align;
  }

  /**
   * Sets the font color
   * @param {Color} color - The color of the font
   * @returns {void}
   */
  setBackgroundColor(color) {
    this.backgroundColor = color;
  }

  /**
   * Sets if the font is bold
   * @param {boolean} bold - If the font is bold
   * @returns {void}
   */
  setBold(bold) {
    this.bold = bold;
  }

  /**
   * Sets the border of the cell
   * @param {Border} border - The border of the cell
   * @returns {void}
   */
  setBorder(border) {
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
    this.bottomBorder = border;
  }

  /**
   * Sets the left border of the cell
   * @param {Border} border - The left border of the cell
   * @returns {void}
   */
  setBorderLeft(border) {
    this.leftBorder = border;
  }

  /**
   * Sets the right border of the cell
   * @param {Border} border - The right border of the cell
   * @returns {void}
   */
  setBorderRight(border) {
    this.rightBorder = border;
  }

  /**
   * Sets the top border of the cell
   * @param {Border} border - The top border of the cell
   * @returns {void}
   */
  setBorderTop(border) {
    this.topBorder = border;
  }

  /**
   * Sets the diagonal border of the cell
   * @param {DiagonalBorder} border - The diagonal border of the cell
   * @returns {void}
   */
  setBorderDiagonal(border) {
    this.diagonalBorder = border;
  }

  /**
   * Sets the font charset
   * @param {number} charset - The charset of the font
   * @returns {void}
   */
  setFontCharset(charset) {
    this.charset = charset;
  }

  /**
   * Sets the font color
   * @param {Color} color - The color of the font
   * @returns {void}
   */
  setFontColor(color) {
    this.fontColor = color;
  }

  /**
   * Sets the font family
   * @param {number} fontFamily - The family of the font
   * @returns {void}
   */
  setFontFamily(fontFamily) {
    this.fontFamily = fontFamily;
  }

  /**
   * Sets the font name
   * @param {string} fontName - The name of the font
   * @returns {void}
   */
  setFontName(fontName) {
    this.fontName = fontName;
  }

  /**
   * Set the Format font scheme property.
   * @param {string} fontScheme - The font scheme
   * @returns {void}
   */
  setFontScheme(fontScheme) {
    this.fontScheme = fontScheme;
  }

  /**
   * Set the font size property
   * @param {number} fontSize - The font size
   * @returns {void}
   */
  setFontSize(fontSize) {
    this.fontSize = fontSize;
  }

  /**
   * Set the font strike through property
   * @returns {void}
   */
  setFontStrikeThrough() {
    this.strikeThrough = true;
  }

  /**
   * Set the foreground color property
   * @param {Color} color - The color
   * @returns {void}
   */
  setForegroundColor(color) {
    this.foregroundColor = color;
  }

  /**
   * Set if the format is hidden
   * @param {boolean} hidden - If the format is hidden
   * @returns {void}
   */
  setHidden(hidden) {
    this.hidden = hidden;
  }

  /**
   * Set if the format in hyperlinked
   * @param {boolean} hyperlink - If the format is hyperlinked
   * @returns {void}
   */
  setHyperlink(hyperlink) {
    this.hyperlink = hyperlink;
  }

  /**
   * Set if the format is indented
   * @param {number} indent - The indent level
   * @returns {void}
   */
  setIndent(indent) {
    this.indent = indent;
  }

  /**
   * Set the italic property
   * @param {boolean} italic - If the format is italic
   * @returns {void}
   */
  setItalic(italic) {
    this.italic = italic;
  }

  /**
   * Set the lock property
   * @param {boolean} locked - If the format is locked
   * @returns {void}
   */
  setLocked(locked) {
    this.locked = locked;
  }

  /**
   * Set the number format property
   * @param {string} numFmt - The number format
   * @returns {void}
   */
  setNumFmt(numFmt) {
    this.numFmt = numFmt;
  }

  /**
   * Set the number format id property
   * @param {number} numFmtId - The number format id
   * @returns {void}
   */
  setNumFmtId(numFmtId) {
    this.numFmtId = numFmtId;
  }

  /**
   * Set the format pattern
   * @param {FormatPattern} pattern - The pattern
   * @returns {void}
   */
  setPattern(pattern) {
    this.pattern = pattern;
  }

  /**
   * Set the underline properties for a format.
   * @param {FormatUnderline} underline - The underline style
   * @returns {void}
   */
  setUnderline(underline) {
    this.underline = underline;
  }
}

module.exports = Format;
