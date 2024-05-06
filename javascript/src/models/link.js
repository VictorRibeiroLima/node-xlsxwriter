// @ts-check

/**
 * @class Link
 * @classdesc Represents a link
 * @property {string} url - The URL of the link
 * @property {string} [text] - The text of the link
 * @property {string} [tip] - The tooltip of the link
 */
class Link {
  /**
   * @param {string} url - The URL of the link
   * @param {string} [text] - The text of the link
   * @param {string} [tip] - The tooltip of the link
   */
  constructor(url, text, tip) {
    this.url = url;
    this.text = text;
    this.tip = tip;
  }
}

module.exports = Link;
