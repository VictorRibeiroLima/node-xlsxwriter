export = Link;
/**
 * @class Link
 * @classdesc Represents a link
 * @property {string} url - The URL of the link
 * @property {string} [text] - The text of the link
 * @property {string} [tip] - The tooltip of the link
 */
declare class Link {
    /**
     * @param {string} url - The URL of the link
     * @param {string} [text] - The text of the link
     * @param {string} [tip] - The tooltip of the link
     */
    constructor(url: string, text?: string, tip?: string);
    url: string;
    text: string;
    tip: string;
}
//# sourceMappingURL=link.d.ts.map