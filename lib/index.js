import { convertHtmlToDocx, convertHtmlToDocxFile } from './converter.js';

class HtmlToDocx {
  constructor(options = {}) {
    this.options = {
      fontSize: options.fontSize || 11,
      fontFamily: options.fontFamily || 'Calibri',
      lineHeight: options.lineHeight || 1.15,
      ...options
    };
  }

  /**
   * Convert HTML string to DOCX buffer
   * @param {string} html - HTML content to convert
   * @returns {Promise<Buffer>} - DOCX file buffer
   */
  async convertHtmlToDocx(html) {
    return convertHtmlToDocx(html, this.options);
  }

  /**
   * Convert HTML string to DOCX file
   * @param {string} html - HTML content to convert
   * @param {string} outputPath - Path where to save the DOCX file
   * @returns {Promise<void>}
   */
  async convertHtmlToDocxFile(html, outputPath) {
    return convertHtmlToDocxFile(html, outputPath, this.options);
  }
}

export default HtmlToDocx;
export { convertHtmlToDocx, convertHtmlToDocxFile };
